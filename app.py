import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from streamlit_mermaid import st_mermaid
import io
import zipfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Librerías para Word
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 1. Configuración centrada y minimalista
st.set_page_config(page_title="MatrixDevTesis", layout="centered", page_icon="🍷")

# 2. Estilos personalizados
st.markdown("""
    <style>
    /* Fondo principal y color de texto */
    .stApp {
        background-color: #2c0f14;
        color: #f4ecec;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
    }
    
    /* Encabezados */
    h1, h2, h3 {
        color: #e09f9f !important;
        font-weight: 300 !important;
        text-align: center;
    }
    
    /* Botones */
    .stButton>button, .stFormSubmitButton>button {
        background-color: #5c1e28;
        color: #ffffff !important;
        border: 1px solid #8a2d3b;
        border-radius: 8px;
        transition: 0.3s;
        font-weight: bold;
    }
    .stButton>button:hover, .stFormSubmitButton>button:hover {
        background-color: #8a2d3b;
        border-color: #ffffff;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.5);
    }

    /* Estilo de los contenedores de items */
    div[data-testid="stHorizontalBlock"] {
        align-items: center;
    }

    /* Ocultar elementos nativos innecesarios */
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# INICIALIZACIÓN DE VARIABLES DE ESTADO
# ==========================================
if "nombre_proj" not in st.session_state:
    st.session_state.nombre_proj = "MatrixDev Core"
if "integrantes" not in st.session_state:
    st.session_state.integrantes = "Juan Pérez"
if "objetivo_text" not in st.session_state:
    st.session_state.objetivo_text = "Automatizar el flujo de inventario con una arquitectura minimalista."

# Requisitos
if "rf_list" not in st.session_state:
    st.session_state.rf_list = ["Autenticación OAuth2.", "CRUD de usuarios."]
if "rnf_list" not in st.session_state:
    st.session_state.rnf_list = ["Latencia < 200ms.", "Cifrado AES-256 en base de datos."]

# Kanban
if "kanban_tasks" not in st.session_state:
    st.session_state.kanban_tasks = [
        {"Tarea": "Modelo BD", "Estado": "Completado"},
        {"Tarea": "Endpoints API", "Estado": "En Proceso"}
    ]

# EDT / WBS
if "edt_list" not in st.session_state:
    st.session_state.edt_list = [
        {"Fase": "1. Inicio y Requisitos", "Paquete": "Levantamiento de Información", "Horas": 20},
        {"Fase": "1. Inicio y Requisitos", "Paquete": "Especificación ERS", "Horas": 15},
        {"Fase": "2. Diseño y Arquitectura", "Paquete": "Modelado de Base de Datos", "Horas": 30},
        {"Fase": "2. Diseño y Arquitectura", "Paquete": "Diagramación de Flujos", "Horas": 25},
        {"Fase": "3. Desarrollo", "Paquete": "Implementación Backend API", "Horas": 80},
        {"Fase": "3. Desarrollo", "Paquete": "Interfaz Frontend Streamlit", "Horas": 60},
    ]

# Riesgos
if "riesgos_list" not in st.session_state:
    st.session_state.riesgos_list = [
        {"Riesgo": "Retraso en entrega de APIs", "Probabilidad": 3, "Impacto": 4},
        {"Riesgo": "Incompatibilidad de base de datos", "Probabilidad": 2, "Impacto": 5},
        {"Riesgo": "Cambio no planificado de requisitos", "Probabilidad": 4, "Impacto": 3}
    ]

# ==========================================
# FUNCIONES DE GENERACIÓN DE DOCUMENTOS
# ==========================================
def generar_word_ers():
    doc = Document()
    doc.add_heading(f"Especificación de Requisitos: {st.session_state.nombre_proj}", level=1)
    doc.add_paragraph(f"Objetivo: {st.session_state.objetivo_text}")
    
    doc.add_heading("Requisitos Funcionales", level=2)
    for idx, rf in enumerate(st.session_state.rf_list, 1):
        if rf.strip():
            doc.add_paragraph(f"RF-{idx:02d}: {rf.strip()}")
            
    doc.add_heading("Requisitos No Funcionales", level=2)
    for idx, rnf in enumerate(st.session_state.rnf_list, 1):
        if rnf.strip():
            doc.add_paragraph(f"RNF-{idx:02d}: {rnf.strip()}")
            
    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

def generar_excel_estilizado():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # Hoja Requisitos
    ws_req = wb.active
    ws_req.title = "Requisitos"
    ws_req.append(["Código", "Tipo", "Descripción"])
    for idx, rf in enumerate(st.session_state.rf_list, 1):
        if rf.strip():
            ws_req.append([f"RF-{idx:02d}", "Funcional", rf.strip()])
    for idx, rnf in enumerate(st.session_state.rnf_list, 1):
        if rnf.strip():
            ws_req.append([f"RNF-{idx:02d}", "No Funcional", rnf.strip()])

    # Hoja EDT
    ws_edt = wb.create_sheet(title="EDT (WBS)")
    ws_edt.append(["Fase", "Paquete de Trabajo", "Horas Estimadas"])
    for item in st.session_state.edt_list:
        ws_edt.append([item["Fase"], item["Paquete"], item["Horas"]])

    # Hoja Riesgos
    ws_risk = wb.create_sheet(title="Matriz Riesgos")
    ws_risk.append(["Riesgo", "Probabilidad (1-5)", "Impacto (1-5)", "Nivel Severidad"])
    for r in st.session_state.riesgos_list:
        ws_risk.append([r["Riesgo"], r["Probabilidad"], r["Impacto"], r["Probabilidad"] * r["Impacto"]])

    wb.save(output)
    return output.getvalue()

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.title("✦ MatrixDev ✦")
st.markdown("<p style='text-align: center; color: #a68a8d;'>Gestión y Arquitectura Avanzada de Proyectos</p>", unsafe_allow_html=True)
st.write("---")

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab_full = st.tabs([
    "1. Requisitos", 
    "2. Modelado", 
    "3. Gestión", 
    "4. Monte Carlo",
    "5. Flujos",
    "6. EDT (WBS)",
    "7. Riesgos",
    "📦 Descarga Full"
])

# --- TAB 1: REQUISITOS ---
with tab1:
    st.markdown("### Definición del Sistema")
    st.session_state.nombre_proj = st.text_input("Nombre del Proyecto", st.session_state.nombre_proj)
    st.session_state.objetivo_text = st.text_area("Objetivo Principal", st.session_state.objetivo_text, height=80)
    
    st.write("---")
    st.markdown("### Requisitos Funcionales (RF)")
    
    with st.form("form_add_rf", clear_on_submit=True):
        col_in, col_btn = st.columns([4, 1])
        with col_in:
            nuevo_rf_val = st.text_input("Nuevo RF", placeholder="Escribe un requisito...", label_visibility="collapsed")
        with col_btn:
            btn_add_rf = st.form_submit_button(" Añadir", use_container_width=True)
            
        if btn_add_rf and nuevo_rf_val.strip():
            st.session_state.rf_list.append(nuevo_rf_val.strip())
            st.rerun()

    if not st.session_state.rf_list:
        st.info("No hay requisitos funcionales registrados.")
    else:
        rf_to_delete = None
        for i, item in enumerate(st.session_state.rf_list):
            c_tag, c_input, c_del = st.columns([0.8, 5, 0.8])
            c_tag.markdown(f"**RF-{i+1:02d}**")
            new_val = c_input.text_input(f"rf_in_{i}", value=item, label_visibility="collapsed", key=f"rf_field_{i}")
            st.session_state.rf_list[i] = new_val
            
            if c_del.button("🗑️", key=f"del_rf_{i}"):
                rf_to_delete = i

        if rf_to_delete is not None:
            st.session_state.rf_list.pop(rf_to_delete)
            st.rerun()

    st.write("---")
    st.markdown("### Requisitos No Funcionales (RNF)")
    
    with st.form("form_add_rnf", clear_on_submit=True):
        col_in_rnf, col_btn_rnf = st.columns([4, 1])
        with col_in_rnf:
            nuevo_rnf_val = st.text_input("Nuevo RNF", placeholder="Escribe un requisito no funcional...", label_visibility="collapsed")
        with col_btn_rnf:
            btn_add_rnf = st.form_submit_button(" Añadir", use_container_width=True)
            
        if btn_add_rnf and nuevo_rnf_val.strip():
            st.session_state.rnf_list.append(nuevo_rnf_val.strip())
            st.rerun()

    if not st.session_state.rnf_list:
        st.info("No hay requisitos no funcionales registrados.")
    else:
        rnf_to_delete = None
        for i, item in enumerate(st.session_state.rnf_list):
            c_tag, c_input, c_del = st.columns([0.8, 5, 0.8])
            c_tag.markdown(f"**RNF-{i+1:02d}**")
            new_val = c_input.text_input(f"rnf_in_{i}", value=item, label_visibility="collapsed", key=f"rnf_field_{i}")
            st.session_state.rnf_list[i] = new_val
            
            if c_del.button("🗑️", key=f"del_rnf_{i}"):
                rnf_to_delete = i

        if rnf_to_delete is not None:
            st.session_state.rnf_list.pop(rnf_to_delete)
            st.rerun()

# --- TAB 2: MODELADO ---
with tab2:
    st.markdown("### Arquitectura Visual")
    default_mermaid = "graph TD\n A[Inicio] --> B{Validar}\n B -- Sí --> C[Éxito]\n B -- No --> D[Error]"
    codigo_mermaid = st.text_area("Sintaxis Mermaid", value=default_mermaid, height=120)
    st_mermaid(codigo_mermaid)

# --- TAB 3: GESTIÓN ---
with tab3:
    st.markdown("### Tablero Kanban")
    
    with st.form("form_kanban", clear_on_submit=True):
        col_t, col_s, col_b = st.columns([3, 2, 1])
        with col_t:
            nueva_t = st.text_input("Tarea", placeholder="Nueva tarea...", label_visibility="collapsed")
        with col_s:
            estado_t = st.selectbox("Estado", ["Pendiente", "En Proceso", "Completado"], label_visibility="collapsed")
        with col_b:
            btn_add_k = st.form_submit_button(" Añadir", use_container_width=True)
            
        if btn_add_k and nueva_t.strip():
            st.session_state.kanban_tasks.append({"Tarea": nueva_t.strip(), "Estado": estado_t})
            st.rerun()

    if not st.session_state.kanban_tasks:
        st.info("No hay tareas registradas en el tablero.")
    else:
        task_to_delete = None
        for i, t in enumerate(st.session_state.kanban_tasks):
            col_txt, col_sel, col_d = st.columns([3, 2, 0.8])
            updated_text = col_txt.text_input(f"kt_{i}", value=t["Tarea"], label_visibility="collapsed", key=f"kt_in_{i}")
            updated_status = col_sel.selectbox(
                f"ks_{i}", 
                ["Pendiente", "En Proceso", "Completado"], 
                index=["Pendiente", "En Proceso", "Completado"].index(t["Estado"]), 
                label_visibility="collapsed", 
                key=f"ks_in_{i}"
            )
            
            st.session_state.kanban_tasks[i]["Tarea"] = updated_text
            st.session_state.kanban_tasks[i]["Estado"] = updated_status
            
            if col_d.button("🗑️", key=f"del_k_{i}"):
                task_to_delete = i
                
        if task_to_delete is not None:
            st.session_state.kanban_tasks.pop(task_to_delete)
            st.rerun()

# --- TAB 4: MONTE CARLO ---
with tab4:
    st.markdown("### Simulación Monte Carlo (Estimación PERT)")
    st.write("Calcula la probabilidad de cumplir plazos estimando tiempos en días (Optimista, Probable, Pesimista).")
    
    c1, c2, c3 = st.columns(3)
    opt = c1.number_input("Días Optimista (a)", min_value=1, value=10)
    mod = c2.number_input("Días Probable (m)", min_value=1, value=20)
    pes = c3.number_input("Días Pesimista (b)", min_value=1, value=35)
    
    sims = st.slider("Número de Simulaciones", 100, 5000, 1000, step=100)
    
    if pes < mod or mod < opt:
        st.warning("Asegúrate de que: Optimista <= Probable <= Pesimista.")
    else:
        # Generar estimaciones mediante Distribución Triangular
        data_sim = np.random.triangular(opt, mod, pes, sims)
        
        # Percentiles
        p50 = np.percentile(data_sim, 50)
        p80 = np.percentile(data_sim, 80)
        p90 = np.percentile(data_sim, 90)
        
        m1, m2, m3 = st.columns(3)
        m1.metric("P50 (50% Confianza)", f"{p50:.1f} días")
        m2.metric("P80 (80% Confianza)", f"{p80:.1f} días")
        m3.metric("P90 (90% Confianza)", f"{p90:.1f} días")
        
        # Gráfico Plotly
        fig_mc = px.histogram(
            x=data_sim, 
            nbins=30, 
            title="Distribución de Duración Estimada del Proyecto",
            labels={'x': 'Días', 'count': 'Frecuencia'},
            color_discrete_sequence=['#e09f9f']
        )
        fig_mc.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#f4ecec'),
            bargap=0.05
        )
        st.plotly_chart(fig_mc, use_container_width=True)

# --- TAB 5: DIAGRAMAS DE FLUJO ---
with tab5:
    st.markdown("### Diagramas de Flujo del Sistema")
    st.write("Selecciona una plantilla o escribe tu diagrama personalizado.")
    
    plantilla = st.selectbox(
        "Cargar Plantilla",
        ["Proceso de Autenticación", "Flujo de Datos (DFD)", "Diagrama de Secuencia API"]
    )
    
    if plantilla == "Proceso de Autenticación":
        code_flujo = "graph TD\n  A[Usuario] -->|Credenciales| B(API Login)\n  B -->|Validar| C{¿Correcto?}\n  C -- Sí --> D[Generar JWT Token]\n  C -- No --> E[Error 401 Unauthorized]"
    elif plantilla == "Flujo de Datos (DFD)":
        code_flujo = "graph LR\n  Cliente -->|Request POST| Router\n  Router --> Controller\n  Controller -->|Query| DB[(Base de Datos)]\n  DB -->|Respuesta| Controller\n  Controller -->|JSON| Cliente"
    else:
        code_flujo = "sequenceDiagram\n  autonumber\n  Cliente->>Servidor: POST /login\n  Servidor-->>BaseDeDatos: Consulta Usuario\n  BaseDeDatos-->>Servidor: Datos Ok\n  Servidor-->>Cliente: 200 OK + Token"

    flujo_editado = st.text_area("Código Mermaid", value=code_flujo, height=140)
    st_mermaid(flujo_editado)

# --- TAB 6: EDT (WBS) ---
with tab6:
    st.markdown("### Estructura de Descomposición del Trabajo (EDT / WBS)")
    
    with st.form("form_edt", clear_on_submit=True):
        col_f, col_p, col_h, col_b = st.columns([2, 3, 1.5, 1])
        fase_in = col_f.text_input("Fase", placeholder="Ej: 1. Diseño")
        paq_in = col_p.text_input("Paquete", placeholder="Ej: Diagramas ER")
        hrs_in = col_h.number_input("Horas", min_value=1, value=10)
        btn_edt = col_b.form_submit_button("Añadir")
        
        if btn_edt and fase_in.strip() and paq_in.strip():
            st.session_state.edt_list.append({"Fase": fase_in.strip(), "Paquete": paq_in.strip(), "Horas": hrs_in})
            st.rerun()

    df_edt = pd.DataFrame(st.session_state.edt_list)
    if not df_edt.empty:
        # Visualización Treemap
        fig_wbs = px.treemap(
            df_edt, 
            path=['Fase', 'Paquete'], 
            values='Horas',
            title="Distribución del Esfuerzo por Fases y Paquetes (Horas)",
            color_discrete_sequence=['#8a2d3b', '#e09f9f', '#5c1e28']
        )
        fig_wbs.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#f4ecec')
        )
        st.plotly_chart(fig_wbs, use_container_width=True)
        
        # Eliminar items de la EDT
        edt_del = None
        for i, r in enumerate(st.session_state.edt_list):
            c1, c2, c3, c4 = st.columns([2, 3, 1, 0.8])
            c1.write(r["Fase"])
            c2.write(r["Paquete"])
            c3.write(f"{r['Horas']} hrs")
            if c4.button("🗑️", key=f"del_edt_{i}"):
                edt_del = i
        if edt_del is not None:
            st.session_state.edt_list.pop(edt_del)
            st.rerun()

# --- TAB 7: RIESGOS ---
with tab7:
    st.markdown("### Matriz de Evaluación de Riesgos")
    
    with st.form("form_riesgos", clear_on_submit=True):
        col_r, col_prob, col_imp, col_btn_r = st.columns([3, 1.5, 1.5, 1])
        r_txt = col_r.text_input("Riesgo", placeholder="Descripción del riesgo...")
        r_prob = col_prob.slider("Prob (1-5)", 1, 5, 3)
        r_imp = col_imp.slider("Imp (1-5)", 1, 5, 3)
        btn_add_r = col_btn_r.form_submit_button("Añadir")
        
        if btn_add_r and r_txt.strip():
            st.session_state.riesgos_list.append({"Riesgo": r_txt.strip(), "Probabilidad": r_prob, "Impacto": r_imp})
            st.rerun()

    df_r = pd.DataFrame(st.session_state.riesgos_list)
    if not df_r.empty:
        df_r["Severidad"] = df_r["Probabilidad"] * df_r["Impacto"]
        
        # Plot Scatter Matriz de Riesgos
        fig_r = px.scatter(
            df_r, 
            x="Impacto", 
            y="Probabilidad", 
            size="Severidad", 
            text="Riesgo",
            title="Matriz Impacto vs Probabilidad",
            color="Severidad",
            color_continuous_scale=['#5c1e28', '#e09f9f', '#ff4d4d']
        )
        fig_r.update_layout(
            xaxis=dict(range=[0, 6], dtick=1),
            yaxis=dict(range=[0, 6], dtick=1),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#f4ecec')
        )
        st.plotly_chart(fig_r, use_container_width=True)

        # Eliminar ítems
        r_del = None
        for i, r in enumerate(st.session_state.riesgos_list):
            c1, c2, c3, c4 = st.columns([3, 1, 1, 0.8])
            c1.write(f"**{r['Riesgo']}**")
            c2.write(f"P: {r['Probabilidad']}")
            c3.write(f"I: {r['Impacto']}")
            if c4.button("🗑️", key=f"del_risk_{i}"):
                r_del = i
        if r_del is not None:
            st.session_state.riesgos_list.pop(r_del)
            st.rerun()

# --- TAB 8: REPORTE FULL ---
with tab_full:
    st.markdown("### 📦 Exportación General Consolidada")
    st.write("Descarga un paquete (.ZIP) con la documentación completa generada en la plataforma.")
    
    word_doc = generar_word_ers()
    excel_doc = generar_excel_estilizado()
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"ERS_{st.session_state.nombre_proj}.docx", word_doc)
        zip_file.writestr(f"Proyecto_Consolidado_{st.session_state.nombre_proj}.xlsx", excel_doc)
        
        df_kanban = pd.DataFrame(st.session_state.kanban_tasks)
        zip_file.writestr("Kanban_Tareas.csv", df_kanban.to_csv(index=False).encode('utf-8'))

    st.write("") 
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.download_button(
            label="⬇️ Descargar Reporte Full (.ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"Reporte_Full_{st.session_state.nombre_proj.replace(' ', '_')}.zip",
            mime="application/zip",
            use_container_width=True
        )
