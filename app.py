import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from streamlit_mermaid import st_mermaid
import io
import zipfile
import openpyxl

# Librerías para Word
from docx import Document
from docx.shared import Pt, RGBColor

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ==========================================
st.set_page_config(page_title="MatrixDev Tesis", layout="wide", page_icon="🍷")

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

    /* Ocultar elementos nativos innecesarios */
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. PLANTILLAS Y VARIABLES DE ESTADO
# ==========================================
PLANTILLAS_FLUJO = {
    "Autenticación": """graph TD
A["Usuario"] -->|Credenciales| B["API Login"]
B -->|Validar| C{"¿Válido?"}
C -->|Sí| D["Generar Token JWT"]
C -->|No| E["Error 401 Unauthorized"]""",

    "Procesamiento API": """graph LR
Cliente["Cliente"] -->|Request POST| Router["Router"]
Router --> Controller["Controller"]
Controller -->|Query| DB[("Base de Datos")]
DB -->|Respuesta| Controller
Controller -->|JSON| Cliente""",

    "Secuencia de Usuario": """sequenceDiagram
autonumber
actor Cliente
participant Servidor
participant BD as Base de Datos

Cliente->>Servidor: POST /login
Servidor->>BD: Consulta Usuario
BD-->>Servidor: Datos OK
Servidor-->>Cliente: 200 OK + Token""",

    "Crear desde cero": """graph TD
A["Inicio"] --> B["Tu Nuevo Proceso"]
B --> C{"¿Aprobado?"}
C -->|Sí| D["Resultado Éxito"]
C -->|No| E["Resultado Fallo"]"""
}

if "nombre_proj" not in st.session_state:
    st.session_state.nombre_proj = "MatrixDev Core"
if "integrantes" not in st.session_state:
    st.session_state.integrantes = "Juan Pérez"
if "objetivo_text" not in st.session_state:
    st.session_state.objetivo_text = "Automatizar el flujo de inventario con una arquitectura minimalista."

# Listas de datos
if "rf_list" not in st.session_state:
    st.session_state.rf_list = ["Autenticación OAuth2.", "CRUD de usuarios."]
if "rnf_list" not in st.session_state:
    st.session_state.rnf_list = ["Latencia < 200ms.", "Cifrado AES-256 en base de datos."]
if "kanban_tasks" not in st.session_state:
    st.session_state.kanban_tasks = [
        {"Tarea": "Modelo BD", "Estado": "Completado"},
        {"Tarea": "Endpoints API", "Estado": "En Proceso"}
    ]

# Estructura jerárquica multinivel para la EDT
if "edt_list" not in st.session_state:
    st.session_state.edt_list = [
        {
            "Fase": "Fase 1: Preparación",
            "Subfase": "Limpieza y Despeje",
            "Paquete": "Despejar Habitación y Muebles",
            "Horas": 5,
            "Costo": 20000,
            "Documentos": "Plan de Trabajo"
        },
        {
            "Fase": "Fase 1: Preparación",
            "Subfase": "Pintura Aerosol",
            "Paquete": "Pintar Arbusto / Detalles",
            "Horas": 4,
            "Costo": 35000,
            "Documentos": "Ficha Técnica Aerosol"
        },
        {
            "Fase": "Fase 2: Aplicación Manual",
            "Subfase": "Capas de Pintura",
            "Paquete": "Mano de Pintura Muros (1 y 2)",
            "Horas": 16,
            "Costo": 80000,
            "Documentos": "Guía de Pintado Manual"
        },
        {
            "Fase": "Fase 2: Aplicación Manual",
            "Subfase": "Documentación y Entrega",
            "Paquete": "Informe de Inspección y Cotizaciones",
            "Horas": 6,
            "Costo": 15000,
            "Documentos": "Cotización + Certificados"
        }
    ]

if "riesgos_list" not in st.session_state:
    st.session_state.riesgos_list = [
        {"Riesgo": "Retraso en entrega API", "Probabilidad": 3, "Impacto": 4},
        {"Riesgo": "Falla en servidor cloud", "Probabilidad": 2, "Impacto": 5},
        {"Riesgo": "Incompatibilidad de navegador", "Probabilidad": 1, "Impacto": 2}
    ]

if "flujo_codigo" not in st.session_state:
    st.session_state.flujo_codigo = PLANTILLAS_FLUJO["Autenticación"]

def cambiar_plantilla_flujo():
    sel = st.session_state.select_tipo_flujo
    if sel in PLANTILLAS_FLUJO:
        st.session_state.flujo_codigo = PLANTILLAS_FLUJO[sel]

def generar_mermaid_edt(edt_items, nombre_proyecto):
    proj_name = str(nombre_proyecto).replace('"', '').replace("'", "")
    lines = ["graph TD", f'    ROOT["📦 {proj_name}"]']
    
    fases = {}
    for item in edt_items:
        fase = str(item.get("Fase", "Sin Fase")).strip().replace('"', '')
        subfase = str(item.get("Subfase", "")).strip().replace('"', '')
        paquete = str(item.get("Paquete", "Tarea")).strip().replace('"', '')
        horas = float(item.get("Horas", 0))
        costo = float(item.get("Costo", 0))
        doc = str(item.get("Documentos", "")).strip().replace('"', '')
        
        if fase not in fases:
            fases[fase] = {}
        if subfase not in fases[fase]:
            fases[fase][subfase] = []
        
        fases[fase][subfase].append({
            "paquete": paquete,
            "horas": horas,
            "costo": costo,
            "doc": doc
        })

    fase_idx = 0
    for fase_name, subfases in fases.items():
        fase_id = f"F{fase_idx}"
        
        total_fase_hrs = sum(pkg["horas"] for sub in subfases.values() for pkg in sub)
        total_fase_cost = sum(pkg["costo"] for sub in subfases.values() for pkg in sub)
        
        fase_label = f"📂 {fase_name}<br/><i>Total: {total_fase_hrs:.0f}h | ${total_fase_cost:,.0f}</i>"
        lines.append(f'    ROOT --> {fase_id}["{fase_label}"]')
        
        sub_idx = 0
        for sub_name, pkgs in subfases.items():
            if sub_name:
                sub_id = f"{fase_id}_S{sub_idx}"
                sub_hrs = sum(pkg["horas"] for pkg in pkgs)
                sub_cost = sum(pkg["costo"] for pkg in pkgs)
                sub_label = f"📁 {sub_name}<br/><i>{sub_hrs:.0f}h | ${sub_cost:,.0f}</i>"
                lines.append(f'    {fase_id} --> {sub_id}["{sub_label}"]')
                parent_id = sub_id
            else:
                parent_id = fase_id
            
            for pkg_idx, pkg in enumerate(pkgs):
                pkg_id = f"{parent_id}_P{pkg_idx}"
                details = [f"⏱️ {pkg['horas']:.0f} hrs"]
                if pkg['costo'] > 0:
                    details.append(f"💰 ${pkg['costo']:,.0f}")
                if pkg['doc']:
                    details.append(f"📄 {pkg['doc']}")
                
                details_str = "<br/>".join(details)
                lines.append(f'    {parent_id} --> {pkg_id}["<b>{pkg["paquete"]}</b><br/>{details_str}"]')
            
            sub_idx += 1
        fase_idx += 1

    return "\n".join(lines)

# ==========================================
# 3. FUNCIONES DE GENERACIÓN DE DOCUMENTOS
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
    
    # Hoja 1: Requisitos
    ws1 = wb.active
    ws1.title = "Requisitos"
    ws1.append(["Código", "Tipo", "Descripción"])
    for idx, rf in enumerate(st.session_state.rf_list, 1):
        if rf.strip():
            ws1.append([f"RF-{idx:02d}", "Funcional", rf.strip()])
    for idx, rnf in enumerate(st.session_state.rnf_list, 1):
        if rnf.strip():
            ws1.append([f"RNF-{idx:02d}", "No Funcional", rnf.strip()])

    # Hoja 2: Kanban
    ws2 = wb.create_sheet(title="Kanban")
    ws2.append(["Tarea", "Estado"])
    for k in st.session_state.kanban_tasks:
        ws2.append([k["Tarea"], k["Estado"]])

    # Hoja 3: EDT (WBS Ampliada)
    ws3 = wb.create_sheet(title="EDT (WBS)")
    ws3.append(["Fase Principal", "Subfase / Grupo", "Paquete de Trabajo", "Horas Estimadas", "Costo Estimado", "Documentos / Adjuntos"])
    for e in st.session_state.edt_list:
        ws3.append([
            e.get("Fase", ""),
            e.get("Subfase", ""),
            e.get("Paquete", ""),
            e.get("Horas", 0),
            e.get("Costo", 0),
            e.get("Documentos", "")
        ])

    wb.save(output)
    return output.getvalue()

# ==========================================
# 4. INTERFAZ PRINCIPAL
# ==========================================
st.title("✦ MatrixDev ✦")
st.markdown("<p style='text-align: center; color: #a68a8d;'>Gestión y Arquitectura de Proyectos</p>", unsafe_allow_html=True)
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

# ------------------------------------------
# TAB 1: REQUISITOS
# ------------------------------------------
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
            btn_add_rf = st.form_submit_button("Añadir", use_container_width=True)
            
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
            btn_add_rnf = st.form_submit_button("Añadir", use_container_width=True)
            
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

# ------------------------------------------
# TAB 2: MODELADO
# ------------------------------------------
with tab2:
    st.markdown("### Arquitectura Visual")
    default_mermaid = """graph TD
    A["Inicio"] --> B{"Validar"}
    B -->|Sí| C["Éxito"]
    B -->|No| D["Error"]"""
    
    codigo_mermaid = st.text_area("Sintaxis Mermaid", value=default_mermaid, height=120)
    st_mermaid(codigo_mermaid)

# ------------------------------------------
# TAB 3: GESTIÓN
# ------------------------------------------
with tab3:
    st.markdown("### Tablero Kanban")
    
    with st.form("form_kanban", clear_on_submit=True):
        col_t, col_s, col_b = st.columns([3, 2, 1])
        with col_t:
            nueva_t = st.text_input("Tarea", placeholder="Nueva tarea...", label_visibility="collapsed")
        with col_s:
            estado_t = st.selectbox("Estado", ["Pendiente", "En Proceso", "Completado"], label_visibility="collapsed")
        with col_b:
            btn_add_k = st.form_submit_button("Añadir", use_container_width=True)
            
        if btn_add_k and nueva_t.strip():
            st.session_state.kanban_tasks.append({"Tarea": nueva_t.strip(), "Estado": estado_t})
            st.rerun()

    if not st.session_state.kanban_tasks:
        st.info("No hay tareas registradas.")
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

# ------------------------------------------
# TAB 4: MONTE CARLO
# ------------------------------------------
with tab4:
    st.markdown("### Simulación Monte Carlo de Estimación")
    c1, c2, c3 = st.columns(3)
    opt = c1.number_input("Días Optimista", value=10, min_value=1)
    prob = c2.number_input("Días Más Probable", value=20, min_value=1)
    pes = c3.number_input("Días Pesimista", value=45, min_value=1)
    
    simulaciones = st.slider("Número de Simulaciones", 500, 10000, 2000, step=500)
    
    if opt <= prob <= pes:
        datos_sim = np.random.triangular(left=opt, mode=prob, right=pes, size=simulaciones)
        
        fig = px.histogram(datos_sim, nbins=30, title="Distribución de Duración (Días)", labels={'value': 'Días'}, color_discrete_sequence=['#e09f9f'])
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color='#f4ecec')
        st.plotly_chart(fig, use_container_width=True)
        
        p85 = np.percentile(datos_sim, 85)
        st.metric("Estimación con 85% de Confianza", f"{p85:.1f} Días")
    else:
        st.error("Asegúrate de que: Optimista ≤ Más Probable ≤ Pesimista.")

# ------------------------------------------
# TAB 5: FLUJOS
# ------------------------------------------
with tab5:
    st.markdown("### Diagramas de Flujo y Secuencia")
    
    st.selectbox(
        "Seleccionar Plantilla de Flujo", 
        options=list(PLANTILLAS_FLUJO.keys()),
        key="select_tipo_flujo",
        on_change=cambiar_plantilla_flujo
    )
    
    st.session_state.flujo_codigo = st.text_area(
        "Código del Diagrama (Editable)", 
        value=st.session_state.flujo_codigo, 
        height=180,
        key="txt_flujo_code"
    )
    
    st.markdown("#### Vista Previa del Diagrama")
    st_mermaid(st.session_state.flujo_codigo)

# ------------------------------------------
# TAB 6: EDT (WBS) - MULTINIVEL & MINIMALISTA
# ------------------------------------------
with tab6:
    st.markdown("### Estructura de Desglose de Trabajo (EDT / WBS)")
    
    with st.expander("📖 **Guía Rápida EDT: Fases, Agrupaciones y Documentos**", expanded=False):
        st.markdown(f"""
        **Estructura Jerárquica del Proyecto:**
        * 📦 **Proyecto (Nivel 0):** *{st.session_state.nombre_proj}*
        * 📂 **Fase Principal (Nivel 1):** Ej: *Fase 1: Preparación*, *Fase 2: Manual*.
        * 📁 **Subfase / Grupo (Nivel 2):** Agrupación secundaria opcional (Ej: *Aerosol*, *Documentación*).
        * 📄 **Paquete de Trabajo (Nivel 3):** Tarea entregable con **Horas**, **Costo** y **Documentos**.
        """)

    df_edt = pd.DataFrame(st.session_state.edt_list)
    
    # MÉTRICAS GLOBALES
    if not df_edt.empty:
        total_horas = df_edt["Horas"].sum()
        total_costo = df_edt["Costo"].sum() if "Costo" in df_edt.columns else 0
        total_paquetes = len(df_edt)
        total_fases = df_edt["Fase"].nunique()
    else:
        total_horas = 0
        total_costo = 0
        total_paquetes = 0
        total_fases = 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("⏳ Total Horas", f"{total_horas:.0f} hrs")
    m2.metric("💰 Costo Total Estimado", f"${total_costo:,.0f}")
    m3.metric("📄 Paquetes de Trabajo", total_paquetes)
    m4.metric("📂 Fases Principales", total_fases)
    
    st.write("---")
    
    # DIAGRAMA DE ÁRBOL JERÁRQUICO (WBS VISUAL EN MERMAID)
    st.markdown("#### 🌳 Diagrama Jerárquico Multinivel (EDT)")
    if not df_edt.empty:
        mermaid_edt_code = generar_mermaid_edt(st.session_state.edt_list, st.session_state.nombre_proj)
        st_mermaid(mermaid_edt_code)
    else:
        st.info("No hay paquetes agregados para mostrar el árbol jerárquico.")

    st.write("---")

    # DISTRIBUCIÓN DE HORAS Y COSTOS Y FORMULARIO
    col_graph, col_manage = st.columns([1.1, 1])

    with col_graph:
        st.markdown("#### 📊 Distribución de Recursos por Fase")
        if not df_edt.empty and "Costo" in df_edt.columns:
            df_fases = df_edt.groupby("Fase", as_index=False)[["Horas", "Costo"]].sum()
            
            tipo_metric = st.radio("Métrica visualizada", ["Horas", "Costo"], horizontal=True, key="rad_edt_metric")
            
            fig_donut = px.pie(
                df_fases, 
                names='Fase', 
                values=tipo_metric, 
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Dark24
            )
            fig_donut.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', 
                plot_bgcolor='rgba(0,0,0,0)', 
                font_color='#f4ecec',
                margin=dict(l=10, r=10, t=30, b=10)
            )
            st.plotly_chart(fig_donut, use_container_width=True)
        else:
            st.info("Agrega datos para generar el gráfico de distribución.")

    with col_manage:
        st.markdown("#### ➕ Agregar Paquete de Trabajo / Tarea")
        with st.form("form_add_edt_multi", clear_on_submit=True):
            fase_in = st.text_input("Fase Principal", placeholder="Ej: Fase 1: Preparación")
            subfase_in = st.text_input("Subfase / Grupo (Opcional)", placeholder="Ej: Pintura Aerosol")
            paquete_in = st.text_input("Paquete de Trabajo / Tarea", placeholder="Ej: Pintar arbusto")
            
            c_h, c_c = st.columns(2)
            horas_in = c_h.number_input("Horas Estimadas", min_value=0, value=4, step=1)
            costo_in = c_c.number_input("Costo Estimado ($)", min_value=0, value=15000, step=1000)
            
            doc_in = st.text_input("Documentos / Referencias", placeholder="Ej: Ficha Técnica Aerosol")
            
            btn_add_edt = st.form_submit_button("Añadir Paquete a EDT", use_container_width=True)
            if btn_add_edt:
                if fase_in.strip() and paquete_in.strip():
                    st.session_state.edt_list.append({
                        "Fase": fase_in.strip(),
                        "Subfase": subfase_in.strip(),
                        "Paquete": paquete_in.strip(),
                        "Horas": float(horas_in),
                        "Costo": float(costo_in),
                        "Documentos": doc_in.strip()
                    })
                    st.rerun()
                else:
                    st.error("Debes completar al menos la Fase Principal y el Paquete de Trabajo.")

    st.write("---")
    st.markdown("#### 📝 Lista Dinámica y Edición de EDT")
    
    if not st.session_state.edt_list:
        st.info("La estructura de desglose de trabajo está vacía.")
    else:
        edt_del_idx = None
        
        # Cabecera de la tabla editable
        h1, h2, h3, h4, h5, h6, h7 = st.columns([2, 1.8, 2.2, 1, 1.2, 1.8, 0.6])
        h1.markdown("**Fase**")
        h2.markdown("**Subfase**")
        h3.markdown("**Paquete / Tarea**")
        h4.markdown("**Horas**")
        h5.markdown("**Costo ($)**")
        h6.markdown("**Documentos**")
        h7.markdown("")
        
        for i, item in enumerate(st.session_state.edt_list):
            col_f, col_sf, col_p, col_h, col_c, col_d, col_del = st.columns([2, 1.8, 2.2, 1, 1.2, 1.8, 0.6])
            
            f_val = col_f.text_input(f"f_{i}", value=item.get("Fase", ""), key=f"edt_fase_{i}", label_visibility="collapsed")
            sf_val = col_sf.text_input(f"sf_{i}", value=item.get("Subfase", ""), key=f"edt_sfase_{i}", label_visibility="collapsed")
            p_val = col_p.text_input(f"p_{i}", value=item.get("Paquete", ""), key=f"edt_pkg_{i}", label_visibility="collapsed")
            h_val = col_h.number_input(f"h_{i}", value=float(item.get("Horas", 0)), min_value=0.0, step=1.0, key=f"edt_hrs_{i}", label_visibility="collapsed")
            c_val = col_c.number_input(f"c_{i}", value=float(item.get("Costo", 0)), min_value=0.0, step=1000.0, key=f"edt_cost_{i}", label_visibility="collapsed")
            doc_val = col_d.text_input(f"d_{i}", value=item.get("Documentos", ""), key=f"edt_doc_{i}", label_visibility="collapsed")
            
            st.session_state.edt_list[i]["Fase"] = f_val
            st.session_state.edt_list[i]["Subfase"] = sf_val
            st.session_state.edt_list[i]["Paquete"] = p_val
            st.session_state.edt_list[i]["Horas"] = h_val
            st.session_state.edt_list[i]["Costo"] = c_val
            st.session_state.edt_list[i]["Documentos"] = doc_val
            
            if col_del.button("🗑️", key=f"del_edt_{i}"):
                edt_del_idx = i

        if edt_del_idx is not None:
            st.session_state.edt_list.pop(edt_del_idx)
            st.rerun()

# ------------------------------------------
# TAB 7: RIESGOS
# ------------------------------------------
with tab7:
    st.markdown("### Matriz de Riesgos (Impacto vs Probabilidad)")
    
    df_r = pd.DataFrame(st.session_state.riesgos_list)
    if not df_r.empty:
        df_r["Severidad"] = df_r["Probabilidad"] * df_r["Impacto"]
        
        fig_r = px.scatter(
            df_r, 
            x="Probabilidad", 
            y="Impacto", 
            text="Riesgo", 
            size="Severidad",
            color="Severidad",
            color_continuous_scale="Reds",
            title="Mapa de Calor de Riesgos"
        )
        fig_r.update_layout(
            xaxis=dict(range=[0, 6]), 
            yaxis=dict(range=[0, 6]),
            paper_bgcolor='rgba(0,0,0,0)', 
            plot_bgcolor='rgba(0,0,0,0)', 
            font_color='#f4ecec'
        )
        st.plotly_chart(fig_r, use_container_width=True)
        
    st.dataframe(df_r, use_container_width=True)

# ------------------------------------------
# TAB 8: REPORTE FULL
# ------------------------------------------
with tab_full:
    st.markdown("### 📦 Exportación General Consolidada")
    st.write("Descarga un archivo ZIP con todos los entregables generados (Word, Excel y CSVs).")
    
    word_doc = generar_word_ers()
    excel_doc = generar_excel_estilizado()
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"ERS_{st.session_state.nombre_proj}.docx", word_doc)
        zip_file.writestr(f"Matrices_{st.session_state.nombre_proj}.xlsx", excel_doc)
        
        df_kanban = pd.DataFrame(st.session_state.kanban_tasks)
        zip_file.writestr("Kanban_Tareas.csv", df_kanban.to_csv(index=False).encode('utf-8'))
        
        df_edt_export = pd.DataFrame(st.session_state.edt_list)
        zip_file.writestr("EDT_WBS.csv", df_edt_export.to_csv(index=False).encode('utf-8'))

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
