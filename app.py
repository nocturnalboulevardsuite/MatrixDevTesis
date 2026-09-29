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

# Librerías para generación de Word (.docx)
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="MatrixDevTesis", layout="wide", page_icon="🎓")

# ==========================================
# INICIALIZACIÓN DE VARIABLES DE ESTADO
# ==========================================
if "nombre_proj" not in st.session_state:
    st.session_state.nombre_proj = "Sistema de Control de Inventario MatrixDev"
if "integrantes" not in st.session_state:
    st.session_state.integrantes = "Juan Pérez, María González"
if "profesor" not in st.session_state:
    st.session_state.profesor = "Dr. Roberto Gómez"
if "fecha" not in st.session_state:
    st.session_state.fecha = "10 de Septiembre de 2026"
if "seccion" not in st.session_state:
    st.session_state.seccion = "Sección 1 - Taller de Proyecto de Título"
if "objetivo_text" not in st.session_state:
    st.session_state.objetivo_text = "Automatizar el flujo de inventario y optimizar la generación de reportes universitarios."
if "mvp_text" not in st.session_state:
    st.session_state.mvp_text = "Módulo de autenticación, gestión CRUD de productos y exportación del documento ERS."
if "df_rf" not in st.session_state:
    st.session_state.df_rf = pd.DataFrame([
        {"Descripción": "Autenticación con credenciales universitarias."},
        {"Descripción": "Registro y edición de tareas."},
        {"Descripción": "Exportación en formato Markdown."}
    ])
if "df_rnf" not in st.session_state:
    st.session_state.df_rnf = pd.DataFrame([
        {"Descripción": "Tiempo de respuesta menor a 1.5 segundos."},
        {"Descripción": "Cifrado SSL en todas las peticiones."}
    ])
if "user_stories" not in st.session_state:
    st.session_state.user_stories = pd.DataFrame([
        {"ID": "US-01", "Como": "Estudiante", "Quiero": "Generar mi documento ERS en un clic", "Para": "Entregarlo en la memoria de título"},
        {"ID": "US-02", "Como": "Profesor Evaluador", "Quiero": "Visualizar la ruta crítica", "Para": "Analizar la factibilidad del proyecto"}
    ])
if "kanban_tasks" not in st.session_state:
    st.session_state.kanban_tasks = pd.DataFrame([
        {"Tarea": "Diseñar Modelo de Datos", "Estado": "Completado", "Asignado": "Ana", "Prioridad": "Alta"},
        {"Tarea": "Endpoints API Rest", "Estado": "En Proceso", "Asignado": "Carlos", "Prioridad": "Alta"},
        {"Tarea": "Pruebas de Integración", "Estado": "Pendiente", "Asignado": "Beatriz", "Prioridad": "Media"}
    ])
if "df_mc_tasks" not in st.session_state:
    st.session_state.df_mc_tasks = pd.DataFrame([
        {"Tarea": "Análisis y ERS", "Optimista": 3, "Mas_Probable": 5, "Pesimista": 10},
        {"Tarea": "Modelado UML y Base de Datos", "Optimista": 4, "Mas_Probable": 7, "Pesimista": 12},
        {"Tarea": "Desarrollo Backend & Frontend", "Optimista": 10, "Mas_Probable": 15, "Pesimista": 25},
        {"Tarea": "Pruebas y Despliegue", "Optimista": 3, "Mas_Probable": 5, "Pesimista": 8}
    ])

# ==========================================
# FUNCIONES DE GENERACIÓN DE DOCUMENTOS
# ==========================================
def generar_word_ers(nombre_proyecto, integrantes, profesor, fecha, seccion, objetivo, alcance, df_rf, df_rnf, df_us):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(3)
        sec.bottom_margin = Cm(3)
        sec.left_margin = Cm(3)
        sec.right_margin = Cm(3)

    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Calibri'
    style_normal.font.size = Pt(11)
    style_normal.paragraph_format.line_spacing = 1.5
    style_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    style_normal.paragraph_format.space_after = Pt(6)

    def set_cell_background(cell, fill_hex):
        tcPr = cell._element.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        tcPr.append(shd)

    # Portada
    p_top_space = doc.add_paragraph()
    p_top_space.paragraph_format.space_before = Pt(50)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("ESPECIFICACIÓN DE REQUISITOS DE SOFTWARE\n(ERS)")
    r_title.bold = True
    r_title.font.size = Pt(22)
    r_title.font.color.rgb = RGBColor(31, 78, 120)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(f"\nPROYECTO: {nombre_proyecto.upper()}")
    r_sub.bold = True
    r_sub.font.size = Pt(14)
    r_sub.font.color.rgb = RGBColor(89, 89, 89)

    doc.add_paragraph().paragraph_format.space_before = Pt(180)
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    datos_portada = [
        ("Nombre del Proyecto:", nombre_proyecto),
        ("Integrante/s:", integrantes),
        ("Profesor/a o Responsable:", profesor),
        ("Fecha:", fecha),
        ("Sección:", seccion)
    ]
    
    for campo, val in datos_portada:
        r_c = p_meta.add_run(f"• {campo} ")
        r_c.bold = True
        r_v = p_meta.add_run(f"{val}\n")

    doc.add_page_break()

    def agregar_encabezado_seccion(texto, nivel=1, color=RGBColor(31, 78, 120)):
        h = doc.add_heading(level=nivel)
        r = h.add_run(texto)
        r.bold = True
        r.font.name = 'Calibri'
        r.font.color.rgb = color
        return h

    # Secciones
    agregar_encabezado_seccion("1. Objetivo Principal del Sistema")
    doc.add_paragraph(objetivo)

    agregar_encabezado_seccion("2. Alcance del Producto Mínimo Viable (MVP)")
    doc.add_paragraph(alcance)

    agregar_encabezado_seccion("3. Requerimientos Funcionales (RF)")
    table_rf = doc.add_table(rows=1, cols=2)
    table_rf.style = 'Table Grid'
    hdr_rf = table_rf.rows[0].cells
    hdr_rf[0].text = "ID"
    hdr_rf[1].text = "Descripción del Requerimiento Funcional"
    
    for cell in hdr_rf:
        set_cell_background(cell, "1F4E78")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)

    for _, row in df_rf.iterrows():
        row_cells = table_rf.add_row().cells
        row_cells[0].text = str(row["ID"])
        row_cells[1].text = str(row["Descripción"])

    doc.add_paragraph()

    agregar_encabezado_seccion("4. Requerimientos No Funcionales (RNF)", color=RGBColor(192, 0, 0))
    table_rnf = doc.add_table(rows=1, cols=2)
    table_rnf.style = 'Table Grid'
    hdr_rnf = table_rnf.rows[0].cells
    hdr_rnf[0].text = "ID"
    hdr_rnf[1].text = "Descripción del Requerimiento No Funcional"
    
    for cell in hdr_rnf:
        set_cell_background(cell, "C00000")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)

    for _, row in df_rnf.iterrows():
        row_cells = table_rnf.add_row().cells
        row_cells[0].text = str(row["ID"])
        row_cells[1].text = str(row["Descripción"])

    doc.add_paragraph()

    agregar_encabezado_seccion("5. Historias de Usuario (User Stories)")
    table_us = doc.add_table(rows=1, cols=4)
    table_us.style = 'Table Grid'
    hdr_us = table_us.rows[0].cells
    hdr_us[0].text = "ID"
    hdr_us[1].text = "Como..."
    hdr_us[2].text = "Quiero..."
    hdr_us[3].text = "Para..."
    
    for cell in hdr_us:
        set_cell_background(cell, "333333")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)

    for _, row in df_us.iterrows():
        row_cells = table_us.add_row().cells
        row_cells[0].text = str(row.get("ID", ""))
        row_cells[1].text = str(row.get("Como", ""))
        row_cells[2].text = str(row.get("Quiero", ""))
        row_cells[3].text = str(row.get("Para", ""))

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

def generar_excel_estilizado(df_rf, df_rnf, nombre_proyecto):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    thin_border = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'), top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))
    
    def aplicar_estilo_hoja(ws, df, titulo_hoja, color_principal, color_suave):
        ws.views.sheetView[0].showGridLines = True
        ws.merge_cells("A1:B1")
        title_cell = ws["A1"]
        title_cell.value = f"📌 {titulo_hoja.upper()}"
        title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color=color_principal, end_color=color_principal, fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 35
        
        headers = ["ID", "Descripción del Requerimiento"]
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=color_principal, end_color=color_principal, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        for r_idx, row in df.iterrows():
            row_num = r_idx + 4
            bg_color = color_suave if r_idx % 2 == 1 else "FFFFFF"
            fill_zebra = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            
            ws.cell(row=row_num, column=1, value=row["ID"]).fill = fill_zebra
            ws.cell(row=row_num, column=2, value=row["Descripción"]).fill = fill_zebra

        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 80

    ws_rf = wb.active
    ws_rf.title = "Requerimientos Funcionales"
    aplicar_estilo_hoja(ws_rf, df_rf, "Requerimientos Funcionales", "1F4E78", "F2F5F9")
    
    ws_rnf = wb.create_sheet(title="Requerimientos No Funcionales")
    aplicar_estilo_hoja(ws_rnf, df_rnf, "Requerimientos No Funcionales", "C00000", "FDF2F2")

    wb.save(output)
    return output.getvalue()

def simular_montecarlo(df_tasks, N=2500):
    np.random.seed(42)
    total_duraciones = np.zeros(N)
    for _, row in df_tasks.iterrows():
        o, m, p = float(row["Optimista"]), float(row["Mas_Probable"]), float(row["Pesimista"])
        low, high = min(o, m, p), max(o, m, p)
        mode = max(low, min(m, high))
        total_duraciones += np.random.triangular(left=low, mode=mode, right=high, size=N)
    return total_duraciones

def generar_excel_montecarlo(df_tasks, duraciones, n_sim, nombre_p):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen Montecarlo"
    ws.cell(row=1, column=1, value=f"REPORTE MONTECARLO - {nombre_p}").font = Font(bold=True)
    ws.cell(row=3, column=1, value="Métrica").font = Font(bold=True)
    ws.cell(row=3, column=2, value="Valor (Días)").font = Font(bold=True)
    
    metricas = [
        ("Promedio Esperado", float(np.mean(duraciones))),
        ("P50 (Mediana)", float(np.percentile(duraciones, 50))),
        ("P80 (Recomendado)", float(np.percentile(duraciones, 80))),
        ("P90 (Conservador)", float(np.percentile(duraciones, 90)))
    ]
    
    for idx, (m_nombre, m_val) in enumerate(metricas, 4):
        ws.cell(row=idx, column=1, value=m_nombre)
        ws.cell(row=idx, column=2, value=round(m_val, 2))

    wb.save(output)
    return output.getvalue()

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.title("🎓 MatrixDevTesis")
st.caption("Suite web all-in-one para la gestión, modelado y documentación de proyectos informáticos.")

tab1, tab2, tab3, tab4 = st.tabs([
    " Requisitos & ERS",
    " Diagramas Visuales",
    " Gestión Estilo Jira & RACI",
    " Analítica, Ruta Crítica & Montecarlo"
])

# --- TAB 1: REQUISITOS ---
with tab1:
    st.subheader(" Documentación Base, Requerimientos e Historias de Usuario")
    
    col_acta1, col_acta2 = st.columns([1, 1])
    with col_acta1:
        st.write("**Metadatos para Portada del Documento ERS**")
        st.session_state.nombre_proj = st.text_input("Nombre del Proyecto", st.session_state.nombre_proj)
        
        col_meta1, col_meta2 = st.columns(2)
        with col_meta1:
            st.session_state.integrantes = st.text_input("Integrante/s", st.session_state.integrantes)
            st.session_state.fecha = st.text_input("Fecha", st.session_state.fecha)
        with col_meta2:
            st.session_state.profesor = st.text_input("Profesor/a o Responsable", st.session_state.profesor)
            st.session_state.seccion = st.text_input("Sección / Asignatura", st.session_state.seccion)

        st.write("**Objetivo Principal del Sistema**")
        st.session_state.objetivo_text = st.text_area("Objetivo", value=st.session_state.objetivo_text, height=90, label_visibility="collapsed")
        
        st.write("**Alcance MVP (Producto Mínimo Viable)**")
        st.session_state.mvp_text = st.text_area("Alcance", value=st.session_state.mvp_text, height=90, label_visibility="collapsed")

    with col_acta2:
        st.write("**Requisitos Funcionales (RF)** — *Añade filas al final*")
        rf_edited = st.data_editor(st.session_state.df_rf, num_rows="dynamic", use_container_width=True, key="rf_editor")
        
        st.write("**Requisitos No Funcionales (RNF)** — *Añade filas al final*")
        rnf_edited = st.data_editor(st.session_state.df_rnf, num_rows="dynamic", use_container_width=True, key="rnf_editor")

    # Autonumeración de IDs
    df_rf_final = rf_edited.dropna(subset=["Descripción"]).reset_index(drop=True)
    df_rf_final["ID"] = [f"RF{i+1:02d}" for i in range(len(df_rf_final))]
    df_rf_final = df_rf_final[["ID", "Descripción"]]
    st.session_state.df_rf = df_rf_final

    df_rnf_final = rnf_edited.dropna(subset=["Descripción"]).reset_index(drop=True)
    df_rnf_final["ID"] = [f"RNF{i+1:02d}" for i in range(len(df_rnf_final))]
    df_rnf_final = df_rnf_final[["ID", "Descripción"]]
    st.session_state.df_rnf = df_rnf_final

    st.markdown("---")
    st.write("**Historias de Usuario (User Stories)**")
    us_df = st.data_editor(st.session_state.user_stories, num_rows="dynamic", use_container_width=True, key="us_editor")
    st.session_state.user_stories = us_df

    st.markdown("---")
    st.subheader("📥 Exportación Individual de Entregables")
    
    word_data = generar_word_ers(
        st.session_state.nombre_proj, st.session_state.integrantes, st.session_state.profesor,
        st.session_state.fecha, st.session_state.seccion, st.session_state.objetivo_text, 
        st.session_state.mvp_text, df_rf_final, df_rnf_final, us_df
    )
    excel_data = generar_excel_estilizado(df_rf_final, df_rnf_final, st.session_state.nombre_proj)

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button("📝 Descargar Documento ERS en WORD (.docx)", data=word_data, file_name=f"ERS_{st.session_state.nombre_proj.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with col_dl2:
        st.download_button("📊 Descargar Excel de Requisitos (.xlsx)", data=excel_data, file_name=f"Requisitos_{st.session_state.nombre_proj.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# --- TAB 2: MERMAID ---
with tab2:
    st.subheader("📐 Diagramado & Modelado (Flowchart, UML, EDT)")
    plantilla = st.selectbox("Cargar Plantilla de Ejemplo:", ["Diagrama de Flujo (Flowchart)", "Diagrama de Clases (UML)", "Estructura de Desglose de Trabajo (EDT / WBS)"])
    
    if plantilla == "Diagrama de Flujo (Flowchart)":
        default_mermaid = "graph TD\n A[Inicio] --> B{¿Usuario registrado?}\n B -- Sí --> C[Login Exitoso]\n B -- No --> D[Formulario de Registro]\n D --> C\n C --> E[Acceso al Panel MatrixDev]"
    elif plantilla == "Diagrama de Clases (UML)":
        default_mermaid = "classDiagram\n class Usuario {\n +String nombre\n +String email\n +login()\n }\n class Proyecto {\n +String nombre\n +List tareas\n +exportarERS()\n }\n Usuario \"1\" -- \"*\" Proyecto : administra"
    else:
        default_mermaid = "graph TD\n A[Proyecto MatrixDevTesis] --> B[1.0 Ingeniería de Requisitos]\n A --> C[2.0 Desarrollo Software]\n B --> B1[1.1 Levantamiento RF/RNF]\n B --> B2[1.2 Diagramas UML]\n C --> C1[2.1 Frontend Streamlit]\n C --> C2[2.2 Lógica en Python]"

    codigo_mermaid = st.text_area("Código en sintaxis Mermaid.js", value=default_mermaid, height=180)
    st_mermaid(codigo_mermaid)

# --- TAB 3: KANBAN & RACI ---
with tab3:
    st.subheader("📌 Tablero de Tareas, Matriz RACI y Gestión de Riesgos")
    col_k1, col_k2 = st.columns([2, 1])
    with col_k1:
        st.write("**Tablero Kanban (Estilo Jira)**")
        kanban_df = st.data_editor(st.session_state.kanban_tasks, num_rows="dynamic", use_container_width=True, column_config={"Estado": st.column_config.SelectboxColumn(options=["Pendiente", "En Proceso", "Completado"])})
        st.session_state.kanban_tasks = kanban_df

    with col_k2:
        st.write("**Matriz de Riesgos**")
        st.dataframe(pd.DataFrame({"Riesgo": ["Retraso en Entregables", "Falla de Integración API"], "Impacto": ["Alto", "Medio"], "Estrategia Mitigación": ["Ajustar alcance MVP", "Crear respuestas Mock"]}), use_container_width=True)

    st.markdown("---")
    st.write("**Matriz RACI (Asignación de Responsabilidades)**")
    st.dataframe(pd.DataFrame({
        "Entregable / Fase": ["Acta de Constitución", "Diagramas UML", "Desarrollo Backend", "Pruebas QA"],
        "Líder de Proyecto": ["A", "R", "I", "C"],
        "Desarrollador": ["C", "C", "R", "I"],
        "Tester / QA": ["I", "I", "C", "R"]
    }), use_container_width=True)

# --- TAB 4: MONTECARLO ---
with tab4:
    st.subheader("📈 Métricas del Proyecto y Simulación Montecarlo")
    
    col_mc_inputs, col_mc_params = st.columns([2, 1])
    with col_mc_inputs:
        st.write("**Definición de Tareas y Estimación PERT (Días)**")
        mc_tasks_edited = st.data_editor(st.session_state.df_mc_tasks, num_rows="dynamic", use_container_width=True, key="mc_tasks_editor")
        st.session_state.df_mc_tasks = mc_tasks_edited

    with col_mc_params:
        st.write("**Configuración de Simulación**")
        n_simulaciones = st.slider("Número de Simulaciones", 500, 10000, 2500, 500)
        st.button(" Ejecutar Simulación Montecarlo", use_container_width=True)

    duraciones_simuladas = simular_montecarlo(mc_tasks_edited, n_simulaciones)
    p50, p80, p90 = np.percentile(duraciones_simuladas, 50), np.percentile(duraciones_simuladas, 80), np.percentile(duraciones_simuladas, 90)

    mc_col1, mc_col2, mc_col3, mc_col4 = st.columns(4)
    mc_col1.metric("Duración Promedio", f"{np.mean(duraciones_simuladas):.1f} Días")
    mc_col2.metric("P50 (Probabilidad 50%)", f"{p50:.1f} Días")
    mc_col3.metric("P80 (Recomendado 80%)", f"{p80:.1f} Días", delta=f"+{p80-p50:.1f}d")
    mc_col4.metric("P90 (Conservador 90%)", f"{p90:.1f} Días")

    fig_mc = go.Figure(data=[go.Histogram(x=duraciones_simuladas, marker_color='#5B2C6F')])
    st.plotly_chart(fig_mc, use_container_width=True)

    excel_mc_data = generar_excel_montecarlo(mc_tasks_edited, duraciones_simuladas, n_simulaciones, st.session_state.nombre_proj)
    st.download_button("📥 Descargar Reporte de Montecarlo (.xlsx)", data=excel_mc_data, file_name=f"Montecarlo_{st.session_state.nombre_proj.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# ==========================================
# BARRA LATERAL: DESCARGA MASIVA EN ZIP
# ==========================================
with st.sidebar:
    st.title("📦 Exportación General")
    st.write("Genera un paquete consolidado con todos los informes generados en tu sesión actual.")
    
    # Crear el archivo ZIP en memoria
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        # Añadir Word ERS
        zip_file.writestr(f"ERS_{st.session_state.nombre_proj}.docx", word_data)
        # Añadir Excel Requisitos
        zip_file.writestr(f"Requisitos_{st.session_state.nombre_proj}.xlsx", excel_data)
        # Añadir Excel Montecarlo
        zip_file.writestr(f"Montecarlo_{st.session_state.nombre_proj}.xlsx", excel_mc_data)
    
    st.download_button(
        label="⬇️ Descargar Todo en formato ZIP",
        data=zip_buffer.getvalue(),
        file_name=f"Proyecto_Completo_{st.session_state.nombre_proj.replace(' ', '_')}.zip",
        mime="application/zip",
        use_container_width=True,
        type="primary"
    )
    st.caption("El archivo ZIP incluirá el documento Word (ERS), la matriz de Requisitos y el análisis de Montecarlo actualizados.")
