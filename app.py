import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from streamlit_mermaid import st_mermaid
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import google.generativeai as genai

# Librerías para generación de Word (.docx)
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="MatrixDevTesis", layout="wide", page_icon="🎓")

# ==========================================
# CONFIGURACIÓN Y LÍMITE DE USO DE IA
# ==========================================
MAX_USOS_IA = 5

if "usos_ia" not in st.session_state:
    st.session_state.usos_ia = 0

# Barra lateral limpia mostrando únicamente los créditos restantes del usuario
st.sidebar.title("⚙️ Estado de Sesión")
creditos_restantes = MAX_USOS_IA - st.session_state.usos_ia
st.sidebar.metric("Créditos de IA disponibles", f"{creditos_restantes} / {MAX_USOS_IA}")
st.sidebar.caption("Cada usuario cuenta con 5 mejoras automáticas con IA por sesión.")

def mejorar_texto_con_ia_limite(texto_original, tipo_campo):
    # 1. Validar si el usuario superó sus 5 usos gratuitos
    if st.session_state.usos_ia >= MAX_USOS_IA:
        st.error(f"🚫 Has alcanzado el límite máximo de {MAX_USOS_IA} mejoras con IA en esta sesión.")
        return None

    # 2. Leer la API Key oculta cargada en el servidor (secrets.toml)
    api_key_server = st.secrets.get("GEMINI_API_KEY", None)
    
    if not api_key_server:
        st.error("⚠️ La API Key del servidor no está configurada en los Secrets de Streamlit.")
        return None

    if not texto_original or texto_original.strip() == "":
        st.warning("⚠️ Escribe una idea o borrador inicial antes de solicitar la mejora.")
        return None
        
    try:
        # Limpieza de comillas o espacios accidentales en la API Key
        clean_key = str(api_key_server).strip().strip('"').strip("'")
        genai.configure(api_key=clean_key)
        
        prompt = f"""
        Eres un Ingeniero de Software Senior y revisor de memorias de título/tesis universitarias.
        Reescribe y mejora el siguiente borrador de texto para la sección '{tipo_campo}' de un documento de Especificación de Requisitos de Software (ERS).
        
        Instrucciones:
        - Transforma cualquier idea informal, vaga o mal redactada en un lenguaje sumamente profesional, técnico, formal y claro.
        - Mantén la intención original del usuario pero exprésala con estándares de la IEEE / Ingeniería de Software.
        - Devuelve ÚNICAMENTE el texto mejorado final, sin introducciones, saludos ni explicaciones.

        Borrador original:
        "{texto_original}"
        """
        
        modelos = ["gemini-1.5-flash", "gemini-3.6-flash"]
        ultimo_error = None
        
        for modelo_nombre in modelos:
            try:
                model = genai.GenerativeModel(modelo_nombre)
                response = model.generate_content(prompt)
                if response and response.text:
                    # Descontar un uso al tener éxito
                    st.session_state.usos_ia += 1
                    return response.text.strip()
            except Exception as err:
                ultimo_error = str(err)
                continue
                
        st.error(f"❌ No se pudo conectar con Gemini. Detalle: {ultimo_error}")
        return None

    except Exception as e:
        st.error(f"❌ Error al procesar con IA: {str(e)}")
        return None

st.title("🎓 MatrixDevTesis")
st.caption("Suite web all-in-one para la gestión, modelado y documentación de proyectos informáticos.")

tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Requisitos & ERS",
    "📐 Diagramas Visuales",
    "📌 Gestión Estilo Jira & RACI",
    "📈 Analítica, Ruta Crítica & Montecarlo"
])

# ==========================================
# TAB 1: REQUISITOS, ACTA Y HISTORIAS DE USUARIO
# ==========================================
with tab1:
    st.subheader("📋 Documentación Base, Requerimientos e Historias de Usuario")
    
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

        # Objetivo con IA transparente
        st.write("**Objetivo Principal del Sistema**")
        st.text_area("Objetivo", key="objetivo_text", height=90, label_visibility="collapsed")
        
        if st.button("✨ Mejorar Objetivo con IA", key="btn_ai_obj"):
            with st.spinner("Optimizando redacción con IA..."):
                resultado = mejorar_texto_con_ia_limite(st.session_state.objetivo_text, "Objetivo Principal")
                if resultado:
                    st.session_state.objetivo_text = resultado
                    st.success("¡Objetivo optimizado con éxito!")
                    st.rerun()

        # Alcance MVP con IA transparente
        st.write("**Alcance MVP (Producto Mínimo Viable)**")
        st.text_area("Alcance", key="mvp_text", height=90, label_visibility="collapsed")
        
        if st.button("✨ Mejorar Alcance MVP con IA", key="btn_ai_mvp"):
            with st.spinner("Optimizando redacción con IA..."):
                resultado = mejorar_texto_con_ia_limite(st.session_state.mvp_text, "Alcance MVP")
                if resultado:
                    st.session_state.mvp_text = resultado
                    st.success("¡Alcance optimizado con éxito!")
                    st.rerun()

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

    with col_acta2:
        st.write("**Requisitos Funcionales (RF)** — *Añade filas al final*")
        rf_edited = st.data_editor(
            st.session_state.df_rf, 
            num_rows="dynamic", 
            use_container_width=True, 
            key="rf_editor"
        )
        
        st.write("**Requisitos No Funcionales (RNF)** — *Añade filas al final*")
        rnf_edited = st.data_editor(
            st.session_state.df_rnf, 
            num_rows="dynamic", 
            use_container_width=True, 
            key="rnf_editor"
        )

    # Autonumeración de IDs
    df_rf_final = rf_edited.dropna(subset=["Descripción"]).reset_index(drop=True)
    df_rf_final["ID"] = [f"RF{i+1:02d}" for i in range(len(df_rf_final))]
    df_rf_final = df_rf_final[["ID", "Descripción"]]

    df_rnf_final = rnf_edited.dropna(subset=["Descripción"]).reset_index(drop=True)
    df_rnf_final["ID"] = [f"RNF{i+1:02d}" for i in range(len(df_rnf_final))]
    df_rnf_final = df_rnf_final[["ID", "Descripción"]]

    st.markdown("---")
    st.write("**Historias de Usuario (User Stories)**")
    
    if "user_stories" not in st.session_state:
        st.session_state.user_stories = pd.DataFrame([
            {"ID": "US-01", "Como": "Estudiante", "Quiero": "Generar mi documento ERS en un clic", "Para": "Entregarlo en la memoria de título"},
            {"ID": "US-02", "Como": "Profesor Evaluador", "Quiero": "Visualizar la ruta crítica", "Para": "Analizar la factibilidad del proyecto"}
        ])
    
    us_df = st.data_editor(st.session_state.user_stories, num_rows="dynamic", use_container_width=True, key="us_editor")

    st.markdown("---")
    st.subheader("📥 Exportación de Entregables")
    
    # Generador Word
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
        p_title.paragraph_format.line_spacing = 1.3
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

        p_mid_space = doc.add_paragraph()
        p_mid_space.paragraph_format.space_before = Pt(180)

        p_meta = doc.add_paragraph()
        p_meta.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_meta.paragraph_format.line_spacing = 1.5
        
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
            r_c.font.size = Pt(11)
            r_v = p_meta.add_run(f"{val}\n")
            r_v.font.size = Pt(11)

        doc.add_page_break()
        doc.add_paragraph("")
        doc.add_page_break()

        def agregar_encabezado_seccion(texto, nivel=1, color=RGBColor(31, 78, 120)):
            h = doc.add_heading(level=nivel)
            h.alignment = WD_ALIGN_PARAGRAPH.LEFT
            h.paragraph_format.space_before = Pt(16)
            h.paragraph_format.space_after = Pt(6)
            r = h.add_run(texto)
            r.bold = True
            r.font.name = 'Calibri'
            r.font.size = Pt(14 if nivel == 1 else 12)
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
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.name = 'Calibri'
                    run.font.bold = True
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(255, 255, 255)

        for _, row in df_rf.iterrows():
            row_cells = table_rf.add_row().cells
            row_cells[0].text = str(row["ID"])
            row_cells[1].text = str(row["Descripción"])
            for idx, cell in enumerate(row_cells):
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if idx == 0 else WD_ALIGN_PARAGRAPH.JUSTIFY
                    for run in p.runs:
                        run.font.name = 'Calibri'
                        run.font.size = Pt(10)

        doc.add_paragraph().paragraph_format.space_before = Pt(10)

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
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.name = 'Calibri'
                    run.font.bold = True
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(255, 255, 255)

        for _, row in df_rnf.iterrows():
            row_cells = table_rnf.add_row().cells
            row_cells[0].text = str(row["ID"])
            row_cells[1].text = str(row["Descripción"])
            for idx, cell in enumerate(row_cells):
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if idx == 0 else WD_ALIGN_PARAGRAPH.JUSTIFY
                    for run in p.runs:
                        run.font.name = 'Calibri'
                        run.font.size = Pt(10)

        doc.add_paragraph().paragraph_format.space_before = Pt(10)

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
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.name = 'Calibri'
                    run.font.bold = True
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(255, 255, 255)

        for _, row in df_us.iterrows():
            row_cells = table_us.add_row().cells
            row_cells[0].text = str(row.get("ID", ""))
            row_cells[1].text = str(row.get("Como", ""))
            row_cells[2].text = str(row.get("Quiero", ""))
            row_cells[3].text = str(row.get("Para", ""))
            
            for idx, cell in enumerate(row_cells):
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
                    for run in p.runs:
                        run.font.name = 'Calibri'
                        run.font.size = Pt(10)

        target_stream = io.BytesIO()
        doc.save(target_stream)
        return target_stream.getvalue()

    # Generador Excel
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
            ws.row_dimensions[3].height = 25
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col_idx, value=header)
                cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color=color_principal, end_color=color_principal, fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            for r_idx, row in df.iterrows():
                row_num = r_idx + 4
                ws.row_dimensions[row_num].height = 22
                bg_color = color_suave if r_idx % 2 == 1 else "FFFFFF"
                fill_zebra = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                
                cell_id = ws.cell(row=row_num, column=1, value=row["ID"])
                cell_id.alignment = Alignment(horizontal="center", vertical="center")
                cell_id.border = thin_border
                cell_id.fill = fill_zebra
                
                cell_desc = ws.cell(row=row_num, column=2, value=row["Descripción"])
                cell_desc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                cell_desc.border = thin_border
                cell_desc.fill = fill_zebra

            ws.column_dimensions["A"].width = 14
            ws.column_dimensions["B"].width = 80

        ws_rf = wb.active
        ws_rf.title = "Requerimientos Funcionales"
        ws_rf.sheet_properties.tabColor = "1F4E78"
        aplicar_estilo_hoja(ws_rf, df_rf, "Requerimientos Funcionales", "1F4E78", "F2F5F9")
        
        ws_rnf = wb.create_sheet(title="Requerimientos No Funcionales")
        ws_rnf.sheet_properties.tabColor = "C00000"
        aplicar_estilo_hoja(ws_rnf, df_rnf, "Requerimientos No Funcionales", "C00000", "FDF2F2")

        wb.save(output)
        return output.getvalue()

    word_data = generar_word_ers(
        st.session_state.nombre_proj,
        st.session_state.integrantes,
        st.session_state.profesor,
        st.session_state.fecha,
        st.session_state.seccion,
        st.session_state.objetivo_text, 
        st.session_state.mvp_text, 
        df_rf_final, 
        df_rnf_final, 
        us_df
    )
    
    excel_data = generar_excel_estilizado(df_rf_final, df_rnf_final, st.session_state.nombre_proj)

    col_dl1, col_dl2 = st.columns(2)
    
    with col_dl1:
        st.download_button(
            label="📝 Descargar Documento ERS en WORD (.docx)", 
            data=word_data, 
            file_name=f"ERS_{st.session_state.nombre_proj.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
    with col_dl2:
        st.download_button(
            label="📊 Descargar Excel de Requisitos (.xlsx)",
            data=excel_data,
            file_name=f"Requisitos_{st.session_state.nombre_proj.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ==========================================
# TAB 2: MODELADO VISUAL (MERMAID.JS)
# ==========================================
with tab2:
    st.subheader("📐 Diagramado & Modelado (Flowchart, UML, EDT)")
    
    plantilla = st.selectbox("Cargar Plantilla de Ejemplo:", [
        "Diagrama de Flujo (Flowchart)",
        "Diagrama de Clases (UML)",
        "Estructura de Desglose de Trabajo (EDT / WBS)"
    ])
    
    if plantilla == "Diagrama de Flujo (Flowchart)":
        default_mermaid = """graph TD
    A[Inicio] --> B{¿Usuario registrado?}
    B -- Sí --> C[Login Exitoso]
    B -- No --> D[Formulario de Registro]
    D --> C
    C --> E[Acceso al Panel MatrixDev]"""
    elif plantilla == "Diagrama de Clases (UML)":
        default_mermaid = """classDiagram
    class Usuario {
        +String nombre
        +String email
        +login()
    }
    class Proyecto {
        +String nombre
        +List tareas
        +exportarERS()
    }
    Usuario "1" -- "*" Proyecto : administra"""
    else:
        default_mermaid = """graph TD
    A[Proyecto MatrixDevTesis] --> B[1.0 Ingeniería de Requisitos]
    A --> C[2.0 Desarrollo Software]
    B --> B1[1.1 Levantamiento RF/RNF]
    B --> B2[1.2 Diagramas UML]
    C --> C1[2.1 Frontend Streamlit]
    C --> C2[2.2 Lógica en Python]"""

    codigo_mermaid = st.text_area("Código en sintaxis Mermaid.js", value=default_mermaid, height=180)
    st_mermaid(codigo_mermaid)

# ==========================================
# TAB 3: GESTIÓN DE PROYECTOS, JIRA, RACI & RIESGOS
# ==========================================
with tab3:
    st.subheader("📌 Tablero de Tareas, Matriz RACI y Gestión de Riesgos")
    
    col_k1, col_k2 = st.columns([2, 1])
    
    with col_k1:
        st.write("**Tablero Kanban (Estilo Jira)**")
        if "kanban_tasks" not in st.session_state:
            st.session_state.kanban_tasks = pd.DataFrame([
                {"Tarea": "Diseñar Modelo de Datos", "Estado": "Completado", "Asignado": "Ana", "Prioridad": "Alta"},
                {"Tarea": "Endpoints API Rest", "Estado": "En Proceso", "Asignado": "Carlos", "Prioridad": "Alta"},
                {"Tarea": "Pruebas de Integración", "Estado": "Pendiente", "Asignado": "Beatriz", "Prioridad": "Media"}
            ])
        
        kanban_df = st.data_editor(
            st.session_state.kanban_tasks, 
            num_rows="dynamic", 
            use_container_width=True,
            column_config={
                "Estado": st.column_config.SelectboxColumn(options=["Pendiente", "En Proceso", "Completado"])
            }
        )

    with col_k2:
        st.write("**Matriz de Riesgos**")
        riesgos_df = pd.DataFrame({
            "Riesgo": ["Retraso en Entregables", "Falla de Integración API"],
            "Impacto": ["Alto", "Medio"],
            "Estrategia Mitigación": ["Ajustar alcance MVP", "Crear respuestas Mock"]
        })
        st.dataframe(riesgos_df, use_container_width=True)

    st.markdown("---")
    st.write("**Matriz RACI (Asignación de Responsabilidades)**")
    raci_df = pd.DataFrame({
        "Entregable / Fase": ["Acta de Constitución", "Diagramas UML", "Desarrollo Backend", "Pruebas QA"],
        "Líder de Proyecto": ["A (Accountable)", "R (Responsible)", "I (Informed)", "C (Consulted)"],
        "Desarrollador": ["C (Consulted)", "C (Consulted)", "R (Responsible)", "I (Informed)"],
        "Tester / QA": ["I (Informed)", "I (Informed)", "C (Consulted)", "R (Responsible)"]
    })
    st.dataframe(raci_df, use_container_width=True)

# ==========================================
# TAB 4: ANALÍTICA, RUTA CRÍTICA & MONTECARLO
# ==========================================
with tab4:
    st.subheader("📈 Métricas del Proyecto, Ruta Crítica y Simulación Montecarlo")
    
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Progreso del Sprint", "68%", "+12%")
    kpi2.metric("Historias de Usuario", "6 / 10", "+2")
    kpi3.metric("Riesgos Activos", "2", "Estable")
    kpi4.metric("Días para Entrega", "14 Días", "-1 día")

    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.write("**Distribución de Tareas por Estado**")
        if not kanban_df.empty and "Estado" in kanban_df.columns:
            estado_counts = kanban_df["Estado"].value_counts().reset_index()
            estado_counts.columns = ["Estado", "Cantidad"]
            fig_bar = px.bar(estado_counts, x="Estado", y="Cantidad", color="Estado", text_auto=True)
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("Añade tareas al tablero para actualizar las estadísticas.")

    with col_g2:
        st.write("**Diagrama de Gantt & Ruta Crítica**")
        gantt_data = pd.DataFrame([
            dict(Tarea="Levantamiento Requisitos", Inicio='2026-03-01', Fin='2026-03-05', Tipo='Secundario'),
            dict(Tarea="Diseño Arquitectura", Inicio='2026-03-06', Fin='2026-03-12', Tipo='Ruta Crítica'),
            dict(Tarea="Desarrollo Core API", Inicio='2026-03-13', Fin='2026-03-24', Tipo='Ruta Crítica'),
            dict(Tarea="Pruebas & Despliegue", Inicio='2026-03-25', Fin='2026-03-30', Tipo='Ruta Crítica')
        ])
        fig_gantt = px.timeline(gantt_data, x_start="Inicio", x_end="Fin", y="Tarea", color="Tipo", title="Planificación temporal de Hitos")
        fig_gantt.update_yaxes(autorange="reversed")
        st.plotly_chart(fig_gantt, use_container_width=True)

    st.markdown("---")
    st.subheader("🎲 Estimación de Duración de Proyecto con Método de Montecarlo")
    
    col_mc_inputs, col_mc_params = st.columns([2, 1])
    
    with col_mc_inputs:
        st.write("**Definición de Tareas y Estimación PERT (Días)**")
        if "df_mc_tasks" not in st.session_state:
            st.session_state.df_mc_tasks = pd.DataFrame([
                {"Tarea": "Análisis y ERS", "Optimista": 3, "Mas_Probable": 5, "Pesimista": 10},
                {"Tarea": "Modelado UML y Base de Datos", "Optimista": 4, "Mas_Probable": 7, "Pesimista": 12},
                {"Tarea": "Desarrollo Backend & Frontend", "Optimista": 10, "Mas_Probable": 15, "Pesimista": 25},
                {"Tarea": "Pruebas y Despliegue", "Optimista": 3, "Mas_Probable": 5, "Pesimista": 8}
            ])
        
        mc_tasks_edited = st.data_editor(
            st.session_state.df_mc_tasks,
            num_rows="dynamic",
            use_container_width=True,
            key="mc_tasks_editor"
        )

    with col_mc_params:
        st.write("**Configuración de Simulación**")
        n_simulaciones = st.slider("Número de Simulaciones", min_value=500, max_value=10000, value=2500, step=500)
        
        if st.button("🔄 Ejecutar Simulación Montecarlo", use_container_width=True):
            st.success("Simulación ejecutada correctamente.")

    def simular_montecarlo(df_tasks, N):
        np.random.seed(42)
        total_duraciones = np.zeros(N)
        
        for _, row in df_tasks.iterrows():
            o = float(row["Optimista"])
            m = float(row["Mas_Probable"])
            p = float(row["Pesimista"])
            
            low = min(o, m, p)
            high = max(o, m, p)
            mode = max(low, min(m, high))
            
            duraciones_tarea = np.random.triangular(left=low, mode=mode, right=high, size=N)
            total_duraciones += duraciones_tarea
            
        return total_duraciones

    duraciones_simuladas = simular_montecarlo(mc_tasks_edited, n_simulaciones)

    p50 = np.percentile(duraciones_simuladas, 50)
    p80 = np.percentile(duraciones_simuladas, 80)
    p90 = np.percentile(duraciones_simuladas, 90)
    promedio = np.mean(duraciones_simuladas)

    mc_col1, mc_col2, mc_col3, mc_col4 = st.columns(4)
    mc_col1.metric("Duración Promedio", f"{promedio:.1f} Días")
    mc_col2.metric("P50 (Probabilidad 50%)", f"{p50:.1f} Días")
    mc_col3.metric("P80 (Recomendado 80%)", f"{p80:.1f} Días", delta=f"+{p80-p50:.1f}d vs P50")
    mc_col4.metric("P90 (Conservador 90%)", f"{p90:.1f} Días")

    fig_mc = go.Figure()

    fig_mc.add_trace(go.Histogram(
        x=duraciones_simuladas,
        name='Simulaciones',
        marker_color='#5B2C6F',
        opacity=0.75,
        nbinsx=40
    ))

    fig_mc.add_vline(x=p50, line_dash="dash", line_color="#3498DB", annotation_text=f"P50 ({p50:.1f}d)", annotation_position="top left")
    fig_mc.add_vline(x=p80, line_dash="dash", line_color="#E67E22", annotation_text=f"P80 ({p80:.1f}d)", annotation_position="top right")
    fig_mc.add_vline(x=p90, line_dash="dash", line_color="#E74C3C", annotation_text=f"P90 ({p90:.1f}d)", annotation_position="top right")

    fig_mc.update_layout(
        title="Distribución de Duración Total Estimada (Días)",
        xaxis_title="Días de Duración",
        yaxis_title="Frecuencia (N° de Escenarios)",
        bargap=0.05
    )

    st.plotly_chart(fig_mc, use_container_width=True)

    def generar_excel_montecarlo(df_tasks, duraciones, n_sim, nombre_p):
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        purple_color = "4A235A"
        thin_border = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'), top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))
        
        ws = wb.active
        ws.title = "Resumen Montecarlo"
        ws.sheet_properties.tabColor = purple_color
        ws.views.sheetView[0].showGridLines = True
        
        ws.merge_cells("A1:D1")
        title = ws["A1"]
        title.value = "🎲 REPORTE DE SIMULACIÓN DE MONTECARLO"
        title.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        title.fill = PatternFill(start_color=purple_color, end_color=purple_color, fill_type="solid")
        title.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 35

        ws.merge_cells("A2:D2")
        sub = ws["A2"]
        sub.value = f"Proyecto: {nombre_p} | Iteraciones: {n_sim}"
        sub.font = Font(name="Calibri", size=10, italic=True, color="595959")
        sub.alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=4, column=1, value="Métrica de Riesgo").font = Font(bold=True)
        ws.cell(row=4, column=2, value="Valor Estimado (Días)").font = Font(bold=True)
        
        metricas = [
            ("Promedio Esperado", float(np.mean(duraciones))),
            ("Percentil 10 (P10)", float(np.percentile(duraciones, 10))),
            ("Percentil 50 (Mediana - P50)", float(np.percentile(duraciones, 50))),
            ("Percentil 80 (Recomendado - P80)", float(np.percentile(duraciones, 80))),
            ("Percentil 90 (Conservador - P90)", float(np.percentile(duraciones, 90)))
        ]
        
        for idx, (m_nombre, m_val) in enumerate(metricas, 5):
            c1 = ws.cell(row=idx, column=1, value=m_nombre)
            c2 = ws.cell(row=idx, column=2, value=round(m_val, 2))
            c1.border = thin_border
            c2.border = thin_border

        start_row_tasks = 12
        ws.cell(row=start_row_tasks, column=1, value="Tarea").font = Font(bold=True)
        ws.cell(row=start_row_tasks, column=2, value="Optimista (O)").font = Font(bold=True)
        ws.cell(row=start_row_tasks, column=3, value="Más Probable (M)").font = Font(bold=True)
        ws.cell(row=start_row_tasks, column=4, value="Pesimista (P)").font = Font(bold=True)

        for r_i, row in df_tasks.iterrows():
            curr_row = start_row_tasks + 1 + r_i
            ws.cell(row=curr_row, column=1, value=row["Tarea"]).border = thin_border
            ws.cell(row=curr_row, column=2, value=row["Optimista"]).border = thin_border
            ws.cell(row=curr_row, column=3, value=row["Mas_Probable"]).border = thin_border
            ws.cell(row=curr_row, column=4, value=row["Pesimista"]).border = thin_border

        ws.column_dimensions["A"].width = 35
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["C"].width = 20
        ws.column_dimensions["D"].width = 20

        wb.save(output)
        return output.getvalue()

    excel_mc_data = generar_excel_montecarlo(mc_tasks_edited, duraciones_simuladas, n_simulaciones, st.session_state.nombre_proj)

    st.download_button(
        label="📥 Descargar Reporte de Montecarlo (.xlsx)",
        data=excel_mc_data,
        file_name=f"Montecarlo_{st.session_state.nombre_proj.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
