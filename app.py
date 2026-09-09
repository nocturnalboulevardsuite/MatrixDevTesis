import streamlit as st
import pandas as pd
import plotly.express as px
from streamlit_mermaid import st_mermaid
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="MatrixDevTesis", layout="wide", page_icon="🎓")

st.title("🎓 MatrixDevTesis")
st.caption("Suite web all-in-one para la gestión, modelado y documentación de proyectos informáticos.")

tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Requisitos & ERS",
    "📐 Diagramas Visuales",
    "📌 Gestión Estilo Jira & RACI",
    "📈 Analítica & Ruta Crítica"
])

# ==========================================
# TAB 1: REQUISITOS, ACTA Y HISTORIAS DE USUARIO
# ==========================================
with tab1:
    st.subheader("📋 Documentación Base, Requerimientos e Historias de Usuario")
    
    col_acta1, col_acta2 = st.columns([1, 1])
    
    with col_acta1:
        nombre_proj = st.text_input("Nombre del Proyecto", "Sistema de Control de Inventario MatrixDev")
        objetivo = st.text_area("Objetivo Principal del Sistema", "Automatizar el flujo de inventario y optimizar la generación de reportes universitarios.", height=100)
        mvp_scope = st.text_area("Alcance MVP (Producto Mínimo Viable)", "Módulo de autenticación, gestión CRUD de productos y exportación del documento ERS.", height=100)

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
    
    # ---------------------------------------------------------
    # GENERADOR DE EXCEL ESTILIZADO PROFESIONAL
    # ---------------------------------------------------------
    def generar_excel_estilizado(df_rf, df_rnf, nombre_proyecto):
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        
        thin_border = Border(
            left=Side(style='thin', color='D3D3D3'),
            right=Side(style='thin', color='D3D3D3'),
            top=Side(style='thin', color='D3D3D3'),
            bottom=Side(style='thin', color='D3D3D3')
        )
        
        def aplicar_estilo_hoja(ws, df, titulo_hoja, color_principal, color_suave):
            ws.views.sheetView[0].showGridLines = True
            
            # 1. Banner Principal de Título
            ws.merge_cells("A1:B1")
            title_cell = ws["A1"]
            title_cell.value = f"📌 {titulo_hoja.upper()}"
            title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color=color_principal, end_color=color_principal, fill_type="solid")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 35
            
            # 2. Subtítulo con Nombre del Proyecto
            ws.merge_cells("A2:B2")
            sub_cell = ws["A2"]
            sub_cell.value = f"Proyecto: {nombre_proyecto}"
            sub_cell.font = Font(name="Calibri", size=10, italic=True, color="595959")
            sub_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[2].height = 18

            # 3. Encabezados de Tabla
            headers = ["ID", "Descripción del Requerimiento"]
            ws.row_dimensions[4].height = 25
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col_idx, value=header)
                cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color=color_principal, end_color=color_principal, fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # 4. Datos con estilo Zebra y bordes suaves
            for r_idx, row in df.iterrows():
                row_num = r_idx + 5
                ws.row_dimensions[row_num].height = 22
                
                bg_color = color_suave if r_idx % 2 == 1 else "FFFFFF"
                fill_zebra = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                
                cell_id = ws.cell(row=row_num, column=1, value=row["ID"])
                cell_id.font = Font(name="Calibri", size=11, bold=True, color="333333")
                cell_id.alignment = Alignment(horizontal="center", vertical="center")
                cell_id.border = thin_border
                cell_id.fill = fill_zebra
                
                cell_desc = ws.cell(row=row_num, column=2, value=row["Descripción"])
                cell_desc.font = Font(name="Calibri", size=11, color="333333")
                cell_desc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                cell_desc.border = thin_border
                cell_desc.fill = fill_zebra

            # 5. Ancho de Columnas
            ws.column_dimensions["A"].width = 14
            ws.column_dimensions["B"].width = 80

        # Configuración Hoja Azul (Funcionales)
        ws_rf = wb.active
        ws_rf.title = "Requerimientos Funcionales"
        ws_rf.sheet_properties.tabColor = "1F4E78"
        aplicar_estilo_hoja(ws_rf, df_rf, "Requerimientos Funcionales", "1F4E78", "F2F5F9")
        
        # Configuración Hoja Roja (No Funcionales)
        ws_rnf = wb.create_sheet(title="Requerimientos No Funcionales")
        ws_rnf.sheet_properties.tabColor = "C00000"
        aplicar_estilo_hoja(ws_rnf, df_rnf, "Requerimientos No Funcionales", "C00000", "FDF2F2")

        wb.save(output)
        return output.getvalue()

    excel_data = generar_excel_estilizado(df_rf_final, df_rnf_final, nombre_proj)

    # Markdown Document
    rf_md_text = "\n".join([f"- **{row['ID']}**: {row['Descripción']}" for _, row in df_rf_final.iterrows()])
    rnf_md_text = "\n".join([f"- **{row['ID']}**: {row['Descripción']}" for _, row in df_rnf_final.iterrows()])

    doc_ers = f"""# ERS & Acta de Constitución: {nombre_proj}

## 1. Objetivo del Sistema
{objetivo}

## 2. Alcance del MVP
{mvp_scope}

## 3. Requerimientos Funcionales
{rf_md_text}

## 4. Requerimientos No Funcionales
{rnf_md_text}

## 5. Historias de Usuario
{us_df.to_markdown(index=False)}
"""
    
    col_dl1, col_dl2 = st.columns(2)
    
    with col_dl1:
        st.download_button(
            label="📄 Exportar ERS Completo (.md)", 
            data=doc_ers, 
            file_name=f"ERS_{nombre_proj.replace(' ', '_')}.md",
            use_container_width=True
        )
        
    with col_dl2:
        st.download_button(
            label="📊 Descargar Excel Requisitos Funcionales y no Funcionales",
            data=excel_data,
            file_name=f"Requisitos_{nombre_proj.replace(' ', '_')}.xlsx",
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
# TAB 4: ANALÍTICA, KPIS Y RUTA CRÍTICA
# ==========================================
with tab4:
    st.subheader("📈 Métricas del Proyecto y Análisis de Ruta Crítica")
    
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
