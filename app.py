import streamlit as st
import pandas as pd
import plotly.express as px
from streamlit_mermaid import st_mermaid
import io # Librería nativa para manejar el archivo Excel en memoria

# Configuración inicial de la aplicación
st.set_page_config(page_title="MatrixDevTesis", layout="wide", page_icon="🎓")

st.title("🎓 MatrixDevTesis")
st.caption("Suite web all-in-one para la gestión, modelado y documentación de proyectos informáticos.")

# Navegación en pestañas organizadas por área metodológica
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
    
    col_acta1, col_acta2 = st.columns(2)
    with col_acta1:
        nombre_proj = st.text_input("Nombre del Proyecto", "Sistema de Control de Inventario MatrixDev")
        objetivo = st.text_area("Objetivo Principal del Sistema", "Automatizar el flujo de inventario y optimizar la generación de reportes universitarios.")
        mvp_scope = st.text_area("Alcance MVP (Producto Mínimo Viable)", "Módulo de autenticación, gestión CRUD de productos y exportación del documento ERS.")
    
    with col_acta2:
        req_funcionales = st.text_area("Requisitos Funcionales (RF)", "RF01: Autenticación con credenciales universitarias.\nRF02: Registro y edición de tareas.\nRF03: Exportación en formato Markdown.")
        req_no_funcionales = st.text_area("Requisitos No Funcionales (RNF)", "RNF01: Tiempo de respuesta menor a 1.5 segundos.\nRNF02: Cifrado SSL en todas las peticiones.")

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
    # LÓGICA PARA EXPORTAR EXCEL DE REQUERIMIENTOS
    # ---------------------------------------------------------
    def parse_requerimientos(texto):
        """Función auxiliar para convertir el texto en un DataFrame estructurado"""
        lineas = [linea.strip() for linea in texto.split('\n') if linea.strip()]
        datos = []
        for linea in lineas:
            if ':' in linea:
                req_id, req_desc = linea.split(':', 1)
                datos.append({"ID": req_id.strip(), "Descripción": req_desc.strip()})
            else:
                datos.append({"ID": "-", "Descripción": linea.strip()})
        return pd.DataFrame(datos)

    # Convertimos los textos en DataFrames
    df_rf = parse_requerimientos(req_funcionales)
    df_rnf = parse_requerimientos(req_no_funcionales)

    # Creamos el archivo Excel en memoria con dos pestañas
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_rf.to_excel(writer, index=False, sheet_name="Funcionales")
        df_rnf.to_excel(writer, index=False, sheet_name="No Funcionales")
    
    # ---------------------------------------------------------
    # GENERACIÓN DEL MARKDOWN ERS
    # ---------------------------------------------------------
    doc_ers = f"""# ERS & Acta de Constitución: {nombre_proj}

## 1. Objetivo del Sistema
{objetivo}

## 2. Alcance del MVP
{mvp_scope}

## 3. Requerimientos Funcionales
{req_funcionales}

## 4. Requerimientos No Funcionales
{req_no_funcionales}

## 5. Historias de Usuario
{us_df.to_markdown(index=False)}
"""
    
    # Botones de descarga alineados en columnas
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
            data=excel_buffer.getvalue(),
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
