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
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 1. Configuración centrada y minimalista
st.set_page_config(page_title="MatrixDevTesis", layout="centered", page_icon="🍷")

# 2. Inyección de CSS (Tema Vino Mate Épico)
st.markdown("""
    <style>
    /* Fondo principal y color de texto */
    .stApp {
        background-color: #2c0f14;
        color: #f4ecec;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
    }
    
    /* Encabezados épicos */
    h1, h2, h3 {
        color: #e09f9f !important;
        font-weight: 300 !important;
        text-align: center;
    }
    
    /* Botones minimalistas */
    .stButton>button {
        background-color: #5c1e28;
        color: #ffffff !important;
        border: 1px solid #8a2d3b;
        border-radius: 8px;
        transition: 0.3s;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #8a2d3b;
        border-color: #ffffff;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.5);
    }

    /* Ocultar elementos innecesarios */
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
if "df_rf" not in st.session_state:
    st.session_state.df_rf = pd.DataFrame({"Descripción": ["Autenticación OAuth2.", "CRUD de usuarios."]}, dtype=str)
if "df_rnf" not in st.session_state:
    st.session_state.df_rnf = pd.DataFrame({"Descripción": ["Latencia < 200ms.", "Cifrado AES-256 en base de datos."]}, dtype=str)
if "kanban_tasks" not in st.session_state:
    st.session_state.kanban_tasks = pd.DataFrame({"Tarea": ["Modelo BD"], "Estado": ["Completado"]}, dtype=str)

# Variables auxiliares para los campos de texto
if "input_nuevo_rf" not in st.session_state:
    st.session_state.input_nuevo_rf = ""
if "input_nuevo_rnf" not in st.session_state:
    st.session_state.input_nuevo_rnf = ""
if "input_nueva_tarea" not in st.session_state:
    st.session_state.input_nueva_tarea = ""

# ==========================================
# FUNCIONES CALLBACK (Añadir sin errores)
# ==========================================
def agregar_rf():
    if st.session_state.input_nuevo_rf.strip():
        nuevo_df = pd.DataFrame({"Descripción": [st.session_state.input_nuevo_rf.strip()]})
        st.session_state.df_rf = pd.concat([st.session_state.df_rf, nuevo_df], ignore_index=True)
        st.session_state.input_nuevo_rf = "" # Limpia la barra al añadir

def agregar_rnf():
    if st.session_state.input_nuevo_rnf.strip():
        nuevo_df = pd.DataFrame({"Descripción": [st.session_state.input_nuevo_rnf.strip()]})
        st.session_state.df_rnf = pd.concat([st.session_state.df_rnf, nuevo_df], ignore_index=True)
        st.session_state.input_nuevo_rnf = ""

def agregar_tarea():
    if st.session_state.input_nueva_tarea.strip():
        nuevo_df = pd.DataFrame({"Tarea": [st.session_state.input_nueva_tarea.strip()], "Estado": ["Pendiente"]})
        st.session_state.kanban_tasks = pd.concat([st.session_state.kanban_tasks, nuevo_df], ignore_index=True)
        st.session_state.input_nueva_tarea = ""

# ==========================================
# FUNCIONES DE GENERACIÓN DE DOCUMENTOS
# ==========================================
def generar_word_ers():
    doc = Document()
    doc.add_heading(f"Especificación de Requisitos: {st.session_state.nombre_proj}", level=1)
    doc.add_paragraph(f"Objetivo: {st.session_state.objetivo_text}")
    
    doc.add_heading("Requisitos Funcionales", level=2)
    df_rf_clean = st.session_state.df_rf.dropna(subset=["Descripción"])
    for i, row in df_rf_clean.iterrows():
        if str(row['Descripción']).strip() != "":
            doc.add_paragraph(f"- {row['Descripción']}")
            
    doc.add_heading("Requisitos No Funcionales", level=2)
    df_rnf_clean = st.session_state.df_rnf.dropna(subset=["Descripción"])
    for i, row in df_rnf_clean.iterrows():
        if str(row['Descripción']).strip() != "":
            doc.add_paragraph(f"- {row['Descripción']}")
            
    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

def generar_excel_estilizado():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Requisitos"
    ws.append(["Tipo", "Descripción"])
    
    df_rf_clean = st.session_state.df_rf.dropna(subset=["Descripción"])
    for i, row in df_rf_clean.iterrows():
        if str(row['Descripción']).strip() != "":
            ws.append(["Funcional", row['Descripción']])
            
    df_rnf_clean = st.session_state.df_rnf.dropna(subset=["Descripción"])
    for i, row in df_rnf_clean.iterrows():
        if str(row['Descripción']).strip() != "":
            ws.append(["No Funcional", row['Descripción']])
            
    wb.save(output)
    return output.getvalue()

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.title("✦ MatrixDev ✦")
st.markdown("<p style='text-align: center; color: #a68a8d;'>Gestión y Arquitectura de Proyectos</p>", unsafe_allow_html=True)
st.write("---")

tab1, tab2, tab3, tab_full = st.tabs([
    "1. Requisitos", 
    "2. Modelado", 
    "3. Gestión", 
    "📦 Descarga Full"
])

# --- TAB 1: REQUISITOS ---
with tab1:
    st.markdown("### Definición del Sistema")
    st.session_state.nombre_proj = st.text_input("Nombre del Proyecto", st.session_state.nombre_proj)
    st.session_state.objetivo_text = st.text_area("Objetivo Principal", st.session_state.objetivo_text, height=80)
    
    st.markdown("### Requisitos Funcionales")
    
    # Nuevo formulario robusto para añadir RF
    col1, col2 = st.columns([4, 1])
    with col1:
        st.text_input("Nuevo RF", key="input_nuevo_rf", label_visibility="collapsed", placeholder="Escribe un requisito y presiona Enter...", on_change=agregar_rf)
    with col2:
        st.button("➕ Añadir", key="btn_rf", on_click=agregar_rf, use_container_width=True)

    # La tabla ahora se usa principalmente para visualizar, editar datos existentes o borrar
    st.session_state.df_rf = st.data_editor(
        st.session_state.df_rf, 
        num_rows="dynamic", 
        use_container_width=True,
        hide_index=True,
        column_config={"Descripción": st.column_config.TextColumn("Descripción", required=True)},
        key="editor_rf"
    )
    
    st.markdown("### Requisitos No Funcionales")
    
    # Nuevo formulario robusto para añadir RNF
    col3, col4 = st.columns([4, 1])
    with col3:
        st.text_input("Nuevo RNF", key="input_nuevo_rnf", label_visibility="collapsed", placeholder="Escribe un requisito no funcional y presiona Enter...", on_change=agregar_rnf)
    with col4:
        st.button("➕ Añadir", key="btn_rnf", on_click=agregar_rnf, use_container_width=True)

    st.session_state.df_rnf = st.data_editor(
        st.session_state.df_rnf, 
        num_rows="dynamic", 
        use_container_width=True,
        hide_index=True,
        column_config={"Descripción": st.column_config.TextColumn("Descripción", required=True)},
        key="editor_rnf"
    )

# --- TAB 2: MODELADO ---
with tab2:
    st.markdown("### Arquitectura Visual")
    default_mermaid = "graph TD\n A[Inicio] --> B{Validar}\n B -- Sí --> C[Éxito]\n B -- No --> D[Error]"
    codigo_mermaid = st.text_area("Sintaxis Mermaid", value=default_mermaid, height=100)
    st_mermaid(codigo_mermaid)

# --- TAB 3: GESTIÓN ---
with tab3:
    st.markdown("### Tablero Kanban")
    
    # Nuevo formulario robusto para añadir Tareas
    col5, col6 = st.columns([4, 1])
    with col5:
        st.text_input("Nueva Tarea", key="input_nueva_tarea", label_visibility="collapsed", placeholder="Escribe una nueva tarea y presiona Enter...", on_change=agregar_tarea)
    with col6:
        st.button("➕ Añadir", key="btn_tarea", on_click=agregar_tarea, use_container_width=True)

    st.session_state.kanban_tasks = st.data_editor(
        st.session_state.kanban_tasks, 
        num_rows="dynamic", 
        use_container_width=True,
        hide_index=True,
        column_config={
            "Tarea": st.column_config.TextColumn("Tarea", required=True),
            "Estado": st.column_config.SelectboxColumn("Estado", options=["Pendiente", "En Proceso", "Completado"], required=True)
        },
        key="editor_kanban"
    )

# --- TAB 4: REPORTE FULL (DESCARGA DE TODO) ---
with tab_full:
    st.markdown("### 📦 Exportación General Consolidada")
    st.write("Descarga un archivo ZIP que contiene todos los modelos, matrices y requerimientos configurados en las pestañas anteriores.")
    
    # Pre-generar archivos
    word_doc = generar_word_ers()
    excel_doc = generar_excel_estilizado()
    
    # Crear ZIP en memoria
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"ERS_{st.session_state.nombre_proj}.docx", word_doc)
        zip_file.writestr(f"Matrices_{st.session_state.nombre_proj}.xlsx", excel_doc)
        
        # Guardar también los datos del Kanban en CSV dentro del ZIP
        kanban_csv = st.session_state.kanban_tasks.to_csv(index=False).encode('utf-8')
        zip_file.writestr(f"Kanban_Tareas.csv", kanban_csv)

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
