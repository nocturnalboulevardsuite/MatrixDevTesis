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
    
    /* Encabezados */
    h1, h2, h3 {
        color: #e09f9f !important;
        font-weight: 300 !important;
        text-align: center;
    }
    
    /* Botones principales y de formulario */
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

# Listas principales de requisitos
if "rf_list" not in st.session_state:
    st.session_state.rf_list = ["Autenticación OAuth2.", "CRUD de usuarios."]
if "rnf_list" not in st.session_state:
    st.session_state.rnf_list = ["Latencia < 200ms.", "Cifrado AES-256 en base de datos."]
if "kanban_tasks" not in st.session_state:
    st.session_state.kanban_tasks = [
        {"Tarea": "Modelo BD", "Estado": "Completado"},
        {"Tarea": "Endpoints API", "Estado": "En Proceso"}
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
    ws = wb.active
    ws.title = "Requisitos"
    ws.append(["Código", "Tipo", "Descripción"])
    
    for idx, rf in enumerate(st.session_state.rf_list, 1):
        if rf.strip():
            ws.append([f"RF-{idx:02d}", "Funcional", rf.strip()])
            
    for idx, rnf in enumerate(st.session_state.rnf_list, 1):
        if rnf.strip():
            ws.append([f"RNF-{idx:02d}", "No Funcional", rnf.strip()])
            
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
    
    st.write("---")
    st.markdown("### Requisitos Funcionales (RF)")
    
    # Campo para agregar un nuevo RF
    with st.form("form_add_rf", clear_on_submit=True):
        col_in, col_btn = st.columns([4, 1])
        with col_in:
            nuevo_rf_val = st.text_input("Nuevo RF", placeholder="Escribe un requisito y presiona Enter o Añadir...", label_visibility="collapsed")
        with col_btn:
            btn_add_rf = st.form_submit_button("➕ Añadir", use_container_width=True)
            
        if btn_add_rf and nuevo_rf_val.strip():
            st.session_state.rf_list.append(nuevo_rf_val.strip())
            st.rerun()

    # Listado editable de RF
    if not st.session_state.rf_list:
        st.info("No hay requisitos funcionales registrados.")
    else:
        rf_to_delete = None
        for i, item in enumerate(st.session_state.rf_list):
            c_tag, c_input, c_del = st.columns([0.8, 5, 0.8])
            c_tag.markdown(f"**RF-{i+1:02d}**")
            new_val = c_input.text_input(f"rf_in_{i}", value=item, label_visibility="collapsed", key=f"rf_field_{i}")
            st.session_state.rf_list[i] = new_val
            
            if c_del.button("🗑️", key=f"del_rf_{i}", help="Eliminar requisito"):
                rf_to_delete = i

        if rf_to_delete is not None:
            st.session_state.rf_list.pop(rf_to_delete)
            st.rerun()

    st.write("---")
    st.markdown("### Requisitos No Funcionales (RNF)")
    
    # Campo para agregar un nuevo RNF
    with st.form("form_add_rnf", clear_on_submit=True):
        col_in_rnf, col_btn_rnf = st.columns([4, 1])
        with col_in_rnf:
            nuevo_rnf_val = st.text_input("Nuevo RNF", placeholder="Escribe un requisito no funcional y presiona Enter o Añadir...", label_visibility="collapsed")
        with col_btn_rnf:
            btn_add_rnf = st.form_submit_button("➕ Añadir", use_container_width=True)
            
        if btn_add_rnf and nuevo_rnf_val.strip():
            st.session_state.rnf_list.append(nuevo_rnf_val.strip())
            st.rerun()

    # Listado editable de RNF
    if not st.session_state.rnf_list:
        st.info("No hay requisitos no funcionales registrados.")
    else:
        rnf_to_delete = None
        for i, item in enumerate(st.session_state.rnf_list):
            c_tag, c_input, c_del = st.columns([0.8, 5, 0.8])
            c_tag.markdown(f"**RNF-{i+1:02d}**")
            new_val = c_input.text_input(f"rnf_in_{i}", value=item, label_visibility="collapsed", key=f"rnf_field_{i}")
            st.session_state.rnf_list[i] = new_val
            
            if c_del.button("🗑️", key=f"del_rnf_{i}", help="Eliminar requisito"):
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
    
    # Agregar nueva tarea
    with st.form("form_kanban", clear_on_submit=True):
        col_t, col_s, col_b = st.columns([3, 2, 1])
        with col_t:
            nueva_t = st.text_input("Tarea", placeholder="Nueva tarea...", label_visibility="collapsed")
        with col_s:
            estado_t = st.selectbox("Estado", ["Pendiente", "En Proceso", "Completado"], label_visibility="collapsed")
        with col_b:
            btn_add_k = st.form_submit_button("➕ Añadir", use_container_width=True)
            
        if btn_add_k and nueva_t.strip():
            st.session_state.kanban_tasks.append({"Tarea": nueva_t.strip(), "Estado": estado_t})
            st.rerun()

    # Mostrar lista editable de tareas
    if not st.session_state.kanban_tasks:
        st.info("No hay tareas registradas en el tablero.")
    else:
        task_to_delete = None
        for i, t in enumerate(st.session_state.kanban_tasks):
            col_txt, col_sel, col_d = st.columns([3, 2, 0.8])
            updated_text = col_txt.text_input(f"kt_{i}", value=t["Tarea"], label_visibility="collapsed", key=f"kt_in_{i}")
            updated_status = col_sel.selectbox(f"ks_{i}", ["Pendiente", "En Proceso", "Completado"], index=["Pendiente", "En Proceso", "Completado"].index(t["Estado"]), label_visibility="collapsed", key=f"ks_in_{i}")
            
            st.session_state.kanban_tasks[i]["Tarea"] = updated_text
            st.session_state.kanban_tasks[i]["Estado"] = updated_status
            
            if col_d.button("🗑️", key=f"del_k_{i}"):
                task_to_delete = i
                
        if task_to_delete is not None:
            st.session_state.kanban_tasks.pop(task_to_delete)
            st.rerun()

# --- TAB 4: REPORTE FULL ---
with tab_full:
    st.markdown("### 📦 Exportación General Consolidada")
    st.write("Descarga un archivo ZIP que contiene los requisitos en Word, la matriz en Excel y las tareas en CSV.")
    
    word_doc = generar_word_ers()
    excel_doc = generar_excel_estilizado()
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"ERS_{st.session_state.nombre_proj}.docx", word_doc)
        zip_file.writestr(f"Matrices_{st.session_state.nombre_proj}.xlsx", excel_doc)
        
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
