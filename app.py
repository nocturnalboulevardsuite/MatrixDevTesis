import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from streamlit_mermaid import st_mermaid
import io
import zipfile
import openpyxl

# Librerías de Estilos para Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Librerías para Word
from docx import Document
from docx.shared import Pt, RGBColor

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y TEMAS DINÁMICOS
# ==========================================
st.set_page_config(page_title="MatrixDev Tesis", layout="wide", page_icon="🍷")

# Inicialización del tema en session_state (Por defecto: Claro)
if "tema" not in st.session_state:
    st.session_state.tema = "claro"  # Opciones: "vino", "oscuro", "claro"

# Configuración de colores según el tema seleccionado
if st.session_state.tema == "oscuro":
    bg_color = "#18181b"
    text_color = "#f4f4f5"
    heading_color = "#38bdf8"
    sub_color = "#a1a1aa"
    btn_bg = "#27272a"
    btn_border = "#3f3f46"
    btn_hover = "#3f3f46"
    graph_text = "#f4f4f5"
    graph_bar = "#38bdf8"
    grid_color = "#27272a"
elif st.session_state.tema == "vino":
    bg_color = "#2c0f14"
    text_color = "#f4ecec"
    heading_color = "#e09f9f"
    sub_color = "#a68a8d"
    btn_bg = "#5c1e28"
    btn_border = "#8a2d3b"
    btn_hover = "#8a2d3b"
    graph_text = "#f4ecec"
    graph_bar = "#e09f9f"
    grid_color = "#442026"
else:  # "claro" (Por defecto)
    bg_color = "#ffffff"
    text_color = "#0f172a"
    heading_color = "#1d4ed8"
    sub_color = "#475569"
    btn_bg = "#f1f5f9"
    btn_border = "#cbd5e1"
    btn_hover = "#e2e8f0"
    graph_text = "#0f172a"
    graph_bar = "#2563eb"
    grid_color = "#e2e8f0"

# Inyección de CSS dinámico
st.markdown(f"""
    <style>
    /* Fondo principal y color de texto global */
    .stApp, .stAppHeader, [data-testid="stHeader"] {{
        background-color: {bg_color} !important;
        color: {text_color} !important;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
    }}
    
    /* Encabezados */
    h1, h2, h3, h4, h5, h6 {{
        color: {heading_color} !important;
        font-weight: 300 !important;
        text-align: center;
    }}
    
    /* Pestañas (Tabs) */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
    }}
    .stTabs [data-baseweb="tab-list"] button {{
        background-color: transparent !important;
        border-radius: 4px;
        padding: 8px 16px;
    }}
    .stTabs [data-baseweb="tab-list"] button p {{
        color: {sub_color} !important;
        font-size: 1rem;
        font-weight: 500;
    }}
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] p {{
        color: {heading_color} !important;
        font-weight: bold;
    }}
    .stTabs [data-baseweb="tab-highlight"] {{
        background-color: {heading_color} !important;
    }}

    /* Textos generales, párrafos y etiquetas */
    p, span, label, div, .stMarkdown {{
        color: {text_color};
    }}

    /* Inputs, Selects y Textareas */
    div[data-baseweb="input"] > div, div[data-baseweb="select"] > div {{
        background-color: {btn_bg} !important;
        color: {text_color} !important;
        border-color: {btn_border} !important;
    }}
    input, textarea {{
        color: {text_color} !important;
    }}

    /* Métricas */
    [data-testid="stMetricValue"], [data-testid="stMetricLabel"] {{
        color: {text_color} !important;
    }}

    /* Botones */
    .stButton>button, .stFormSubmitButton>button, .stDownloadButton>button {{
        background-color: {btn_bg} !important;
        color: {text_color} !important;
        border: 1px solid {btn_border} !important;
        border-radius: 8px;
        transition: 0.3s;
        font-weight: bold;
    }}
    .stButton>button:hover, .stFormSubmitButton>button:hover, .stDownloadButton>button:hover {{
        background-color: {btn_hover} !important;
        border-color: {heading_color} !important;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.2);
    }}

    /* Ocultar elementos nativos innecesarios */
    header {{visibility: hidden;}}
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. PLANTILLAS Y VARIABLES DE ESTADO
# ==========================================
PLANTILLAS_FLUJO = {
    "Autenticación": """graph TD
    A[Usuario] -->|Credenciales| B[API Login]
    B -->|Validar| C{"¿Es Válido?"}
    C -->|Sí| D[Generar Token JWT]
    C -->|No| E[Error 401 Unauthorized]""",

    "Procesamiento API": """graph LR
    Cliente[Cliente] -->|Request POST| Router[Router]
    Router --> Controller[Controller]
    Controller -->|Query| DB[(Base de Datos)]
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
    A[Inicio] --> B[Tu Nuevo Proceso]
    B --> C{"¿Aprobado?"}
    C -->|Sí| D[Resultado Éxito]
    C -->|No| E[Resultado Fallo]"""
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
        st.session_state.txt_flujo_code = PLANTILLAS_FLUJO[sel]

def generar_mermaid_edt(edt_items, nombre_proyecto):
    proj_name = str(nombre_proyecto).replace('"', '').replace("'", "").replace("$", "")
    lines = ["graph TD", f'    ROOT["📦 {proj_name}"]']
    
    fases = {}
    for item in edt_items:
        fase = str(item.get("Fase", "Sin Fase")).strip().replace('"', '').replace("'", "")
        subfase = str(item.get("Subfase", "")).strip().replace('"', '').replace("'", "")
        paquete = str(item.get("Paquete", "Tarea")).strip().replace('"', '').replace("'", "")
        horas = float(item.get("Horas", 0))
        costo = float(item.get("Costo", 0))
        doc = str(item.get("Documentos", "")).strip().replace('"', '').replace("'", "")
        
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

def clasificar_riesgo(sev):
    if sev <= 5:
        return "Bajo 🟢"
    elif sev <= 10:
        return "Medio 🟡"
    elif sev <= 15:
        return "Alto 🟠"
    else:
        return "Crítico 🔴"

# ==========================================
# 3. FUNCIONES DE GENERACIÓN DE DOCUMENTOS INDIVIDUALES
# ==========================================

# 3.1 Requisitos Excel
def generar_excel_requisitos():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Requisitos"
    ws.views.sheetView[0].showGridLines = True
    
    ws.append(["Proyecto:", st.session_state.nombre_proj])
    ws.append(["Objetivo:", st.session_state.objetivo_text])
    ws.append([])
    ws.append(["Código", "Tipo de Requisito", "Descripción"])
    
    for idx, rf in enumerate(st.session_state.rf_list, 1):
        if rf.strip():
            ws.append([f"RF-{idx:02d}", "Funcional", rf.strip()])
            
    for idx, rnf in enumerate(st.session_state.rnf_list, 1):
        if rnf.strip():
            ws.append([f"RNF-{idx:02d}", "No Funcional", rnf.strip()])

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    wb.save(output)
    return output.getvalue()

# 3.2 Kanban Excel
def generar_excel_kanban():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tablero Kanban"
    ws.views.sheetView[0].showGridLines = True
    
    ws.append(["Proyecto:", st.session_state.nombre_proj])
    ws.append([])
    ws.append(["N°", "Tarea", "Estado"])
    
    for idx, k in enumerate(st.session_state.kanban_tasks, 1):
        ws.append([idx, k["Tarea"], k["Estado"]])

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    wb.save(output)
    return output.getvalue()

# 3.3 MONTE CARLO EXCEL ALTAMENTE ESTILIZADO Y DINÁMICO
def generar_excel_montecarlo(datos_dias, datos_costos=None):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Simulación Monte Carlo"
    ws.views.sheetView[0].showGridLines = True

    # 🎨 ESTILOS DE CORTE EJECUTIVO / MINIMALISTA
    HEADER_MAIN_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")  # Azul Marino
    SUBHEADER_FILL = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")   # Azul Acento
    ACCENT_BG_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")   # Gris ultra claro
    ZEBRA_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")       # Gris alternado
    
    SUCCESS_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")     # Verde suave
    WARNING_FILL = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")     # Rojo suave

    FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    FONT_BOLD = Font(name="Calibri", size=11, bold=True, color="0F172A")
    FONT_REGULAR = Font(name="Calibri", size=11, color="334155")

    ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
    ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

    THIN_BORDER = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    # 1. TÍTULO Y LOGO/BANNER
    ws.merge_cells("A1:C1")
    title_cell = ws["A1"]
    title_cell.value = f"📊 REPORTE DE SIMULACIÓN MONTE CARLO: {st.session_state.nombre_proj.upper()}"
    title_cell.font = FONT_TITLE
    title_cell.fill = HEADER_MAIN_FILL
    title_cell.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 35

    # 2. PARÁMETROS CONFIGURADOS
    dist_name = st.session_state.get("mc_dist_type", "Triangular")
    dias_meta = st.session_state.get("mc_dias_meta", 25)
    costo_meta = st.session_state.get("mc_costo_meta", 500000)
    percentil_target = st.session_state.get("mc_target_perc", 85)

    params = [
        ("Proyecto:", st.session_state.nombre_proj),
        ("Modelo de Distribución Usado:", dist_name),
        ("Días Meta (SLA Tiempo):", dias_meta),
        ("Presupuesto Meta (SLA Costo):", costo_meta),
        ("Muestras Simuladas (Iteraciones):", len(datos_dias))
    ]

    ws.cell(row=3, column=1, value="PARÁMETROS DEL PROYECTO Y CONFIGURACIÓN").font = FONT_BOLD
    ws.merge_cells("A3:C3")
    ws.row_dimensions[3].height = 22

    r = 4
    for label, val in params:
        c1 = ws.cell(row=r, column=1, value=label)
        c2 = ws.cell(row=r, column=2, value=val)
        c3 = ws.cell(row=r, column=3, value="")

        c1.font = FONT_BOLD
        c2.font = FONT_REGULAR
        c1.fill = ACCENT_BG_FILL
        
        c1.border = THIN_BORDER
        c2.border = THIN_BORDER
        c3.border = THIN_BORDER

        if isinstance(val, (int, float)):
            if "Presupuesto" in label:
                c2.number_format = "$#,##0"
            else:
                c2.number_format = "#,##0"
            c2.alignment = ALIGN_RIGHT
        else:
            c2.alignment = ALIGN_LEFT

        ws.row_dimensions[r].height = 20
        r += 1

    # 3. SECCIÓN 1: MÉTRICAS DE TIEMPO
    r += 1
    ws.cell(row=r, column=1, value="Métrica de Tiempo / SLA").font = FONT_HEADER
    ws.cell(row=r, column=1).fill = SUBHEADER_FILL
    ws.cell(row=r, column=2, value="Valor Estimado").font = FONT_HEADER
    ws.cell(row=r, column=2).fill = SUBHEADER_FILL
    ws.cell(row=r, column=3, value="Unidad / Formato").font = FONT_HEADER
    ws.cell(row=r, column=3).fill = SUBHEADER_FILL
    ws.row_dimensions[r].height = 25
    r += 1

    p50_dias = np.percentile(datos_dias, 50)
    p_target_dias = np.percentile(datos_dias, percentil_target)
    p95_dias = np.percentile(datos_dias, 95)

    exitos_dias = (datos_dias <= dias_meta).sum()
    tasa_exito_dias = (exitos_dias / len(datos_dias))
    riesgo_retraso = 1.0 - tasa_exito_dias
    capacidad_anual = (365 / p_target_dias) if p_target_dias > 0 else 0

    time_metrics = [
        ("Percentil 50 (P50 - Mediana Esperada)", p50_dias, "Días", "0.0"),
        (f"Percentil Objetivo (P{percentil_target} - Compromiso Seguros)", p_target_dias, "Días", "0.0"),
        ("Percentil 95 (P95 - Escenario Pesimista)", p95_dias, "Días", "0.0"),
        ("Tasa de Éxito en Tiempo", tasa_exito_dias, "Porcentaje", "0.0%"),
        ("Riesgo de Retraso", riesgo_retraso, "Porcentaje", "0.0%"),
        ("Capacidad Anual Proyectada", capacidad_anual, "Proyectos / Año", "0.0")
    ]

    for idx, (label, val, unit, num_fmt) in enumerate(time_metrics):
        c1 = ws.cell(row=r, column=1, value=label)
        c2 = ws.cell(row=r, column=2, value=val)
        c3 = ws.cell(row=r, column=3, value=unit)

        c1.font = FONT_REGULAR
        c2.font = FONT_BOLD if ("Objetivo" in label or "Éxito" in label) else FONT_REGULAR
        c3.font = FONT_REGULAR

        c1.border = THIN_BORDER
        c2.border = THIN_BORDER
        c3.border = THIN_BORDER

        c2.number_format = num_fmt
        c2.alignment = ALIGN_RIGHT
        c3.alignment = ALIGN_CENTER

        if idx % 2 == 1:
            c1.fill = ZEBRA_FILL
            c2.fill = ZEBRA_FILL
            c3.fill = ZEBRA_FILL

        if "Tasa de Éxito" in label:
            if tasa_exito_dias >= 0.8:
                c2.fill = SUCCESS_FILL
            elif tasa_exito_dias < 0.5:
                c2.fill = WARNING_FILL
        elif "Riesgo de Retraso" in label:
            if riesgo_retraso >= 0.5:
                c2.fill = WARNING_FILL

        ws.row_dimensions[r].height = 20
        r += 1

    # 4. SECCIÓN 2: MÉTRICAS FINANCIERAS / COSTOS
    if datos_costos is not None:
        r += 1
        ws.cell(row=r, column=1, value="Métrica Financiera / Costos").font = FONT_HEADER
        ws.cell(row=r, column=1).fill = SUBHEADER_FILL
        ws.cell(row=r, column=2, value="Valor Estimado ($)").font = FONT_HEADER
        ws.cell(row=r, column=2).fill = SUBHEADER_FILL
        ws.cell(row=r, column=3, value="Unidad / Formato").font = FONT_HEADER
        ws.cell(row=r, column=3).fill = SUBHEADER_FILL
        ws.row_dimensions[r].height = 25
        r += 1

        p50_cost = np.percentile(datos_costos, 50)
        p_target_cost = np.percentile(datos_costos, percentil_target)
        p95_cost = np.percentile(datos_costos, 95)
        exitos_cost = (datos_costos <= costo_meta).sum()
        tasa_exito_cost = (exitos_cost / len(datos_costos))

        cost_metrics = [
            ("Costo Mediano (P50)", p50_cost, "Moneda ($)", "$#,##0"),
            (f"Costo Recomendado (P{percentil_target})", p_target_cost, "Moneda ($)", "$#,##0"),
            ("Costo Pesimista (P95)", p95_cost, "Moneda ($)", "$#,##0"),
            ("Tasa de Éxito Presupuestario", tasa_exito_cost, "Porcentaje", "0.0%")
        ]

        for idx, (label, val, unit, num_fmt) in enumerate(cost_metrics):
            c1 = ws.cell(row=r, column=1, value=label)
            c2 = ws.cell(row=r, column=2, value=val)
            c3 = ws.cell(row=r, column=3, value=unit)

            c1.font = FONT_REGULAR
            c2.font = FONT_BOLD if ("Recomendado" in label or "Éxito" in label) else FONT_REGULAR
            c3.font = FONT_REGULAR

            c1.border = THIN_BORDER
            c2.border = THIN_BORDER
            c3.border = THIN_BORDER

            c2.number_format = num_fmt
            c2.alignment = ALIGN_RIGHT
            c3.alignment = ALIGN_CENTER

            if idx % 2 == 1:
                c1.fill = ZEBRA_FILL
                c2.fill = ZEBRA_FILL
                c3.fill = ZEBRA_FILL

            if "Tasa de Éxito" in label:
                if tasa_exito_cost >= 0.8:
                    c2.fill = SUCCESS_FILL
                elif tasa_exito_cost < 0.5:
                    c2.fill = WARNING_FILL

            ws.row_dimensions[r].height = 20
            r += 1

    # 5. AJUSTE DINÁMICO DE ANCHO DE COLUMNAS (AUTO-FIT PERFECTO)
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row == 1:  # Ignorar celda de título unificada
                continue
            if cell.value is not None:
                # Si es un número formateado como moneda o porcentaje, estimar largo visual
                val_str = str(cell.value)
                if isinstance(cell.value, float) and "0.0%" in (cell.number_format or ""):
                    val_str = f"{cell.value * 100:.1f}%"
                elif isinstance(cell.value, (int, float)) and "$#" in (cell.number_format or ""):
                    val_str = f"${cell.value:,.0f}"
                
                if len(val_str) > max_len:
                    max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 6, 16)

    wb.save(output)
    return output.getvalue()

# 3.4 EDT Excel
def generar_excel_edt():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EDT WBS"
    ws.views.sheetView[0].showGridLines = True
    
    ws.append(["Proyecto:", st.session_state.nombre_proj])
    ws.append([])
    ws.append(["Fase Principal", "Subfase / Grupo", "Paquete de Trabajo", "Horas Estimadas", "Costo Estimado ($)", "Documentos / Referencias"])
    
    for item in st.session_state.edt_list:
        ws.append([
            item.get("Fase", ""),
            item.get("Subfase", ""),
            item.get("Paquete", ""),
            item.get("Horas", 0),
            item.get("Costo", 0),
            item.get("Documentos", "")
        ])

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    wb.save(output)
    return output.getvalue()

# 3.5 Riesgos Excel
def generar_excel_riesgos():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Matriz de Riesgos"
    ws.views.sheetView[0].showGridLines = True
    
    ws.append(["Proyecto:", st.session_state.nombre_proj])
    ws.append([])
    ws.append(["Riesgo", "Probabilidad (1-5)", "Impacto (1-5)", "Severidad (P x I)", "Nivel de Riesgo"])
    
    for r_item in st.session_state.riesgos_list:
        prob_v = r_item.get("Probabilidad", 1)
        imp_v = r_item.get("Impacto", 1)
        sev = prob_v * imp_v
        ws.append([
            r_item.get("Riesgo", ""),
            prob_v,
            imp_v,
            sev,
            clasificar_riesgo(sev)
        ])

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    wb.save(output)
    return output.getvalue()

# 3.6 Informes Word
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

def generar_word_mementomori():
    doc = Document()
    
    # Encabezado principal
    p_main = doc.add_heading(level=0)
    r_main = p_main.add_run(f"REPORTE MEMENTO MORI: {st.session_state.nombre_proj.upper()}")
    r_main.font.bold = True
    
    doc.add_paragraph(f"Integrantes / Responsables: {st.session_state.integrantes}")
    doc.add_paragraph("Análisis Crítico Integral de Requisitos, Arquitectura, Tiempos, EDT y Riesgos del Sistema.")
    doc.add_paragraph("="*65)
    
    # 1. INTRODUCCIÓN Y OBJETIVOS
    doc.add_heading("1. Introducción y Objetivos del Proyecto", level=1)
    doc.add_paragraph(f"Objetivo Principal: {st.session_state.objetivo_text}")
    
    p1 = doc.add_paragraph()
    p1.add_run("🔍 Análisis Crítico e Interpretación: ").bold = True
    p1.add_run(
        f"El objetivo planteado para el proyecto '{st.session_state.nombre_proj}' aborda la necesidad clave de automatización e integración. "
        "Sin embargo, el punto crítico principal radica en controlar el alcance funcional inicial. "
        "Bajo el análisis MementoMori, el principal riesgo de esta fase es la expansión descontrolada del alcance (Scope Creep), "
        "la cual puede desalinear las expectativas de los interesados antes de consolidar el producto mínimo viable (MVP)."
    )

    # 2. REQUISITOS
    doc.add_heading("2. Especificación de Requisitos (RF y RNF)", level=1)
    doc.add_heading("Requisitos Funcionales (RF)", level=2)
    for idx, rf in enumerate(st.session_state.rf_list, 1):
        if rf.strip():
            doc.add_paragraph(f"• RF-{idx:02d}: {rf.strip()}")
            
    doc.add_heading("Requisitos No Funcionales (RNF)", level=2)
    for idx, rnf in enumerate(st.session_state.rnf_list, 1):
        if rnf.strip():
            doc.add_paragraph(f"• RNF-{idx:02d}: {rnf.strip()}")
            
    p2 = doc.add_paragraph()
    p2.add_run("🔍 Análisis Crítico de Requisitos: ").bold = True
    p2.add_run(
        "Al evaluar la matriz de requisitos, los puntos críticos se concentran en los Requisitos No Funcionales (RNF). "
        "Si bien los requisitos funcionales aseguran la operatividad del software, un fallo en el cumplimiento de los RNF "
        "(como latencias elevadas o vulnerabilidades de seguridad) generará la falla total de adopción por parte de los usuarios finales, "
        "imposibilitando la escalabilidad del sistema."
    )

    # 3. DIAGRAMAS Y ARQUITECTURA
    doc.add_heading("3. Arquitectura y Diagramas de Flujo / Secuencia", level=1)
    doc.add_paragraph("Sintaxis del Diagrama de Flujo / Secuencia Activo:")
    doc.add_paragraph(st.session_state.flujo_codigo)
    
    p3 = doc.add_paragraph()
    p3.add_run("🔍 Interpretación de Arquitectura y Puntos Críticos: ").bold = True
    p3.add_run(
        "Con estos diagramas de flujo y secuencia podemos entender la interacción de componentes y el recorrido de la información. "
        "El punto crítico identificado es el acoplamiento directo entre el controlador y los servicios de base de datos o autenticación. "
        "Si no se implementan políticas de reintento, caché o manejo de excepciones HTTP (ej. 401 Unauthorized o 500 Server Error), "
        "cualquier caída puntual de la base de datos provocará un colapso en cadena en los endpoints expuestos al usuario."
    )

    # 4. GESTIÓN KANBAN
    doc.add_heading("4. Estado de Gestión de Tareas (Tablero Kanban)", level=1)
    if st.session_state.kanban_tasks:
        tbl_k = doc.add_table(rows=1, cols=2)
        tbl_k.rows[0].cells[0].text = "Tarea"
        tbl_k.rows[0].cells[1].text = "Estado"
        for t in st.session_state.kanban_tasks:
            row = tbl_k.add_row().cells
            row[0].text = t["Tarea"]
            row[1].text = t["Estado"]
    else:
        doc.add_paragraph("No existen tareas registradas.")

    p4 = doc.add_paragraph()
    p4.add_run("🔍 Análisis Crítico del Flujo Kanban: ").bold = True
    p4.add_run(
        "El análisis del tablero de tareas revela que la acumulación de ítems en estado 'En Proceso' "
        "representa la mayor amenaza para el ritmo de entrega (Work In Progress). Se sugiere definir un límite estricto de tareas simultáneas "
        "para evitar la sobrecarga de trabajo y la sobreestimación del rendimiento del equipo."
    )

    # 5. EDT / WBS
    doc.add_heading("5. Estructura de Desglose de Trabajo (EDT / WBS)", level=1)
    if st.session_state.edt_list:
        tbl_e = doc.add_table(rows=1, cols=5)
        hdr_e = tbl_e.rows[0].cells
        hdr_e[0].text = "Fase"
        hdr_e[1].text = "Paquete de Trabajo"
        hdr_e[2].text = "Horas"
        hdr_e[3].text = "Costo ($)"
        hdr_e[4].text = "Documentos"
        
        tot_hrs = 0
        tot_cost = 0
        for item in st.session_state.edt_list:
            r = tbl_e.add_row().cells
            r[0].text = str(item.get("Fase", ""))
            r[1].text = str(item.get("Paquete", ""))
            r[2].text = str(item.get("Horas", 0))
            r[3].text = f"${item.get('Costo', 0):,.0f}"
            r[4].text = str(item.get("Documentos", ""))
            tot_hrs += item.get("Horas", 0)
            tot_cost += item.get("Costo", 0)
            
        doc.add_paragraph(f"Métricas Totales Planificadas: {tot_hrs:.1f} Horas Totales | Presupuesto Total: ${tot_cost:,.0f}")
    else:
        doc.add_paragraph("No hay paquetes de EDT registrados.")

    p5 = doc.add_paragraph()
    p5.add_run("🔍 Puntos Críticos Presupuestarios y de Carga Horaria: ").bold = True
    p5.add_run(
        "Al examinar la EDT, se observa que las actividades con mayor asignación de costo y horas conforman la ruta crítica del proyecto. "
        "Cualquier imprevisto de estimación en estas fases clave provocará desviaciones acumulativas, "
        "agotando la reserva de contingencia financiera y extendiendo la fecha límite original."
    )

    # 6. MONTE CARLO
    doc.add_heading("6. Análisis y Simulación Monte Carlo", level=1)
    opt = st.session_state.get("mc_opt", 10)
    prob = st.session_state.get("mc_prob", 20)
    pes = st.session_state.get("mc_pes", 45)
    sims = st.session_state.get("mc_sims", 2000)

    if opt <= prob <= pes:
        np.random.seed(42)
        sim_data = np.random.triangular(left=opt, mode=prob, right=pes, size=sims)
        p50 = np.percentile(sim_data, 50)
        p85 = np.percentile(sim_data, 85)
        p95 = np.percentile(sim_data, 95)

        doc.add_paragraph(f"Parámetros Configurados: Optimista={opt} días | Más Probable={prob} días | Pesimista={pes} días (Muestras: {sims})")
        doc.add_paragraph(f"• Percentil 50 (Mediana Esperada): {p50:.1f} Días")
        doc.add_paragraph(f"• Percentil 85 (Recomendación de Compromiso): {p85:.1f} Días")
        doc.add_paragraph(f"• Percentil 95 (Escenario de Alta Incertidumbre): {p95:.1f} Días")

        p6 = doc.add_paragraph()
        p6.add_run("🔍 Veredicto y Puntos Críticos de Monte Carlo: ").bold = True
        p6.add_run(
            f"Con esta simulación Monte Carlo podemos concluir que la entrega es factible a los {p85:.1f} días con un 85% de nivel de confianza. "
            f"Sin embargo, existe un margen de incertidumbre del 15% asociado a la cola pesimista ({pes} días). "
            f"Si se materializan complicaciones técnicas de manera simultánea, la probabilidad de sobrepasar la estimación central es considerable, "
            "lo que generará retrasos en la entrega final salvo que se mantengan controles semanales de seguimiento."
        )
    else:
        doc.add_paragraph("Los parámetros introducidos para la simulación Monte Carlo no son válidos.")

    # 7. MATRIZ DE RIESGOS
    doc.add_heading("7. Evaluación y Filosofía de Riesgos (MementoMori)", level=1)
    if st.session_state.riesgos_list:
        tbl_r = doc.add_table(rows=1, cols=5)
        hdr_r = tbl_r.rows[0].cells
        hdr_r[0].text = "Riesgo"
        hdr_r[1].text = "Probabilidad"
        hdr_r[2].text = "Impacto"
        hdr_r[3].text = "Severidad"
        hdr_r[4].text = "Nivel"

        for r_item in st.session_state.riesgos_list:
            r_row = tbl_r.add_row().cells
            prob_v = r_item.get("Probabilidad", 1)
            imp_v = r_item.get("Impacto", 1)
            sev = prob_v * imp_v
            r_row[0].text = str(r_item.get("Riesgo", ""))
            r_row[1].text = str(prob_v)
            r_row[2].text = str(imp_v)
            r_row[3].text = str(sev)
            r_row[4].text = clasificar_riesgo(sev)
    else:
        doc.add_paragraph("No hay riesgos registrados en la matriz.")

    p7 = doc.add_paragraph()
    p7.add_run("🔍 Estrategia de Mitigación MementoMori: ").bold = True
    p7.add_run(
        "Bajo el principio 'Memento Mori' (recordar la vulnerabilidad y planificar anticipadamente los fallos), "
        "los riesgos con severidad alta y crítica deben contar con planes de contingencia documentados antes de iniciar el desarrollo. "
        "Ignorar las señales de alerta tempranas en los componentes críticos provocará que problemas menores escalen a fallos sistémicos indeseados."
    )

    # 8. CONCLUSIONES
    doc.add_heading("8. Conclusiones y Recomendaciones Finales", level=1)
    doc.add_paragraph(
        f"El proyecto '{st.session_state.nombre_proj}' cuenta con una base metodológica y técnica sólida. "
        f"Se recomienda formalizar el compromiso de entrega considerando la simulación de Monte Carlo "
        "y monitorear de forma continua los requisitos no funcionales y la matriz de riesgos para asegurar el éxito del proyecto."
    )

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

def generar_excel_estilizado():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # Hoja 1: Requisitos
    ws1 = wb.active
    ws1.title = "Requisitos"
    ws1.views.sheetView[0].showGridLines = True
    ws1.append(["Código", "Tipo", "Descripción"])
    for idx, rf in enumerate(st.session_state.rf_list, 1):
        if rf.strip():
            ws1.append([f"RF-{idx:02d}", "Funcional", rf.strip()])
    for idx, rnf in enumerate(st.session_state.rnf_list, 1):
        if rnf.strip():
            ws1.append([f"RNF-{idx:02d}", "No Funcional", rnf.strip()])

    # Hoja 2: Kanban
    ws2 = wb.create_sheet(title="Kanban")
    ws2.views.sheetView[0].showGridLines = True
    ws2.append(["Tarea", "Estado"])
    for k in st.session_state.kanban_tasks:
        ws2.append([k["Tarea"], k["Estado"]])

    # Hoja 3: EDT (WBS Ampliada)
    ws3 = wb.create_sheet(title="EDT (WBS)")
    ws3.views.sheetView[0].showGridLines = True
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

    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output)
    return output.getvalue()

# ==========================================
# 4. INTERFAZ PRINCIPAL Y CONMUTADOR DE TEMA
# ==========================================
st.title("✦ MatrixDev ✦")
st.markdown(f"<p style='text-align: center; color: {sub_color};'>Gestión y Arquitectura de Proyectos</p>", unsafe_allow_html=True)

# BOTONES DE CAMBIO DE TEMA VISUAL CENTRADOS
col_sp1, col_btn1, col_btn2, col_btn3, col_sp2 = st.columns([1.5, 1, 1, 1, 1.5])

with col_btn1:
    if st.button("Vino Tinto", use_container_width=True):
        st.session_state.tema = "vino"
        st.rerun()

with col_btn2:
    if st.button("Gris Oscuro", use_container_width=True):
        st.session_state.tema = "oscuro"
        st.rerun()

with col_btn3:
    if st.button("Claro", use_container_width=True):
        st.session_state.tema = "claro"
        st.rerun()

st.write("---")

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab_full = st.tabs([
    "1. Requisitos", 
    "2. Modelado", 
    "3. Gestión", 
    "4. Monte Carlo", 
    "5. Flujos", 
    "6. EDT (WBS)", 
    "7. Riesgos", 
    "8. Reporte MementoMori",
    "📦 Descarga Full"
])

# ------------------------------------------
# TAB 1: REQUISITOS
# ------------------------------------------
with tab1:
    st.markdown("### Definición del Sistema")
    c_p1, c_p2 = st.columns(2)
    st.session_state.nombre_proj = c_p1.text_input("Nombre del Proyecto", st.session_state.nombre_proj)
    st.session_state.integrantes = c_p2.text_input("Integrantes / Equipo", st.session_state.integrantes)
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

    # DESCARGA INDIVIDUAL - TAB 1
    st.write("---")
    st.markdown("#### 📥 Exportar Punto 1: Requisitos")
    c_dl1, c_dl2 = st.columns(2)
    with c_dl1:
        st.download_button(
            label="📊 Descargar Requisitos en Excel (.xlsx)",
            data=generar_excel_requisitos(),
            file_name=f"Requisitos_{st.session_state.nombre_proj.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with c_dl2:
        st.download_button(
            label="📄 Descargar Documento ERS en Word (.docx)",
            data=generar_word_ers(),
            file_name=f"ERS_Requisitos_{st.session_state.nombre_proj.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# ------------------------------------------
# TAB 2: MODELADO
# ------------------------------------------
with tab2:
    st.markdown("### Arquitectura Visual")
    default_mermaid = """graph TD
    A[Inicio] --> B{Validar}
    B -->|Sí| C[Éxito]
    B -->|No| D[Error]"""
    
    codigo_mermaid = st.text_area("Sintaxis Mermaid", value=default_mermaid, height=120)
    st_mermaid(codigo_mermaid)

    # DESCARGA INDIVIDUAL - TAB 2
    st.write("---")
    st.markdown("#### 📥 Exportar Punto 2: Diagrama de Modelado")
    col_d2_1, col_d2_2 = st.columns([1, 1])
    with col_d2_1:
        st.download_button(
            label="📌 Descargar Código Diagrama Mermaid (.mmd)",
            data=codigo_mermaid.encode('utf-8'),
            file_name=f"Diagrama_Modelado_{st.session_state.nombre_proj.replace(' ', '_')}.mmd",
            mime="text/plain",
            use_container_width=True
        )

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

    # DESCARGA INDIVIDUAL - TAB 3
    st.write("---")
    st.markdown("#### 📥 Exportar Punto 3: Tablero Kanban")
    c_dk1, c_dk2 = st.columns(2)
    with c_dk1:
        st.download_button(
            label="📊 Descargar Tablero Kanban en Excel (.xlsx)",
            data=generar_excel_kanban(),
            file_name=f"Kanban_{st.session_state.nombre_proj.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with c_dk2:
        df_k_exp = pd.DataFrame(st.session_state.kanban_tasks)
        st.download_button(
            label="📄 Descargar Tareas en CSV (.csv)",
            data=df_k_exp.to_csv(index=False).encode('utf-8'),
            file_name=f"Kanban_{st.session_state.nombre_proj.replace(' ', '_')}.csv",
            mime="text/csv",
            use_container_width=True
        )

# ------------------------------------------
# TAB 4: MONTE CARLO (AVANZADO, CONFIGURABLE Y COMPLETO)
# ------------------------------------------
with tab4:
    st.markdown("### 🎲 Simulación Monte Carlo de Tiempos y Presupuestos")
    st.caption("Motor de análisis probabilístico de riesgo, factibilidad y capacidad proyectada.")

    # GUÍA EXPLICATIVA INTUITIVA
    with st.expander("📖 **Guía Intuitiva: ¿Cómo entender la Simulación Monte Carlo?**", expanded=False):
        st.markdown("""
        **¿Qué es Monte Carlo?** Es un método estadístico que ejecuta miles de escenarios simulados (iteraciones) variando los tiempos y costos según la incertidumbre real del proyecto.
        
        * **Modelos de Distribución:**
            * 📐 **Triangular:** Basada en 3 estimaciones simples (Optimista, Más Probable, Pesimista).
            * 📊 **Beta-PERT:** Estándar de la metodología PMBOK. Da más peso visual al valor 'Más Probable'.
            * 🔔 **Normal / Gaussiana:** Para procesos maduros con comportamiento en campana (Media y Desviación).
            * 🟦 **Uniforme:** Incertidumbre total, cualquier duración entre el Mínimo y Máximo es igual de probable.
        * **Percentiles Clave:**
            * **P50 (Mediana):** 50% de probabilidad de terminar antes de este tiempo. Es como tirar una moneda.
            * **P85 (Compromiso Recomendado):** Estándar de la industria. 85% de probabilidad de éxito.
            * **P95 (Escenario Pesimista):** Cubre el 95% de todos los imprevistos y bloqueos posibles.
        """)

    st.write("---")

    # CONFIGURACIÓN COMPLETA DE PARÁMETROS
    c_sec1, c_sec2, c_sec3 = st.columns([1.2, 1.2, 1])

    with c_sec1:
        st.markdown("##### 📐 1. Modelo Estadístico & Tiempos")
        tipo_dist = st.selectbox(
            "Distribución Probabilística",
            ["Triangular", "Beta-PERT", "Normal", "Uniforme"],
            key="mc_dist_type",
            help="Selecciona el modelo matemático para generar la variabilidad aleatoria"
        )

        if tipo_dist in ["Triangular", "Beta-PERT"]:
            mc_opt = st.number_input("Días Optimista (Mínimo)", value=10, min_value=1, key="mc_opt")
            mc_prob = st.number_input("Días Más Probable (Moda)", value=20, min_value=1, key="mc_prob")
            mc_pes = st.number_input("Días Pesimista (Máximo)", value=45, min_value=1, key="mc_pes")
        elif tipo_dist == "Normal":
            mc_mean = st.number_input("Duración Promedio (Días)", value=22.0, min_value=1.0, key="mc_mean")
            mc_std = st.number_input("Desviación Estándar (Días)", value=5.0, min_value=0.1, key="mc_std")
        elif tipo_dist == "Uniforme":
            mc_min = st.number_input("Días Mínimos", value=12, min_value=1, key="mc_min_u")
            mc_max = st.number_input("Días Máximos", value=40, min_value=1, key="mc_max_u")

    with c_sec2:
        st.markdown("##### 💰 2. Modelo Financiero / Costos")
        costo_fijo = st.number_input("Costo Fijo Base del Proyecto ($)", value=100000, step=10000, key="mc_cfijo")
        costo_diario = st.number_input("Costo Operativo por Día ($/Día)", value=15000, step=1000, key="mc_cdiario")
        costo_meta = st.number_input("Presupuesto Límite Meta ($)", value=500000, step=25000, key="mc_costo_meta")

    with c_sec3:
        st.markdown("##### 📅 3. Metas & Calendario")
        fecha_inicio = st.date_input("Fecha de Inicio", value=pd.to_datetime("today"), key="mc_f_inicio")
        dias_meta = st.number_input("Días Meta (SLA Tiempo)", value=25, min_value=1, key="mc_dias_meta")
        percentil_target = st.slider("Percentil de Compromiso (%)", 50, 99, 85, step=5, key="mc_target_perc")

    # MUESTREO Y CONTROL
    st.markdown("##### ⚙️ 4. Muestreo & Motor de Cálculo")
    col_s1, col_s2 = st.columns([3, 1])
    simulaciones = col_s1.slider("Número de Simulaciones (Iteraciones)", 500, 20000, 5000, step=500, key="mc_sims")
    usar_semilla = col_s2.checkbox("Fijar Semilla (Reproducible)", value=True, key="mc_seed")

    # VALIDACIÓN Y GENERACIÓN DE SIMULACIONES
    valido = True
    if tipo_dist in ["Triangular", "Beta-PERT"] and not (mc_opt <= mc_prob <= mc_pes):
        st.error("Error: Se debe cumplir que Días Optimista ≤ Días Más Probable ≤ Días Pesimista.")
        valido = False
    elif tipo_dist == "Uniforme" and not (mc_min < mc_max):
        st.error("Error: Se debe cumplir que Días Mínimos < Días Máximos.")
        valido = False

    if valido:
        if usar_semilla:
            np.random.seed(42)

        # Generación aleatoria de duraciones según la distribución
        if tipo_dist == "Triangular":
            datos_dias = np.random.triangular(left=mc_opt, mode=mc_prob, right=mc_pes, size=simulaciones)
        elif tipo_dist == "Beta-PERT":
            # Parámetros Beta-PERT
            a = 1 + 4 * (mc_prob - mc_opt) / (mc_pes - mc_opt) if (mc_pes > mc_opt) else 1
            b = 1 + 4 * (mc_pes - mc_prob) / (mc_pes - mc_opt) if (mc_pes > mc_opt) else 1
            datos_dias = mc_opt + np.random.beta(a, b, size=simulaciones) * (mc_pes - mc_opt)
        elif tipo_dist == "Normal":
            datos_dias = np.random.normal(loc=mc_mean, scale=mc_std, size=simulaciones)
            datos_dias = np.clip(datos_dias, 1.0, None)  # Evitar días menores a 1
        elif tipo_dist == "Uniforme":
            datos_dias = np.random.uniform(low=mc_min, high=mc_max, size=simulaciones)

        # Cálculo de costos por simulación
        # Se agrega un margen de variabilidad aleatoria del 5% al costo diario por imprevistos
        variabilidad_costo = np.random.normal(1.0, 0.05, size=simulaciones)
        datos_costos = costo_fijo + (datos_dias * costo_diario * variabilidad_costo)

        # CÁLCULO DE METRICAS PRINCIPALES
        p50 = np.percentile(datos_dias, 50)
        p_target = np.percentile(datos_dias, percentil_target)
        p95 = np.percentile(datos_dias, 95)

        p50_cost = np.percentile(datos_costos, 50)
        p_target_cost = np.percentile(datos_costos, percentil_target)
        p95_cost = np.percentile(datos_costos, 95)

        exitos_tiempo = (datos_dias <= dias_meta).sum()
        tasa_exito_tiempo = (exitos_tiempo / simulaciones) * 100
        riesgo_retraso = 100 - tasa_exito_tiempo

        exitos_costo = (datos_costos <= costo_meta).sum()
        tasa_exito_costo = (exitos_costo / simulaciones) * 100

        exitos_ambos = ((datos_dias <= dias_meta) & (datos_costos <= costo_meta)).sum()
        tasa_exito_conjunta = (exitos_ambos / simulaciones) * 100

        fecha_p50 = fecha_inicio + pd.Timedelta(days=int(np.round(p50)))
        fecha_target = fecha_inicio + pd.Timedelta(days=int(np.round(p_target)))
        fecha_p95 = fecha_inicio + pd.Timedelta(days=int(np.round(p95)))

        st.write("---")
        st.markdown("#### 📊 Diagnóstico Ejecutivo & Veredicto de Viabilidad")

        # Tarjeta dinámica de veredicto
        if tasa_exito_conjunta >= 80:
            v_color, v_title = "🟢", "PROYECTO ALTAMENTE VIABLE"
            v_msg = f"Alta certidumbre: {tasa_exito_tiempo:.1f}% de éxito en plazo y {tasa_exito_costo:.1f}% en presupuesto."
        elif tasa_exito_conjunta >= 50:
            v_color, v_title = "🟡", "VIABLE CON RIESGO MODERADO"
            v_msg = f"Factible pero ajustado: Se recomienda reservar una contingencia hasta el percentil P{percentil_target} ({p_target:.1f} días)."
        else:
            v_color, v_title = "🔴", "ALTO RIESGO DE INCUMPLIMIENTO"
            v_msg = f"Incertidumbre elevada: El riesgo de retraso es del {riesgo_retraso:.1f}%. Ajusta la meta a {p_target:.0f} días o amplía el presupuesto."

        st.info(f"### {v_color} {v_title}\n{v_msg}")

        # Tarjetas de Métricas
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("🎯 Éxito en Tiempo", f"{tasa_exito_tiempo:.1f}%", f"Meta: {dias_meta}d")
        m2.metric("⚠️ Riesgo de Retraso", f"{riesgo_retraso:.1f}%", f"Riesgo de Tiempo", delta_color="inverse")
        m3.metric("💰 Éxito en Presupuesto", f"{tasa_exito_costo:.1f}%", f"Meta: ${costo_meta:,.0f}")
        m4.metric(f"📆 Entrega P{percentil_target}", fecha_target.strftime("%d/%m/%Y"), f"{p_target:.1f} días")
        m5.metric(f"💵 Presupuesto P{percentil_target}", f"${p_target_cost:,.0f}", f"P50: ${p50_cost:,.0f}")

        st.write("---")

        # GRÁFICOS INTERACTIVOS CON TABULADORES
        st.markdown("#### 📈 Panel de Gráficos y Análisis de Distribución")
        v_tab1, v_tab2, v_tab3 = st.tabs(["📊 Histograma (PDF)", "📈 Curva S Acumulada (CDF)", "💥 Dispersión Tiempo vs Costo"])

        with v_tab1:
            fig_hist = go.Figure()
            fig_hist.add_trace(go.Histogram(
                x=datos_dias,
                nbinsx=40,
                name="Simulaciones",
                marker=dict(color=graph_bar, opacity=0.75, line=dict(color=heading_color, width=1)),
                hovertemplate="Duración: %{x:.1f} días<br>Escenarios: %{y}<extra></extra>"
            ))
            fig_hist.add_vline(x=dias_meta, line_dash="dash", line_color="#ef4444", line_width=2,
                               annotation_text=f"Meta ({dias_meta}d)", annotation_position="top left")
            fig_hist.add_vline(x=p50, line_dash="dot", line_color="#f59e0b", line_width=2,
                               annotation_text=f"P50 ({p50:.1f}d)", annotation_position="top right")
            fig_hist.add_vline(x=p_target, line_dash="solid", line_color="#10b981", line_width=2.5,
                               annotation_text=f"P{percentil_target} ({p_target:.1f}d)", annotation_position="top right")

            fig_hist.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=graph_text,
                xaxis=dict(title="Días Estimados de Ejecución", gridcolor=grid_color),
                yaxis=dict(title="Frecuencia (N° Muestras)", gridcolor=grid_color),
                margin=dict(l=20, r=20, t=30, b=20), height=360, showlegend=False
            )
            st.plotly_chart(fig_hist, use_container_width=True)

        with v_tab2:
            # Curva S (CDF)
            dias_sorted = np.sort(datos_dias)
            prob_acumulada = np.arange(1, len(dias_sorted) + 1) / len(dias_sorted) * 100

            fig_cdf = go.Figure()
            fig_cdf.add_trace(go.Scatter(
                x=dias_sorted, y=prob_acumulada, mode='lines',
                line=dict(color=graph_bar, width=3),
                hovertemplate="Días: %{x:.1f}<br>Probabilidad Cumplimiento: %{y:.1f}%<extra></extra>"
            ))
            fig_cdf.add_hline(y=percentil_target, line_dash="dash", line_color="#10b981",
                              annotation_text=f"Objetivo {percentil_target}%", annotation_position="bottom right")
            fig_cdf.add_vline(x=dias_meta, line_dash="dash", line_color="#ef4444",
                              annotation_text=f"Meta ({dias_meta}d)", annotation_position="top left")

            fig_cdf.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=graph_text,
                xaxis=dict(title="Días de Ejecución", gridcolor=grid_color),
                yaxis=dict(title="Probabilidad Acumulada de Éxito (%)", gridcolor=grid_color, range=[0, 105]),
                margin=dict(l=20, r=20, t=30, b=20), height=360, showlegend=False
            )
            st.plotly_chart(fig_cdf, use_container_width=True)

        with v_tab3:
            # Scatter Plot de Tiempo vs Costo
            # Muestrear máximo 1000 puntos para rendimiento visual limpio
            indices_sample = np.random.choice(simulaciones, size=min(1000, simulaciones), replace=False)
            df_scatter = pd.DataFrame({
                "Días": datos_dias[indices_sample],
                "Costo": datos_costos[indices_sample]
            })

            fig_scatter = px.scatter(
                df_scatter, x="Días", y="Costo",
                color_discrete_sequence=[graph_bar],
                opacity=0.6,
                labels={"Días": "Duración (Días)", "Costo": "Costo Total ($)"}
            )
            fig_scatter.add_vline(x=dias_meta, line_dash="dash", line_color="#ef4444", annotation_text="Límite Tiempo")
            fig_scatter.add_hline(y=costo_meta, line_dash="dash", line_color="#ef4444", annotation_text="Límite Costo")

            fig_scatter.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=graph_text,
                xaxis=dict(gridcolor=grid_color), yaxis=dict(gridcolor=grid_color),
                margin=dict(l=20, r=20, t=30, b=20), height=360
            )
            st.plotly_chart(fig_scatter, use_container_width=True)

        st.write("---")

        # INSPECCIÓN INDIVIDUAL (PESCAR ITERACIÓN) & ANÁLISIS DE EXTREMOS
        col_p1, col_p2 = st.columns([1.2, 1])

        with col_p1:
            st.markdown("#### 🎣 Pescar e Inspeccionar Iteración")
            st.caption("Selecciona una iteración puntual para auditar cómo se comporta una muestra aleatoria específica.")

            iter_idx = st.slider("Número de Iteración", 1, simulaciones, 100, key="mc_slider_iter")
            val_iter_dias = datos_dias[iter_idx - 1]
            val_iter_costo = datos_costos[iter_idx - 1]
            fecha_iter = fecha_inicio + pd.Timedelta(days=int(np.round(val_iter_dias)))

            cumple_tiempo = val_iter_dias <= dias_meta
            cumple_costo = val_iter_costo <= costo_meta

            st.markdown(f"""
            * 📌 **Iteración Auditada:** **#{iter_idx}** de {simulaciones}
            * ⏱️ **Duración Resultante:** **{val_iter_dias:.1f} Días**
            * 💵 **Costo Estimado:** **${val_iter_costo:,.0f}**
            * 📆 **Fecha Proyectada de Término:** **{fecha_iter.strftime('%d/%m/%Y')}**
            * 🏁 **Resultado Tiempo:** {"🟢 Cumple Meta" if cumple_tiempo else "🔴 Retrasado"}
            * 💳 **Resultado Presupuesto:** {"🟢 Dentro de Presupuesto" if cumple_costo else "🔴 Sobrecosto"}
            """)

        with col_p2:
            st.markdown("#### 🚀 Proyección Anual (1 Año) y Capacidad")
            dias_laborables_opcion = st.radio("Tipo de Calendario Anual", ["Días Hábiles (250 días/año)", "Días Corridos (365 días/año)"], horizontal=True, key="mc_cal_type")
            dias_ano_base = 250 if "250" in dias_laborables_opcion else 365

            ciclos_ano_p50 = dias_ano_base / p50 if p50 > 0 else 0
            ciclos_ano_target = dias_ano_base / p_target if p_target > 0 else 0

            st.markdown(f"""
            * 🔄 **Capacidad Mediana (P50):** **{ciclos_ano_p50:.1f} Proyectos / Año**
            * 🛡️ **Capacidad Segura (P{percentil_target}):** **{ciclos_ano_target:.1f} Proyectos / Año**
            * 💡 **Margen de Seguridad:** Utilizar el percentil P{percentil_target} garantiza que el equipo cumplirá sus metas anuales con solo un **{100 - percentil_target}% de riesgo de imprevistos**.
            """)

        st.write("---")

        # TABLA DE CASOS EXTREMOS (TOP 5 MEJORES / PEORES)
        st.markdown("#### 🔍 Tabla de Auditoría: Escenarios Extremos")
        df_all_sims = pd.DataFrame({
            "N° Iteración": np.arange(1, simulaciones + 1),
            "Duración (Días)": np.round(datos_dias, 1),
            "Costo Total ($)": np.round(datos_costos, 0),
            "Cumple Tiempo": datos_dias <= dias_meta,
            "Cumple Costo": datos_costos <= costo_meta
        })

        t_col1, t_col2 = st.columns(2)
        with t_col1:
            st.markdown("**Top 5 Escenarios Más Rápidos (Optimistas)**")
            st.dataframe(df_all_sims.sort_values(by="Duración (Días)").head(5), use_container_width=True, hide_index=True)

        with t_col2:
            st.markdown("**Top 5 Escenarios Más Lentos (Pesimistas)**")
            st.dataframe(df_all_sims.sort_values(by="Duración (Días)", ascending=False).head(5), use_container_width=True, hide_index=True)

        # DESCARGA INDIVIDUAL - TAB 4
        st.write("---")
        st.markdown("#### 📥 Exportar Punto 4: Resultados Monte Carlo")
        c_dmc1, _ = st.columns([1, 1])
        with c_dmc1:
            st.download_button(
                label="📊 Descargar Simulación Monte Carlo en Excel (.xlsx)",
                data=generar_excel_montecarlo(datos_dias, datos_costos),
                file_name=f"MonteCarlo_{st.session_state.nombre_proj.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

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
    
    codigo_actualizado = st.text_area(
        "Código del Diagrama (Editable)", 
        value=st.session_state.flujo_codigo, 
        height=180,
        key="txt_flujo_code"
    )
    st.session_state.flujo_codigo = codigo_actualizado
    
    st.markdown("#### Vista Previa del Diagrama")
    st_mermaid(st.session_state.flujo_codigo)

    # DESCARGA INDIVIDUAL - TAB 5
    st.write("---")
    st.markdown("#### 📥 Exportar Punto 5: Diagrama de Flujos")
    c_df1, _ = st.columns([1, 1])
    with c_df1:
        st.download_button(
            label="📌 Descargar Diagrama de Flujo (.mmd)",
            data=st.session_state.flujo_codigo.encode('utf-8'),
            file_name=f"Flujo_Secuencia_{st.session_state.nombre_proj.replace(' ', '_')}.mmd",
            mime="text/plain",
            use_container_width=True
        )

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
                font_color=graph_text,
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

    # DESCARGA INDIVIDUAL - TAB 6
    st.write("---")
    st.markdown("#### 📥 Exportar Punto 6: Estructura EDT (WBS)")
    c_dedt1, c_dedt2 = st.columns(2)
    with c_dedt1:
        st.download_button(
            label="📊 Descargar EDT en Excel (.xlsx)",
            data=generar_excel_edt(),
            file_name=f"EDT_{st.session_state.nombre_proj.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with c_dedt2:
        df_e_exp = pd.DataFrame(st.session_state.edt_list)
        st.download_button(
            label="📄 Descargar EDT en CSV (.csv)",
            data=df_e_exp.to_csv(index=False).encode('utf-8'),
            file_name=f"EDT_{st.session_state.nombre_proj.replace(' ', '_')}.csv",
            mime="text/csv",
            use_container_width=True
        )

# ------------------------------------------
# TAB 7: RIESGOS
# ------------------------------------------
with tab7:
    st.markdown("### Matriz de Riesgos (5x5 Minimalista)")

    df_r = pd.DataFrame(st.session_state.riesgos_list)
    
    if not df_r.empty:
        df_r["Severidad"] = df_r["Probabilidad"] * df_r["Impacto"]
        df_r["Nivel"] = df_r["Severidad"].apply(clasificar_riesgo)
    else:
        df_r = pd.DataFrame(columns=["Riesgo", "Probabilidad", "Impacto", "Severidad", "Nivel"])

    # MÉTRICAS RÁPIDAS
    r_total = len(df_r)
    r_criticos = len(df_r[df_r["Severidad"] >= 12]) if not df_r.empty else 0
    r_prom = df_r["Severidad"].mean() if not df_r.empty else 0.0

    mr1, mr2, mr3 = st.columns(3)
    mr1.metric("⚠ Total de Riesgos", r_total)
    mr2.metric("🔥 Riesgos Altos / Críticos", r_criticos)
    mr3.metric("📊 Severidad Promedio", f"{r_prom:.1f} / 25")

    st.write("---")

    col_matriz, col_form = st.columns([1.3, 1])

    with col_matriz:
        st.markdown("#### 🗺️ Mapa de Calor (Probabilidad vs Impacto)")
        
        # Construir matriz de fondo 5x5
        z_grid = [[i * j for j in range(1, 6)] for i in range(1, 6)]

        fig_matrix = go.Figure()

        # Fondo codificado por nivel de riesgo
        fig_matrix.add_trace(go.Heatmap(
            z=z_grid,
            x=[1, 2, 3, 4, 5],
            y=[1, 2, 3, 4, 5],
            colorscale=[
                [0.0, "rgba(46, 125, 50, 0.35)"],   # Verde (Bajo)
                [0.35, "rgba(251, 192, 45, 0.35)"], # Amarillo (Medio)
                [0.6, "rgba(245, 124, 0, 0.35)"],   # Naranja (Alto)
                [1.0, "rgba(211, 47, 47, 0.4)"]     # Rojo (Crítico)
            ],
            showscale=False,
            hoverinfo="skip"
        ))

        # Puntos de riesgo
        if not df_r.empty:
            fig_matrix.add_trace(go.Scatter(
                x=df_r["Probabilidad"],
                y=df_r["Impacto"],
                mode="markers+text",
                text=df_r["Riesgo"],
                textposition="top center",
                textfont=dict(color=graph_text, size=11),
                marker=dict(
                    size=16,
                    color=df_r["Severidad"],
                    colorscale="Reds",
                    line=dict(width=1.5, color="#ffffff")
                ),
                hovertemplate="<b>%{text}</b><br>Probabilidad: %{x}<br>Impacto: %{y}<br>Severidad: %{marker.color}<extra></extra>"
            ))

        fig_matrix.update_layout(
            xaxis=dict(title="Probabilidad (1-5)", range=[0.5, 5.5], dtick=1, gridcolor=grid_color, zeroline=False),
            yaxis=dict(title="Impacto (1-5)", range=[0.5, 5.5], dtick=1, gridcolor=grid_color, zeroline=False),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font_color=graph_text,
            margin=dict(l=30, r=20, t=10, b=30),
            height=380
        )

        st.plotly_chart(fig_matrix, use_container_width=True)

    with col_form:
        st.markdown("#### ➕ Registrar Nuevo Riesgo")
        with st.form("form_add_riesgo", clear_on_submit=True):
            r_nombre = st.text_input("Nombre / Descripción del Riesgo", placeholder="Ej: Retraso en entregables")
            
            c_p, c_i = st.columns(2)
            r_prob = c_p.slider("Probabilidad (1 a 5)", min_value=1, max_value=5, value=3)
            r_imp = c_i.slider("Impacto (1 a 5)", min_value=1, max_value=5, value=3)

            btn_add_r = st.form_submit_button("Agregar Riesgo", use_container_width=True)
            if btn_add_r:
                if r_nombre.strip():
                    st.session_state.riesgos_list.append({
                        "Riesgo": r_nombre.strip(),
                        "Probabilidad": int(r_prob),
                        "Impacto": int(r_imp)
                    })
                    st.rerun()
                else:
                    st.error("Escribe un nombre para el riesgo.")

    st.write("---")
    st.markdown("#### 📝 Inventario y Edición de Riesgos")

    if not st.session_state.riesgos_list:
        st.info("No hay riesgos registrados.")
    else:
        r_del_idx = None
        
        # Cabecera
        rh1, rh2, rh3, rh4, rh5, rh6 = st.columns([3, 1.2, 1.2, 1.2, 1.5, 0.6])
        rh1.markdown("**Riesgo**")
        rh2.markdown("**Probabilidad**")
        rh3.markdown("**Impacto**")
        rh4.markdown("**Severidad**")
        rh5.markdown("**Nivel**")
        rh6.markdown("")

        for idx, item in enumerate(st.session_state.riesgos_list):
            rc1, rc2, rc3, rc4, rc5, rc6 = st.columns([3, 1.2, 1.2, 1.2, 1.5, 0.6])
            
            r_txt = rc1.text_input(f"rt_{idx}", value=item["Riesgo"], key=f"r_name_{idx}", label_visibility="collapsed")
            r_p_val = rc2.number_input(f"rp_{idx}", value=int(item["Probabilidad"]), min_value=1, max_value=5, key=f"r_p_{idx}", label_visibility="collapsed")
            r_i_val = rc3.number_input(f"ri_{idx}", value=int(item["Impacto"]), min_value=1, max_value=5, key=f"r_i_{idx}", label_visibility="collapsed")
            
            sev_val = r_p_val * r_i_val
            nivel_str = clasificar_riesgo(sev_val)
            
            rc4.markdown(f"**{sev_val}**")
            rc5.markdown(nivel_str)

            st.session_state.riesgos_list[idx]["Riesgo"] = r_txt
            st.session_state.riesgos_list[idx]["Probabilidad"] = r_p_val
            st.session_state.riesgos_list[idx]["Impacto"] = r_i_val

            if rc6.button("🗑️", key=f"del_r_{idx}"):
                r_del_idx = idx

        if r_del_idx is not None:
            st.session_state.riesgos_list.pop(r_del_idx)
            st.rerun()

    # DESCARGA INDIVIDUAL - TAB 7
    st.write("---")
    st.markdown("#### 📥 Exportar Punto 7: Matriz de Riesgos")
    c_dr1, _ = st.columns([1, 1])
    with c_dr1:
        st.download_button(
            label="📊 Descargar Matriz de Riesgos en Excel (.xlsx)",
            data=generar_excel_riesgos(),
            file_name=f"Matriz_Riesgos_{st.session_state.nombre_proj.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ------------------------------------------
# TAB 8: REPORTE MEMENTOMORI
# ------------------------------------------
with tab8:
    st.markdown("### 📜 Reporte MementoMori (Análisis de Puntos Críticos)")
    st.write(
        "Este reporte genera un documento en formato Word (.docx) que consolida un análisis crítico y exhaustivo de todos los componentes "
        "del proyecto: requisitos, diagramas de flujo/arquitectura, estimación Monte Carlo, EDT y matriz de riesgos."
    )
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.session_state.integrantes = st.text_input("Integrantes / Responsables", st.session_state.integrantes, key="mem_tab_integrantes")
    with col_m2:
        st.session_state.nombre_proj = st.text_input("Nombre del Proyecto", st.session_state.nombre_proj, key="mem_tab_nombre")

    st.markdown("#### 🔍 Cobertura del Análisis MementoMori:")
    st.markdown("""
    * **1. Introducción y Alcance:** Evaluación de viabilidad, alineación estratégica y prevención de Scope Creep.
    * **2. Requisitos (RF / RNF):** Análisis de vulnerabilidad en los Requisitos No Funcionales (latencia, seguridad).
    * **3. Arquitectura y Flujos:** Interpretación del acoplamiento de componentes y puntos de fallo en la secuencia de usuario.
    * **4. Gestión Kanban:** Identificación del trabajo en proceso (WIP) y cuellos de botella.
    * **5. EDT (WBS):** Evaluación de actividades críticas con alto impacto en costo y horas.
    * **6. Simulación Monte Carlo:** Diagnóstico de niveles de confianza (P85 / P95) y análisis de dispersión pesimista.
    * **7. Matriz de Riesgos:** Filosofía MementoMori ('Planificar anticipadamente ante fallos') y mitigación activa.
    """)

    word_mementomori_bytes = generar_word_mementomori()

    st.write("")
    col_btn_center, _ = st.columns([2, 1])
    with col_btn_center:
        st.download_button(
            label="📄 Descargar Reporte MementoMori (.docx)",
            data=word_mementomori_bytes,
            file_name=f"Reporte_MementoMori_{st.session_state.nombre_proj.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# ------------------------------------------
# TAB 9: DESCARGA FULL
# ------------------------------------------
with tab_full:
    st.markdown("### 📦 Exportación General Consolidada")
    st.write("Descarga un archivo ZIP con todos los entregables generados (Reporte MementoMori, ERS Word, Excel y CSVs).")
    
    word_doc = generar_word_ers()
    word_mementomori = generar_word_mementomori()
    excel_doc = generar_excel_estilizado()
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"Reporte_MementoMori_{st.session_state.nombre_proj}.docx", word_mementomori)
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
            label="⬇ Descargar Reporte Full (.ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"Reporte_Full_{st.session_state.nombre_proj.replace(' ', '_')}.zip",
            mime="application/zip",
            use_container_width=True
        )
