import io
import json
import zipfile
from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from streamlit_mermaid import st_mermaid

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from docx import Document
from docx.shared import Pt, RGBColor

# =====================================================================
# 1. CONFIGURACIÓN GENERAL
# =====================================================================
st.set_page_config(page_title="MatrixDev · Tesis", layout="wide", page_icon="🍷",
                   initial_sidebar_state="collapsed")


def _version_streamlit():
    try:
        return tuple(int(p) for p in st.__version__.split(".")[:2])
    except Exception:
        return (1, 0)


# Streamlit >= 1.50 reemplazó use_container_width por width="stretch"
ANCHO = {"width": "stretch"} if _version_streamlit() >= (1, 50) else {"use_container_width": True}

# ---------------------------------------------------------------------
# Paletas de color por tema
#   claro       → celeste y blanco
#   oscuro      → negro y blanco
#   informatico → negro con rojo neón
# ---------------------------------------------------------------------
TEMAS = {
    "claro": {
        "nombre": "Claro", "base": "light", "icono": "☀️", "bg": "#f2f9fe", "card": "#ffffff", "text": "#0b2a3f", "sub": "#56768c",
        "accent": "#0ea5e9", "accent2": "#0369a1", "border": "#d3e9f6", "grid": "#e5f2fa",
        "input": "#ffffff", "ok": "#10b981", "warn": "#f59e0b", "bad": "#ef4444", "on_accent": "#ffffff",
        "fuente": "'Plus Jakarta Sans', 'Segoe UI', Helvetica, Arial, sans-serif", "neon": False,
        "fondo": "radial-gradient(1200px 500px at 85% -10%, rgba(14,165,233,0.14), transparent 60%), "
                 "radial-gradient(900px 400px at -10% 10%, rgba(56,189,248,0.10), transparent 60%)",
    },
    "oscuro": {
        "nombre": "Oscuro", "base": "dark", "icono": "🌙", "bg": "#000000", "card": "#0b0b0b", "text": "#ffffff", "sub": "#a1a1a1",
        "accent": "#ffffff", "accent2": "#bdbdbd", "border": "#262626", "grid": "#191919",
        "input": "#121212", "ok": "#4ade80", "warn": "#facc15", "bad": "#f87171", "on_accent": "#000000",
        "fuente": "'Plus Jakarta Sans', 'Segoe UI', Helvetica, Arial, sans-serif", "neon": False,
        "fondo": "radial-gradient(1000px 420px at 80% -10%, rgba(255,255,255,0.06), transparent 60%)",
    },
    "informatico": {
        "nombre": "Informático", "base": "dark", "icono": "💻", "bg": "#050505", "card": "#0d090a", "text": "#f5f5f5", "sub": "#a08f92",
        "accent": "#ff1f4b", "accent2": "#ff6b86", "border": "#3a0d17", "grid": "#1c0a0f",
        "input": "#110a0c", "ok": "#2bff88", "warn": "#ffb020", "bad": "#ff3b3b", "on_accent": "#050505",
        "fuente": "'JetBrains Mono', 'Fira Code', Consolas, monospace", "neon": True,
        "fondo": "linear-gradient(rgba(255,31,75,0.045) 1px, transparent 1px) 0 0 / 32px 32px, "
                 "linear-gradient(90deg, rgba(255,31,75,0.045) 1px, transparent 1px) 0 0 / 32px 32px, "
                 "radial-gradient(900px 420px at 85% -10%, rgba(255,31,75,0.16), transparent 60%)",
    },
}

if st.session_state.get("tema") not in TEMAS:
    st.session_state.tema = "claro"
T = TEMAS[st.session_state.tema]
FUENTE = T["fuente"]


def aplicar_tema_nativo(tema):
    """Sincroniza el tema nativo de Streamlit (sliders, toggles, tablas, calendarios) con el tema elegido."""
    try:
        from streamlit import config as _cfg
        opciones = {
            "theme.base": "light" if tema["base"] == "light" else "dark",
            "theme.primaryColor": tema["accent"],
            "theme.backgroundColor": tema["bg"],
            "theme.secondaryBackgroundColor": tema["input"],
            "theme.textColor": tema["text"],
        }
        cambio = False
        for clave, valor in opciones.items():
            if _cfg.get_option(clave) != valor:
                _cfg.set_option(clave, valor)
                cambio = True
        return cambio
    except Exception:
        return False  # versiones antiguas: el CSS de abajo cubre lo esencial


def rgba(hex_color, alpha):
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


# Si el tema nativo cambió, se recarga una vez para que Streamlit lo aplique de inmediato
if aplicar_tema_nativo(T) and st.session_state.get("_tema_sincronizado") != st.session_state.tema:
    st.session_state._tema_sincronizado = st.session_state.tema
    st.rerun()
st.session_state._tema_sincronizado = st.session_state.tema


def cambiar_tema(nuevo):
    st.session_state.tema = nuevo


# Efectos de brillo solo para el tema informático
GLOW = f"0 0 18px {rgba(T['accent'], 0.45)}, 0 0 2px {rgba(T['accent'], 0.9)}" if T["neon"] else f"0 10px 28px {rgba(T['accent'], 0.16)}"
GLOW_TXT = f"0 0 10px {rgba(T['accent'], 0.75)}, 0 0 24px {rgba(T['accent'], 0.35)}" if T["neon"] else "none"
SOMBRA = f"0 1px 2px {rgba('#000000', 0.04)}, 0 4px 16px {rgba(T['accent'], 0.06)}" if st.session_state.tema == "claro" else "none"

# ---------------------------------------------------------------------
# CSS global
# ---------------------------------------------------------------------
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap');

/* ---------- Animaciones ---------- */
@keyframes mdFadeUp {{ from {{ opacity: 0; transform: translateY(10px); }} to {{ opacity: 1; transform: none; }} }}
@keyframes mdFade {{ from {{ opacity: 0; }} to {{ opacity: 1; }} }}
@keyframes mdPop {{ 0% {{ opacity: 0; transform: scale(0.2); }} 70% {{ transform: scale(1.15); }} 100% {{ opacity: 1; transform: scale(1); }} }}
@keyframes mdSlide {{ from {{ background-position: 0% 50%; }} to {{ background-position: 200% 50%; }} }}
@keyframes mdPulse {{ 0%, 100% {{ box-shadow: 0 0 0 0 {rgba(T['accent'], 0.45)}; }} 50% {{ box-shadow: 0 0 0 7px {rgba(T['accent'], 0)}; }} }}
@keyframes mdBlink {{ 0%, 49% {{ opacity: 1; }} 50%, 100% {{ opacity: 0; }} }}
@keyframes mdFlicker {{ 0%, 100% {{ opacity: 1; }} 92% {{ opacity: 1; }} 93% {{ opacity: .55; }} 94% {{ opacity: 1; }} 96% {{ opacity: .75; }} 97% {{ opacity: 1; }} }}

/* ---------- Base ---------- */
.stApp {{ background: {T['fondo']}, {T['bg']} !important; background-attachment: fixed !important;
  color: {T['text']} !important; font-family: {FUENTE}; transition: background-color .4s ease, color .4s ease; }}
[data-testid="stHeader"] {{ background: transparent !important; }}
#MainMenu, footer, [data-testid="stDecoration"] {{ visibility: hidden; display: none; }}
.block-container {{ padding-top: 1.6rem; max-width: 1380px; }}

p, span, label, div, li, td, th, .stMarkdown {{ color: {T['text']}; font-family: {FUENTE}; }}
h1, h2, h3, h4, h5, h6 {{ color: {T['text']} !important; font-family: {FUENTE} !important; letter-spacing: -0.01em; }}
strong, b {{ color: {T['text']}; }}
a {{ color: {T['accent']} !important; }}
[data-testid="stCaptionContainer"] p, .stCaption, small {{ color: {T['sub']} !important; }}
::selection {{ background: {rgba(T['accent'], 0.30)}; }}
code {{ color: {T['accent']} !important; background: {rgba(T['accent'], 0.08)} !important; }}

/* Scrollbar */
::-webkit-scrollbar {{ width: 10px; height: 10px; }}
::-webkit-scrollbar-thumb {{ background: {rgba(T['accent'], 0.35)}; border-radius: 10px; }}
::-webkit-scrollbar-track {{ background: transparent; }}

/* Sidebar */
[data-testid="stSidebar"] {{ background-color: {T['card']} !important; border-right: 1px solid {T['border']}; }}
[data-testid="stSidebar"] * {{ color: {T['text']}; }}

/* ---------- Tabs ---------- */
.stTabs [data-baseweb="tab-list"] {{ gap: 4px; border-bottom: 1px solid {T['border']}; flex-wrap: wrap; background: transparent; }}
.stTabs [data-baseweb="tab-list"] button {{ background: transparent !important; padding: 10px 14px; border-radius: 10px 10px 0 0;
  transition: background-color .25s ease, transform .2s ease; }}
.stTabs [data-baseweb="tab-list"] button:hover {{ background: {rgba(T['accent'], 0.07)} !important; }}
.stTabs [data-baseweb="tab-list"] button p {{ color: {T['sub']} !important; font-size: 0.95rem; font-weight: 600; transition: color .25s; }}
.stTabs [data-baseweb="tab-list"] button:hover p {{ color: {T['text']} !important; }}
.stTabs [data-baseweb="tab-list"] button[aria-selected="true"] p {{ color: {T['accent']} !important; text-shadow: {GLOW_TXT}; }}
.stTabs [data-baseweb="tab-highlight"] {{ background-color: {T['accent']} !important; height: 3px; border-radius: 3px;
  box-shadow: {GLOW if T['neon'] else 'none'}; }}
.stTabs [data-baseweb="tab-border"] {{ display: none; }}
.stTabs [data-baseweb="tab-panel"] {{ animation: mdFade .35s ease both; padding-top: 10px; }}
[data-testid="stTab"] {{ border-radius: 10px 10px 0 0; transition: background-color .25s ease; padding-left: 12px !important; padding-right: 12px !important; }}
[data-testid="stTab"]:hover {{ background: {rgba(T['accent'], 0.07)}; }}
[data-testid="stTab"] p {{ color: {T['sub']} !important; font-weight: 600; transition: color .25s; }}
[data-testid="stTab"]:hover p {{ color: {T['text']} !important; }}
[data-testid="stTab"][aria-selected="true"] p {{ color: {T['accent']} !important; text-shadow: {GLOW_TXT}; }}
[data-testid="stTabPanel"] {{ animation: mdFade .35s ease both; padding-top: 10px; }}
input, textarea, select, button {{ font-family: {FUENTE} !important; }}

/* ---------- Inputs ---------- */
div[data-baseweb="input"] > div, div[data-baseweb="select"] > div, div[data-baseweb="textarea"] > div,
div[data-baseweb="base-input"], div[data-baseweb="input"], div[data-baseweb="textarea"] {{
  background-color: {T['input']} !important; border-color: {T['border']} !important; border-radius: 10px !important;
  transition: border-color .2s ease, box-shadow .2s ease; }}
div[data-baseweb="input"]:focus-within, div[data-baseweb="select"] > div:focus-within, div[data-baseweb="textarea"]:focus-within {{
  border-color: {T['accent']} !important; box-shadow: 0 0 0 3px {rgba(T['accent'], 0.18)} !important; }}
input, textarea {{ color: {T['text']} !important; background-color: {T['input']} !important; caret-color: {T['accent']}; }}
input::placeholder, textarea::placeholder {{ color: {T['sub']} !important; opacity: .8; }}
[data-baseweb="popover"] ul, [data-baseweb="popover"] li, [data-baseweb="menu"], [data-baseweb="popover"] > div {{
  background-color: {T['input']} !important; color: {T['text']} !important; }}
[data-baseweb="popover"] li:hover {{ background-color: {rgba(T['accent'], 0.12)} !important; }}
[data-baseweb="calendar"], [data-baseweb="calendar"] * {{ background-color: {T['input']} !important; color: {T['text']} !important; }}
[data-testid="stNumberInput"] button {{ background-color: {T['input']} !important; border-color: {T['border']} !important; }}
[data-testid="stNumberInput"] button:hover {{ background-color: {rgba(T['accent'], 0.15)} !important; }}
[data-testid="stNumberInput"] button svg {{ fill: {T['text']} !important; }}
[data-baseweb="select"] svg {{ fill: {T['sub']} !important; }}

/* Sliders, toggles, radios, checkboxes → color de acento */
[data-testid="stSlider"] [role="slider"] {{ background-color: {T['accent']} !important; border-color: {T['accent']} !important;
  box-shadow: 0 0 0 4px {rgba(T['accent'], 0.18)} !important; }}
[data-testid="stSlider"] [data-testid="stThumbValue"], [data-testid="stSliderThumbValue"] {{ color: {T['accent']} !important; font-weight: 700; }}
[data-testid="stSlider"] [data-baseweb="slider"] > div > div > div:first-child {{ background: {rgba(T['accent'], 0.85)} !important; }}
[data-testid="stTickBarMin"], [data-testid="stTickBarMax"] {{ color: {T['sub']} !important; }}
label[data-baseweb="checkbox"] input:checked + div, label[data-baseweb="checkbox"] [aria-checked="true"] {{ background-color: {T['accent']} !important; }}
label[data-baseweb="radio"] input:checked + div {{ background-color: {T['accent']} !important; border-color: {T['accent']} !important; }}
label[data-baseweb="radio"] input:checked + div > div {{ background-color: {T['on_accent']} !important; }}
[data-testid="stProgress"] [role="progressbar"] > div > div > div {{ background: linear-gradient(90deg, {T['accent']}, {T['accent2']}) !important; }}

/* Alertas */
[data-testid="stAlert"] {{ background: {rgba(T['accent'], 0.07)} !important; border: 1px solid {T['border']} !important; border-radius: 12px !important; }}
[data-testid="stAlert"] * {{ color: {T['text']} !important; }}

/* ---------- Botones ---------- */
.stButton>button, .stFormSubmitButton>button, .stDownloadButton>button {{
  background-color: {T['card']} !important; color: {T['text']} !important; border: 1px solid {T['border']} !important;
  border-radius: 12px; font-weight: 600; box-shadow: {SOMBRA};
  transition: transform .18s ease, border-color .2s ease, background-color .2s ease, box-shadow .25s ease; }}
.stButton>button:hover, .stFormSubmitButton>button:hover, .stDownloadButton>button:hover {{
  border-color: {T['accent']} !important; background-color: {rgba(T['accent'], 0.08)} !important;
  transform: translateY(-2px); box-shadow: {GLOW}; }}
.stButton>button:active, .stDownloadButton>button:active {{ transform: translateY(0) scale(.98); }}
.stButton>button p, .stDownloadButton>button p, .stFormSubmitButton>button p {{ color: inherit !important; }}
.stButton>button:focus-visible, .stDownloadButton>button:focus-visible {{ outline: 2px solid {T['accent']}; outline-offset: 2px; }}
.stDownloadButton>button {{ border-style: solid !important; }}

/* Selector de tema (botones con icono) */
[class*="st-key-tema_btn_"] button {{ border-radius: 999px !important; min-height: 42px; font-size: 1.15rem; padding: 0 !important; }}
.st-key-tema_btn_{st.session_state.tema} button {{ background: {T['accent']} !important; border-color: {T['accent']} !important;
  box-shadow: {GLOW} !important; animation: mdPulse 2.4s ease-in-out 1; }}

/* ---------- Contenedores ---------- */
[data-testid="stExpander"] details {{ background: {T['card']}; border: 1px solid {T['border']} !important; border-radius: 14px;
  transition: border-color .2s; }}
[data-testid="stExpander"] details:hover {{ border-color: {rgba(T['accent'], 0.6)} !important; }}
[data-testid="stExpander"] summary:hover * {{ color: {T['accent']} !important; }}
[data-testid="stExpander"] svg {{ fill: {T['sub']} !important; }}
[data-testid="stForm"] {{ background: {T['card']}; border: 1px solid {T['border']} !important; border-radius: 16px; box-shadow: {SOMBRA}; }}
[data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"]) {{ border-radius: 16px !important; }}
[data-testid="stVerticalBlockBorderWrapper"] {{ border-color: {T['border']} !important; }}
[data-testid="stLayoutWrapper"] > div[data-testid="stVerticalBlock"], .stVerticalBlock[class*="border"] {{ border-color: {T['border']} !important; }}
hr {{ border-color: {T['border']} !important; opacity: .8; }}
iframe {{ border-radius: 14px; }}

/* ---------- Componentes propios ---------- */
.md-hero {{ position: relative; padding: 6px 0 16px 0; animation: mdFadeUp .5s ease both; }}
.md-hero::after {{ content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 3px; border-radius: 3px;
  background: linear-gradient(90deg, {T['accent']}, {T['accent2']}, {rgba(T['accent'], 0.1)}, {T['accent']});
  background-size: 200% 100%; animation: mdSlide 6s linear infinite; box-shadow: {GLOW if T['neon'] else 'none'}; }}
.md-hero h1 {{ font-size: 2.4rem; font-weight: 800; margin: 0; padding: 0; line-height: 1.1; text-shadow: {GLOW_TXT};
  {'animation: mdFlicker 7s linear infinite;' if T['neon'] else ''} }}
.md-hero h1 .cur {{ color: {T['accent']}; animation: mdBlink 1.1s step-end infinite; font-weight: 400; }}
.md-hero .md-hero-sub {{ color: {T['sub']}; font-size: 1rem; margin-top: 6px; }}
.md-hero .md-chip {{ display: inline-block; margin-top: 10px; border: 1px solid {T['border']}; background: {T['card']};
  border-radius: 999px; padding: 5px 14px; font-size: .88rem; color: {T['text']}; }}
.md-hero .md-chip b {{ color: {T['accent']}; }}
.md-hero .md-chip .live {{ display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: {T['ok']};
  margin-right: 7px; vertical-align: middle; animation: mdPulse 2s infinite; }}

.md-sec {{ margin: 6px 0 14px 0; animation: mdFadeUp .45s ease both; }}
.md-sec .md-sec-t {{ font-size: 1.3rem; font-weight: 700; color: {T['text']}; display: flex; align-items: center; gap: 10px; }}
.md-sec .md-sec-t::before {{ content: ""; width: 4px; height: 1.1em; border-radius: 4px; background: {T['accent']};
  box-shadow: {GLOW if T['neon'] else 'none'}; }}
.md-sec .md-sec-s {{ color: {T['sub']}; font-size: .95rem; margin-top: 3px; max-width: 85ch; padding-left: 14px; }}

.md-kpi {{ background: {T['card']}; border: 1px solid {T['border']}; border-radius: 16px; padding: 16px 18px;
  position: relative; overflow: hidden; height: 100%; box-shadow: {SOMBRA};
  animation: mdFadeUp .5s ease both; animation-delay: var(--d, 0s);
  transition: transform .22s ease, box-shadow .25s ease, border-color .25s ease; }}
.md-kpi::before {{ content: ""; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: var(--c); }}
.md-kpi:hover {{ transform: translateY(-4px); border-color: var(--c); box-shadow: {GLOW}; }}
.md-kpi .k-l {{ color: {T['sub']}; font-size: .84rem; font-weight: 600; }}
.md-kpi .k-v {{ color: {T['text']}; font-size: 1.6rem; font-weight: 800; margin-top: 6px; font-variant-numeric: tabular-nums; }}
.md-kpi .k-s {{ color: {T['sub']}; font-size: .82rem; margin-top: 3px; }}

.md-verdict {{ display: flex; gap: 28px; align-items: center; justify-content: space-between; flex-wrap: wrap;
  background: linear-gradient(135deg, var(--cb), {T['card']} 55%); border: 1px solid var(--c); border-radius: 18px;
  padding: 22px 26px; animation: mdFadeUp .55s ease both; box-shadow: 0 0 0 1px {rgba('#000000', 0)}, {GLOW}; }}
.md-verdict .v-tag {{ display: inline-flex; align-items: center; gap: 8px; font-size: .85rem; font-weight: 700; color: var(--c);
  background: var(--cb); padding: 5px 13px; border-radius: 999px; }}
.md-verdict .v-tag::before {{ content: ""; width: 8px; height: 8px; border-radius: 50%; background: var(--c); }}
.md-verdict .v-big {{ font-size: 1.45rem; font-weight: 700; margin: 12px 0 8px 0; color: {T['text']}; line-height: 1.3; max-width: 46ch; }}
.md-verdict .v-big b {{ color: var(--c); }}
.md-verdict .v-txt {{ color: {T['sub']}; font-size: .96rem; max-width: 70ch; line-height: 1.55; }}
.md-verdict .v-txt b {{ color: {T['text']}; }}

.dotwrap {{ display: flex; flex-direction: column; align-items: center; gap: 10px; }}
.dotgrid {{ display: grid; grid-template-columns: repeat(10, 13px); gap: 5px; }}
.dot {{ width: 13px; height: 13px; border-radius: 50%; animation: mdPop .35s ease both; }}
.dot.ok {{ background: {T['ok']}; {'box-shadow: 0 0 6px ' + rgba(T['ok'], 0.7) + ';' if T['neon'] else ''} }}
.dot.mid {{ background: {T['warn']}; }}
.dot.bad {{ background: {rgba(T['bad'], 0.85)}; {'box-shadow: 0 0 6px ' + rgba(T['bad'], 0.7) + ';' if T['neon'] else ''} }}
.dotleg {{ display: flex; gap: 12px; font-size: .78rem; color: {T['sub']}; flex-wrap: wrap; justify-content: center; }}
.dotleg span {{ color: {T['sub']}; }}
.dotleg i {{ display: inline-block; width: 9px; height: 9px; border-radius: 50%; margin-right: 4px; vertical-align: middle; }}

.md-box {{ background: {T['card']}; border: 1px solid {T['border']}; border-radius: 16px; padding: 16px 18px; box-shadow: {SOMBRA};
  animation: mdFadeUp .5s ease both; transition: border-color .25s ease; }}
.md-box:hover {{ border-color: {rgba(T['accent'], 0.55)}; }}
.md-box .b-t {{ font-weight: 700; font-size: 1rem; margin-bottom: 8px; color: {T['text']}; }}
.md-row {{ display: flex; justify-content: space-between; gap: 12px; padding: 8px 0; border-bottom: 1px dashed {T['border']}; font-size: .93rem; }}
.md-row:last-child {{ border-bottom: none; }}
.md-row .r-l {{ color: {T['sub']}; }}
.md-row .r-v {{ color: {T['text']}; font-weight: 600; font-variant-numeric: tabular-nums; text-align: right; }}
.pill {{ display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: .8rem; font-weight: 700; }}
.pill.ok {{ color: {T['ok']}; background: {rgba(T['ok'], 0.14)}; }}
.pill.bad {{ color: {T['bad']}; background: {rgba(T['bad'], 0.14)}; }}
.pill.warn {{ color: {T['warn']}; background: {rgba(T['warn'], 0.16)}; }}

.md-answer {{ background: {rgba(T['accent'], 0.07)}; border: 1px solid {rgba(T['accent'], 0.3)}; border-radius: 14px;
  padding: 12px 16px; font-size: 1rem; margin-top: 6px; animation: mdFade .3s ease both; }}
.md-answer b {{ color: {T['accent']}; font-size: 1.15rem; text-shadow: {GLOW_TXT}; }}

.kb-col {{ background: {T['card']}; border: 1px solid {T['border']}; border-radius: 16px; padding: 12px; min-height: 140px;
  box-shadow: {SOMBRA}; animation: mdFadeUp .5s ease both; animation-delay: var(--d, 0s); }}
.kb-h {{ display: flex; justify-content: space-between; align-items: center; font-weight: 700; margin-bottom: 10px; }}
.kb-h span.n {{ color: {T['sub']}; font-weight: 600; background: {rgba(T['accent'], 0.1)}; border-radius: 999px; padding: 1px 9px; font-size: .82rem; }}
.kb-card {{ background: {T['bg']}; border: 1px solid {T['border']}; border-left: 3px solid var(--c); border-radius: 10px;
  padding: 9px 11px; margin-bottom: 8px; font-size: .92rem; transition: transform .18s ease, box-shadow .2s ease; }}
.kb-card:hover {{ transform: translateX(3px); box-shadow: {GLOW}; }}

.md-step {{ display: flex; align-items: center; gap: 10px; font-weight: 700; font-size: 1.02rem; margin: 2px 0 10px 0; }}
.md-step .n {{ width: 28px; height: 28px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center;
  background: {T['accent']}; color: {T['on_accent']} !important; font-size: .85rem; box-shadow: {GLOW if T['neon'] else 'none'}; }}

/* Tablas HTML propias */
.md-tablewrap {{ overflow-x: auto; border: 1px solid {T['border']}; border-radius: 14px; background: {T['card']};
  animation: mdFadeUp .5s ease both; box-shadow: {SOMBRA}; }}
.md-table {{ width: 100%; border-collapse: collapse; font-size: .92rem; }}
.md-table th {{ text-align: left; padding: 11px 14px; color: {T['sub']}; font-weight: 700; font-size: .82rem;
  background: {rgba(T['accent'], 0.07)}; border-bottom: 1px solid {T['border']}; white-space: nowrap; }}
.md-table th, .md-table td {{ border-left: none !important; border-right: none !important; border-top: none !important; }}
.md-table td {{ padding: 10px 14px; border-bottom: 1px solid {T['border']}; font-variant-numeric: tabular-nums; }}
.md-table tr:last-child td {{ border-bottom: none; }}
.md-table tbody tr {{ transition: background-color .15s ease; }}
.md-table tbody tr:hover {{ background: {rgba(T['accent'], 0.06)}; }}
.md-table tr.hl td {{ background: {rgba(T['accent'], 0.14)}; font-weight: 700; }}
.md-table tr.hl td:first-child {{ box-shadow: inset 3px 0 0 {T['accent']}; }}

.md-foot {{ text-align: center; color: {T['sub']}; font-size: .82rem; padding: 28px 0 8px 0; }}

/* Accesibilidad: sin animaciones si el sistema lo pide */
@media (prefers-reduced-motion: reduce) {{
  *, *::before, *::after {{ animation: none !important; transition: none !important; }}
}}
</style>
""", unsafe_allow_html=True)


# =====================================================================
# 2. UTILIDADES DE FORMATO Y UI
# =====================================================================
def fmt_num(x, dec=0):
    """Formato chileno: punto de miles y coma decimal."""
    try:
        s = f"{float(x):,.{dec}f}"
    except (TypeError, ValueError):
        return str(x)
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_money(x):
    return "$" + fmt_num(x, 0)


def slug(texto):
    return "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in str(texto).strip()) or "Proyecto"


def seccion(titulo, subtitulo=""):
    sub = f'<div class="md-sec-s">{subtitulo}</div>' if subtitulo else ""
    st.markdown(f'<div class="md-sec"><div class="md-sec-t">{titulo}</div>{sub}</div>', unsafe_allow_html=True)


def paso(numero, titulo):
    st.markdown(f'<div class="md-step"><span class="n">{numero}</span>{titulo}</div>', unsafe_allow_html=True)


def kpi(col, etiqueta, valor, sub="", color=None, delay=0.0):
    color = color or T["accent"]
    col.markdown(
        f'<div class="md-kpi" style="--c:{color};--d:{delay:.2f}s"><div class="k-l">{etiqueta}</div>'
        f'<div class="k-v">{valor}</div><div class="k-s">{sub}</div></div>',
        unsafe_allow_html=True,
    )


def caja_filas(titulo, filas):
    """filas: lista de (etiqueta, valor_html)"""
    html = "".join(f'<div class="md-row"><span class="r-l">{l}</span><span class="r-v">{v}</span></div>' for l, v in filas)
    return f'<div class="md-box"><div class="b-t">{titulo}</div>{html}</div>'


def tabla_html(df, resaltar=None):
    """Tabla HTML con el estilo del tema (st.dataframe no respeta los temas oscuros personalizados)."""
    cab = "".join(f"<th>{c}</th>" for c in df.columns)
    filas = []
    for i, fila in enumerate(df.itertuples(index=False)):
        clase = ' class="hl"' if resaltar is not None and resaltar(fila) else ""
        celdas = "".join(f"<td>{v}</td>" for v in fila)
        filas.append(f"<tr{clase}>{celdas}</tr>")
    return f'<div class="md-tablewrap"><table class="md-table"><thead><tr>{cab}</tr></thead><tbody>{"".join(filas)}</tbody></table></div>'


def pill(texto, tipo):
    return f'<span class="pill {tipo}">{texto}</span>'


def limpiar_widgets(*prefijos):
    """Evita que los widgets conserven valores de filas eliminadas."""
    for k in list(st.session_state.keys()):
        if any(str(k).startswith(p) for p in prefijos):
            del st.session_state[k]


def mostrar_mermaid(codigo, alto="420px", key=None):
    """Renderiza Mermaid con los colores del tema activo (si el código no trae su propio %%{init}%%)."""
    codigo = str(codigo or "").strip()
    if not codigo:
        st.info("Escribe código Mermaid para ver la vista previa.")
        return
    if not codigo.startswith("%%{init"):
        variables = {
            "background": T["card"], "mainBkg": T["card"], "primaryColor": T["card"],
            "primaryTextColor": T["text"], "primaryBorderColor": T["accent"], "lineColor": T["accent"],
            "secondaryColor": T["bg"], "tertiaryColor": T["bg"], "textColor": T["text"],
            "nodeBorder": T["accent"], "clusterBkg": T["bg"], "edgeLabelBackground": T["card"],
            "actorBkg": T["card"], "actorBorder": T["accent"], "actorTextColor": T["text"],
            "actorLineColor": T["sub"], "signalColor": T["text"], "signalTextColor": T["text"],
            "labelBoxBkgColor": T["card"], "labelTextColor": T["text"], "noteBkgColor": T["bg"],
            "noteTextColor": T["text"], "sequenceNumberColor": T["on_accent"],
            "fontFamily": "JetBrains Mono, Consolas, monospace" if T["neon"] else "Segoe UI, Helvetica, Arial, sans-serif",
        }
        init = json.dumps({"theme": "base", "themeVariables": variables})
        codigo = f"%%{{init: {init}}}%%\n" + codigo
    with st.container(border=True):
        st_mermaid(codigo, height=alto, key=key)


def estilo_fig(fig, height=380, **kwargs):
    base = dict(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FUENTE, color=T["text"], size=13),
        height=height, margin=dict(l=10, r=10, t=40, b=10),
        hoverlabel=dict(bgcolor=T["card"], font_color=T["text"], bordercolor=T["border"]),
        legend=dict(orientation="h", y=1.1, x=0, bgcolor="rgba(0,0,0,0)"),
        separators=",.",
    )
    base.update(kwargs)
    fig.update_layout(**base)
    fig.update_xaxes(gridcolor=T["grid"], zeroline=False, linecolor=T["border"])
    fig.update_yaxes(gridcolor=T["grid"], zeroline=False, linecolor=T["border"])
    return fig


PLOTLY_CFG = {"displaylogo": False, "modeBarButtonsToRemove": ["lasso2d", "select2d"]}


def mostrar_fig(fig):
    st.plotly_chart(fig, config=PLOTLY_CFG, **ANCHO)


# =====================================================================
# 3. PLANTILLAS Y ESTADO
# =====================================================================
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

DEFAULTS = {
    "nombre_proj": "MatrixDev Core",
    "integrantes": "Juan Pérez",
    "objetivo_text": "Automatizar el flujo de inventario con una arquitectura minimalista.",
    "rf_list": ["Autenticación OAuth2.", "CRUD de usuarios."],
    "rnf_list": ["Latencia < 200ms.", "Cifrado AES-256 en base de datos."],
    "kanban_tasks": [
        {"Tarea": "Modelo BD", "Estado": "Completado"},
        {"Tarea": "Endpoints API", "Estado": "En Proceso"},
    ],
    "edt_list": [
        {"Fase": "Fase 1: Preparación", "Subfase": "Limpieza y Despeje", "Paquete": "Despejar Habitación y Muebles",
         "Horas": 5, "Costo": 20000, "Documentos": "Plan de Trabajo"},
        {"Fase": "Fase 1: Preparación", "Subfase": "Pintura Aerosol", "Paquete": "Pintar Arbusto / Detalles",
         "Horas": 4, "Costo": 35000, "Documentos": "Ficha Técnica Aerosol"},
        {"Fase": "Fase 2: Aplicación Manual", "Subfase": "Capas de Pintura", "Paquete": "Mano de Pintura Muros (1 y 2)",
         "Horas": 16, "Costo": 80000, "Documentos": "Guía de Pintado Manual"},
        {"Fase": "Fase 2: Aplicación Manual", "Subfase": "Documentación y Entrega",
         "Paquete": "Informe de Inspección y Cotizaciones", "Horas": 6, "Costo": 15000,
         "Documentos": "Cotización + Certificados"},
    ],
    "riesgos_list": [
        {"Riesgo": "Retraso en entrega API", "Probabilidad": 3, "Impacto": 4},
        {"Riesgo": "Falla en servidor cloud", "Probabilidad": 2, "Impacto": 5},
        {"Riesgo": "Incompatibilidad de navegador", "Probabilidad": 1, "Impacto": 2},
    ],
    "flujo_codigo": PLANTILLAS_FLUJO["Autenticación"],
    "modelo_mermaid": """graph TD
    A[Inicio] --> B{Validar}
    B -->|Sí| C[Éxito]
    B -->|No| D[Error]""",
    "mc_semilla_val": 42,
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

ESTADOS_KANBAN = ["Pendiente", "En Proceso", "Completado"]

# Resultado vigente de Monte Carlo (lo rellena la pestaña 4 y lo leen los reportes)
MC_ACTUAL = None


def cambiar_plantilla_flujo():
    sel = st.session_state.select_tipo_flujo
    if sel in PLANTILLAS_FLUJO:
        st.session_state.flujo_codigo = PLANTILLAS_FLUJO[sel]
        st.session_state.txt_flujo_code = PLANTILLAS_FLUJO[sel]


def generar_mermaid_edt(edt_items, nombre_proyecto):
    def limpio(x):
        return str(x).strip().replace('"', '').replace("'", "").replace("$", "")

    lines = ["graph TD", f'    ROOT["📦 {limpio(nombre_proyecto)}"]']
    fases = {}
    for item in edt_items:
        fase = limpio(item.get("Fase", "Sin Fase")) or "Sin Fase"
        subfase = limpio(item.get("Subfase", ""))
        fases.setdefault(fase, {}).setdefault(subfase, []).append({
            "paquete": limpio(item.get("Paquete", "Tarea")),
            "horas": float(item.get("Horas", 0) or 0),
            "costo": float(item.get("Costo", 0) or 0),
            "doc": limpio(item.get("Documentos", "")),
        })

    for fase_idx, (fase_name, subfases) in enumerate(fases.items()):
        fase_id = f"F{fase_idx}"
        t_h = sum(p["horas"] for s in subfases.values() for p in s)
        t_c = sum(p["costo"] for s in subfases.values() for p in s)
        lines.append(f'    ROOT --> {fase_id}["📂 {fase_name}<br/><i>Total: {t_h:.0f}h | ${fmt_num(t_c)}</i>"]')
        for sub_idx, (sub_name, pkgs) in enumerate(subfases.items()):
            if sub_name:
                sub_id = f"{fase_id}_S{sub_idx}"
                s_h = sum(p["horas"] for p in pkgs)
                s_c = sum(p["costo"] for p in pkgs)
                lines.append(f'    {fase_id} --> {sub_id}["📁 {sub_name}<br/><i>{s_h:.0f}h | ${fmt_num(s_c)}</i>"]')
                parent = sub_id
            else:
                parent = fase_id
            for pkg_idx, pkg in enumerate(pkgs):
                det = [f"⏱️ {pkg['horas']:.0f} hrs"]
                if pkg["costo"] > 0:
                    det.append(f"💰 ${fmt_num(pkg['costo'])}")
                if pkg["doc"]:
                    det.append(f"📄 {pkg['doc']}")
                lines.append(f'    {parent} --> {parent}_P{pkg_idx}["<b>{pkg["paquete"]}</b><br/>{"<br/>".join(det)}"]')
    return "\n".join(lines)


def clasificar_riesgo(sev):
    if sev <= 5:
        return "Bajo 🟢"
    elif sev <= 10:
        return "Medio 🟡"
    elif sev <= 15:
        return "Alto 🟠"
    return "Crítico 🔴"


# =====================================================================
# 4. MOTOR MONTE CARLO
# =====================================================================
DISTRIBUCIONES = ["Beta-PERT", "Triangular", "Normal", "Uniforme"]
NIVELES_PERCENTIL = [10, 25, 50, 75, 80, 85, 90, 95, 99]


def _muestra(rng, dist, a, m, b, n):
    """Genera n muestras de una distribución de tres puntos (a=mín, m=más probable, b=máx)."""
    if b <= a:
        return np.full(n, float(m))
    if dist == "Triangular":
        return rng.triangular(a, min(max(m, a), b), b, n)
    if dist == "Uniforme":
        return rng.uniform(a, b, n)
    if dist == "Normal":
        media = (a + 4 * m + b) / 6
        desv = (b - a) / 6
        return np.maximum(rng.normal(media, desv, n), a * 0.5)
    # Beta-PERT (PMBOK)
    alfa = 1 + 4 * (m - a) / (b - a)
    beta = 1 + 4 * (b - m) / (b - a)
    return a + rng.beta(alfa, beta, n) * (b - a)


@st.cache_data(show_spinner=False, max_entries=24)
def simular_manual(dist, p1, p2, p3, n, costo_fijo, costo_diario, var_costo_pct, seed):
    rng = np.random.default_rng(seed)
    if dist == "Normal":  # p1 = media, p2 = desviación
        dias = np.maximum(rng.normal(p1, p2, n), 0.5)
    elif dist == "Uniforme":  # p1 = mín, p3 = máx
        dias = rng.uniform(p1, p3, n)
    else:
        dias = _muestra(rng, dist, p1, p2, p3, n)
    factor = np.maximum(rng.normal(1.0, var_costo_pct / 100.0, n), 0.0)
    costos = costo_fijo + dias * costo_diario * factor
    return {"dias": dias, "costos": costos, "componentes": None, "nombres": None}


@st.cache_data(show_spinner=False, max_entries=24)
def simular_edt(paquetes, horas_dia, pct_opt, pct_pes, dist, n, costo_fijo, seed):
    """Cada paquete de la EDT se simula por separado y se suman (ejecución secuencial)."""
    rng = np.random.default_rng(seed)
    nombres = [p[0] for p in paquetes]
    comp = np.zeros((n, len(paquetes)))
    costos = np.full(n, float(costo_fijo))
    for j, (_, horas, costo) in enumerate(paquetes):
        m = max(float(horas), 0.0) / horas_dia
        if m <= 0:
            continue
        a, b = m * (1 - pct_opt / 100.0), m * (1 + pct_pes / 100.0)
        s = _muestra(rng, dist, a, m, b, n)
        comp[:, j] = s
        costos += float(costo) * (s / m)  # el costo escala con el esfuerzo real
    return {"dias": comp.sum(axis=1), "costos": costos, "componentes": comp, "nombres": nombres}


def calcular_metricas(dias, costos, dias_meta, costo_meta, perc):
    ok_t = dias <= dias_meta
    ok_c = costos <= costo_meta
    m = {
        "n": int(len(dias)),
        "media": float(np.mean(dias)), "desv": float(np.std(dias)),
        "min": float(np.min(dias)), "max": float(np.max(dias)),
        "p50": float(np.percentile(dias, 50)), "pt": float(np.percentile(dias, perc)),
        "p95": float(np.percentile(dias, 95)),
        "p50c": float(np.percentile(costos, 50)), "ptc": float(np.percentile(costos, perc)),
        "p95c": float(np.percentile(costos, 95)), "media_c": float(np.mean(costos)),
        "prob_t": float(ok_t.mean() * 100), "prob_c": float(ok_c.mean() * 100),
        "prob_ambos": float((ok_t & ok_c).mean() * 100),
        "prob_uno": float((ok_t ^ ok_c).mean() * 100),
    }
    m["prob_ninguno"] = 100 - m["prob_ambos"] - m["prob_uno"]
    m["riesgo_t"] = 100 - m["prob_t"]
    m["reserva_dias"] = m["pt"] - m["p50"]
    m["reserva_costo"] = m["ptc"] - m["p50c"]
    m["cv"] = m["desv"] / m["media"] * 100 if m["media"] else 0
    return m


def fecha_fin(inicio, dias, habiles):
    d = int(np.ceil(max(float(dias), 0)))
    if habiles:
        return pd.Timestamp(np.busday_offset(np.datetime64(inicio, "D"), d, roll="forward"))
    return pd.Timestamp(inicio) + pd.Timedelta(days=d)


def tabla_percentiles(dias, costos, inicio, habiles, perc):
    niveles = sorted(set(NIVELES_PERCENTIL + [int(perc)]))
    filas = []
    for p in niveles:
        d = float(np.percentile(dias, p))
        c = float(np.percentile(costos, p))
        filas.append({"Confianza": f"P{p}", "p": p, "Días": d, "Fecha": fecha_fin(inicio, d, habiles), "Costo": c})
    return filas


def veredicto(met, perc):
    pa = met["prob_ambos"]
    if pa >= perc:
        return "ok", "Viable con la confianza exigida"
    if pa >= 50:
        return "warn", "Viable, con riesgo moderado"
    return "bad", "Alto riesgo de incumplimiento"


# =====================================================================
# 5. GENERACIÓN DE DOCUMENTOS (EXCEL / WORD)
# =====================================================================
XL_NAVY, XL_ACC = "1E3A8A", "2563EB"
XL_ZEBRA, XL_SOFT = "F1F5F9", "F8FAFC"
XL_OK, XL_BAD, XL_HL = "DCFCE7", "FEE2E2", "FEF3C7"
XL_BORDE = Border(*(Side(style="thin", color="E2E8F0") for _ in range(4)))


def _xl_titulo(ws, texto, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(ncols, 2))
    c = ws.cell(row=1, column=1, value=texto)
    c.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", start_color=XL_NAVY, end_color=XL_NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 32


def _xl_meta(ws, fila, pares):
    for etiqueta, valor in pares:
        a = ws.cell(row=fila, column=1, value=etiqueta)
        b = ws.cell(row=fila, column=2, value=valor)
        a.font = Font(name="Calibri", bold=True, color="0F172A")
        a.fill = PatternFill("solid", start_color=XL_SOFT, end_color=XL_SOFT)
        b.font = Font(name="Calibri", color="334155")
        a.border = b.border = XL_BORDE
        fila += 1
    return fila


def _xl_tabla(ws, fila, encabezados, filas, formatos=None, resaltar=None):
    """Escribe una tabla con estilo. formatos: {col_idx(1-based): number_format}. Devuelve la siguiente fila libre."""
    formatos = formatos or {}
    for j, h in enumerate(encabezados, 1):
        c = ws.cell(row=fila, column=j, value=h)
        c.font = Font(name="Calibri", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=XL_ACC, end_color=XL_ACC)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = XL_BORDE
    ws.row_dimensions[fila].height = 24
    for i, f in enumerate(filas):
        r = fila + 1 + i
        for j, v in enumerate(f, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = Font(name="Calibri", color="334155")
            c.border = XL_BORDE
            if j in formatos:
                c.number_format = formatos[j]
            if isinstance(v, (int, float)):
                c.alignment = Alignment(horizontal="right")
            if resaltar is not None and resaltar(i, f):
                c.fill = PatternFill("solid", start_color=XL_HL, end_color=XL_HL)
                c.font = Font(name="Calibri", bold=True, color="0F172A")
            elif i % 2 == 1:
                c.fill = PatternFill("solid", start_color=XL_ZEBRA, end_color=XL_ZEBRA)
    return fila + len(filas) + 1


def _xl_autofit(ws, minimo=12, maximo=70):
    anchos = {}
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if c.value is None:
                continue
            v = c.value
            if isinstance(v, float) and "%" in (c.number_format or ""):
                txt = f"{v * 100:.1f}%"
            elif isinstance(v, (int, float)) and "$" in (c.number_format or ""):
                txt = f"${v:,.0f}"
            elif isinstance(v, float):
                txt = f"{v:,.2f}"
            else:
                txt = str(v)
            anchos[c.column_letter] = max(anchos.get(c.column_letter, 0), len(txt))
    for col, w in anchos.items():
        ws.column_dimensions[col].width = min(max(w + 4, minimo), maximo)


def _xl_bytes(wb):
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _filas_requisitos():
    filas = [[f"RF-{i:02d}", "Funcional", t.strip()] for i, t in enumerate(st.session_state.rf_list, 1) if t.strip()]
    filas += [[f"RNF-{i:02d}", "No funcional", t.strip()] for i, t in enumerate(st.session_state.rnf_list, 1) if t.strip()]
    return filas


def _filas_riesgos():
    filas = []
    for r in st.session_state.riesgos_list:
        p, i = int(r.get("Probabilidad", 1)), int(r.get("Impacto", 1))
        filas.append([r.get("Riesgo", ""), p, i, p * i, clasificar_riesgo(p * i)])
    return filas


def _filas_edt():
    return [[e.get("Fase", ""), e.get("Subfase", ""), e.get("Paquete", ""), float(e.get("Horas", 0) or 0),
             float(e.get("Costo", 0) or 0), e.get("Documentos", "")] for e in st.session_state.edt_list]


ENC_EDT = ["Fase principal", "Subfase / grupo", "Paquete de trabajo", "Horas estimadas", "Costo estimado ($)", "Documentos"]


def generar_excel_requisitos():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Requisitos"
    _xl_titulo(ws, f"Matriz de requisitos · {st.session_state.nombre_proj}", 3)
    f = _xl_meta(ws, 3, [("Proyecto", st.session_state.nombre_proj), ("Objetivo", st.session_state.objetivo_text)])
    _xl_tabla(ws, f + 1, ["Código", "Tipo de requisito", "Descripción"], _filas_requisitos())
    ws.freeze_panes = ws.cell(row=f + 2, column=1)
    _xl_autofit(ws)
    return _xl_bytes(wb)


def generar_excel_kanban():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tablero Kanban"
    _xl_titulo(ws, f"Tablero Kanban · {st.session_state.nombre_proj}", 3)
    filas = [[i, t["Tarea"], t["Estado"]] for i, t in enumerate(st.session_state.kanban_tasks, 1)]
    _xl_tabla(ws, 3, ["N°", "Tarea", "Estado"], filas)
    ws.freeze_panes = "A4"
    _xl_autofit(ws)
    return _xl_bytes(wb)


def generar_excel_edt():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EDT WBS"
    _xl_titulo(ws, f"Estructura de desglose de trabajo · {st.session_state.nombre_proj}", 6)
    filas = _filas_edt()
    fin = _xl_tabla(ws, 3, ENC_EDT, filas, formatos={4: "#,##0.0", 5: "$#,##0"})
    if filas:
        ws.cell(row=fin, column=3, value="TOTAL").font = Font(bold=True)
        ws.cell(row=fin, column=4, value=f"=SUM(D4:D{fin - 1})").number_format = "#,##0.0"
        ws.cell(row=fin, column=5, value=f"=SUM(E4:E{fin - 1})").number_format = "$#,##0"
        ws.cell(row=fin, column=4).font = ws.cell(row=fin, column=5).font = Font(bold=True)
    ws.freeze_panes = "A4"
    _xl_autofit(ws)
    return _xl_bytes(wb)


def generar_excel_riesgos():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Matriz de Riesgos"
    _xl_titulo(ws, f"Matriz de riesgos · {st.session_state.nombre_proj}", 5)
    _xl_tabla(ws, 3, ["Riesgo", "Probabilidad (1-5)", "Impacto (1-5)", "Severidad (P×I)", "Nivel"],
              _filas_riesgos(), resaltar=lambda i, f: f[3] >= 16)
    ws.freeze_panes = "A4"
    _xl_autofit(ws)
    return _xl_bytes(wb)


@st.cache_data(show_spinner=False, max_entries=8)
def generar_excel_montecarlo(dias, costos, cfg):
    wb = openpyxl.Workbook()
    met = calcular_metricas(dias, costos, cfg["dias_meta"], cfg["costo_meta"], cfg["perc"])
    perc = cfg["perc"]

    # --- Hoja 1: Resumen ejecutivo
    ws = wb.active
    ws.title = "Resumen"
    _xl_titulo(ws, f"Simulación Monte Carlo · {cfg['proyecto']}", 3)
    f = _xl_meta(ws, 3, [
        ("Proyecto", cfg["proyecto"]),
        ("Fuente de estimación", cfg["fuente"]),
        ("Distribución", cfg["dist"]),
        ("Parámetros", cfg["parametros"]),
        ("Iteraciones", met["n"]),
        ("Fecha de inicio", cfg["inicio"]),
        ("Calendario", "Días hábiles" if cfg["habiles"] else "Días corridos"),
        ("Plazo meta (días)", cfg["dias_meta"]),
        ("Presupuesto meta ($)", cfg["costo_meta"]),
        ("Confianza exigida", f"P{perc}"),
    ])
    ws.cell(row=f - 2, column=2).number_format = "#,##0.0"
    ws.cell(row=f - 1, column=2).number_format = "$#,##0"
    _, titulo_v = veredicto(met, perc)
    f += 1
    ws.cell(row=f, column=1, value="Veredicto").font = Font(bold=True, size=12, color="0F172A")
    ws.cell(row=f, column=2, value=titulo_v).font = Font(bold=True, size=12, color=XL_NAVY)
    f += 2

    filas_t = [
        ["Duración mediana (P50)", met["p50"], "días"],
        [f"Duración comprometida (P{perc})", met["pt"], "días"],
        ["Duración pesimista (P95)", met["p95"], "días"],
        ["Duración promedio", met["media"], "días"],
        ["Desviación estándar", met["desv"], "días"],
        ["Reserva de contingencia de tiempo", met["reserva_dias"], "días"],
        ["Probabilidad de cumplir el plazo", met["prob_t"] / 100, "%"],
        ["Probabilidad de cumplir el presupuesto", met["prob_c"] / 100, "%"],
        ["Probabilidad de cumplir ambos", met["prob_ambos"] / 100, "%"],
    ]
    fmt_t = {}
    f_ini = f
    f = _xl_tabla(ws, f, ["Indicador de tiempo y éxito", "Valor", "Unidad"], filas_t,
                  resaltar=lambda i, r: "comprometida" in r[0] or "ambos" in r[0])
    for i, r in enumerate(filas_t):
        ws.cell(row=f_ini + 1 + i, column=2).number_format = "0.0%" if r[2] == "%" else "#,##0.0"
        if r[2] == "%":
            v = r[1]
            ws.cell(row=f_ini + 1 + i, column=2).fill = PatternFill(
                "solid", start_color=XL_OK if v >= perc / 100 else (XL_BAD if v < 0.5 else XL_HL))
    f += 1
    filas_c = [
        ["Costo mediano (P50)", met["p50c"]],
        [f"Costo recomendado (P{perc})", met["ptc"]],
        ["Costo pesimista (P95)", met["p95c"]],
        ["Reserva de contingencia de costo", met["reserva_costo"]],
    ]
    _xl_tabla(ws, f, ["Indicador financiero", "Valor ($)"], filas_c, formatos={2: "$#,##0"},
              resaltar=lambda i, r: "recomendado" in r[0])
    _xl_autofit(ws)
    ws.column_dimensions["A"].width = 42

    # --- Hoja 2: Percentiles
    ws2 = wb.create_sheet("Percentiles")
    _xl_titulo(ws2, "Tabla de confianza (percentiles)", 4)
    tp = tabla_percentiles(dias, costos, pd.Timestamp(cfg["inicio"]), cfg["habiles"], perc)
    _xl_tabla(ws2, 3, ["Confianza", "Días", "Fecha estimada de término", "Costo ($)"],
              [[r["Confianza"], r["Días"], r["Fecha"].strftime("%d/%m/%Y"), r["Costo"]] for r in tp],
              formatos={2: "#,##0.0", 4: "$#,##0"}, resaltar=lambda i, r: r[0] == f"P{perc}")
    _xl_autofit(ws2)

    # --- Hoja 3: Histograma con gráfico nativo de Excel
    ws3 = wb.create_sheet("Histograma")
    _xl_titulo(ws3, "Distribución de duraciones simuladas", 4)
    conteos, bordes = np.histogram(dias, bins=30)
    filas_h = [[f"{bordes[i]:.1f} – {bordes[i + 1]:.1f}", int(conteos[i]), float(conteos[i] / len(dias)),
                float((dias <= bordes[i + 1]).mean())] for i in range(len(conteos))]
    _xl_tabla(ws3, 3, ["Rango (días)", "Escenarios", "% del total", "% acumulado"], filas_h,
              formatos={3: "0.0%", 4: "0.0%"})
    _xl_autofit(ws3)
    graf = BarChart()
    graf.title = "Frecuencia de duraciones"
    graf.y_axis.title = "Escenarios"
    graf.x_axis.title = "Días"
    graf.add_data(Reference(ws3, min_col=2, min_row=3, max_row=3 + len(filas_h)), titles_from_data=True)
    graf.set_categories(Reference(ws3, min_col=1, min_row=4, max_row=3 + len(filas_h)))
    graf.gapWidth = 10
    graf.height, graf.width = 9, 18
    graf.legend = None
    ws3.add_chart(graf, "F3")
    curva = LineChart()
    curva.title = "Curva S (probabilidad acumulada)"
    curva.add_data(Reference(ws3, min_col=4, min_row=3, max_row=3 + len(filas_h)), titles_from_data=True)
    curva.set_categories(Reference(ws3, min_col=1, min_row=4, max_row=3 + len(filas_h)))
    curva.height, curva.width = 9, 18
    curva.legend = None
    ws3.add_chart(curva, "F23")

    # --- Hoja 4: Datos crudos
    ws4 = wb.create_sheet("Simulaciones")
    ws4.append(["Iteración", "Duración (días)", "Costo ($)", "Cumple plazo", "Cumple presupuesto"])
    for c in ws4[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=XL_ACC, end_color=XL_ACC)
    for i, (d, c) in enumerate(zip(dias, costos), 1):
        ws4.append([i, round(float(d), 2), round(float(c), 0),
                    "Sí" if d <= cfg["dias_meta"] else "No", "Sí" if c <= cfg["costo_meta"] else "No"])
    for col, w in zip("ABCDE", [12, 18, 16, 14, 20]):
        ws4.column_dimensions[col].width = w
    ws4.freeze_panes = "A2"
    return _xl_bytes(wb)


def generar_excel_estilizado():
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Requisitos"
    _xl_titulo(ws1, f"Requisitos · {st.session_state.nombre_proj}", 3)
    _xl_tabla(ws1, 3, ["Código", "Tipo", "Descripción"], _filas_requisitos())
    _xl_autofit(ws1)

    ws2 = wb.create_sheet("Kanban")
    _xl_titulo(ws2, "Tablero Kanban", 2)
    _xl_tabla(ws2, 3, ["Tarea", "Estado"], [[t["Tarea"], t["Estado"]] for t in st.session_state.kanban_tasks])
    _xl_autofit(ws2)

    ws3 = wb.create_sheet("EDT (WBS)")
    _xl_titulo(ws3, "Estructura de desglose de trabajo", 6)
    _xl_tabla(ws3, 3, ENC_EDT, _filas_edt(), formatos={4: "#,##0.0", 5: "$#,##0"})
    _xl_autofit(ws3)

    ws4 = wb.create_sheet("Riesgos")
    _xl_titulo(ws4, "Matriz de riesgos", 5)
    _xl_tabla(ws4, 3, ["Riesgo", "Probabilidad", "Impacto", "Severidad", "Nivel"], _filas_riesgos(),
              resaltar=lambda i, f: f[3] >= 16)
    _xl_autofit(ws4)

    if MC_ACTUAL:
        met, cfg = MC_ACTUAL["met"], MC_ACTUAL["cfg"]
        ws5 = wb.create_sheet("Monte Carlo")
        _xl_titulo(ws5, "Resumen Monte Carlo", 2)
        _xl_tabla(ws5, 3, ["Indicador", "Valor"], [
            ["Distribución", cfg["dist"]], ["Fuente", cfg["fuente"]], ["Iteraciones", met["n"]],
            ["P50 (días)", round(met["p50"], 1)], [f"P{cfg['perc']} (días)", round(met["pt"], 1)],
            ["P95 (días)", round(met["p95"], 1)], [f"Costo P{cfg['perc']} ($)", round(met["ptc"])],
            ["Prob. cumplir plazo", f"{met['prob_t']:.1f}%"], ["Prob. cumplir presupuesto", f"{met['prob_c']:.1f}%"],
            ["Prob. cumplir ambos", f"{met['prob_ambos']:.1f}%"],
        ])
        _xl_autofit(ws5)
    return _xl_bytes(wb)


def tabla_word(doc, encabezados, filas):
    t = doc.add_table(rows=1, cols=len(encabezados))
    try:
        t.style = "Light Grid Accent 1"
    except Exception:
        t.style = "Table Grid"
    for i, h in enumerate(encabezados):
        celda = t.rows[0].cells[i]
        celda.text = ""
        celda.paragraphs[0].add_run(str(h)).bold = True
    for f in filas:
        cells = t.add_row().cells
        for i, v in enumerate(f):
            cells[i].text = str(v)
    doc.add_paragraph()
    return t


def parrafo_analisis(doc, titulo, texto):
    p = doc.add_paragraph()
    r = p.add_run(titulo + " ")
    r.bold = True
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    p.add_run(texto)


def _doc_base():
    doc = Document()
    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(11)
    return doc


def generar_word_ers():
    doc = _doc_base()
    doc.add_heading(f"Especificación de requisitos: {st.session_state.nombre_proj}", level=0)
    doc.add_paragraph(f"Equipo: {st.session_state.integrantes}")
    doc.add_paragraph(f"Objetivo: {st.session_state.objetivo_text}")
    doc.add_heading("Matriz de requisitos", level=1)
    filas = _filas_requisitos()
    if filas:
        tabla_word(doc, ["Código", "Tipo", "Descripción"], filas)
    else:
        doc.add_paragraph("No hay requisitos registrados.")
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def generar_word_mementomori():
    doc = _doc_base()
    nombre = st.session_state.nombre_proj

    doc.add_heading(f"Reporte Memento Mori: {nombre}", level=0)
    doc.add_paragraph(f"Integrantes / responsables: {st.session_state.integrantes}")
    doc.add_paragraph("Análisis crítico integral de requisitos, arquitectura, tiempos, EDT y riesgos del sistema.")

    # 1. Introducción
    doc.add_heading("1. Introducción y objetivos del proyecto", level=1)
    doc.add_paragraph(f"Objetivo principal: {st.session_state.objetivo_text}")
    parrafo_analisis(doc, "Análisis crítico e interpretación:",
        f"El objetivo planteado para el proyecto '{nombre}' aborda la necesidad clave de automatización e integración. "
        "Sin embargo, el punto crítico principal radica en controlar el alcance funcional inicial. "
        "Bajo el análisis Memento Mori, el principal riesgo de esta fase es la expansión descontrolada del alcance (scope creep), "
        "la cual puede desalinear las expectativas de los interesados antes de consolidar el producto mínimo viable (MVP).")

    # 2. Requisitos
    doc.add_heading("2. Especificación de requisitos (RF y RNF)", level=1)
    filas = _filas_requisitos()
    if filas:
        tabla_word(doc, ["Código", "Tipo", "Descripción"], filas)
    parrafo_analisis(doc, "Análisis crítico de requisitos:",
        "Los puntos críticos se concentran en los requisitos no funcionales (RNF). "
        "Si bien los requisitos funcionales aseguran la operatividad del software, un fallo en el cumplimiento de los RNF "
        "(latencias elevadas o vulnerabilidades de seguridad) puede provocar el rechazo de los usuarios finales "
        "e impedir la escalabilidad del sistema.")

    # 3. Arquitectura
    doc.add_heading("3. Arquitectura y diagramas de flujo / secuencia", level=1)
    doc.add_paragraph("Sintaxis del diagrama activo:")
    p_code = doc.add_paragraph()
    r_code = p_code.add_run(st.session_state.flujo_codigo)
    r_code.font.name = "Consolas"
    r_code.font.size = Pt(9)
    parrafo_analisis(doc, "Interpretación de arquitectura y puntos críticos:",
        "Los diagramas permiten entender la interacción de componentes y el recorrido de la información. "
        "El punto crítico es el acoplamiento directo entre el controlador y los servicios de base de datos o autenticación. "
        "Sin políticas de reintento, caché o manejo de excepciones HTTP (401, 500), una caída puntual de la base de datos "
        "provocará un colapso en cadena en los endpoints expuestos al usuario.")

    # 4. Kanban
    doc.add_heading("4. Estado de gestión de tareas (tablero Kanban)", level=1)
    tareas = st.session_state.kanban_tasks
    if tareas:
        tabla_word(doc, ["Tarea", "Estado"], [[t["Tarea"], t["Estado"]] for t in tareas])
        n_proc = sum(1 for t in tareas if t["Estado"] == "En Proceso")
        n_ok = sum(1 for t in tareas if t["Estado"] == "Completado")
        doc.add_paragraph(f"Avance: {n_ok} de {len(tareas)} tareas completadas ({n_ok / len(tareas) * 100:.0f}%). "
                          f"Trabajo en proceso (WIP): {n_proc} tareas.")
    else:
        doc.add_paragraph("No existen tareas registradas.")
    parrafo_analisis(doc, "Análisis crítico del flujo Kanban:",
        "La acumulación de ítems 'En Proceso' es la mayor amenaza para el ritmo de entrega (WIP). "
        "Se sugiere definir un límite estricto de tareas simultáneas para evitar sobrecarga y sobreestimación del rendimiento.")

    # 5. EDT
    doc.add_heading("5. Estructura de desglose de trabajo (EDT / WBS)", level=1)
    edt = st.session_state.edt_list
    if edt:
        tabla_word(doc, ["Fase", "Paquete de trabajo", "Horas", "Costo ($)", "Documentos"],
                   [[e.get("Fase", ""), e.get("Paquete", ""), fmt_num(e.get("Horas", 0), 1),
                     fmt_money(e.get("Costo", 0)), e.get("Documentos", "")] for e in edt])
        tot_h = sum(float(e.get("Horas", 0) or 0) for e in edt)
        tot_c = sum(float(e.get("Costo", 0) or 0) for e in edt)
        doc.add_paragraph(f"Totales planificados: {fmt_num(tot_h, 1)} horas | Presupuesto: {fmt_money(tot_c)}")
    else:
        doc.add_paragraph("No hay paquetes de EDT registrados.")
    parrafo_analisis(doc, "Puntos críticos presupuestarios y de carga horaria:",
        "Las actividades con mayor asignación de costo y horas conforman la ruta crítica. Cualquier imprevisto de estimación "
        "en ellas provoca desviaciones acumulativas que agotan la reserva de contingencia y extienden la fecha límite.")

    # 6. Monte Carlo
    doc.add_heading("6. Análisis y simulación Monte Carlo", level=1)
    if MC_ACTUAL:
        met, cfg = MC_ACTUAL["met"], MC_ACTUAL["cfg"]
        perc = cfg["perc"]
        doc.add_paragraph(f"Fuente de estimación: {cfg['fuente']} | Distribución: {cfg['dist']} | "
                          f"Iteraciones: {fmt_num(met['n'])} | Parámetros: {cfg['parametros']}")
        doc.add_paragraph(f"Metas evaluadas: plazo de {fmt_num(cfg['dias_meta'], 1)} días y presupuesto de "
                          f"{fmt_money(cfg['costo_meta'])}, con confianza exigida P{perc}.")
        tabla_word(doc, ["Indicador", "Valor"], [
            ["Duración mediana (P50)", f"{fmt_num(met['p50'], 1)} días"],
            [f"Duración comprometida (P{perc})", f"{fmt_num(met['pt'], 1)} días"],
            ["Duración pesimista (P95)", f"{fmt_num(met['p95'], 1)} días"],
            ["Reserva de contingencia (tiempo)", f"{fmt_num(met['reserva_dias'], 1)} días"],
            [f"Costo recomendado (P{perc})", fmt_money(met["ptc"])],
            ["Reserva de contingencia (costo)", fmt_money(met["reserva_costo"])],
            ["Probabilidad de cumplir el plazo", f"{fmt_num(met['prob_t'], 1)}%"],
            ["Probabilidad de cumplir el presupuesto", f"{fmt_num(met['prob_c'], 1)}%"],
            ["Probabilidad de cumplir ambos", f"{fmt_num(met['prob_ambos'], 1)}%"],
        ])
        _, titulo_v = veredicto(met, perc)
        parrafo_analisis(doc, f"Veredicto: {titulo_v}.",
            f"De cada 100 escenarios simulados, {round(met['prob_ambos'])} terminan dentro de plazo y presupuesto. "
            f"Para comprometer la entrega con un {perc}% de confianza se requieren {fmt_num(met['pt'], 1)} días "
            f"(fecha estimada {MC_ACTUAL['fecha_pt'].strftime('%d/%m/%Y')}) y un presupuesto de {fmt_money(met['ptc'])}. "
            f"La diferencia frente a la mediana ({fmt_num(met['reserva_dias'], 1)} días y {fmt_money(met['reserva_costo'])}) "
            "corresponde a la reserva de contingencia recomendada. El escenario pesimista (P95) "
            f"alcanza {fmt_num(met['p95'], 1)} días, por lo que se aconsejan controles semanales de seguimiento.")
    else:
        doc.add_paragraph("Los parámetros configurados para la simulación Monte Carlo no son válidos.")

    # 7. Riesgos
    doc.add_heading("7. Evaluación y filosofía de riesgos (Memento Mori)", level=1)
    filas_r = _filas_riesgos()
    if filas_r:
        tabla_word(doc, ["Riesgo", "Probabilidad", "Impacto", "Severidad", "Nivel"], filas_r)
    else:
        doc.add_paragraph("No hay riesgos registrados en la matriz.")
    parrafo_analisis(doc, "Estrategia de mitigación Memento Mori:",
        "Bajo el principio 'Memento Mori' (recordar la vulnerabilidad y planificar anticipadamente los fallos), "
        "los riesgos de severidad alta y crítica deben contar con planes de contingencia documentados antes de iniciar el desarrollo.")

    # 8. Conclusiones
    doc.add_heading("8. Conclusiones y recomendaciones finales", level=1)
    doc.add_paragraph(
        f"El proyecto '{nombre}' cuenta con una base metodológica y técnica sólida. Se recomienda formalizar el compromiso "
        "de entrega según la simulación Monte Carlo y monitorear de forma continua los requisitos no funcionales y la matriz "
        "de riesgos para asegurar el éxito del proyecto.")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# =====================================================================
# 6. ENCABEZADO
# =====================================================================
col_hero, col_tema = st.columns([5, 1.25], vertical_alignment="bottom")
with col_hero:
    cursor = '<span class="cur">_</span>' if T["neon"] else ""
    st.markdown(
        f'<div class="md-hero"><h1>MatrixDev{cursor}</h1>'
        f'<div class="md-hero-sub">Gestión y arquitectura de proyectos de software</div>'
        f'<div class="md-chip"><span class="live"></span>Proyecto <b>{st.session_state.nombre_proj}</b></div></div>',
        unsafe_allow_html=True,
    )
with col_tema:
    st.markdown(f'<div style="text-align:center;color:{T["sub"]};font-size:.8rem;margin-bottom:4px">'
                f'Tema: <b style="color:{T["accent"]}">{T["nombre"]}</b></div>', unsafe_allow_html=True)
    b_cols = st.columns(3)
    for b_col, (clave, datos) in zip(b_cols, TEMAS.items()):
        b_col.button(datos["icono"], key=f"tema_btn_{clave}", help=f"Cambiar a tema {datos['nombre'].lower()}",
                     on_click=cambiar_tema, args=(clave,), **ANCHO)

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab_full = st.tabs([
    "Requisitos", "Modelado", "Gestión", "🎲 Monte Carlo", "Flujos", "EDT (WBS)", "Riesgos",
    "Reporte Memento Mori", "📦 Descarga completa",
])

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# ---------------------------------------------------------------------
# TAB 1: REQUISITOS
# ---------------------------------------------------------------------
with tab1:
    seccion("Definición del sistema", "Datos generales del proyecto. Se usan en todos los reportes exportados.")
    c_p1, c_p2 = st.columns(2)
    st.session_state.nombre_proj = c_p1.text_input("Nombre del proyecto", st.session_state.nombre_proj)
    st.session_state.integrantes = c_p2.text_input("Integrantes / equipo", st.session_state.integrantes)
    st.session_state.objetivo_text = st.text_area("Objetivo principal", st.session_state.objetivo_text, height=80)

    def bloque_requisitos(titulo, lista_key, prefijo, form_key, placeholder):
        st.write("")
        seccion(titulo)
        with st.form(form_key, clear_on_submit=True):
            c_in, c_b = st.columns([4, 1])
            nuevo = c_in.text_input("Nuevo", placeholder=placeholder, label_visibility="collapsed")
            if c_b.form_submit_button("Añadir", **ANCHO) and nuevo.strip():
                st.session_state[lista_key].append(nuevo.strip())
                st.rerun()
        lista = st.session_state[lista_key]
        if not lista:
            st.info("Aún no hay requisitos. Escribe uno arriba y pulsa Añadir.")
            return
        borrar = None
        for i, item in enumerate(lista):
            c_tag, c_input, c_del = st.columns([0.8, 5, 0.6])
            c_tag.markdown(f"**{prefijo}-{i + 1:02d}**")
            lista[i] = c_input.text_input(f"{prefijo}_{i}", value=item, label_visibility="collapsed",
                                          key=f"{prefijo.lower()}_field_{i}")
            if c_del.button("🗑️", key=f"del_{prefijo.lower()}_{i}", help="Eliminar requisito"):
                borrar = i
        if borrar is not None:
            lista.pop(borrar)
            limpiar_widgets(f"{prefijo.lower()}_field_")
            st.rerun()

    bloque_requisitos("Requisitos funcionales (RF)", "rf_list", "RF", "form_add_rf", "Escribe un requisito funcional…")
    bloque_requisitos("Requisitos no funcionales (RNF)", "rnf_list", "RNF", "form_add_rnf",
                      "Escribe un requisito no funcional…")

    st.divider()
    seccion("Exportar requisitos")
    c_dl1, c_dl2 = st.columns(2)
    c_dl1.download_button("📊 Requisitos en Excel (.xlsx)", generar_excel_requisitos(),
                          f"Requisitos_{slug(st.session_state.nombre_proj)}.xlsx", XLSX, **ANCHO)
    c_dl2.download_button("📄 Documento ERS en Word (.docx)", generar_word_ers(),
                          f"ERS_Requisitos_{slug(st.session_state.nombre_proj)}.docx", DOCX, **ANCHO)

# ---------------------------------------------------------------------
# TAB 2: MODELADO
# ---------------------------------------------------------------------
with tab2:
    seccion("Arquitectura visual", "Escribe sintaxis Mermaid a la izquierda y revisa el diagrama a la derecha.")
    c_code, c_prev = st.columns([1, 1.3])
    with c_code:
        st.text_area("Sintaxis Mermaid", key="modelo_mermaid", height=260)
        st.download_button("📌 Descargar diagrama (.mmd)", st.session_state.modelo_mermaid.encode("utf-8"),
                           f"Diagrama_Modelado_{slug(st.session_state.nombre_proj)}.mmd", "text/plain", **ANCHO)
    with c_prev:
        mostrar_mermaid(st.session_state.modelo_mermaid, "380px", key=f"mm_modelo_{st.session_state.tema}")

# ---------------------------------------------------------------------
# TAB 3: GESTIÓN (KANBAN)
# ---------------------------------------------------------------------
with tab3:
    seccion("Tablero Kanban", "Agrega tareas y muévelas cambiando su estado.")
    tareas = st.session_state.kanban_tasks

    with st.form("form_kanban", clear_on_submit=True):
        col_t, col_s, col_b = st.columns([3, 2, 1])
        nueva_t = col_t.text_input("Tarea", placeholder="Nueva tarea…", label_visibility="collapsed")
        estado_t = col_s.selectbox("Estado", ESTADOS_KANBAN, label_visibility="collapsed")
        if col_b.form_submit_button("Añadir", **ANCHO) and nueva_t.strip():
            tareas.append({"Tarea": nueva_t.strip(), "Estado": estado_t})
            st.rerun()

    if tareas:
        n_ok = sum(t["Estado"] == "Completado" for t in tareas)
        n_proc = sum(t["Estado"] == "En Proceso" for t in tareas)
        st.progress(n_ok / len(tareas), text=f"Avance: {n_ok} de {len(tareas)} tareas completadas ({n_ok / len(tareas) * 100:.0f}%)")
        if n_proc > 3:
            st.warning(f"Hay {n_proc} tareas en proceso a la vez. Limitar el WIP a 3 ayuda a terminar antes.")

        colores_estado = {"Pendiente": T["sub"], "En Proceso": T["warn"], "Completado": T["ok"]}
        cols_kb = st.columns(3)
        for n_col, (col, estado) in enumerate(zip(cols_kb, ESTADOS_KANBAN)):
            items = [t for t in tareas if t["Estado"] == estado]
            cards = "".join(f'<div class="kb-card" style="--c:{colores_estado[estado]}">{t["Tarea"]}</div>' for t in items)
            if not items:
                cards = f'<div style="color:{T["sub"]};font-size:.88rem">Sin tareas</div>'
            col.markdown(f'<div class="kb-col" style="--d:{n_col * 0.08:.2f}s"><div class="kb-h"><span>{estado}</span><span class="n">{len(items)}</span></div>'
                         f'{cards}</div>', unsafe_allow_html=True)

        st.write("")
        with st.expander("Editar o eliminar tareas", expanded=False):
            borrar = None
            for i, t in enumerate(tareas):
                c_txt, c_sel, c_d = st.columns([3, 2, 0.6])
                t["Tarea"] = c_txt.text_input(f"kt_{i}", value=t["Tarea"], label_visibility="collapsed", key=f"kt_in_{i}")
                t["Estado"] = c_sel.selectbox(f"ks_{i}", ESTADOS_KANBAN, index=ESTADOS_KANBAN.index(t["Estado"]),
                                              label_visibility="collapsed", key=f"ks_in_{i}")
                if c_d.button("🗑️", key=f"del_k_{i}", help="Eliminar tarea"):
                    borrar = i
            if borrar is not None:
                tareas.pop(borrar)
                limpiar_widgets("kt_in_", "ks_in_")
                st.rerun()
    else:
        st.info("No hay tareas. Escribe la primera arriba para empezar el tablero.")

    st.divider()
    seccion("Exportar tablero")
    c_dk1, c_dk2 = st.columns(2)
    c_dk1.download_button("📊 Kanban en Excel (.xlsx)", generar_excel_kanban(),
                          f"Kanban_{slug(st.session_state.nombre_proj)}.xlsx", XLSX, **ANCHO)
    c_dk2.download_button("📄 Tareas en CSV (.csv)", pd.DataFrame(tareas).to_csv(index=False).encode("utf-8"),
                          f"Kanban_{slug(st.session_state.nombre_proj)}.csv", "text/csv", **ANCHO)

# ---------------------------------------------------------------------
# TAB 4: MONTE CARLO
# ---------------------------------------------------------------------
with tab4:
    seccion("Simulación Monte Carlo de plazo y presupuesto",
            "Simula miles de versiones posibles del proyecto para responder dos preguntas: "
            "¿qué probabilidad hay de terminar a tiempo y dentro del presupuesto? y ¿qué plazo conviene comprometer?")

    with st.expander("¿Cómo funciona y cómo leer los resultados?", expanded=False):
        st.markdown("""
**La idea.** En vez de una sola estimación ("tardaremos 20 días"), defines un rango de incertidumbre.
El motor sortea miles de duraciones dentro de ese rango, como si ejecutaras el proyecto miles de veces, y cuenta cuántas veces se cumplen tus metas.

**Distribuciones disponibles**
- **Beta-PERT** (recomendada, estándar PMBOK): usa optimista, más probable y pesimista, y concentra los resultados cerca del valor más probable.
- **Triangular**: mismos tres valores, pero da más peso a los extremos. Es más conservadora.
- **Normal**: para procesos conocidos, definidos por un promedio y una desviación.
- **Uniforme**: cuando no sabes nada más que el mínimo y el máximo.

**Percentiles (P)**
- **P50**: la mitad de los escenarios termina antes. Es una apuesta 50/50, no un compromiso.
- **P85** (u otro que elijas): el plazo que puedes prometer con esa confianza. Es el estándar de la industria.
- **P95**: cubre casi todos los imprevistos; útil como escenario pesimista.

**Reserva de contingencia**: la diferencia entre el P50 y el percentil de compromiso. Es el colchón de tiempo y dinero que conviene negociar.
        """)

    # ---------------- Paso 1: estimación ----------------
    with st.container(border=True):
        paso(1, "Estimación de la duración")
        fuente = st.radio("Fuente de la estimación", ["Estimación manual (3 puntos)", "Desde la EDT (por paquete)"],
                          horizontal=True, key="mc_fuente",
                          help="La opción EDT simula cada paquete de trabajo por separado y suma sus duraciones.")
        desde_edt = fuente.startswith("Desde")

        c_e1, c_e2 = st.columns([1, 2])
        with c_e1:
            tipo_dist = st.selectbox("Distribución", DISTRIBUCIONES, key="mc_dist_type",
                                     help="Modelo matemático con el que se sortea la incertidumbre.")
        with c_e2:
            valido, error_msg = True, ""
            parametros_txt = ""
            if desde_edt:
                ce1, ce2, ce3 = st.columns(3)
                horas_dia = ce1.number_input("Horas de trabajo por día", 1.0, 24.0, 8.0, 0.5, key="mc_horas_dia")
                pct_opt = ce2.slider("Mejor caso (% menos)", 0, 80, 20, 5, key="mc_pct_opt",
                                     help="Cuánto más rápido podría salir cada paquete si todo va bien.")
                pct_pes = ce3.slider("Peor caso (% más)", 0, 300, 60, 10, key="mc_pct_pes",
                                     help="Cuánto más podría tardar cada paquete si hay problemas.")
                paquetes = tuple((str(e.get("Paquete", f"Paquete {i + 1}")) or f"Paquete {i + 1}",
                                  float(e.get("Horas", 0) or 0), float(e.get("Costo", 0) or 0))
                                 for i, e in enumerate(st.session_state.edt_list))
                base_h = sum(p[1] for p in paquetes)
                if not paquetes or base_h <= 0:
                    valido, error_msg = False, "La EDT no tiene paquetes con horas. Agrégalos en la pestaña EDT (WBS) o usa la estimación manual."
                else:
                    st.caption(f"Base de la EDT: {len(paquetes)} paquetes, {fmt_num(base_h, 1)} horas ≈ "
                               f"{fmt_num(base_h / horas_dia, 1)} días de trabajo si se ejecutan uno tras otro.")
                parametros_txt = f"{horas_dia} h/día, mejor caso −{pct_opt}%, peor caso +{pct_pes}%"
            elif tipo_dist in ("Triangular", "Beta-PERT"):
                ce1, ce2, ce3 = st.columns(3)
                mc_opt = ce1.number_input("Optimista (días)", min_value=0.5, value=10.0, step=1.0, format="%.1f", key="mc_opt")
                mc_prob = ce2.number_input("Más probable (días)", min_value=0.5, value=20.0, step=1.0, format="%.1f", key="mc_prob")
                mc_pes = ce3.number_input("Pesimista (días)", min_value=0.5, value=45.0, step=1.0, format="%.1f", key="mc_pes")
                if not (mc_opt <= mc_prob <= mc_pes) or mc_opt >= mc_pes:
                    valido, error_msg = False, "Revisa los valores: optimista ≤ más probable ≤ pesimista, y optimista < pesimista."
                else:
                    st.caption(f"Estimación PERT clásica (O + 4M + P) / 6 = {fmt_num((mc_opt + 4 * mc_prob + mc_pes) / 6, 1)} días.")
                parametros_txt = f"O={fmt_num(mc_opt, 1)}, M={fmt_num(mc_prob, 1)}, P={fmt_num(mc_pes, 1)} días"
            elif tipo_dist == "Normal":
                ce1, ce2 = st.columns(2)
                mc_mean = ce1.number_input("Duración promedio (días)", min_value=1.0, value=22.0, step=1.0, format="%.1f", key="mc_mean")
                mc_std = ce2.number_input("Desviación estándar (días)", min_value=0.1, value=5.0, step=0.5, format="%.1f", key="mc_std")
                st.caption(f"El 95% de los escenarios caerá aprox. entre {fmt_num(max(mc_mean - 2 * mc_std, 0.5), 1)} "
                           f"y {fmt_num(mc_mean + 2 * mc_std, 1)} días.")
                parametros_txt = f"media={fmt_num(mc_mean, 1)}, desv.={fmt_num(mc_std, 1)} días"
            else:
                ce1, ce2 = st.columns(2)
                mc_min = ce1.number_input("Mínimo (días)", min_value=0.5, value=12.0, step=1.0, format="%.1f", key="mc_min_u")
                mc_max = ce2.number_input("Máximo (días)", min_value=0.5, value=40.0, step=1.0, format="%.1f", key="mc_max_u")
                if mc_min >= mc_max:
                    valido, error_msg = False, "El mínimo debe ser menor que el máximo."
                parametros_txt = f"mín={fmt_num(mc_min, 1)}, máx={fmt_num(mc_max, 1)} días"

    # ---------------- Paso 2: costos y metas ----------------
    with st.container(border=True):
        paso(2, "Costos y metas a evaluar")
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            st.markdown("**Costos**")
            costo_fijo = st.number_input("Costo fijo base ($)", min_value=0, value=100000, step=10000, key="mc_cfijo")
            if desde_edt:
                st.caption("El costo variable sale de cada paquete de la EDT y crece si el paquete se demora.")
                costo_diario, var_costo = 0, 0
            else:
                costo_diario = st.number_input("Costo operativo por día ($)", min_value=0, value=15000, step=1000, key="mc_cdiario")
                var_costo = st.slider("Variabilidad del costo diario (±%)", 0, 30, 5, key="mc_var_costo",
                                      help="Imprevistos en el gasto diario, independientes de la duración.")
        with cc2:
            st.markdown("**Metas**")
            dias_meta = st.number_input("Plazo meta (días)", min_value=0.5, value=25.0, step=1.0, format="%.1f", key="mc_dias_meta")
            costo_meta = st.number_input("Presupuesto meta ($)", min_value=0, value=500000, step=25000, key="mc_costo_meta")
            percentil_target = st.slider("Confianza exigida (percentil)", 50, 99, 85, step=1, key="mc_target_perc",
                                         help="Con qué probabilidad quieres poder cumplir lo que prometas.")
        with cc3:
            st.markdown("**Calendario**")
            fecha_inicio = st.date_input("Fecha de inicio", value=date.today(), key="mc_f_inicio", format="DD/MM/YYYY")
            habiles = st.toggle("Contar solo días hábiles (lun–vie)", value=True, key="mc_habiles")
            st.caption("Afecta las fechas de término y la capacidad anual.")

    # ---------------- Paso 3: motor ----------------
    with st.container(border=True):
        paso(3, "Motor de simulación")
        cm1, cm2, cm3 = st.columns([3, 1.3, 1])
        simulaciones = cm1.select_slider("Número de iteraciones", options=[500, 1000, 2000, 5000, 10000, 20000, 50000],
                                         value=10000, key="mc_sims",
                                         help="Más iteraciones dan resultados más estables. 10.000 suele bastar.")
        usar_semilla = cm2.toggle("Resultados reproducibles", value=True, key="mc_seed",
                                  help="Usa siempre la misma semilla aleatoria para obtener los mismos números.")
        if not usar_semilla:
            if cm3.button("🎲 Nueva corrida", **ANCHO):
                st.session_state.mc_semilla_val = int(np.random.default_rng().integers(1, 10**9))
        semilla = 42 if usar_semilla else st.session_state.mc_semilla_val

    if not valido:
        st.error(error_msg)
    else:
        with st.spinner("Simulando escenarios…"):
            if desde_edt:
                res = simular_edt(paquetes, float(horas_dia), float(pct_opt), float(pct_pes), tipo_dist,
                                  int(simulaciones), float(costo_fijo), int(semilla))
            elif tipo_dist == "Normal":
                res = simular_manual(tipo_dist, float(mc_mean), float(mc_std), 0.0, int(simulaciones),
                                     float(costo_fijo), float(costo_diario), float(var_costo), int(semilla))
            elif tipo_dist == "Uniforme":
                res = simular_manual(tipo_dist, float(mc_min), 0.0, float(mc_max), int(simulaciones),
                                     float(costo_fijo), float(costo_diario), float(var_costo), int(semilla))
            else:
                res = simular_manual(tipo_dist, float(mc_opt), float(mc_prob), float(mc_pes), int(simulaciones),
                                     float(costo_fijo), float(costo_diario), float(var_costo), int(semilla))

        dias, costos = res["dias"], res["costos"]
        perc = int(percentil_target)
        met = calcular_metricas(dias, costos, dias_meta, costo_meta, perc)
        f_p50 = fecha_fin(fecha_inicio, met["p50"], habiles)
        f_pt = fecha_fin(fecha_inicio, met["pt"], habiles)
        f_p95 = fecha_fin(fecha_inicio, met["p95"], habiles)
        cfg = {
            "proyecto": st.session_state.nombre_proj, "fuente": fuente, "dist": tipo_dist,
            "parametros": parametros_txt, "dias_meta": float(dias_meta), "costo_meta": float(costo_meta),
            "perc": perc, "inicio": str(fecha_inicio), "habiles": bool(habiles),
        }
        MC_ACTUAL = {"res": res, "met": met, "cfg": cfg, "fecha_pt": f_pt}

        ok_t = dias <= dias_meta
        ok_c = costos <= costo_meta

        # ---------------- Veredicto ----------------
        st.write("")
        nivel, titulo_v = veredicto(met, perc)
        color_v = T[nivel]
        n_ok = int(round(met["prob_ambos"]))
        n_mid = int(round(met["prob_ambos"] + met["prob_uno"])) - n_ok
        n_mid = max(0, min(n_mid, 100 - n_ok))
        clases = ["ok"] * n_ok + ["mid"] * n_mid + ["bad"] * (100 - n_ok - n_mid)
        puntos = "".join(f'<span class="dot {c}" style="animation-delay:{i * 0.008:.3f}s"></span>' for i, c in enumerate(clases))

        if nivel == "ok":
            recomendacion = (f"Puedes comprometer <b>{fmt_num(met['pt'], 1)} días</b> (término el <b>{f_pt.strftime('%d/%m/%Y')}</b>) "
                             f"y <b>{fmt_money(met['ptc'])}</b> con un {perc}% de confianza, dentro de tus metas.")
        else:
            falta = []
            if met["pt"] > dias_meta:
                falta.append(f"ampliar el plazo a <b>{fmt_num(np.ceil(met['pt']), 0)} días</b>")
            if met["ptc"] > costo_meta:
                falta.append(f"ampliar el presupuesto a <b>{fmt_money(met['ptc'])}</b>")
            recomendacion = (f"Para alcanzar un {perc}% de confianza habría que " + " y ".join(falta) + ", o reducir el alcance."
                             if falta else
                             f"Cada meta por separado se cumple con {perc}% de confianza, pero ambas a la vez no: "
                             "revisa la holgura de plazo o presupuesto.")

        st.markdown(
            f'<div class="md-verdict" style="--c:{color_v};--cb:{rgba(color_v, 0.13)}">'
            f'<div><span class="v-tag">{titulo_v}</span>'
            f'<div class="v-big">De cada 100 escenarios simulados, <b>{n_ok}</b> terminan a tiempo y dentro del presupuesto.</div>'
            f'<div class="v-txt">{recomendacion}<br>Reserva de contingencia sugerida: '
            f'<b>+{fmt_num(met["reserva_dias"], 1)} días</b> y <b>+{fmt_money(met["reserva_costo"])}</b> sobre la mediana.</div></div>'
            f'<div class="dotwrap"><div class="dotgrid">{puntos}</div>'
            f'<div class="dotleg"><span><i style="background:{T["ok"]}"></i>Cumple ambas</span>'
            f'<span><i style="background:{T["warn"]}"></i>Cumple una</span>'
            f'<span><i style="background:{T["bad"]}"></i>Ninguna</span></div></div></div>',
            unsafe_allow_html=True,
        )

        # ---------------- KPIs ----------------
        st.write("")
        k1, k2, k3, k4 = st.columns(4)
        kpi(k1, "Entrega probable (P50)", f"{fmt_num(met['p50'], 1)} días",
            f"Término {f_p50.strftime('%d/%m/%Y')} · apuesta 50/50", T["sub"], 0.0)
        kpi(k2, f"Plazo a comprometer (P{perc})", f"{fmt_num(met['pt'], 1)} días",
            f"Término {f_pt.strftime('%d/%m/%Y')}", T["accent"], 0.08)
        kpi(k3, "Escenario pesimista (P95)", f"{fmt_num(met['p95'], 1)} días",
            f"Término {f_p95.strftime('%d/%m/%Y')}", T["warn"], 0.16)
        kpi(k4, f"Presupuesto a comprometer (P{perc})", fmt_money(met["ptc"]),
            f"Mediana {fmt_money(met['p50c'])}", T["accent2"], 0.24)

        # ---------------- Indicadores tipo velocímetro ----------------
        def color_prob(v):
            return T["ok"] if v >= perc else (T["warn"] if v >= 50 else T["bad"])

        fig_g = go.Figure()
        for i, (val, tit) in enumerate([(met["prob_t"], f"Cumplir plazo ({fmt_num(dias_meta, 0)} días)"),
                                        (met["prob_c"], f"Cumplir presupuesto ({fmt_money(costo_meta)})"),
                                        (met["prob_ambos"], "Cumplir ambos")]):
            fig_g.add_trace(go.Indicator(
                mode="gauge+number", value=val,
                number={"suffix": "%", "valueformat": ".1f", "font": {"size": 30, "color": color_prob(val)}},
                title={"text": tit, "font": {"size": 14, "color": T["sub"]}},
                domain={"x": [i / 3 + 0.04, (i + 1) / 3 - 0.04], "y": [0, 1]},
                gauge={
                    "axis": {"range": [0, 100], "tickcolor": T["sub"], "tickwidth": 1, "ticksuffix": "%"},
                    "bar": {"color": color_prob(val), "thickness": 0.28},
                    "bgcolor": "rgba(0,0,0,0)", "borderwidth": 0,
                    "steps": [{"range": [0, 50], "color": rgba(T["bad"], 0.12)},
                              {"range": [50, perc], "color": rgba(T["warn"], 0.12)},
                              {"range": [perc, 100], "color": rgba(T["ok"], 0.12)}],
                    "threshold": {"line": {"color": T["text"], "width": 3}, "thickness": 0.85, "value": perc},
                },
            ))
        estilo_fig(fig_g, height=230, margin=dict(l=20, r=20, t=50, b=0))
        mostrar_fig(fig_g)
        st.caption(f"La marca vertical en cada indicador es tu confianza exigida (P{perc}). "
                   "Verde: la supera · Ámbar: entre 50% y la meta · Rojo: menos de 50%.")

        # ---------------- Gráficos ----------------
        st.write("")
        seccion("Explorar los resultados")
        nombres_tabs = ["Distribución del plazo", "Curva S y calculadora", "Plazo vs. costo", "Distribución del costo",
                        "Convergencia"]
        if res["componentes"] is not None:
            nombres_tabs.append("Sensibilidad por paquete")
        g_tabs = st.tabs(nombres_tabs)

        def marcar(fig, marcas):
            for i, (x, txt, col, dash) in enumerate(marcas):
                fig.add_vline(x=x, line_color=col, line_dash=dash, line_width=2)
                fig.add_annotation(x=x, y=1 - i * 0.085, yref="paper", text=txt, showarrow=False, xanchor="left",
                                   xshift=5, font=dict(color=col, size=12), bgcolor=rgba(T["card"], 0.85))

        def histograma_coloreado(datos, meta, etiqueta_x, es_dinero=False):
            conteos, bordes = np.histogram(datos, bins=45)
            centros = (bordes[:-1] + bordes[1:]) / 2
            pct = conteos / len(datos) * 100
            ancho = np.diff(bordes)
            dentro = bordes[1:] <= meta
            fmt_h = "$%{customdata[0]:,.0f} – $%{customdata[1]:,.0f}" if es_dinero else "%{customdata[0]:.1f} – %{customdata[1]:.1f} días"
            fig = go.Figure()
            for mask, nombre, col in [(dentro, "Cumple la meta", T["ok"]), (~dentro, "Excede la meta", T["bad"])]:
                if mask.any():
                    fig.add_bar(x=centros[mask], y=pct[mask], width=ancho[mask] * 0.95, name=nombre,
                                marker=dict(color=rgba(col, 0.75), line=dict(width=0)),
                                customdata=np.stack([bordes[:-1][mask], bordes[1:][mask]], axis=-1),
                                hovertemplate=fmt_h + "<br>%{y:.2f}% de los escenarios<extra></extra>")
            fig.update_layout(barmode="overlay", bargap=0)
            estilo_fig(fig, height=420)
            fig.update_xaxes(title=etiqueta_x)
            fig.update_yaxes(title="% de escenarios", ticksuffix="%")
            return fig

        with g_tabs[0]:
            fig_h = histograma_coloreado(dias, dias_meta, "Duración del proyecto (días)")
            marcar(fig_h, [(dias_meta, f"Meta {fmt_num(dias_meta, 1)} d", T["text"], "dash"),
                           (met["p50"], f"P50 {fmt_num(met['p50'], 1)} d", T["sub"], "dot"),
                           (met["pt"], f"P{perc} {fmt_num(met['pt'], 1)} d", T["accent"], "solid"),
                           (met["p95"], f"P95 {fmt_num(met['p95'], 1)} d", T["warn"], "dot")])
            mostrar_fig(fig_h)
            st.caption(f"Cada barra agrupa escenarios con duración parecida. Las barras verdes ({fmt_num(met['prob_t'], 1)}%) "
                       f"terminan dentro del plazo meta; las rojas ({fmt_num(met['riesgo_t'], 1)}%) se atrasan. "
                       f"Promedio {fmt_num(met['media'], 1)} días, desviación {fmt_num(met['desv'], 1)} días "
                       f"(variabilidad {fmt_num(met['cv'], 0)}%).")

        with g_tabs[1]:
            xs = np.sort(dias)
            ys = np.arange(1, len(xs) + 1) / len(xs) * 100
            idx = np.linspace(0, len(xs) - 1, min(len(xs), 700)).astype(int)
            fig_s = go.Figure()
            fig_s.add_scatter(x=xs[idx], y=ys[idx], mode="lines", line=dict(color=T["accent"], width=3),
                              fill="tozeroy", fillcolor=rgba(T["accent"], 0.10), name="Probabilidad acumulada",
                              hovertemplate="Terminar en ≤ %{x:.1f} días<br>Probabilidad: %{y:.1f}%<extra></extra>")
            pts = [(met["p50"], 50, "P50", T["sub"]), (met["pt"], perc, f"P{perc}", T["accent"]), (met["p95"], 95, "P95", T["warn"])]
            fig_s.add_scatter(x=[p[0] for p in pts], y=[p[1] for p in pts], mode="markers+text",
                              text=[f"{p[2]}: {fmt_num(p[0], 1)} d" for p in pts], textposition="middle left",
                              marker=dict(size=11, color=[p[3] for p in pts], line=dict(width=2, color=T["card"])),
                              textfont=dict(color=T["text"]), hoverinfo="skip", showlegend=False)
            fig_s.add_shape(type="line", x0=dias_meta, x1=dias_meta, y0=0, y1=met["prob_t"],
                            line=dict(color=T["text"], dash="dash", width=1.5))
            fig_s.add_shape(type="line", x0=xs.min(), x1=dias_meta, y0=met["prob_t"], y1=met["prob_t"],
                            line=dict(color=T["text"], dash="dash", width=1.5))
            fig_s.add_annotation(x=dias_meta, y=met["prob_t"], text=f"Meta: {fmt_num(met['prob_t'], 1)}% de probabilidad",
                                 showarrow=True, arrowhead=0, ax=60, ay=30, font=dict(color=T["text"]),
                                 bgcolor=rgba(T["card"], 0.9), bordercolor=T["border"])
            estilo_fig(fig_s, height=420, showlegend=False)
            fig_s.update_xaxes(title="Duración del proyecto (días)")
            fig_s.update_yaxes(title="Probabilidad de terminar a esa fecha o antes", ticksuffix="%", range=[0, 103])
            mostrar_fig(fig_s)
            st.caption("Lee la curva así: elige un número de días en el eje horizontal y sube hasta la curva; "
                       "la altura es la probabilidad de terminar en ese plazo o antes.")

            st.markdown("**Calculadora de confianza**")
            q1, q2 = st.columns(2)
            with q1:
                x_dias = st.number_input("Si me comprometo a … días", min_value=0.5,
                                         value=float(np.ceil(met["pt"])), step=1.0, format="%.1f", key="mc_calc_dias")
                prob_x = (dias <= x_dias).mean() * 100
                st.markdown(f'<div class="md-answer">Probabilidad de cumplir: <b>{fmt_num(prob_x, 1)}%</b><br>'
                            f'Término el {fecha_fin(fecha_inicio, x_dias, habiles).strftime("%d/%m/%Y")}</div>',
                            unsafe_allow_html=True)
            with q2:
                conf = st.slider("Quiero esta confianza (%)", 50, 99, 90, key="mc_calc_conf")
                d_conf = float(np.percentile(dias, conf))
                c_conf = float(np.percentile(costos, conf))
                st.markdown(f'<div class="md-answer">Necesitas: <b>{fmt_num(d_conf, 1)} días</b> y '
                            f'<b>{fmt_money(c_conf)}</b><br>Término el '
                            f'{fecha_fin(fecha_inicio, d_conf, habiles).strftime("%d/%m/%Y")}</div>',
                            unsafe_allow_html=True)

        with g_tabs[2]:
            cat = np.where(ok_t & ok_c, "Cumple plazo y presupuesto",
                           np.where(ok_t, "Solo cumple plazo", np.where(ok_c, "Solo cumple presupuesto", "No cumple ninguno")))
            rng_vis = np.random.default_rng(7)
            muestra = rng_vis.choice(len(dias), size=min(2500, len(dias)), replace=False)
            colores_cat = {"Cumple plazo y presupuesto": T["ok"], "Solo cumple plazo": T["accent"],
                           "Solo cumple presupuesto": T["warn"], "No cumple ninguno": T["bad"]}
            fig_sc = go.Figure()
            for nombre_c, col in colores_cat.items():
                m_full = cat == nombre_c
                m = m_full[muestra]
                if not m_full.any():
                    continue
                fig_sc.add_trace(go.Scattergl(
                    x=dias[muestra][m], y=costos[muestra][m], mode="markers",
                    name=f"{nombre_c} ({fmt_num(m_full.mean() * 100, 1)}%)",
                    marker=dict(color=rgba(col, 0.55), size=6),
                    hovertemplate="%{x:.1f} días<br>$%{y:,.0f}<extra></extra>"))
            fig_sc.add_vline(x=dias_meta, line_dash="dash", line_color=T["text"], line_width=1.5)
            fig_sc.add_hline(y=costo_meta, line_dash="dash", line_color=T["text"], line_width=1.5)
            estilo_fig(fig_sc, height=440)
            fig_sc.update_xaxes(title="Duración (días)")
            fig_sc.update_yaxes(title="Costo total ($)", tickprefix="$")
            mostrar_fig(fig_sc)
            st.caption("Cada punto es un escenario. Las líneas punteadas son tus metas: el cuadrante inferior izquierdo "
                       "es donde se cumple todo. Se muestran hasta 2.500 puntos; los porcentajes usan todas las iteraciones.")

        with g_tabs[3]:
            fig_c = histograma_coloreado(costos, costo_meta, "Costo total del proyecto ($)", es_dinero=True)
            marcar(fig_c, [(costo_meta, f"Meta {fmt_money(costo_meta)}", T["text"], "dash"),
                           (met["p50c"], f"P50 {fmt_money(met['p50c'])}", T["sub"], "dot"),
                           (met["ptc"], f"P{perc} {fmt_money(met['ptc'])}", T["accent"], "solid")])
            fig_c.update_xaxes(tickprefix="$")
            mostrar_fig(fig_c)
            st.caption(f"{fmt_num(met['prob_c'], 1)}% de los escenarios queda dentro del presupuesto meta. "
                       f"Costo promedio: {fmt_money(met['media_c'])}.")

        with g_tabs[4]:
            cps = np.unique(np.linspace(min(100, len(dias)), len(dias), 70).astype(int))
            media_run = np.cumsum(dias)[cps - 1] / cps
            p_run = np.array([np.percentile(dias[:c], perc) for c in cps])
            fig_cv = go.Figure()
            fig_cv.add_hrect(y0=met["pt"] * 0.99, y1=met["pt"] * 1.01, fillcolor=rgba(T["accent"], 0.12), line_width=0)
            fig_cv.add_scatter(x=cps, y=p_run, mode="lines", name=f"P{perc} acumulado", line=dict(color=T["accent"], width=2.5))
            fig_cv.add_scatter(x=cps, y=media_run, mode="lines", name="Promedio acumulado", line=dict(color=T["sub"], width=2))
            estilo_fig(fig_cv, height=380)
            fig_cv.update_xaxes(title="Iteraciones ejecutadas")
            fig_cv.update_yaxes(title="Días")
            mostrar_fig(fig_cv)
            estable = np.all(np.abs(p_run[len(p_run) // 2:] - met["pt"]) <= met["pt"] * 0.01)
            st.caption(("✅ El resultado es estable: " if estable else "⚠️ El resultado todavía varía: ")
                       + f"en la segunda mitad de la simulación el P{perc} se mueve "
                       + ("menos de ±1% (banda sombreada)." if estable else "más de ±1%. Aumenta las iteraciones."))

        if res["componentes"] is not None:
            with g_tabs[5]:
                comp = res["componentes"]
                var_total = np.var(dias)
                aporte = np.array([np.cov(comp[:, j], dias)[0, 1] / var_total * 100 if var_total > 0 else 0
                                   for j in range(comp.shape[1])])
                df_sens = (pd.DataFrame({"Paquete": res["nombres"], "Aporte": aporte,
                                         "Mediana (días)": np.median(comp, axis=0)})
                           .sort_values("Aporte").tail(12))
                fig_t = go.Figure(go.Bar(
                    x=df_sens["Aporte"], y=df_sens["Paquete"], orientation="h",
                    marker_color=[T["bad"] if v == df_sens["Aporte"].max() else T["accent"] for v in df_sens["Aporte"]],
                    text=[f"{fmt_num(v, 1)}%" for v in df_sens["Aporte"]], textposition="outside",
                    hovertemplate="%{y}<br>Aporte a la incertidumbre: %{x:.1f}%<extra></extra>"))
                estilo_fig(fig_t, height=max(300, 60 * len(df_sens)), showlegend=False)
                fig_t.update_xaxes(title="Aporte a la incertidumbre total (%)", ticksuffix="%")
                mostrar_fig(fig_t)
                top = df_sens.iloc[-1]
                st.caption(f"'{top['Paquete']}' explica el {fmt_num(top['Aporte'], 1)}% de la variación del plazo total. "
                           "Afinar su estimación o reducir su riesgo es lo que más acota la incertidumbre del proyecto.")

        # ---------------- Tabla de confianza ----------------
        st.write("")
        seccion("Tabla de confianza", "Qué plazo, fecha y presupuesto corresponde a cada nivel de confianza.")
        tp = tabla_percentiles(dias, costos, fecha_inicio, habiles, perc)
        df_tp = pd.DataFrame([{
            "Confianza": r["Confianza"],
            "Días": fmt_num(r["Días"], 1),
            "Fecha de término": r["Fecha"].strftime("%d/%m/%Y"),
            "Presupuesto": fmt_money(r["Costo"]),
            "Cómo leerlo": f"{r['p']}% de los escenarios termina en {fmt_num(r['Días'], 1)} días o menos",
        } for r in tp])
        objetivo = f"P{perc}"
        st.markdown(tabla_html(df_tp, resaltar=lambda fila: fila[0] == objetivo), unsafe_allow_html=True)

        # ---------------- Inspector y capacidad ----------------
        st.write("")
        col_i1, col_i2 = st.columns(2)
        with col_i1:
            seccion("Inspeccionar un escenario", "Revisa una iteración concreta del sorteo.")
            atajo = st.radio("Ir a", ["Elegir número", "Más rápido", "Mediano", "Más lento"], horizontal=True,
                             key="mc_atajo", label_visibility="collapsed")
            if atajo == "Elegir número":
                it = st.slider("Iteración", 1, len(dias), min(100, len(dias)), key="mc_slider_iter")
            else:
                orden = np.argsort(dias)
                it = int({"Más rápido": orden[0], "Mediano": orden[len(orden) // 2], "Más lento": orden[-1]}[atajo]) + 1
            d_it, c_it = float(dias[it - 1]), float(costos[it - 1])
            pct_rank = (dias <= d_it).mean() * 100
            st.markdown(caja_filas(f"Escenario #{fmt_num(it)} de {fmt_num(len(dias))}", [
                ("Duración", f"{fmt_num(d_it, 1)} días"),
                ("Costo", fmt_money(c_it)),
                ("Fecha de término", fecha_fin(fecha_inicio, d_it, habiles).strftime("%d/%m/%Y")),
                ("Posición", f"más lento que el {fmt_num(pct_rank, 0)}% de los escenarios"),
                ("Plazo", pill("Cumple", "ok") if d_it <= dias_meta else pill("Atrasado", "bad")),
                ("Presupuesto", pill("Dentro", "ok") if c_it <= costo_meta else pill("Sobrecosto", "bad")),
            ]), unsafe_allow_html=True)

        with col_i2:
            seccion("Capacidad anual del equipo", "Cuántos proyectos de este tamaño caben en un año.")
            dias_ano = 250 if habiles else 365
            cap50 = dias_ano / met["p50"] if met["p50"] > 0 else 0
            capt = dias_ano / met["pt"] if met["pt"] > 0 else 0
            st.markdown(caja_filas(f"Base: {dias_ano} {'días hábiles' if habiles else 'días corridos'} por año", [
                ("Capacidad mediana (P50)", f"{fmt_num(cap50, 1)} proyectos/año"),
                (f"Capacidad segura (P{perc})", f"{fmt_num(capt, 1)} proyectos/año"),
                ("Proyectos completos garantizados", f"{int(capt)} por año"),
                ("Riesgo asumido al planificar con P" + str(perc), f"{100 - perc}%"),
            ]), unsafe_allow_html=True)

        # ---------------- Escenarios extremos ----------------
        st.write("")
        df_sims = pd.DataFrame({
            "Iteración": np.arange(1, len(dias) + 1),
            "Duración (días)": np.round(dias, 1),
            "Costo ($)": np.round(costos, 0),
            "Cumple plazo": ok_t,
            "Cumple presupuesto": ok_c,
        })
        with st.expander("Ver escenarios extremos y datos crudos", expanded=False):
            def formatear_extremos(df):
                return pd.DataFrame({
                    "Iteración": df["Iteración"].map(fmt_num),
                    "Duración": df["Duración (días)"].map(lambda v: f"{fmt_num(v, 1)} d"),
                    "Costo": df["Costo ($)"].map(fmt_money),
                    "Plazo": df["Cumple plazo"].map(lambda v: pill("Cumple", "ok") if v else pill("Atraso", "bad")),
                    "Presupuesto": df["Cumple presupuesto"].map(lambda v: pill("Dentro", "ok") if v else pill("Excede", "bad")),
                })
            t1, t2 = st.columns(2)
            t1.markdown("**10 escenarios más rápidos**")
            t1.markdown(tabla_html(formatear_extremos(df_sims.nsmallest(10, "Duración (días)"))), unsafe_allow_html=True)
            t2.markdown("**10 escenarios más lentos**")
            t2.markdown(tabla_html(formatear_extremos(df_sims.nlargest(10, "Duración (días)"))), unsafe_allow_html=True)

        # ---------------- Exportación ----------------
        st.divider()
        seccion("Exportar la simulación",
                "El Excel incluye resumen ejecutivo, tabla de percentiles, histograma y curva S nativos, y todas las iteraciones.")
        e1, e2 = st.columns(2)
        with st.spinner("Preparando Excel…"):
            excel_mc = generar_excel_montecarlo(dias, costos, cfg)
        e1.download_button("📊 Reporte Monte Carlo en Excel (.xlsx)", excel_mc,
                           f"MonteCarlo_{slug(st.session_state.nombre_proj)}.xlsx", XLSX, **ANCHO)
        e2.download_button("📄 Iteraciones en CSV (.csv)", df_sims.to_csv(index=False).encode("utf-8"),
                           f"MonteCarlo_iteraciones_{slug(st.session_state.nombre_proj)}.csv", "text/csv", **ANCHO)

# ---------------------------------------------------------------------
# TAB 5: FLUJOS
# ---------------------------------------------------------------------
with tab5:
    seccion("Diagramas de flujo y secuencia", "Parte de una plantilla y edítala; la vista previa se actualiza sola.")
    c_f1, c_f2 = st.columns([1, 1.3])
    with c_f1:
        st.selectbox("Plantilla", options=list(PLANTILLAS_FLUJO.keys()), key="select_tipo_flujo",
                     on_change=cambiar_plantilla_flujo)
        if "txt_flujo_code" not in st.session_state:
            st.session_state.txt_flujo_code = st.session_state.flujo_codigo
        st.text_area("Código del diagrama", height=240, key="txt_flujo_code")
        st.session_state.flujo_codigo = st.session_state.txt_flujo_code
        st.download_button("📌 Descargar diagrama de flujo (.mmd)", st.session_state.flujo_codigo.encode("utf-8"),
                           f"Flujo_Secuencia_{slug(st.session_state.nombre_proj)}.mmd", "text/plain", **ANCHO)
    with c_f2:
        mostrar_mermaid(st.session_state.flujo_codigo, "400px", key=f"mm_flujo_{st.session_state.tema}")

# ---------------------------------------------------------------------
# TAB 6: EDT (WBS)
# ---------------------------------------------------------------------
with tab6:
    seccion("Estructura de desglose de trabajo (EDT / WBS)",
            "Proyecto → fases → subfases → paquetes de trabajo con horas, costo y documentos. "
            "Estos paquetes también pueden alimentar la simulación Monte Carlo.")

    with st.expander("Guía rápida de niveles", expanded=False):
        st.markdown(f"""
- 📦 **Proyecto (nivel 0):** *{st.session_state.nombre_proj}*
- 📂 **Fase principal (nivel 1):** p. ej. *Fase 1: Preparación*.
- 📁 **Subfase / grupo (nivel 2):** agrupación opcional (p. ej. *Aerosol*, *Documentación*).
- 📄 **Paquete de trabajo (nivel 3):** entregable con **horas**, **costo** y **documentos**.
        """)

    df_edt = pd.DataFrame(st.session_state.edt_list)
    if not df_edt.empty:
        total_horas, total_costo = df_edt["Horas"].sum(), df_edt["Costo"].sum()
        total_paquetes, total_fases = len(df_edt), df_edt["Fase"].nunique()
    else:
        total_horas = total_costo = total_paquetes = total_fases = 0

    m1, m2, m3, m4 = st.columns(4)
    kpi(m1, "Horas totales", f"{fmt_num(total_horas)} h", f"≈ {fmt_num(total_horas / 8, 1)} días de 8 h")
    kpi(m2, "Costo total estimado", fmt_money(total_costo), "", T["accent2"])
    kpi(m3, "Paquetes de trabajo", fmt_num(total_paquetes), "", T["sub"])
    kpi(m4, "Fases principales", fmt_num(total_fases), "", T["sub"])

    st.write("")
    seccion("Diagrama jerárquico")
    if not df_edt.empty:
        mostrar_mermaid(generar_mermaid_edt(st.session_state.edt_list, st.session_state.nombre_proj), "560px",
                        key=f"mm_edt_{st.session_state.tema}")
    else:
        st.info("Agrega un paquete de trabajo para ver el árbol jerárquico.")

    st.divider()
    col_graph, col_manage = st.columns([1.1, 1])
    with col_graph:
        seccion("Distribución de recursos por fase")
        if not df_edt.empty:
            df_fases = df_edt.groupby("Fase", as_index=False)[["Horas", "Costo"]].sum()
            tipo_metric = st.radio("Métrica", ["Horas", "Costo"], horizontal=True, key="rad_edt_metric")
            fig_donut = px.pie(df_fases, names="Fase", values=tipo_metric, hole=0.55,
                               color_discrete_sequence=[T["accent"], T["accent2"], T["warn"], T["ok"], T["bad"], T["sub"]])
            fig_donut.update_traces(textinfo="percent", hovertemplate="%{label}<br>%{value:,.0f}<extra></extra>")
            estilo_fig(fig_donut, height=360)
            mostrar_fig(fig_donut)
        else:
            st.info("Agrega datos para generar el gráfico de distribución.")

    with col_manage:
        seccion("Agregar paquete de trabajo")
        with st.form("form_add_edt_multi", clear_on_submit=True):
            fase_in = st.text_input("Fase principal", placeholder="Ej: Fase 1: Preparación")
            subfase_in = st.text_input("Subfase / grupo (opcional)", placeholder="Ej: Pintura aerosol")
            paquete_in = st.text_input("Paquete de trabajo / tarea", placeholder="Ej: Pintar arbusto")
            c_h, c_c = st.columns(2)
            horas_in = c_h.number_input("Horas estimadas", min_value=0, value=4, step=1)
            costo_in = c_c.number_input("Costo estimado ($)", min_value=0, value=15000, step=1000)
            doc_in = st.text_input("Documentos / referencias", placeholder="Ej: Ficha técnica aerosol")
            if st.form_submit_button("Añadir paquete a la EDT", **ANCHO):
                if fase_in.strip() and paquete_in.strip():
                    st.session_state.edt_list.append({
                        "Fase": fase_in.strip(), "Subfase": subfase_in.strip(), "Paquete": paquete_in.strip(),
                        "Horas": float(horas_in), "Costo": float(costo_in), "Documentos": doc_in.strip()})
                    st.rerun()
                else:
                    st.error("Completa al menos la fase principal y el paquete de trabajo.")

    st.divider()
    seccion("Editar paquetes")
    if not st.session_state.edt_list:
        st.info("La EDT está vacía. Usa el formulario de arriba para agregar el primer paquete.")
    else:
        anchos = [2, 1.8, 2.2, 1, 1.2, 1.8, 0.5]
        cab = st.columns(anchos)
        for c, t in zip(cab, ["Fase", "Subfase", "Paquete / tarea", "Horas", "Costo ($)", "Documentos", ""]):
            c.markdown(f"**{t}**")
        borrar = None
        for i, item in enumerate(st.session_state.edt_list):
            cs = st.columns(anchos)
            item["Fase"] = cs[0].text_input(f"f_{i}", value=item.get("Fase", ""), key=f"edt_fase_{i}", label_visibility="collapsed")
            item["Subfase"] = cs[1].text_input(f"sf_{i}", value=item.get("Subfase", ""), key=f"edt_sfase_{i}", label_visibility="collapsed")
            item["Paquete"] = cs[2].text_input(f"p_{i}", value=item.get("Paquete", ""), key=f"edt_pkg_{i}", label_visibility="collapsed")
            item["Horas"] = cs[3].number_input(f"h_{i}", value=float(item.get("Horas", 0)), min_value=0.0, step=1.0,
                                               key=f"edt_hrs_{i}", label_visibility="collapsed")
            item["Costo"] = cs[4].number_input(f"c_{i}", value=float(item.get("Costo", 0)), min_value=0.0, step=1000.0,
                                               key=f"edt_cost_{i}", label_visibility="collapsed")
            item["Documentos"] = cs[5].text_input(f"d_{i}", value=item.get("Documentos", ""), key=f"edt_doc_{i}", label_visibility="collapsed")
            if cs[6].button("🗑️", key=f"del_edt_{i}", help="Eliminar paquete"):
                borrar = i
        if borrar is not None:
            st.session_state.edt_list.pop(borrar)
            limpiar_widgets("edt_fase_", "edt_sfase_", "edt_pkg_", "edt_hrs_", "edt_cost_", "edt_doc_")
            st.rerun()

    st.divider()
    seccion("Exportar EDT")
    c_dedt1, c_dedt2 = st.columns(2)
    c_dedt1.download_button("📊 EDT en Excel (.xlsx)", generar_excel_edt(),
                            f"EDT_{slug(st.session_state.nombre_proj)}.xlsx", XLSX, **ANCHO)
    c_dedt2.download_button("📄 EDT en CSV (.csv)", pd.DataFrame(st.session_state.edt_list).to_csv(index=False).encode("utf-8"),
                            f"EDT_{slug(st.session_state.nombre_proj)}.csv", "text/csv", **ANCHO)

# ---------------------------------------------------------------------
# TAB 7: RIESGOS
# ---------------------------------------------------------------------
with tab7:
    seccion("Matriz de riesgos 5×5", "Severidad = probabilidad × impacto. Los riesgos sobre 15 son críticos.")

    df_r = pd.DataFrame(st.session_state.riesgos_list)
    if not df_r.empty:
        df_r["Severidad"] = df_r["Probabilidad"] * df_r["Impacto"]
        df_r["Nivel"] = df_r["Severidad"].apply(clasificar_riesgo)
    else:
        df_r = pd.DataFrame(columns=["Riesgo", "Probabilidad", "Impacto", "Severidad", "Nivel"])

    r_total = len(df_r)
    r_criticos = int((df_r["Severidad"] >= 12).sum()) if not df_r.empty else 0
    r_prom = float(df_r["Severidad"].mean()) if not df_r.empty else 0.0
    mr1, mr2, mr3 = st.columns(3)
    kpi(mr1, "Riesgos registrados", fmt_num(r_total), "", T["sub"])
    kpi(mr2, "Riesgos altos o críticos", fmt_num(r_criticos), "Severidad ≥ 12", T["bad"])
    kpi(mr3, "Severidad promedio", f"{fmt_num(r_prom, 1)} / 25", "", T["warn"])

    st.write("")
    col_matriz, col_form = st.columns([1.3, 1])
    with col_matriz:
        seccion("Mapa de calor")
        z_grid = [[i * j for i in range(1, 6)] for j in range(1, 6)]
        fig_matrix = go.Figure(go.Heatmap(
            z=z_grid, x=[1, 2, 3, 4, 5], y=[1, 2, 3, 4, 5], showscale=False, hoverinfo="skip",
            text=z_grid, texttemplate="%{text}", textfont=dict(color=rgba(T["text"], 0.35), size=11),
            colorscale=[[0.0, rgba(T["ok"], 0.30)], [0.35, rgba(T["warn"], 0.30)],
                        [0.6, "rgba(245,124,0,0.35)"], [1.0, rgba(T["bad"], 0.45)]]))
        if not df_r.empty:
            fig_matrix.add_trace(go.Scatter(
                x=df_r["Probabilidad"], y=df_r["Impacto"], mode="markers+text", text=df_r["Riesgo"],
                textposition="top center", textfont=dict(color=T["text"], size=11),
                marker=dict(size=16, color=T["text"], line=dict(width=2, color=T["card"])),
                customdata=df_r["Severidad"],
                hovertemplate="<b>%{text}</b><br>Probabilidad: %{x}<br>Impacto: %{y}<br>Severidad: %{customdata}<extra></extra>"))
        estilo_fig(fig_matrix, height=420, showlegend=False)
        fig_matrix.update_xaxes(title="Probabilidad (1-5)", range=[0.5, 5.5], dtick=1, showgrid=False)
        fig_matrix.update_yaxes(title="Impacto (1-5)", range=[0.5, 5.5], dtick=1, showgrid=False)
        mostrar_fig(fig_matrix)

    with col_form:
        seccion("Registrar riesgo")
        with st.form("form_add_riesgo", clear_on_submit=True):
            r_nombre = st.text_input("Descripción del riesgo", placeholder="Ej: Retraso en entregables")
            c_p, c_i = st.columns(2)
            r_prob = c_p.slider("Probabilidad", 1, 5, 3)
            r_imp = c_i.slider("Impacto", 1, 5, 3)
            if st.form_submit_button("Agregar riesgo", **ANCHO):
                if r_nombre.strip():
                    st.session_state.riesgos_list.append({"Riesgo": r_nombre.strip(), "Probabilidad": int(r_prob), "Impacto": int(r_imp)})
                    st.rerun()
                else:
                    st.error("Escribe una descripción para el riesgo.")

    st.divider()
    seccion("Inventario de riesgos")
    if not st.session_state.riesgos_list:
        st.info("No hay riesgos. Registra el primero con el formulario de arriba.")
    else:
        anchos_r = [3, 1.2, 1.2, 1, 1.4, 0.5]
        cab = st.columns(anchos_r)
        for c, t in zip(cab, ["Riesgo", "Probabilidad", "Impacto", "Severidad", "Nivel", ""]):
            c.markdown(f"**{t}**")
        borrar = None
        for idx, item in enumerate(st.session_state.riesgos_list):
            rc = st.columns(anchos_r)
            item["Riesgo"] = rc[0].text_input(f"rt_{idx}", value=item["Riesgo"], key=f"r_name_{idx}", label_visibility="collapsed")
            item["Probabilidad"] = rc[1].number_input(f"rp_{idx}", value=int(item["Probabilidad"]), min_value=1, max_value=5,
                                                      key=f"r_p_{idx}", label_visibility="collapsed")
            item["Impacto"] = rc[2].number_input(f"ri_{idx}", value=int(item["Impacto"]), min_value=1, max_value=5,
                                                 key=f"r_i_{idx}", label_visibility="collapsed")
            sev = item["Probabilidad"] * item["Impacto"]
            rc[3].markdown(f"**{sev}**")
            rc[4].markdown(clasificar_riesgo(sev))
            if rc[5].button("🗑️", key=f"del_r_{idx}", help="Eliminar riesgo"):
                borrar = idx
        if borrar is not None:
            st.session_state.riesgos_list.pop(borrar)
            limpiar_widgets("r_name_", "r_p_", "r_i_")
            st.rerun()

    st.divider()
    seccion("Exportar matriz")
    c_dr1, _ = st.columns(2)
    c_dr1.download_button("📊 Matriz de riesgos en Excel (.xlsx)", generar_excel_riesgos(),
                          f"Matriz_Riesgos_{slug(st.session_state.nombre_proj)}.xlsx", XLSX, **ANCHO)

# ---------------------------------------------------------------------
# TAB 8: REPORTE MEMENTO MORI
# ---------------------------------------------------------------------
with tab8:
    seccion("Reporte Memento Mori",
            "Documento Word con el análisis crítico de todos los componentes del proyecto. "
            "Incluye los resultados reales de la simulación Monte Carlo configurada.")
    col_m1, col_m2 = st.columns(2)
    st.session_state.integrantes = col_m1.text_input("Integrantes / responsables", st.session_state.integrantes, key="mem_tab_integrantes")
    st.session_state.nombre_proj = col_m2.text_input("Nombre del proyecto", st.session_state.nombre_proj, key="mem_tab_nombre")

    st.markdown("""
**Contenido del reporte**
1. **Introducción y alcance:** viabilidad, alineación estratégica y prevención de scope creep.
2. **Requisitos (RF / RNF):** vulnerabilidad en los requisitos no funcionales.
3. **Arquitectura y flujos:** acoplamiento de componentes y puntos de fallo.
4. **Gestión Kanban:** trabajo en proceso (WIP) y cuellos de botella.
5. **EDT (WBS):** actividades críticas por costo y horas.
6. **Simulación Monte Carlo:** niveles de confianza, reservas de contingencia y veredicto.
7. **Matriz de riesgos:** planificación anticipada ante fallos y mitigación.
8. **Conclusiones y recomendaciones.**
    """)
    if MC_ACTUAL is None:
        st.warning("La simulación Monte Carlo tiene parámetros inválidos; el reporte lo indicará en la sección 6.")

    col_btn_center, _ = st.columns([2, 1])
    col_btn_center.download_button("📄 Descargar reporte Memento Mori (.docx)", generar_word_mementomori(),
                                   f"Reporte_MementoMori_{slug(st.session_state.nombre_proj)}.docx", DOCX, **ANCHO)

# ---------------------------------------------------------------------
# TAB 9: DESCARGA COMPLETA
# ---------------------------------------------------------------------
with tab_full:
    seccion("Exportación consolidada",
            "Un archivo ZIP con todos los entregables: reporte Memento Mori, ERS, matrices en Excel, "
            "reporte Monte Carlo y CSVs.")
    nombre_s = slug(st.session_state.nombre_proj)
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"Reporte_MementoMori_{nombre_s}.docx", generar_word_mementomori())
        zf.writestr(f"ERS_{nombre_s}.docx", generar_word_ers())
        zf.writestr(f"Matrices_{nombre_s}.xlsx", generar_excel_estilizado())
        if MC_ACTUAL:
            zf.writestr(f"MonteCarlo_{nombre_s}.xlsx",
                        generar_excel_montecarlo(MC_ACTUAL["res"]["dias"], MC_ACTUAL["res"]["costos"], MC_ACTUAL["cfg"]))
        zf.writestr("Kanban_Tareas.csv", pd.DataFrame(st.session_state.kanban_tasks).to_csv(index=False).encode("utf-8"))
        zf.writestr("EDT_WBS.csv", pd.DataFrame(st.session_state.edt_list).to_csv(index=False).encode("utf-8"))
        zf.writestr("Mermaid/Modelado.mmd", st.session_state.modelo_mermaid)
        zf.writestr("Mermaid/Flujo.mmd", st.session_state.flujo_codigo)

    col1, col2, col3 = st.columns([1, 2, 1])
    col2.download_button("⬇ Descargar paquete completo (.zip)", zip_buffer.getvalue(),
                         f"Reporte_Full_{nombre_s}.zip", "application/zip", **ANCHO)

# =====================================================================
# 7. BARRA LATERAL (tema y resumen)
# =====================================================================
with st.sidebar:
    st.markdown("### MatrixDev")
    st.caption("Cambia el tema con los iconos ☀️ 🌙 💻 de la esquina superior.")
    s_cols = st.columns(3)
    for s_col, (clave, datos) in zip(s_cols, TEMAS.items()):
        s_col.button(datos["icono"], key=f"tema_btn_side_{clave}", help=f"Tema {datos['nombre'].lower()}",
                     on_click=cambiar_tema, args=(clave,), **ANCHO)
    st.divider()
    st.markdown("**Resumen del proyecto**")
    resumen = [
        ("Requisitos", f"{len(st.session_state.rf_list)} RF · {len(st.session_state.rnf_list)} RNF"),
        ("Tareas", f"{sum(t['Estado'] == 'Completado' for t in st.session_state.kanban_tasks)}/{len(st.session_state.kanban_tasks)} completadas"),
        ("EDT", f"{fmt_num(sum(float(e.get('Horas', 0) or 0) for e in st.session_state.edt_list))} h · "
                f"{fmt_money(sum(float(e.get('Costo', 0) or 0) for e in st.session_state.edt_list))}"),
        ("Riesgos", f"{len(st.session_state.riesgos_list)} registrados"),
    ]
    if MC_ACTUAL:
        m_ = MC_ACTUAL["met"]
        resumen.append((f"Plazo P{MC_ACTUAL['cfg']['perc']}", f"{fmt_num(m_['pt'], 1)} días"))
        resumen.append(("Éxito conjunto", f"{fmt_num(m_['prob_ambos'], 1)}%"))
    st.markdown(caja_filas("Indicadores", resumen), unsafe_allow_html=True)
