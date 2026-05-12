"""
Sistema 5S INOVA - Version PRO Ejecutiva
Archivo: appi.py

Incluye lo solicitado en el Word:
- Programar auditoria 5S.
- Meta por bodega >= 90%.
- Promedio general >= 90%.
- Responsables base: Darwin Herrera, Nelson Meza, Carlos Lugo, Aldair Montes.
- Opcion de agregar nuevos responsables.
- Cronograma visual tipo Gantt con eje X por fecha y etiqueta de responsable.
- Checklist por bodega segun requerimientos.
- Inspeccion con evidencia fotografica multiple.
- Dashboard ejecutivo con KPIs, ranking, tendencia, semaforos y hallazgos.
- Exportacion Excel, PDF y grafica del cronograma.
"""

import os
import io
import json
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    PageBreak,
)

# =========================================================
# CONFIGURACION GENERAL
# =========================================================
LOGO_INOVA = "INOVA.png"
LOGO_EASY = "FOND EASY.png"

DATA_DIR = "data_5s"
REPORTS_DIR = os.path.join(DATA_DIR, "reportes")
EVIDENCE_DIR = os.path.join(DATA_DIR, "evidencias")
DB_PATH = os.path.join(DATA_DIR, "inspecciones.json")
SCHEDULE_PATH = os.path.join(DATA_DIR, "cronograma.json")
RESPONSIBLES_PATH = os.path.join(DATA_DIR, "responsables.json")
EXCEL_PATH = os.path.join(DATA_DIR, "historico_5s.xlsx")

for folder in [DATA_DIR, REPORTS_DIR, EVIDENCE_DIR]:
    os.makedirs(folder, exist_ok=True)

if os.path.exists(LOGO_INOVA):
    try:
        logo_icon = Image.open(LOGO_INOVA)
    except Exception:
        logo_icon = "📦"
else:
    logo_icon = "📦"

st.set_page_config(
    page_title="5S INOVA PRO",
    page_icon=logo_icon,
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# CONSTANTES DE NEGOCIO
# =========================================================
META_BODEGA = 90.0
META_GENERAL = 90.0

USUARIOS_SISTEMA = {
    "DHERRERA": "1397",
    "GVISBAL": "0768",
}

RESPONSABLES_DEFAULT = [
    {
        "id": "darwin_herrera",
        "nombre": "Darwin Herrera",
        "cargo": "Responsable 5S",
        "area": "Operaciones",
        "color": "#156CC1",
        "activo": True,
    },
    {
        "id": "nelson_meza",
        "nombre": "Nelson Meza",
        "cargo": "Responsable 5S",
        "area": "Operaciones",
        "color": "#13A35B",
        "activo": True,
    },
    {
        "id": "carlos_lugo",
        "nombre": "Carlos Lugo",
        "cargo": "Responsable 5S",
        "area": "Operaciones",
        "color": "#FF8A00",
        "activo": True,
    },
    {
        "id": "aldair_montes",
        "nombre": "Aldair Montes",
        "cargo": "Responsable 5S",
        "area": "Operaciones",
        "color": "#7C3AED",
        "activo": True,
    },
]

BODEGA_COLORS = {
    "Bodega General": "#156CC1",
    "Bodega Tierras": "#7EC0EE",
    "Bodega Preforma": "#76E09B",
    "Bodega Químico": "#E24B4B",
    "Bodega Cuarto Frío": "#5DADEC",
    "Bodega Cuarto Atemperado": "#2BB3A3",
}

# Checklists cargados desde el Word del usuario.
BODEGAS = {
    "Bodega General": [
        "Limpieza pisos pasillos (pasillo 1-2, zonas de tránsito)",
        "Limpieza pisos naves de almacenamiento (Lata y zona de alistamiento)",
        "Limpieza pisos zona de revisión y ETO",
        "Limpieza pisos oficina administrativa",
        "Limpieza pisos muelles de descargue",
        "Limpieza patio 1, línea de vida y muelles",
        "Limpieza piso zona de parqueo equipos eléctricos",
        "Limpieza estaciones de aseo (Bodega general y patio)",
        "Limpieza de estanterías pasillo 1-2 y estantería azúcar (polvo, suciedad, telarañas, etc.)",
        "Limpieza de materiales (cambio de pelex y retiro de polvo)",
        "Almacenamiento de materiales cumpliendo layout",
        "Limpieza herramientas manuales y ubicados en su layout",
        "Orden y aseo gabinetes (gabinete implementos 5S y gabinete oficina administrativa)",
        "Validación estibas en buen estado (0 estibas partidas)",
    ],
    "Bodega Tierras": [
        "Limpieza pisos zona de tránsito y muelle de descargue",
        "Limpieza pisos naves de almacenamiento",
        "Limpieza de estanterías (polvo, suciedad, telarañas, etc.)",
        "Limpieza de materiales (cambio de pelex y retiro de polvo)",
        "Limpieza herramientas 5S y portón muelle descargue",
        "Almacenamiento de materiales cumpliendo layout",
        "Limpieza herramientas manuales y ubicados en su layout",
        "Muelle de descargue libre de estibas",
        "Limpieza rampa zona externa de la bodega",
        "Cumplimiento patrón de estibado de materiales",
        "Validación estibas en buen estado (0 estibas partidas)",
    ],
    "Bodega Preforma": [
        "Limpieza pisos zona de tránsito y muelle de descargue",
        "Limpieza pisos naves de almacenamiento",
        "Limpieza de estanterías (polvo, suciedad, telarañas, etc.)",
        "Limpieza de materiales (cambio de pelex y retiro de polvo)",
        "Limpieza muelle de preforma y zona detrás buffer de lata",
        "Almacenamiento de materiales cumpliendo layout",
        "Bodega y muelle de descargue libre de estibas",
        "Validación estibas en buen estado (0 estibas partidas)",
    ],
    "Bodega Químico": [
        "Limpieza de pisos pasillo externo bodega y muelle de descargue",
        "Limpieza de rampa",
        "Limpieza de pisos pasillos (pasillo 1-2-3)",
        "Limpieza de estanterías (polvo, suciedad, telarañas, etc.)",
        "Limpieza de materiales (cambio de pelex y retiro de polvo)",
        "Limpieza gabinete EPPS",
        "Limpieza herramientas manuales y ubicados en su layout",
        "Almacenamiento de materiales cumpliendo layout",
        "Validación estibas en buen estado (0 estibas partidas)",
        "Cumplimiento compatibilidad SQ almacenamiento",
        "Limpieza de estibas plásticas rojas fuera de bodega",
    ],
    "Bodega Cuarto Frío": [
        "Limpiezas pisos pasillo externo bodega",
        "Limpieza zona de almacenamiento color caramelo",
        "Limpieza pisos bodega",
        "Limpieza de estanterías (polvo, suciedad, telarañas, etc.)",
        "Limpieza de materiales (cambio de pelex y retiro de polvo)",
        "Almacenamiento de materiales cumpliendo layout",
        "Validación estibas en buen estado (0 estibas partidas)",
        "Validación de control temperatura dentro de parámetro",
    ],
    "Bodega Cuarto Atemperado": [
        "Limpiezas pisos pasillo externo bodega",
        "Limpieza pisos bodega",
        "Limpieza de estanterías (polvo, suciedad, telarañas, etc.)",
        "Limpieza de materiales (cambio de pelex y retiro de polvo)",
        "Almacenamiento de materiales cumpliendo layout",
        "Validación estibas en buen estado (0 estibas partidas)",
        "Validación de control temperatura dentro de parámetro",
    ],
}

ESTADOS_CRONOGRAMA = ["Programada", "En ejecución", "Finalizada", "Vencida", "Crítica"]

# Paleta visual viva alineada con la interfaz INOVA
PALETA_VIVA = [
    "#156CC1",  # azul INOVA
    "#00B8D9",  # cyan vivo
    "#13A35B",  # verde
    "#FF8A00",  # naranja
    "#7C3AED",  # violeta
    "#D53333",  # rojo
    "#2BB3A3",  # teal
    "#F5C542",  # amarillo
]

ESCALA_CUMPLIMIENTO_VIVA = [
    [0.00, "#D53333"],
    [0.55, "#FF8A00"],
    [0.75, "#F5C542"],
    [0.90, "#13A35B"],
    [1.00, "#156CC1"],
]


COLOR_ESTADO_VIVO = {
    "Programada": "#156CC1",
    "En ejecución": "#FF8A00",
    "Finalizada": "#13A35B",
    "Vencida": "#667085",
    "Crítica": "#D53333",
}



# =========================================================
# CONFIGURACION VISUAL PERSONALIZABLE
# =========================================================
if "color_primario" not in st.session_state:
    st.session_state.color_primario = "#156CC1"
if "color_secundario" not in st.session_state:
    st.session_state.color_secundario = "#13A35B"
if "color_alerta" not in st.session_state:
    st.session_state.color_alerta = "#D53333"

if st.session_state.get("autenticado", False):
    with st.expander("🎨 Personalizar colores del dashboard", expanded=False):
        c1, c2, c3 = st.columns(3)

        with c1:
            st.session_state.color_primario = st.color_picker(
                "Color principal",
                st.session_state.color_primario
            )

        with c2:
            st.session_state.color_secundario = st.color_picker(
                "Color secundario",
                st.session_state.color_secundario
            )

        with c3:
            st.session_state.color_alerta = st.color_picker(
                "Color alerta/meta",
                st.session_state.color_alerta
            )

PALETA_VIVA = [
    st.session_state.color_primario,
    st.session_state.color_secundario,
    "#00B8D9",
    "#FF8A00",
    "#7C3AED",
    st.session_state.color_alerta,
    "#2BB3A3",
    "#F5C542",
]

ESCALA_CUMPLIMIENTO_VIVA = [
    [0.00, st.session_state.color_alerta],
    [0.55, "#FF8A00"],
    [0.75, "#F5C542"],
    [0.90, st.session_state.color_secundario],
    [1.00, st.session_state.color_primario],
]

# =========================================================
# CSS NIVEL PRO
# =========================================================
st.markdown(
    """
<style>
:root{
    --azul:#061f45;
    --azul2:#082b5c;
    --azul3:#0d3b73;
    --fondo:#edf3f9;
    --panel:#ffffff;
    --borde:#d7e3ef;
    --texto:#1f2937;
    --muted:#667085;
    --verde:#13a35b;
    --amarillo:#d99b00;
    --rojo:#d53333;
    --shadow:0 18px 42px rgba(8,32,68,0.08);
}

html, body, [class*="css"] {
    font-family: "Segoe UI", Inter, Arial, sans-serif;
}

.stApp{
    background:
        radial-gradient(circle at 12% 10%, rgba(21,108,193,0.14), transparent 24%),
        radial-gradient(circle at 90% 0%, rgba(6,31,69,0.12), transparent 22%),
        linear-gradient(180deg, #f6f9fc 0%, #edf3f9 100%);
}

.block-container{
    max-width: 100% !important;
    padding: 0.75rem 1.25rem 2rem 1.25rem;
}

section[data-testid="stSidebar"]{
    background: linear-gradient(180deg, #061f45 0%, #082b5c 100%);
    border-right: 1px solid rgba(255,255,255,0.08);
}

section[data-testid="stSidebar"] *{
    color: white !important;
}

section[data-testid="stSidebar"] .stButton > button{
    background: rgba(255,255,255,0.10) !important;
    color: white !important;
    border: 1px solid rgba(255,255,255,0.18) !important;
}

[data-testid="stSidebar"] input,
[data-testid="stSidebar"] select,
[data-testid="stSidebar"] textarea{
    color:#061f45 !important;
}

.top-shell{
    background: rgba(255,255,255,0.86);
    border: 1px solid rgba(215,227,239,0.95);
    border-radius: 28px;
    padding: 16px 18px;
    box-shadow: var(--shadow);
    margin-bottom: 16px;
    backdrop-filter: blur(10px);
}

.top-banner{
    position: relative;
    overflow: hidden;
    background:
        radial-gradient(circle at 95% 10%, rgba(255,255,255,0.20), transparent 18%),
        linear-gradient(135deg, #061f45 0%, #0d3b73 54%, #156cc1 100%);
    color: white;
    border-radius: 24px;
    padding: 22px 28px;
    min-height: 126px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 18px 36px rgba(6,31,69,0.28);
}

.top-title{
    font-size: 2.15rem;
    font-weight: 900;
    line-height: 1.05;
    letter-spacing: -0.04em;
    margin-bottom: 7px;
}

.top-subtitle{
    font-size: 0.98rem;
    color: #dce8f8;
    max-width: 900px;
}

.header-badge{
    padding: 10px 14px;
    border-radius: 999px;
    background: rgba(255,255,255,0.14);
    border: 1px solid rgba(255,255,255,0.20);
    color: white;
    font-weight: 800;
    white-space: nowrap;
}

.main-wrap{
    background: rgba(255,255,255,0.86);
    backdrop-filter: blur(10px);
    border: 1px solid rgba(215,227,239,0.98);
    border-radius: 28px;
    padding: 22px;
    box-shadow: var(--shadow);
    margin-bottom: 18px;
}

.section-title{
    color: #061f45;
    font-size: 1.38rem;
    font-weight: 900;
    letter-spacing: -0.02em;
    margin-bottom: 4px;
}

.section-subtitle{
    color:#667085;
    margin-bottom: 18px;
    font-size: 0.96rem;
}

.kpi-card{
    background:
        radial-gradient(circle at top right, rgba(21,108,193,0.08), transparent 30%),
        linear-gradient(180deg, #ffffff 0%, #f7fbff 100%);
    border: 1px solid #dbe6f1;
    border-radius: 22px;
    padding: 20px;
    min-height: 124px;
    box-shadow: 0 12px 30px rgba(8,32,68,0.07);
}

.kpi-label{
    color: #667085;
    font-size: 0.78rem;
    margin-bottom: 8px;
    font-weight: 900;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}

.kpi-value{
    color: #061f45;
    font-size: 2.15rem;
    font-weight: 900;
    line-height: 1;
}

.kpi-note{
    margin-top: 8px;
    color:#667085;
    font-size:0.88rem;
}

.exec-card{
    background:white;
    border:1px solid #dbe6f1;
    border-radius:22px;
    padding:18px;
    box-shadow:0 10px 26px rgba(8,32,68,0.06);
    margin-bottom:14px;
}

.ins-card{
    background: white;
    border: 1px solid #d9e6f2;
    border-radius: 20px;
    padding: 16px 16px 10px 16px;
    box-shadow: 0 8px 22px rgba(8, 32, 68, 0.055);
    margin-bottom: 14px;
}

.punto-title{
    font-size: 1.02rem;
    font-weight: 900;
    color: #082b5c;
    margin-bottom: 8px;
}

.punto-sub{
    color: #344054;
    font-size: 0.95rem;
    line-height: 1.38;
}

.badge-green, .badge-yellow, .badge-red, .badge-blue, .badge-gray{
    display:inline-flex;
    align-items:center;
    gap:6px;
    padding:7px 13px;
    border-radius:999px;
    font-weight:900;
    font-size:0.86rem;
    letter-spacing:0.02em;
}
.badge-green{background:#e9f8ef;color:#118a4c;border:1px solid #bfe9ce;}
.badge-yellow{background:#fff7df;color:#9b7400;border:1px solid #ecd995;}
.badge-red{background:#fdecec;color:#bf2525;border:1px solid #efbbbb;}
.badge-blue{background:#eef5ff;color:#155bb5;border:1px solid #c9dcff;}
.badge-gray{background:#eef2f6;color:#475467;border:1px solid #d0d5dd;}

.stButton > button,
div[data-testid="stDownloadButton"] > button,
div[data-testid="stFormSubmitButton"] > button{
    border-radius: 15px !important;
    min-height: 46px !important;
    font-weight: 900 !important;
    border: 1px solid #c8d8e8 !important;
    background: linear-gradient(180deg, #ffffff 0%, #f4f8fd 100%) !important;
    color: #082b5c !important;
}

button[kind="primary"], .stButton > button[kind="primary"]{
    background: linear-gradient(135deg, #156cc1 0%, #0b4fc4 100%) !important;
    color: white !important;
    border: none !important;
}

[data-testid="stMetricValue"]{
    color:#061f45;
    font-weight:900;
}

hr{
    border: none;
    border-top: 1px solid #e3edf6;
    margin: 16px 0;
}
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# HELPERS
# =========================================================
def safe_load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def safe_save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def slugify(value):
    text = str(value or "").strip().lower()
    replacements = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ñ": "n",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return "_".join(text.split()) or datetime.now().strftime("item_%Y%m%d_%H%M%S")


def get_responsables_activos():
    return [r for r in st.session_state.responsables if r.get("activo", True)]


def get_responsable_color(nombre):
    for r in st.session_state.responsables:
        if r.get("nombre") == nombre:
            return r.get("color", "#156CC1")
    return "#156CC1"


def get_week_label(fecha_ref=None):
    if fecha_ref is None:
        fecha_ref = datetime.today()
    year, week_num, _ = fecha_ref.isocalendar()
    return f"Semana_{week_num:02d}_{year}"


def cumplimiento_texto(valor):
    if valor >= META_BODEGA:
        return "Excelente"
    if valor >= 75:
        return "Atención"
    return "Crítico"


def cumplimiento_badge(valor):
    if valor >= META_BODEGA:
        return '<span class="badge-green">🟢 Excelente</span>'
    if valor >= 75:
        return '<span class="badge-yellow">🟡 Atención</span>'
    return '<span class="badge-red">🔴 Crítico</span>'


def estado_badge(estado):
    if estado == "Finalizada":
        return '<span class="badge-green">Finalizada</span>'
    if estado == "En ejecución":
        return '<span class="badge-yellow">En ejecución</span>'
    if estado == "Crítica":
        return '<span class="badge-red">Crítica</span>'
    if estado == "Vencida":
        return '<span class="badge-gray">Vencida</span>'
    return '<span class="badge-blue">Programada</span>'


def infer_schedule_status(item):
    status = item.get("estado", "Programada")
    try:
        fecha_fin = pd.to_datetime(item.get("fecha_fin")).date()
        if status == "Programada" and fecha_fin < date.today():
            return "Vencida"
    except Exception:
        pass
    return status


def save_uploaded_image(uploaded_file, folder, file_prefix):
    if uploaded_file is None:
        return None
    os.makedirs(folder, exist_ok=True)
    ext = uploaded_file.name.split(".")[-1].lower()
    filename = f"{file_prefix}.{ext}"
    path = os.path.join(folder, filename)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path


def resize_image(path, max_width=1200):
    try:
        img = Image.open(path)
        w, h = img.size
        if w > max_width:
            ratio = max_width / w
            img = img.resize((int(w * ratio), int(h * ratio)))
            img.save(path)
    except Exception:
        pass


def fit_image_box(path, max_w_px=900, max_h_px=520):
    try:
        img = Image.open(path)
        w, h = img.size
        ratio = min(max_w_px / w, max_h_px / h)
        ratio = min(ratio, 1.0)
        return int(w * ratio), int(h * ratio)
    except Exception:
        return 800, 450


def normalizar_items_legacy(items):
    normalizados = []
    for item in items or []:
        if "fotos" not in item:
            foto_unica = item.get("foto")
            item["fotos"] = [foto_unica] if foto_unica else []
        if "severidad" not in item:
            item["severidad"] = "Media" if not item.get("cumple") else "Sin novedad"
        normalizados.append(item)
    return normalizados


def append_to_excel(registro):
    rows = []
    for item in registro["items"]:
        fotos = item.get("fotos", [])
        rows.append(
            {
                "Fecha": registro["fecha"],
                "Semana": registro.get("semana", ""),
                "Responsable": registro["responsable"],
                "Bodega": registro["bodega"],
                "Area": registro["area"],
                "Punto": item["punto"],
                "Cumple": "Si" if item["cumple"] else "No",
                "Severidad": item.get("severidad", ""),
                "Observacion": item.get("observacion", ""),
                "Fotos": " | ".join(fotos) if fotos else "",
                "Cantidad fotos": len(fotos),
                "Cumplimiento Total %": registro["cumplimiento"],
                "Meta Bodega %": META_BODEGA,
                "Estado": cumplimiento_texto(registro["cumplimiento"]),
            }
        )

    df_new = pd.DataFrame(rows)
    if os.path.exists(EXCEL_PATH):
        try:
            df_old = pd.read_excel(EXCEL_PATH)
            df_all = pd.concat([df_old, df_new], ignore_index=True)
        except Exception:
            df_all = df_new
    else:
        df_all = df_new

    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        df_all.to_excel(writer, index=False, sheet_name="Historico")
    return EXCEL_PATH


def rebuild_excel_from_inspections(inspecciones):
    rows = []
    for registro in inspecciones:
        for item in normalizar_items_legacy(registro.get("items", [])):
            rows.append(
                {
                    "Fecha": registro.get("fecha"),
                    "Semana": registro.get("semana", ""),
                    "Responsable": registro.get("responsable"),
                    "Bodega": registro.get("bodega"),
                    "Area": registro.get("area"),
                    "Punto": item.get("punto"),
                    "Cumple": "Si" if item.get("cumple") else "No",
                    "Severidad": item.get("severidad", ""),
                    "Observacion": item.get("observacion", ""),
                    "Fotos": " | ".join(item.get("fotos", [])),
                    "Cantidad fotos": len(item.get("fotos", [])),
                    "Cumplimiento Total %": registro.get("cumplimiento", 0),
                    "Meta Bodega %": META_BODEGA,
                    "Estado": cumplimiento_texto(float(registro.get("cumplimiento", 0))),
                }
            )
    if rows:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            pd.DataFrame(rows).to_excel(writer, index=False, sheet_name="Historico")
    elif os.path.exists(EXCEL_PATH):
        os.remove(EXCEL_PATH)


def export_dataframe_excel(df, sheet_name="Datos"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer


def build_inspection_dataframe(inspecciones):
    rows = []
    for reg in inspecciones:
        items = normalizar_items_legacy(reg.get("items", []))
        total = len(items)
        no_conformes = sum(1 for x in items if not x.get("cumple"))
        observaciones = sum(1 for x in items if str(x.get("observacion", "")).strip())
        rows.append(
            {
                "Fecha": pd.to_datetime(reg.get("fecha")),
                "Responsable": reg.get("responsable", ""),
                "Bodega": reg.get("bodega", ""),
                "Cumplimiento": float(reg.get("cumplimiento", 0)),
                "Meta": META_BODEGA,
                "Estado": cumplimiento_texto(float(reg.get("cumplimiento", 0))),
                "Puntos evaluados": total,
                "No conformes": no_conformes,
                "Observaciones": observaciones,
            }
        )
    return pd.DataFrame(rows)


def build_items_dataframe(inspecciones):
    rows = []
    for reg in inspecciones:
        for item in normalizar_items_legacy(reg.get("items", [])):
            rows.append(
                {
                    "Fecha": pd.to_datetime(reg.get("fecha")),
                    "Responsable": reg.get("responsable", ""),
                    "Bodega": reg.get("bodega", ""),
                    "Punto": item.get("punto", ""),
                    "Cumple": bool(item.get("cumple")),
                    "Severidad": item.get("severidad", ""),
                    "Observacion": item.get("observacion", ""),
                    "Fotos": len(item.get("fotos", [])),
                }
            )
    return pd.DataFrame(rows)


def generar_pdf(registro):
    fecha_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_path = os.path.join(
        REPORTS_DIR,
        f"Informe_5S_PRO_{registro['bodega'].replace(' ', '_')}_{fecha_id}.pdf",
    )

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm,
    )

    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="TitleBlue",
            parent=styles["Title"],
            alignment=TA_CENTER,
            fontSize=18,
            leading=22,
            textColor=colors.HexColor("#061f45"),
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="HBlue",
            parent=styles["Heading2"],
            fontSize=14,
            leading=16,
            textColor=colors.HexColor("#061f45"),
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="NormalSmall",
            parent=styles["Normal"],
            fontSize=8.5,
            leading=10.5,
            alignment=TA_LEFT,
        )
    )
    styles.add(
        ParagraphStyle(
            name="MetaCenterValue",
            parent=styles["Normal"],
            alignment=TA_CENTER,
            fontSize=12,
            leading=16,
            textColor=colors.HexColor("#061f45"),
        )
    )

    story = []
    header_logo = ""
    if os.path.exists(LOGO_INOVA):
        try:
            header_logo = RLImage(LOGO_INOVA, width=2.2 * cm, height=2.2 * cm)
        except Exception:
            header_logo = ""

    score_color = "#13A35B" if registro["cumplimiento"] >= META_BODEGA else "#D53333"
    title = Paragraph(
        "<b>INFORME EJECUTIVO DE AUDITORIA 5S</b><br/>"
        "<font size='9' color='#52667A'>Control visual, cumplimiento por bodega, hallazgos y evidencias fotograficas</font>",
        styles["TitleBlue"],
    )
    top_header = Table([[header_logo, title]], colWidths=[2.6 * cm, 13.2 * cm])
    top_header.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (1, 0), "CENTER"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(top_header)

    meta_html = f"""
    <b>Fecha:</b> {registro['fecha']}<br/>
    <b>Semana:</b> {registro.get('semana', '')}<br/>
    <b>Responsable:</b> {registro['responsable']}<br/>
    <b>Bodega:</b> {registro['bodega']}<br/>
    <b>Area:</b> {registro['area']}<br/>
    <b>Meta requerida:</b> >= {META_BODEGA:.0f}%<br/><br/>
    <font size='18' color='{score_color}'><b>{registro['cumplimiento']:.1f}%</b></font><br/>
    <b>{cumplimiento_texto(registro['cumplimiento'])}</b>
    """
    meta_box = Table([[Paragraph(meta_html, styles["MetaCenterValue"])]], colWidths=[15.8 * cm])
    meta_box.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F7FAFD")),
                ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#D5E0EB")),
                ("LEFTPADDING", (0, 0), (-1, -1), 14),
                ("RIGHTPADDING", (0, 0), (-1, -1), 14),
                ("TOPPADDING", (0, 0), (-1, -1), 14),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]
        )
    )
    story.append(meta_box)
    story.append(Spacer(1, 0.35 * cm))

    items = normalizar_items_legacy(registro["items"])
    total_items = len(items)
    cumplidos = sum(1 for x in items if x["cumple"])
    no_conformes = total_items - cumplidos
    hallazgos = [x for x in items if (not x["cumple"]) or str(x.get("observacion", "")).strip()]

    story.append(Paragraph("1. Resumen ejecutivo", styles["HBlue"]))
    resumen_data = [
        ["Indicador", "Resultado"],
        ["Cumplimiento de bodega", f"{registro['cumplimiento']:.1f}%"],
        ["Meta bodega", f">= {META_BODEGA:.0f}%"],
        ["Nivel de desempeno", cumplimiento_texto(registro["cumplimiento"])],
        ["Puntos evaluados", str(total_items)],
        ["Puntos conformes", str(cumplidos)],
        ["Puntos no conformes", str(no_conformes)],
        ["Hallazgos / novedades", str(len(hallazgos))],
    ]
    resumen_table = Table(resumen_data, colWidths=[7.8 * cm, 7.1 * cm])
    resumen_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#061f45")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#c7d6e5")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("PADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.append(resumen_table)
    story.append(Spacer(1, 0.4 * cm))

    story.append(Paragraph("2. Matriz tecnica de verificacion", styles["HBlue"]))
    detalle = [["Item", "Punto evaluado", "Estado", "Severidad", "Observacion", "Fotos"]]
    for i, item in enumerate(items, start=1):
        detalle.append(
            [
                str(i),
                Paragraph(item["punto"], styles["NormalSmall"]),
                "Conforme" if item["cumple"] else "No conforme",
                item.get("severidad", ""),
                Paragraph(item.get("observacion") or "-", styles["NormalSmall"]),
                str(len(item.get("fotos", []))),
            ]
        )

    detalle_table = Table(
        detalle,
        colWidths=[0.9 * cm, 6.3 * cm, 2.0 * cm, 1.8 * cm, 3.9 * cm, 0.9 * cm],
        repeatRows=1,
    )
    detalle_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#061f45")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#c7d6e5")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("PADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(detalle_table)

    if hallazgos:
        story.append(PageBreak())
        story.append(Paragraph("3. Hallazgos y acciones sugeridas", styles["HBlue"]))
        for i, item in enumerate(hallazgos, start=1):
            accion = "Ejecutar accion correctiva y validar cierre en proxima auditoria."
            if item.get("cumple"):
                accion = "Mantener seguimiento preventivo y documentar oportunidad de mejora."
            texto = (
                f"<b>{i}. {item['punto']}</b><br/>"
                f"Estado: {'Conforme con observacion' if item.get('cumple') else 'No conforme'}<br/>"
                f"Severidad: {item.get('severidad', '-')}<br/>"
                f"Observacion: {item.get('observacion') or 'Sin detalle adicional.'}<br/>"
                f"Accion sugerida: {accion}"
            )
            story.append(Paragraph(texto, styles["Normal"]))
            story.append(Spacer(1, 0.18 * cm))

    evidencias = [x for x in items if x.get("fotos")]
    if evidencias:
        story.append(PageBreak())
        story.append(Paragraph("4. Evidencias fotograficas", styles["HBlue"]))
        contador = 1
        for item in evidencias:
            story.append(Paragraph(f"<b>Punto:</b> {item['punto']}", styles["Normal"]))
            story.append(Spacer(1, 0.1 * cm))
            for foto in item.get("fotos", []):
                try:
                    resize_image(foto, max_width=1300)
                    w_px, h_px = fit_image_box(foto)
                    img = RLImage(foto, width=min(w_px / 96, 15.2) * cm, height=min(h_px / 96, 8.3) * cm)
                    img_table = Table([[img]], colWidths=[16.0 * cm])
                    img_table.setStyle(
                        TableStyle(
                            [
                                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#d1dce8")),
                                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8fbff")),
                                ("PADDING", (0, 0), (-1, -1), 8),
                            ]
                        )
                    )
                    story.append(Paragraph(f"Evidencia {contador}", styles["NormalSmall"]))
                    story.append(img_table)
                    story.append(Spacer(1, 0.25 * cm))
                    contador += 1
                except Exception:
                    story.append(Paragraph("No fue posible cargar una evidencia fotografica.", styles["NormalSmall"]))

    doc.build(story)
    return pdf_path


def exportar_gantt_html(fig):
    semana = get_week_label()
    filename = f"Cronograma_5S_PRO_{semana}.html"
    html = fig.to_html(full_html=True, include_plotlyjs=True)
    buffer = io.StringIO()
    buffer.write(html)
    buffer.seek(0)
    return buffer, filename


def guardar_gantt_png(fig):
    import plotly.io as pio

    semana = get_week_label()
    filename = f"Cronograma_5S_PRO_{semana}.png"
    output_path = os.path.join(DATA_DIR, filename)
    img_bytes = pio.to_image(fig, format="png", width=2200, height=1100, scale=2)
    with open(output_path, "wb") as f:
        f.write(img_bytes)
    return output_path, filename

# =========================================================
# LOGIN
# =========================================================
def init_login():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False
    if "usuario_actual" not in st.session_state:
        st.session_state.usuario_actual = ""


def mostrar_login():
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"]{display:none !important;}
        #MainMenu, footer, header{visibility:hidden !important; height:0 !important;}
        html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"]{
            overflow-x:hidden !important;
        }
        .stApp{
            min-height:100vh !important;
            background:
                radial-gradient(circle at 18% 15%, rgba(21,108,193,0.18), transparent 20%),
                radial-gradient(circle at 82% 20%, rgba(6,31,69,0.14), transparent 22%),
                linear-gradient(135deg, #eef4fb 0%, #dfeaf5 100%) !important;
        }

        [data-testid="stMainBlockContainer"],
        .block-container{
            max-width:380px !important;
            width:380px !important;
            padding:0 !important;
            margin:0 auto !important;
            box-sizing:border-box !important;
            transform:translateY(calc(50vh - 190px));
        }

        [data-testid="stMainBlockContainer"] > div,
        .block-container > div{
            background:rgba(255,255,255,0.96) !important;
            border:1px solid rgba(255,255,255,0.88) !important;
            box-shadow:0 18px 45px rgba(6,31,69,0.20) !important;
            border-radius:18px !important;
            padding:18px 20px 20px 20px !important;
            box-sizing:border-box !important;
        }

        [data-testid="stVerticalBlock"]{gap:0.35rem !important;}

        div[data-testid="stImage"]{
            display:flex !important;
            justify-content:center !important;
            margin:0 !important;
        }
        div[data-testid="stImage"] img{
            width:72px !important;
            max-width:72px !important;
        }
        .login-title{
            text-align:center;
            color:#061f45;
            font-size:1.55rem;
            font-weight:900;
            letter-spacing:-0.035em;
            margin:4px 0 2px 0;
            line-height:1.05;
        }
        .login-sub{
            text-align:center;
            color:#667085;
            font-size:0.78rem;
            line-height:1.25;
            margin:0 0 12px 0;
        }

        div[data-testid="stForm"]{
            width:100% !important;
            max-width:340px !important;
            margin:0 auto !important;
            background:transparent !important;
            border:none !important;
            box-shadow:none !important;
            padding:0 !important;
        }
        div[data-testid="stTextInput"]{
            width:100% !important;
            max-width:340px !important;
            margin:0 auto 8px auto !important;
        }
        div[data-testid="stTextInput"] label,
        div[data-testid="stTextInput"] p{
            color:#061f45 !important;
            font-weight:900 !important;
            font-size:0.72rem !important;
            margin-bottom:2px !important;
        }
        div[data-testid="stTextInput"] input{
            height:38px !important;
            min-height:38px !important;
            color:#061f45 !important;
            background:#ffffff !important;
            border:1px solid #c8d8e8 !important;
            border-radius:8px !important;
            font-size:0.86rem !important;
        }
        div[data-testid="stFormSubmitButton"]{
            width:100% !important;
            max-width:340px !important;
            margin:8px auto 0 auto !important;
        }
        div[data-testid="stFormSubmitButton"] > button{
            width:100% !important;
            min-height:40px !important;
            height:40px !important;
            border-radius:8px !important;
            font-size:0.80rem !important;
            font-weight:900 !important;
            background:linear-gradient(135deg,#156cc1 0%,#0b4fc4 100%) !important;
            color:white !important;
            border:none !important;
        }
        .stAlert{
            max-width:340px !important;
            margin:8px auto 0 auto !important;
        }

        @media (max-width:520px){
            [data-testid="stMainBlockContainer"], .block-container{
                width:92vw !important;
                max-width:380px !important;
                transform:translateY(12vh);
            }
        }
        @media (max-height:560px){
            [data-testid="stMainBlockContainer"], .block-container{
                transform:none !important;
                margin-top:12px !important;
            }
            .login-sub{display:none !important;}
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    if os.path.exists(LOGO_INOVA):
        st.image(LOGO_INOVA, width=72)
    st.markdown('<div class="login-title">5S INOVA PRO</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="login-sub">Auditoría, control visual y excelencia operacional 5S</div>',
        unsafe_allow_html=True,
    )

    with st.form("login_form", clear_on_submit=False):
        usuario = st.text_input("USUARIO", placeholder="Ingrese su usuario", key="login_usuario").strip().upper()
        clave = st.text_input("CONTRASEÑA", type="password", placeholder="Ingrese su contraseña", key="login_clave")
        entrar = st.form_submit_button("ACCEDER", use_container_width=True)

    if entrar:
        if usuario in USUARIOS_SISTEMA and USUARIOS_SISTEMA[usuario] == clave:
            st.session_state.autenticado = True
            st.session_state.usuario_actual = usuario
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")


# =========================================================
# SESION
# =========================================================
def init_session():
    if "selected_bodega" not in st.session_state:
        st.session_state.selected_bodega = list(BODEGAS.keys())[0]
    if "cronograma" not in st.session_state:
        st.session_state.cronograma = safe_load_json(SCHEDULE_PATH, [])
    if "inspecciones" not in st.session_state:
        inspecciones_cargadas = safe_load_json(DB_PATH, [])
        for reg in inspecciones_cargadas:
            reg["items"] = normalizar_items_legacy(reg.get("items", []))
        st.session_state.inspecciones = inspecciones_cargadas
    if "responsables" not in st.session_state:
        st.session_state.responsables = safe_load_json(RESPONSIBLES_PATH, RESPONSABLES_DEFAULT)
        if not st.session_state.responsables:
            st.session_state.responsables = RESPONSABLES_DEFAULT
            safe_save_json(RESPONSIBLES_PATH, RESPONSABLES_DEFAULT)


init_login()
init_session()

if not st.session_state.autenticado:
    mostrar_login()
    st.stop()

# =========================================================
# HEADER PRINCIPAL
# =========================================================
def render_header():
    promedio = 0.0
    if st.session_state.inspecciones:
        df_head = build_inspection_dataframe(st.session_state.inspecciones)
        promedio = float(df_head["Cumplimiento"].mean()) if not df_head.empty else 0.0

    st.markdown('<div class="top-shell">', unsafe_allow_html=True)
    h1, h2 = st.columns([1.05, 7.4], gap="small")
    with h1:
        if os.path.exists(LOGO_INOVA):
            st.image(LOGO_INOVA, width=124)
        else:
            st.markdown("### 5S")
    with h2:
        status_html = cumplimiento_badge(promedio) if promedio else '<span class="badge-blue">Sistema activo</span>'
        st.markdown(
            f"""
            <div class="top-banner">
                <div>
                    <div class="top-title">Sistema 5S INOVA PRO</div>
                    <div class="top-subtitle">Cronograma ejecutivo, auditoria fotografica, cumplimiento por bodega, responsables, indicadores y reportes profesionales.</div>
                </div>
                <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;justify-content:flex-end;">
                    <div class="header-badge">Usuario: {st.session_state.usuario_actual}</div>
                    <div class="header-badge">Meta global: >= {META_GENERAL:.0f}%</div>
                    {status_html}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)


render_header()

# =========================================================
# SIDEBAR
# =========================================================
menu = st.sidebar.radio(
    "Módulos",
    ["Inicio Ejecutivo", "Cronograma 5S", "Inspección 5S", "Responsables", "Dashboard Ejecutivo"],
)

st.sidebar.markdown("---")
st.sidebar.subheader("Eliminar por día")
fecha_borrar_crono = st.sidebar.date_input("Fecha cronograma", value=date.today(), key="fecha_borrar_crono")
if st.sidebar.button("Eliminar cronograma del día", use_container_width=True):
    antes = len(st.session_state.cronograma)
    fecha_txt = str(fecha_borrar_crono)
    st.session_state.cronograma = [
        x for x in st.session_state.cronograma if str(x.get("fecha_inicio", ""))[:10] != fecha_txt
    ]
    safe_save_json(SCHEDULE_PATH, st.session_state.cronograma)
    st.sidebar.success(f"Actividades eliminadas: {antes - len(st.session_state.cronograma)}")

fecha_borrar_insp = st.sidebar.date_input("Fecha inspección", value=date.today(), key="fecha_borrar_insp")
if st.sidebar.button("Eliminar inspecciones del día", use_container_width=True):
    antes = len(st.session_state.inspecciones)
    fecha_txt = str(fecha_borrar_insp)
    st.session_state.inspecciones = [
        x for x in st.session_state.inspecciones if str(x.get("fecha", ""))[:10] != fecha_txt
    ]
    safe_save_json(DB_PATH, st.session_state.inspecciones)
    rebuild_excel_from_inspections(st.session_state.inspecciones)
    st.sidebar.success(f"Inspecciones eliminadas: {antes - len(st.session_state.inspecciones)}")

st.sidebar.markdown("---")
st.sidebar.markdown(f"**Sesión activa:** {st.session_state.usuario_actual}")
if st.sidebar.button("Cerrar sesión", use_container_width=True):
    st.session_state.autenticado = False
    st.session_state.usuario_actual = ""
    st.rerun()

# =========================================================
# INICIO EJECUTIVO
# =========================================================
if menu == "Inicio Ejecutivo":
    st.markdown('<div class="main-wrap">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Centro ejecutivo 5S</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-subtitle">Vista general del sistema, metas corporativas, auditorias y criticidad operacional.</div>',
        unsafe_allow_html=True,
    )

    df_ini = build_inspection_dataframe(st.session_state.inspecciones) if st.session_state.inspecciones else pd.DataFrame()
    promedio = float(df_ini["Cumplimiento"].mean()) if not df_ini.empty else 0.0
    bodegas_criticas = int((df_ini["Cumplimiento"] < META_BODEGA).sum()) if not df_ini.empty else 0
    total_bodegas = len(BODEGAS)
    total_puntos = sum(len(v) for v in BODEGAS.values())
    auditorias_semana = 0
    if not df_ini.empty:
        hoy = pd.Timestamp(date.today())
        inicio_semana = hoy - pd.Timedelta(days=hoy.weekday())
        auditorias_semana = int((df_ini["Fecha"] >= inicio_semana).sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Promedio general</div><div class="kpi-value">{promedio:.1f}%</div><div class="kpi-note">Meta >= {META_GENERAL:.0f}%</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Bodegas activas</div><div class="kpi-value">{total_bodegas}</div><div class="kpi-note">Checklist operativo</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Puntos de control</div><div class="kpi-value">{total_puntos}</div><div class="kpi-note">Distribuidos por bodega</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Auditorías semana</div><div class="kpi-value">{auditorias_semana}</div><div class="kpi-note">ISO semanal</div></div>', unsafe_allow_html=True)
    with c5:
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Alertas críticas</div><div class="kpi-value">{bodegas_criticas}</div><div class="kpi-note">Bajo meta</div></div>', unsafe_allow_html=True)

    st.markdown("---")
    c_left, c_right = st.columns([1.2, 1])
    with c_left:
        st.markdown("#### Mapa ejecutivo de bodegas")
        for bodega, puntos in BODEGAS.items():
            ult = None
            if not df_ini.empty:
                filtro = df_ini[df_ini["Bodega"] == bodega].sort_values("Fecha", ascending=False)
                if not filtro.empty:
                    ult = float(filtro.iloc[0]["Cumplimiento"])
            valor = ult if ult is not None else 0
            badge = cumplimiento_badge(valor) if ult is not None else '<span class="badge-gray">Sin auditoría</span>'
            st.markdown(
                f"""
                <div class="exec-card">
                    <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;">
                        <div>
                            <div style="font-weight:900;color:#061f45;font-size:1.05rem;">{bodega}</div>
                            <div style="color:#667085;font-size:0.9rem;">{len(puntos)} puntos de control · Meta >= {META_BODEGA:.0f}%</div>
                        </div>
                        <div>{badge}</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
    with c_right:
        st.markdown("#### Indicador corporativo")
        fig_gauge = go.Figure(
            go.Indicator(
                mode="gauge+number+delta",
                value=promedio,
                delta={"reference": META_GENERAL},
                gauge={
                    "axis": {"range": [0, 100]},
                    "bar": {"color": "#156CC1"},
                    "steps": [
                        {"range": [0, 75], "color": "#FDECEC"},
                        {"range": [75, 90], "color": "#FFF7DF"},
                        {"range": [90, 100], "color": "#E9F8EF"},
                    ],
                    "threshold": {"line": {"color": "#061F45", "width": 4}, "thickness": 0.75, "value": META_GENERAL},
                },
                title={"text": "Promedio General 5S"},
            )
        )
        fig_gauge.update_layout(height=360, margin=dict(l=20, r=20, t=50, b=20), paper_bgcolor="white")
        st.plotly_chart(fig_gauge, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# RESPONSABLES
# =========================================================
elif menu == "Responsables":
    st.markdown('<div class="main-wrap">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Gestión de responsables 5S</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Responsables base del Word y opción para añadir nuevos responsables.</div>', unsafe_allow_html=True)

    c1, c2 = st.columns([0.85, 1.4])
    with c1:
        st.markdown("#### Nuevo responsable")
        with st.form("form_responsable"):
            nombre = st.text_input("Nombre completo", placeholder="Ej. Nuevo responsable")
            cargo = st.text_input("Cargo", value="Responsable 5S")
            area = st.text_input("Área", value="Operaciones")
            color = st.color_picker("Color visual", value="#156CC1")
            guardar_resp = st.form_submit_button("Agregar responsable", use_container_width=True)
        if guardar_resp:
            if not nombre.strip():
                st.error("Debes ingresar el nombre del responsable.")
            else:
                nuevo = {
                    "id": slugify(nombre),
                    "nombre": nombre.strip(),
                    "cargo": cargo.strip() or "Responsable 5S",
                    "area": area.strip() or "Operaciones",
                    "color": color,
                    "activo": True,
                }
                st.session_state.responsables.append(nuevo)
                safe_save_json(RESPONSIBLES_PATH, st.session_state.responsables)
                st.success("Responsable agregado correctamente.")
                st.rerun()
    with c2:
        st.markdown("#### Responsables activos")
        for idx, r in enumerate(st.session_state.responsables):
            col_a, col_b, col_c = st.columns([0.1, 1.3, 0.35])
            with col_a:
                st.markdown(f"<div style='width:18px;height:18px;border-radius:50%;background:{r.get('color','#156CC1')};margin-top:10px;'></div>", unsafe_allow_html=True)
            with col_b:
                estado = "Activo" if r.get("activo", True) else "Inactivo"
                st.markdown(f"**{r.get('nombre')}**  ")
                st.caption(f"{r.get('cargo','')} · {r.get('area','')} · {estado}")
            with col_c:
                if st.button("Desactivar" if r.get("activo", True) else "Activar", key=f"toggle_resp_{idx}"):
                    st.session_state.responsables[idx]["activo"] = not r.get("activo", True)
                    safe_save_json(RESPONSIBLES_PATH, st.session_state.responsables)
                    st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# CRONOGRAMA
# =========================================================
elif menu == "Cronograma 5S":
    st.markdown('<div class="main-wrap">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Programar auditoría 5S</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Cronograma visual ejecutivo con eje X por fecha, bodega, día y responsable.</div>', unsafe_allow_html=True)

    responsables_activos = get_responsables_activos()
    nombres_responsables = [r["nombre"] for r in responsables_activos]

    with st.form("form_cronograma_pro"):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            bodega = st.selectbox("Bodega", list(BODEGAS.keys()))
        with c2:
            responsable = st.selectbox("Responsable", nombres_responsables if nombres_responsables else ["Sin responsable"])
        with c3:
            fecha_inicio = st.date_input("Día auditoría", value=date.today())
        with c4:
            estado = st.selectbox("Estado", ESTADOS_CRONOGRAMA, index=0)

        c5, c6, c7 = st.columns([1, 1, 1])
        with c5:
            fecha_fin = st.date_input("Fecha fin visual", value=date.today())
        with c6:
            prioridad = st.selectbox("Prioridad", ["Alta", "Media", "Baja"], index=1)
        with c7:
            actividad = st.text_input("Actividad", value="Auditoría 5S")

        observacion = st.text_area("Observación / alcance", height=90)
        guardar = st.form_submit_button("Guardar auditoría programada", use_container_width=True)

    if guardar:
        fecha_fin_real = fecha_fin
        if fecha_fin_real <= fecha_inicio:
            fecha_fin_real = fecha_inicio + timedelta(days=1)
        nuevo = {
            "id": datetime.now().strftime("%Y%m%d_%H%M%S_%f"),
            "bodega": bodega,
            "responsable": responsable,
            "actividad": actividad,
            "fecha_inicio": str(fecha_inicio),
            "fecha_fin": str(fecha_fin_real),
            "estado": estado,
            "prioridad": prioridad,
            "meta_bodega": META_BODEGA,
            "observacion": observacion,
        }
        st.session_state.cronograma.append(nuevo)
        safe_save_json(SCHEDULE_PATH, st.session_state.cronograma)
        st.success("Auditoría agregada al cronograma.")
        st.rerun()

    if st.session_state.cronograma:
        df_crono = pd.DataFrame(st.session_state.cronograma)

        # Blindaje para cronogramas antiguos o registros incompletos.
        # Plotly Express lanza ValueError si alguna columna usada en x/y/color/text/hover_data no existe.
        columnas_default_crono = {
            "bodega": "Sin bodega",
            "responsable": "Sin responsable",
            "actividad": "Auditoría 5S",
            "estado": "Programada",
            "prioridad": "Media",
            "meta_bodega": META_BODEGA,
            "observacion": "",
            "fecha_inicio": date.today(),
            "fecha_fin": date.today() + timedelta(days=1),
        }
        for col, default in columnas_default_crono.items():
            if col not in df_crono.columns:
                df_crono[col] = default
            df_crono[col] = df_crono[col].fillna(default)

        df_crono["fecha_inicio"] = pd.to_datetime(df_crono["fecha_inicio"], errors="coerce")
        df_crono["fecha_fin"] = pd.to_datetime(df_crono["fecha_fin"], errors="coerce")
        df_crono = df_crono.dropna(subset=["fecha_inicio", "fecha_fin"]).copy()

        # Evita barras con duración cero o fecha final menor a la inicial.
        mascara_fechas_invalidas = df_crono["fecha_fin"] <= df_crono["fecha_inicio"]
        df_crono.loc[mascara_fechas_invalidas, "fecha_fin"] = (
            df_crono.loc[mascara_fechas_invalidas, "fecha_inicio"] + pd.Timedelta(days=1)
        )

        fig = None
        if df_crono.empty:
            st.warning("No hay actividades válidas para graficar. Revisa las fechas del cronograma.")
        else:
            df_crono["estado_visual"] = df_crono.apply(lambda row: infer_schedule_status(row.to_dict()), axis=1)
            df_crono["estado_visual"] = df_crono["estado_visual"].fillna("Programada").astype(str)
            df_crono["etiqueta"] = df_crono["responsable"].fillna("Sin responsable").astype(str)
            df_crono["bodega"] = df_crono["bodega"].fillna("Sin bodega").astype(str)
            df_crono["dia"] = df_crono["fecha_inicio"].dt.strftime("%A %d %b")

            hover_cols = [
                col for col in ["actividad", "responsable", "dia", "prioridad", "meta_bodega", "observacion"]
                if col in df_crono.columns
            ]

            st.markdown("#### Cronograma visual ejecutivo")
            try:
                fig = px.timeline(
                    df_crono,
                    x_start="fecha_inicio",
                    x_end="fecha_fin",
                    y="bodega",
                    color="estado_visual",
                    text="etiqueta",
                    hover_data=hover_cols,
                    color_discrete_map=COLOR_ESTADO_VIVO,
                )
            except ValueError as e:
                st.error("No se pudo dibujar el cronograma porque hay datos antiguos/incompletos. Borra o corrige el cronograma guardado del día y vuelve a intentarlo.")
                st.caption(str(e))
                fig = None

            if fig is not None:
                fig.update_yaxes(autorange="reversed", title="Bodega", showgrid=True, gridcolor="#e5edf6")
                fig.update_xaxes(title="Fecha", showgrid=True, gridcolor="#dbe6f1", tickformat="%a %d %b")
                fig.update_traces(textposition="inside", insidetextanchor="middle", marker_line_color="white", marker_line_width=2.2, opacity=0.96)
                fig.update_layout(
                    height=720,
                    title="Cronograma 5S por bodega, día y responsable",
                    plot_bgcolor="#ffffff",
                    paper_bgcolor="#ffffff",
                    legend_title="Estado",
                    font=dict(size=13, color="#1f2937"),
                    margin=dict(l=150, r=40, t=72, b=45),
                    title_font=dict(size=22, color="#061f45"),
                    colorway=PALETA_VIVA,
                )
                st.plotly_chart(fig, use_container_width=True)

            if fig is not None:
                cex1, cex2, cex3 = st.columns(3)
                with cex1:
                    buffer = export_dataframe_excel(df_crono, "Cronograma")
                    st.download_button(
                        "Exportar cronograma Excel",
                        data=buffer.getvalue(),
                        file_name=f"Cronograma_5S_PRO_{get_week_label()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with cex2:
                    try:
                        gantt_png, gantt_name = guardar_gantt_png(fig)
                        with open(gantt_png, "rb") as f:
                            st.download_button(
                                "Exportar imagen Gantt",
                                data=f.read(),
                                file_name=gantt_name,
                                mime="image/png",
                                use_container_width=True,
                            )
                    except Exception:
                        st.info("PNG no disponible en este entorno. Usa HTML interactivo.")
                with cex3:
                    html_buffer, html_name = exportar_gantt_html(fig)
                    st.download_button(
                        "Exportar Gantt HTML",
                        data=html_buffer.getvalue(),
                        file_name=html_name,
                        mime="text/html",
                        use_container_width=True,
                    )

        with st.expander("Ver tabla de programación", expanded=False):
            st.dataframe(df_crono, use_container_width=True)
    else:
        st.info("No hay auditorías programadas todavía.")

    st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# INSPECCION
# =========================================================
elif menu == "Inspección 5S":
    st.markdown('<div class="main-wrap">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Inspección 5S por bodega</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Checklist técnico por bodega, evidencia fotográfica múltiple, severidad, cumplimiento y PDF ejecutivo.</div>', unsafe_allow_html=True)

    cols = st.columns(3)
    for i, bod in enumerate(BODEGAS.keys()):
        with cols[i % 3]:
            if st.button(bod, use_container_width=True):
                st.session_state.selected_bodega = bod

    bodega_actual = st.session_state.selected_bodega
    puntos = BODEGAS[bodega_actual]
    st.markdown(f"#### Bodega seleccionada: **{bodega_actual}**")
    st.markdown(f"Meta de cumplimiento: **>= {META_BODEGA:.0f}%** · Puntos de control: **{len(puntos)}**")

    responsables_activos = get_responsables_activos()
    nombres_responsables = [r["nombre"] for r in responsables_activos]
    c1, c2, c3 = st.columns(3)
    with c1:
        fecha_inspeccion = st.date_input("Fecha de inspección", value=date.today())
    with c2:
        responsable = st.selectbox("Responsable de inspección", nombres_responsables if nombres_responsables else ["Sin responsable"])
    with c3:
        area = st.text_input("Área / proceso", value="Almacenamiento")

    st.markdown("---")
    items = []
    for idx, punto in enumerate(puntos, start=1):
        st.markdown('<div class="ins-card">', unsafe_allow_html=True)
        box1, box2 = st.columns([0.9, 2.7])
        with box1:
            cumple = st.checkbox("Cumple", key=f"{bodega_actual}_{idx}_cumple")
            severidad = st.selectbox(
                "Severidad",
                ["Sin novedad", "Baja", "Media", "Alta", "Crítica"],
                index=0 if cumple else 2,
                key=f"{bodega_actual}_{idx}_sev",
            )
            fotos = st.file_uploader(
                f"Evidencias {idx}",
                type=["png", "jpg", "jpeg"],
                accept_multiple_files=True,
                key=f"{bodega_actual}_{idx}_foto",
            )
            if fotos:
                with st.expander(f"Ver {len(fotos)} evidencia(s)", expanded=False):
                    for n_foto, foto_item in enumerate(fotos, start=1):
                        st.image(foto_item, caption=f"Evidencia {idx}.{n_foto}", use_container_width=True)
        with box2:
            st.markdown(f'<div class="punto-title">Punto {idx}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="punto-sub">{punto}</div>', unsafe_allow_html=True)
            observacion = st.text_area(
                "Observación / hallazgo / acción requerida",
                key=f"{bodega_actual}_{idx}_obs",
                height=105,
                placeholder="Describe hallazgo, novedad, condición observada o acción correctiva requerida...",
            )
        st.markdown("</div>", unsafe_allow_html=True)
        items.append({"punto": punto, "cumple": cumple, "severidad": severidad, "observacion": observacion, "foto_obj": fotos})

    if st.button("Guardar inspección y generar informe PDF ejecutivo", type="primary", use_container_width=True):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder = os.path.join(EVIDENCE_DIR, f"{bodega_actual.replace(' ', '_')}_{timestamp}")
        os.makedirs(folder, exist_ok=True)

        total = len(items)
        cumplidos = 0
        items_final = []
        for i, item in enumerate(items, start=1):
            if item["cumple"]:
                cumplidos += 1
            foto_paths = []
            if item["foto_obj"]:
                for j, foto_subida in enumerate(item["foto_obj"], start=1):
                    foto_path = save_uploaded_image(foto_subida, folder, f"evidencia_{i}_{j}")
                    if foto_path:
                        foto_paths.append(foto_path)
            items_final.append(
                {
                    "punto": item["punto"],
                    "cumple": item["cumple"],
                    "severidad": item["severidad"],
                    "observacion": item["observacion"],
                    "fotos": foto_paths,
                }
            )

        cumplimiento = (cumplidos / total) * 100 if total > 0 else 0
        registro = {
            "id": timestamp,
            "fecha": str(fecha_inspeccion),
            "semana": get_week_label(datetime.combine(fecha_inspeccion, datetime.min.time())),
            "responsable": responsable,
            "area": area,
            "bodega": bodega_actual,
            "cumplimiento": round(cumplimiento, 2),
            "meta_bodega": META_BODEGA,
            "items": items_final,
        }

        st.session_state.inspecciones.append(registro)
        safe_save_json(DB_PATH, st.session_state.inspecciones)
        excel_file = append_to_excel(registro)

        try:
            pdf_file = generar_pdf(registro)
            st.success("Inspección guardada correctamente.")
            st.markdown(cumplimiento_badge(cumplimiento), unsafe_allow_html=True)
            with open(pdf_file, "rb") as f:
                st.download_button("Descargar informe PDF ejecutivo", data=f.read(), file_name=os.path.basename(pdf_file), mime="application/pdf", use_container_width=True)
        except Exception as e:
            st.error(f"Error generando PDF: {e}")

        if os.path.exists(excel_file):
            with open(excel_file, "rb") as f:
                st.download_button("Descargar histórico Excel", data=f.read(), file_name=os.path.basename(excel_file), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# DASHBOARD EJECUTIVO
# =========================================================
elif menu == "Dashboard Ejecutivo":
    st.markdown('<div class="main-wrap">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Dashboard ejecutivo de cumplimiento 5S</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Indicadores, ranking, tendencia, responsables, semáforos y hallazgos críticos.</div>', unsafe_allow_html=True)

    inspecciones = st.session_state.inspecciones
    if not inspecciones:
        st.info("Aún no hay inspecciones guardadas.")
    else:
        df = build_inspection_dataframe(inspecciones)
        df_items = build_items_dataframe(inspecciones)
        promedio = float(df["Cumplimiento"].mean())
        total_inspecciones = len(df)
        bajo_meta = int((df["Cumplimiento"] < META_BODEGA).sum())
        mejor_bodega = "-"
        peor_bodega = "-"
        if not df.empty:
            df_bodega_aux = df.groupby("Bodega", as_index=False)["Cumplimiento"].mean()
            mejor_bodega = df_bodega_aux.sort_values("Cumplimiento", ascending=False).iloc[0]["Bodega"]
            peor_bodega = df_bodega_aux.sort_values("Cumplimiento", ascending=True).iloc[0]["Bodega"]

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Promedio general</div><div class="kpi-value">{promedio:.1f}%</div><div class="kpi-note">Meta >= {META_GENERAL:.0f}%</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Inspecciones</div><div class="kpi-value">{total_inspecciones}</div><div class="kpi-note">Registros históricos</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Bajo meta</div><div class="kpi-value">{bajo_meta}</div><div class="kpi-note">Auditorías críticas</div></div>', unsafe_allow_html=True)
        with c4:
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Mejor bodega</div><div style="font-weight:900;color:#061f45;font-size:1.2rem;">{mejor_bodega}</div><div class="kpi-note">Ranking cumplimiento</div></div>', unsafe_allow_html=True)
        with c5:
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Bodega crítica</div><div style="font-weight:900;color:#061f45;font-size:1.2rem;">{peor_bodega}</div><div class="kpi-note">Prioridad de acción</div></div>', unsafe_allow_html=True)

        st.markdown("---")
        f1, f2, f3 = st.columns(3)
        with f1:
            bodega_filter = st.multiselect("Filtrar bodega", sorted(df["Bodega"].dropna().unique()))
        with f2:
            responsable_filter = st.multiselect("Filtrar responsable", sorted(df["Responsable"].dropna().unique()))
        with f3:
            estado_filter = st.multiselect("Filtrar estado", ["Excelente", "Atención", "Crítico"])

        df_view = df.copy()
        if bodega_filter:
            df_view = df_view[df_view["Bodega"].isin(bodega_filter)]
        if responsable_filter:
            df_view = df_view[df_view["Responsable"].isin(responsable_filter)]
        if estado_filter:
            df_view = df_view[df_view["Estado"].isin(estado_filter)]

        if df_view.empty:
            st.warning("No hay datos para los filtros seleccionados.")
        else:
            left, right = st.columns([1.2, 1])
            with left:
                st.markdown("#### Cumplimiento promedio por bodega")
                df_bodega = df_view.groupby("Bodega", as_index=False)["Cumplimiento"].mean().sort_values("Cumplimiento", ascending=False)
                fig_bar = px.bar(
                    df_bodega,
                    x="Bodega",
                    y="Cumplimiento",
                    text="Cumplimiento",
                    color="Cumplimiento",
                    color_continuous_scale=ESCALA_CUMPLIMIENTO_VIVA,
                    range_color=[0, 100],
                )
                fig_bar.add_hline(y=META_BODEGA, line_dash="dash", line_color="#061F45", annotation_text=f"Meta {META_BODEGA:.0f}%")
                fig_bar.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
                fig_bar.update_layout(height=500, plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=25, r=25, t=40, b=25), font=dict(color="#1f2937"), coloraxis_colorbar=dict(title="Cumplimiento"))
                st.plotly_chart(fig_bar, use_container_width=True)
            with right:
                st.markdown("#### Gauge ejecutivo")
                promedio_view = float(df_view["Cumplimiento"].mean())
                fig_gauge = go.Figure(
                    go.Indicator(
                        mode="gauge+number+delta",
                        value=promedio_view,
                        delta={"reference": META_GENERAL},
                        gauge={
                            "axis": {"range": [0, 100]},
                            "bar": {"color": "#156CC1"},
                            "steps": [
                                {"range": [0, 75], "color": "#FDECEC"},
                                {"range": [75, 90], "color": "#FFF7DF"},
                                {"range": [90, 100], "color": "#E9F8EF"},
                            ],
                            "threshold": {"line": {"color": "#061F45", "width": 4}, "thickness": 0.75, "value": META_GENERAL},
                        },
                        title={"text": "Cumplimiento filtrado"},
                    )
                )
                fig_gauge.update_layout(height=500, margin=dict(l=20, r=20, t=50, b=20), paper_bgcolor="white")
                st.plotly_chart(fig_gauge, use_container_width=True)

            st.markdown("#### Tendencia histórica")
            fig_line = px.line(
                df_view.sort_values("Fecha"),
                x="Fecha",
                y="Cumplimiento",
                color="Bodega",
                color_discrete_sequence=PALETA_VIVA,
                markers=True,
                hover_data=["Responsable", "Estado", "No conformes", "Observaciones"],
            )
            fig_line.add_hline(y=META_BODEGA, line_dash="dash", line_color="#061F45", annotation_text=f"Meta {META_BODEGA:.0f}%")
            fig_line.update_layout(height=440, plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=25, r=25, t=40, b=25), font=dict(color="#1f2937"))
            st.plotly_chart(fig_line, use_container_width=True)

            st.markdown("#### Ranking de responsables")
            df_resp = df_view.groupby("Responsable", as_index=False).agg(Cumplimiento=("Cumplimiento", "mean"), Auditorias=("Cumplimiento", "count")).sort_values("Cumplimiento", ascending=False)
            fig_resp = px.bar(df_resp, x="Responsable", y="Cumplimiento", text="Cumplimiento", color="Cumplimiento", color_continuous_scale=ESCALA_CUMPLIMIENTO_VIVA, hover_data=["Auditorias"])
            fig_resp.add_hline(y=META_BODEGA, line_dash="dash", line_color="#061F45")
            fig_resp.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig_resp.update_layout(height=420, plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=25, r=25, t=40, b=25))
            st.plotly_chart(fig_resp, use_container_width=True)

            if not df_items.empty:
                st.markdown("#### Hallazgos más repetidos")
                incumplidos = df_items[~df_items["Cumple"]].copy()
                if not incumplidos.empty:
                    df_hallazgos = incumplidos.groupby("Punto", as_index=False).size().rename(columns={"size": "No conformidades"}).sort_values("No conformidades", ascending=False).head(10)
                    fig_h = px.bar(df_hallazgos, x="No conformidades", y="Punto", orientation="h", text="No conformidades", color="No conformidades", color_continuous_scale=["#FFB703", "#FF8A00", "#D53333", "#7C3AED"])
                    fig_h.update_layout(height=520, yaxis=dict(autorange="reversed"), plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=20, r=20, t=30, b=20))
                    st.plotly_chart(fig_h, use_container_width=True)
                else:
                    st.success("No hay puntos no conformes registrados.")

            with st.expander("Ver histórico ejecutivo", expanded=False):
                st.dataframe(df_view.sort_values("Fecha", ascending=False), use_container_width=True)
                buffer = export_dataframe_excel(df_view, "Dashboard")
                st.download_button("Exportar dashboard filtrado Excel", data=buffer.getvalue(), file_name="Dashboard_5S_PRO.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)
