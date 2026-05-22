import html
import io
import os
import re
import unicodedata
import requests
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from datetime import datetime

# Fuente única: Excel en SharePoint / OneDrive (sin archivos locales)
# En Render/GitHub configure la variable de entorno EXCEL_URL (o secrets en Streamlit Cloud).
EXCEL_URL_DEFAULT = (
    "https://sceperucom-my.sharepoint.com/personal/daniel_bravo_sce-peru_com/"
    "_layouts/15/download.aspx?share=IQCWbu_Gzz_BS4AmOAjBzLZAAT058WpaGvFXW1RTLeRibuU"
)


def _excel_url():
    try:
        if hasattr(st, "secrets") and "EXCEL_URL" in st.secrets:
            return str(st.secrets["EXCEL_URL"]).strip()
    except Exception:
        pass
    return os.environ.get("EXCEL_URL", EXCEL_URL_DEFAULT).strip()


def _normalizar_url_sharepoint(url):
    """
    Convierte enlaces tipo :x:/g/... (vista web) a download.aspx (archivo .xlsx).
    """
    url = (url or "").strip().strip('"').strip("'")
    if not url:
        return url
    m = re.search(
        r"(https://[^/]+)/:x:/g/personal/([^/]+)/([^/?#]+)",
        url,
        re.IGNORECASE,
    )
    if m:
        base, usuario, share_id = m.group(1), m.group(2), m.group(3)
        return f"{base}/personal/{usuario}/_layouts/15/download.aspx?share={share_id}"
    return url


def _urls_a_probar(url):
    """Variantes de enlace SharePoint/OneDrive para descarga anónima."""
    url = _normalizar_url_sharepoint(url)
    if not url:
        return []
    urls = [url]
    if "download=1" not in url.lower() and "download.aspx" in url.lower():
        sep = "&" if "?" in url else "?"
        urls.append(f"{url}{sep}download=1")
    return list(dict.fromkeys(urls))


def _abrir_excel_remoto(url):
    """
    Descarga el .xlsx desde SharePoint (Render no puede usar el enlace como en tu PC).
    Requiere enlace compartido: cualquier persona con el enlace puede ver/descargar.
    """
    candidatos = _urls_a_probar(url)
    if not candidatos:
        raise ValueError(
            "EXCEL_URL no está configurada. En Render → Environment → agregue EXCEL_URL."
        )
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "*/*",
    }
    ultimo_error = None
    for intento_url in candidatos:
        try:
            resp = requests.get(
                intento_url, headers=headers, allow_redirects=True, timeout=120
            )
            if resp.status_code in (401, 403):
                raise PermissionError(
                    f"SharePoint respondió {resp.status_code}. El enlace no es público para internet: "
                    "use Compartir → *Cualquier persona con el enlace* (no solo SCE)."
                )
            resp.raise_for_status()
            if len(resp.content) < 2048 or resp.content[:2] != b"PK":
                tipo = (resp.headers.get("content-type") or "desconocido")[:60]
                raise ValueError(
                    f"La URL devolvió HTML o vacío (tipo: {tipo}), no un Excel. "
                    "Copie el enlace de *descarga* del archivo .xlsx."
                )
            buf = io.BytesIO(resp.content)
            buf.seek(0)
            return buf
        except Exception as exc:
            ultimo_error = exc
    raise ultimo_error or ValueError("No se pudo descargar el Excel.")


def _origen_excel_url():
    try:
        if hasattr(st, "secrets") and "EXCEL_URL" in st.secrets:
            return "Streamlit secrets"
    except Exception:
        pass
    if os.environ.get("EXCEL_URL"):
        return "variable Render EXCEL_URL"
    return "enlace por defecto del código (configure EXCEL_URL en Render)"


MSG_NO_DATA = "Dato no disponible por el momento"
VISTA_TRIMESTRE = "Trimestre"
VISTA_POR_FLOTA = "Por Flota"
VISTA_REDES = "Red de Estaciones"
VISTA_FICHA_VEHICULO = "Ficha vehículo"
VISTAS_MENU = [VISTA_TRIMESTRE, VISTA_POR_FLOTA, VISTA_REDES, VISTA_FICHA_VEHICULO]
_ALIASES_VISTA = {
    "Dashboard Combustible": VISTA_POR_FLOTA,
    "Por flota": VISTA_POR_FLOTA,
    "por flota": VISTA_POR_FLOTA,
}
MES_A_NUM = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SET": 9, "OCT": 10, "NOV": 11, "DIC": 12,
}
CHART_COLORS = ["#6366F1", "#14B8A6", "#F59E0B", "#EC4899", "#8B5CF6", "#06B6D4"]
PLOTLY_CFG = {"displayModeBar": False}

st.set_page_config(page_title="LOGISTIX AI | SCE", layout="wide")

st.markdown("""
    <style>
    * {
        margin: 0;
        padding: 0;
    }
    
    .stApp { 
        background: #000000;
        color: #ffffff; 
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }
    
    [data-testid="stSidebar"] { 
        background: linear-gradient(180deg, #0a0a0a 0%, #000000 100%) !important; 
        min-width: 260px !important; 
        border-right: 1px solid #1a1a1a;
    }
    
    [data-testid="stMain"] {
        background: #000000;
    }
    
    .sidebar-brand { 
        color: #ffffff; 
        font-size: 13px; 
        font-weight: 900; 
        text-align: center; 
        margin-bottom: 20px; 
        padding: 16px 0;
        letter-spacing: 2.5px;
        border-bottom: 2px solid #ffffff;
        text-transform: uppercase;
    }
    
    .kpi-card { 
        background: linear-gradient(145deg, #111827 0%, #0a0a0a 55%, #050505 100%);
        border-radius: 14px; 
        padding: 14px 12px; 
        border: 1px solid rgba(99, 102, 241, 0.35);
        margin-bottom: 8px;
        height: 96px;
        min-height: 96px;
        max-height: 96px;
        box-sizing: border-box;
        display: flex;
        flex-direction: column;
        justify-content: center;
        overflow: hidden;
        transition: border-color 0.2s ease;
        box-shadow: 0 6px 20px rgba(255, 255, 255, 0.05), inset 0 1px 0 rgba(255, 255, 255, 0.08);
    }
    
    .kpi-card:hover {
        border-color: #d4af37;
    }
    
    .kpi-label { 
        color: #a0a0a0; 
        font-size: 10px; 
        font-weight: 800; 
        text-transform: uppercase; 
        margin-bottom: 10px;
        letter-spacing: 1.2px;
    }
    
    .kpi-value { 
        font-size: 20px; 
        font-weight: 800; 
        color: #e8c547;
        letter-spacing: -0.3px;
        line-height: 1.15;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    
    .kpi-unit{
        font-size: 11px;
        color: #ffffff;
        font-weight: 700;
        margin-left: 2px;
    }
    div[data-testid="column"] .kpi-card {
        width: 100%;
    }
    .audit-table-wrap {
        max-height: 420px;
        overflow: auto;
        border: 1px solid #333;
        border-radius: 8px;
        margin: 8px 0;
    }
    table.audit-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
    }
    table.audit-table th {
        background: #1a1a1a;
        color: #FFD700;
        padding: 10px 8px;
        text-align: left;
        position: sticky;
        top: 0;
        z-index: 1;
        border-bottom: 2px solid #FFD700;
    }
    table.audit-table td {
        padding: 8px;
        border-bottom: 1px solid #2a2a2a;
    }
    .red-net-toolbar {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        align-items: center;
        margin: 8px 0 14px 0;
        padding: 10px 12px;
        background: #0a0a0a;
        border: 1px solid #2a2a2a;
        border-radius: 10px;
    }
    .red-net-toolbar .toolbar-label {
        color: #aaaaaa;
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-right: 4px;
    }
    .btn-grafico-row {
        display: flex;
        gap: 8px;
        margin-bottom: 8px;
    }
    
    .main-title {
        font-size: 42px;
        font-weight: 700;
        color: #ffffff;
        text-align: left;
        margin: 30px 0 28px 0;
        letter-spacing: -1px;
    }
    
    .section-title {
        font-size: 12px;
        font-weight: 800;
        color: #ffffff;
        margin: 24px 0 16px 0;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        padding-bottom: 8px;
        border-bottom: 1px solid #1a1a1a;
    }
    
    .header-top { 
        padding: 10px 14px; 
        border-radius: 4px 4px 0 0; 
        font-weight: 700; 
        font-size: 10px; 
        color: #ffffff;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .bg-success { background: #1a3a1a; border: 1px solid #ffffff; }
    .bg-danger { background: #3a1a1a; border: 1px solid #ffffff; }
    .bg-warning { background: #3a3a1a; border: 1px solid #ffffff; }
    .bg-info { background: #1a2a3a; border: 1px solid #ffffff; }
    
    .divider-line {
        height: 1px;
        background: linear-gradient(90deg, transparent, #1a1a1a, transparent);
        margin: 20px 0;
        border: none;
    }
    
    .info-box {
        background: #0a0a0a;
        border-left: 3px solid #d4af37;
        padding: 12px 14px;
        border-radius: 2px;
        color: #ffffff;
        font-size: 12px;
        line-height: 1.5;
    }
    
    .warning-box {
        background: #0a0a0a;
        border-left: 3px solid #808080;
        padding: 12px 14px;
        border-radius: 2px;
        color: #ffffff;
        font-size: 12px;
    }
    
    .update-info {
        background: #0a1a0a;
        border-left: 3px solid #00ff00;
        padding: 8px 12px;
        border-radius: 2px;
        color: #00ff00;
        font-size: 10px;
        font-weight: 600;
        margin-bottom: 12px;
    }
    
    [data-testid="stRadio"] {
        display: flex;
        gap: 10px;
    }
    
    [data-testid="stRadio"] > label > div:first-child {
        accent-color: #ffd700 !important;
    }
    
    [role="radio"] {
        accent-color: #ffd700 !important;
    }
    
    .stSlider > div > div > div > input {
        accent-color: #ffd700 !important;
    }
    
    [data-baseweb="slider"] {
        --slider-color: #ffd700 !important;
    }
    
    [data-testid="stSlider"] input {
        accent-color: #ffd700 !important;
    }
    /* Meta KM/G: slider nativo arrastrable (único slider del sidebar) */
    .meta-kmg-labels {
        display: flex;
        justify-content: space-between;
        align-items: center;
        color: #ffffff;
        font-size: 11px;
        margin: 4px 0 12px 0;
        background: #000000;
    }
    .meta-kmg-labels .val {
        font-weight: 700;
        font-size: 14px;
    }
    section[data-testid="stSidebar"] [data-testid="stSlider"] {
        background: transparent !important;
        padding: 8px 0 4px 0 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSlider"] > div {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSlider"] [data-baseweb="slider"] {
        margin-top: 4px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSlider"] [data-baseweb="slider"] > div {
        background-color: #333333 !important;
        height: 3px !important;
        border-radius: 2px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSlider"] [data-baseweb="slider"] > div > div {
        background-color: #FFD700 !important;
        height: 3px !important;
        border-radius: 2px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSlider"] [role="slider"] {
        background-color: #FFD700 !important;
        border: 2px solid #000000 !important;
        width: 16px !important;
        height: 16px !important;
        cursor: grab !important;
        box-shadow: none !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSlider"] div[data-testid="stThumbValue"] {
        display: none !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSlider"] [data-testid="stTickBarMin"],
    section[data-testid="stSidebar"] [data-testid="stSlider"] [data-testid="stTickBarMax"] {
        display: none !important;
    }
    .trimestre-mes-bloque {
        margin-bottom: 28px;
        padding-bottom: 16px;
        border-bottom: 1px solid #2a2a2a;
    }
    .trimestre-mes-titulo {
        color: #FFD700;
        font-size: 14px;
        font-weight: 800;
        letter-spacing: 1px;
        margin-bottom: 10px;
        text-transform: uppercase;
    }
    .sidebar-simple-label {
        color: #cccccc;
        font-size: 13px;
        font-weight: 600;
        margin: 12px 0 4px 0;
    }
    section[data-testid="stSidebar"] [data-testid="stSelectbox"] label {
        display: none !important;
    }
    .map-legend-redes {
        display: flex;
        gap: 24px;
        margin: 8px 0 16px 0;
        flex-wrap: wrap;
    }
    .map-legend-item {
        display: flex;
        align-items: center;
        gap: 8px;
        color: #ffffff;
        font-size: 13px;
        font-weight: 600;
    }
    .map-dot-primax { width: 14px; height: 14px; border-radius: 50%; background: #2563eb; border: 2px solid #fff; }
    .map-dot-redcol { width: 14px; height: 14px; border-radius: 50%; background: #ef4444; border: 2px solid #fff; }
    
    .semaforo-excelente {
        background-color: #1a4d1a;
        border: 1px solid #00ff00;
        color: #00ff00;
        font-weight: 700;
        text-align: center;
        padding: 4px 8px;
        border-radius: 3px;
    }
    
    .semaforo-alerta {
        background-color: #4d4d00;
        border: 1px solid #ffff00;
        color: #ffff00;
        font-weight: 700;
        text-align: center;
        padding: 4px 8px;
        border-radius: 3px;
    }
    
    .semaforo-critico {
        background-color: #4d1a1a;
        border: 1px solid #ff0000;
        color: #ff0000;
        font-weight: 700;
        text-align: center;
        padding: 4px 8px;
        border-radius: 3px;
    }
    
    .legend-semaforo {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 12px;
        margin: 16px 0;
    }
    
    .legend-item {
        padding: 8px 12px;
        border-radius: 3px;
        font-size: 11px;
        font-weight: 700;
        text-align: center;
    }
    
    .legend-excelente {
        background-color: #16a34a;
        border: 2px solid #22c55e;
        color: #ffffff;
    }
    
    .legend-alerta {
        background-color: #eab308;
        border: 2px solid #facc15;
        color: #1a1a1a;
    }
    
    .legend-critico {
        background-color: #dc2626;
        border: 2px solid #ef4444;
        color: #ffffff;
    }
    
    .tabla-heatmap-dep {
        width: max-content;
        min-width: 100%;
        border-collapse: collapse;
        font-size: 12px;
        background: #000000;
        margin: 0;
        table-layout: auto;
    }
    .tabla-heatmap-dep th {
        background: #0f0f0f;
        color: #cccccc;
        padding: 10px 8px;
        text-align: center;
        border: 1px solid #222222;
        font-weight: 800;
        text-transform: uppercase;
        font-size: 11px;
        white-space: nowrap;
    }
    .tabla-heatmap-dep td {
        border: 1px solid #1a1a1a;
        padding: 8px 10px;
        text-align: right;
        white-space: nowrap;
        vertical-align: middle;
    }
    .tabla-heatmap-dep .col-prov {
        text-align: left;
        padding: 8px 12px;
        color: #ffffff;
        background: #0a0a0a;
        font-weight: 700;
        min-width: 140px;
        max-width: 180px;
        position: sticky;
        left: 0;
        z-index: 2;
    }
    .tabla-heatmap-dep tfoot td {
        background: #111111;
        color: #ffffff;
        font-weight: 900;
        padding: 10px 8px;
        text-align: right;
        border-top: 2px solid #ffd700;
    }
    .tabla-heatmap-dep .titulo-anual {
        font-size: 22px;
        font-weight: 900;
        color: #ffffff;
        text-align: center;
        margin: 20px 0 12px 0;
        letter-spacing: 1px;
    }
    .combustible-table {
        background: #0a0a0a;
        padding: 16px;
        border-radius: 8px;
        border: 1px solid #1a1a1a;
        margin: 16px 0;
        width: 100%;
        box-sizing: border-box;
    }
    .tabla-heatmap-scroll {
        width: 100%;
        overflow-x: auto;
        overflow-y: visible;
        -webkit-overflow-scrolling: touch;
    }
    
    .combustible-header {
        display: flex;
        gap: 12px;
        margin-bottom: 12px;
        flex-wrap: wrap;
    }
    
    .combustible-item {
        background: linear-gradient(135deg, #0a1a0a 0%, #050505 100%);
        border: 2px solid #ffd700;
        border-radius: 8px;
        padding: 12px 20px;
        text-align: center;
        min-width: 150px;
    }
    
    .combustible-item.no-data {
        background: linear-gradient(135deg, #1a0a0a 0%, #050505 100%);
        border: 2px solid #555555;
        opacity: 0.7;
    }
    
    .combustible-tipo {
        font-size: 12px;
        font-weight: 700;
        color: #ffd700;
        text-transform: uppercase;
        margin-bottom: 4px;
    }
    
    .combustible-tipo.no-data {
        color: #888888;
    }
    
    .combustible-valor {
        font-size: 28px;
        font-weight: 800;
        color: #ffffff;
    }
    
    .combustible-valor.no-data {
        color: #666666;
        font-size: 20px;
    }
    
    .combustible-unidad {
        font-size: 10px;
        color: #a0a0a0;
        margin-top: 4px;
    }
    
    .combustible-unidad.no-data {
        color: #555555;
    }

    /* Selección amarilla: panel lateral y controles */
    section[data-testid="stSidebar"] [data-baseweb="tag"] {
        background-color: #FFD700 !important;
        color: #111111 !important;
        border: 1px solid #e6c200 !important;
    }
    section[data-testid="stSidebar"] [data-baseweb="tag"] span,
    section[data-testid="stSidebar"] [data-baseweb="tag"] svg {
        color: #111111 !important;
        fill: #111111 !important;
    }
    [data-testid="stMultiSelect"] [data-baseweb="tag"] {
        background-color: #FFD700 !important;
        color: #111111 !important;
    }
    /* Radio: solo punto/círculo amarillo — texto sin sombreado */
    [data-testid="stRadio"] label[data-baseweb="radio"] {
        background: transparent !important;
        background-color: transparent !important;
    }
    [data-testid="stRadio"] label[data-baseweb="radio"] p,
    [data-testid="stRadio"] label[data-baseweb="radio"] span,
    [data-testid="stRadio"] label[data-baseweb="radio"] div[data-testid="stMarkdownContainer"] {
        background: transparent !important;
        color: #ffffff !important;
    }
    [data-testid="stRadio"] label[data-baseweb="radio"] input + div {
        background: transparent !important;
        border: 2px solid #555555 !important;
    }
    [data-testid="stRadio"] label[data-baseweb="radio"] input:checked + div {
        border-color: #FFD700 !important;
    }
    [data-testid="stRadio"] label[data-baseweb="radio"] input:checked + div > div {
        background-color: #FFD700 !important;
    }
    [data-testid="stRadio"] label[data-baseweb="radio"] div[style*="rgb(255, 75"],
    [data-testid="stRadio"] label[data-baseweb="radio"] div[style*="rgb(255, 43"] {
        background-color: #FFD700 !important;
    }
    [data-testid="stRadio"] input[type="radio"] {
        accent-color: #FFD700 !important;
    }
    [data-testid="stRadio"] label:has(input:checked),
    [data-testid="stRadio"] label:has(input:checked) > div,
    [data-testid="stRadio"] label:has(input:checked) p,
    [data-testid="stRadio"] label:has(input:checked) span {
        background: transparent !important;
        background-color: transparent !important;
        color: #ffffff !important;
        box-shadow: none !important;
    }
    /* Desplegables y listas del panel: fondo oscuro (evita blanco) */
    section[data-testid="stSidebar"] [data-baseweb="select"] > div,
    section[data-testid="stSidebar"] [data-baseweb="input"],
    section[data-testid="stSidebar"] input {
        background-color: #111111 !important;
        color: #ffffff !important;
        border-color: #333333 !important;
    }
    [data-baseweb="popover"],
    [data-baseweb="menu"],
    ul[role="listbox"],
    div[data-baseweb="popover"] > div {
        background-color: #1a1a1a !important;
        color: #ffffff !important;
    }
    [data-baseweb="menu"] li,
    [role="option"],
    [data-baseweb="menu"] ul li {
        background-color: #1a1a1a !important;
        color: #ffffff !important;
    }
    [role="option"]:hover,
    [data-baseweb="menu"] li:hover {
        background-color: #2a2a2a !important;
        color: #FFD700 !important;
    }
    [role="option"][aria-selected="true"] {
        background-color: #333333 !important;
        color: #FFD700 !important;
    }
    .sidebar-panel-title {
        color: #FFD700;
        font-size: 11px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin: 8px 0 12px 0;
    }
    .sidebar-vista-label {
        color: #FFD700;
        font-size: 12px;
        font-weight: 700;
        margin: 0 0 8px 0;
    }
    section[data-testid="stSidebar"] button[kind="primary"] {
        background: linear-gradient(135deg, #2a2200 0%, #1a1a00 100%) !important;
        border: 2px solid #FFD700 !important;
        color: #FFD700 !important;
        font-weight: 700 !important;
    }
    section[data-testid="stSidebar"] button[kind="secondary"] {
        background: #111111 !important;
        border: 1px solid #333333 !important;
        color: #cccccc !important;
    }
    section[data-testid="stSidebar"] button[kind="secondary"]:hover {
        border-color: #FFD700 !important;
        color: #ffffff !important;
    }
    section[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] {
        flex-direction: row !important;
        flex-wrap: wrap !important;
        gap: 10px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"],
    section[data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked),
    section[data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"] > div:first-child {
        background: transparent !important;
        background-color: transparent !important;
        box-shadow: none !important;
    }
    section[data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"] p,
    section[data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"] span {
        background: transparent !important;
        color: #ffffff !important;
    }
    /* Tabla de control: fila seleccionada */
    [data-testid="stDataFrame"] [aria-selected="true"],
    [data-testid="stDataFrame"] tr:focus-within,
    [data-testid="stDataFrame"] tbody tr:hover {
        background-color: rgba(255, 215, 0, 0.35) !important;
    }
    [data-testid="stDataFrame"] [aria-selected="true"] td {
        background-color: #FFD700 !important;
        color: #111111 !important;
    }
    .vehiculo-holo-stage {
        position: relative;
        width: 100%;
        min-height: 520px;
        border-radius: 18px;
        overflow: hidden;
        background:
            radial-gradient(ellipse 80% 60% at 50% 45%, rgba(255,215,0,0.18) 0%, transparent 55%),
            radial-gradient(ellipse 100% 80% at 50% 100%, rgba(37,99,235,0.12) 0%, transparent 50%),
            linear-gradient(180deg, #0d0d0d 0%, #050505 100%);
        border: 1px solid rgba(255, 215, 0, 0.4);
        box-shadow: 0 0 40px rgba(255, 215, 0, 0.08), inset 0 0 60px rgba(255, 215, 0, 0.04);
    }
    .vehiculo-holo-stage model-viewer {
        width: 100%;
        height: 520px;
        --poster-color: transparent;
    }
    .holo-ring {
        position: absolute;
        left: 50%;
        top: 52%;
        width: 280px;
        height: 280px;
        margin: -140px 0 0 -140px;
        border-radius: 50%;
        border: 1px solid rgba(255, 215, 0, 0.25);
        pointer-events: none;
        animation: holo-spin 12s linear infinite;
    }
    .holo-ring-2 {
        width: 340px;
        height: 340px;
        margin: -170px 0 0 -170px;
        border-color: rgba(37, 99, 235, 0.2);
        animation-duration: 18s;
        animation-direction: reverse;
    }
    @keyframes holo-spin {
        from { transform: rotateX(62deg) rotateZ(0deg); }
        to { transform: rotateX(62deg) rotateZ(360deg); }
    }
    .sce-logo-3d {
        position: absolute;
        left: 50%;
        top: 58%;
        transform: translate(-50%, -50%) rotateY(-18deg);
        z-index: 5;
        font-size: 42px;
        font-weight: 900;
        letter-spacing: 6px;
        color: #ffd700;
        text-shadow: 0 0 20px rgba(255,215,0,0.9), 0 2px 8px rgba(0,0,0,0.9);
        pointer-events: none;
        font-family: "Segoe UI", Arial, sans-serif;
    }
    .holo-veh-titulo {
        position: absolute;
        top: 14px;
        left: 0;
        right: 0;
        text-align: center;
        z-index: 6;
        color: #fff;
        font-size: 15px;
        font-weight: 700;
        text-shadow: 0 1px 4px #000;
    }
    .holo-veh-sub {
        position: absolute;
        top: 36px;
        left: 0;
        right: 0;
        text-align: center;
        z-index: 6;
        color: #ffd700;
        font-size: 12px;
        font-weight: 600;
    }
    .holo-veh-hint {
        position: absolute;
        bottom: 10px;
        left: 0;
        right: 0;
        text-align: center;
        color: #888;
        font-size: 11px;
        z-index: 6;
    }
    .vehiculo-specs-panel {
        max-height: 520px;
        overflow-y: auto;
        padding-right: 4px;
    }
    .vehiculo-spec-card {
        background: linear-gradient(145deg, #141414 0%, #0a0a0a 100%);
        border: 1px solid #2a2a2a;
        border-left: 3px solid #ffd700;
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 10px;
    }
    .vehiculo-spec-card .spec-label {
        color: #888;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .vehiculo-spec-card .spec-value {
        color: #fff;
        font-size: 15px;
        font-weight: 700;
        margin-top: 2px;
    }
    .watermark-turromzita {
        position: fixed;
        bottom: 14px;
        right: 18px;
        z-index: 99999;
        font-family: "Segoe UI", "Palatino Linotype", Georgia, serif;
        font-size: 13px;
        font-style: italic;
        font-weight: 500;
        color: rgba(255, 255, 255, 0.28);
        letter-spacing: 0.4px;
        pointer-events: none;
        user-select: none;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.5);
    }
    </style>
    """, unsafe_allow_html=True)

MESES_COMPLETOS = {
    "ENE": "ENERO", "FEB": "FEBRERO", "MAR": "MARZO", "ABR": "ABRIL",
    "MAY": "MAYO", "JUN": "JUNIO", "JUL": "JULIO", "AGO": "AGOSTO",
    "SET": "SETIEMBRE", "OCT": "OCTUBRE", "NOV": "NOVIEMBRE", "DIC": "DICIEMBRE"
}
NUM_A_MES = {v: k for k, v in MES_A_NUM.items()}
ORDEN_MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SET", "OCT", "NOV", "DIC"]

PROVINCIAS_PERU = {
    'LIMA': {'lat': -12.0464, 'lon': -77.0428}, 'CALLAO': {'lat': -12.0561, 'lon': -77.1181},
    'BARRANCA': {'lat': -10.7512, 'lon': -77.7665},
    'HUAURA': {'lat': -11.1239, 'lon': -77.6122}, 'OYÓN': {'lat': -10.7067, 'lon': -76.7573},
    'HUAROCHIRÍ': {'lat': -11.8482, 'lon': -76.3876}, 'CAÑETE': {'lat': -13.0788, 'lon': -76.3812},
    'YAUYOS': {'lat': -11.8969, 'lon': -75.7711}, 'TRUJILLO': {'lat': -8.1117, 'lon': -79.0288},
    'PACASMAYO': {'lat': -7.3823, 'lon': -79.5161}, 'CHEPÉN': {'lat': -7.1827, 'lon': -79.3227},
    'ASCOPE': {'lat': -7.5876, 'lon': -78.9822}, 'VIRÚ': {'lat': -8.3846, 'lon': -79.1494},
    'JULCÁN': {'lat': -7.8667, 'lon': -78.3667}, 'SÁNCHEZ CARRIÓN': {'lat': -7.5167, 'lon': -78.1667},
    'PATAZ': {'lat': -7.6333, 'lon': -77.3333}, 'GRAN CHIMÚ': {'lat': -7.4917, 'lon': -78.8583},
    'CHICLAYO': {'lat': -6.7713, 'lon': -79.8409}, 'LAMBAYEQUE': {'lat': -6.5453, 'lon': -79.8947},
    'FERREÑAFE': {'lat': -6.6667, 'lon': -79.7333}, 'PIURA': {'lat': -5.1945, 'lon': -80.6328},
    'SULLANA': {'lat': -4.9128, 'lon': -80.6703}, 'PAITA': {'lat': -5.0912, 'lon': -81.1043},
    'TALARA': {'lat': -4.5781, 'lon': -81.2700}, 'AYABACA': {'lat': -4.6333, 'lon': -79.6667},
    'HUANCABAMBA': {'lat': -5.2561, 'lon': -79.3232}, 'MORROPÓN': {'lat': -5.5833, 'lon': -80.0333},
    'SECHURA': {'lat': -5.5597, 'lon': -80.8238}, 'CAJAMARCA': {'lat': -7.1617, 'lon': -78.5128},
    'CELENDÍN': {'lat': -6.8667, 'lon': -78.1333}, 'CONTUMAZÁ': {'lat': -7.0333, 'lon': -78.8167},
    'HUAMACHUCO': {'lat': -7.8167, 'lon': -77.8333}, 'JAÉN': {'lat': -5.7131, 'lon': -78.8131},
    'SAN IGNACIO': {'lat': -5.2667, 'lon': -78.6167}, 'SAN MIGUEL': {'lat': -6.5667, 'lon': -78.9333},
    'CHOTA': {'lat': -6.5333, 'lon': -79.1667}, 'BAMBAMARCA': {'lat': -7.0167, 'lon': -78.9167},
    'HUARAZ': {'lat': -9.5332, 'lon': -77.5287}, 'AIJA': {'lat': -10.1833, 'lon': -77.3333},
    'ANTONIO RAYMONDI': {'lat': -8.6667, 'lon': -77.8333}, 'ASUNCIÓN': {'lat': -9.1667, 'lon': -77.1333},
    'BOLOGNESI': {'lat': -9.5667, 'lon': -77.4667}, 'CARHUAZ': {'lat': -9.0667, 'lon': -77.7333},
    'CASMA': {'lat': -9.5000, 'lon': -78.5167}, 'CORONGO': {'lat': -9.1667, 'lon': -77.9833},
    'HUARI': {'lat': -10.0167, 'lon': -77.3333}, 'HUARMEY': {'lat': -10.0667, 'lon': -78.1500},
    'MARISCAL LUZURIAGA': {'lat': -10.5500, 'lon': -77.2833}, 'OCROS': {'lat': -10.1333, 'lon': -76.6333},
    'PALLASCA': {'lat': -7.5667, 'lon': -77.6333}, 'POMABAMBA': {'lat': -9.6667, 'lon': -76.9500},
    'RECUAY': {'lat': -10.0333, 'lon': -77.2167}, 'SANTA': {'lat': -9.1833, 'lon': -78.1167},
    'SIHUAS': {'lat': -9.7667, 'lon': -77.1333}, 'TINCO': {'lat': -9.3167, 'lon': -77.7500},
    'YUNGAY': {'lat': -9.1500, 'lon': -77.5833}, 'HUANCAYO': {'lat': -12.0651, 'lon': -75.2049},
    'CONCEPCIÓN': {'lat': -12.3833, 'lon': -75.2333}, 'CHANCHAMAYO': {'lat': -11.6333, 'lon': -75.3333},
    'JAUJA': {'lat': -12.2689, 'lon': -75.4833}, 'JUNÍN': {'lat': -12.1167, 'lon': -75.2333},
    'SATIPO': {'lat': -11.2833, 'lon': -74.6333}, 'TARMA': {'lat': -11.4167, 'lon': -75.6833},
    'YAULI': {'lat': -12.2667, 'lon': -75.1667}, 'CHUPACA': {'lat': -12.0667, 'lon': -75.3667},
    'CERRO DE PASCO': {'lat': -10.6821, 'lon': -76.2565}, 'DANIEL ALCIDES CARRIÓN': {'lat': -10.6667, 'lon': -75.8333},
    'OXAPAMPA': {'lat': -10.5833, 'lon': -75.3833}, 'HUÁNUCO': {'lat': -9.9306, 'lon': -76.2422},
    'AMBO': {'lat': -10.2333, 'lon': -76.1667}, 'LA UNIÓN': {'lat': -10.0500, 'lon': -76.8167},
    'LAURICOCHA': {'lat': -10.8333, 'lon': -76.3333}, 'MARAÑÓN': {'lat': -10.1333, 'lon': -75.5667},
    'PACHITEA': {'lat': -10.4000, 'lon': -75.4333}, 'PUERTO INCA': {'lat': -9.7167, 'lon': -75.1167},
    'YAROWILCA': {'lat': -10.1833, 'lon': -76.6167}, 'ICA': {'lat': -14.0681, 'lon': -75.7286},
    'CHINCHA': {'lat': -13.5000, 'lon': -75.9667}, 'NAZCA': {'lat': -14.8367, 'lon': -74.9333},
    'PISCO': {'lat': -13.7083, 'lon': -76.2167}, 'PALPA': {'lat': -14.5667, 'lon': -75.7167},
    'AYACUCHO': {'lat': -13.1588, 'lon': -74.2239}, 'CANGALLO': {'lat': -13.8333, 'lon': -74.1667},
    'HUAMANGA': {'lat': -13.1583, 'lon': -74.2242}, 'HUANTA': {'lat': -12.9167, 'lon': -74.1667},
    'LA MAR': {'lat': -12.7333, 'lon': -74.2333}, 'LUCANAS': {'lat': -13.5667, 'lon': -74.5333},
    'PÁUCAR DEL SARA SARA': {'lat': -14.0333, 'lon': -74.6667}, 'SUCRE': {'lat': -13.9333, 'lon': -74.1333},
    'VÍCTOR FAJARDO': {'lat': -13.8333, 'lon': -74.0333}, 'VILCASHUAMÁN': {'lat': -13.6167, 'lon': -74.0167},
    'ABANCAY': {'lat': -13.6356, 'lon': -72.8816}, 'ANDAHUAYLAS': {'lat': -13.6551, 'lon': -73.3842},
    'ANTABAMBA': {'lat': -15.5333, 'lon': -73.6333}, 'AYMARAES': {'lat': -14.4667, 'lon': -72.8167},
    'CHINCHEROS': {'lat': -13.3833, 'lon': -72.8167}, 'GRAU': {'lat': -14.4333, 'lon': -73.1667},
    'CUSCO': {'lat': -13.5319, 'lon': -71.9675}, 'ACOMAYO': {'lat': -13.8167, 'lon': -71.8333},
    'ANTA': {'lat': -13.4667, 'lon': -72.1833}, 'CALCA': {'lat': -12.0917, 'lon': -71.9000},
    'CANAS': {'lat': -14.3333, 'lon': -71.4333}, 'CANCHIS': {'lat': -13.6667, 'lon': -71.6667},
    'CHUMBIVILCAS': {'lat': -14.7333, 'lon': -71.7333}, 'ESPINAR': {'lat': -14.8167, 'lon': -71.4167},
    'LA CONVENCIÓN': {'lat': -12.0333, 'lon': -72.3333}, 'PARURO': {'lat': -13.7333, 'lon': -71.8333},
    'PAUCARTAMBO': {'lat': -13.1333, 'lon': -71.5333}, 'QUISPICANCHI': {'lat': -13.8333, 'lon': -71.4167},
    'URUBAMBA': {'lat': -12.2667, 'lon': -72.1667}, 'PUNO': {'lat': -15.8402, 'lon': -70.0219},
    'ACORA': {'lat': -15.4167, 'lon': -69.9333}, 'AZÁNGARO': {'lat': -14.9167, 'lon': -70.1833},
    'CARABAYA': {'lat': -14.3667, 'lon': -69.5333}, 'CHUCUITO': {'lat': -15.9500, 'lon': -70.1167},
    'EL COLLAO': {'lat': -15.8167, 'lon': -70.1333}, 'HUANCANÉ': {'lat': -15.2167, 'lon': -69.7500},
    'LAMPA': {'lat': -15.3667, 'lon': -70.3667}, 'MELGAR': {'lat': -14.4083, 'lon': -70.3667},
    'MOHO': {'lat': -15.2500, 'lon': -69.7667}, 'SAN ANTONIO DE PUTINA': {'lat': -14.3167, 'lon': -70.0333},
    'SAN ROMÁN': {'lat': -15.5167, 'lon': -70.1333}, 'YUNGUYO': {'lat': -15.9500, 'lon': -69.6167},
    'AREQUIPA': {'lat': -16.4090, 'lon': -71.5375}, 'CAMANÁ': {'lat': -16.6167, 'lon': -72.7167},
    'CARAVELÍ': {'lat': -15.8333, 'lon': -74.8667}, 'CASTILLA': {'lat': -15.8167, 'lon': -70.4333},
    'CAYLLOMA': {'lat': -15.5333, 'lon': -71.3833}, 'CONDESUYOS': {'lat': -16.1833, 'lon': -72.5333},
    'ISLAY': {'lat': -17.0333, 'lon': -71.3667}, 'MOQUEGUA': {'lat': -17.1927, 'lon': -70.9352},
    'GENERAL SÁNCHEZ CERRO': {'lat': -17.1667, 'lon': -70.6833}, 'MARISCAL NIETO': {'lat': -17.6500, 'lon': -70.6500},
    'TACNA': {'lat': -18.0066, 'lon': -70.2463}, 'CANDARAVE': {'lat': -17.6667, 'lon': -69.3333},
    'JORGE BASADRE': {'lat': -17.8333, 'lon': -70.3333}, 'TARATA': {'lat': -17.4167, 'lon': -69.6667},
    'MAYNAS': {'lat': -3.7437, 'lon': -73.2516}, 'ALTO AMAZONAS': {'lat': -5.5000, 'lon': -76.5000},
    'LORETO': {'lat': -4.8500, 'lon': -75.0500}, 'MARISCAL RAMÓN CASTILLA': {'lat': -4.2333, 'lon': -71.5167},
    'REQUENA': {'lat': -5.1833, 'lon': -74.3667}, 'UCAYALI': {'lat': -6.6333, 'lon': -74.1167},
    'CORONEL PORTILLO': {'lat': -8.3791, 'lon': -74.5539}, 'ATALAYA': {'lat': -8.8833, 'lon': -73.7333},
    'PADRE ABAD': {'lat': -9.6667, 'lon': -75.3667}, 'TAMBOPATA': {'lat': -12.5933, 'lon': -69.1895},
    'MANU': {'lat': -11.9833, 'lon': -71.4167}, 'TAHUAMANU': {'lat': -10.9667, 'lon': -68.9333},
    'MOYOBAMBA': {'lat': -6.0346, 'lon': -76.9716}, 'EL DORADO': {'lat': -6.5333, 'lon': -76.1500},
    'HUALLAGA': {'lat': -6.9333, 'lon': -76.1667}, 'PICOTA': {'lat': -6.9167, 'lon': -76.3333},
    'RIOJA': {'lat': -6.1167, 'lon': -77.1667}, 'SAN MARTÍN': {'lat': -6.4845, 'lon': -76.3756},
    'TARAPOTO': {'lat': -6.4904, 'lon': -76.3656}, 'TOCACHE': {'lat': -6.8833, 'lon': -75.9667},
    'CHACHAPOYAS': {'lat': -6.2318, 'lon': -77.8690}, 'BAGUA': {'lat': -5.4500, 'lon': -78.5333},
    'BONGARÁ': {'lat': -6.2167, 'lon': -78.2333}, 'CONDORCANQUI': {'lat': -4.2167, 'lon': -77.8333},
    'LUYA': {'lat': -6.5667, 'lon': -77.8667}, 'RODRÍGUEZ DE MENDOZA': {'lat': -6.4333, 'lon': -77.4667},
    'UTCUBAMBA': {'lat': -6.4500, 'lon': -78.6000}, 'TUMBES': {'lat': -3.5669, 'lon': -80.4515},
    'CONTRALMIRANTE VILLAR': {'lat': -3.2500, 'lon': -80.3667}, 'ZARUMILLA': {'lat': -3.4167, 'lon': -80.2333},
}

def format_number(value, decimals=2):
    try:
        return float(f"{float(value):.{decimals}f}")
    except:
        return 0.0

def get_semaforo_rendimiento(rendimiento, meta):
    if pd.isna(rendimiento) or pd.isna(meta):
        return "CRÍTICO", "semaforo-critico"
    porcentaje = (rendimiento / meta) * 100
    if porcentaje >= 100:
        return "EXCELENTE", "semaforo-excelente"
    elif porcentaje >= 85:
        return "PRECAUCIÓN", "semaforo-alerta"
    else:
        return "CRÍTICO", "semaforo-critico"


COLORES_SEMAFORO_FILA = {
    "EXCELENTE": {"background": "#16a34a", "color": "#ffffff", "badge": "#22c55e"},
    "PRECAUCIÓN": {"background": "#eab308", "color": "#1a1a1a", "badge": "#facc15"},
    "ALERTA": {"background": "#eab308", "color": "#1a1a1a", "badge": "#facc15"},
    "CRÍTICO": {"background": "#dc2626", "color": "#ffffff", "badge": "#ef4444"},
}


def _texto_sin_acentos(s):
    if s is None:
        return ""
    t = unicodedata.normalize("NFD", str(s))
    return "".join(c for c in t if unicodedata.category(c) != "Mn").upper()


def codigo_ceco_a_str(val):
    """Normaliza código CECO (1001.0 → '1001')."""
    if pd.isna(val) or val is None:
        return None
    try:
        f = float(val)
        if f == int(f):
            return str(int(f))
        return str(f).rstrip("0").rstrip(".")
    except (ValueError, TypeError):
        s = str(val).strip().upper()
        return s if s and s not in ("NAN", "N/A", "") else None


def _es_fila_encabezado_tmceco(cod, nom):
    n = _texto_sin_acentos(nom)
    if not cod and not n:
        return True
    if n in ("NAN", "N/A"):
        return True
    if "AREA DEL CECO" in n or "CENTRO DE COSTO" in n:
        return True
    if "CECO" in n and ("AREA" in n or "CENTRO" in n):
        return True
    return False


def cargar_mapa_ceco(xl):
    """Hoja TMCeco: código (col. CECO / Centro de costo) → nombre del área."""
    mapa = {}
    hoja = _resolver_hoja(xl, "TMCeco")
    if not hoja:
        return mapa
    try:
        df = pd.read_excel(xl, sheet_name=hoja, engine="openpyxl")
    except Exception:
        return mapa
    if df.empty or df.shape[1] < 2:
        return mapa
    col_cod = buscar_columna(df, "CECO", "CENTRO DE COSTO", "CENTRODECOSTO") or df.columns[0]
    otras = [c for c in df.columns if c != col_cod]
    col_nom = buscar_columna(df, "AREA", "NOMBRE", "DESCRIPCION")
    if not col_nom or col_nom == col_cod:
        col_nom = otras[0] if otras else df.columns[1]
    for _, row in df.iterrows():
        cod = codigo_ceco_a_str(row[col_cod])
        nom = str(row[col_nom]).strip().upper()
        if _es_fila_encabezado_tmceco(cod, nom):
            continue
        if not cod or not nom:
            continue
        mapa[cod] = nom
    return mapa


def resolver_nombre_ceco(val, mapa_ceco):
    """Código → nombre (TMCeco). Vacío, 0 o sin mapeo → etiqueta visible."""
    if pd.isna(val) or val is None:
        return "SIN CECO ASIGNADO"
    s_raw = str(val).strip()
    if s_raw.upper() in ("", "NAN", "N/A", "NONE", "NULL", "0"):
        return "SIN CECO ASIGNADO"
    cod = codigo_ceco_a_str(val)
    if cod in ("0", "00"):
        return "SIN CECO ASIGNADO"
    if mapa_ceco:
        if cod and cod in mapa_ceco:
            return mapa_ceco[cod]
        s = s_raw.upper()
        if s in mapa_ceco.values():
            return s
        if cod:
            return f"CECO NO MAPEADO (cód. {cod})"
        return f"CECO NO MAPEADO ({s})"
    if cod:
        return f"CECO {cod} (sin catálogo TMCeco)"
    return s_raw.upper()


def aplicar_nombres_ceco_df(df, mapa_ceco):
    if df is None or df.empty or "CECO" not in df.columns:
        return df
    out = df.copy()
    mapa = mapa_ceco or {}
    out["CECO"] = out["CECO"].apply(lambda v: resolver_nombre_ceco(v, mapa))
    return out


def estilo_fila_semaforo(row):
    """Toda la fila con color sólido del semáforo."""
    est = row.get("ESTADO") or row.get("_ESTADO_TXT") or "CRÍTICO"
    colores = COLORES_SEMAFORO_FILA.get(est, COLORES_SEMAFORO_FILA["CRÍTICO"])
    css = (
        f"background-color: {colores['background']} !important; "
        f"color: {colores['color']} !important; "
        f"font-weight: 600 !important;"
    )
    return [css] * len(row)


def estilo_celda_estado(val):
    """Columna ESTADO resaltada sobre la fila."""
    colores = COLORES_SEMAFORO_FILA.get(str(val), COLORES_SEMAFORO_FILA["CRÍTICO"])
    return (
        f"background-color: {colores['badge']} !important; "
        f"color: #000000 !important; "
        f"font-weight: 900 !important; text-align: center !important; "
        f"border: 3px solid #ffffff !important; "
        f"border-radius: 6px; font-size: 13px; letter-spacing: 0.5px;"
    )


def tipo_proveedor_combustible(valor):
    """Clasifica registro como PRIMAX, REDCOL u OTRO."""
    if pd.isna(valor):
        return "OTRO"
    s = str(valor).upper()
    if "PRIMAX" in s:
        return "PRIMAX"
    if "REDCOL" in s:
        return "REDCOL"
    return "OTRO"


def acortar_nombre_conductor(nombre):
    """Nombre + inicial del primer apellido (ej. CARLOS RODRIGUEZ -> Carlos R.)."""
    if nombre is None or (isinstance(nombre, float) and pd.isna(nombre)):
        return "Sin nombre"
    texto = str(nombre).strip()
    if not texto or texto.upper() in ("NAN", "N/A", ""):
        return "Sin nombre"
    partes = [p for p in texto.split() if p]
    if len(partes) == 1:
        return partes[0].title()
    return f"{partes[0].title()} {partes[1][0].upper()}."


def serie_display(df, col_display):
    """Etiqueta unificada (placa o nombre corto) según selector del sidebar."""
    if col_display == "CONDUCTOR" and "CONDUCTOR" in df.columns:
        return df["CONDUCTOR"].apply(acortar_nombre_conductor)
    if "PLACA" in df.columns:
        return df["PLACA"].astype(str)
    return pd.Series(["—"] * len(df), index=df.index)


def etiqueta_columna_id(col_display):
    return "NOMBRE" if col_display == "CONDUCTOR" else "PLACA"


def color_heatmap_galones(val, vmax):
    """Semáforo volumen: alto=verde, medio=amarillo, bajo=rojo."""
    if val is None or (isinstance(val, float) and (pd.isna(val) or val <= 0)):
        return "background:#0d0d0d;color:#555555;"
    if vmax <= 0:
        return "background:#1a1a1a;color:#ffffff;"
    t = min(1.0, max(0.0, float(val) / float(vmax)))
    if t >= 0.55:
        bg, fg = "#16a34a", "#ffffff"
    elif t >= 0.2:
        bg, fg = "#eab308", "#111111"
    else:
        bg, fg = "#dc2626", "#ffffff"
    return (
        f"background:{bg};color:{fg};font-weight:700;text-align:right;"
        f"padding:7px 10px;font-size:12px;"
    )


def color_heatmap_participacion(pct):
    return color_heatmap_galones(pct, 100.0)


def precio_promedio_visitas(df, gal_col, precio_col):
    """Promedio del precio/galón en cada abastecimiento (monto÷galones por visita)."""
    if df is None or df.empty or not gal_col:
        return np.nan
    g = to_numeric_locale(df[gal_col])
    m = calcular_monto_fila(df, gal_col, precio_col)
    mask = (g > 0) & (m > 0)
    if not mask.any():
        return np.nan
    return float((m[mask] / g[mask]).mean())


def filtrar_registro_redes(df_reg, red_col, red_tipo="REDCOL", ano=None, mes_corto=None):
    """Filtra registro por año, mes fiscal y red (PRIMAX / REDCOL / ambas)."""
    if df_reg is None or df_reg.empty:
        return pd.DataFrame()
    df = df_reg.copy()
    if ano is not None and "ANO" in df.columns:
        df = df[pd.to_numeric(df["ANO"], errors="coerce") == ano]
    if mes_corto and "MES_NUM" in df.columns:
        mes_num = MES_A_NUM.get(mes_corto)
        if mes_num is not None:
            df = df[df["MES_NUM"] == mes_num]
    if red_col and red_col in df.columns:
        if red_tipo == "PRIMAX":
            df = df[df[red_col].astype(str).str.upper().str.contains("PRIMAX", na=False)]
        elif red_tipo == "REDCOL":
            df = df[df[red_col].astype(str).str.upper().str.contains("REDCOL", na=False)]
        elif red_tipo == "AMBAS":
            tipos = df[red_col].apply(tipo_proveedor_combustible)
            df = df[tipos.isin(["PRIMAX", "REDCOL"])]
    return df


def meses_hasta(mes_corto, meses_disponibles=None):
    """Meses del año hasta el mes seleccionado (inclusive)."""
    base = meses_disponibles if meses_disponibles else ORDEN_MESES
    if mes_corto not in base:
        return list(base)
    idx = base.index(mes_corto) if mes_corto in ORDEN_MESES else len(base) - 1
    orden = [m for m in ORDEN_MESES if m in base]
    if mes_corto in orden:
        return orden[: orden.index(mes_corto) + 1]
    return list(base)


def meses_en_registro(df_reg, ano=2026, red_col=None, red_tipo="AMBAS"):
    if df_reg is None or df_reg.empty or "MES_NUM" not in df_reg.columns:
        return []
    df = filtrar_registro_redes(df_reg, red_col, red_tipo=red_tipo, ano=ano, mes_corto=None)
    nums = sorted(pd.to_numeric(df["MES_NUM"], errors="coerce").dropna().unique())
    return [m for m in ORDEN_MESES if MES_A_NUM.get(m) in [int(x) for x in nums]]


def calcular_kpis_estaciones(df_reg, red_col, gal_col, estacion_col, prov_col):
    """KPIs coherentes: estaciones únicas vs abastecimientos (registros) por red."""
    if df_reg is None or df_reg.empty:
        return 0, 0.0, 0, 0, 0, 0
    n_abast = len(df_reg)
    if estacion_col and estacion_col in df_reg.columns:
        est = df_reg[estacion_col].astype(str).str.strip()
        est = est[~est.isin(["", "NAN", "NONE"])]
        n_estaciones = int(est.nunique())
    else:
        n_estaciones = n_abast
    if gal_col and gal_col in df_reg.columns:
        total_gal = float(to_numeric_locale(df_reg[gal_col]).sum())
    else:
        total_gal = 0.0
    n_primax = contar_proveedor(df_reg[red_col], "PRIMAX") if red_col and red_col in df_reg.columns else 0
    n_redcol = contar_proveedor(df_reg[red_col], "REDCOL") if red_col and red_col in df_reg.columns else 0
    if prov_col and prov_col in df_reg.columns:
        n_prov = int(df_reg[prov_col].astype(str).nunique())
    else:
        n_prov = 0
    return n_estaciones, total_gal, n_primax, n_redcol, n_prov, n_abast


def construir_matriz_departamento(
    df_reg, prov_col, gal_col, precio_col, red_col, ano=2026, red_tipo="REDCOL", mes_corto=None
):
    df = filtrar_registro_redes(df_reg, red_col, red_tipo=red_tipo, ano=ano, mes_corto=None)
    if df.empty or not prov_col or not gal_col:
        return pd.DataFrame(), pd.DataFrame(), [], False, None, {}, {}
    df = df.copy()
    df["PROV_KEY"] = df[prov_col].apply(etiqueta_provincia_tabla)
    df = df[df["PROV_KEY"].notna()]
    df[gal_col] = to_numeric_locale(df[gal_col])
    df["MES_CORTO"] = pd.to_numeric(df.get("MES_NUM", pd.Series()), errors="coerce").map(
        lambda x: NUM_A_MES.get(int(x), "") if pd.notna(x) else ""
    )
    df = df[df["MES_CORTO"] != ""]
    meses_disp = meses_en_registro(df_reg, ano, red_col, red_tipo)
    if mes_corto:
        meses_cols = meses_hasta(mes_corto, meses_disp)
    else:
        meses_cols = meses_disp
    if not meses_cols:
        meses_cols = sorted(
            [m for m in df["MES_CORTO"].unique() if m in ORDEN_MESES],
            key=lambda m: ORDEN_MESES.index(m),
        )
    totales_verif = {m: float(df.loc[df["MES_CORTO"] == m, gal_col].sum()) for m in meses_cols}

    pivot_gal = df.groupby(["PROV_KEY", "MES_CORTO"])[gal_col].sum().unstack(fill_value=0)
    for m in meses_cols:
        if m not in pivot_gal.columns:
            pivot_gal[m] = 0.0
    pivot_gal = pivot_gal[[m for m in meses_cols if m in pivot_gal.columns]]
    pivot_gal["TOTAL"] = pivot_gal.sum(axis=1)
    total_gen_gal = float(pivot_gal["TOTAL"].sum())
    pivot_gal["PART_%"] = (pivot_gal["TOTAL"] / max(total_gen_gal, 1) * 100).round(2)
    pivot_gal = pivot_gal.sort_values("TOTAL", ascending=False)

    pivot_sol = pd.DataFrame()
    tiene_soles = False
    col_monto = precio_col if precio_col else resolver_columna_monto(df)
    df["_MONTO"] = calcular_monto_fila(df, gal_col, col_monto)
    totales_verif_sol = {}
    if float(df["_MONTO"].sum()) > 0:
        totales_verif_sol = {
            m: float(df.loc[df["MES_CORTO"] == m, "_MONTO"].sum()) for m in meses_cols
        }
        pivot_sol = df.groupby(["PROV_KEY", "MES_CORTO"])["_MONTO"].sum().unstack(fill_value=0)
        for m in meses_cols:
            if m not in pivot_sol.columns:
                pivot_sol[m] = 0.0
        pivot_sol = pivot_sol[[m for m in meses_cols if m in pivot_sol.columns]]
        pivot_sol["TOTAL"] = pivot_sol.sum(axis=1)
        total_gen_sol = float(pivot_sol["TOTAL"].sum())
        pivot_sol["PART_%"] = (pivot_sol["TOTAL"] / max(total_gen_sol, 1) * 100).round(2)
        pivot_sol = pivot_sol.reindex(pivot_gal.index).fillna(0)
        tiene_soles = total_gen_sol > 0

    totales_verif["_TOTAL"] = sum(totales_verif.get(m, 0.0) for m in meses_cols)
    if totales_verif_sol:
        totales_verif_sol["_TOTAL"] = sum(totales_verif_sol.get(m, 0.0) for m in meses_cols)

    return pivot_gal, pivot_sol, meses_cols, tiene_soles, col_monto, totales_verif, totales_verif_sol


def render_selector_red_departamento():
    """Botones PRIMAX / REDCOL para tabla por departamento."""
    if "dep_red_tipo" not in st.session_state:
        st.session_state.dep_red_tipo = "PRIMAX"
    b1, b2 = st.columns(2)
    with b1:
        if st.button(
            "PRIMAX",
            key="btn_dep_red_primax",
            use_container_width=True,
            type="primary" if st.session_state.dep_red_tipo == "PRIMAX" else "secondary",
        ):
            st.session_state.dep_red_tipo = "PRIMAX"
            st.rerun()
    with b2:
        if st.button(
            "REDCOL",
            key="btn_dep_red_redcol",
            use_container_width=True,
            type="primary" if st.session_state.dep_red_tipo == "REDCOL" else "secondary",
        ):
            st.session_state.dep_red_tipo = "REDCOL"
            st.rerun()
    return st.session_state.dep_red_tipo


def render_tabla_abastecimiento_departamento(
    df_reg, prov_col, gal_col, precio_col, red_col, simbolo, ano=2026, mes_corto=None, mes_label=""
):
    st.markdown('<div class="section-title">Abastecimiento por departamento</div>', unsafe_allow_html=True)
    mes_txt = mes_label or (MESES_COMPLETOS.get(mes_corto, mes_corto) if mes_corto else "año completo")

    t1, t2, t3 = st.columns([1.1, 1.1, 1.3])
    with t1:
        st.markdown('<span class="toolbar-label">Red de estaciones</span>', unsafe_allow_html=True)
        red_tipo = render_selector_red_departamento()
    with t2:
        st.markdown('<span class="toolbar-label">Unidad</span>', unsafe_allow_html=True)
        _preview = filtrar_registro_redes(df_reg, red_col, red_tipo=red_tipo, ano=ano, mes_corto=None)
        _col_monto = resolver_columna_monto(_preview) or precio_col
        _tiene_preview = (
            _col_monto
            and gal_col
            and float(calcular_monto_fila(_preview, gal_col, _col_monto).sum()) > 0
        )
        if _tiene_preview:
            unidad = st.radio(
                "Unidad",
                ["Galones (GL)", f"Soles ({simbolo})"],
                horizontal=True,
                key="dep_unidad_tabla",
                label_visibility="collapsed",
            )
        else:
            unidad = "Galones (GL)"
            st.caption("Soles: use columna SUBTOTAL o TOTAL en hoja registro")
    with t3:
        st.caption(f"Acumulado ene.–**{mes_txt}** {ano}. Mes fiscal en panel lateral.")

    matriz_gal, matriz_sol, meses_cols, tiene_soles, col_monto_usado, totales_verif, totales_verif_sol = (
        construir_matriz_departamento(
            df_reg, prov_col, gal_col, precio_col, red_col, ano=ano, red_tipo=red_tipo, mes_corto=mes_corto
        )
    )
    ver_soles = tiene_soles and unidad.startswith("Soles")
    matriz = matriz_sol if ver_soles else matriz_gal
    totales_pie = totales_verif_sol if ver_soles and totales_verif_sol else totales_verif
    if ver_soles and matriz.empty:
        aviso_amigable(
            f"Sin montos en soles para {red_tipo}. "
            f"Columna detectada: **{col_monto_usado or 'ninguna'}**."
        )
        return
    unidad_lbl = f"SOLES ({simbolo})" if ver_soles else "GALONES"

    st.markdown(
        f'<div class="titulo-anual">ABASTECIMIENTO POR DEPARTAMENTO {red_tipo} — {ano}</div>',
        unsafe_allow_html=True,
    )
    if matriz.empty:
        aviso_amigable(f"Sin datos {red_tipo} por departamento ({mes_txt}, {ano}).")
        return
    vmax = float(matriz[meses_cols].max().max()) if meses_cols else 1.0
    headers = (
        ["PROVINCIA"]
        + [MESES_COMPLETOS.get(m, m) for m in meses_cols]
        + ["TOTAL", "PARTICIPACIÓN"]
    )
    thead = "".join(f"<th>{html.escape(h)}</th>" for h in headers)

    def fmt_celda(val):
        if val <= 0:
            return ""
        if ver_soles:
            return f"{simbolo} {format_number(val, 2):,.2f}"
        return f"{format_number(val, 2):,.2f}"

    filas = []
    for prov, fila in matriz.iterrows():
        celdas = [f'<td class="col-prov">{html.escape(str(prov))}</td>']
        for m in meses_cols:
            val = float(fila.get(m, 0))
            celdas.append(
                f'<td style="{color_heatmap_galones(val, vmax)}">{fmt_celda(val)}</td>'
            )
        total_v = float(fila["TOTAL"])
        part_v = float(fila["PART_%"])
        celdas.append(f'<td style="{color_heatmap_galones(total_v, vmax)}">{fmt_celda(total_v)}</td>')
        celdas.append(f'<td style="{color_heatmap_participacion(part_v)}">{part_v:.2f}%</td>')
        filas.append(f"<tr>{''.join(celdas)}</tr>")

    foot = [f'<td class="col-prov">{unidad_lbl}</td>']
    for m in meses_cols:
        s = float(totales_pie.get(m, matriz[m].sum() if m in matriz.columns else 0.0))
        foot.append(f'<td style="{color_heatmap_galones(s, vmax)}">{fmt_celda(s)}</td>')
    t_total = float(
        totales_pie.get("_TOTAL")
        if "_TOTAL" in totales_pie
        else sum(float(totales_pie.get(m, 0.0)) for m in meses_cols)
    )
    if t_total <= 0:
        t_total = float(matriz["TOTAL"].sum())
    foot.append(f'<td style="background:#111;color:#fff;font-weight:900;">{fmt_celda(t_total)}</td>')
    foot.append('<td style="background:#111;color:#fff;font-weight:900;">100.00%</td>')

    st.markdown(
        f'<div class="combustible-table"><div class="tabla-heatmap-scroll">'
        f'<table class="tabla-heatmap-dep">'
        f"<thead><tr>{thead}</tr></thead><tbody>{''.join(filas)}</tbody>"
        f"<tfoot><tr>{''.join(foot)}</tr></tfoot></table></div></div>",
        unsafe_allow_html=True,
    )
    st.caption(
        "Semáforo por volumen: **verde** alto, **amarillo** medio, **rojo** bajo. "
        + (
            f"Soles = suma de **{col_monto_usado or precio_col}** por provincia y mes (registro)."
            if ver_soles and (col_monto_usado or precio_col)
            else (
                f"Totales del pie = suma directa del registro ({red_tipo}). "
                "Incluye **todas** las provincias del Excel (Callao, Lima, etc.)."
            )
        )
    )



def render_grafico_km_costo_km(df_dashboard, col_display, c_km, c_sub, factor, simbolo):
    """Costo por KM — cuánto cuesta cada kilómetro recorrido."""
    if df_dashboard is None or df_dashboard.empty:
        return
    st.markdown(
        '<div class="section-title">¿Cuánto cuesta cada kilómetro?</div>',
        unsafe_allow_html=True,
    )
    df_g = df_dashboard.copy()
    df_g["DISPLAY"] = serie_display(df_g, col_display)
    df_g["KM_VAL"] = pd.to_numeric(df_g[c_km], errors="coerce").fillna(0)
    df_g["COSTO_KM"] = np.where(
        df_g["KM_VAL"] > 0,
        (pd.to_numeric(df_g[c_sub], errors="coerce").fillna(0) * factor) / df_g["KM_VAL"],
        0.0,
    )
    df_g = df_g[df_g["KM_VAL"] > 0].copy()
    if df_g.empty:
        aviso_amigable("Sin kilómetros para calcular costo/KM.")
        return
    gasto_total = float(df_g[c_sub].sum()) * factor
    km_total = float(df_g["KM_VAL"].sum())
    prom_flota = gasto_total / max(km_total, 1)

    def color_bar_costo(v):
        if v <= prom_flota * 0.9:
            return "#16a34a"
        if v <= prom_flota * 1.1:
            return "#eab308"
        return "#dc2626"

    df_g = df_g.sort_values("COSTO_KM", ascending=True)
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            y=df_g["DISPLAY"],
            x=df_g["COSTO_KM"],
            orientation="h",
            marker_color=[color_bar_costo(v) for v in df_g["COSTO_KM"]],
            text=[f"{simbolo} {v:,.2f}" for v in df_g["COSTO_KM"]],
            textposition="outside",
            textfont=dict(color="#ffffff", size=11),
            hovertemplate="%{y}<br>" + simbolo + "/KM: %{x:,.2f}<extra></extra>",
        )
    )
    fig.add_vline(
        x=prom_flota,
        line_dash="dash",
        line_color="#ffd700",
        line_width=2,
        annotation_text=f"Promedio flota {simbolo} {prom_flota:,.2f}/KM",
        annotation_position="top",
        annotation_font_color="#ffd700",
    )
    fig.update_layout(
        height=max(380, len(df_g) * 28),
        template="plotly_dark",
        paper_bgcolor="#000000",
        plot_bgcolor="#000000",
        margin=dict(l=120, r=90, t=50, b=40),
        xaxis_title=f"Costo por kilómetro ({simbolo})",
        yaxis_title="",
        showlegend=False,
    )
    fig.update_xaxes(showgrid=True, gridcolor="#1a1a1a")
    fig.update_yaxes(showgrid=False)
    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CFG)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.caption(f"**Promedio flota:** {simbolo} {format_number(prom_flota, 2):,.2f} / KM")
    with c2:
        st.caption(f"**KM totales:** {format_number(km_total, 2):,.2f}")
    with c3:
        st.caption(f"**Gasto combustible:** {simbolo} {format_number(gasto_total, 2):,.2f}")
    with st.expander("Ver kilómetros recorridos"):
        df_km = df_g.sort_values("KM_VAL", ascending=False)
        fig_km = px.bar(df_km, x="DISPLAY", y="KM_VAL", template="plotly_dark", color_discrete_sequence=["#14B8A6"])
        fig_km.update_layout(height=280, paper_bgcolor="#000", plot_bgcolor="#000", yaxis_title="KM")
        fig_km.update_xaxes(tickangle=-35)
        st.plotly_chart(fig_km, use_container_width=True, config=PLOTLY_CFG)



def construir_tabla_estaciones(df_reg, estacion_col, red_col, gal_col, precio_col=None, red_tipo=None):
    """Galones y soles por estación, comparando PRIMAX vs REDCOL."""
    if df_reg is None or df_reg.empty or not estacion_col or not red_col or not gal_col:
        return pd.DataFrame()
    df = df_reg.copy()
    df[gal_col] = to_numeric_locale(df[gal_col])
    col_m = precio_col if precio_col else resolver_columna_monto(df)
    df["_MONTO"] = calcular_monto_fila(df, gal_col, col_m)
    df["_RED"] = df[red_col].apply(tipo_proveedor_combustible)
    df = df[df["_RED"].isin(["PRIMAX", "REDCOL"])]
    if red_tipo in ("PRIMAX", "REDCOL"):
        df = df[df["_RED"] == red_tipo]
    if df.empty:
        return pd.DataFrame()
    filas = []
    for est in lista_unica_texto(df[estacion_col]):
        if not est or est.upper() in ("NAN", ""):
            continue
        sub = df[df[estacion_col].astype(str) == est]
        p_mask = sub["_RED"] == "PRIMAX"
        r_mask = sub["_RED"] == "REDCOL"
        p_gal = float(sub.loc[p_mask, gal_col].sum())
        r_gal = float(sub.loc[r_mask, gal_col].sum())
        p_sol = float(sub.loc[p_mask, "_MONTO"].sum())
        r_sol = float(sub.loc[r_mask, "_MONTO"].sum())
        t_gal = p_gal + r_gal
        t_sol = p_sol + r_sol
        precio_prom = precio_promedio_visitas(sub, gal_col, precio_col)
        filas.append(
            {
                "ESTACIÓN": est,
                "CARGAS": len(sub),
                "PRIMAX_GL": p_gal,
                "REDCOL_GL": r_gal,
                "TOTAL_GL": t_gal,
                "PRIMAX_S": p_sol,
                "REDCOL_S": r_sol,
                "TOTAL_S": t_sol,
                "PRECIO_PROM_S_GL": precio_prom,
            }
        )
    return pd.DataFrame(filas).sort_values("TOTAL_GL", ascending=False)


def _radio_unidad_gal_soles(simbolo, tiene_soles, key):
    opciones = ["Galones (GL)"]
    if tiene_soles:
        opciones.append(f"Soles ({simbolo})")
    return st.radio("Ver en", opciones, horizontal=True, key=key, label_visibility="collapsed")


def _formatear_tabla_red_dual(df_tab, modo, simbolo, id_col):
    """Muestra PRIMAX y REDCOL en galones o soles."""
    cols_gal = [id_col, "CARGAS", "PRIMAX_GL", "REDCOL_GL", "TOTAL_GL"]
    cols_sol = [id_col, "CARGAS", "PRIMAX_S", "REDCOL_S", "TOTAL_S"]
    if modo.startswith("Galones"):
        show = df_tab[[c for c in cols_gal if c in df_tab.columns]].copy()
        ren = {
            "CARGAS": "# CARGAS",
            "PRIMAX_GL": "PRIMAX (GL)",
            "REDCOL_GL": "REDCOL (GL)",
            "TOTAL_GL": "TOTAL (GL)",
        }
        show = show.rename(columns={k: v for k, v in ren.items() if k in show.columns})
        for c in show.columns:
            if c not in (id_col, "# CARGAS"):
                show[c] = show[c].apply(lambda x: f"{format_number(x, 2):,.2f}")
    else:
        show = df_tab[[c for c in cols_sol if c in df_tab.columns]].copy()
        ren = {
            "CARGAS": "# CARGAS",
            "PRIMAX_S": f"PRIMAX ({simbolo})",
            "REDCOL_S": f"REDCOL ({simbolo})",
            "TOTAL_S": f"TOTAL ({simbolo})",
        }
        show = show.rename(columns={k: v for k, v in ren.items() if k in show.columns})
        for c in show.columns:
            if c not in (id_col, "# CARGAS"):
                show[c] = show[c].apply(lambda x: f"{simbolo} {format_number(x, 2):,.2f}")
    return show


def construir_tabla_provincia(df_reg, prov_col, gal_col, precio_col, red_col):
    """PRIMAX y REDCOL por provincia (galones y soles)."""
    if df_reg is None or df_reg.empty or not prov_col or not gal_col or not red_col:
        return pd.DataFrame()
    df = df_reg.copy()
    df[gal_col] = to_numeric_locale(df[gal_col])
    col_m = precio_col if precio_col else resolver_columna_monto(df)
    df["_MONTO"] = calcular_monto_fila(df, gal_col, col_m)
    df["_RED"] = df[red_col].apply(tipo_proveedor_combustible)
    df = df[df["_RED"].isin(["PRIMAX", "REDCOL"])]
    if df.empty:
        return pd.DataFrame()
    filas = []
    for prov in lista_unica_texto(df[prov_col]):
        if not prov or str(prov).upper() in ("NAN", ""):
            continue
        sub = df[df[prov_col].astype(str) == prov]
        p_mask = sub["_RED"] == "PRIMAX"
        r_mask = sub["_RED"] == "REDCOL"
        p_gal = float(sub.loc[p_mask, gal_col].sum())
        r_gal = float(sub.loc[r_mask, gal_col].sum())
        p_sol = float(sub.loc[p_mask, "_MONTO"].sum())
        r_sol = float(sub.loc[r_mask, "_MONTO"].sum())
        filas.append(
            {
                "PROVINCIA": prov,
                "CARGAS": len(sub),
                "PRIMAX_GL": p_gal,
                "REDCOL_GL": r_gal,
                "TOTAL_GL": p_gal + r_gal,
                "PRIMAX_S": p_sol,
                "REDCOL_S": r_sol,
                "TOTAL_S": p_sol + r_sol,
                "PART_%": 0.0,
            }
        )
    out = pd.DataFrame(filas)
    if out.empty:
        return out
    tot = float(out["TOTAL_GL"].sum())
    out["PART_%"] = (out["TOTAL_GL"] / max(tot, 1) * 100).round(2)
    return out.sort_values("TOTAL_GL", ascending=False)


def render_tabla_provincia_red(df_reg, prov_col, gal_col, precio_col, red_col, simbolo):
    st.markdown(
        '<div class="section-title">Análisis de combustible por provincia — PRIMAX y REDCOL</div>',
        unsafe_allow_html=True,
    )
    df_tab = construir_tabla_provincia(df_reg, prov_col, gal_col, precio_col, red_col)
    if df_tab.empty:
        aviso_amigable("Sin datos PRIMAX/REDCOL por provincia en el mes.")
        return
    tiene_soles = float(df_tab["TOTAL_S"].sum()) > 0
    st.markdown('<span class="toolbar-label">Unidad</span>', unsafe_allow_html=True)
    modo = _radio_unidad_gal_soles(simbolo, tiene_soles, "red_prov_unidad")
    show = _formatear_tabla_red_dual(df_tab, modo, simbolo, "PROVINCIA")
    show["Participación %"] = df_tab["PART_%"].apply(lambda x: f"{x:.2f}%")
    st.caption(
        "**Atenciones** = cantidad de abastecimientos en esa provincia (mes fiscal del panel). "
        "Soles = suma SUB TOTAL/TOTAL del registro."
    )
    if not tiene_soles:
        st.caption("Soles: use columna SUB TOTAL o TOTAL en hoja registro.")
    st.dataframe(show, use_container_width=True, hide_index=True, height=420)


def render_tabla_estaciones_red(df_reg, estacion_col, red_col, gal_col, precio_col, simbolo, red_tipo=None):
    """Tabla por estación: siempre PRIMAX y REDCOL; galones o soles."""
    st.markdown(
        '<div class="section-title">Resumen por estación — PRIMAX y REDCOL</div>',
        unsafe_allow_html=True,
    )
    df_tab = construir_tabla_estaciones(
        df_reg, estacion_col, red_col, gal_col, precio_col, red_tipo=None
    )
    if df_tab.empty:
        aviso_amigable("Sin abastecimientos PRIMAX/REDCOL por estación en el mes.")
        return
    tiene_soles = float(df_tab["TOTAL_S"].sum()) > 0
    st.markdown('<span class="toolbar-label">Unidad</span>', unsafe_allow_html=True)
    modo = _radio_unidad_gal_soles(simbolo, tiene_soles, "red_est_unidad")
    show = _formatear_tabla_red_dual(df_tab, modo, simbolo, "ESTACIÓN")
    if not tiene_soles:
        st.caption(
            "Soles no disponibles: agregue columna de importe/monto en la hoja **registro**."
        )
    st.dataframe(show, use_container_width=True, hide_index=True, height=420)


def render_selector_graficos(df_raw):
    """Botones Placa / Conductor."""
    tiene_conductor = "CONDUCTOR" in df_raw.columns
    st.markdown(
        '<div class="sidebar-simple-label">Mostrar por (gráficos, comparativo y auditoría)</div>',
        unsafe_allow_html=True,
    )
    if "mostrar_grafico_col" not in st.session_state:
        st.session_state.mostrar_grafico_col = "PLACA"
    if tiene_conductor:
        b1, b2 = st.columns(2)
        with b1:
            if st.button(
                "Placa",
                key="btn_graf_placa",
                use_container_width=True,
                type="primary" if st.session_state.mostrar_grafico_col == "PLACA" else "secondary",
            ):
                st.session_state.mostrar_grafico_col = "PLACA"
                st.rerun()
        with b2:
            if st.button(
                "Conductor",
                key="btn_graf_conductor",
                use_container_width=True,
                type="primary" if st.session_state.mostrar_grafico_col == "CONDUCTOR" else "secondary",
            ):
                st.session_state.mostrar_grafico_col = "CONDUCTOR"
                st.rerun()
    else:
        st.session_state.mostrar_grafico_col = "PLACA"
    return st.session_state.mostrar_grafico_col


def render_slider_meta(vmin=4.0, vmax=18.0, default=11.5, key="slider_meta_kmg"):
    """Meta KM/G: barra arrastrable (st.slider)."""
    st.markdown('<div class="sidebar-simple-label">Meta KM/G</div>', unsafe_allow_html=True)
    cur = float(st.session_state.get(key, default))
    meta = st.slider(
        "Meta KM/G",
        min_value=float(vmin),
        max_value=float(vmax),
        value=cur,
        step=0.01,
        label_visibility="collapsed",
        key=key,
    )
    st.markdown(
        f'<div class="meta-kmg-labels">'
        f'<span>{vmin:.2f}</span><span class="val">{meta:.2f}</span><span>{vmax:.2f}</span></div>',
        unsafe_allow_html=True,
    )
    return meta


def render_selector_vista():
    """Selector de vista moderno (botones en lugar de radio antiguo)."""
    opciones = VISTAS_MENU
    mv = st.session_state.get("menu_vista")
    if mv in _ALIASES_VISTA:
        st.session_state.menu_vista = _ALIASES_VISTA[mv]
    if "menu_vista" not in st.session_state or st.session_state.menu_vista not in opciones:
        st.session_state.menu_vista = opciones[0]
    st.markdown('<div class="sidebar-vista-label">Seleccione Vista</div>', unsafe_allow_html=True)
    for op in opciones:
        activo = st.session_state.menu_vista == op
        if st.button(
            op,
            key=f"btn_vista_{op}",
            use_container_width=True,
            type="primary" if activo else "secondary",
        ):
            st.session_state.menu_vista = op
            st.rerun()
    return st.session_state.menu_vista


def construir_excel_auditoria(df_a, c_km, c_gal, c_ren, c_sub, meta_kmg, col_display, factor=1.0):
    """Genera .xlsx con valores numéricos (aptos para análisis en Excel)."""
    col_id = etiqueta_columna_id(col_display)
    filas = {col_id: serie_display(df_a, col_display)}
    if "CECO" in df_a.columns:
        filas["CECO"] = df_a["CECO"].astype(str)
    for c in ("TIPO", "MARCA", "MODELO", "CONDUCTOR"):
        if c in df_a.columns:
            filas[c] = df_a[c]
    filas["KM"] = pd.to_numeric(df_a[c_km], errors="coerce")
    filas["GALONES"] = pd.to_numeric(df_a[c_gal], errors="coerce")
    filas["RENDIMIENTO"] = pd.to_numeric(df_a[c_ren], errors="coerce")
    filas["COSTO"] = pd.to_numeric(df_a[c_sub], errors="coerce") * factor
    filas["ESTADO"] = df_a[c_ren].apply(lambda x: get_semaforo_rendimiento(x, meta_kmg)[0])
    df_x = pd.DataFrame(filas)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_x.to_excel(writer, index=False, sheet_name="Auditoria")
    buf.seek(0)
    return buf.getvalue()


def render_tabla_auditoria(
    df_dashboard,
    c_km,
    c_gal,
    c_ren,
    c_sub,
    meta_kmg,
    simbolo,
    factor,
    col_display="PLACA",
    mes_label="",
):
    """Tabla de auditoría con semáforo por fila y orden por rendimiento."""
    if df_dashboard is None or df_dashboard.empty:
        aviso_amigable()
        return
    c1, c2, c3 = st.columns([1.4, 1, 0.55])
    with c1:
        st.markdown('<div class="section-title">Auditoría Completa - Tabla de Control</div>', unsafe_allow_html=True)
    with c2:
        orden_rend = st.selectbox(
            "Ordenar rendimiento",
            ["Mayor a menor", "Menor a mayor", "Por placa (A-Z)"],
            key="orden_audit_rend",
        )

    df_a = df_dashboard.copy()
    if orden_rend == "Mayor a menor":
        df_a = df_a.sort_values(c_ren, ascending=False)
    elif orden_rend == "Menor a mayor":
        df_a = df_a.sort_values(c_ren, ascending=True)
    else:
        df_a = df_a.sort_values("PLACA")

    mes_slug = re.sub(r"[^\w\-]+", "_", (mes_label or "mes")).strip("_") or "mes"
    nombre_archivo = f"auditoria_sce_{mes_slug}_{datetime.now():%Y%m%d_%H%M}.xlsx"
    excel_bytes = construir_excel_auditoria(
        df_a, c_km, c_gal, c_ren, c_sub, meta_kmg, col_display, factor
    )
    with c3:
        st.download_button(
            label="Exportar Excel",
            data=excel_bytes,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="btn_export_auditoria",
            help="Descarga la tabla visible (filtros y orden aplicados), con CECO y valores numéricos.",
        )

    df_a["_ESTADO_TXT"] = df_a[c_ren].apply(lambda x: get_semaforo_rendimiento(x, meta_kmg)[0])
    df_a["_ID"] = serie_display(df_a, col_display)
    col_id = etiqueta_columna_id(col_display)

    df_show = df_a[["_ID", "TIPO", "MARCA", c_km, c_gal, c_ren, c_sub, "_ESTADO_TXT"]].copy()
    df_show.columns = [col_id, "TIPO", "MARCA", "KM", "GALONES", "RENDIMIENTO", "COSTO", "ESTADO"]
    df_show["KM"] = df_show["KM"].apply(lambda x: f"{format_number(x, 2):,.2f}")
    df_show["GALONES"] = df_show["GALONES"].apply(lambda x: f"{format_number(x, 2):,.2f}")
    df_show["RENDIMIENTO"] = df_show["RENDIMIENTO"].apply(lambda x: f"{format_number(x, 2):,.2f}")
    df_show["COSTO"] = df_show["COSTO"].apply(lambda x: f"{simbolo} {format_number(x * factor, 2):,.2f}")

    columnas = [col_id, "TIPO", "MARCA", "KM", "GALONES", "RENDIMIENTO", "COSTO", "ESTADO"]
    thead = "".join(f"<th>{html.escape(str(c))}</th>" for c in columnas)
    filas_html = []
    for _, fila in df_show.iterrows():
        est = fila["ESTADO"]
        colores = COLORES_SEMAFORO_FILA.get(est, COLORES_SEMAFORO_FILA["CRÍTICO"])
        celdas = "".join(
            f'<td>{html.escape(str(fila[c]))}</td>' for c in columnas[:-1]
        )
        celdas += (
            f'<td style="background:{colores["badge"]};color:#000;font-weight:900;'
            f'text-align:center;border:2px solid #fff;">{est}</td>'
        )
        filas_html.append(
            f'<tr style="background:{colores["background"]};color:{colores["color"]};">{celdas}</tr>'
        )
    st.markdown(
        '<div class="audit-table-wrap"><table class="audit-table">'
        f"<thead><tr>{thead}</tr></thead><tbody>{''.join(filas_html)}</tbody></table></div>",
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="legend-semaforo">
            <div class="legend-item legend-excelente">● EXCELENTE — verde (≥ 100% meta)</div>
            <div class="legend-item legend-alerta">● PRECAUCIÓN — ámbar (85% – 99% meta)</div>
            <div class="legend-item legend-critico">● CRÍTICO — rojo (&lt; 85% meta)</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def aviso_amigable(mensaje=None):
    st.markdown(
        f'<div class="warning-box">{mensaje or MSG_NO_DATA}</div>',
        unsafe_allow_html=True,
    )


def normalizar_provincia(nombre):
    """Mapea textos del Excel (ej. 'LURIN - LIMA') a claves de PROVINCIAS_PERU (solo mapa)."""
    if nombre is None or (isinstance(nombre, float) and pd.isna(nombre)):
        return None
    texto = _texto_sin_acentos(str(nombre).strip())
    if not texto or texto in ("NAN", "NONE", ""):
        return None
    if texto in PROVINCIAS_PERU:
        return texto
    if " - " in str(nombre).upper():
        partes = [p.strip() for p in str(nombre).upper().split(" - ") if p.strip()]
        for parte in reversed(partes):
            pnorm = _texto_sin_acentos(parte)
            if pnorm in PROVINCIAS_PERU:
                return pnorm
    for prov in PROVINCIAS_PERU:
        if prov in texto:
            return prov
    return None


def etiqueta_provincia_tabla(nombre):
    """Etiqueta para tablas: conserva TODAS las provincias del Excel (no descarta filas)."""
    if nombre is None or (isinstance(nombre, float) and pd.isna(nombre)):
        return None
    texto = str(nombre).strip()
    if not texto or texto.upper() in ("NAN", "N/A", "NONE", ""):
        return None
    catalogada = normalizar_provincia(nombre)
    if catalogada:
        return catalogada.title()
    return _texto_sin_acentos(texto).title()


def buscar_columna_proveedor(df):
    """Columna PRIMAX/REDCOL — nunca CONTROL COSTO (códigos CECO)."""
    if df is None or df.empty:
        return None
    if "PROVEEDOR" in df.columns:
        return "PROVEEDOR"
    for col in df.columns:
        compact = nombre_columna_compacto(col)
        if compact == "PROVEEDOR" or compact.endswith("PROVEEDOR"):
            return col
    return buscar_columna(df, "PROVEEDOR")


def contar_proveedor(serie, clave):
    s = serie.astype(str).str.upper().str.strip()
    return int(s.str.contains(clave, na=False).sum())


def limpiar_nombre_columna(nombre):
    s = str(nombre).strip().upper().replace("MY_", "MAY_")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip()


def nombre_columna_compacto(nombre):
    return re.sub(r"[\s_]+", "", limpiar_nombre_columna(nombre))


def buscar_columna(df, *patrones):
    if df is None or df.empty:
        return None
    for col in df.columns:
        nombre = limpiar_nombre_columna(col)
        compact = nombre_columna_compacto(col)
        for pat in patrones:
            p = limpiar_nombre_columna(pat)
            pc = nombre_columna_compacto(pat)
            if nombre == p or compact == pc or (pc and pc in compact):
                return col
    return None


def buscar_columna_galones(df):
    if df is None or df.empty:
        return None
    for col in df.columns:
        c = nombre_columna_compacto(col)
        if "GALON" in c and ("CANTIDAD" in c or c.startswith("CANT")):
            return col
    return buscar_columna(df, "CANTIDAD DE GALONES", "GALONES", "GALON")


def buscar_columna_estacion(df):
    if df is None or df.empty:
        return None
    for col in df.columns:
        if nombre_columna_compacto(col).startswith("ESTACI"):
            return col
    return buscar_columna(df, "ESTACION", "ESTACIÓN")


def resolver_columnas_registro(df):
    """Mapea columnas de la hoja Registro aunque se hayan insertado columnas nuevas."""
    if df is None or df.empty:
        return {}
    return {
        "provincia": buscar_columna(df, "PROVINCIA"),
        "estacion": buscar_columna_estacion(df),
        "proveedor": buscar_columna_proveedor(df),
        "estado": buscar_columna(df, "ESTADO"),
        "galones": buscar_columna_galones(df),
        "monto": resolver_columna_monto(df),
        "conductor": buscar_columna(df, "CONDUCTOR", "PORTADOR"),
        "placa": "PLACA" if "PLACA" in df.columns else buscar_columna(df, "PLACA"),
    }


def _redes_filtros_activos(estaciones_sel, provincias_sel, placas_sel):
    return bool(estaciones_sel or provincias_sel or placas_sel)


def _limpiar_filtros_redes():
    for key in (
        "red_filtro_provincia",
        "red_filtro_estacion",
        "red_filtro_placa",
    ):
        if key in st.session_state:
            del st.session_state[key]


def lista_unica_texto(serie):
    """Valores únicos como texto ordenados (evita error float vs str con NaN)."""
    if serie is None:
        return []
    out = set()
    for val in serie.dropna().tolist():
        txt = str(val).strip().upper()
        if txt and txt not in ("NAN", "NONE", "NAT", "<NA>"):
            out.add(txt)
    return sorted(out)


def parse_numero_locale(val):
    """
    Número con formato peruano (coma = decimal: 14,258 → 14.258)
    o US (coma = miles: 1,234.56 → 1234.56).
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s.upper() in ("NAN", "NONE", ""):
        return 0.0
    s = s.replace("S/.", "").replace("S/", "").replace("$", "").strip()
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1].strip()
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        partes = s.split(",")
        if len(partes) == 2 and len(partes[1]) <= 3:
            s = partes[0] + "." + partes[1]
        else:
            s = s.replace(",", "")
    import re

    m = re.search(r"-?\d+\.?\d*", s)
    if not m:
        return 0.0
    n = float(m.group(0))
    return -n if neg else n


def to_numeric_locale(serie):
    if serie is None:
        return pd.Series(dtype=float)
    return serie.map(parse_numero_locale)


def to_numeric_monto(serie):
    """Convierte montos del Excel (S/., comas, texto) a número."""
    return to_numeric_locale(serie)


def resolver_columna_monto(df):
    """
    Columna de importe total por abastecimiento (no precio unitario).
    Prioridad: SUBTOTAL > TOTAL > IMPORTE > MONTO > PAGO.
    """
    if df is None or df.empty:
        return None
    candidatos = [(nombre_columna_compacto(c), c) for c in df.columns]

    def mejor_col(test_fn):
        for compact, col in candidatos:
            if test_fn(compact):
                vals = to_numeric_monto(df[col])
                if float(vals.sum()) > 0:
                    return col
        return None

    col = mejor_col(lambda n: "SUBTOTAL" in n)
    if col:
        return col
    col = mejor_col(lambda n: n.endswith("TOTAL") and "SUBTOTAL" not in n)
    if col:
        return col
    for pat in ("IMPORTE", "MONTO", "PAGO", "VALOR"):
        col = mejor_col(lambda n, p=pat: p in n and "PRECIO" not in n and "UNIT" not in n)
        if col:
            return col
    col = mejor_col(lambda n: "PRECIO" in n and "PROM" not in n)
    if col:
        return col
    return None


def calcular_monto_fila(df, gal_col, precio_col=None):
    """Monto en soles por registro: SUBTOTAL/TOTAL; si solo PRECIO unitario, × galones."""
    if df is None or df.empty:
        return pd.Series(dtype=float)
    col = precio_col if precio_col else resolver_columna_monto(df)
    if not col or col not in df.columns:
        return pd.Series(0.0, index=df.index)
    monto = to_numeric_monto(df[col])
    nombre = nombre_columna_compacto(col)
    if "SUBTOTAL" in nombre or (nombre.endswith("TOTAL") and "SUBTOTAL" not in nombre):
        return monto
    if gal_col and gal_col in df.columns:
        gal = to_numeric_locale(df[gal_col])
        if float(gal.sum()) > 0 and float(monto.max()) < 80 and float(monto.median()) < 25:
            if "PRECIO" in nombre:
                monto = monto * gal
    return monto


def normalizar_columnas(df):
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()
    df.columns = [str(c).strip().upper().replace("MY_", "MAY_") for c in df.columns]
    return df


def _resolver_hoja(xl, nombre_objetivo):
    objetivo = nombre_objetivo.strip().lower()
    for hoja in xl.sheet_names:
        if hoja.strip().lower() == objetivo:
            return hoja
    return None


def _resolver_hoja_fragmento(xl, fragmento):
    frag = fragmento.strip().lower()
    for hoja in xl.sheet_names:
        hn = hoja.strip().lower().replace("ó", "o").replace("í", "i")
        if frag in hn:
            return hoja
    return None


URL_MODELO_CAMION_3D = (
    "https://raw.githubusercontent.com/KhronosGroup/glTF-Sample-Models/master/"
    "2.0/CesiumMilkTruck/glTF-Binary/CesiumMilkTruck.glb"
)
URL_MODELO_TRACTO_3D = URL_MODELO_CAMION_3D
URL_MODELO_UTIL_3D = (
    "https://raw.githubusercontent.com/mrdoob/three.js/r128/examples/models/gltf/ferrari.glb"
)
HDR_ENV_3D = "https://modelviewer.dev/shared-assets/environments/spruit_sunrise_1k_HDR.hdr"


def _config_modelo_3d(tipo_flota, marca, modelo):
    """URL + cámara según tipo de flota y marca (referencia visual más cercana)."""
    t = _texto_sin_acentos(str(tipo_flota or ""))
    m = _texto_sin_acentos(str(marca or ""))
    mod = str(modelo or "").strip()
    etiqueta = f"{marca} {mod}".strip() if marca not in ("—", "") else mod or "Flota SCE"
    url = URL_MODELO_CAMION_3D
    orbit = "42deg 88% 102%"
    scale = "1.05 1.05 1.05"
    categoria = "Camión de carga"

    marcas_tracto = (
        "IVECO", "VOLVO", "SCANIA", "FREIGHT", "INTERNATIONAL", "KENWORTH",
        "MACK", "PETERBILT", "MERCEDES", "MAN", "DAF", "RENAULT",
    )
    marcas_camion = (
        "DONGFENG", "FAW", "SHACMAN", "JAC", "CHEVROLET", "HINO", "ISUZU",
        "HYUNDAI", "FOTON", "FORLAND", "BEIBEN", "HOWO", "SINOTRUK",
    )

    if any(k in t for k in ("TRACT", "REMOL", "CISTER", "SEMIR")):
        url = URL_MODELO_TRACTO_3D
        orbit = "28deg 72% 125%"
        scale = "1.22 1.22 1.22"
        categoria = "Tracto"
    elif any(k in t for k in ("PICK", "FURG", "VAN", "UTILIT", "AUTO")):
        url = URL_MODELO_UTIL_3D
        orbit = "38deg 92% 96%"
        scale = "0.85 0.85 0.85"
        categoria = "Utilitario"
    elif "VOLQU" in t or "TIPPER" in t:
        orbit = "36deg 82% 112%"
        scale = "1.2 1.2 1.2"
        categoria = "Volquete"

    if any(mk in m for mk in marcas_tracto):
        url = URL_MODELO_TRACTO_3D
        orbit = "26deg 70% 128%"
        scale = "1.25 1.25 1.25"
        categoria = f"Tracto · {marca}"
    elif any(mk in m for mk in marcas_camion) or "DONGFENG" in m:
        url = URL_MODELO_CAMION_3D
        orbit = "38deg 84% 112%"
        scale = "1.14 1.14 1.14"
        categoria = f"Camión · {marca}" if marca not in ("—", "") else "Camión de carga"

    return url, orbit, scale, html.escape(etiqueta), html.escape(categoria)


def _preparar_vehiculos(df):
    df = normalizar_columnas(df)
    if df.empty:
        return df
    if "PLACA" in df.columns:
        df["PLACA"] = df["PLACA"].astype(str).str.strip().str.upper()
    return df


def _valor_ficha(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "—"
    s = str(val).strip()
    return s if s and s.upper() not in ("NAN", "N/A", "-", "NONE") else "—"


def buscar_ficha_vehiculo(placa, df_veh, df_bdmes):
    """Combina hoja TMDatosvehículo + BdMes por placa."""
    placa = str(placa).strip().upper()
    ficha = {"PLACA": placa}
    if df_veh is not None and not df_veh.empty and "PLACA" in df_veh.columns:
        fila = df_veh[df_veh["PLACA"] == placa]
        if not fila.empty:
            row = fila.iloc[0]
            for col in df_veh.columns:
                ficha[str(col)] = row[col]
    if df_bdmes is not None and not df_bdmes.empty and "PLACA" in df_bdmes.columns:
        fila_b = df_bdmes[df_bdmes["PLACA"] == placa]
        if not fila_b.empty:
            row = fila_b.iloc[0]
            for c in ("TIPO", "MARCA", "MODELO", "CONDUCTOR", "CECO"):
                if c in row.index and c not in ficha:
                    ficha[c] = row[c]
    return ficha


def render_holograma_vehiculo(tipo_flota, marca, modelo):
    """Visor 3D estilo holograma (model-viewer + iluminación HDR)."""
    url, orbit, scale, titulo, categoria = _config_modelo_3d(tipo_flota, marca, modelo)
    components.html(
        f"""
        <script type="module" src="https://ajax.googleapis.com/ajax/libs/model-viewer/3.4.0/model-viewer.min.js"></script>
        <div class="vehiculo-holo-stage">
            <div class="holo-veh-titulo">{titulo}</div>
            <div class="holo-veh-sub">{categoria} · Socorro Cargo Express</div>
            <div class="holo-ring"></div>
            <div class="holo-ring holo-ring-2"></div>
            <model-viewer src="{url}"
                alt="{titulo}"
                scale="{scale}"
                camera-orbit="{orbit}"
                min-camera-orbit="auto auto 60%"
                max-camera-orbit="auto auto 140%"
                environment-image="{html.escape(HDR_ENV_3D)}"
                exposure="1.25"
                shadow-intensity="1.35"
                shadow-softness="0.8"
                auto-rotate
                rotation-per-second="22deg"
                camera-controls
                touch-action="pan-y"
                interaction-prompt="none"
                style="background:transparent;">
            </model-viewer>
            <div class="sce-logo-3d">SCE</div>
            <div class="holo-veh-hint">Arrastre para girar · scroll para zoom · referencia 3D según tipo y marca</div>
        </div>
        """,
        height=540,
        scrolling=False,
    )


def render_tarjeta_spec(label, value):
    st.markdown(
        f'<div class="vehiculo-spec-card"><div class="spec-label">{html.escape(label)}</div>'
        f'<div class="spec-value">{html.escape(str(value))}</div></div>',
        unsafe_allow_html=True,
    )


def render_vista_ficha_vehiculo(df_bdmes, df_vehiculos, mes_corto, simbolo):
    show_update_info()
    st.markdown('<div class="main-title">Ficha técnica — Flota SCE</div>', unsafe_allow_html=True)
    placas_bd = (
        sorted(df_bdmes["PLACA"].dropna().unique().tolist())
        if df_bdmes is not None and not df_bdmes.empty and "PLACA" in df_bdmes.columns
        else []
    )
    placas_tm = (
        sorted(df_vehiculos["PLACA"].dropna().unique().tolist())
        if df_vehiculos is not None and not df_vehiculos.empty and "PLACA" in df_vehiculos.columns
        else []
    )
    placas = sorted(set(placas_bd) | set(placas_tm))
    if not placas:
        aviso_amigable("No hay placas en BdMes ni en TMDatosvehículo.")
        return
    c_placa, _ = st.columns([2, 1])
    with c_placa:
        placa_sel = st.selectbox("Seleccione placa", placas, key="ficha_placa_sel")
    ficha = buscar_ficha_vehiculo(placa_sel, df_vehiculos, df_bdmes)
    tipo = _valor_ficha(
        ficha.get("TIPO")
        or ficha.get("TIPO DE VEHICULO")
        or ficha.get("TIPO DE VEHICULO2")
    )
    marca = _valor_ficha(ficha.get("MARCA"))
    modelo = _valor_ficha(ficha.get("MODELO"))

    col_3d, col_specs = st.columns([1.75, 1])
    with col_3d:
        render_holograma_vehiculo(tipo, marca, modelo)
        if df_bdmes is not None and not df_bdmes.empty and mes_corto:
            cols = columnas_mes(mes_corto, df_bdmes.columns)
            fila_m = df_bdmes[df_bdmes["PLACA"] == placa_sel]
            if not fila_m.empty and cols["gal"] in df_bdmes.columns:
                r = fila_m.iloc[0]
                st.markdown(
                    f'<div class="section-title" style="margin-top:12px;">Consumo — '
                    f'{MESES_COMPLETOS.get(mes_corto, mes_corto)}</div>',
                    unsafe_allow_html=True,
                )
                k1, k2, k3, k4 = st.columns(4)
                render_kpi_beautiful(
                    k1, "Galones", f"{format_number(r.get(cols['gal'], 0), 2):,.2f}", "GL"
                )
                render_kpi_beautiful(
                    k2, "Kilómetros", f"{format_number(r.get(cols['km'], 0), 2):,.2f}", "KM"
                )
                render_kpi_beautiful(
                    k3, "Rendimiento", f"{format_number(r.get(cols['ren'], 0), 2):,.2f}", "KM/G"
                )
                render_kpi_beautiful(
                    k4, "Gasto", f"{simbolo} {format_number(r.get(cols['sub'], 0), 2):,.2f}"
                )

    with col_specs:
        st.markdown('<div class="section-title">Ficha técnica</div>', unsafe_allow_html=True)
        st.markdown('<div class="vehiculo-specs-panel">', unsafe_allow_html=True)
        s1, s2 = st.columns(2)
        with s1:
            render_tarjeta_spec("Placa", placa_sel)
            render_tarjeta_spec("Tipo flota", tipo)
            render_tarjeta_spec("Marca", marca)
            render_tarjeta_spec("Modelo", modelo)
            render_tarjeta_spec("Año", _valor_ficha(ficha.get("AÑO") or ficha.get("ANO")))
        with s2:
            render_tarjeta_spec("Conductor", _valor_ficha(ficha.get("CONDUCTOR")))
            render_tarjeta_spec("CECO", _valor_ficha(ficha.get("CECO")))
            render_tarjeta_spec("Tanque (GL)", _valor_ficha(ficha.get("TANQUE DE COMBUSTIBLE")))
            render_tarjeta_spec("Combustible", _valor_ficha(ficha.get("COMBUSTIBLE")))
            render_tarjeta_spec("Ejes", _valor_ficha(ficha.get("EJES")))
        render_tarjeta_spec(
            "Dimensiones L × A × Al (m)",
            f"{_valor_ficha(ficha.get('LARGO'))} × {_valor_ficha(ficha.get('ANCHO'))} × {_valor_ficha(ficha.get('ALTO'))}",
        )
        render_tarjeta_spec("Peso bruto", _valor_ficha(ficha.get("PESO BRUTO")))
        render_tarjeta_spec("Capacidad TN", _valor_ficha(ficha.get("CAPACIDAD TN")))
        rpm_val = "—"
        for clave, val in ficha.items():
            if "RPM" in str(clave).upper() or "REVOLUCION" in str(clave).upper():
                rpm_val = _valor_ficha(val)
                break
        render_tarjeta_spec("RPM", rpm_val)
        st.markdown("</div>", unsafe_allow_html=True)


def _preparar_bdmes(df):
    df = normalizar_columnas(df)
    if df.empty:
        return df
    if "W" in df.columns and "PLACA" not in df.columns:
        df = df.rename(columns={"W": "PLACA"})
    if "PLACA" in df.columns:
        df["PLACA"] = df["PLACA"].astype(str).str.strip().str.upper()
    cond_col = buscar_columna(df, "CONDUCTOR")
    if cond_col and cond_col != "CONDUCTOR":
        df = df.rename(columns={cond_col: "CONDUCTOR"})
    ceco_col = buscar_columna(df, "CECO", "CENTRO DE COSTO", "CENTRODECOSTO")
    if ceco_col and ceco_col != "CECO":
        df = df.rename(columns={ceco_col: "CECO"})
    for col in ["MARCA", "MODELO", "TIPO"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.upper().replace("NAN", "N/A")
    for col in df.columns:
        if any(x in col for x in ["GALONES", "KM", "RENDIMIENTO", "SUBTOTAL"]):
            df[col] = to_numeric_locale(df[col])
    return df


def _preparar_registro(df):
    df = normalizar_columnas(df)
    if df.empty:
        return df
    cond_col = buscar_columna(df, "CONDUCTOR", "PORTADOR")
    if cond_col and cond_col != "CONDUCTOR" and "CONDUCTOR" not in df.columns:
        df = df.rename(columns={cond_col: "CONDUCTOR"})
    col_monto = resolver_columna_monto(df)
    for col in df.columns:
        nombre = str(col).upper()
        if "GALON" in nombre or ("CANTIDAD" in nombre and "GALON" in nombre):
            df[col] = to_numeric_locale(df[col])
        elif col == col_monto or any(
            k in nombre for k in ("SUBTOTAL", "TOTAL", "IMPORTE", "MONTO", "PAGO", "VALOR", "PRECIO")
        ):
            df[col] = to_numeric_monto(df[col])
        elif df[col].dtype == object:
            df[col] = df[col].astype(str).str.upper().str.strip().replace("NAN", "")
    if "PLACA" in df.columns:
        df["PLACA"] = df["PLACA"].astype(str).str.strip().str.upper()
    if "MES_NUM" in df.columns:
        df["MES_NUM"] = pd.to_numeric(df["MES_NUM"], errors="coerce")
    return df


@st.cache_data(ttl=600, show_spinner="Sincronizando Excel desde SharePoint…")
def cargar_datos_sharepoint():
    df_bdmes = pd.DataFrame()
    df_registro = pd.DataFrame()
    df_vehiculos = pd.DataFrame()
    errores = []
    url = _excel_url()
    try:
        buf = _abrir_excel_remoto(url)
        xl = pd.ExcelFile(buf, engine="openpyxl")
        mapa_ceco = cargar_mapa_ceco(xl)
        if not mapa_ceco:
            errores.append(
                "No se cargó la hoja TMCeco (catálogo CECO). "
                "Revise que exista y tenga código + nombre de área."
            )
        hoja_bd = _resolver_hoja(xl, "BdMes")
        hoja_reg = _resolver_hoja(xl, "registro")
        if hoja_bd:
            df_bdmes = _preparar_bdmes(pd.read_excel(xl, sheet_name=hoja_bd, engine="openpyxl"))
            df_bdmes = aplicar_nombres_ceco_df(df_bdmes, mapa_ceco)
        else:
            errores.append("Hoja BdMes no encontrada en SharePoint.")
        if hoja_reg:
            df_registro = _preparar_registro(pd.read_excel(xl, sheet_name=hoja_reg, engine="openpyxl"))
        else:
            errores.append('Hoja "registro" no encontrada en SharePoint.')
        if (df_registro is None or df_registro.empty) and hoja_bd:
            df_bd_raw = pd.read_excel(xl, sheet_name=hoja_bd, engine="openpyxl")
            if es_hoja_registro_detalle(df_bd_raw):
                df_registro = _preparar_registro(df_bd_raw)
        hoja_veh = _resolver_hoja_fragmento(xl, "datosveh")
        if hoja_veh:
            df_vehiculos = _preparar_vehiculos(
                pd.read_excel(xl, sheet_name=hoja_veh, engine="openpyxl")
            )
    except requests.HTTPError as exc:
        cod = getattr(getattr(exc, "response", None), "status_code", "?")
        errores.append(
            f"HTTP {cod} desde Render. Origen URL: {_origen_excel_url()}. "
            "SharePoint debe ser público (cualquier persona con el enlace)."
        )
    except Exception as exc:
        errores.append(f"{exc} (origen URL: {_origen_excel_url()})")
    return df_bdmes, df_registro, df_vehiculos, errores


def filtrar_registro(df_reg, mes_corto, placas=None):
    if df_reg is None or df_reg.empty:
        return pd.DataFrame()
    df = df_reg.copy()
    mes_num = MES_A_NUM.get(mes_corto)
    if mes_num is not None and "MES_NUM" in df.columns:
        df = df[df["MES_NUM"] == mes_num]
    if placas is not None and len(placas) > 0 and "PLACA" in df.columns:
        placas_set = {str(p).strip().upper() for p in placas}
        df = df[df["PLACA"].isin(placas_set)]
    return df


def filtrar_vista_redes(df_reg, mes_corto, reg_cols, estaciones_sel=None, provincias_sel=None, placas_sel=None):
    """Mismos criterios que Excel: mes fiscal + red + provincia + estación + placa."""
    if df_reg is None or df_reg.empty:
        return pd.DataFrame()
    df = df_reg.copy()
    if mes_corto:
        mes_num = MES_A_NUM.get(mes_corto)
        if mes_num is not None and "MES_NUM" in df.columns:
            df = df[pd.to_numeric(df["MES_NUM"], errors="coerce") == mes_num]
    prov_col = reg_cols.get("provincia") if reg_cols else None
    est_col = reg_cols.get("estacion") if reg_cols else None
    placa_col = reg_cols.get("placa") if reg_cols else None
    if provincias_sel and prov_col and prov_col in df.columns:
        prov_set = {str(p).strip().upper() for p in provincias_sel}
        df = df[df[prov_col].astype(str).str.strip().str.upper().isin(prov_set)]
    if estaciones_sel and est_col and est_col in df.columns:
        est_set = {str(e).strip().upper() for e in estaciones_sel}
        df = df[df[est_col].astype(str).str.strip().str.upper().isin(est_set)]
    if placas_sel and placa_col and placa_col in df.columns:
        placa_set = {str(p).strip().upper() for p in placas_sel}
        df = df[df[placa_col].astype(str).str.strip().str.upper().isin(placa_set)]
    return df


def es_hoja_registro_detalle(df):
    """BdMes u otra hoja con filas por abastecimiento (provincia, galones, proveedor)."""
    if df is None or df.empty:
        return False
    compact = [nombre_columna_compacto(c) for c in df.columns]
    tiene_prov = any("PROVINCIA" in c for c in compact)
    tiene_gal = any("GALON" in c for c in compact)
    tiene_red = any("PROVEEDOR" in c for c in compact)
    return tiene_prov and tiene_gal and tiene_red


def filtrar_registro_meses(df_reg, meses_cortos, placas=None):
    if df_reg is None or df_reg.empty or not meses_cortos:
        return pd.DataFrame()
    df = df_reg.copy()
    mes_nums = [MES_A_NUM.get(m) for m in meses_cortos if MES_A_NUM.get(m) is not None]
    if mes_nums and "MES_NUM" in df.columns:
        df = df[df["MES_NUM"].isin(mes_nums)]
    if placas is not None and len(placas) > 0 and "PLACA" in df.columns:
        placas_set = {str(p).strip().upper() for p in placas}
        df = df[df["PLACA"].isin(placas_set)]
    return df


def agregar_metricas_meses(df, meses_cortos):
    """Suma KM, galones y costo de varios meses en columnas temporales."""
    if df is None or df.empty or not meses_cortos:
        return pd.DataFrame()
    out = df.copy()
    km_cols, gal_cols, sub_cols = [], [], []
    for mes in meses_cortos:
        cols = columnas_mes(mes, out.columns)
        if cols["km"] in out.columns:
            km_cols.append(cols["km"])
        if cols["gal"] in out.columns:
            gal_cols.append(cols["gal"])
        if cols["sub"] in out.columns:
            sub_cols.append(cols["sub"])
    out["_KM_T"] = out[km_cols].sum(axis=1) if km_cols else 0
    out["_GAL_T"] = out[gal_cols].sum(axis=1) if gal_cols else 0
    out["_SUB_T"] = out[sub_cols].sum(axis=1) if sub_cols else 0
    return out


def tabla_participacion(df, columna_grupo, km_col="_KM_T", gal_col="_GAL_T", sub_col="_SUB_T"):
    """Tabla de participación % y valores absolutos por CECO o TIPO."""
    if df is None or df.empty or columna_grupo not in df.columns:
        return pd.DataFrame()
    agg = (
        df.groupby(columna_grupo, as_index=False)
        .agg({km_col: "sum", gal_col: "sum", sub_col: "sum"})
        .rename(columns={columna_grupo: "Grupo", km_col: "KM", gal_col: "Galones", sub_col: "Costo"})
    )
    tot_km = agg["KM"].sum()
    tot_gal = agg["Galones"].sum()
    tot_sub = agg["Costo"].sum()
    agg["%_KM"] = (agg["KM"] / tot_km * 100).round(2) if tot_km > 0 else 0
    agg["%_GALONES"] = (agg["Galones"] / tot_gal * 100).round(2) if tot_gal > 0 else 0
    agg["%_COSTO"] = (agg["Costo"] / tot_sub * 100).round(2) if tot_sub > 0 else 0
    return agg.sort_values("Costo", ascending=False)


def ordenar_meses(meses, orden_ref):
    return [m for m in orden_ref if m in meses]


def tabla_participacion_mes(df, columna_grupo, mes_corto):
    """Participación por grupo en un solo mes (sin sumar trimestre)."""
    if df is None or df.empty or columna_grupo not in df.columns:
        return pd.DataFrame()
    df = df.copy()
    if columna_grupo == "CECO":
        df[columna_grupo] = df[columna_grupo].apply(
            lambda v: "SIN CECO ASIGNADO"
            if pd.isna(v) or str(v).strip() in ("", "NAN", "N/A")
            else v
        )
    cols = columnas_mes(mes_corto, df.columns)
    km_c, gal_c, sub_c = cols["km"], cols["gal"], cols["sub"]
    if not all(c in df.columns for c in (km_c, gal_c, sub_c)):
        return pd.DataFrame()
    agg = (
        df.groupby(columna_grupo, as_index=False)
        .agg({km_c: "sum", gal_c: "sum", sub_c: "sum"})
        .rename(columns={columna_grupo: "Grupo", km_c: "KM", gal_c: "Galones", sub_c: "Costo"})
    )
    tot_km, tot_gal, tot_sub = agg["KM"].sum(), agg["Galones"].sum(), agg["Costo"].sum()
    agg["%_KM"] = (agg["KM"] / tot_km * 100).round(2) if tot_km > 0 else 0
    agg["%_GALONES"] = (agg["Galones"] / tot_gal * 100).round(2) if tot_gal > 0 else 0
    agg["%_COSTO"] = (agg["Costo"] / tot_sub * 100).round(2) if tot_sub > 0 else 0
    return agg.sort_values("Costo", ascending=False)


def combustible_por_mes(df_reg, meses_ordenados):
    """Galones por producto y mes (sin agregar meses)."""
    if df_reg is None or df_reg.empty:
        return pd.DataFrame()
    prod_col = buscar_columna(df_reg, "PRODUCTO")
    gal_col = buscar_columna(df_reg, "CANTIDAD", "GALON")
    if not prod_col or not gal_col:
        return pd.DataFrame()
    filas = []
    tipos = ["DIESEL", "GASOLINA", "GLP", "GNV"]
    for mes in meses_ordenados:
        df_m = filtrar_registro(df_reg, mes, None)
        if df_m.empty:
            continue
        df_m = df_m.copy()
        df_m[prod_col] = df_m[prod_col].astype(str).str.upper().str.strip()
        df_m[gal_col] = to_numeric_locale(df_m[gal_col])
        etiqueta = MESES_COMPLETOS.get(mes, mes)
        for prod in tipos:
            gal = df_m.loc[df_m[prod_col] == prod, gal_col].sum()
            filas.append({"Mes": etiqueta, "Mes_Corto": mes, "Producto": prod, "Galones": float(gal)})
    return pd.DataFrame(filas)


def totales_bdmes_por_mes(df, meses_ordenados):
    """KM, galones y costo totales por mes para gráfico comparativo."""
    filas = []
    for mes in meses_ordenados:
        cols = columnas_mes(mes, df.columns)
        if not all(c in df.columns for c in (cols["km"], cols["gal"], cols["sub"])):
            continue
        filas.append({
            "Mes": MESES_COMPLETOS.get(mes, mes),
            "Mes_Corto": mes,
            "KM": df[cols["km"]].sum(),
            "Galones": df[cols["gal"]].sum(),
            "Costo": df[cols["sub"]].sum(),
        })
    return pd.DataFrame(filas)


def formatear_tabla_participacion(df_tab, simbolo, factor):
    if df_tab.empty:
        return df_tab
    out = df_tab.copy()
    out["KM"] = out["KM"].apply(lambda x: f"{format_number(x, 2):,.2f}")
    out["Galones"] = out["Galones"].apply(lambda x: f"{format_number(x, 2):,.2f}")
    out["Costo"] = out["Costo"].apply(lambda x: f"{simbolo} {format_number(x * factor, 2):,.2f}")
    out["%_KM"] = out["%_KM"].apply(lambda x: f"{x:.2f}%")
    out["%_GALONES"] = out["%_GALONES"].apply(lambda x: f"{x:.2f}%")
    out["%_COSTO"] = out["%_COSTO"].apply(lambda x: f"{x:.2f}%")
    return out


def columnas_mes(mes_corto, columnas_df=None):
    """Encuentra columnas del mes en BdMes aunque cambie el orden o se agreguen columnas."""
    pref = str(mes_corto).strip().upper().replace("MY_", "MAY_")
    sufijos = {
        "km": ("KM",),
        "gal": ("GALONES", "GALON"),
        "ren": ("RENDIMIENTO", "REND"),
        "sub": ("SUBTOTAL",),
        "precio": ("PRECIO",),
        "total": ("TOTAL",),
    }
    encontradas = {}
    cols_list = list(columnas_df) if columnas_df is not None else []
    for col in cols_list:
        compact = nombre_columna_compacto(col)
        if not compact.startswith(pref) or len(compact) <= len(pref):
            continue
        resto = compact[len(pref) :]
        for clave, keys in sufijos.items():
            if clave in encontradas:
                continue
            for key in keys:
                if resto == key or resto.endswith(key):
                    encontradas[clave] = col
                    break
    resultado = {}
    for clave, keys in sufijos.items():
        fallback = f"{pref}_{keys[0]}"
        resultado[clave] = encontradas.get(clave, fallback)
    return resultado


def meses_en_bdmes(df, orden_meses=None):
    orden = orden_meses or ORDEN_MESES
    presentes = []
    for col in df.columns:
        compact = nombre_columna_compacto(col)
        for mes in orden:
            if compact.startswith(mes) and len(compact) > len(mes):
                if mes not in presentes:
                    presentes.append(mes)
                break
    return [m for m in orden if m in presentes]


def aplicar_tema_plotly(fig, height=360, y2=False):
    fig.update_layout(
        template="plotly_dark",
        height=height,
        paper_bgcolor="#000000",
        plot_bgcolor="#0a0a0a",
        font=dict(color="#e5e7eb", size=12),
        margin=dict(l=48, r=32, t=28, b=48),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        colorway=CHART_COLORS,
        hovermode="x unified",
    )
    fig.update_xaxes(showgrid=False, zeroline=False)
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor="#1f2937", zeroline=False)
    if y2:
        fig.update_layout(yaxis2=dict(overlaying="y", side="right", range=[0, 105], showgrid=False))
    return fig


def safe_plotly_chart(fig):
    try:
        st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CFG)
    except Exception:
        aviso_amigable()

def render_kpi_beautiful(col, label, value, unit=""):
    if unit:
        html = f'''
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}<span class="kpi-unit"> {unit}</span></div>
        </div>
        '''
    else:
        html = f'''
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
        </div>
        '''
    col.markdown(html, unsafe_allow_html=True)

def show_update_info():
    now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    st.markdown(f'<div class="update-info">✓ Última actualización: {now} | Caché: Cada 5 min</div>', unsafe_allow_html=True)

df_raw, df_estaciones, df_vehiculos, _errores_carga = cargar_datos_sharepoint()
_bdmes_ok = df_raw is not None and not df_raw.empty
_registro_ok = df_estaciones is not None and not df_estaciones.empty
_vehiculos_ok = df_vehiculos is not None and not df_vehiculos.empty

if _bdmes_ok or _registro_ok:
    if _errores_carga:
        with st.sidebar:
            for err in _errores_carga:
                st.caption(f"⚠ {err}")

if _bdmes_ok or _registro_ok or _vehiculos_ok:
    orden_meses = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SET", "OCT", "NOV", "DIC"]
    meses_final = meses_en_bdmes(df_raw, orden_meses) if _bdmes_ok else orden_meses
    factor = 1.0
    simbolo = "S/."

    with st.sidebar:
        st.markdown('<div class="sidebar-brand">Socorro Cargo Express</div>', unsafe_allow_html=True)
        if st.button("Actualizar Excel ahora", use_container_width=True, help="F5 no limpia la caché de 5 minutos"):
            st.cache_data.clear()
            st.rerun()

        menu = render_selector_vista()

        ceco_sel, placa_sel, flota_sel, marca_sel, modelo_sel = [], [], [], [], []
        estacion_red_sel, provincia_red_sel, placa_red_sel = [], [], []
        meses_trim = []

        if menu == "Trimestre":
            st.divider()
            st.markdown('<div class="sidebar-panel-title">Trimestre</div>', unsafe_allow_html=True)
            meses_trim = st.multiselect(
                "Seleccione 3 meses",
                meses_final,
                default=meses_final[:3] if len(meses_final) >= 3 else meses_final,
                max_selections=3,
                help="Comparación mensual sin sumar el trimestre",
            )
            trimestre_label = " – ".join([MESES_COMPLETOS.get(m, m) for m in meses_trim]) if meses_trim else "Sin meses"
            mes_sel_corto = meses_trim[0] if meses_trim else (meses_final[0] if meses_final else "ABR")
            mes_sel = MESES_COMPLETOS.get(mes_sel_corto, mes_sel_corto)
            df_base = df_raw.copy()
            df_f1 = df_base.copy()
            df_f2 = df_base.copy()
            df_ceco = pd.DataFrame()
            col_display = "PLACA"
            meta_kmg = 11.50
            filtro_detalle = False
        else:
            st.divider()
            st.markdown('<div class="sidebar-panel-title">Filtros</div>', unsafe_allow_html=True)
            trimestre_label = ""
            mes_sel_corto = st.selectbox("Mes Fiscal", meses_final if meses_final else ["ABR"])
            mes_sel = MESES_COMPLETOS.get(mes_sel_corto, mes_sel_corto)

            if menu == VISTA_POR_FLOTA:
                df_base = df_raw.copy()

                st.markdown("**Filtros de flota**")
                tipos_lista = (
                    sorted([str(x) for x in df_base["TIPO"].unique() if pd.notna(x) and str(x) not in ("NAN", "N/A", "")])
                    if not df_base.empty and "TIPO" in df_base.columns
                    else []
                )
                flota_sel = st.multiselect("Tipo de Flota", tipos_lista)
                df_f1 = df_base[df_base["TIPO"].isin(flota_sel)].copy() if flota_sel else df_base.copy()

                marcas_lista = (
                    sorted([str(x) for x in df_f1["MARCA"].unique() if pd.notna(x) and str(x) not in ("NAN", "N/A")])
                    if not df_f1.empty
                    else []
                )
                marca_sel = st.multiselect("Marca", marcas_lista)
                df_f2 = df_f1[df_f1["MARCA"].isin(marca_sel)].copy() if marca_sel else df_f1.copy()

                modelos_lista = (
                    sorted([str(x) for x in df_f2["MODELO"].unique() if pd.notna(x) and str(x) not in ("NAN", "N/A")])
                    if not df_f2.empty
                    else []
                )
                modelo_sel = st.multiselect("Modelo", modelos_lista)

                st.markdown("**Filtros independientes**")
                ceco_lista = (
                    sorted([str(x) for x in df_raw["CECO"].unique() if pd.notna(x) and str(x) not in ("NAN", "N/A", "")])
                    if "CECO" in df_raw.columns
                    else []
                )
                ceco_sel = st.multiselect("CECO (área)", ceco_lista, help="Nombre del área — independiente de Placa")

                placas_lista = (
                    sorted([str(x) for x in df_raw["PLACA"].unique() if pd.notna(x) and str(x) not in ("NAN", "")])
                    if "PLACA" in df_raw.columns
                    else []
                )
                placa_sel = st.multiselect("Placa", placas_lista, help="Independiente de CECO")

                if ceco_sel and "CECO" in df_base.columns:
                    df_base = df_base[df_base["CECO"].isin(ceco_sel)].copy()
                if placa_sel and "PLACA" in df_base.columns:
                    df_base = df_base[df_base["PLACA"].isin(placa_sel)].copy()
                if flota_sel and "TIPO" in df_base.columns:
                    df_base = df_base[df_base["TIPO"].isin(flota_sel)].copy()
                if marca_sel and "MARCA" in df_base.columns:
                    df_base = df_base[df_base["MARCA"].isin(marca_sel)].copy()
                if modelo_sel and "MODELO" in df_base.columns:
                    df_base = df_base[df_base["MODELO"].isin(modelo_sel)].copy()

                df_ceco = (
                    df_raw[df_raw["CECO"].isin(ceco_sel)].copy()
                    if ceco_sel and "CECO" in df_raw.columns
                    else pd.DataFrame()
                )

                col_display = render_selector_graficos(df_raw)
                meta_kmg = render_slider_meta(4.0, 18.0, 11.50, key="slider_meta_kmg")

                filtro_detalle = bool(ceco_sel or placa_sel or flota_sel or marca_sel or modelo_sel)
            elif menu == VISTA_REDES and _registro_ok:
                df_base = df_raw.copy()
                df_f1 = df_base.copy()
                df_f2 = df_base.copy()
                df_ceco = pd.DataFrame()
                col_display = "PLACA"
                meta_kmg = 11.50
                filtro_detalle = False
                reg_cols_sb = resolver_columnas_registro(df_estaciones)
                st.markdown("**Filtros Red de estaciones**")
                st.caption("Misma lógica que Excel: mes + REDCOL/PRIMAX + provincia + estación + placa.")
                if reg_cols_sb.get("provincia"):
                    provs = lista_unica_texto(df_estaciones[reg_cols_sb["provincia"]])
                    provincia_red_sel = st.multiselect(
                        "Provincia", provs, key="red_filtro_provincia"
                    )
                if reg_cols_sb.get("estacion"):
                    ests = lista_unica_texto(df_estaciones[reg_cols_sb["estacion"]])
                    estacion_red_sel = st.multiselect(
                        "Estación", ests, key="red_filtro_estacion"
                    )
                if reg_cols_sb.get("placa"):
                    placas_r = lista_unica_texto(df_estaciones[reg_cols_sb["placa"]])
                    placa_red_sel = st.multiselect("Placa", placas_r, key="red_filtro_placa")
                if _redes_filtros_activos(estacion_red_sel, provincia_red_sel, placa_red_sel):
                    if st.button("Limpiar filtros Red", use_container_width=True):
                        _limpiar_filtros_redes()
                        st.rerun()
            elif menu == VISTA_FICHA_VEHICULO:
                df_base = df_raw.copy() if _bdmes_ok else pd.DataFrame()
                df_f1 = df_base.copy()
                df_f2 = df_base.copy()
                df_ceco = pd.DataFrame()
                col_display = "PLACA"
                meta_kmg = 11.50
                filtro_detalle = False
                st.caption("Elija la placa en la pantalla principal.")
            else:
                df_base = df_raw.copy()
                df_f1 = df_base.copy()
                df_f2 = df_base.copy()
                df_ceco = pd.DataFrame()
                col_display = "PLACA"
                meta_kmg = 11.50
                filtro_detalle = False
        
        st.divider()
        if st.button("🔄 Actualizar Datos (F5)"):
            st.cache_data.clear()
            st.rerun()


    cols_mes = columnas_mes(mes_sel_corto, df_raw.columns)
    c_km, c_gal, c_ren, c_sub = cols_mes["km"], cols_mes["gal"], cols_mes["ren"], cols_mes["sub"]
    cols_mes_ok = all(c in df_raw.columns for c in (c_km, c_gal, c_ren, c_sub))
    if not cols_mes_ok:
        st.warning(
            f"Columnas del mes {mes_sel_corto} no disponibles en BdMes "
            f"({c_km}, {c_gal}, {c_ren}, {c_sub})."
        )

    if filtro_detalle and not df_base.empty:
        df_filtrado = df_base.copy()
        if cols_mes_ok and c_km in df_filtrado.columns:
            df_dashboard = df_filtrado[df_filtrado[c_km] > 0].copy()
        else:
            df_dashboard = df_filtrado.copy()
        hay_datos = not df_dashboard.empty
        if ceco_sel and not df_ceco.empty:
            total_unidades = len(df_ceco)
        elif flota_sel and "TIPO" in df_raw.columns:
            total_unidades = len(df_raw[df_raw["TIPO"].isin(flota_sel)])
        elif placa_sel and not df_base.empty:
            total_unidades = len(df_base)
        else:
            total_unidades = len(df_raw)
        unidades_seleccionadas = len(df_filtrado)
    else:
        df_filtrado = pd.DataFrame()
        df_dashboard = pd.DataFrame()
        hay_datos = False
        total_unidades = len(df_raw)
        unidades_seleccionadas = 0

    if hay_datos and "PLACA" in df_dashboard.columns:
        placas_scope = df_dashboard["PLACA"].unique().tolist()
    elif ceco_sel and not df_base.empty and "PLACA" in df_base.columns:
        placas_scope = df_base["PLACA"].unique().tolist()
    else:
        placas_scope = None

    if menu == VISTA_REDES and _registro_ok:
        _reg_cols_vista = resolver_columnas_registro(df_estaciones)
        df_reg_vista = filtrar_vista_redes(
            df_estaciones,
            mes_sel_corto,
            _reg_cols_vista,
            estaciones_sel=estacion_red_sel or None,
            provincias_sel=provincia_red_sel or None,
            placas_sel=placa_red_sel or None,
        )
        df_reg_redes_base = filtrar_vista_redes(
            df_estaciones,
            None,
            _reg_cols_vista,
            estaciones_sel=estacion_red_sel or None,
            provincias_sel=provincia_red_sel or None,
            placas_sel=placa_red_sel or None,
        )
    else:
        df_reg_redes_base = pd.DataFrame()
    df_reg_vista = filtrar_registro(
        df_estaciones if _registro_ok else pd.DataFrame(),
        mes_sel_corto,
            placas_scope,
        )
    df_reg_trim = filtrar_registro_meses(
        df_estaciones if _registro_ok else pd.DataFrame(),
        meses_trim if meses_trim else [mes_sel_corto],
        placas_scope,
    )

    if menu == "Trimestre":
        show_update_info()
        st.markdown(f'<div class="main-title">Trimestre — {trimestre_label}</div>', unsafe_allow_html=True)
        meses_ord = ordenar_meses(meses_trim, orden_meses) if meses_trim else []

        if not meses_ord:
            aviso_amigable("Seleccione 3 meses en el panel lateral para la vista trimestral.")
        else:
            st.markdown('<div class="section-title">Abastecimiento por combustible — comparación mensual</div>', unsafe_allow_html=True)
            if _registro_ok and df_estaciones is not None:
                df_comb_mes = combustible_por_mes(df_estaciones, meses_ord)
                if not df_comb_mes.empty:
                    df_comb_plot = df_comb_mes[df_comb_mes["Galones"] > 0]
                    if not df_comb_plot.empty:
                        fig_comb = px.bar(
                            df_comb_plot,
                            x="Mes",
                            y="Galones",
                            color="Producto",
                            barmode="group",
                            category_orders={"Mes": [MESES_COMPLETOS.get(m, m) for m in meses_ord]},
                            color_discrete_sequence=["#FFD700", "#FF6B6B", "#4ECDC4", "#45B7D1"],
                        )
                        fig_comb.update_layout(xaxis_title="Mes", yaxis_title="Galones")
                        safe_plotly_chart(aplicar_tema_plotly(fig_comb, height=400))

                    st.markdown('<div class="section-title">Detalle por mes (ordenado)</div>', unsafe_allow_html=True)
                    cols_m = st.columns(len(meses_ord))
                    for col_m, mes in zip(cols_m, meses_ord):
                        etiqueta = MESES_COMPLETOS.get(mes, mes)
                        with col_m:
                            st.markdown(f"**{etiqueta}**")
                            df_mes_c = df_comb_mes[df_comb_mes["Mes_Corto"] == mes]
                            if df_mes_c.empty:
                                aviso_amigable()
                            else:
                                for _, row in df_mes_c.iterrows():
                                    val = row["Galones"]
                                    estado = "no-data" if val == 0 else ""
                                    st.markdown(
                                        f'<div class="combustible-item {estado}" style="margin-bottom:8px;padding:10px;">'
                                        f'<div class="combustible-tipo {estado}">{row["Producto"]}</div>'
                                        f'<div class="combustible-valor {estado}">'
                                        f'{"N/A" if val == 0 else f"{format_number(val, 2):.2f}"}</div>'
                                        f'<div class="combustible-unidad {estado}">Galones</div></div>',
                                        unsafe_allow_html=True,
                                    )
                else:
                    aviso_amigable("Sin datos de combustible en registro para los meses elegidos.")
            else:
                aviso_amigable("No se pudo leer la hoja registro en SharePoint.")

            st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)
            st.markdown('<div class="section-title">Participación por CECO — comparación por mes</div>', unsafe_allow_html=True)
            if "CECO" in df_raw.columns:
                sin_map = df_raw["CECO"].astype(str).str.contains("NO MAPEADO|SIN CECO", na=False).sum()
                if sin_map > 0:
                    st.caption(
                        f"ℹ {sin_map} unidad(es) con CECO vacío, código **0** o sin nombre en **TMCeco**. "
                        f"En BdMes asigne el código (ej. 1004) y en TMCeco el nombre del área."
                    )
                for mes in meses_ord:
                    etiqueta = MESES_COMPLETOS.get(mes, mes)
                    st.markdown(
                        f'<div class="trimestre-mes-bloque">'
                        f'<div class="trimestre-mes-titulo">{etiqueta}</div>',
                        unsafe_allow_html=True,
                    )
                    ceco_mes = tabla_participacion_mes(df_raw, "CECO", mes)
                    if ceco_mes.empty:
                        aviso_amigable()
                    else:
                        df_show = formatear_tabla_participacion(
                            ceco_mes.rename(columns={"Grupo": "CECO"}), simbolo, factor
                        )
                        st.dataframe(
                            df_show,
                            use_container_width=True,
                            hide_index=True,
                            height=360,
                            selection_mode="single-row",
                            key=f"ceco_{mes}",
                        )
                    st.markdown("</div>", unsafe_allow_html=True)
            else:
                aviso_amigable("CECO no disponible en BdMes.")

            st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)
        
    elif menu == VISTA_POR_FLOTA:
        show_update_info()
        if filtro_detalle and hay_datos and cols_mes_ok:
            st.markdown(f'<div class="main-title">Por Flota — {mes_sel}</div>', unsafe_allow_html=True)

            g_total = df_dashboard[c_sub].sum() * factor
            
            c1, c2, c3, c4, c5 = st.columns(5)
            render_kpi_beautiful(c1, "GASTO TOTAL", f"{simbolo} {g_total:,.2f}")
            render_kpi_beautiful(c2, "EJECUCIÓN PR.", f"{format_number((df_dashboard[c_sub].sum()/df_raw[c_sub].sum()*100 if df_raw[c_sub].sum()>0 else 0), 2):.2f}%")
            render_kpi_beautiful(c3, "RENDIMIENTO", f"{format_number(df_dashboard[c_ren].mean(), 2):.2f}", "KM/G")
            render_kpi_beautiful(c4, "KM TOTALES", f"{format_number(df_dashboard[c_km].sum(), 2):,.2f}")
            render_kpi_beautiful(c5, "DISPONIBILIDAD", f"{unidades_seleccionadas}/{total_unidades}")
            
            c6, c7, c8, c9, c10 = st.columns(5)
            render_kpi_beautiful(c6, "COSTO x KM", f"{simbolo} {format_number(g_total/max(df_dashboard[c_km].sum(), 1), 2):.2f}")
            render_kpi_beautiful(c7, "< META", f"{len(df_dashboard[df_dashboard[c_ren]<meta_kmg])}", "Und.")
            render_kpi_beautiful(c8, "EFICIENCIA ME.", f"{format_number(len(df_dashboard[df_dashboard[c_ren]>=meta_kmg])/len(df_dashboard)*100 if len(df_dashboard)>0 else 0, 2):.2f}%")
            render_kpi_beautiful(c9, "COSTO/UNIDAD", f"{simbolo} {format_number(g_total/len(df_dashboard) if len(df_dashboard)>0 else 0, 2):.2f}")
            render_kpi_beautiful(c10, "GASTO DIARIO", f"{simbolo} {format_number(g_total/30, 2):.2f}")

            st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)

            c1, c2 = st.columns([1.2, 1])
            with c1:
                tit_rend = (
                    "Rendimiento por Conductor"
                    if col_display == "CONDUCTOR"
                    else "Rendimiento por Placa"
                )
                st.markdown(f'<div class="section-title">{tit_rend}</div>', unsafe_allow_html=True)
                df_graf = df_dashboard.copy()
                df_graf["DISPLAY"] = serie_display(df_graf, col_display)
                fig_rend = px.bar(df_graf, x="DISPLAY", y=c_ren, color="MARCA", template="plotly_dark", 
                                color_discrete_sequence=px.colors.qualitative.Set2)
                fig_rend.add_hline(y=meta_kmg, line_dash="dash", line_color="#ffd700", annotation_text="Meta")
                fig_rend.update_layout(height=320, showlegend=False, hovermode='x unified', paper_bgcolor='#000000', plot_bgcolor='#000000', margin=dict(l=40, r=20, t=20, b=40))
                fig_rend.update_xaxes(showgrid=False)
                fig_rend.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#1a1a1a')
                st.plotly_chart(fig_rend, use_container_width=True, config={'displayModeBar': False})
            with c2:
                st.markdown('<div class="section-title">Distribución de Gasto</div>', unsafe_allow_html=True)
                fig_tree = px.treemap(df_dashboard, path=['TIPO', 'MARCA', 'PLACA'], values=c_sub, color=c_sub, 
                                    color_continuous_scale='Viridis', template="plotly_dark")
                fig_tree.update_layout(height=320, coloraxis_showscale=False, paper_bgcolor='#000000', plot_bgcolor='#000000', margin=dict(l=0, r=0, t=20, b=20))
                st.plotly_chart(fig_tree, use_container_width=True, config={'displayModeBar': False})

            st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)
            st.markdown('<div class="section-title">Análisis Pareto - Impacto Económico</div>', unsafe_allow_html=True)
            df_p = df_dashboard.sort_values(by=c_sub, ascending=False).copy()
            df_p["DISPLAY"] = serie_display(df_p, col_display)
            df_p['%_ACUM'] = (df_p[c_sub].cumsum() / df_p[c_sub].sum()) * 100
            fig_pareto = go.Figure()
            fig_pareto.add_trace(go.Bar(x=df_p["DISPLAY"], y=df_p[c_sub]*factor, name="Gasto", marker_color='#1f77b4'))
            fig_pareto.add_trace(go.Scatter(x=df_p["DISPLAY"], y=df_p['%_ACUM'], name="% Acum", yaxis="y2", line=dict(color="#ffd700", width=3)))
            fig_pareto.update_layout(template="plotly_dark", yaxis2=dict(overlaying="y", side="right", range=[0, 105]), height=320, paper_bgcolor='#000000', plot_bgcolor='#000000', hovermode='x unified', margin=dict(l=40, r=60, t=20, b=40), showlegend=True)
            fig_pareto.update_xaxes(showgrid=False)
            fig_pareto.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#1a1a1a')
            st.plotly_chart(fig_pareto, use_container_width=True, config={'displayModeBar': False})

            st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)
            render_grafico_km_costo_km(df_dashboard, col_display, c_km, c_sub, factor, simbolo)

            st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)
            st.markdown('<div class="section-title">Análisis Comparativo</div>', unsafe_allow_html=True)
            col_id = etiqueta_columna_id(col_display)
            t1, t2, t3, t4 = st.columns(4)
            with t1:
                st.markdown('<div class="header-top bg-success">Top 5 Eficiencia</div>', unsafe_allow_html=True)
                df_top = df_dashboard.nlargest(5, c_ren).copy()
                df_top[col_id] = serie_display(df_top, col_display)
                df_top = df_top[[col_id, c_ren]]
                df_top[c_ren] = df_top[c_ren].apply(lambda x: f"{format_number(x, 2):.2f}")
                st.dataframe(df_top, hide_index=True, use_container_width=True)
            with t2:
                st.markdown('<div class="header-top bg-danger">Críticos</div>', unsafe_allow_html=True)
                df_crit = df_dashboard.nsmallest(5, c_ren).copy()
                df_crit[col_id] = serie_display(df_crit, col_display)
                df_crit = df_crit[[col_id, c_ren]]
                df_crit[c_ren] = df_crit[c_ren].apply(lambda x: f"{format_number(x, 2):.2f}")
                st.dataframe(df_crit, hide_index=True, use_container_width=True)
            with t3:
                st.markdown(f'<div class="header-top bg-warning">Mayor Gasto</div>', unsafe_allow_html=True)
                df_g = df_dashboard.nlargest(5, c_sub).copy()
                df_g[col_id] = serie_display(df_g, col_display)
                df_g[c_sub] = df_g[c_sub].apply(lambda x: f"{simbolo} {format_number(x * factor, 2):.2f}")
                st.dataframe(df_g[[col_id, c_sub]], hide_index=True, use_container_width=True)
            with t4:
                st.markdown('<div class="header-top bg-info">Mayor KM</div>', unsafe_allow_html=True)
                df_km = df_dashboard.nlargest(5, c_km).copy()
                df_km[col_id] = serie_display(df_km, col_display)
                df_km = df_km[[col_id, c_km]]
                df_km[c_km] = df_km[c_km].apply(lambda x: f"{format_number(x, 2):.2f}")
                st.dataframe(df_km, hide_index=True, use_container_width=True)

            st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)
            render_tabla_auditoria(
                df_dashboard,
                c_km,
                c_gal,
                c_ren,
                c_sub,
                meta_kmg,
                simbolo,
                factor,
                col_display,
                mes_label=mes_sel,
            )
        elif filtro_detalle:
            aviso_amigable(
                "Sin datos para los filtros seleccionados en el mes fiscal, "
                "o revise Tipo de Flota / Marca / Modelo."
            )

        else:
            st.markdown(
                '<div class="info-box">Use los <b>filtros del panel lateral</b> (CECO, Placa, Tipo de Flota, etc.) para ver el dashboard mensual.</div>',
                unsafe_allow_html=True,
            )

    elif menu == VISTA_REDES:
        show_update_info()
        st.markdown(f'<div class="main-title">Red de Estaciones - {mes_sel}</div>', unsafe_allow_html=True)
        
        if not _registro_ok or df_estaciones is None or df_estaciones.empty:
            st.markdown(
                '<div class="warning-box">No se pudo cargar la hoja <b>Registro</b> desde SharePoint.</div>',
                unsafe_allow_html=True,
            )
        elif df_reg_vista is None or df_reg_vista.empty:
            if _redes_filtros_activos(estacion_red_sel, provincia_red_sel, placa_red_sel):
                st.warning(
                    "No hay registros con los filtros del panel lateral. "
                    "Use **Limpiar filtros Red** o amplíe Provincia / Estación / Placa."
                )
            else:
                aviso_amigable(
                    f"Sin abastecimientos en **{mes_sel}** en la hoja Registro. "
                    "Pruebe otro mes fiscal o actualice el Excel (F5)."
                )
        else:
            reg_cols = resolver_columnas_registro(df_estaciones)
            filtros_txt = []
            if provincia_red_sel:
                filtros_txt.append(f"provincia: {', '.join(provincia_red_sel)}")
            if estacion_red_sel:
                filtros_txt.append(f"estación: {', '.join(estacion_red_sel)}")
            if placa_red_sel:
                filtros_txt.append(f"placa: {', '.join(placa_red_sel)}")
            if filtros_txt:
                st.info(
                    f"Filtros activos ({mes_sel}, {st.session_state.get('dep_red_tipo', 'REDCOL')}): "
                    + " · ".join(filtros_txt)
                    + ". Los totales coinciden con Excel si aplica los mismos filtros."
                )
            prov_col = reg_cols.get("provincia")
            estado_col = reg_cols.get("estado")
            estacion_col = reg_cols.get("estacion")
            red_col = reg_cols.get("proveedor")
            galones_col = reg_cols.get("galones")
            precio_col_reg = reg_cols.get("monto")
            df_dep = (
                df_reg_redes_base
                if df_reg_redes_base is not None and not df_reg_redes_base.empty
                else df_estaciones
            )
            
            if prov_col is None:
                st.markdown(
                    '<div class="warning-box">Columna PROVINCIA no encontrada en Registro.</div>',
                    unsafe_allow_html=True,
                )
            elif galones_col is None:
                st.markdown(
                    '<div class="warning-box">Columna de galones no encontrada (ej. CANTIDAD DE GALONES).</div>',
                    unsafe_allow_html=True,
                )
            else:
                df_datos = df_reg_vista.copy()

                if red_col and red_col in df_datos.columns:
                    mask_pr = df_datos[red_col].astype(str).str.upper()
                    n_primax = int(mask_pr.str.contains("PRIMAX", na=False).sum())
                    n_redcol = int(mask_pr.str.contains("REDCOL", na=False).sum())
                else:
                    n_primax = n_redcol = 0
                st.caption(
                    f"**{mes_sel}** — {len(df_datos)} abastecimientos en registro "
                    f"(PRIMAX: {n_primax} | REDCOL: {n_redcol})."
                )

                st.markdown('<div class="section-title">Estadísticas de Estaciones</div>', unsafe_allow_html=True)
                est1, est2, est3, est4, est5 = st.columns(5)
                n_est, total_galones, n_primax, n_redcol, provincias, n_abast = calcular_kpis_estaciones(
                    df_datos,
                    red_col,
                    galones_col,
                    estacion_col,
                    prov_col,
                )
                render_kpi_beautiful(est1, "ESTACIONES ÚNICAS", f"{n_est}")
                render_kpi_beautiful(est2, "GALONES TOTALES", f"{format_number(total_galones, 2):.2f}", "GL")
                render_kpi_beautiful(est3, "ABAST. PRIMAX", f"{n_primax}")
                render_kpi_beautiful(est4, "ABAST. REDCOL", f"{n_redcol}")
                render_kpi_beautiful(est5, "PROVINCIAS", f"{provincias}")
                st.caption(
                    f"Mes **{mes_sel}**: {n_abast} cargas. "
                    "**ABAST. PRIMAX/REDCOL** = cantidad de registros, no estaciones físicas."
                )
                st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)
                render_tabla_abastecimiento_departamento(
                    df_dep,
                    prov_col,
                    galones_col,
                    precio_col_reg,
                    red_col,
                    simbolo,
                    ano=2026,
                    mes_corto=mes_sel_corto,
                    mes_label=mes_sel,
                )

                st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)
                if estacion_col and red_col and galones_col:
                    render_tabla_estaciones_red(
                        df_datos,
                        estacion_col,
                        red_col,
                        galones_col,
                        precio_col_reg,
                        simbolo,
                    )

    elif menu == VISTA_FICHA_VEHICULO:
        if not _bdmes_ok and not _vehiculos_ok:
            aviso_amigable("Cargue BdMes o la hoja TMDatosvehículo en el Excel de SharePoint.")
        else:
            render_vista_ficha_vehiculo(df_raw if _bdmes_ok else pd.DataFrame(), df_vehiculos, mes_sel_corto, simbolo)
else:
    st.markdown('<div class="main-title">LOGISTIX AI | SCE</div>', unsafe_allow_html=True)
    aviso_amigable("No se pudo cargar el Excel desde SharePoint.")
    st.markdown(
        """
        **En Render (nube):**
        1. **Environment** → variable `EXCEL_URL` = enlace completo de descarga del Excel.  
        2. En SharePoint: **Compartir** → *Cualquier persona con el enlace* (sin solo usuarios de la empresa).  
        3. **Manual Deploy** después de guardar la variable.
        """
    )
    st.caption(f"Fuente configurada: **{_origen_excel_url()}**")
    if _errores_carga:
        for err in _errores_carga:
            st.error(err)
    else:
        st.error(
            "Sin detalle del error. Suba el último código a GitHub y haga Deploy en Render."
        )
    if st.button("Reintentar conexión (limpiar caché)"):
        st.cache_data.clear()
        st.rerun()

st.markdown(
    '<div class="watermark-turromzita">Turromzita &lt;333333</div>',
    unsafe_allow_html=True,
)
