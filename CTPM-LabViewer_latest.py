# CTPM_LabViewer v5.4.2 — UPDATED7f
# Changes: Altair theme compatibility (alt.theme/alt.themes), fixed map JSON, consistent chart heights via theme (280),
# removed use_container_width, scheduling policy (1→2→>2 only to avoid 2nd week), Aging page retained.

#CTPM_LabViewer v5.4.1-hotfix1 — UPDATED4 (Policy DB + Scheduling + #Map restored)
# CTPM_LabViewer v5.4.1-hotfix1 — UPDATED4
# Adds back Policy DB (persistent), Scheduling & On‑Site Map pages with full functions,
# geocoding helpers, ICS export, and chart exclusion of CTPM.

from __future__ import annotations
import warnings
import os, json, base64, socket, sqlite3, hashlib, math, re
from io import BytesIO
from pathlib import Path
from typing import Dict, Optional, Tuple, List
from datetime import datetime

import pandas as pd
import numpy as np
import altair as alt
import plotly.express as px
import streamlit as st
import pgeocode



# --- Warning hygiene (keep logs clean; avoid future breakages) ---
# openpyxl: some IndySoft exports omit default style
warnings.filterwarnings(
    "ignore",
    message=r"Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
)
# pandas: groupby.apply include_groups upcoming behavior change
warnings.filterwarnings(
    "ignore",
    message=r"DataFrameGroupBy.apply operated on the grouping columns",
    category=FutureWarning,
)
# --- WO / Tech Efficiency (All WOs export) ---

try:
    import wo_efficiency as we

    build_efficiency_datasets = we.build_efficiency_datasets
    compute_skip_trend_per_customer = we.compute_skip_trend_per_customer
    streamlit_customer_skip_trend_panel = we.streamlit_customer_skip_trend_panel
    streamlit_tech_drilldown_panel = we.streamlit_tech_drilldown_panel

    # Optional helper (won't crash if missing)
    add_effective_date_columns = getattr(we, "add_effective_date_columns", None)

except Exception as e:
    build_efficiency_datasets = None
    compute_skip_trend_per_customer = None
    streamlit_customer_skip_trend_panel = None
    streamlit_tech_drilldown_panel = None
    add_effective_date_columns = None

    import traceback
    print("FAILED importing wo_efficiency:", e)
    traceback.print_exc()


# ======= App constants =======

APP_TITLE = "CTPM_LabViewer"
DEFAULT_DATA_FILE = r"C:\\Users\\bholl\\Documents\\CTPM-Calibration-Dashboard\\All Data.xlsx"
DEFAULT_CERT_ROOT = r"S:\\Shared With Me\\CTPM"
LOGO_PATH = "CTPM-Logo_thumbnail reduced.png"
CACHE_DIR = Path(os.environ.get("CTPM_CACHE_DIR", ".ctpm_cache")); CACHE_DIR.mkdir(parents=True, exist_ok=True)
PARQUET_DIR = CACHE_DIR / "parquet"; PARQUET_DIR.mkdir(parents=True, exist_ok=True)
DEFAULT_NOTES_DB = str(CACHE_DIR / ".ctpm_notes.sqlite")
LAB_ZIP_DEFAULT = os.environ.get("CTPM_LAB_ZIP", "45430")
PRESETS_FILE = Path('.ctpm_presets.json')
THEMES_FILE = Path('.ctpm_themes.json')
SETTINGS_FILE = Path(os.environ.get("CTPM_SETTINGS_FILE", ".ctpm_settings.json"))
POLICY_DB = Path(os.environ.get("CTPM_POLICY_DB", ".ctpm_settings.sqlite3"))
APP_DIR = Path(__file__).resolve().parent


def _sig_hash(sig) -> str:
    """Short hash of any hashable sig used as Parquet cache key."""
    return hashlib.md5(repr(sig).encode()).hexdigest()[:12]


# ---- Individual-sheet file helpers ----
_SHEET_KEYS = ["wos", "events", "equipment", "all_companies", "wip_shop"]

def _upl_sheet_paths() -> Dict[str, Optional[Path]]:
    """Return paths to individually-uploaded sheet files (None if not present)."""
    upl_dir = CACHE_DIR / "uploads"
    result: Dict[str, Optional[Path]] = {}
    for key in _SHEET_KEYS:
        found = None
        for ext in (".xlsx", ".xls"):
            p = upl_dir / f"{key}{ext}"
            if p.exists():
                found = p
                break
        result[key] = found
    return result


def _individual_sheets_sig() -> tuple:
    """Stable cache key derived from all individual sheet files present."""
    sheets = _upl_sheet_paths()
    parts = []
    for key in _SHEET_KEYS:
        p = sheets.get(key)
        if p and p.exists():
            s = p.stat()
            parts.append((float(s.st_mtime), int(s.st_size)))
        else:
            parts.append((0.0, 0))
    return tuple(parts)


def _read_sheet_file(p: Path, keep: list) -> pd.DataFrame:
    """Read first sheet from an individual export file, keeping only known columns."""
    engine = "xlrd" if p.suffix.lower() == ".xls" else "openpyxl"
    hdr = pd.read_excel(p, nrows=0, engine=engine)
    cols = [c for c in keep if c in hdr.columns]
    if cols:
        return pd.read_excel(p, usecols=cols, engine=engine)
    return pd.read_excel(p, engine=engine)


# ---- AWS S3 helpers (cloud persistence, optional) ----
def _s3_bucket_name() -> str:
    return os.environ.get("CTPM_S3_BUCKET", "")


def _s3_client():
    """Return a boto3 S3 client when CTPM_S3_BUCKET is set, else None."""
    if not _s3_bucket_name():
        return None
    try:
        import boto3  # noqa: PLC0415
        return boto3.client("s3")
    except Exception:
        return None


def _s3_upload(local_path: Path, s3_key: Optional[str] = None) -> bool:
    """Upload a local file to S3. Returns True on success."""
    s3 = _s3_client()
    bucket = _s3_bucket_name()
    if not s3 or not bucket:
        return False
    key = s3_key or f"uploads/{local_path.name}"
    try:
        s3.upload_file(str(local_path), bucket, key)
        return True
    except Exception:
        return False


def _s3_download(s3_key: str, local_path: Path) -> bool:
    """Download a file from S3 to local_path. Returns True on success."""
    s3 = _s3_client()
    bucket = _s3_bucket_name()
    if not s3 or not bucket:
        return False
    try:
        local_path.parent.mkdir(parents=True, exist_ok=True)
        s3.download_file(bucket, s3_key, str(local_path))
        return True
    except Exception:
        return False


def _sync_from_s3():
    """On cold start: restore uploaded sheet files and users.db from S3."""
    s3 = _s3_client()
    bucket = _s3_bucket_name()
    if not s3 or not bucket:
        return
    upl_dir = CACHE_DIR / "uploads"
    upl_dir.mkdir(parents=True, exist_ok=True)
    for key in _SHEET_KEYS:
        for ext in (".xlsx", ".xls"):
            local = upl_dir / f"{key}{ext}"
            if not local.exists():
                try:
                    s3.download_file(bucket, f"uploads/{key}{ext}", str(local))
                    break
                except Exception:
                    pass
    # Restore user database
    users_db_path = CACHE_DIR / "users.db"
    if not users_db_path.exists():
        try:
            s3.download_file(bucket, "userdata/users.db", str(users_db_path))
        except Exception:
            pass


DEFAULT_SETTINGS = {
    "hours_per_item_default": 0.5,
    "lab_zip":"45430",
    "near_radius":30, "far_radius":80,
    "near_window_days":30, "mid_window_days":45, "far_window_months":3,
    "techs":15, "hours_per_tech":8,
    "vehicles":6, "techs_per_vehicle":2,
    "horizon_months":6, "weekdays_only": True,
    "cluster_enable": True, "cluster_by_zip3": False, "cluster_miles": 10.0,
    "small_hours_threshold": 6.0,
    "assumed_mph": 40.0, "per_stop_travel_hours": 0.25,
    "late_penalty_per_day": 5.0,
    "capacity_weeks": 8
}

# ======= Branding / Theme =======
THEMES = {
    # Premium defaults
    "CTPM Light Premium": {"PRIMARY":"#7E1F23", "SECONDARY":"#9B3A3A","ACCENT":"#F28A94","DARK":"#0E0B0B","BG":"#F5F6FA","TEXT":"#111318","INVERT":"#FFFFFF"},
    "CTPM Dark Premium":  {"PRIMARY":"#F28A94", "SECONDARY":"#B55C5C","ACCENT":"#FFD166","DARK":"#0B0F1A","BG":"#0B0F1A","TEXT":"#EAF0FF","INVERT":"#0B0F1A"},

    # Backward-compatible names (mapped to premium variants)
    "CTPM High Contrast": {"PRIMARY":"#7E1F23", "SECONDARY":"#9B3A3A","ACCENT":"#F28A94","DARK":"#0E0B0B","BG":"#F5F6FA","TEXT":"#111318","INVERT":"#FFFFFF"},
    "Dark": {"PRIMARY":"#F28A94","SECONDARY":"#B55C5C","ACCENT":"#FFD166","DARK":"#0B0F1A","BG":"#0B0F1A","TEXT":"#EAF0FF","INVERT":"#0B0F1A"},
    "Light": {"PRIMARY":"#7E1F23", "SECONDARY":"#B55C5C","ACCENT":"#E4E8F5","DARK":"#0E0B0B","BG":"#F5F6FA","TEXT":"#111318","INVERT":"#FFFFFF"},
}
FONT_FAMILY = "'Inter', 'Roboto', 'Open Sans'"; BASE_FONT_PX = 16

# ======= Theme helpers (robust session_state defaults) =======
REQUIRED_THEME_KEYS = ["PRIMARY","SECONDARY","ACCENT","DARK","BG","TEXT","INVERT"]

def get_active_theme() -> Dict[str,str]:
    """Return an always-valid theme dict.

    Fixes rare cases where session_state is missing keys during reconnects/reruns.
    Never raises KeyError.
    """
    base = THEMES.get('CTPM Light Premium', next(iter(THEMES.values())))
    t = st.session_state.get('active_theme')
    if not isinstance(t, dict):
        name = st.session_state.get('sb_theme_name', 'CTPM Light Premium')
        t = THEMES.get(name, base)
        st.session_state['active_theme'] = t
    # ensure required keys exist
    for k in REQUIRED_THEME_KEYS:
        if k not in t and k in base:
            t[k] = base[k]
    return t


# ======= Utils =======
def sanitize_path(p: str) -> str:
    return str(p).strip().strip('"').strip("'")

def resolve_default_data_file() -> str:
    return sanitize_path(os.environ.get("CTPM_DATA_FILE", DEFAULT_DATA_FILE))

def resolve_default_notes_db(data_file_path: str) -> str:
    try:
        dfp = Path(sanitize_path(data_file_path)); sibling = str(dfp.parent / ".ctpm_notes.sqlite")
    except Exception:
        sibling = str(APP_DIR / ".ctpm_notes.sqlite")
    return sanitize_path(os.environ.get("CTPM_NOTES_DB", sibling))

def resolve_default_cert_root() -> str:
    return sanitize_path(os.environ.get("CTPM_CERT_ROOT", DEFAULT_CERT_ROOT))

def file_signature(path: str) -> Tuple[float, int]:
    p = Path(sanitize_path(path)); s = p.stat(); return (s.st_mtime, s.st_size)

# Exclude CTPM from CHARTS (not tables)
def charts_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty: return df
    if 'Company' in df.columns:
        return df[df['Company'].astype(str).str.upper().ne('CTPM')]
    return df

# ======= Theme & CSS =======
def _inject_css(theme: Dict[str,str], kiosk_mode: bool=False):
    """Inject premium CTPM styling. kiosk_mode increases readability and hides chrome."""
    # Surface colors derived from BG
    bg = theme['BG']
    textc = theme['TEXT']
    is_dark = bg.lower() == "#0b0f1a"
    surface = "rgba(255,255,255,0.78)" if not is_dark else "rgba(255,255,255,0.06)"
    surface2 = "rgba(255,255,255,0.60)" if not is_dark else "rgba(255,255,255,0.10)"
    border = "#00000014" if not is_dark else "#FFFFFF20"
    surface_elevated = "rgba(255,255,255,0.92)" if not is_dark else "rgba(255,255,255,0.08)"
    primary_glow = theme['PRIMARY'] + ("20" if not is_dark else "30")
    input_bg   = "#FFFFFF" if not is_dark else "#1A1F2E"
    popover_bg = "#FFFFFF" if not is_dark else "#161B2A"

    kiosk_css = ""
    if kiosk_mode:
        kiosk_css = f"""
        /* Kiosk mode: hide Streamlit chrome and increase tap targets */
        [data-testid=\"stToolbar\"], [data-testid=\"stStatusWidget\"], #MainMenu, footer {{ display:none !important; }}
        header {{ visibility:hidden; height:0px; }}
        .block-container {{ padding-top: 0.75rem !important; padding-bottom: 1.5rem !important; }}
        html, body, .stApp {{ font-size: 17px !important; }}
        """

    css = f"""
    :root {{
      --ctpm-primary:{theme['PRIMARY']};
      --ctpm-secondary:{theme['SECONDARY']};
      --ctpm-accent:{theme['ACCENT']};
      --ctpm-dark:{theme['DARK']};
      --ctpm-bg:{bg};
      --ctpm-text:{textc};
      --ctpm-invert:{theme['INVERT']};
      --ctpm-surface:{surface};
      --ctpm-surface2:{surface2};
      --ctpm-border:{border};
      --ctpm-surface-elevated:{surface_elevated};
      --ctpm-shadow-sm: 0 2px 8px rgba(0,0,0,0.06);
      --ctpm-shadow-md: 0 8px 24px rgba(0,0,0,0.10);
      --ctpm-shadow-lg: 0 16px 48px rgba(0,0,0,0.12);
      --ctpm-radius-sm: 8px;
      --ctpm-radius-md: 14px;
      --ctpm-radius-lg: 20px;
      --ctpm-transition: 0.2s cubic-bezier(0.4, 0, 0.2, 1);
    }}

    /* Always hide Streamlit's built-in toolbar — app has its own Tools menu */
    [data-testid=\"stToolbar\"], [data-testid=\"stStatusWidget\"],
    [data-testid=\"stDeployButton\"], [data-testid=\"stAppDeployButton\"],
    #MainMenu, footer {{ display: none !important; }}

    html, body, .stApp {{
      background: var(--ctpm-bg) !important;
      color: var(--ctpm-text) !important;
      font-family: "Segoe UI Variable","Segoe UI",Inter,Roboto,"Helvetica Neue",Arial,sans-serif !important;
      font-size: {BASE_FONT_PX}px !important;
      line-height: 1.35;
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
      text-rendering: geometricPrecision;
    }}

    /* App background gradient (subtle, 3-layer) */
    [data-testid=\"stAppViewContainer\"] {{
      background: radial-gradient(1200px 600px at 15% 0%, {theme['ACCENT']}18, transparent 50%),
                  radial-gradient(1000px 650px at 100% 10%, {theme['SECONDARY']}12, transparent 55%),
                  radial-gradient(800px 400px at 50% 100%, {theme['PRIMARY']}08, transparent 40%),
                  var(--ctpm-bg) !important;
    }}

    /* Sidebar styling */
    section[data-testid=\"stSidebar\"] > div:first-child {{
      background: linear-gradient(180deg, {theme['DARK']}, {theme['PRIMARY']}) !important;
    }}

    /* Header accent bar */
    .app-header-bar {{
      height: 4px;
      background: linear-gradient(90deg, var(--ctpm-primary), var(--ctpm-accent), var(--ctpm-secondary));
      border-radius: 999px;
      margin: 8px 0 20px 0;
      box-shadow: 0 4px 16px {primary_glow};
      opacity: 0.85;
    }}

    /* Primary buttons (stBaseButton-primary is the testid Streamlit 1.40+ uses) */
    button[data-testid=\"stBaseButton-primary\"],
    div[data-testid=\"stButton\"] > button {{
      background-color: var(--ctpm-primary) !important;
      color: #fff !important;
      border-radius: var(--ctpm-radius-sm) !important;
      font-weight: 700 !important;
      border: none !important;
      box-shadow: var(--ctpm-shadow-sm) !important;
      padding: 8px 20px !important;
      transition: all var(--ctpm-transition) !important;
      letter-spacing: 0.2px !important;
    }}
    button[data-testid=\"stBaseButton-primary\"]:hover,
    div[data-testid=\"stButton\"] > button:hover {{
      filter: brightness(1.15) !important;
      transform: translateY(-1px) !important;
      box-shadow: var(--ctpm-shadow-md) !important;
    }}
    /* Secondary buttons — nav inactive items, etc. */
    button[data-testid=\"stBaseButton-secondary\"] {{
      background-color: var(--ctpm-surface-elevated) !important;
      color: var(--ctpm-text) !important;
      border: 1px solid var(--ctpm-border) !important;
      border-radius: var(--ctpm-radius-sm) !important;
      font-weight: 600 !important;
      transition: all var(--ctpm-transition) !important;
    }}
    button[data-testid=\"stBaseButton-secondary\"]:hover {{
      border-color: var(--ctpm-primary) !important;
      color: var(--ctpm-primary) !important;
    }}
    /* Popover trigger buttons (data-testid="stPopoverButton" in Streamlit 1.40+) */
    button[data-testid=\"stPopoverButton\"] {{
      background-color: var(--ctpm-surface-elevated) !important;
      color: var(--ctpm-text) !important;
      border: 1px solid var(--ctpm-border) !important;
      border-radius: var(--ctpm-radius-sm) !important;
      font-weight: 600 !important;
      transition: all var(--ctpm-transition) !important;
    }}
    button[data-testid=\"stPopoverButton\"]:hover {{
      border-color: var(--ctpm-primary) !important;
      color: var(--ctpm-primary) !important;
    }}

    /* Download buttons — outline style */
    div[data-testid=\"stDownloadButton\"] > button {{
      background-color: transparent !important;
      color: var(--ctpm-primary) !important;
      border: 1.5px solid var(--ctpm-primary) !important;
      border-radius: var(--ctpm-radius-sm) !important;
      font-weight: 700 !important;
      transition: all var(--ctpm-transition) !important;
    }}
    div[data-testid=\"stDownloadButton\"] > button:hover {{
      background-color: var(--ctpm-primary) !important;
      color: #fff !important;
      transform: translateY(-1px) !important;
    }}

    /* Cards / KPIs */
    .kpi {{
      padding: 16px 20px;
      border: 1px solid var(--ctpm-border);
      border-radius: var(--ctpm-radius-md);
      background: var(--ctpm-surface-elevated);
      box-shadow: var(--ctpm-shadow-md);
      backdrop-filter: blur(12px) saturate(1.2);
      transition: transform var(--ctpm-transition), box-shadow var(--ctpm-transition);
      position: relative;
      overflow: hidden;
    }}
    .kpi:hover {{
      transform: translateY(-2px);
      box-shadow: var(--ctpm-shadow-lg);
    }}
    .kpi::before {{
      content: '';
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 3px;
      background: linear-gradient(90deg, var(--ctpm-primary), var(--ctpm-accent));
      border-radius: var(--ctpm-radius-md) var(--ctpm-radius-md) 0 0;
    }}
    .kpi .label {{
      font-size: 0.78rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: var(--ctpm-text);
      opacity: 0.6;
      margin-bottom: 4px;
    }}
    .kpi .value {{
      font-size: 1.6rem;
      font-weight: 900;
      color: var(--ctpm-primary);
      line-height: 1.2;
    }}

    /* Dataframe/table container */
    .stDataFrame, .stTable {{
      border-radius: var(--ctpm-radius-md) !important;
      border: 1px solid var(--ctpm-border) !important;
      background: var(--ctpm-surface-elevated) !important;
      box-shadow: var(--ctpm-shadow-sm) !important;
      overflow: hidden !important;
    }}

    /* Tab buttons — ensure readable text on both active and inactive states */
    button[role=\"tab\"] {{
      color: var(--ctpm-text) !important;
      font-weight: 600 !important;
    }}
    button[role=\"tab\"][aria-selected=\"true\"] {{
      color: var(--ctpm-primary) !important;
      border-bottom-color: var(--ctpm-primary) !important;
    }}
    button[role=\"tab\"][aria-selected=\"false\"] {{
      opacity: 0.65;
    }}
    button[role=\"tab\"]:hover {{
      opacity: 1 !important;
      color: var(--ctpm-primary) !important;
    }}

    /* Inputs */
    div[data-baseweb=\"select\"] > div, input, textarea {{
      border-radius: var(--ctpm-radius-sm) !important;
      transition: border-color var(--ctpm-transition), box-shadow var(--ctpm-transition) !important;
    }}
    input:focus, textarea:focus {{
      border-color: var(--ctpm-primary) !important;
      box-shadow: 0 0 0 3px {primary_glow} !important;
    }}

    /* Widen popovers so policy table columns are fully visible */
    div[data-testid=\"stPopoverBody\"] {{
      min-width: 580px !important;
    }}

    /* Expander summary text always readable */
    [data-testid=\"stExpander\"] summary,
    [data-testid=\"stExpander\"] summary p,
    [data-testid=\"stExpander\"] summary span {{
      color: {textc} !important;
      background: {surface_elevated} !important;
    }}

    /* Expander cards */
    [data-testid=\"stExpander\"] {{
      border: 1px solid var(--ctpm-border) !important;
      border-radius: var(--ctpm-radius-md) !important;
      background: var(--ctpm-surface) !important;
      box-shadow: var(--ctpm-shadow-sm) !important;
      margin-bottom: 8px !important;
      transition: box-shadow var(--ctpm-transition), border-color var(--ctpm-transition) !important;
      overflow: hidden !important;
    }}
    [data-testid=\"stExpander\"]:hover {{
      box-shadow: var(--ctpm-shadow-md) !important;
      border-color: var(--ctpm-primary) !important;
    }}
    [data-testid=\"stExpander\"] summary {{
      font-weight: 700 !important;
      padding: 12px 16px !important;
      font-size: 0.95rem !important;
    }}
    [data-testid=\"stExpander\"] > div[data-testid=\"stExpanderDetails\"] {{
      padding: 0 16px 16px !important;
      border-top: 1px solid var(--ctpm-border) !important;
    }}

    /* Titles */
    .section-title {{
      font-weight: 900;
      color: var(--ctpm-primary);
      letter-spacing: 0.3px;
      margin: 16px 0 10px;
      font-size: 1.1rem;
      position: relative;
      padding-left: 14px;
    }}
    .section-title::before {{
      content: '';
      position: absolute;
      left: 0;
      top: 2px;
      bottom: 2px;
      width: 4px;
      background: linear-gradient(180deg, var(--ctpm-primary), var(--ctpm-accent));
      border-radius: 999px;
    }}

    /* ── Widget background/text overrides ──────────────────────────────────
       Streamlit's BaseWeb components keep their own background independent of
       our CSS variables.  Force them to match the active theme so they never
       appear black on a white page (or white on a dark page).               */

    /* Text inputs */
    div[data-baseweb="input"],
    div[data-baseweb="input"] > div {{
      background-color: {input_bg} !important;
      color: {textc} !important;
    }}
    div[data-baseweb="input"] input {{
      color: {textc} !important;
      caret-color: {theme['PRIMARY']} !important;
    }}

    /* Selectbox / Multiselect containers */
    div[data-baseweb="select"] > div,
    div[data-baseweb="select"] > div > div {{
      background-color: {input_bg} !important;
      color: {textc} !important;
    }}
    /* Dropdown list panel */
    ul[data-baseweb="menu"],
    div[data-baseweb="popover"] > div {{
      background-color: {input_bg} !important;
      color: {textc} !important;
    }}
    ul[data-baseweb="menu"] li {{
      color: {textc} !important;
    }}
    ul[data-baseweb="menu"] li:hover {{
      background-color: {theme['PRIMARY']}18 !important;
    }}

    /* Number / text area inputs */
    div[data-baseweb="textarea"] > div,
    div[data-baseweb="textarea"] textarea {{
      background-color: {input_bg} !important;
      color: {textc} !important;
    }}

    /* Widget labels */
    label[data-testid="stWidgetLabel"] > div,
    div[data-testid="stWidgetLabel"] > div,
    label[data-testid="stWidgetLabel"] p,
    div[data-testid="stWidgetLabel"] p {{
      color: {textc} !important;
      opacity: 0.85;
    }}

    /* Streamlit popover panel (the floating card that opens) */
    div[data-testid="stPopover"],
    div[data-testid="stPopoverBody"] {{
      background-color: {popover_bg} !important;
      color: {textc} !important;
      border: 1px solid {border} !important;
      border-radius: var(--ctpm-radius-md) !important;
      box-shadow: var(--ctpm-shadow-lg) !important;
    }}
    div[data-testid="stPopover"] p,
    div[data-testid="stPopover"] label,
    div[data-testid="stPopover"] span {{
      color: {textc} !important;
    }}

    /* Slider track and thumb */
    div[data-testid="stSlider"] div[data-baseweb="slider"] div[role="slider"] {{
      background-color: {theme['PRIMARY']} !important;
    }}

    /* Checkbox / Toggle labels */
    div[data-testid="stCheckbox"] label span,
    div[data-testid="stToggle"] label span {{
      color: {textc} !important;
    }}

    /* Caption / small text */
    div[data-testid="stCaptionContainer"] p {{
      color: {textc} !important;
      opacity: 0.6;
    }}

    {kiosk_css}
    """
    st.markdown("<style>" + css + "</style>", unsafe_allow_html=True)


ALT_THEME_NAME = "ctpm"

def _altair_theme_config():
    theme = st.session_state.get('active_theme', THEMES['CTPM Light Premium'])
    return {
        "config": {
            "background": theme['DARK'],
            "view": {"strokeOpacity": 0, "fill": theme['DARK'], "continuousWidth": 1100, "continuousHeight": 280},
            "axis": {"labelColor": "#FFFFFF", "titleColor": "#FFFFFF", "gridColor": "#FFFFFF22",
                     "domainColor":"#FFFFFF55", "tickColor":"#FFFFFF55"},
            "legend": {"titleColor": "#FFFFFF", "labelColor": "#FFFFFF"},
            "range": {
                "category": [ theme['PRIMARY'], theme['ACCENT'], "#2C7FB8", "#7FC97F",
                               "#FDC086", "#386CB0", "#F0027F", "#1B9E77" ]
            }
        }
    }

try:
    # Altair >= 5.5 prefers alt.theme API
    if hasattr(alt, 'theme') and hasattr(alt.theme, 'register'):
        @alt.theme.register(ALT_THEME_NAME, enable=True)
        def _ctpm_theme():
            cfg = _altair_theme_config()
            return alt.theme.ThemeConfig(cfg.get('config', {}))
    else:
        alt.themes.register(ALT_THEME_NAME, _altair_theme_config)
        alt.themes.enable(ALT_THEME_NAME)
except Exception:
    pass
# ======= Header =======
def brand_header():
    col_logo, col_title, col_tools, col_account = st.columns([1,6,2,2])
    with col_logo:
        try: st.image(LOGO_PATH, )
        except Exception: st.write("")
    with col_title:
        st.markdown(
            f"<h1 style='margin-bottom:0;color:var(--ctpm-primary);'>{APP_TITLE}</h1>"
            f"<div style='color:var(--ctpm-secondary);font-weight:800;'>CTPM • Calibration Management</div>",
            unsafe_allow_html=True,
        )
    with col_tools:
        with st.popover("⚙️ Tools", ):
            st.caption("Admin")
            if st.button("🧹 Clear caches", key="btn_clear_cache"):
                st.cache_data.clear()
                try:
                    for f in CACHE_DIR.glob("*.pkl"): f.unlink(missing_ok=True)
                except Exception: pass
                st.session_state.pop("__warmed_key__", None); st.rerun()
            if st.button("🔄 Reload", key="btn_reload"):
                st.cache_data.clear(); st.rerun()
            
            with st.expander("Data file", expanded=False):
                new_path = st.text_input("Path to All Data.xlsx", value=st.session_state['data_file'])
                if new_path and new_path != st.session_state['data_file']:
                    st.session_state['data_file'] = new_path
                    st.success(f"Data file set to: {new_path}")
                    st.rerun()

            st.divider()
            # Appearance
            st.caption("Appearance")
            dark_mode = st.toggle("🌙 Dark mode", value=bool(st.session_state.get('dark_mode', False)), key='tog_dark_mode')
            kiosk_mode = st.toggle("🖥️ Kiosk mode", value=bool(st.session_state.get('kiosk_mode', False)), key='tog_kiosk_mode')
            st.session_state['dark_mode'] = bool(dark_mode)
            st.session_state['kiosk_mode'] = bool(kiosk_mode)

            # Map dark/light to premium themes
            preferred = "CTPM Dark Premium" if dark_mode else "CTPM Light Premium"
            if st.session_state.get('sb_theme_name') not in THEMES:
                st.session_state['sb_theme_name'] = preferred
            if st.session_state.get('sb_theme_name') != preferred:
                st.session_state['sb_theme_name'] = preferred

            # Optional theme override (advanced)
            theme_choice = st.selectbox("Theme (advanced)", list(THEMES.keys()),
                index=list(THEMES.keys()).index(st.session_state.get('sb_theme_name', preferred)))
            if theme_choice != st.session_state.get('sb_theme_name'):
                st.session_state['sb_theme_name'] = theme_choice

            st.session_state['active_theme'] = THEMES[st.session_state['sb_theme_name']]
            _inject_css(get_active_theme(), kiosk_mode=st.session_state.get('kiosk_mode', False))
            # If advanced theme changed, rerun to repaint charts
            if theme_choice != preferred:
                st.rerun()
            if st.button('🔥 Warm caches', key='btn_warm'):
                try: st.cache_data.clear()
                except Exception: pass
                try:
                    path = sanitize_path(st.session_state['data_file']); sig = file_signature(path)
                    _wos, _events, _equip = load_clean_data(path, sig)
                    st.session_state['__aging__']  = compute_aging_from_file(path, sig)
                    st.session_state['__wip__']    = wip_chain(path, sig, _equip)
                    st.session_state['__tatroll__']= rolling_tat_365d_from_file(path, sig)
                    st.success('Caches warmed.')
                except Exception as _ex:
                    st.warning(f'Warmup skipped: {_ex}')

    with col_account:
        auth = st.session_state.get("auth", {})
        if auth.get("is_authenticated"):
            st.caption(f"Signed in as **{auth.get('username','')}** ({auth.get('role','')})")
            if st.button("🚪 Log out", key="btn_logout_header"):
                st.session_state.pop("auth", None); st.rerun()
        else:
            st.caption("Not signed in")
    st.markdown("<div class='app-header-bar'></div>", unsafe_allow_html=True)

# ======= Exports =======
def download_buttons(df: pd.DataFrame, base_name: str, key_prefix: str) -> None:
    if df is None or df.empty: st.caption("No data to export."); return
    csv_bytes = df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ CSV", data=csv_bytes, file_name=f"{base_name}.csv", mime="text/csv", key=f"{key_prefix}_csv")
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w: df.to_excel(w, index=False, sheet_name="Data")
    st.download_button("⬇️ Excel", data=bio.getvalue(), file_name=f"{base_name}.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key=f"{key_prefix}_xlsx")

# ==============================================================
# ----------------------- NOTES (SQLITE) ------------------------
# ==============================================================
# Notes are EID-only: entity_type is always "EID", key is always the Equipment ID in column "I.D.".
def init_notes_db(db_path: str) -> None:
    p = Path(db_path)
    p.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(p)
    try:
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS notes (
              entity_type TEXT NOT NULL,
              entity_key  TEXT NOT NULL,
              note        TEXT NOT NULL,
              updated_ts  TEXT NOT NULL,
              author      TEXT NOT NULL,
              PRIMARY KEY (entity_type, entity_key)
            )
            """
        )
        con.commit()
    finally:
        con.close()

def get_notes_map(db_path: str, keys: List[str]) -> Dict[str, str]:
    keys = [str(k) for k in keys if str(k).strip()]
    if not keys:
        return {}
    init_notes_db(db_path)
    con = sqlite3.connect(db_path)
    try:
        out: Dict[str, str] = {}
        chunk = 400  # keep under SQLite variable limit
        for i in range(0, len(keys), chunk):
            kk = keys[i:i+chunk]
            qmarks = ",".join(["?"] * len(kk))
            rows = con.execute(
                f"SELECT entity_key, note FROM notes WHERE entity_type='EID' AND entity_key IN ({qmarks})",
                kk,
            ).fetchall()
            out.update({k: n for (k, n) in rows})
        return out
    finally:
        con.close()

def upsert_notes(db_path: str, notes: Dict[str, str], author: str = "Bill") -> None:
    if not notes:
        return
    init_notes_db(db_path)
    con = sqlite3.connect(db_path)
    try:
        now = pd.Timestamp.now().isoformat()
        rows = [("EID", str(k), str(v or ""), now, author) for k, v in notes.items()]
        con.executemany(
            """
            INSERT INTO notes(entity_type, entity_key, note, updated_ts, author)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(entity_type, entity_key)
            DO UPDATE SET note=excluded.note, updated_ts=excluded.updated_ts, author=excluded.author
            """,
            rows,
        )
        con.commit()
    finally:
        con.close()


def attach_notes_column(df: pd.DataFrame, db_path: str, note_col: str = "Notes") -> pd.DataFrame:
    """Attach notes to a dataframe that already has an 'I.D.' column."""
    if df is None or df.empty or "I.D." not in df.columns:
        return df
    keys = df["I.D."].astype(str).fillna("").tolist()
    m = get_notes_map(db_path, keys)
    out = df.copy()
    out[note_col] = out["I.D."].astype(str).map(lambda k: m.get(str(k), ""))
    return out


def attach_eid_notes_if_present(df: pd.DataFrame, db_path: str, note_col: str = "Notes") -> pd.DataFrame:
    """Auto-attach Equipment-ID notes to any dataframe that contains an 'I.D.' column."""
    if df is None or (hasattr(df, "empty") and df.empty):
        return df
    if not isinstance(df, pd.DataFrame):
        try:
            df = pd.DataFrame(df)
        except Exception:
            return df
    if "I.D." not in df.columns:
        return df
    return attach_notes_column(df, db_path, note_col=note_col)


def notes_table_editor(df: pd.DataFrame, db_path: str, editor_key: str, help_text: str = "") -> None:
    """Editable notes table for Equipment IDs. Requires an 'I.D.' column."""
    if df is None or df.empty or "I.D." not in df.columns:
        st.caption("No Equipment IDs available for notes.")
        return

    view = df.copy()
    view = view[["I.D."]].drop_duplicates()
    view = attach_notes_column(view, db_path, note_col="Notes")

    st.caption(help_text or "Edit notes directly in the table, then click Save.")

    edited = st.data_editor(
        view,
        key=editor_key,
        hide_index=True,
        column_config={
            "I.D.": st.column_config.TextColumn("I.D.", disabled=True),
            "Notes": st.column_config.TextColumn("Notes", help="Saved notes will appear anywhere this I.D. appears."),
        },
        disabled=["I.D."],
        num_rows="fixed",
    )

    if st.button("💾 Save notes changes", key=f"{editor_key}_save"):
        base_notes = view.set_index("I.D.")["Notes"].astype(str)
        new_notes = edited.set_index("I.D.")["Notes"].astype(str)
        changes = (new_notes != base_notes)
        changed_keys = new_notes.index[changes].tolist()
        payload = {str(k): str(new_notes.loc[k]) for k in changed_keys}
        upsert_notes(db_path, payload, author="")
        st.success(f"Saved {len(payload)} note(s).")
        st.rerun()

# ==============================================================
# ------------------ TABLE RENDERING (AUTO NOTES) --------------
# ==============================================================
def ensure_df(obj) -> pd.DataFrame:
    if obj is None:
        return pd.DataFrame()
    if isinstance(obj, pd.DataFrame):
        return obj
    try:
        return pd.DataFrame(obj)
    except Exception:
        return pd.DataFrame()

def build_date_column_config(df: pd.DataFrame) -> Dict[str, object]:
    """Build Streamlit column_config so datetime columns always show year (YYYY-MM-DD)."""
    cfg: Dict[str, object] = {}
    if df is None or df.empty:
        return cfg
    for c in df.columns:
        try:
            s = df[c]
        except Exception:
            continue
        if pd.api.types.is_datetime64_any_dtype(s):
            cfg[c] = st.column_config.DatetimeColumn(c, format="YYYY-MM-DD")
    return cfg


def with_notes(obj, db_path: str) -> pd.DataFrame:
    df = ensure_df(obj)
    return attach_eid_notes_if_present(df, db_path)


def show_table(obj, db_path: str, *, max_rows: int, table_key: str,
               do_export: bool = False, export_name: str = "export", export_key: str = "export") -> pd.DataFrame:
    """Show a dataframe with auto-notes if it has I.D. Returns the dataframe used (includes Notes if applicable)."""
    df = with_notes(obj, db_path)
    st.dataframe(df.head(max_rows), key=table_key, column_config=build_date_column_config(df))
    if do_export:
        download_buttons(df, base_name=export_name, key_prefix=export_key)
    return df

def brand_metric(
    label: str,
    value: str | int | float | None,
    delta: str | int | float | None = None,
    delta_good: bool | None = None,
):
    theme = st.session_state.get('active_theme', THEMES['CTPM Light Premium'])

    def fmt_val(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "-"
        if isinstance(v, (int, float)):
            try:
                return f"{int(v):,}" if float(v).is_integer() else f"{float(v):,.1f}"
            except Exception:
                return str(v)
        return str(v)

    chip_bg = "#FFFFFF"
    chip_fg = "#000000"
    if delta is not None:
        if delta_good is True:
            chip_bg, chip_fg = theme['ACCENT'], "#000000"
        elif delta_good is False:
            chip_bg, chip_fg = "#2B2B2B", "#FFFFFF"

    delta_html = ""
    if delta is not None:
        arrow = "▲" if (delta_good is True) else ("▼" if (delta_good is False) else "•")
        delta_html = (
            f"<div style='display:inline-block;padding:2px 8px;border-radius:999px;"
            f"background:{chip_bg};color:{chip_fg};font-weight:800;margin-left:8px;"
            f"font-size:0.85rem;line-height:1.2;'>{arrow} {fmt_val(delta)}</div>"
        )

    card_html = (
        f"<div class='kpi'>"
        f"<div class='label'>{label}</div>"
        f"<div class='value' style='display:flex;align-items:center;'>"
        f"{fmt_val(value)}{delta_html}</div></div>"
    )
    st.markdown(card_html, unsafe_allow_html=True)

# ======= Auth (minimal) =======
from datetime import datetime as _dt, timezone as _tz
AUTH_DB_PATH = str(CACHE_DIR / "users.db")
@st.cache_resource(show_spinner=False)
def _auth_conn(path: str):
    c = sqlite3.connect(path, check_same_thread=False)
    c.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE COLLATE NOCASE,
        role TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        password_salt TEXT NOT NULL,
        password_iters INTEGER NOT NULL,
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        last_login_at TEXT
    );
    """)
    c.commit(); return c

def _hash_pw(pw: str, salt: bytes|None=None, iters: int=200_000) -> dict:
    if salt is None: salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, iters)
    import base64 as _b
    return {"algo":"pbkdf2_sha256","iters":iters,"salt":_b.b64encode(salt).decode(),"hash":_b.b64encode(dk).decode()}

def _verify_pw(pw: str, rec: dict) -> bool:
    import base64 as _b
    salt = _b.b64decode(rec["salt"]); iters = int(rec["iters"])
    dk = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, iters)
    return _b.b64encode(dk).decode() == rec["hash"]


def _seed_default_admin():
    """Create a default admin account if CTPM_ADMIN_PASS is set and no users exist."""
    _pw = os.environ.get("CTPM_ADMIN_PASS", "")
    if not _pw:
        return
    _user = os.environ.get("CTPM_ADMIN_USER", "admin")
    conn = _auth_conn(AUTH_DB_PATH)
    if conn.execute("SELECT 1 FROM users LIMIT 1").fetchone():
        return  # users already exist
    _rec = _hash_pw(_pw)
    _now = _dt.now(_tz.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    conn.execute(
        "INSERT INTO users(username,role,password_hash,password_salt,password_iters,active,created_at,updated_at)"
        " VALUES (?,?,?,?,?,?,?,?);",
        (_user, "admin", _rec["hash"], _rec["salt"], _rec["iters"], 1, _now, _now),
    )
    conn.commit()

_seed_default_admin()

def _get_user(conn, username: str):
    cur = conn.execute("SELECT username, role, password_hash, password_salt, password_iters, active FROM users WHERE username=?;", (username,))
    r = cur.fetchone()
    if not r: return None
    u, role, ph, ps, it, active = r
    return {"username":u, "role":role, "password_hash":ph, "password_salt":ps, "password_iters":it, "active":bool(active)}

def _record_login(conn, username: str):
    now = _dt.now(_tz.utc).isoformat(timespec="seconds").replace('+00:00','Z')
    conn.execute("UPDATE users SET last_login_at=?, updated_at=? WHERE username=?;", (now, now, username)); conn.commit()

def _bootstrap_admin_if_needed(conn) -> bool:
    """Create admin/admin123! once, do nothing if it already exists.
       Returns True if created, False otherwise."""
    try:
        cur = conn.execute("SELECT 1 FROM users WHERE username=? LIMIT 1;", ('admin',))
        if cur.fetchone():
            return False  # already present

        rec = _hash_pw('admin123!')
        now = _dt.now(_tz.utc).isoformat(timespec='seconds').replace('+00:00','Z')
        conn.execute(
            "INSERT INTO users(username,role,password_hash,password_salt,password_iters,active,created_at,updated_at) "
            "VALUES (?,?,?,?,?,?,?,?);",
            ('admin', 'admin', rec['hash'], rec['salt'], rec['iters'], 1, now, now)
        )
        conn.commit()
        return True
    except Exception:
        # Final safety: if a concurrent pass inserts first, ignore the uniqueness error
        try:
            conn.rollback()
        except Exception:
            pass
        return False


def _login_ui(conn, source: str = "unknown"):
    # 🔒 Guard: if already authenticated, never draw the form again
    if st.session_state.get("auth", {}).get("is_authenticated", False):
        return

    st.subheader(f"Sign in  •  source: {source}")

    # Use a form so the page doesn't rerun between keystrokes
    with st.form("login_form", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            u = st.text_input("Username", key="auth_username")
        with c2:
            p = st.text_input("Password", type="password", key="auth_password")
        submitted = st.form_submit_button("Login", type="primary")

    if submitted:
        u = u.strip()
        rec = _get_user(conn, u)
        if rec and rec["active"] and _verify_pw(
            p, {"salt": rec["password_salt"], "iters": rec["password_iters"], "hash": rec["password_hash"]}
        ):
            st.session_state["auth"] = {
                "is_authenticated": True,
                "username": rec["username"],
                "role": rec["role"],
            }
            _record_login(conn, u)
            st.success(f"Welcome, {u}!")
            st.rerun()      # single rerun happens here, right after success
        elif rec and not rec["active"]:
            st.error("This account is deactivated.")
        else:
            st.error("Invalid username or password.")


def require_auth(allowed_roles: List[str] | None = None):
    auth = st.session_state.get("auth", {"is_authenticated": False})
    if not auth.get("is_authenticated"):
        conn = _auth_conn(AUTH_DB_PATH)

        if _bootstrap_admin_if_needed(conn):
            st.info("Bootstrap admin 'admin' created. Log in to continue.")

        _login_ui(conn, source="require-auth")
        st.stop()

    if allowed_roles and auth.get("role") not in allowed_roles:
        st.warning("You do not have permission to view this page.")
        st.stop()

# ================= SETTINGS =================
def load_settings() -> Dict:
    if SETTINGS_FILE.exists():
        try:
            s = json.loads(SETTINGS_FILE.read_text())
            for k,v in DEFAULT_SETTINGS.items():
                if k not in s: s[k]=v
            return s
        except Exception: pass
    SETTINGS_FILE.write_text(json.dumps(DEFAULT_SETTINGS, indent=2))
    return DEFAULT_SETTINGS.copy()

# ======= Data I/O =======
WOS_KEEP = ["Work Order","Company","Open Date","Due Date","Completed Date","Status","Sub-Status","Last Modified","P.O. Number","Description"]
EVENTS_KEEP = ["Event Type","Status","Tracking Status","Company","I.D.","Work Order","Entered By","Description","Result","Date","Time","Event Date (Universal)","Event Time (Universal)","Total Time"]
EQUIP_KEEP = ["Company","I.D.","Description","Manufacturer","Model Number","Last Calibration","Calibration Due","Service Site","In Shop","IN SHOP","Status","Tracking Status","Active","Current Event Date","Estimated Cal. Time"]


@st.cache_data(ttl=1200, show_spinner=False)
def load_clean_data(path: str, sig) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    h = _sig_hash(sig)
    pq_wos   = PARQUET_DIR / f"wos_{h}.parquet"
    pq_events = PARQUET_DIR / f"events_{h}.parquet"
    pq_equip = PARQUET_DIR / f"equip_{h}.parquet"

    # Fast path: Parquet cache exists for this version of the data
    if pq_wos.exists() and pq_events.exists() and pq_equip.exists():
        return pd.read_parquet(pq_wos), pd.read_parquet(pq_events), pd.read_parquet(pq_equip)

    # Check for individually-uploaded sheet files
    sheets = _upl_sheet_paths()
    using_individual = sheets["events"] is not None and sheets["equipment"] is not None

    if using_individual:
        wos   = _read_sheet_file(sheets["wos"],       WOS_KEEP)   if sheets["wos"]       else pd.DataFrame(columns=WOS_KEEP)
        events = _read_sheet_file(sheets["events"],   EVENTS_KEEP)
        equip  = _read_sheet_file(sheets["equipment"], EQUIP_KEEP)
    else:
        # Fall back to combined All Data workbook
        p = Path(sanitize_path(path))
        if not p.exists():
            raise FileNotFoundError(f"Data file not found: {p}")
        engine = "openpyxl" if p.suffix.lower() in (".xlsx",".xlsm") else "xlrd"
        xl = pd.ExcelFile(p, engine=engine)
        for need in ("All Events", "All Equipment"):
            if need not in xl.sheet_names:
                raise ValueError(f"Missing sheet: {need}")

        def _usecols(sheet, keep):
            hdr = pd.read_excel(p, sheet_name=sheet, nrows=0, engine=engine)
            cols = [c for c in keep if c in hdr.columns]
            return pd.read_excel(p, sheet_name=sheet, engine=engine) if not cols else pd.read_excel(
                p, sheet_name=sheet, usecols=cols, engine=engine
            )

        wos    = _usecols("All WOs", WOS_KEEP) if "All WOs" in xl.sheet_names else pd.DataFrame(columns=WOS_KEEP)
        events = _usecols("All Events", EVENTS_KEEP)
        equip  = _usecols("All Equipment", EQUIP_KEEP)

    for c in ("Open Date", "Due Date", "Completed Date", "Last Modified"):
        if c in wos.columns:
            wos[c] = pd.to_datetime(wos[c], errors="coerce")

    for c in ("Last Calibration", "Calibration Due", "Current Event Date"):
        if c in equip.columns:
            equip[c] = pd.to_datetime(equip[c], errors="coerce")

    if "IN SHOP" in equip.columns:
        s = equip["IN SHOP"].astype(str).str.strip().str.lower()
        equip["IN_SHOP"] = s.isin(["true","yes","1","y"])
    elif "In Shop" in equip.columns:
        s = equip["In Shop"].astype(str).str.strip().str.lower()
        equip["IN_SHOP"] = s.isin(["true","yes","1","y"])
    else:
        equip["IN_SHOP"] = False

    try:
        wos.to_parquet(pq_wos, index=False)
        events.to_parquet(pq_events, index=False)
        equip.to_parquet(pq_equip, index=False)
    except Exception:
        pass

    return wos, events, equip


@st.cache_data(ttl=1200, show_spinner=False)
def load_companies_df(path: str, sig) -> pd.DataFrame:
    pq = PARQUET_DIR / f"companies_{_sig_hash(sig)}.parquet"
    if pq.exists():
        return pd.read_parquet(pq)

    sheets = _upl_sheet_paths()
    if sheets["all_companies"]:
        df = _read_sheet_file(sheets["all_companies"], [])
    else:
        p = Path(sanitize_path(path))
        engine = "openpyxl" if p.suffix.lower() in (".xlsx",".xlsm") else "xlrd"
        try:
            df = pd.read_excel(p, sheet_name='All Companies', engine=engine)
        except Exception:
            return pd.DataFrame(columns=['Company','ZIP'])

    zip_col = next((c for c in df.columns if str(c).strip().lower() in ['zip','postal','postal code','zipcode']), None)
    if zip_col and zip_col != 'ZIP':
        df = df.rename(columns={zip_col:'ZIP'})
    df['ZIP'] = df.get('ZIP', pd.Series(index=df.index)).astype(str).str.extract(r'(\d{5})', expand=False)
    type_col = next((c for c in df.columns if str(c).strip().lower() == 'type'), None)
    if type_col and type_col != 'Type':
        df = df.rename(columns={type_col:'Type'})
    if 'Type' in df.columns:
        df['Type'] = df['Type'].astype(str).str.strip()
    try:
        df.to_parquet(pq, index=False)
    except Exception:
        pass
    return df


@st.cache_data(ttl=1200, show_spinner=False)
def load_wip_shop_df(path: str, sig) -> pd.DataFrame:
    pq = PARQUET_DIR / f"wip_shop_{_sig_hash(sig)}.parquet"
    if pq.exists():
        return pd.read_parquet(pq)

    sheets = _upl_sheet_paths()
    if sheets["wip_shop"]:
        df = _read_sheet_file(sheets["wip_shop"], [])
    else:
        p = Path(sanitize_path(path))
        engine = "openpyxl" if p.suffix.lower() in (".xlsx",".xlsm") else "xlrd"
        try:
            df = pd.read_excel(p, sheet_name='WIP Shop', engine=engine)
        except Exception:
            return pd.DataFrame(columns=['Company','I.D.','Work Order','Received On','Description','Tracking Status'])

    if 'Received On' in df.columns:
        df['Received On'] = pd.to_datetime(df['Received On'], errors='coerce')
    try:
        df.to_parquet(pq, index=False)
    except Exception:
        pass
    return df


# -----------------------------
# Efficiency datasets (WO export) - cached builder
# -----------------------------
@st.cache_data(ttl=1200, show_spinner=False)
def build_efficiency_datasets_from_file(
    data_file_path: str,
    sig: Tuple[float, int],
    all_wos_export_path: str,
    names_numbers_path: str,
    skip_threshold: float = 0.30,
) -> dict:
    """Build WO/Tech efficiency datasets using the external All WOs export.

    Caches on (data_file signature + external file paths) to avoid hashing large dataframes.
    """
    if build_efficiency_datasets is None:
        return {}
    _wos_df, _events_df, _equip_df = load_clean_data(data_file_path, sig)
    return build_efficiency_datasets(
        all_events_df=_events_df,
        all_wos_export_path=all_wos_export_path,
        names_numbers_path=names_numbers_path,
        skip_threshold=skip_threshold,
    )


def locate_efficiency_files(data_file_path: str) -> Tuple[Path, Path]:
    """Find All WOs export + Names_Numbers mapping next to All Data.xlsx."""
    data_folder = Path(sanitize_path(data_file_path)).resolve().parent
    all_wos_path = data_folder / 'All WOs.xlsx'
    if not all_wos_path.exists():
        candidates = sorted(
            list(data_folder.glob('All WOs*.xlsx'))
            + list(data_folder.glob('All WOs*.XLSX'))
            + list(data_folder.glob('All WOs*.xlsm'))
        )
        if candidates:
            all_wos_path = candidates[-1]
    names_path = data_folder / 'Names_Numbers.xlsx'
    return all_wos_path, names_path

# ================= Geocoding & Distances =================
@st.cache_resource(show_spinner=False)
def pgeo(): return pgeocode.Nominatim('US')

def zip_to_latlon(zip_code: str) -> Optional[Tuple[float,float]]:
    try:
        rec = pgeo().query_postal_code(str(zip_code)[:5])
        if pd.isna(rec.latitude) or pd.isna(rec.longitude): return None
        return float(rec.latitude), float(rec.longitude)
    except Exception:
        return None

def haversine_miles(lat1, lon1, lat2, lon2):
    R = 3958.8
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1); dl = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dl/2)**2
    return 2*R*math.asin(math.sqrt(a))

@st.cache_data(ttl=900)
def compute_company_distance(companies: pd.DataFrame, lab_zip: str) -> pd.DataFrame:
    comp = companies.copy()
    comp['ZIP'] = comp.get('ZIP').astype(str).str.extract(r'(\d{5})', expand=False)
    comp['ZIP3'] = comp['ZIP'].str[:3]
    comp['latlon'] = comp['ZIP'].map(lambda z: zip_to_latlon(z) if pd.notna(z) else None)
    comp['lat'] = comp['latlon'].map(lambda t: None if t is None else t[0])
    comp['lon'] = comp['latlon'].map(lambda t: None if t is None else t[1])
    lab = zip_to_latlon(str(lab_zip)) or (39.7069, -84.0560)
    lab_lat, lab_lon = lab
    comp['MilesFromLab'] = comp.apply(lambda r: haversine_miles(lab_lat, lab_lon, float(r['lat']), float(r['lon'])) if pd.notna(r['lat']) and pd.notna(r['lon']) else np.nan, axis=1)
    return comp[['Company','ZIP','ZIP3','lat','lon','MilesFromLab']]



def companies_shipping_only(companies: pd.DataFrame) -> pd.DataFrame:
    """Filter All Companies to Shipping rows only. Shipping-only even if miles becomes blank."""
    if companies is None or companies.empty:
        return pd.DataFrame(columns=['Company','ZIP','Type'])
    c = companies.copy()
    if 'Type' in c.columns:
        t = c['Type'].astype(str).str.strip().str.lower()
        c = c[t.eq('shipping')].copy()
    return c


def compute_company_urgency_from_wos(wos: pd.DataFrame, *, urgent_days: int = 3, soon_days: int = 10) -> pd.DataFrame:
    """Company-level Urgency from All WOs using Due Date + Priority."""
    if wos is None or wos.empty or 'Company' not in wos.columns:
        return pd.DataFrame(columns=['Company','Urgency','Priority_Best','WO_Earliest_Due'])
    w = wos.copy()
    w['Company'] = w['Company'].astype(str)
    if 'Due Date' in w.columns:
        w['Due Date'] = pd.to_datetime(w['Due Date'], errors='coerce')
    else:
        w['Due Date'] = pd.NaT

    pri = w.get('Priority', pd.Series(index=w.index, dtype='object')).astype(str).str.strip().str.lower()
    pri_map = {
        'rush':0,'expedite':0,'expedited':0,'urgent':0,'emergency':0,'asap':0,'critical':0,
        'high':1,'h':1,
        'normal':2,'medium':2,'m':2,
        'low':3,'l':3,
        '1':0,'2':1,'3':2,'4':3,'5':3
    }
    w['Priority_Score'] = pd.to_numeric(pri.map(pri_map), errors='coerce').fillna(2).astype(int)

    today = pd.Timestamp.today().normalize()
    days_to_due = (w['Due Date'].dt.normalize() - today).dt.days

    w['Due_Score'] = 2
    w.loc[days_to_due < 0, 'Due_Score'] = 0
    w.loc[(days_to_due >= 0) & (days_to_due <= urgent_days), 'Due_Score'] = 0
    w.loc[(days_to_due > urgent_days) & (days_to_due <= soon_days), 'Due_Score'] = 1

    w['Urgency'] = w[['Priority_Score','Due_Score']].min(axis=1).astype(int)

    out = w.groupby('Company', as_index=False).agg(
        Urgency=('Urgency','min'),
        Priority_Best=('Priority_Score','min'),
        WO_Earliest_Due=('Due Date','min')
    )
    return out


@st.cache_data(ttl=3600, show_spinner=False)
def zip_latlon_df(zips: list[str]) -> pd.DataFrame:
    """Return lat/lon for a list of 5-digit ZIPs using pgeocode."""
    rows = []
    for z in sorted({str(x).zfill(5) for x in (zips or []) if str(x).strip()}):
        coords = zip_to_latlon(z)
        if coords is None:
            continue
        rows.append({'ZIP': z, 'lat': coords[0], 'lon': coords[1]})
    return pd.DataFrame(rows)


def plotly_click_zip(fig, key: str = 'map_click') -> str | None:
    """Return clicked ZIP from a Plotly figure.

    Uses streamlit-plotly-events when available; otherwise falls back to no-click.
    """
    try:
        from streamlit_plotly_events import plotly_events
        selected = plotly_events(fig, click_event=True, select_event=False, hover_event=False, override_height=650, key=key)
        if selected and isinstance(selected, list) and len(selected) > 0:
            pt = selected[0]
            # We place ZIP in customdata[0]
            cd = pt.get('customdata')
            if isinstance(cd, (list, tuple)) and len(cd) > 0:
                return str(cd[0])
    except Exception:
        # fallback: render chart without click support
        st.plotly_chart(fig, width='stretch')
        st.caption("Tip: install click support with `pip install streamlit-plotly-events`.")
        return None
    return None

# Optimizer

def companies_in(record_company: str) -> List[str]:
    s = record_company.replace('ROUTE:','').replace('GROUPED JOB:','')
    return [p.strip() for p in s.split('|')] if '|' in s else [record_company]

@st.cache_data(ttl=900)
def optimize_schedule(demand: pd.DataFrame, policies: Dict[str,Dict], techs:int, hours_per_tech:float, vehicles:int, techs_per_vehicle:int,
 weekdays_only: bool = True, late_penalty_per_day: float = 5.0) -> pd.DataFrame:
 """Schedule on-site demand across days.

 Rules:
  - Prefer spreading jobs across multiple days using the smallest tech count possible.
  - Working-day thresholds (Mon-Fri when weekdays_only=True):
     * 1 tech if <= 4 working days.
     * 2 techs if 1 tech would take > 4 working days.
     * If 2 techs would take > 2 workweeks (10 working days), add 3rd/4th tech to bring duration <= 10 working days when possible.
     * Never suggest more than 4 techs.
  - Per-day allocation cap per job = suggested_techs * hours_per_tech (plus company daily cap).
  - Vehicle slots: each job-day consumes ceil(suggested_techs / techs_per_vehicle) vehicles.
  - Policies:
     * daily_hours_cap limits hours/day per company.
     * exclude=True removes a company from optimization. If a GROUPED JOB contains excluded members, they are removed and the rest remain.
 """
 if demand is None or demand.empty:
  return pd.DataFrame(columns=['Date','Company'])

 excluded = {str(c) for c,v in (policies or {}).items() if bool(v.get('exclude', False))}
 dmd = demand.copy()
 if 'Company' in dmd.columns and excluded:
  dmd = dmd[~dmd['Company'].astype(str).isin(excluded)].copy()
 if dmd.empty:
  return pd.DataFrame(columns=['Date','Company'])

 start = pd.Timestamp.today().normalize()
 end = pd.to_datetime(dmd['Window_End']).max().normalize()
 days = pd.date_range(start, end, freq='D')
 if weekdays_only:
  days = days[days.dayofweek < 5]

 daily_capacity = float(techs) * float(hours_per_tech)
 calendar = pd.DataFrame({'Date': days})
 calendar['Cap_Remaining'] = daily_capacity
 calendar['Visits_Remaining'] = int(vehicles)

 daily_caps: Dict[str,float] = {str(c): float(v['daily_hours_cap']) for c,v in (policies or {}).items() if v.get('daily_hours_cap') is not None}

 def suggested_techs_for_hours(total_hours: float) -> int:
  h = float(total_hours or 0)
  hpt = float(hours_per_tech or 8)
  if hpt <= 0:
   return 1
  max_days_one = 4
  max_days_two_weeks = 10 if weekdays_only else 14
  days_one = int(math.ceil(h / hpt))
  if days_one <= max_days_one:
   return 1
  days_two = int(math.ceil(h / (2*hpt)))
  if days_two <= max_days_two_weeks:
   return 2
  for t in (3,4):
   if int(math.ceil(h / (t*hpt))) <= max_days_two_weeks:
    return t
  return 4

 def vehicles_needed(tcount: int) -> int:
  tpv = max(1, int(techs_per_vehicle or 1))
  return int(math.ceil(int(tcount) / tpv))

 if 'Urgency' not in dmd.columns:
  dmd['Urgency'] = 2
 if 'MilesFromLab' not in dmd.columns:
  dmd['MilesFromLab'] = np.nan
 if 'Est_Hours' not in dmd.columns:
  dmd['Est_Hours'] = 0.0
 dmd = dmd.sort_values(['Urgency','MilesFromLab','Est_Hours'], ascending=[True, False, False]).reset_index(drop=True)

 assignments = []
 visits_by_company_month: Dict[tuple[str,str], int] = {}

 for _, r in dmd.iterrows():
  comp = str(r.get('Company',''))
  total_hours = float(r.get('Est_Hours', 0.0) or 0.0)
  if total_hours <= 0:
   continue

  ws = pd.Timestamp(r.get('Window_Start', start)).normalize()
  we = pd.Timestamp(r.get('Window_End', end)).normalize()
  target = pd.Timestamp(r.get('Due_Median', ws)).normalize() if pd.notna(r.get('Due_Median', pd.NaT)) else ws

  members = companies_in(comp)
  if excluded and members:
   members = [m for m in members if str(m) not in excluded]
  if excluded and (not members):
   continue

  t_suggest = suggested_techs_for_hours(total_hours)
  v_need = vehicles_needed(t_suggest)

  remaining = float(total_hours)
  cand = [d for d in days if (d >= ws and d <= we)]
  def score(d):
   late_days = max(0, (d - we).days)
   return (abs((d - target).days) + late_penalty_per_day * late_days, d)
  cand.sort(key=score)

  while remaining > 0:
   chosen = None
   for d in cand:
    row = calendar.loc[calendar['Date'] == d].iloc[0]
    if row['Cap_Remaining'] <= 0:
     continue
    if row['Visits_Remaining'] < v_need:
     continue
    month = d.strftime('%Y-%m')
    ok_visits = True
    for mcomp in members:
     if visits_by_company_month.get((mcomp, month), 0) >= 1 and not any((a['Date'] == d and mcomp in companies_in(a['Company'])) for a in assignments):
      ok_visits = False
      break
    if not ok_visits:
     continue
    chosen = d
    break

   if chosen is None:
    tail = [d for d in days if d > we]
    tail.sort(key=lambda d: ((d - we).days, d))
    for d in tail:
     row = calendar.loc[calendar['Date'] == d].iloc[0]
     if row['Cap_Remaining'] <= 0:
      continue
     if row['Visits_Remaining'] < v_need:
      continue
     month = d.strftime('%Y-%m')
     ok_visits = True
     for mcomp in members:
      if visits_by_company_month.get((mcomp, month), 0) >= 1 and not any((a['Date'] == d and mcomp in companies_in(a['Company'])) for a in assignments):
       ok_visits = False
       break
     if not ok_visits:
      continue
     chosen = d
     break

   if chosen is None:
    assignments.append({'Date':pd.NaT,'Company':comp,'Est_Hours':remaining,'MilesFromLab':r.get('MilesFromLab',np.nan),'Count_Items':r.get('Count_Items',0),
                        'Earliest_Due':r.get('Earliest_Due',pd.NaT),'IDs':r.get('IDs',''),'Members':r.get('RouteMembers',''),
                        'Techs_Suggested':int(t_suggest),'Vehicles_Needed':int(v_need)})
    break

   row = calendar.loc[calendar['Date'] == chosen].iloc[0]
   if members:
    cap_list = [daily_caps.get(str(m), float(row['Cap_Remaining'])) for m in members]
    company_cap = float(min(cap_list))
   else:
    company_cap = float(daily_caps.get(comp, float(row['Cap_Remaining'])))

   job_daily_cap = float(t_suggest) * float(hours_per_tech)
   site_cap = float(row['Cap_Remaining'])
   usable = min(site_cap, remaining, job_daily_cap, company_cap)
   if usable <= 0:
    calendar.loc[calendar['Date'] == chosen, 'Cap_Remaining'] = 0
    continue

   calendar.loc[calendar['Date'] == chosen, 'Cap_Remaining'] = site_cap - usable
   month = chosen.strftime('%Y-%m')
   new_visit = not any((a['Date'] == chosen and a.get('Company') == comp) for a in assignments)
   if new_visit:
    calendar.loc[calendar['Date'] == chosen, 'Visits_Remaining'] = float(row['Visits_Remaining']) - v_need
    for mcomp in members:
     visits_by_company_month[(mcomp, month)] = visits_by_company_month.get((mcomp, month), 0) + 1

   assignments.append({'Date':chosen,'Company':comp,'Est_Hours':round(float(usable),1),'MilesFromLab':r.get('MilesFromLab',np.nan),'Count_Items':r.get('Count_Items',0),
                       'Earliest_Due':r.get('Earliest_Due',pd.NaT),'IDs':r.get('IDs',''),'Members':r.get('RouteMembers',''),
                       'Techs_Suggested':int(t_suggest),'Vehicles_Needed':int(v_need)})
   remaining = round(float(remaining) - float(usable), 6)

 out = pd.DataFrame(assignments)
 if out.empty:
  return out

 grp = out.groupby(['Date','Company'], as_index=False).agg({
  'Est_Hours':'sum',
  'MilesFromLab':'max',
  'Count_Items':'sum',
  'Earliest_Due':'min',
  'IDs': lambda s: ', '.join(sorted(set(', '.join(map(str,s)).split(', ')))) ,
  'Techs_Suggested':'max',
  'Vehicles_Needed':'max',
 })
 return grp.sort_values(['Date','Company'])
# ICS export

def ics_from_schedule(df: pd.DataFrame, org: str = "CTPM Scheduling") -> bytes:
    def esc(s: str) -> str:
        return s.replace('\\', '\\\\').replace(';', '\\;').replace(',', '\\,').replace('\n', '\\n')
    lines = ["BEGIN:VCALENDAR","VERSION:2.0",f"PRODID:-//{org}//EN"]
    for _,r in df.iterrows():
        if pd.isna(r.get('Date')): continue
        uid = f"{re.sub('[^A-Za-z0-9]','',str(r.get('Company','')))[:40]}-{str(r.get('Date',''))[:10]}@ctpm"
        dtstamp = datetime.now(_tz.utc).strftime('%Y%m%dT%H%M%SZ')
        dtdue = pd.Timestamp(r.get('Date')).strftime('%Y%m%d')
        dtend = (pd.Timestamp(r.get('Date')) + pd.Timedelta(days=1)).strftime('%Y%m%d')
        summary = esc(f"On-Site: {str(r.get('Company',''))[:60]} ({r.get('Est_Hours',0):.1f} hrs, ~{int(r.get('Techs_Suggested',0))} techs)")
        desc = esc(f"Miles from lab: {r.get('MilesFromLab',''):.0f}\nEarliest due: {str(r.get('Earliest_Due',''))[:10]}\nIDs: {str(r.get('IDs',''))[:450]}")
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{dtstamp}",
            f"DTSTART;VALUE=DATE:{dtdue}",
            f"DTEND;VALUE=DATE:{dtend}",
            f"SUMMARY:{summary}",
            f"DESCRIPTION:{desc}",
            "END:VEVENT"
        ]
    lines.append("END:VCALENDAR")
    return ("\r\n".join(lines)).encode('utf-8')

# ================= ZIP CENTROIDS =================

def build_zip_centroids(companies: pd.DataFrame) -> pd.DataFrame:
    if companies is None or companies.empty or 'ZIP' not in companies.columns:
        return pd.DataFrame(columns=['zip','lat','lon'])
    zips = companies['ZIP'].dropna().astype(str).str.extract(r'(\d{5})', expand=False).dropna().unique()
    rows=[]
    for z in zips:
        coords = zip_to_latlon(z)
        if coords is not None:
            rows.append({'zip': str(z).zfill(5), 'lat': coords[0], 'lon': coords[1]})
    return pd.DataFrame(rows)

# ======= Map Bubble Diagnostics (ZIP normalization + centroid join) =======

ZIP5_RE = re.compile(r"(\d{5})")
EXCEL_DOT_ZERO_RE = re.compile(r"^\d+\.0$")

def normalize_zip(z):
 """Normalize a ZIP-like value to a 5-digit US ZIP string."""
 if z is None:
  return None
 try:
  if pd.isna(z):
   return None
 except Exception:
  pass
 s = str(z).strip()
 if not s or s.lower() in {"nan","none","null"}:
  return None
 if EXCEL_DOT_ZERO_RE.fullmatch(s):
  s = s[:-2]
 m = ZIP5_RE.search(s)
 return m.group(1) if m else None


def prepare_zip(centroids: pd.DataFrame, zip_col: str = 'ZIP', lat_col: str = 'lat', lon_col: str = 'lon') -> pd.DataFrame:
 """Prepare/normalize ZIP centroid table to columns: zip_norm, lat, lon."""
 c = centroids.copy()
 if zip_col not in c.columns:
  raise KeyError(f"Centroid ZIP column '{zip_col}' not found.")
 c['zip_norm'] = c[zip_col].apply(normalize_zip)
 c['lat'] = pd.to_numeric(c.get(lat_col), errors='coerce')
 c['lon'] = pd.to_numeric(c.get(lon_col), errors='coerce')
 c = c.dropna(subset=['zip_norm','lat','lon']).drop_duplicates(subset=['zip_norm'])
 return c[['zip_norm','lat','lon']]


def attach_zip(agg: pd.DataFrame, centroids_prepared: pd.DataFrame, agg_zip_col: str = 'ZIP') -> pd.DataFrame:
 """Left-join prepared centroids onto aggregated ZIP dataframe."""
 out = agg.copy()
 out['zip_norm'] = out[agg_zip_col].apply(normalize_zip)
 out = out.merge(centroids_prepared, how='left', on='zip_norm')
 return out


def coerce_lat_lon(df: pd.DataFrame, lat_col: str = 'lat', lon_col: str = 'lon'):
 """Coerce lat/lon to numeric, drop NaNs, and compute sanity metrics."""
 out = df.copy()
 out[lat_col] = pd.to_numeric(out.get(lat_col), errors='coerce')
 out[lon_col] = pd.to_numeric(out.get(lon_col), errors='coerce')
 out = out.dropna(subset=[lat_col, lon_col])
 metrics = {
  'lat_min': float(out[lat_col].min()) if len(out) else None,
  'lat_max': float(out[lat_col].max()) if len(out) else None,
  'lon_min': float(out[lon_col].min()) if len(out) else None,
  'lon_max': float(out[lon_col].max()) if len(out) else None,
 }
 if metrics['lat_min'] is not None:
  metrics['lat_in_us_range'] = 24 <= metrics['lat_min'] <= 50 and 24 <= metrics['lat_max'] <= 50
  metrics['lon_in_us_range'] = -125 <= metrics['lon_min'] <= -66 and -125 <= metrics['lon_max'] <= -66
  metrics['maybe_swapped'] = (metrics['lat_in_us_range'] is False and metrics['lon_in_us_range'] is False)
 else:
  metrics['lat_in_us_range'] = None
  metrics['lon_in_us_range'] = None
  metrics['maybe_swapped'] = None
 return out, metrics

# Backward-compatible aliases
prepare_zip_centroids = prepare_zip
attach_zip_centroids = attach_zip


# ======= Grouping helpers (distance-based) =======
def cluster_points_greedy(df: pd.DataFrame, radius_miles: float) -> pd.DataFrame:
 """Greedy clustering for a small number of points (ZIP centroids)."""
 if df is None or df.empty or float(radius_miles or 0) <= 0:
  out = df.copy() if df is not None else pd.DataFrame()
  if not out.empty and 'ZIP' in out.columns:
   out['ClusterID'] = out['ZIP'].astype(str)
   out['ZIPs'] = out['ZIP'].astype(str)
  else:
   out['ClusterID'] = [str(i) for i in range(len(out))]
   out['ZIPs'] = ''
  out['Members'] = 1
  return out

 pts = df.reset_index(drop=True).copy()
 pts['lat'] = pd.to_numeric(pts.get('lat'), errors='coerce')
 pts['lon'] = pd.to_numeric(pts.get('lon'), errors='coerce')
 pts = pts.dropna(subset=['lat','lon']).reset_index(drop=True)
 if pts.empty:
  return pts

 used = [False]*len(pts)
 cid = 0
 cids = [-1]*len(pts)
 for i in range(len(pts)):
  if used[i]:
   continue
  used[i] = True
  cids[i] = cid
  base_lat, base_lon = float(pts.loc[i,'lat']), float(pts.loc[i,'lon'])
  for j in range(i+1, len(pts)):
   if used[j]:
    continue
   dmi = haversine_miles(base_lat, base_lon, float(pts.loc[j,'lat']), float(pts.loc[j,'lon']))
   if dmi <= float(radius_miles):
    used[j] = True
    cids[j] = cid
  cid += 1

 pts['ClusterID'] = cids

 def wavg(g, col):
  w = pd.to_numeric(g.get('Count', 1), errors='coerce').fillna(1).astype(float).clip(lower=1)
  return float((pd.to_numeric(g[col], errors='coerce')*w).sum()/w.sum())

 out = pts.groupby('ClusterID', as_index=False).apply(
  lambda g: pd.Series({
   'Count': int(pd.to_numeric(g.get('Count', 1), errors='coerce').fillna(1).sum()),
   'lat': wavg(g,'lat'),
   'lon': wavg(g,'lon'),
   'ZIPs': ', '.join(g.get('ZIP', pd.Series([], dtype=str)).astype(str).tolist()[:250]),
   'Companies': ', '.join(sorted(set(', '.join(g.get('Companies', pd.Series([], dtype=str)).astype(str)).split(', '))))[:600] if 'Companies' in g.columns else '',
   'Members': int(len(g)),
  })
 ).reset_index(drop=True)
 out['ClusterID'] = out['ClusterID'].astype(str)
 return out


def group_companies_by_distance(demand: pd.DataFrame, company_locs: pd.DataFrame, radius_miles: float, small_hours_threshold: float) -> pd.DataFrame:
 """Group small-demand companies within radius_miles into combined rows."""
 if demand is None or demand.empty or float(radius_miles or 0) <= 0:
  return demand
 if 'Company' not in demand.columns or 'Est_Hours' not in demand.columns:
  return demand

 d = demand.copy()
 loc = company_locs.copy() if company_locs is not None else pd.DataFrame(columns=['Company','lat','lon'])
 if not loc.empty:
  loc['Company'] = loc['Company'].astype(str)
  loc['lat'] = pd.to_numeric(loc.get('lat'), errors='coerce')
  loc['lon'] = pd.to_numeric(loc.get('lon'), errors='coerce')

 d['Company'] = d['Company'].astype(str)
 d = d.merge(loc[['Company','lat','lon']], on='Company', how='left')

 small = d[pd.to_numeric(d['Est_Hours'], errors='coerce').fillna(0) < float(small_hours_threshold or 0)].copy()
 big = d.drop(index=small.index).copy()
 small = small.dropna(subset=['lat','lon']).reset_index(drop=True)
 if small.empty:
  return demand

 used = [False]*len(small)
 rows = []
 for i in range(len(small)):
  if used[i]:
   continue
  used[i] = True
  base_lat, base_lon = float(small.loc[i,'lat']), float(small.loc[i,'lon'])
  members = [i]
  for j in range(i+1, len(small)):
   if used[j]:
    continue
   dmi = haversine_miles(base_lat, base_lon, float(small.loc[j,'lat']), float(small.loc[j,'lon']))
   if dmi <= float(radius_miles):
    used[j] = True
    members.append(j)

  g = small.loc[members].copy()
  if len(g) == 1:
   rows.append(g.drop(columns=['lat','lon'], errors='ignore'))
   continue

  comps = g['Company'].astype(str).tolist()
  row = g.iloc[0].copy()
  row['Company'] = 'GROUPED JOB: ' + ' | '.join(comps[:8]) + (' ...' if len(comps) > 8 else '')
  row['RouteMembers'] = ' | '.join([str(x) for x in comps])
  row['Est_Hours'] = float(pd.to_numeric(g['Est_Hours'], errors='coerce').fillna(0).sum())
  if 'Count_Items' in g.columns:
   row['Count_Items'] = int(pd.to_numeric(g['Count_Items'], errors='coerce').fillna(0).sum())
  if 'MilesFromLab' in g.columns:
   row['MilesFromLab'] = float(pd.to_numeric(g['MilesFromLab'], errors='coerce').fillna(0).max())
  if 'Earliest_Due' in g.columns:
   row['Earliest_Due'] = pd.to_datetime(g['Earliest_Due'], errors='coerce').min()
  if 'Due_Median' in g.columns:
   row['Due_Median'] = pd.to_datetime(g['Due_Median'], errors='coerce').median()
  if 'Window_Start' in g.columns:
   row['Window_Start'] = pd.to_datetime(g['Window_Start'], errors='coerce').min()
  if 'Window_End' in g.columns:
   row['Window_End'] = pd.to_datetime(g['Window_End'], errors='coerce').max()
  if 'IDs' in g.columns:
   ids = []
   for s in g['IDs'].astype(str).tolist():
    ids.extend([x.strip() for x in s.split(',') if x.strip()])
   row['IDs'] = ', '.join(sorted(set(ids)))
  rows.append(pd.DataFrame([row.drop(labels=['lat','lon'], errors='ignore')]))

 grouped = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=demand.columns)
 out = pd.concat([big.drop(columns=['lat','lon'], errors='ignore'), grouped], ignore_index=True)
 return out

# ================= Policy DB (load/save) =================
import sqlite3 as _sqlite3

def _ensure_policy_db(db_path: Path = POLICY_DB):
 try:
  con = _sqlite3.connect(str(db_path))
  con.execute("CREATE TABLE IF NOT EXISTS policies (company TEXT PRIMARY KEY COLLATE NOCASE, cadence TEXT, daily_hours_cap REAL, exclude_from_opt INTEGER NOT NULL DEFAULT 0);")
  # Migration: add missing column
  try:
   cols = [r[1] for r in con.execute('PRAGMA table_info(policies);').fetchall()]
   if 'exclude_from_opt' not in cols:
    con.execute('ALTER TABLE policies ADD COLUMN exclude_from_opt INTEGER NOT NULL DEFAULT 0;')
  except Exception:
   pass
  con.commit()
 finally:
  con.close()

@st.cache_data(ttl=600, show_spinner=False)
def pol_load_all() -> Dict[str, Dict]:
    _ensure_policy_db(POLICY_DB)
    con = _sqlite3.connect(str(POLICY_DB))
    try:
        rows = con.execute("SELECT company, cadence, daily_hours_cap, exclude_from_opt FROM policies").fetchall()
        return {str(c): {'cadence': (cad or 'default'), 'daily_hours_cap': (None if cap is None else float(cap)), 'exclude': bool(ex or 0)} for c,cad,cap,ex in rows}
    finally:
        con.close()

def pol_save_many(policies: Dict[str, Dict]) -> None:
    if not policies: return
    _ensure_policy_db(POLICY_DB)
    con = _sqlite3.connect(str(POLICY_DB))
    try:
        payload = [(str(comp), str(rec.get('cadence','default')), rec.get('daily_hours_cap'), (1 if rec.get('exclude', False) else 0)) for comp,rec in policies.items()]
        con.executemany("INSERT INTO policies(company,cadence,daily_hours_cap,exclude_from_opt) VALUES (?,?,?,?) ON CONFLICT(company) DO UPDATE SET cadence=excluded.cadence, daily_hours_cap=excluded.daily_hours_cap, exclude_from_opt=excluded.exclude_from_opt", payload)
        con.commit(); pol_load_all.clear()
    finally:
        con.close()

@st.cache_data(ttl=900, show_spinner=False)
def events_enriched_from_file(path: str, sig: Tuple[float,int]) -> pd.DataFrame:
    _, events, _ = load_clean_data(path, sig)
    if events is None or events.empty:
        return pd.DataFrame(columns=['Work Order','I.D.','Company','Tracking Status','__etype_norm__','event_ts'])
    e = events.copy()
    for c in ['Work Order','I.D.','Company','Tracking Status','Event Type','Status','Result','Description']:
        if c not in e.columns: e[c] = ''
    d = pd.to_datetime(e.get('Event Date (Universal)', e.get('Date')), errors='coerce')
    if 'Event Time (Universal)' in e.columns:
        hhmm = pd.to_timedelta(e['Event Time (Universal)'].astype(str), errors='coerce')
        e['event_ts'] = (d + hhmm)
    else:
        e['event_ts'] = d
    _CAL_TYPES = {
        'shop calibration', 'pipette cal in-shop', 'pippette cal in-shop',
        'cover letter cal cert',
    }
    def _etype_norm(row):
        et = str(row.get('Event Type', '') or '').strip().lower()
        if et == 'receiving in-shop':
            return 'receiving'
        if et in _CAL_TYPES:
            return 'shop calibration'
        if et == 'qc':
            return 'qc'
        if et == 'shipping':
            return 'shipping'
        # Narrow fallback: use only Tracking Status (structured field), never Description/Result
        ts = str(row.get('Tracking Status', '') or '').strip().lower()
        if 'receiv' in ts:
            return 'receiving'
        if ts == 'qc':
            return 'qc'
        if any(k in ts for k in ('ship', 'return to customer', 'courier delivered', 'picked up')):
            return 'shipping'
        if 'cal' in ts:
            return 'shop calibration'
        if 'in shop' in ts or 'in-shop' in ts:
            return 'in shop'
        if 'wip' in ts:
            return 'wip'
        return 'event'
    e['__etype_norm__'] = e.apply(_etype_norm, axis=1)
    keep = ['Work Order','I.D.','Company','Tracking Status','__etype_norm__','event_ts']
    return e[keep].dropna(subset=['Work Order'])

@st.cache_data(ttl=900, show_spinner=False)
def wip_chain(path: str, sig: Tuple[float,int], equip: pd.DataFrame) -> pd.DataFrame:
    e = events_enriched_from_file(path, sig)
    if e.empty:
        return pd.DataFrame(columns=['Company','Work Order','I.D.','Recv','Days_Since_Received','Stage_Status',
                                     'Days_R2Cal','Days_Cal2QC','Days_QC2Ship','Days_Total_R2Ship'])
    rec = (e[e['__etype_norm__']=='receiving'].dropna(subset=['event_ts']).groupby(['Work Order','I.D.'], as_index=False)['event_ts'].min().rename(columns={'event_ts':'Recv'}))
    latest = (e.dropna(subset=['event_ts']).sort_values('event_ts').groupby(['Work Order','I.D.'], as_index=False).last())
    out = rec.merge(latest[['Work Order','I.D.','Company','Tracking Status','__etype_norm__']], on=['Work Order','I.D.'], how='left')
    today = pd.Timestamp.today().normalize()
    out['Days_Since_Received'] = (today - pd.to_datetime(out['Recv']).dt.normalize()).dt.days
    out['Stage_Status'] = out['__etype_norm__'].fillna('wip').replace({'receiving':'Receiving','qc':'QC','shipping':'Shipping/Returned','shop calibration':'Shop Calibration','in shop':'In Shop','wip':'WIP'})

    # Stage-to-stage timestamps for TAT analysis
    ev = e.dropna(subset=['event_ts']).copy()
    first_cal = (ev[ev['__etype_norm__']=='shop calibration']
                 .groupby(['Work Order','I.D.'], as_index=False)['event_ts'].min()
                 .rename(columns={'event_ts':'Cal_ts'}))
    first_qc = (ev[ev['__etype_norm__']=='qc']
                .groupby(['Work Order','I.D.'], as_index=False)['event_ts'].min()
                .rename(columns={'event_ts':'QC_ts'}))
    last_ship = (ev[ev['__etype_norm__']=='shipping']
                 .groupby(['Work Order','I.D.'], as_index=False)['event_ts'].max()
                 .rename(columns={'event_ts':'Ship_ts'}))
    out = out.merge(first_cal, on=['Work Order','I.D.'], how='left')
    out = out.merge(first_qc, on=['Work Order','I.D.'], how='left')
    out = out.merge(last_ship, on=['Work Order','I.D.'], how='left')
    out['Days_R2Cal'] = (out['Cal_ts'] - out['Recv']).dt.total_seconds() / 86400
    out['Days_Cal2QC'] = (out['QC_ts'] - out['Cal_ts']).dt.total_seconds() / 86400
    out['Days_QC2Ship'] = (out['Ship_ts'] - out['QC_ts']).dt.total_seconds() / 86400
    out['Days_Total_R2Ship'] = (out['Ship_ts'] - out['Recv']).dt.total_seconds() / 86400

    # TAT sanity: negative durations imply ship-before-receive or timestamp issues
    dur_cols = ['Days_R2Cal','Days_Cal2QC','Days_QC2Ship','Days_Total_R2Ship']
    out['TAT_Anomaly'] = False
    for c in dur_cols:
        out[c] = pd.to_numeric(out.get(c), errors='coerce')
        neg = out[c].notna() & (out[c] < 0)
        out.loc[neg, 'TAT_Anomaly'] = True
        out.loc[neg, c] = np.nan

    # In_Shop = True when the item has NOT yet reached the shipping stage
    out['In_Shop'] = ~out['__etype_norm__'].eq('shipping')
    return out[['Company','Work Order','I.D.','Recv','Days_Since_Received','Stage_Status',
                'Days_R2Cal','Days_Cal2QC','Days_QC2Ship','Days_Total_R2Ship','TAT_Anomaly','In_Shop']]

@st.cache_data(ttl=600, show_spinner=False)
def build_demand(equip: pd.DataFrame, wos: pd.DataFrame, companies: pd.DataFrame, lab_zip: str, near_radius:int, far_radius:int, near_days:int, mid_days:int, far_months:int, est_per_item: float, horizon_months:int, policies: Dict[str,Dict], cluster_enable: bool, cluster_zip3: bool, cluster_miles: float, small_hours_threshold: float, assumed_mph: float, per_stop_travel_hours: float) -> pd.DataFrame:
    if equip is None or equip.empty:
        return pd.DataFrame(columns=['Company','Kind','MilesFromLab','Window_Start','Window_End','Due_Median','Earliest_Due','Count_Items','Est_Hours','IDs','RouteMembers'])
    df = equip.copy()
    # ACTIVE-only: items not active do not need scheduling
    # Keep blanks as active so we don't miss anything (conservative).
    if 'Active' in df.columns:
        def _is_active(v):
            try:
                if pd.isna(v):
                    return True
            except Exception:
                pass
            if isinstance(v, (bool, np.bool_)):
                return bool(v)
            s = str(v).strip().lower()
            if s in ('false','0','no','n','inactive','in-active'):
                return False
            return True
        df = df[df['Active'].map(_is_active)].copy()

    # Exclude equipment already in the lab (Tracking Status == 'WIP Shop')
    if 'Tracking Status' in df.columns:
        _ts = df['Tracking Status'].astype(str).str.strip().str.lower()
        df = df[~_ts.eq('wip shop')].copy()
    # Ignore items more than N days past due (default 90)
    ignore_days = int(st.session_state.get('overdue_ignore_days', 90))
    if 'Calibration Due' in df.columns:
        _due = pd.to_datetime(df['Calibration Due'], errors='coerce')
        _cutoff = pd.Timestamp.today().normalize() - pd.Timedelta(days=ignore_days)
        df = df[(_due.isna()) | (_due >= _cutoff)].copy()
    if 'Calibration Due' in df.columns:
        df['Calibration Due'] = pd.to_datetime(df['Calibration Due'], errors='coerce')
    today = pd.Timestamp.today().normalize()
    horizon_end = today + pd.DateOffset(months=int(horizon_months or 6))
    if 'IN_SHOP' in df.columns:
        df = df[df['IN_SHOP']==False]
    if 'Calibration Due' in df.columns:
        df = df[(df['Calibration Due'].notna()) & (df['Calibration Due']<=horizon_end)]
    # company ZIP map
    zc = {}
    if companies is not None and not companies.empty:
        zcol = next((c for c in companies.columns if str(c).lower() in ['zip','postal','postal code','zipcode']), None)
        if zcol is not None:
            for _,r in companies[['Company',zcol]].dropna().iterrows(): zc[str(r['Company'])] = str(r[zcol])[:5]
    df['Company'] = df.get('Company', pd.Series(index=df.index)).astype(str)
    df['ZIP'] = df['Company'].map(lambda c: zc.get(str(c), None))
    # For safety, skip distance calc in this minimal helper
    # MilesFromLab based on Shipping address ZIPs only (All Companies Type='Shipping')
    try:
        ship_companies = companies_shipping_only(companies)
        dist = compute_company_distance(ship_companies, str(lab_zip).zfill(5)) if ship_companies is not None and not ship_companies.empty else pd.DataFrame()
        if dist is not None and not dist.empty:
            df = df.merge(dist[['Company','MilesFromLab']], on='Company', how='left')
        else:
            df['MilesFromLab'] = np.nan
    except Exception:
        df['MilesFromLab'] = np.nan
    df['Window_Start'] = today
    df['I.D.'] = df.get('I.D.', pd.Series(index=df.index)).astype(str)
    _m = pd.to_numeric(df.get('MilesFromLab', pd.Series(np.nan, index=df.index)), errors='coerce')
    df['_Kind'] = np.where(_m.isna(), 'near',
                  np.where(_m <= float(near_radius), 'near',
                  np.where(_m <= float(far_radius), 'mid', 'far')))
    _near_end = today + pd.Timedelta(days=int(near_days))
    _mid_end  = today + pd.Timedelta(days=int(mid_days))
    _far_end  = today + pd.DateOffset(months=int(far_months))
    df['Window_End'] = df['_Kind'].map({'near': _near_end, 'mid': _mid_end, 'far': _far_end})
    agg = df.groupby('Company', as_index=False).agg(
        Count_Items=('I.D.','nunique'),
        Est_Hours=('I.D.', lambda s: float(len(set(s))) * float(est_per_item or 0.5)),
        MilesFromLab=('MilesFromLab','max'),
        Earliest_Due=('Calibration Due','min'),
        Due_Median=('Calibration Due','median'),
        Window_Start=('Window_Start','min'),
        Window_End=('Window_End','max'),
        IDs=('I.D.', lambda s: ', '.join(sorted({str(x) for x in s if str(x).strip()}))),
        Kind=('_Kind', 'first'),
    )
    agg['RouteMembers'] = ''
    # Company-level Urgency from All WOs (adjustable thresholds)
    urgent_days = int(st.session_state.get('urgency_urgent_days', 3))
    soon_days = int(st.session_state.get('urgency_soon_days', 10))
    ucomp = compute_company_urgency_from_wos(wos, urgent_days=urgent_days, soon_days=soon_days)
    if ucomp is not None and not ucomp.empty:
        agg = agg.merge(ucomp, on='Company', how='left')
    agg['Urgency'] = pd.to_numeric(agg.get('Urgency', 2), errors='coerce').fillna(2).astype(int)
    return agg[['Company','Kind','MilesFromLab','Window_Start','Window_End','Due_Median','Earliest_Due','Count_Items','Est_Hours','IDs','RouteMembers','Urgency']]

@st.cache_data(ttl=600, show_spinner=False)
def capacity_planning_chart(
    demand: pd.DataFrame,
    weeks: int,
    techs: int,
    hours_per_tech: float,
    weekdays_only: bool = True,
    schedule_df: pd.DataFrame | None = None,
):
    """Upcoming weeks workload vs capacity.

    If schedule_df is provided, workload is computed from the Optimized Schedule and the chart covers
    the same horizon as the schedule.

    Otherwise, it falls back to demand-based workload anchored to Window_End.
    """
    today = pd.Timestamp.today().normalize()
    workdays = 5 if weekdays_only else 7
    weekly_capacity = float(techs) * float(hours_per_tech) * float(workdays)

    end_date = None
    if schedule_df is not None and isinstance(schedule_df, pd.DataFrame) and (not schedule_df.empty) and 'Date' in schedule_df.columns:
        end_date = pd.to_datetime(schedule_df['Date'], errors='coerce').dropna().max()

    if end_date is None:
        if demand is None or demand.empty:
            end_date = today + pd.Timedelta(weeks=int(weeks or 8))
        else:
            end_date = pd.to_datetime(demand.get('Window_End', pd.Series([today])), errors='coerce').dropna().max()
            if end_date is None or pd.isna(end_date):
                end_date = today + pd.Timedelta(weeks=int(weeks or 8))

    end_date = pd.Timestamp(end_date).normalize()

    wk0 = today.to_period('W-MON').start_time
    wk_last = end_date.to_period('W-MON').start_time
    week_starts = pd.date_range(wk0, wk_last, freq='W-MON')
    base = pd.DataFrame({'WeekStart': week_starts})

    if schedule_df is not None and isinstance(schedule_df, pd.DataFrame) and (not schedule_df.empty) and {'Date','Est_Hours'}.issubset(schedule_df.columns):
        s = schedule_df.copy()
        s['Date'] = pd.to_datetime(s['Date'], errors='coerce')
        s = s.dropna(subset=['Date']).copy()
        s['WeekStart'] = s['Date'].dt.to_period('W-MON').dt.start_time
        req = s.groupby('WeekStart', as_index=False)['Est_Hours'].sum().rename(columns={'Est_Hours':'Hours'})
        suffix = ' (based on Optimized Schedule)'
    else:
        suffix = ''
        if demand is None or demand.empty:
            req = pd.DataFrame({'WeekStart': week_starts, 'Hours': [0.0]*len(week_starts)})
        else:
            d = demand.copy()
            d['Anchor'] = pd.to_datetime(d.get('Window_End', pd.Series([today]*len(d))), errors='coerce').fillna(today)
            d['WeekStart'] = d['Anchor'].dt.to_period('W-MON').dt.start_time
            req = d.groupby('WeekStart', as_index=False)['Est_Hours'].sum().rename(columns={'Est_Hours':'Hours'})

    m = base.merge(req, on='WeekStart', how='left').sort_values('WeekStart')
    m['Hours'] = pd.to_numeric(m['Hours'], errors='coerce').fillna(0.0)
    m['Capacity'] = weekly_capacity
    m['Utilization'] = (m['Hours'] / m['Capacity']).replace([np.inf, -np.inf], np.nan)

    title = 'Upcoming Weekly Workload vs Capacity' + suffix

    base_chart = alt.Chart(m).encode(x=alt.X('WeekStart:T', title='Week'))
    bars = base_chart.mark_bar().encode(
        y=alt.Y('Hours:Q', title='Hours'),
        color=alt.condition(alt.datum.Utilization > 1.0, alt.value('#F28A94'), alt.value(get_active_theme()['ACCENT'])),
        tooltip=[
            alt.Tooltip('WeekStart:T', title='Week'),
            alt.Tooltip('Hours:Q', title='Hours', format=',.1f'),
            alt.Tooltip('Capacity:Q', title='Capacity', format=',.1f'),
            alt.Tooltip('Utilization:Q', title='Utilization', format='.0%'),
        ],
    )
    cap_line = base_chart.mark_line(color=get_active_theme()['PRIMARY'], strokeWidth=2).encode(
        y=alt.Y('Capacity:Q', title=''),
    )

    chart = (bars + cap_line).properties(width=1100, height=280, title=title)
    return chart, m

# ======= Aging/TAT helpers =======
def normalize_tracking_status(s: str) -> str:
    s = str(s or '').strip().lower()
    table = {
        'in shop': 'in shop',
        'in-shop': 'in shop',
        'wip': 'wip',
        'work in progress': 'wip',
        'receiving': 'receiving',
        'qc': 'qc',
        'quality control': 'qc',
        'awaiting qc': 'qc pending',
        'shop calibration': 'shop calibration',
        'shipping': 'shipping/returned',
        'courier delivered': 'shipping/returned',
        'picked up': 'shipping/returned',
        'return to customer': 'shipping/returned',
        'returned to customer': 'shipping/returned',
        'awaiting customer': 'awaiting customer',
        'on hold': 'on hold',
        'awaiting info': 'awaiting info',
    }
    if s in table: return table[s]
    if 'receiv' in s: return 'receiving'
    if 'cal' in s and 'shop' in s: return 'shop calibration'
    if s.startswith('qc') or 'qc ' in s: return 'qc'
    if any(k in s for k in ['ship','return to customer','courier delivered','picked up']): return 'shipping/returned'
    if 'wip' in s: return 'wip'
    if 'hold' in s: return 'on hold'
    if 'awaiting' in s: return 'awaiting'
    if 'shop' in s: return 'in shop'
    return s or '(unknown)'
def latest_tracking_status_from_file(path: str, sig: Tuple[float,int]) -> pd.DataFrame:
    e = events_enriched_from_file(path, sig)
    if e.empty or "Tracking Status" not in e.columns: return pd.DataFrame(columns=["Work Order","I.D.","Latest Tracking Status"])
    ev = e.dropna(subset=["event_ts"]).sort_values("event_ts")
    ev = ev[ev["Tracking Status"].astype("string").str.len() > 0]
    latest = ev.groupby(["Work Order","I.D."], as_index=False)["Tracking Status"].last()
    return latest.rename(columns={"Tracking Status":"Latest Tracking Status"})

@st.cache_data(ttl=1200, show_spinner=False)
def compute_aging_from_file(path: str, sig: Tuple[float,int]) -> pd.DataFrame:
    wos, _, equip = load_clean_data(path, sig)
    e = events_enriched_from_file(path, sig)
    if not e.empty and set(["Work Order","I.D."]).issubset(e.columns):
        base = e[["Work Order","I.D."]].dropna().drop_duplicates()
    else:
        base = pd.DataFrame({"Work Order": wos.get("Work Order", pd.Series(dtype="string")).dropna().unique()}); base["I.D."] = ""
    w = wos.copy() if wos is not None else pd.DataFrame()
    keep_cols = ["Work Order","Company","Open Date","Due Date","Completed Date","Status","Sub-Status"]
    for c in keep_cols:
        if c not in w.columns: w[c] = pd.NA
    w = w[keep_cols].drop_duplicates()
    out = base.merge(w, on="Work Order", how="left")
    shop_pairs = pd.DataFrame(columns=["Work Order","I.D.","is_shop"])  # not used on this view
    latest = latest_tracking_status_from_file(path, sig)
    out = out.merge(latest, on=["Work Order","I.D."], how="left")
    if equip is not None and not equip.empty and set(["Tracking Status","I.D."]).issubset(equip.columns):
        eq_ts = equip[["I.D.","Tracking Status"]].dropna().drop_duplicates().rename(columns={"Tracking Status":"Equip Tracking Status"})
        out = out.merge(eq_ts, on="I.D.", how="left"); out["Equipment Tracking Status"] = out["Latest Tracking Status"].combine_first(out["Equip Tracking Status"])
    else:
        out["Equipment Tracking Status"] = out.get("Latest Tracking Status", pd.NA)
    today = pd.Timestamp.today().normalize(); out["days_open"] = pd.NA; out["days_to_complete"] = pd.NA
    if "Open Date" in out.columns:
        mask_open = out["Open Date"].notna() & (out["Completed Date"].isna()); out.loc[mask_open, "days_open"] = (today - pd.to_datetime(out.loc[mask_open, "Open Date"]).dt.normalize()).dt.days
        mask_done = out["Open Date"].notna() & (out["Completed Date"].notna()); out.loc[mask_done, "days_to_complete"] = (pd.to_datetime(out.loc[mask_done, "Completed Date"]).dt.normalize() - pd.to_datetime(out.loc[mask_done, "Open Date"]).dt.normalize()).dt.days
    keep = ["Work Order","I.D.","Company","Open Date","Due Date","Completed Date","Status","Sub-Status","Equipment Tracking Status","days_open","days_to_complete"]
    for k in keep:
        if k not in out.columns: out[k] = pd.NA
    return out[keep]

@st.cache_data(ttl=1200, show_spinner=False)
def rolling_tat_365d_from_file(path: str, sig: Tuple[float,int], window_days: int=365) -> pd.DataFrame:
    # This uses events-enriched data restricted to shop? We reuse shipping timestamps where present
    e = events_enriched_from_file(path, sig)
    if e.empty: return pd.DataFrame(columns=["Ship Day","daily_avg","tat_count","rolling_avg","rolling_count"])
    # Infer simple TAT: receiving -> shipping per (WO, I.D.)
    e2 = e.copy()
    def stage_tag(s: str):
        s = str(s)
        if 'receiv' in s: return 'receiving'
        if ('ship' in s) or ('return to customer' in s) or ('courier delivered' in s) or ('picked up' in s): return 'shipping'
        return None
    e2['Stage'] = e2['__etype_norm__'].map(stage_tag)
    xrows = []
    for key, g in e2.dropna(subset=['event_ts']).sort_values('event_ts').groupby(['Work Order','I.D.']):
        r = g.loc[g['Stage']=='receiving','event_ts']
        s = g.loc[g['Stage']=='shipping','event_ts']
        if r.empty or s.empty: continue
        r = r.min(); s = s.max()
        _dd = (s - r).days
        if _dd >= 0:
            xrows.append({'Ship Day': s.normalize(), 'd_total_tat': _dd})
    x = pd.DataFrame(xrows)
    if x.empty: return pd.DataFrame(columns=["Ship Day","daily_avg","tat_count","rolling_avg","rolling_count"])
    daily = x.groupby('Ship Day').agg(tat_sum=('d_total_tat','sum'), tat_count=('d_total_tat','size'), daily_avg=('d_total_tat','mean')).sort_index()
    all_days = pd.date_range(daily.index.min(), daily.index.max(), freq="D"); daily = daily.reindex(all_days)
    daily.index.name = 'Ship Day'; daily[['tat_sum','tat_count']] = daily[['tat_sum','tat_count']].fillna(0)
    roll_sum = daily['tat_sum'].rolling(f"{int(window_days)}D", min_periods=1).sum(); roll_cnt = daily['tat_count'].rolling(f"{int(window_days)}D", min_periods=1).sum()
    daily['rolling_avg'] = (roll_sum / roll_cnt).where(roll_cnt > 0); daily['rolling_count'] = roll_cnt
    return daily.reset_index()[["Ship Day","daily_avg","tat_count","rolling_avg","rolling_count"]]

def _brand_header_min():
    col_logo, col_title, col_account = st.columns([1,6,2])
    with col_logo:
        try:
            st.image(LOGO_PATH)
        except Exception:
            st.write("")
    with col_title:
        st.markdown(
            f"<h1 style='margin-bottom:0;color:var(--ctpm-primary);'>{APP_TITLE}</h1>"
            f"<div style='color:var(--ctpm-secondary);font-weight:800;'>CTPM • Calibration Management</div>",
            unsafe_allow_html=True,
        )
    with col_account:
        auth = st.session_state.get("auth", {})
        if auth.get("is_authenticated"):
            st.caption(f"Signed in as **{auth.get('username','')}** ({auth.get('role','')})")
        else:
            st.caption("Not signed in")
    st.markdown("<div class='app-header-bar'></div>", unsafe_allow_html=True)

# ======= Presets storage =======
def _load_presets() -> Dict[str, Dict]:
    try:
        if PRESETS_FILE.exists(): return json.loads(PRESETS_FILE.read_text())
    except Exception: pass
    return {"tat":{}, "aging":{}}

def _save_presets(presets: Dict[str, Dict]):
    try: PRESETS_FILE.write_text(json.dumps(presets, indent=2))
    except Exception: pass


# ======= Technician Efficiency helpers =======
@st.cache_data(ttl=1200, show_spinner=False)
def tech_efficiency_from_events(events: pd.DataFrame, wos: pd.DataFrame) -> pd.DataFrame:
    """Compute technician efficiency from All Events.

    Business rules (per CTPM workflow):
      - Source: All Events sheet.
      - Calibration events: Event Type in {Calibration, Shop Calibration, Field Calibration} (plus other strings containing 'calibration').
      - Receiving timestamp: earliest Event Type == "Receiving In-Shop" for the SAME (Work Order, I.D.).
      - Lab time (days): (Calibration event completed Date+Time) - (Receiving In-Shop completed Date+Time).
      - Technician attribution: Entered By on the calibration event.

    Output rows are calibration-event rows. A piece may appear multiple times if it has multiple calibration events.
    The UI can optionally reduce to the latest calibration per (Work Order, I.D.).

    IMPORTANT: Events are tied strictly by (Work Order, I.D.) (no cross-WO fallback).
    """
    cols_out = [
        'Technician','EventType','Work Order','I.D.','Company',
        'Completed_TS','Recv_TS','Lab_Days','LabTime_Anomaly',
        'Due Date','OnTime','DaysLate'
    ]
    if events is None or not isinstance(events, pd.DataFrame) or events.empty:
        return pd.DataFrame(columns=cols_out)

    e = events.copy()

    # Ensure expected columns exist
    for c in ['Work Order','I.D.','Company','Event Type','Entered By','Date','Time']:
        if c not in e.columns:
            e[c] = pd.NA

    # Completed timestamp uses Date + Time columns (when event was completed)
    dt = pd.to_datetime(e['Date'], errors='coerce')
    tt = pd.to_timedelta(e['Time'].astype(str), errors='coerce') if 'Time' in e.columns else pd.to_timedelta(pd.Series([None]*len(e)))
    e['event_ts'] = dt + tt

    # Normalize event types
    e['Event Type'] = e['Event Type'].astype(str).str.strip()
    etype_l = e['Event Type'].str.lower()

    calib_mask = etype_l.isin(['calibration','shop calibration','field calibration']) | etype_l.str.contains('calibration', na=False)
    recv_mask = e['Event Type'].astype(str).str.strip().eq('Receiving In-Shop')

    # Receiving timestamp: earliest Receiving In-Shop per (Work Order, I.D.)
    recv = (
        e[recv_mask & e['event_ts'].notna()]
        .groupby(['Work Order','I.D.'], as_index=False)['event_ts']
        .min()
        .rename(columns={'event_ts':'Recv_TS'})
    )

    cal = e[calib_mask & e['event_ts'].notna()].copy()
    if cal.empty:
        return pd.DataFrame(columns=cols_out)

    cal = cal.rename(columns={'event_ts':'Completed_TS', 'Event Type':'EventType', 'Entered By':'Technician'})

    out = cal[['Technician','EventType','Work Order','I.D.','Company','Completed_TS']].merge(
        recv, on=['Work Order','I.D.'], how='left'
    )

    # Lab time
    out['Lab_Days'] = (out['Completed_TS'] - out['Recv_TS']).dt.total_seconds() / 86400.0
    out['LabTime_Anomaly'] = out['Lab_Days'].notna() & (out['Lab_Days'] < 0)
    out.loc[out['LabTime_Anomaly'], 'Lab_Days'] = np.nan

    # Due Date (WO-level) and On-time
    if wos is not None and isinstance(wos, pd.DataFrame) and (not wos.empty) and 'Work Order' in wos.columns:
        w = wos.copy()
        if 'Due Date' in w.columns:
            w['Due Date'] = pd.to_datetime(w['Due Date'], errors='coerce')
        else:
            w['Due Date'] = pd.NaT
        due = w[['Work Order','Due Date']].drop_duplicates()
        out = out.merge(due, on='Work Order', how='left')
    else:
        out['Due Date'] = pd.NaT

    out['OnTime'] = out['Completed_TS'].notna() & out['Due Date'].notna() & (out['Completed_TS'] <= out['Due Date'])
    out['DaysLate'] = np.where(
        out['Completed_TS'].notna() & out['Due Date'].notna(),
        (out['Completed_TS'].dt.normalize() - out['Due Date'].dt.normalize()).dt.days,
        np.nan
    )

    # Clean technician
    out['Technician'] = out['Technician'].astype(str).str.strip()
    out.loc[out['Technician'].isin(['', 'nan', 'None', 'NULL', 'null']), 'Technician'] = '(Unknown)'

    for c in cols_out:
        if c not in out.columns:
            out[c] = pd.NA

    return out[cols_out]

# ======= App start =======
# Restore cloud-persisted files on cold start (no-op if S3 not configured)
_sync_from_s3()

st.set_page_config(page_title=APP_TITLE, layout="wide", page_icon="📊", initial_sidebar_state="collapsed")

# State defaults
if 'sb_theme_name' not in st.session_state: st.session_state['sb_theme_name'] = 'CTPM Light Premium'
if 'active_theme' not in st.session_state: st.session_state['active_theme'] = THEMES['CTPM Light Premium']
if 'data_file' not in st.session_state:
    # Prefer a previously-uploaded file; fall back to the local Windows default
    _upl_candidates = sorted((CACHE_DIR / "uploads").glob("All Data.*"), reverse=True)
    _upl_default = str(_upl_candidates[0]) if _upl_candidates else DEFAULT_DATA_FILE
    _default_path = _upl_default if Path(_upl_default).exists() else DEFAULT_DATA_FILE
    st.session_state['data_file'] = _default_path
if 'cert_root' not in st.session_state: st.session_state['cert_root'] = DEFAULT_CERT_ROOT
if 'notes_db' not in st.session_state: st.session_state['notes_db'] = DEFAULT_NOTES_DB
if 'dark_mode' not in st.session_state: st.session_state['dark_mode'] = False
if 'kiosk_mode' not in st.session_state: st.session_state['kiosk_mode'] = False


# Repair theme state defensively (prevents KeyError on reconnect)
get_active_theme()  # ensure theme present
# Shared grouping radius (miles) used on Scheduling + Map pages
if 'group_radius_miles' not in st.session_state:
 try:
  _s0 = load_settings()
  st.session_state['group_radius_miles'] = float(_s0.get('cluster_miles', 10.0))
 except Exception:
  st.session_state['group_radius_miles'] = 10.0

# ======= Sign-In page gate (global) =======
auth = st.session_state.get('auth', {'is_authenticated': False})

# EARLY EXIT when already authenticated — fall through to the rest of the app
if auth.get('is_authenticated', False):
    pass
else:
    st.title('🔐 Sign in to CTPM_LabViewer')
    conn = _auth_conn(AUTH_DB_PATH)

    try:
        user_count = conn.execute('SELECT COUNT(*) FROM users;').fetchone()[0]
    except Exception:
        user_count = 0

    if user_count == 0:
        rec = _hash_pw('admin123!')
        now = _dt.now(_tz.utc).isoformat(timespec='seconds').replace('+00:00','Z')
        conn.execute(
            'INSERT INTO users(username,role,password_hash,password_salt,password_iters,active,created_at,updated_at) '
            'VALUES (?,?,?,?,?,?,?,?);',
            ('admin','admin',rec['hash'],rec['salt'],rec['iters'],1,now,now)
        )
        conn.commit()
        st.info("Bootstrap admin 'admin' created. Use password: admin123! (change it after login)")

    _login_ui(conn, source="global-gate")  # draw the form only while unauthenticated
    st.stop()                              # <- halt the page on this pass

# ======= CSS + Header (render only after auth) =======
_inject_css(get_active_theme(), kiosk_mode=st.session_state.get('kiosk_mode', False))
brand_header()

# ======= Load data =======
try:
    with st.spinner("Loading workbook ..."):
        _upl_sheets = _upl_sheet_paths()
        _using_individual = _upl_sheets["events"] is not None and _upl_sheets["equipment"] is not None

        if _using_individual:
            path = st.session_state.get('data_file', DEFAULT_DATA_FILE)  # may not exist on cloud; ignored by loader
            sig  = _individual_sheets_sig()
        else:
            path = sanitize_path(st.session_state['data_file'])
            sig  = file_signature(path)

        wos, events, equip = load_clean_data(path, sig)
        companies    = load_companies_df(path, sig)
        wip_shop_df  = load_wip_shop_df(path, sig)
except FileNotFoundError:
    st.warning(
        "**No data file found.** "
        "Please go to **📤 Upload Data** and upload your IndySoft exports to get started."
    )
    st.session_state['nav_page'] = "📤 Upload Data"
    st.rerun()
except Exception as ex:
    st.error(f"Failed to load workbook: {ex}")
    st.stop()

# Cache computed views for this run
if '__aging__' not in st.session_state:
    st.session_state['__aging__'] = compute_aging_from_file(path, sig)
if '__wip__' not in st.session_state:
    st.session_state['__wip__'] = wip_chain(path, sig, equip)
if '__tatroll__' not in st.session_state:
    st.session_state['__tatroll__'] = rolling_tat_365d_from_file(path, sig)

aging_all = st.session_state['__aging__']
wc_all = st.session_state['__wip__']
tat_roll_all = st.session_state['__tatroll__']

PAGES = [
    "🏠 Dashboard",
    "📦 WIP Tracking",
    "🧰 Ops — Assets & Due",
    "🗓️ Scheduling (Smart)",
    "🗺️ On‑Site Map",
    "⌛ Work Order Aging",
    "⏱️ TAT (Shop)",
    "👷 Technician Efficiency",
    "📈 WO Efficiency",
    "📋 Weekly Report",
    "📤 Upload Data",
    "🔐 Admin — Users",
]

if 'nav_page' not in st.session_state:
    st.session_state['nav_page'] = PAGES[0]
if st.session_state['nav_page'] not in PAGES:
    st.session_state['nav_page'] = PAGES[0]

_nav_cols = st.columns(len(PAGES))
_nav_clicked = None
for _i, (_col, _p) in enumerate(zip(_nav_cols, PAGES)):
    with _col:
        _active = st.session_state['nav_page'] == _p
        if st.button(_p, key=f"_nav_{_i}", use_container_width=True,
                     type="primary" if _active else "secondary"):
            _nav_clicked = _p

if _nav_clicked and _nav_clicked != st.session_state['nav_page']:
    st.session_state['nav_page'] = _nav_clicked
    st.rerun()

page = st.session_state['nav_page']

# ---- Dashboard ----
if page=="🏠 Dashboard":
    alt.themes.enable("ctpm")
    _today = pd.Timestamp.today().normalize()

    # WIP Health — from WIP Shop sheet (authoritative current-shop list)
    _wip_shop = wip_shop_df.copy() if wip_shop_df is not None and not wip_shop_df.empty else pd.DataFrame()
    items_in_shop = int(_wip_shop['I.D.'].nunique()) if not _wip_shop.empty and 'I.D.' in _wip_shop.columns else 0
    if not _wip_shop.empty and 'Received On' in _wip_shop.columns:
        _shop_days = (_today - _wip_shop['Received On'].dt.normalize()).dt.days
        med_days_shop = int(_shop_days.dropna().median()) if not _shop_days.dropna().empty else 0
    else:
        med_days_shop = 0

    # Work Order Pipeline — from wos
    _wos = wos if wos is not None and not wos.empty else pd.DataFrame()
    wo_in_progress    = int((_wos['Status'] == 'INCOMPLETE').sum()) if 'Status' in _wos.columns else 0
    wo_complete_not_inv = int((_wos['Status'] == 'COMPLETE').sum()) if 'Status' in _wos.columns else 0
    wo_awaiting_qc    = int((_wos['Sub-Status'].astype(str).str.strip() == 'Awaiting QC').sum()) if 'Sub-Status' in _wos.columns else 0
    wo_invoiced_30d   = 0
    if 'Status' in _wos.columns and 'Completed Date' in _wos.columns:
        _cutoff = _today - pd.Timedelta(days=30)
        wo_invoiced_30d = int(((_wos['Status'] == 'INVOICED') & (_wos['Completed Date'].notna()) & (_wos['Completed Date'] >= _cutoff)).sum())

    st.markdown("<div class='section-title'>WIP Health</div>", unsafe_allow_html=True)
    _a, _b = st.columns(2)
    with _a: st.markdown(f"<div class='kpi'><div class='label'>Items in Shop</div><div class='value'>{items_in_shop:,}</div></div>", unsafe_allow_html=True)
    with _b: st.markdown(f"<div class='kpi'><div class='label'>Median Days in Shop</div><div class='value'>{med_days_shop} d</div></div>", unsafe_allow_html=True)

    st.markdown("<div class='section-title'>Work Order Pipeline</div>", unsafe_allow_html=True)
    _c1, _c2, _c3, _c4 = st.columns(4)
    with _c1: st.markdown(f"<div class='kpi'><div class='label'>In Progress</div><div class='value'>{wo_in_progress:,}</div></div>", unsafe_allow_html=True)
    with _c2: st.markdown(f"<div class='kpi' style='border-color:var(--ctpm-primary)'><div class='label'>Complete — Not Invoiced</div><div class='value'>{wo_complete_not_inv:,}</div></div>", unsafe_allow_html=True)
    with _c3: st.markdown(f"<div class='kpi'><div class='label'>Awaiting QC</div><div class='value'>{wo_awaiting_qc:,}</div></div>", unsafe_allow_html=True)
    with _c4: st.markdown(f"<div class='kpi'><div class='label'>Invoiced (30 d)</div><div class='value'>{wo_invoiced_30d:,}</div></div>", unsafe_allow_html=True)

    # Charts still use wip_chain for stage/aging/TAT views
    wc = wc_all[wc_all['In_Shop']].copy() if (wc_all is not None and not wc_all.empty and 'In_Shop' in wc_all.columns) else wc_all

    if wc is None or wc.empty:
        st.info("No event history found."); st.stop()

    wc_plot = charts_df(wc)

    c1,c2 = st.columns(2)
    with c1:
        mix = wc_plot['Stage_Status'].value_counts(dropna=False).rename_axis('Stage').reset_index(name='Count')
        mix['Stage'] = mix['Stage'].astype(str).replace({'nan': '(Unknown)'})
        st.altair_chart(
            alt.Chart(mix).mark_bar().encode(
                x=alt.X('Count:Q'), y=alt.Y('Stage:N', sort='-x'), color=alt.Color('Stage:N'), tooltip=['Stage','Count']
            ).properties(height=280, title='Stage mix').interactive(), )
    with c2:
        last = pd.DataFrame({'Date': wc_plot['Recv'].dt.floor('D')}).dropna()
        W = last.groupby('Date').size().rename('Count').reset_index()
        st.altair_chart(
            alt.Chart(W).mark_line(point=True).encode(
                x=alt.X('Date:T', title='Date'), y=alt.Y('Count:Q', title='WIP Received'),
                color=alt.value(get_active_theme()['PRIMARY'])
            ).properties(height=280, title='WIP — last 30–90 days (zoomable)').interactive(), )

    c3,c4 = st.columns(2)
    with c3:
        ages = pd.to_numeric(wc_plot['Days_Since_Received'], errors='coerce').dropna()
        ages = ages[ages <= 365]
        bins = pd.DataFrame({'Days': ages})
        st.altair_chart(
            alt.Chart(bins).mark_bar().encode(
                x=alt.X('Days:Q', bin=alt.Bin(maxbins=25), title='Days since Receiving',
                         scale=alt.Scale(domain=[0, 365])),
                y=alt.Y('count():Q', title='Items'),
                color=alt.value(get_active_theme()['ACCENT'])
            ).properties(height=280, title='Aging Histogram (≤365 days)').interactive(), )
    with c4:
        tat_roll = tat_roll_all.dropna(subset=['rolling_avg']) if tat_roll_all is not None else pd.DataFrame()
        st.altair_chart(
            alt.Chart(tat_roll).mark_line().encode(
                x=alt.X('Ship Day:T', title='Ship Day'), y=alt.Y('rolling_avg:Q', title='Rolling Avg TAT (days)'),
                color=alt.value(get_active_theme()['PRIMARY'])
            ).properties(height=280, title='TAT (Shop) — 365d rolling avg').interactive(), )

# ---- WIP Tracking ----
elif page=="📦 WIP Tracking":
    alt.themes.enable("ctpm")
    st.markdown("<div class='section-title'>WIP Tracking — In Shop</div>", unsafe_allow_html=True)

    today = pd.Timestamp.today().normalize()

    # ---- Build WIP from Equipment sheet (has Cal Due) + WIP Shop sheet (has WO, Received On) ----
    wip = equip.copy() if equip is not None else pd.DataFrame()
    if wip.empty or 'Tracking Status' not in wip.columns:
        st.info("No equipment data available."); st.stop()

    wip = wip[wip['Tracking Status'].astype(str).str.strip().str.lower() == 'wip shop'].copy()
    if wip.empty:
        st.info("No items currently in WIP Shop."); st.stop()

    # Days in shop from Current Event Date (complete — no nulls)
    if 'Current Event Date' in wip.columns:
        wip['Days_In_Shop'] = (today - pd.to_datetime(wip['Current Event Date'], errors='coerce').dt.normalize()).dt.days
    else:
        wip['Days_In_Shop'] = np.nan

    # Work Order — pull directly from WIP Shop sheet (faster + more accurate than events join)
    if wip_shop_df is not None and not wip_shop_df.empty and 'Work Order' in wip_shop_df.columns and 'I.D.' in wip_shop_df.columns:
        wo_map = wip_shop_df[['I.D.','Work Order']].dropna(subset=['Work Order']).drop_duplicates('I.D.')
        wip = wip.merge(wo_map, on='I.D.', how='left')
    else:
        wip['Work Order'] = pd.NA
    wip['Work Order'] = wip['Work Order'].fillna('(No WO)')

    # Enrich with WO status from wos
    if wos is not None and not wos.empty and 'Work Order' in wos.columns:
        wo_status = wos[['Work Order','Status','Sub-Status','Due Date']].drop_duplicates('Work Order')
        wip = wip.merge(wo_status, on='Work Order', how='left')

    # ---- Filters ----
    with st.expander("Filters", expanded=False):
        fc1, fc2 = st.columns(2)
        with fc1:
            company_opts = sorted(wip['Company'].dropna().astype(str).unique().tolist())
            company_filter = st.multiselect("Company", company_opts, key='wip_company')
        with fc2:
            search_q = st.text_input("Search I.D., Description, or Work Order", key='wip_search')

    if company_filter:
        wip = wip[wip['Company'].astype(str).isin(company_filter)]
    if search_q:
        _mask = wip['I.D.'].astype(str).str.contains(search_q, case=False, na=False)
        if 'Description' in wip.columns:
            _mask = _mask | wip['Description'].astype(str).str.contains(search_q, case=False, na=False)
        _mask = _mask | wip['Work Order'].astype(str).str.contains(search_q, case=False, na=False)
        wip = wip[_mask]

    if wip.empty:
        st.info("No items match the current filters."); st.stop()

    wip_plot = charts_df(wip)

    # ---- KPIs ----
    total_items    = int(wip['I.D.'].nunique())
    total_wos      = int(wip.loc[wip['Work Order'] != '(No WO)', 'Work Order'].nunique())
    total_companies = int(wip['Company'].nunique())
    _days          = pd.to_numeric(wip['Days_In_Shop'], errors='coerce').dropna()
    med_age        = int(_days.median()) if not _days.empty else 0
    over10         = int((_days > 10).sum())
    aging_14       = int((_days > 14).sum())

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1: st.markdown(f"<div class='kpi'><div class='label'>Items In Shop</div><div class='value'>{total_items:,}</div></div>", unsafe_allow_html=True)
    with k2: st.markdown(f"<div class='kpi'><div class='label'>Work Orders</div><div class='value'>{total_wos:,}</div></div>", unsafe_allow_html=True)
    with k3: st.markdown(f"<div class='kpi'><div class='label'>Companies</div><div class='value'>{total_companies:,}</div></div>", unsafe_allow_html=True)
    with k4: st.markdown(f"<div class='kpi'><div class='label'>Median Days</div><div class='value'>{med_age} d</div></div>", unsafe_allow_html=True)
    with k5: st.markdown(f"<div class='kpi' style='border-color:#FF9800'><div class='label'>&gt;10 d (yellow)</div><div class='value'>{over10:,}</div></div>", unsafe_allow_html=True)
    with k6: st.markdown(f"<div class='kpi' style='border-color:#F44336'><div class='label'>&gt;14 d (red)</div><div class='value'>{aging_14:,}</div></div>", unsafe_allow_html=True)

    # ---- Charts (collapsible) ----
    with st.expander("Charts & Analytics", expanded=False):
        _c1, _c2 = st.columns(2)
        with _c1:
            if not wip_plot.empty and wip_plot['Days_In_Shop'].notna().any():
                st.altair_chart(
                    alt.Chart(wip_plot).mark_circle(size=50).encode(
                        x=alt.X('Days_In_Shop:Q', title='Days in Shop'),
                        y=alt.Y('Company:N', sort='-x'),
                        color=alt.Color('Company:N'),
                        tooltip=['Company','Work Order','I.D.','Days_In_Shop',
                                 alt.Tooltip('Description:N') if 'Description' in wip_plot.columns else alt.Tooltip('I.D.:N')]
                    ).properties(height=280, title='WIP Age Strip — each dot = one item').interactive())
        with _c2:
            grp = wip_plot.groupby('Company').agg(
                Items=('I.D.', 'nunique'),
                MedDays=('Days_In_Shop', 'median')
            ).reset_index().sort_values('Items', ascending=False)
            if not grp.empty:
                st.altair_chart(
                    alt.Chart(grp).mark_bar().encode(
                        x=alt.X('Items:Q'), y=alt.Y('Company:N', sort='-x'),
                        color=alt.Color('MedDays:Q', scale=alt.Scale(scheme='reds')),
                        tooltip=['Company','Items','MedDays']
                    ).properties(height=280, title='Items by Company — color = median days').interactive())

        if not wip_plot.empty and wip_plot['Days_In_Shop'].notna().any():
            _hist = pd.DataFrame({'Days': pd.to_numeric(wip_plot['Days_In_Shop'], errors='coerce').dropna()})
            if not _hist.empty:
                st.altair_chart(
                    alt.Chart(_hist).mark_bar().encode(
                        x=alt.X('Days:Q', bin=alt.Bin(maxbins=20), title='Days in Shop'),
                        y=alt.Y('count():Q', title='Items'),
                        color=alt.value(get_active_theme()['ACCENT'])
                    ).properties(height=280, title='Age Distribution').interactive())

    # ---- Pre-attach notes ONCE before grouping ----
    db_path = st.session_state.get('notes_db', DEFAULT_NOTES_DB)
    wip_noted = with_notes(wip, db_path)

    # ---- Work Orders — grouped expanders ----
    st.markdown("<div class='section-title'>Work Orders</div>", unsafe_allow_html=True)

    wo_summary = wip_noted.groupby('Work Order').agg(
        company=('Company', 'first'),
        item_count=('I.D.', 'nunique'),
        oldest_days=('Days_In_Shop', 'max'),
        median_days=('Days_In_Shop', 'median'),
        wo_status=('Status', 'first'),
    ).reset_index().sort_values('oldest_days', ascending=False)

    WO_DISPLAY_LIMIT = 50
    show_all = False
    if len(wo_summary) > WO_DISPLAY_LIMIT:
        show_all = st.toggle(f"Show all {len(wo_summary)} work orders (top {WO_DISPLAY_LIMIT} shown)", value=False, key='wip_show_all')
    wo_display = wo_summary if show_all or len(wo_summary) <= WO_DISPLAY_LIMIT else wo_summary.head(WO_DISPLAY_LIMIT)

    # Build item-level column list
    _item_cols = ['I.D.']
    for _c in ['Description', 'Manufacturer', 'Model Number']:
        if _c in wip_noted.columns: _item_cols.append(_c)
    _item_cols.append('Days_In_Shop')
    for _c in ['Current Event Date', 'Estimated Cal. Time']:
        if _c in wip_noted.columns: _item_cols.append(_c)
    if 'Notes' in wip_noted.columns: _item_cols.append('Notes')
    _item_cols = [c for c in _item_cols if c in wip_noted.columns]

    for _, row in wo_display.iterrows():
        wo        = str(row['Work Order'])
        company   = str(row['company'])
        count     = int(row['item_count'])
        age       = int(row['oldest_days']) if pd.notna(row['oldest_days']) else 0
        wo_stat   = str(row['wo_status']) if pd.notna(row['wo_status']) else ''
        age_color = "#4CAF50" if age <= 10 else ("#FF9800" if age <= 14 else "#F44336")
        stat_badge = f"<span style='background:var(--ctpm-border);color:var(--ctpm-text);padding:2px 8px;border-radius:999px;font-size:0.78rem;font-weight:600;margin-left:4px;'>{wo_stat}</span>" if wo_stat else ""

        with st.expander(f"WO {wo}  —  {company}  |  {count} item(s)  |  {age}d oldest", expanded=False):
            st.markdown(
                f"<div style='display:flex;gap:12px;align-items:center;padding:8px 0 12px;"
                f"border-bottom:1px solid var(--ctpm-border);margin-bottom:12px;flex-wrap:wrap;'>"
                f"<span style='font-weight:900;color:var(--ctpm-primary);font-size:1.05rem;'>{company}</span>"
                f"<span style='background:{age_color};color:#fff;padding:2px 10px;border-radius:999px;"
                f"font-weight:800;font-size:0.85rem;'>{age}d</span>"
                f"<span style='color:var(--ctpm-text);opacity:0.7;'>{count} item(s)</span>"
                f"{stat_badge}"
                f"</div>", unsafe_allow_html=True)

            wo_items = wip_noted[wip_noted['Work Order'] == wo].copy()
            st.dataframe(
                wo_items[_item_cols].sort_values('Days_In_Shop', ascending=False),
                hide_index=True,
                column_config=build_date_column_config(wo_items),
            )

    # ---- Export ----
    st.divider()
    download_buttons(wip_noted, base_name='wip_shop', key_prefix='wip_export')

# ---- Ops — Assets & Due ----
elif page=="🧰 Ops — Assets & Due":
    alt.themes.enable("ctpm")
    st.subheader("Assets & Calibration Due")
    eq = equip.copy()
    if "Calibration Due" in eq.columns: eq["Calibration Due"] = pd.to_datetime(eq["Calibration Due"], errors='coerce')
    if "Last Calibration" in eq.columns: eq["Last Calibration"] = pd.to_datetime(eq["Last Calibration"], errors='coerce')

    with st.expander("Filters", expanded=False):
        due_rng = st.date_input("Calibration Due range", [])
        hide_no_due = st.checkbox("Hide items with no due date", value=True)
        active_only = st.checkbox("Active only", value=True)
        in_shop_only = st.checkbox("IN_SHOP = True only", value=False)
        max_days = st.number_input("Show due within N days", value=60, min_value=0, step=5)

    today = pd.Timestamp.today().normalize()
    if "Calibration Due" in eq.columns: eq["Days_To_Due"] = (eq["Calibration Due"].dt.normalize() - today).dt.days
    if hide_no_due and 'Calibration Due' in eq.columns: eq = eq[eq['Calibration Due'].notna()]
    if active_only and "Active" in eq.columns: eq = eq[eq["Active"]==True]
    if in_shop_only and "IN_SHOP" in eq.columns: eq = eq[eq["IN_SHOP"]==True]
    if len(due_rng)==2 and "Calibration Due" in eq.columns:
        eq = eq[(eq['Calibration Due']>=pd.to_datetime(due_rng[0])) & (eq['Calibration Due']<=pd.to_datetime(due_rng[1]))]
    if "Days_To_Due" in eq.columns and max_days is not None:
        eq = eq[eq["Days_To_Due"].notna() & (eq["Days_To_Due"] <= max_days)]

    # Exclude CTPM from charts only
    ov = eq[eq['Days_To_Due']<0].copy()
    ov_plot = charts_df(ov)

    st.dataframe(eq[[c for c in ["Company","I.D.","Description","Tracking Status","IN_SHOP","Last Calibration","Calibration Due","Days_To_Due"] if c in eq.columns]].sort_values(by=["Days_To_Due","Company","I.D."], na_position="last"), hide_index=True)
    download_buttons(eq[[c for c in ["Company","I.D.","Description","Tracking Status","IN_SHOP","Last Calibration","Calibration Due","Days_To_Due"] if c in eq.columns]], base_name='assets_due', key_prefix='assets_due')

    if not ov_plot.empty and 'Company' in ov_plot.columns:
        st.markdown("### Overdue Pareto — by Company")
        pareto = ov_plot.groupby('Company').size().reset_index(name='Overdue').sort_values('Overdue', ascending=False)
        pareto['Cum%'] = (pareto['Overdue'].cumsum() / pareto['Overdue'].sum()*100).round(1)
        left = alt.Chart(pareto).mark_bar().encode(y=alt.Y('Company:N', sort='-x'), x='Overdue:Q', tooltip=['Company','Overdue'])
        right = alt.Chart(pareto).mark_line(color=get_active_theme()['ACCENT']).encode(y=alt.Y('Cum%:Q', axis=alt.Axis(format='%')), x=alt.X('Company:N', sort=None))
        st.altair_chart(alt.layer(left, right).resolve_scale(y='independent').properties(height=280).interactive(), )


# ---- Scheduling (Smart) — FULL ----
elif page == "🗓️ Scheduling (Smart)":
    alt.themes.enable("ctpm")
    st.subheader("Smart Scheduling — capacity · distance · due windows · policies · grouping · lateness penalty")

    settings = load_settings()
    policies = pol_load_all()
    excluded_companies = {str(c) for c, v in (policies or {}).items() if bool(v.get("exclude", False))}

    # ============================
    # Settings + Policies + Presets  |  Filters + Inputs
    # ============================
    tools_col1, tools_col2 = st.columns([1, 1])
    with tools_col1:
        settings_pop = st.popover("⚙️ Scheduling Settings")
    with tools_col2:
        filters_pop = st.popover("🔎 Filters")

    # =========================================================
    # SETTINGS POPOVER (Policies + Presets)
    # =========================================================
    with settings_pop:
        # --- Reset / Clear Scheduling Settings ---
        if st.button("↩️ Reset scheduling settings", key="sched_reset_settings"):
            keys_to_clear = {
                "urgency_urgent_days",
                "urgency_soon_days",
                "overdue_ignore_days",
                "group_radius_miles",
                "pol_edit",
                "sched_preset_sel",
                "sched_preset_new",
                # Filters + Inputs
                "sched_company_sel", "sched_kind_sel", "sched_urgency_sel", "sched_search",
                "sched_miles_rng", "sched_hours_rng", "sched_due_rng",
                "sched_techs", "sched_hpt", "sched_vehicles", "sched_tpv", "sched_lab_zip",
                "sched_near_radius", "sched_far_radius", "sched_near_days", "sched_mid_days", "sched_far_months",
                "sched_est_per_item", "sched_horizon_months", "sched_late_penalty",
                "sched_weekdays_only", "sched_cluster_enable", "sched_cluster_zip3", "sched_small_thr",
            }
            for k in list(st.session_state.keys()):
                if k.startswith("sched_") or k in keys_to_clear:
                    st.session_state.pop(k, None)
            st.success("Scheduling settings cleared.")
            st.rerun()

        st.markdown("**Company-specific policies (persistent)**")
        st.caption(
            "Default cadence = distance windows; pick '6week' or 'quarterly'. "
            "Daily cap blank = no cap. Exclude removes company from optimization."
        )

        dfp = (
            pd.DataFrame({"Company": sorted(companies["Company"].dropna().astype(str).unique().tolist())})
            if companies is not None and (not companies.empty)
            else pd.DataFrame({"Company": []})
        )
        dfp["cadence"] = dfp["Company"].map(lambda c: policies.get(str(c), {}).get("cadence", "default"))
        dfp["daily_hours_cap"] = dfp["Company"].map(lambda c: policies.get(str(c), {}).get("daily_hours_cap", None))
        dfp["exclude"] = dfp["Company"].map(lambda c: bool(policies.get(str(c), {}).get("exclude", False)))

        _pol_search = st.text_input("🔍 Search companies", value="", key="pol_search", placeholder="Type to filter list…")
        _dfp_view = dfp[dfp["Company"].str.contains(_pol_search.strip(), case=False, na=False)].copy() if _pol_search.strip() else dfp.copy()
        st.caption(f"Showing {len(_dfp_view)} of {len(dfp)} companies" + (" — clear search to see all" if _pol_search.strip() else ""))

        edited = st.data_editor(
            _dfp_view,
            key=f"pol_edit_{_pol_search}",
            num_rows="fixed",
            column_config={
                "cadence": st.column_config.SelectboxColumn("Cadence", options=["default", "6week", "quarterly"], required=True),
                "daily_hours_cap": st.column_config.NumberColumn("Daily Cap (hrs)", min_value=0.0, step=0.5),
                "exclude": st.column_config.CheckboxColumn(
                    "Exclude",
                    help="Checked = this company is excluded from schedule optimization.",
                ),
            },
            use_container_width=True,
        )

        if st.button("💾 Save policies", type="primary", key="pol_save_btn"):
            new_pol = {}
            for _, r in edited.iterrows():
                cad = (str(r.get("cadence", "default")).lower() if pd.notna(r.get("cadence")) else "default")
                cap = float(r["daily_hours_cap"]) if pd.notna(r.get("daily_hours_cap")) else None
                ex = bool(r.get("exclude", False))
                # Save ALL visible rows so that unchecking exclude (ex=False) is written to DB
                new_pol[str(r["Company"])] = {"cadence": cad, "daily_hours_cap": cap, "exclude": ex}

            pol_save_many(new_pol)
            st.success(f"Saved {len(new_pol)} company policy/policies.")
            st.rerun()

        st.divider()

        # ---- Presets (Scheduling) ----
        st.markdown("**Presets**")
        presets = _load_presets()
        if "scheduling" not in presets:
            presets["scheduling"] = {}
        sched_presets = presets.get("scheduling", {})

        p1, p2, p3 = st.columns([2, 2, 3])
        with p1:
            preset_names = sorted(list(sched_presets.keys()))
            sel_preset = st.selectbox(
                "Load preset",
                options=["(None)"] + preset_names,
                index=0,
                key="sched_preset_sel",
            )
        with p2:
            new_preset_name = st.text_input("New preset name", value="", key="sched_preset_new")
        with p3:
            bsave, bapply, bdel = st.columns(3)
            with bsave:
                save_p = st.button("💾 Save current", key="sched_preset_save")
            with bapply:
                apply_p = st.button("📥 Apply preset", key="sched_preset_apply")
            with bdel:
                del_p = st.button("🗑️ Delete preset", key="sched_preset_delete")

        def _sched_cfg_from_state() -> dict:
            return {
                "sched_techs": int(st.session_state.get("sched_techs", int(settings.get("techs", 15)))),
                "sched_hpt": float(st.session_state.get("sched_hpt", float(settings.get("hours_per_tech", 8)))),
                "sched_vehicles": int(st.session_state.get("sched_vehicles", int(settings.get("vehicles", 6)))),
                "sched_tpv": int(st.session_state.get("sched_tpv", int(settings.get("techs_per_vehicle", 2)))),
                "sched_lab_zip": str(st.session_state.get("sched_lab_zip", str(settings.get("lab_zip", "45430")).zfill(5))),
                "sched_near_radius": int(st.session_state.get("sched_near_radius", int(settings.get("near_radius", 30)))),
                "sched_far_radius": int(st.session_state.get("sched_far_radius", int(settings.get("far_radius", 80)))),
                "sched_near_days": int(st.session_state.get("sched_near_days", int(settings.get("near_window_days", 30)))),
                "sched_mid_days": int(st.session_state.get("sched_mid_days", int(settings.get("mid_window_days", 45)))),
                "sched_far_months": int(st.session_state.get("sched_far_months", int(settings.get("far_window_months", 3)))),
                "sched_est_per_item": float(st.session_state.get("sched_est_per_item", float(settings.get("hours_per_item_default", 0.5)))),
                "sched_horizon_months": int(st.session_state.get("sched_horizon_months", int(settings.get("horizon_months", 6)))),
                "sched_late_penalty": float(st.session_state.get("sched_late_penalty", float(settings.get("late_penalty_per_day", 5.0)))),
                "urgency_urgent_days": int(st.session_state.get("urgency_urgent_days", int(settings.get("urgency_urgent_days", 3)))),
                "urgency_soon_days": int(st.session_state.get("urgency_soon_days", int(settings.get("urgency_soon_days", 10)))),
                "overdue_ignore_days": int(st.session_state.get("overdue_ignore_days", int(settings.get("overdue_ignore_days", 90)))),
                "sched_weekdays_only": bool(st.session_state.get("sched_weekdays_only", bool(settings.get("weekdays_only", True)))),
                "sched_cluster_enable": bool(st.session_state.get("sched_cluster_enable", bool(settings.get("cluster_enable", True)))),
                "sched_cluster_zip3": bool(st.session_state.get("sched_cluster_zip3", bool(settings.get("cluster_by_zip3", False)))),
                "group_radius_miles": float(st.session_state.get("group_radius_miles", float(settings.get("cluster_miles", 10.0)))),
                "sched_small_thr": float(st.session_state.get("sched_small_thr", float(settings.get("small_hours_threshold", 6.0)))),
            }

        if save_p:
            name = (new_preset_name or "").strip()
            if not name:
                st.warning("Enter a preset name to save.")
            else:
                presets["scheduling"][name] = _sched_cfg_from_state()
                _save_presets(presets)
                st.success(f"Saved preset: {name}")
                st.rerun()

        if apply_p and sel_preset and sel_preset != "(None)":
            cfg = presets.get("scheduling", {}).get(sel_preset, {})
            if isinstance(cfg, dict):
                for k, v in cfg.items():
                    st.session_state[k] = v
                st.success(f"Applied preset: {sel_preset}")
                st.rerun()

        if del_p and sel_preset and sel_preset != "(None)":
            presets.get("scheduling", {}).pop(sel_preset, None)
            _save_presets(presets)
            st.success(f"Deleted preset: {sel_preset}")
            st.rerun()

    # =========================================================
    # FILTERS POPOVER (Demand Filters + Scheduling Inputs)
    # =========================================================
    with filters_pop:
        with st.form("sched_filters_form", clear_on_submit=False):

            st.markdown("### Demand Filters")

            fc1, fc2, fc3, fc4 = st.columns(4)
            with fc1:
                st.multiselect(
                    "Company",
                    options=sorted(
                        companies.get("Company", pd.Series(dtype="string"))
                        .dropna().astype(str).unique().tolist()
                    ) if companies is not None and not companies.empty else [],
                    key="sched_company_sel",
                )
            with fc2:
                st.multiselect("Kind", options=["near", "mid", "far"], key="sched_kind_sel")
            with fc3:
                st.multiselect(
                    "Urgency",
                    options=["URGENT", "HIGH", "Normal"],
                    help="Filter by urgency level",
                    key="sched_urgency_sel",
                )
            with fc4:
                st.text_input("Search Company/IDs", key="sched_search")

            rc1, rc2, rc3 = st.columns(3)
            with rc1:
                st.slider("Miles from Lab", 0, 600, (0, 600), step=10, key="sched_miles_rng")
            with rc2:
                st.slider("Estimated hours", 0.0, 200.0, (0.0, 200.0), step=1.0, key="sched_hours_rng")
            with rc3:
                st.date_input("Earliest due date range", value=[], key="sched_due_rng")

            st.divider()
            st.markdown("### Scheduling Inputs")

            r1c1, r1c2, r1c3, r1c4, r1c5 = st.columns(5)
            with r1c1:
                st.number_input("Technicians", value=int(st.session_state.get("sched_techs", settings.get("techs", 15))), min_value=1, key="sched_techs")
            with r1c2:
                st.number_input("Hours/tech/day", value=float(st.session_state.get("sched_hpt", settings.get("hours_per_tech", 8))), step=0.5, key="sched_hpt")
            with r1c3:
                st.number_input("Vehicles (max routes/day)", value=int(st.session_state.get("sched_vehicles", settings.get("vehicles", 6))), min_value=1, key="sched_vehicles")
            with r1c4:
                st.number_input("Techs/vehicle", value=int(st.session_state.get("sched_tpv", settings.get("techs_per_vehicle", 2))), min_value=1, key="sched_tpv")
            with r1c5:
                st.text_input("Lab ZIP", value=str(st.session_state.get("sched_lab_zip", str(settings.get("lab_zip", "45430")).zfill(5))), key="sched_lab_zip")

            r2c1, r2c2, r2c3, r2c4, r2c5 = st.columns(5)
            with r2c1:
                st.number_input("Near radius (mi)", value=int(st.session_state.get("sched_near_radius", settings.get("near_radius", 30))), key="sched_near_radius")
            with r2c2:
                st.number_input("Far radius (mi)", value=int(st.session_state.get("sched_far_radius", settings.get("far_radius", 80))), key="sched_far_radius")
            with r2c3:
                st.number_input("Near window (days)", value=int(st.session_state.get("sched_near_days", settings.get("near_window_days", 30))), key="sched_near_days")
            with r2c4:
                st.number_input("Mid window (days)", value=int(st.session_state.get("sched_mid_days", settings.get("mid_window_days", 45))), key="sched_mid_days")
            with r2c5:
                st.number_input("Far window (months)", value=int(st.session_state.get("sched_far_months", settings.get("far_window_months", 3))), key="sched_far_months")

            r3c1, r3c2, r3c3 = st.columns(3)
            with r3c1:
                st.number_input("Est hours per item", value=float(st.session_state.get("sched_est_per_item", settings.get("hours_per_item_default", 0.5))), step=0.1, key="sched_est_per_item")
            with r3c2:
                st.number_input("Horizon (months)", value=int(st.session_state.get("sched_horizon_months", settings.get("horizon_months", 6))), min_value=1, key="sched_horizon_months")
            with r3c3:
                st.number_input("Late penalty (hrs/day)", value=float(st.session_state.get("sched_late_penalty", settings.get("late_penalty_per_day", 5.0))), step=0.5, key="sched_late_penalty")

            st.number_input("Urgency: due within (days) ⇒ URGENT", value=int(st.session_state.get("urgency_urgent_days", settings.get("urgency_urgent_days", 3))), min_value=0, step=1, key="urgency_urgent_days")
            st.number_input("Urgency: due within (days) ⇒ HIGH", value=int(st.session_state.get("urgency_soon_days", settings.get("urgency_soon_days", 10))), min_value=0, step=1, key="urgency_soon_days")
            st.number_input("Ignore items more than (days) past due", value=int(st.session_state.get("overdue_ignore_days", settings.get("overdue_ignore_days", 90))), min_value=0, step=5, key="overdue_ignore_days")

            t1, t2, t3, t4, t5 = st.columns(5)
            with t1:
                st.toggle("Weekdays only", value=bool(st.session_state.get("sched_weekdays_only", settings.get("weekdays_only", True))), key="sched_weekdays_only")
            with t2:
                st.toggle("Enable grouping (routes)", value=bool(st.session_state.get("sched_cluster_enable", settings.get("cluster_enable", True))), key="sched_cluster_enable")
            with t3:
                st.toggle("Cluster by ZIP3 (fast)", value=bool(st.session_state.get("sched_cluster_zip3", settings.get("cluster_by_zip3", False))), key="sched_cluster_zip3")
            with t4:
                st.number_input("Group radius (mi)", value=float(st.session_state.get("group_radius_miles", settings.get("cluster_miles", 10.0))), step=1.0, key="group_radius_miles")
            with t5:
                st.number_input("Group small loads < hrs", value=float(st.session_state.get("sched_small_thr", settings.get("small_hours_threshold", 6.0))), step=0.5, key="sched_small_thr")

            cfa, cfr = st.columns([1, 1])
            with cfa:
                apply_filters = st.form_submit_button("✅ Apply", type="primary")
            with cfr:
                reset_filters = st.form_submit_button("↩️ Reset")

        if reset_filters:
            for k in (
                "sched_company_sel","sched_kind_sel","sched_urgency_sel","sched_search",
                "sched_miles_rng","sched_hours_rng","sched_due_rng",
                "sched_techs","sched_hpt","sched_vehicles","sched_tpv","sched_lab_zip",
                "sched_near_radius","sched_far_radius","sched_near_days","sched_mid_days","sched_far_months",
                "sched_est_per_item","sched_horizon_months","sched_late_penalty",
                "urgency_urgent_days","urgency_soon_days","overdue_ignore_days",
                "sched_weekdays_only","sched_cluster_enable","sched_cluster_zip3",
                "group_radius_miles","sched_small_thr",
            ):
                st.session_state.pop(k, None)
            st.success("Filters + inputs cleared.")
            st.rerun()

    # ============================
    # Demand (drives whole page)
    # ============================
    # Read the form-backed values from session_state
    techs = int(st.session_state.get("sched_techs", settings.get("techs", 15)))
    hpt = float(st.session_state.get("sched_hpt", settings.get("hours_per_tech", 8)))
    vehicles = int(st.session_state.get("sched_vehicles", settings.get("vehicles", 6)))
    t_per_v = int(st.session_state.get("sched_tpv", settings.get("techs_per_vehicle", 2)))
    lab_zip = str(st.session_state.get("sched_lab_zip", str(settings.get("lab_zip", "45430")).zfill(5)))

    near_radius = int(st.session_state.get("sched_near_radius", settings.get("near_radius", 30)))
    far_radius = int(st.session_state.get("sched_far_radius", settings.get("far_radius", 80)))
    near_days = int(st.session_state.get("sched_near_days", settings.get("near_window_days", 30)))
    mid_days = int(st.session_state.get("sched_mid_days", settings.get("mid_window_days", 45)))
    far_months = int(st.session_state.get("sched_far_months", settings.get("far_window_months", 3)))

    est_per_item = float(st.session_state.get("sched_est_per_item", settings.get("hours_per_item_default", 0.5)))
    horizon_months = int(st.session_state.get("sched_horizon_months", settings.get("horizon_months", 6)))
    late_penalty = float(st.session_state.get("sched_late_penalty", settings.get("late_penalty_per_day", 5.0)))

    wkdays = bool(st.session_state.get("sched_weekdays_only", settings.get("weekdays_only", True)))
    clus_enable = bool(st.session_state.get("sched_cluster_enable", settings.get("cluster_enable", True)))
    clus_zip3 = bool(st.session_state.get("sched_cluster_zip3", settings.get("cluster_by_zip3", False)))
    clus_miles = float(st.session_state.get("group_radius_miles", settings.get("cluster_miles", 10.0)))
    small_thr = float(st.session_state.get("sched_small_thr", settings.get("small_hours_threshold", 6.0)))

    demand = build_demand(
        equip, wos, companies,
        lab_zip, near_radius, far_radius,
        near_days, mid_days, far_months,
        est_per_item, horizon_months,
        policies, clus_enable, clus_zip3, clus_miles, small_thr,
        settings.get("assumed_mph", 40.0),
        settings.get("per_stop_travel_hours", 0.25),
    )
    if demand is None:
        demand = pd.DataFrame()

    if excluded_companies and (not demand.empty) and "Company" in demand.columns:
        demand = demand[~demand["Company"].astype(str).isin(excluded_companies)].copy()

    # Optional grouping by distance
    try:
        if clus_enable and float(st.session_state.get("group_radius_miles", 0)) > 0 and (not demand.empty):
            ship = companies_shipping_only(companies)
            locs = (
                compute_company_distance(ship, str(lab_zip).zfill(5))
                if ship is not None and not ship.empty
                else pd.DataFrame(columns=["Company", "lat", "lon"])
            )
            demand = group_companies_by_distance(
                demand,
                locs[["Company", "lat", "lon"]] if "lat" in locs.columns else locs,
                float(st.session_state.get("group_radius_miles", 0)),
                float(small_thr),
            )
    except Exception:
        pass

    # Apply filters to demand
    if not demand.empty:
        if st.session_state.get("sched_company_sel") and "Company" in demand.columns:
            demand = demand[demand["Company"].astype(str).isin([str(x) for x in st.session_state["sched_company_sel"]])].copy()

        if st.session_state.get("sched_kind_sel") and "Kind" in demand.columns:
            demand = demand[demand["Kind"].astype(str).str.lower().isin([str(x).lower() for x in st.session_state["sched_kind_sel"]])].copy()

        if st.session_state.get("sched_urgency_sel") and "Urgency" in demand.columns:
            _u_label_to_int = {"URGENT": 0, "HIGH": 1, "Normal": 2}
            _u_ints = [_u_label_to_int[s] for s in st.session_state["sched_urgency_sel"] if s in _u_label_to_int]
            if _u_ints:
                demand = demand[pd.to_numeric(demand["Urgency"], errors="coerce").isin(_u_ints)].copy()

        q = str(st.session_state.get("sched_search", "") or "").strip()
        if q:
            m1 = demand.get("Company", pd.Series(index=demand.index, dtype="string")).astype(str).str.contains(q, case=False, na=False)
            m2 = demand.get("IDs", pd.Series(index=demand.index, dtype="string")).astype(str).str.contains(q, case=False, na=False)
            demand = demand[m1 | m2].copy()

        if "MilesFromLab" in demand.columns:
            lo, hi = st.session_state.get("sched_miles_rng", (0, 600))
            mm = pd.to_numeric(demand["MilesFromLab"], errors="coerce")
            demand = demand[mm.isna() | mm.between(float(lo), float(hi))].copy()

        if "Est_Hours" in demand.columns:
            loh, hih = st.session_state.get("sched_hours_rng", (0.0, 200.0))
            hh = pd.to_numeric(demand["Est_Hours"], errors="coerce")
            demand = demand[hh.isna() | hh.between(float(loh), float(hih))].copy()

        due_rng = st.session_state.get("sched_due_rng", [])
        if isinstance(due_rng, (list, tuple)) and len(due_rng) == 2 and all(due_rng) and "Earliest_Due" in demand.columns:
            sdt = pd.to_datetime(due_rng[0])
            edt = pd.to_datetime(due_rng[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            ddd = pd.to_datetime(demand["Earliest_Due"], errors="coerce")
            demand = demand[ddd.isna() | ddd.between(sdt, edt)].copy()

    if demand.empty:
        st.info("No on-site demand found in the selected horizon (or filters removed all rows).")
        st.stop()

    # ---- Demand KPIs ----
    _d_items   = int(demand['Count_Items'].sum()) if 'Count_Items' in demand.columns else 0
    _d_cos     = int(demand['Company'].nunique())
    _d_hrs     = float(demand['Est_Hours'].sum()) if 'Est_Hours' in demand.columns else 0.0
    _d_urgent  = int((demand['Urgency'] == 0).sum()) if 'Urgency' in demand.columns else 0
    _d_high    = int((demand['Urgency'] == 1).sum()) if 'Urgency' in demand.columns else 0
    _daily_cap = float(techs) * float(hpt)
    _days_need = math.ceil(_d_hrs / _daily_cap) if _daily_cap > 0 else 0
    _near_ct   = int((demand['Kind'] == 'near').sum()) if 'Kind' in demand.columns else 0
    _mid_ct    = int((demand['Kind'] == 'mid').sum()) if 'Kind' in demand.columns else 0
    _far_ct    = int((demand['Kind'] == 'far').sum()) if 'Kind' in demand.columns else 0

    dk1, dk2, dk3, dk4, dk5, dk6, dk7 = st.columns(7)
    with dk1: st.markdown(f"<div class='kpi'><div class='label'>Companies</div><div class='value'>{_d_cos:,}</div></div>", unsafe_allow_html=True)
    with dk2: st.markdown(f"<div class='kpi'><div class='label'>Items</div><div class='value'>{_d_items:,}</div></div>", unsafe_allow_html=True)
    with dk3: st.markdown(f"<div class='kpi'><div class='label'>Est Hours</div><div class='value'>{_d_hrs:,.0f}</div></div>", unsafe_allow_html=True)
    with dk4: st.markdown(f"<div class='kpi'><div class='label'>Work Days Needed</div><div class='value'>{_days_need}</div></div>", unsafe_allow_html=True)
    with dk5: st.markdown(f"<div class='kpi' style='border-color:#F44336'><div class='label'>Urgent</div><div class='value'>{_d_urgent:,}</div></div>", unsafe_allow_html=True)
    with dk6: st.markdown(f"<div class='kpi' style='border-color:#FF9800'><div class='label'>High Priority</div><div class='value'>{_d_high:,}</div></div>", unsafe_allow_html=True)
    with dk7: st.markdown(f"<div class='kpi'><div class='label'>Near / Mid / Far</div><div class='value'>{_near_ct} / {_mid_ct} / {_far_ct}</div></div>", unsafe_allow_html=True)

    # ---- 1) Demand Table ----
    st.markdown("### 1) Demand windows (with policies & grouping)")
    _URGENCY_LABELS = {0: "URGENT", 1: "HIGH", 2: "Normal"}
    show_cols = ["Company", "Kind", "MilesFromLab", "Window_Start", "Window_End", "Earliest_Due", "Due_Median", "Count_Items", "Est_Hours", "Urgency"]
    show_cols = [c for c in show_cols if c in demand.columns]
    show_dem = demand[show_cols].copy()
    if 'Urgency' in show_dem.columns:
        show_dem['Urgency'] = show_dem['Urgency'].map(lambda u: _URGENCY_LABELS.get(int(u) if pd.notna(u) else 2, "Normal"))
    show_dem = show_dem.sort_values([c for c in ["Kind", "Window_Start", "Company"] if c in show_cols])
    with st.expander(f"Demand table — {len(show_dem)} companies", expanded=True):
        st.dataframe(show_dem, hide_index=True, column_config=build_date_column_config(show_dem))
        st.download_button("⬇️ Demand CSV", data=show_dem.to_csv(index=False).encode("utf-8"), file_name="demand_windows.csv")

    # ---- 2) Capacity Planning ----
    st.markdown("### 2) Capacity planning (Utilization)")
    sched_for_chart = st.session_state.get('__sched__')
    chart, cap_df = capacity_planning_chart(
        demand,
        weeks=int(settings.get('capacity_weeks', 8)),
        techs=int(techs),
        hours_per_tech=float(hpt),
        weekdays_only=bool(wkdays),
        schedule_df=sched_for_chart if isinstance(sched_for_chart, pd.DataFrame) and (not sched_for_chart.empty) else None,
    )
    st.altair_chart(chart)

    # Over-capacity weeks warning
    if cap_df is not None and not cap_df.empty and 'Utilization' in cap_df.columns:
        _over = cap_df[cap_df['Utilization'] > 1.0]
        if not _over.empty:
            _over_strs = [f"{r['WeekStart'].strftime('%b %d')} ({r['Utilization']:.0%})" for _, r in _over.iterrows()]
            st.warning(f"**Over-capacity weeks ({len(_over)}):** {', '.join(_over_strs)}  —  Increase techs, adjust horizon, or activate grouping to resolve.")
        else:
            st.success(f"All weeks within capacity across {len(cap_df)} week(s) scheduled. Daily cap: {_daily_cap:.0f} hrs.")

    # ---- 3) Optimize ----
    st.markdown("### 3) Optimize schedule")
    _oc1, _oc2 = st.columns([1, 4])
    with _oc1:
        run_opt = st.button("Optimize now", type="primary")
    with _oc2:
        if st.session_state.get('__sched__') is not None:
            if st.button("Clear schedule", key="sched_clear"):
                st.session_state.pop('__sched__', None)
                st.rerun()

    if run_opt:
        sched = optimize_schedule(
            demand,
            policies,
            int(techs),
            float(hpt),
            int(vehicles),
            int(t_per_v),
            weekdays_only=bool(wkdays),
            late_penalty_per_day=float(late_penalty),
        )
        st.session_state['__sched__'] = sched
        st.rerun()

    sched_saved = st.session_state.get('__sched__')
    if isinstance(sched_saved, pd.DataFrame) and (not sched_saved.empty):

        # ---- Schedule summary KPIs ----
        _sc_placed   = sched_saved.dropna(subset=['Date'])
        _sc_unplaced = sched_saved[sched_saved['Date'].isna()]
        _sc_cos      = int(_sc_placed['Company'].nunique())
        _sc_hrs      = float(_sc_placed['Est_Hours'].sum())
        _sc_days     = int(_sc_placed['Date'].nunique())
        _sc_unco     = int(_sc_unplaced['Company'].nunique()) if not _sc_unplaced.empty else 0
        _peak_util   = float(cap_df['Utilization'].max()) if cap_df is not None and 'Utilization' in cap_df.columns else 0.0
        _sc_items    = int(_sc_placed['Count_Items'].sum()) if 'Count_Items' in _sc_placed.columns else 0

        sk1, sk2, sk3, sk4, sk5, sk6 = st.columns(6)
        with sk1: st.markdown(f"<div class='kpi'><div class='label'>Companies Placed</div><div class='value'>{_sc_cos:,}</div></div>", unsafe_allow_html=True)
        with sk2: st.markdown(f"<div class='kpi'><div class='label'>Items Scheduled</div><div class='value'>{_sc_items:,}</div></div>", unsafe_allow_html=True)
        with sk3: st.markdown(f"<div class='kpi'><div class='label'>Total Hours</div><div class='value'>{_sc_hrs:,.0f}</div></div>", unsafe_allow_html=True)
        with sk4: st.markdown(f"<div class='kpi'><div class='label'>Work Days Span</div><div class='value'>{_sc_days}</div></div>", unsafe_allow_html=True)
        with sk5: st.markdown(f"<div class='kpi' style='border-color:#{'F44336' if _peak_util>1 else '4CAF50'}'><div class='label'>Peak Week Util</div><div class='value'>{_peak_util:.0%}</div></div>", unsafe_allow_html=True)
        with sk6: st.markdown(f"<div class='kpi' style='border-color:#{'F44336' if _sc_unco>0 else '4CAF50'}'><div class='label'>Unscheduled Cos</div><div class='value'>{_sc_unco}</div></div>", unsafe_allow_html=True)

        # Unscheduled warning
        if not _sc_unplaced.empty:
            _unpl_names = ', '.join(_sc_unplaced['Company'].astype(str).unique()[:10])
            st.warning(f"**{_sc_unco} company/companies could not be placed** (no day with enough capacity + visits in their window): {_unpl_names}. "
                       f"Try increasing techs/vehicles, extending horizon, or widening windows.")

        # Lateness risk: companies whose earliest scheduled date > their Earliest_Due
        if 'Earliest_Due' in _sc_placed.columns:
            _late_check = _sc_placed.groupby('Company').agg(
                First_Sched=('Date', 'min'),
                Earliest_Due=('Earliest_Due', 'min'),
            ).reset_index()
            _late_check['First_Sched'] = pd.to_datetime(_late_check['First_Sched'], errors='coerce')
            _late_check['Earliest_Due'] = pd.to_datetime(_late_check['Earliest_Due'], errors='coerce')
            _late_check['Days_Late'] = (_late_check['First_Sched'] - _late_check['Earliest_Due']).dt.days
            _late_risk = _late_check[_late_check['Days_Late'] > 0].sort_values('Days_Late', ascending=False)
            if not _late_risk.empty:
                st.error(f"**Lateness risk — {len(_late_risk)} company/companies scheduled after their earliest due date:**")
                st.dataframe(_late_risk[['Company','Earliest_Due','First_Sched','Days_Late']].rename(
                    columns={'Days_Late':'Days_Past_Due'}),
                    hide_index=True, column_config=build_date_column_config(_late_risk))

        # ---- Job spans ----
        st.markdown('#### Job spans (Start → End)')
        sched_span = _sc_placed.groupby('Company', as_index=False).agg(
            Start=('Date','min'),
            End=('Date','max'),
            Days_Scheduled=('Date', lambda s: int(pd.to_datetime(s, errors='coerce').dt.normalize().nunique())),
            Total_Hours=('Est_Hours','sum'),
            Techs_Suggested=('Techs_Suggested','max'),
            Vehicles_Needed=('Vehicles_Needed','max'),
            MilesFromLab=('MilesFromLab','max'),
            Count_Items=('Count_Items','sum'),
            Earliest_Due=('Earliest_Due','min'),
        )
        sched_span = sched_span.sort_values(['Start','Company'])
        st.dataframe(sched_span, hide_index=True, column_config=build_date_column_config(sched_span))
        st.download_button('⬇️ Proposed_Schedule_Spans.csv', data=sched_span.to_csv(index=False).encode('utf-8'), file_name='Proposed_Schedule_Spans.csv')

        # ---- Daily detail ----
        st.markdown('#### Daily schedule detail')
        cols = ['Date','Company','Kind','Est_Hours','Techs_Suggested','Vehicles_Needed','MilesFromLab','Count_Items','Earliest_Due','IDs']
        cols = [c for c in cols if c in sched_saved.columns]
        sched_view = sched_saved[cols].sort_values(['Date','Company'])
        st.dataframe(sched_view, hide_index=True, column_config=build_date_column_config(sched_view))
        st.download_button('⬇️ Proposed_Schedule.csv', data=sched_view.to_csv(index=False).encode('utf-8'), file_name='Proposed_Schedule.csv')
        st.download_button('⬇️ CTPM_Optimized_Schedule.ics', data=ics_from_schedule(sched_saved), file_name='CTPM_Optimized_Schedule.ics', mime='text/calendar')

    elif isinstance(sched_saved, pd.DataFrame) and sched_saved.empty:
        st.warning('No schedule could be generated. Try relaxing windows, increasing capacity, or widening windows.')

elif page == "🗺️ On‑Site Map":
    alt.themes.enable("ctpm")
    st.subheader("On‑Site Planning — ZIP bubbles (Shipping ZIP only) — scroll zoom · drag pan")

    settings = load_settings()
    lab_zip = str(settings.get("lab_zip", "45430")).zfill(5)
    eq = equip.copy()
    cmp = companies_shipping_only(companies)

    with st.popover("🔎 Filters"):
        lookahead = st.slider("Look‑ahead window (days)", min_value=7, max_value=180,
                              value=int(st.session_state.get("map_lookahead", 60)), step=1, key="map_lookahead")
        threshold = st.number_input("Non‑local threshold (miles) — ZIPs below this appear in 'Local' section only",
                                    value=int(st.session_state.get("map_threshold", 80)), step=5, key="map_threshold")
        min_count = st.number_input("Min items per ZIP (bubble)", value=1, min_value=1, step=1, key="map_min_count")
        ignore_days = st.number_input(
            "Ignore items more than (days) past due",
            value=int(st.session_state.get("map_ignore_days", int(settings.get("overdue_ignore_days", 90)))),
            min_value=0, step=5, key="map_ignore_days",
        )
        basemap = st.selectbox("Basemap style",
                               options=["carto-positron", "open-street-map", "white-bg"],
                               index=0, key="map_basemap")
        color_by = st.selectbox("Color bubbles by",
                                options=["Count (items)", "MilesFromLab (distance)"],
                                index=0, key="map_color_by")
        group_radius = st.slider("Customer grouping distance (miles)", min_value=0, max_value=100,
                                 value=int(st.session_state.get("group_radius_miles", 10.0)),
                                 step=1, key="map_group_radius")
    group_radius = int(st.session_state.get("map_group_radius", group_radius))
    st.session_state["group_radius_miles"] = float(group_radius)
    enable_map_grouping = st.toggle("Group nearby bubbles on map", value=(group_radius > 0), key="map_enable_grouping")

    # Base filters
    if "Calibration Due" in eq.columns:
        eq["Calibration Due"] = pd.to_datetime(eq["Calibration Due"], errors="coerce")

    today = pd.Timestamp.today().normalize()
    due_mask = pd.Series(True, index=eq.index)

    if "Calibration Due" in eq.columns:
        due_mask = (eq["Calibration Due"] <= today + pd.Timedelta(days=int(lookahead)))
        overdue_cutoff = today - pd.Timedelta(days=int(ignore_days))
        due_mask = due_mask & (eq["Calibration Due"].isna() | (eq["Calibration Due"] >= overdue_cutoff))

    on_site = eq[(eq.get("IN_SHOP", False) == False) & due_mask].copy()

    if "Tracking Status" in on_site.columns:
        _ts = on_site["Tracking Status"].astype(str).str.strip().str.lower()
        on_site = on_site[~_ts.eq("wip shop")].copy()

    # Map company -> Shipping ZIP
    zip_col = "ZIP" if "ZIP" in cmp.columns else None
    if zip_col is None:
        for c in cmp.columns:
            if str(c).strip().lower() in ["zip", "postal", "postal code", "zipcode", "zip5", "zip code"]:
                zip_col = c
                break

    if zip_col is None:
        st.warning("No ZIP column found on All Companies Shipping rows.")
        st.stop()

    cmp = cmp.copy()
    cmp["ZIP"] = cmp[zip_col].apply(normalize_zip)

    on_site["CompanyKey"] = (
        on_site.get("Company", pd.Series(index=on_site.index))
        .astype(str).str.strip().str.upper()
    )
    cmp["CompanyKey"] = (
        cmp.get("Company", pd.Series(index=cmp.index))
        .astype(str).str.strip().str.upper()
    )

    zip_map = (
        cmp[["CompanyKey", "ZIP"]]
        .dropna()
        .drop_duplicates(subset=["CompanyKey"])
        .set_index("CompanyKey")["ZIP"]
        .to_dict()
    )

    on_site["ZIP"] = on_site["CompanyKey"].map(zip_map)
    on_site["ZIP"] = on_site["ZIP"].apply(normalize_zip)

    diag = {
        "equip_rows": int(len(eq)),
        "on_site_rows_after_filters": int(len(on_site)),
        "unique_companies": int(on_site.get("CompanyKey", pd.Series(dtype=str)).nunique()),
        "unique_zips_after_map": int(on_site.get("ZIP", pd.Series(dtype=str)).dropna().nunique()),
    }

    on_site = on_site[on_site["ZIP"].notna()].copy()

    # Aggregate by ZIP
    agg = on_site.groupby("ZIP", as_index=False).agg(
        Count=("I.D.", "nunique") if "I.D." in on_site.columns else ("CompanyKey", "size"),
        Companies=("Company", lambda s: ", ".join(sorted(set(map(str, s.dropna()))))[:600]),
    )
    agg = agg[agg["Count"] >= int(min_count)].copy()

    # Load zip_centroids.csv
    data_folder = Path(sanitize_path(st.session_state["data_file"])).resolve().parent
    zc_path = data_folder / "zip_centroids.csv"

    if not zc_path.exists():
        st.info("zip_centroids.csv not found next to All Data.xlsx. Generate it once and reload.")
        if st.button("Generate zip_centroids.csv (one-time)", key="gen_zc_plot"):
           zdf = build_zip_centroids(companies_shipping_only(companies))
           zdf.to_csv(zc_path, index=False)
           st.success(f"zip_centroids.csv written to: {zc_path}")
        st.stop()

    try:
        zc = pd.read_csv(zc_path)
        cols = {c.lower(): c for c in zc.columns}
        zc = zc.rename(columns={
            cols.get("zip", "zip"): "ZIP",
            cols.get("lat", "lat"): "lat",
            cols.get("lon", "lon"): "lon",
        })
    except Exception as ex:
        st.error(f"Failed to read zip_centroids.csv: {ex}")
        st.stop()

    centroids = prepare_zip(zc, zip_col="ZIP", lat_col="lat", lon_col="lon")
    agg2 = attach_zip(
        agg.rename(columns={"ZIP": "zip_norm"}).rename(columns={"zip_norm": "ZIP"}),
        centroids,
        agg_zip_col="ZIP",
    )

    # Compute miles from lab
    lab = zip_to_latlon(lab_zip) or (39.7069, -84.0560)  
    lab_lat, lab_lon = lab

    with_latlon = agg2.dropna(subset=["lat", "lon"]).copy()
    if not with_latlon.empty:
        with_latlon["MilesFromLab"] = with_latlon.apply(
            lambda r: haversine_miles(lab_lat, lab_lon, float(r["lat"]), float(r["lon"])),
            axis=1,
    )	

    plot_df = (
        with_latlon[with_latlon["MilesFromLab"] >= float(threshold)].copy()
        if not with_latlon.empty else with_latlon
    )
    plot_df, ll_metrics = coerce_lat_lon(plot_df)

    # ---- KPIs ----
    _local_df = with_latlon[with_latlon["MilesFromLab"] < float(threshold)].copy() if not with_latlon.empty else pd.DataFrame()
    _map_items     = int(on_site["I.D."].nunique()) if "I.D." in on_site.columns else len(on_site)
    _map_cos       = int(on_site["Company"].nunique()) if "Company" in on_site.columns else 0
    _map_nonlocal_zips = int(plot_df["ZIP"].nunique()) if not plot_df.empty and "ZIP" in plot_df.columns else 0
    _map_local_zips    = int(_local_df["ZIP"].nunique()) if not _local_df.empty and "ZIP" in _local_df.columns else 0
    _max_miles     = float(plot_df["MilesFromLab"].max()) if not plot_df.empty and "MilesFromLab" in plot_df.columns else 0.0

    mk1, mk2, mk3, mk4, mk5 = st.columns(5)
    with mk1: st.markdown(f"<div class='kpi'><div class='label'>Items Due</div><div class='value'>{_map_items:,}</div></div>", unsafe_allow_html=True)
    with mk2: st.markdown(f"<div class='kpi'><div class='label'>Companies</div><div class='value'>{_map_cos:,}</div></div>", unsafe_allow_html=True)
    with mk3: st.markdown(f"<div class='kpi'><div class='label'>Non-Local ZIPs</div><div class='value'>{_map_nonlocal_zips}</div></div>", unsafe_allow_html=True)
    with mk4: st.markdown(f"<div class='kpi'><div class='label'>Local ZIPs (&lt;{int(threshold)} mi)</div><div class='value'>{_map_local_zips}</div></div>", unsafe_allow_html=True)
    with mk5: st.markdown(f"<div class='kpi'><div class='label'>Farthest ZIP</div><div class='value'>{_max_miles:.0f} mi</div></div>", unsafe_allow_html=True)

    if plot_df.empty:
        st.info("No non-local ZIPs to plot. All items are within the threshold — see Local section below.")
    else:
        # ---- Optional grouping ----
        map_plot_df = plot_df.copy()
        if enable_map_grouping and float(st.session_state.get("group_radius_miles", 0)) > 0:
            map_plot_df = cluster_points_greedy(plot_df, float(st.session_state.get("group_radius_miles", 0)))

        _is_grouped = enable_map_grouping and float(st.session_state.get("group_radius_miles", 0)) > 0
        _color_col = "Count" if "MilesFromLab" not in st.session_state.get("map_color_by", "") else "MilesFromLab"
        _color_scale = "Reds" if _color_col == "Count" else "RdYlGn_r"

        # ---- Map ----
        fig = px.scatter_map(
            map_plot_df,
            lat="lat", lon="lon",
            size="Count", size_max=40,
            color=_color_col if _color_col in map_plot_df.columns else "Count",
            color_continuous_scale=_color_scale,
            hover_name=("ClusterID" if _is_grouped else "ZIP"),
            hover_data=(
                {"Count": True, "Members": True, "ZIPs": True, "Companies": True, "lat": False, "lon": False}
                if _is_grouped
                else {"Count": True, "MilesFromLab": ":.0f", "Companies": True, "lat": False, "lon": False}
            ),
            zoom=3.4,
            height=640,
            center={"lat": 39.0, "lon": -98.0},
            map_style=basemap,
        )
        # Add lab location marker
        fig.add_scattermap(
            lat=[lab_lat], lon=[lab_lon],
            mode="markers+text",
            text=["Lab"],
            textposition="top right",
            marker=dict(size=14, color="#7E1F23"),
            hovertemplate=f"<b>CTPM Lab</b><br>ZIP {lab_zip}<extra></extra>",
            name="CTPM Lab",
            showlegend=False,
        )
        fig.update_layout(margin={"l": 0, "r": 0, "t": 0, "b": 0})
        st.plotly_chart(fig, use_container_width=True, config={"scrollZoom": True, "displaylogo": False})

        # ---- ZIP summary table ----
        st.markdown("### Non-local ZIP summary")
        _zip_sum = plot_df[["ZIP", "Count", "MilesFromLab", "Companies"]].sort_values("MilesFromLab").copy()
        _zip_sum["MilesFromLab"] = _zip_sum["MilesFromLab"].round(1)
        st.dataframe(_zip_sum, hide_index=True, use_container_width=True)

        # ---- Per-ZIP equipment expanders ----
        st.markdown("### Equipment due by ZIP")
        _detail_cols = [c for c in [
            "Company", "I.D.", "Description", "Manufacturer", "Model Number",
            "Calibration Due", "Tracking Status", "Status",
        ] if c in on_site.columns]

        zset = set(plot_df["ZIP"].astype(str).unique()) if "ZIP" in plot_df.columns else set()
        detail = on_site[on_site["ZIP"].astype(str).isin(zset)].copy() if zset else on_site.copy()
        if "Calibration Due" in detail.columns:
            detail["Calibration Due"] = pd.to_datetime(detail["Calibration Due"], errors="coerce")

        zip_order = plot_df.sort_values("MilesFromLab")["ZIP"].astype(str).tolist() if "ZIP" in plot_df.columns else []

        for z in zip_order[:60]:
            sub = detail[detail["ZIP"].astype(str) == str(z)].copy()
            if sub.empty:
                continue
            item_ct = int(sub["I.D."].nunique()) if "I.D." in sub.columns else len(sub)
            comp_ct = int(sub["Company"].nunique()) if "Company" in sub.columns else 0
            miles_val = None
            try:
                miles_val = float(plot_df.loc[plot_df["ZIP"].astype(str) == str(z), "MilesFromLab"].iloc[0])
            except Exception:
                pass
            earliest = None
            if "Calibration Due" in sub.columns:
                earliest = sub["Calibration Due"].dropna().min()
            hdr = (f"ZIP {z} — {item_ct:,} item(s) · {comp_ct:,} company(ies)"
                   + (f" · {miles_val:.1f} mi" if miles_val is not None else "")
                   + (f" · earliest due {earliest.strftime('%b %d, %Y')}" if pd.notna(earliest) else ""))
            with st.expander(hdr, expanded=False):
                view = sub[_detail_cols].copy() if _detail_cols else sub.copy()
                if "Calibration Due" in view.columns:
                    view = view.sort_values("Calibration Due", na_position="last")
                st.dataframe(view, hide_index=True, use_container_width=True,
                             column_config=build_date_column_config(view))

    # ---- Local companies (within threshold) ----
    if not _local_df.empty:
        _local_zips_set = set(_local_df["ZIP"].astype(str).unique()) if "ZIP" in _local_df.columns else set()
        _local_items = on_site[on_site["ZIP"].astype(str).isin(_local_zips_set)].copy() if _local_zips_set else pd.DataFrame()
        _local_item_ct = int(_local_items["I.D."].nunique()) if not _local_items.empty and "I.D." in _local_items.columns else 0
        with st.expander(f"Local companies (within {int(threshold)} mi) — {_local_item_ct:,} items, {_map_local_zips} ZIP(s)", expanded=False):
            if _local_items.empty:
                st.info("No equipment rows matched local ZIPs.")
            else:
                _lc = [c for c in ["Company","I.D.","Description","Manufacturer","Model Number","Calibration Due","Tracking Status","Status"] if c in _local_items.columns]
                if "Calibration Due" in _local_items.columns:
                    _local_items["Calibration Due"] = pd.to_datetime(_local_items["Calibration Due"], errors="coerce")
                st.dataframe(_local_items[_lc].sort_values("Calibration Due", na_position="last") if _lc else _local_items,
                             hide_index=True, use_container_width=True,
                             column_config=build_date_column_config(_local_items))

    with st.expander("Map pipeline diagnostics", expanded=False):
        st.write({**diag, "agg_zips": int(len(agg)), "agg_zips_with_centroids": int(len(with_latlon)),
                  "agg_zips_plotted_after_threshold": int(len(plot_df)), "latlon_metrics": ll_metrics})

elif page=="⌛ Work Order Aging":
    alt.themes.enable("ctpm")
    st.subheader("Work Order Aging")

    # Build/refresh the Aging dataset (from WOs + latest Tracking Status + Equipment status)
    try:
        aging_df = compute_aging_from_file(path, sig)
    except Exception as ex:
        st.error(f"Failed to compute aging: {ex}")
        aging_df = pd.DataFrame()

    if aging_df is None or aging_df.empty:
        st.info("No Work Orders available for aging.")
    else:
        # Normalize for charts only (table remains full)
        plot_df = charts_df(aging_df)

        # ---- Filters --------------------------------------------------------
        with st.expander("Filters", expanded=False):
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                status_sel = st.multiselect(
                    "Status",
                    sorted(plot_df.get("Status", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
                )
            with c2:
                substatus_sel = st.multiselect(
                    "Sub-Status",
                    sorted(plot_df.get("Sub-Status", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
                )
            with c3:
                track_sel = st.multiselect(
                    "Tracking (equip/latest)",
                    sorted(plot_df.get("Equipment Tracking Status", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
                )
            with c4:
                id_query = st.text_input("Filter I.D. (contains)")

        f = aging_df.copy()
        if status_sel:
            f = f[f.get("Status", pd.Series(index=f.index)).astype(str).isin(status_sel)]
        if substatus_sel:
            f = f[f.get("Sub-Status", pd.Series(index=f.index)).astype(str).isin(substatus_sel)]
        if track_sel:
            f = f[f.get("Equipment Tracking Status", pd.Series(index=f.index)).astype(str).isin(track_sel)]
        if id_query:
            f = f[f.get("I.D.", pd.Series(index=f.index)).astype(str).str.contains(id_query, case=False, na=False)]

        # ---- KPIs -----------------------------------------------------------
        open_days = pd.to_numeric(f.get("days_open", pd.Series(index=f.index)), errors="coerce")
        to_complete = pd.to_numeric(f.get("days_to_complete", pd.Series(index=f.index)), errors="coerce")

        _stale_mask = open_days > 365
        _stale_ct   = int(_stale_mask.sum())
        _normal_days = open_days[~_stale_mask]

        k1, k2, k3, k4, k5 = st.columns(5)
        with k1:
            st.markdown(
                f"<div class='kpi'><div class='label'>Open WOs</div>"
                f"<div class='value'>{int((f.get('Completed Date', pd.Series(index=f.index)).isna()).sum()):,}</div></div>",
                unsafe_allow_html=True
            )
        with k2:
            st.markdown(
                f"<div class='kpi'><div class='label'>Open — median age</div>"
                f"<div class='value'>{0 if _normal_days.dropna().empty else int(_normal_days.dropna().median())} d</div></div>",
                unsafe_allow_html=True
            )
        with k3:
            st.markdown(
                f"<div class='kpi'><div class='label'>Closed — median TAT</div>"
                f"<div class='value'>{0 if to_complete.dropna().empty else int(to_complete.dropna().median())} d</div></div>",
                unsafe_allow_html=True
            )
        with k4:
            st.markdown(
                f"<div class='kpi'><div class='label'>Open &gt; 10 days</div>"
                f"<div class='value'>{int((_normal_days > 10).sum())}</div></div>",
                unsafe_allow_html=True
            )
        with k5:
            st.markdown(
                f"<div class='kpi' style='border-color:#F44336'><div class='label'>Stale &gt; 365 days</div>"
                f"<div class='value'>{_stale_ct}</div></div>",
                unsafe_allow_html=True
            )

        # Stale WO warning
        if _stale_ct > 0:
            _stale_rows = f[_stale_mask].copy()
            with st.expander(f"Stale WOs (>{365} days open) — {_stale_ct} record(s) excluded from histogram", expanded=False):
                st.caption("These items have been open over a year. They are excluded from the histogram below but included in all other counts. Investigate or close them.")
                _sc = [c for c in ["Work Order","I.D.","Company","Open Date","Status","Sub-Status","days_open"] if c in _stale_rows.columns]
                st.dataframe(_stale_rows[_sc].sort_values("days_open", ascending=False), hide_index=True,
                             column_config=build_date_column_config(_stale_rows))

        # ---- Charts (fixed sizes) ------------------------------------------
        # Age histogram for OPEN WOs — capped at 365 days
        open_hist = pd.DataFrame({"days": _normal_days.dropna()})
        if not open_hist.empty:
            st.altair_chart(
                alt.Chart(open_hist)
                   .mark_bar()
                   .encode(
                        x=alt.X('days:Q', bin=alt.Bin(maxbins=25), title='Days open (≤365)',
                                scale=alt.Scale(domain=[0, 365])),
                        y=alt.Y('count():Q', title='Count'),
                        color=alt.condition(
                            alt.datum.days > 14,
                            alt.value(get_active_theme()['PRIMARY']),
                            alt.value(get_active_theme()['ACCENT'])
                        ),
                        tooltip=[alt.Tooltip('days:Q', bin=alt.Bin(maxbins=25), title='Days open'),
                                 alt.Tooltip('count():Q', title='WOs')]
                   )
                   .properties(width=1100, height=280, title='Open WO Age Distribution (stale >365d excluded)')
            )

        # Company Pareto of OPEN WOs
        open_tbl = f[f.get('Completed Date', pd.Series(index=f.index)).isna()]
        if not open_tbl.empty and 'Company' in open_tbl.columns:
            pareto = (
                open_tbl.groupby('Company')
                        .size()
                        .reset_index(name='Open')
                        .sort_values('Open', ascending=False)
            )
            pareto['Cum%'] = (pareto['Open'].cumsum() / max(1, pareto['Open'].sum()) * 100).round(1)

            left = alt.Chart(pareto).mark_bar().encode(
                y=alt.Y('Company:N', sort='-x'),
                x='Open:Q',
                tooltip=['Company','Open']
            )
            right = alt.Chart(pareto).mark_line(
                color=get_active_theme()['ACCENT']
            ).encode(
                y=alt.Y('Cum%:Q', axis=alt.Axis(format='%')),
                x=alt.X('Company:N', sort=None)
            )

            st.altair_chart(
                alt.layer(left, right).resolve_scale(y='independent').properties(width=1100, height=280)
            )

        # ---- Table + Export -------------------------------------------------
        cols = [
            "Work Order","I.D.","Company","Open Date","Due Date","Completed Date",
            "Status","Sub-Status","Equipment Tracking Status","days_open","days_to_complete"
        ]
        cols = [c for c in cols if c in f.columns]
        st.dataframe(
            f[cols].sort_values(["Completed Date","Open Date","Company","Work Order"],
                                ascending=[True,True,True,True]),
            hide_index=True
        )

        st.download_button(
            "⬇️ Aging.csv",
            data=f[cols].to_csv(index=False).encode("utf-8"),
            file_name="Aging.csv",
            mime="text/csv"
        )


elif page=="⏱️ TAT (Shop)":
    alt.themes.enable("ctpm")
    st.subheader("Turnaround Time — Shop (overview)")
    wc = wc_all.copy()

    _lb_opts = {"3 months": 3, "6 months": 6, "12 months": 12, "All time": None}
    _lb_col, _info_col = st.columns([2, 6])
    with _lb_col:
        _lb_choice = st.selectbox("Lookback window", list(_lb_opts.keys()), index=1, key="tat_lookback")
    _lb_months = _lb_opts[_lb_choice]

    # Completed cycles only (shipped) — in-progress items have no valid total TAT
    wc_done = wc[~wc['In_Shop'].fillna(False)].copy()
    if _lb_months is not None and 'Recv' in wc_done.columns:
        _cutoff = pd.Timestamp.today().normalize() - pd.DateOffset(months=_lb_months)
        _ship_est = pd.to_datetime(wc_done['Recv'], errors='coerce') + pd.to_timedelta(
            pd.to_numeric(wc_done['Days_Total_R2Ship'], errors='coerce').fillna(0), unit='D')
        wc_done = wc_done[_ship_est >= _cutoff]

    wc_plot = charts_df(wc_done)
    n_anomalies = int(pd.Series(wc.get('TAT_Anomaly', []), dtype='bool').fillna(False).sum()) if hasattr(wc, 'get') else 0
    with _info_col:
        st.caption(f"{len(wc_plot):,} completed cycles in window  •  {n_anomalies:,} anomalies (ship-before-receive) across all data")

    if wc_plot is None or wc_plot.empty:
        st.info("No completed WIP cycles in the selected window.")
    else:
        # Median stats
        def med(s): s = pd.to_numeric(s, errors='coerce').dropna(); return int(s.median()) if len(s)>0 else 0
        k1,k2,k3,k4 = st.columns(4)
        with k1: st.markdown(f"<div class='kpi'><div class='label'>R→Cal (median)</div><div class='value'>{med(wc_plot['Days_R2Cal'])}</div></div>", unsafe_allow_html=True)
        with k2: st.markdown(f"<div class='kpi'><div class='label'>Cal→QC (median)</div><div class='value'>{med(wc_plot['Days_Cal2QC'])}</div></div>", unsafe_allow_html=True)
        with k3: st.markdown(f"<div class='kpi'><div class='label'>QC→Ship (median)</div><div class='value'>{med(wc_plot['Days_QC2Ship'])}</div></div>", unsafe_allow_html=True)
        with k4: st.markdown(f"<div class='kpi'><div class='label'>R→Ship (median)</div><div class='value'>{med(wc_plot['Days_Total_R2Ship'])}</div></div>", unsafe_allow_html=True)

        # Boxplots
        m = wc_plot[['Days_R2Cal','Days_Cal2QC','Days_QC2Ship','Days_Total_R2Ship']].melt(var_name='Stage', value_name='Days').dropna()
        m['Days'] = pd.to_numeric(m['Days'], errors='coerce')
        st.altair_chart(
            alt.Chart(m).mark_boxplot().encode(
                x='Stage:N', y='Days:Q', color=alt.value(get_active_theme()['ACCENT'])
            ).properties(height=280, title='Boxplots by stage').interactive(), )



elif page=="👷 Technician Efficiency":
    alt.themes.enable("ctpm")
    st.subheader("Technician Efficiency")

    st.caption(
        "WO-based: Tech = Performed By on COMPLETE line items (All WOs export). "
        "Hours = All Events 'Total Time' (or line Actual Time fallback) per (Work Order, I.D.). "
        "This page shows weekly calibration hours output (Mon–Sun week buckets starting Monday)."
    )

    if build_efficiency_datasets is None:
        st.error("Missing wo_efficiency.py. Add wo_efficiency.py next to this app file to enable Technician Efficiency (WO-based).")
        st.stop()

    all_wos_path, names_path = locate_efficiency_files(st.session_state['data_file'])
    if not all_wos_path.exists():
        st.error(f"All WOs export not found next to All Data.xlsx. Expected: {all_wos_path}")
        st.stop()
    if not names_path.exists():
        st.error(f"Names_Numbers.xlsx not found next to All Data.xlsx. Expected: {names_path}")
        st.stop()

    datasets = build_efficiency_datasets_from_file(path, sig, str(all_wos_path), str(names_path), skip_threshold=0.30)
    df = datasets.get('events_merged_kpis', pd.DataFrame()).copy()

    if df is None or df.empty:
        st.info("No merged line-level data available (events + work-completed join by Work Order + I.D.).")
        st.stop()

    if 'PerformedBy_Name' not in df.columns:
        st.error("PerformedBy_Name not available. Verify All WOs export is present and Names_Numbers.xlsx is correct.")
        st.stop()

    # Datetimes (Completed Date basis)
    df['CompletedDateTime'] = pd.to_datetime(df.get('CompletedDateTime'), errors='coerce')

    # Keep only completed line items from WO sub-table status
    status_col = 'Status_wo' if 'Status_wo' in df.columns else ('Status' if 'Status' in df.columns else None)
    if status_col is not None:
        df = df[df[status_col].astype(str).str.strip().eq('Complete')].copy()

    # Optional: Field only
    field_only = st.toggle("Field Cal only", value=False, key='te_field_only')
    if field_only and 'Related Event Type' in df.columns:
        df = df[df['Related Event Type'].astype(str).str.strip().eq('FIELD CALIBRATION')].copy()

    # Effort hours
    df['EffortHours'] = pd.to_numeric(df.get('EffortHours', df.get('Total Time', pd.NA)), errors='coerce')
    df['EffortHours'] = df['EffortHours'].fillna(pd.to_numeric(df.get('Actual_Time', pd.NA), errors='coerce'))

    # Technician label
    df['Technician'] = df['PerformedBy_Name'].astype(str).str.strip()
    df.loc[df['Technician'].isin(['', 'nan', 'None', 'NULL', 'null']), 'Technician'] = '(Unknown)'

    # ---------- Filters ----------
    with st.expander("Filters", expanded=False):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            tech_options = sorted(df['Technician'].dropna().astype(str).unique().tolist())
            te_tech_sel = st.multiselect("Technician(s)", options=tech_options, default=[], key='te_hours_tech_sel')
        with c2:
            comp_options = sorted(df.get('Company', pd.Series(dtype='string')).dropna().astype(str).str.strip().replace('', np.nan).dropna().unique().tolist())
            te_company_sel = st.multiselect("Company", options=comp_options, default=[], key='te_hours_company_sel')
        with c3:
            dts = df['CompletedDateTime'].dropna()
            if not dts.empty:
                max_d = dts.max().date()
                min_d = dts.min().date()
                default_start = max(min_d, (pd.Timestamp(max_d) - pd.Timedelta(days=84)).date())
                te_date_range = st.date_input('Completed date range', value=(default_start, max_d), key='te_hours_date_range')
            else:
                te_date_range = None
        with c4:
            te_wo_query = st.text_input('Work Order search (contains)', value='', key='te_hours_wo_search')

        st.toggle('Show 40-hr/week target line', value=bool(st.session_state.get('te_show_40hr', True)), key='te_show_40hr')

        if st.button('🧹 Clear filters', key='te_hours_clear'):
            for k in ('te_hours_tech_sel','te_hours_company_sel','te_hours_date_range','te_hours_wo_search','te_show_40hr'):
                st.session_state.pop(k, None)
            st.rerun()

    # Apply filters
    if te_tech_sel:
        df = df[df['Technician'].astype(str).isin([str(x) for x in te_tech_sel])].copy()
    if te_company_sel and 'Company' in df.columns:
        df = df[df['Company'].astype(str).isin([str(x) for x in te_company_sel])].copy()
    if isinstance(te_date_range, (list, tuple)) and len(te_date_range) == 2 and all(te_date_range):
        sdt = pd.to_datetime(te_date_range[0])
        edt = pd.to_datetime(te_date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        df = df[pd.to_datetime(df['CompletedDateTime'], errors='coerce').between(sdt, edt)].copy()
    if te_wo_query:
        df = df[df.get('Work Order', pd.Series(index=df.index, dtype='string')).astype(str).str.contains(str(te_wo_query), case=False, na=False)].copy()

    df = df.dropna(subset=['CompletedDateTime', 'EffortHours']).copy()
    if df.empty:
        st.info("No records match the selected filters.")
        st.stop()

    # Weekly aggregation (week starts Monday) based on CompletedDateTime
    df['WeekStart'] = df['CompletedDateTime'].dt.to_period('W-MON').dt.start_time
    weekly = (
        df.groupby(['Technician', 'WeekStart'], as_index=False)
          .agg(
              Total_Hours=('EffortHours', 'sum'),
              Completed_Items=('I.D.', 'nunique') if 'I.D.' in df.columns else ('Work Order', 'size'),
              WOs_Touched=('Work Order', 'nunique') if 'Work Order' in df.columns else ('Technician', 'size'),
          )
    )

    week_now = pd.Timestamp.today().to_period('W-MON').start_time
    last_week = week_now - pd.Timedelta(days=7)

    total_hours = float(pd.to_numeric(df['EffortHours'], errors='coerce').fillna(0).sum())
    hours_this_week = float(weekly.loc[weekly['WeekStart'] == week_now, 'Total_Hours'].sum())
    hours_last_week = float(weekly.loc[weekly['WeekStart'] == last_week, 'Total_Hours'].sum())

    pertech_avg = (
        weekly.groupby('Technician', as_index=False)['Total_Hours']
              .mean()
              .rename(columns={'Total_Hours': 'Avg_Hours_Per_Week'})
    )

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        brand_metric('Total Hours (filtered)', round(total_hours, 1))
    with k2:
        brand_metric('Hours This Week', round(hours_this_week, 1))
    with k3:
        brand_metric('Hours Last Week', round(hours_last_week, 1))
    with k4:
        avg_selected = float(pertech_avg['Avg_Hours_Per_Week'].mean()) if not pertech_avg.empty else 0.0
        brand_metric('Avg Hours/Tech/Week', f"{avg_selected:.1f}")

    tab1, tab2, tab3 = st.tabs(["Weekly Hours (Table)", "Weekly Hours Trend", "Line Item Drilldown"])

    with tab1:
        st.markdown("<div class='section-title'>Weekly Hours by Technician</div>", unsafe_allow_html=True)
        view = weekly.sort_values(['WeekStart','Total_Hours'], ascending=[False, False]).copy()
        st.dataframe(view, hide_index=True, column_config=build_date_column_config(view))
        download_buttons(view, base_name='tech_weekly_hours', key_prefix='te_weekly')

    with tab2:
        st.markdown("<div class='section-title'>Weekly Hours Trend</div>", unsafe_allow_html=True)

        techs_all = sorted(weekly['Technician'].dropna().astype(str).unique().tolist())
        default_pick = techs_all[:6] if len(techs_all) > 6 else techs_all
        chart_techs = st.multiselect('Show techs on chart', options=techs_all, default=default_pick, key='te_chart_techs')
        wplot = weekly[weekly['Technician'].astype(str).isin([str(x) for x in chart_techs])].copy() if chart_techs else weekly.copy()

        line = alt.Chart(wplot).mark_line(point=True).encode(
            x=alt.X('WeekStart:T', title='Week (Mon start)', axis=alt.Axis(format='%b %Y')),
            y=alt.Y('Total_Hours:Q', title='Total Hours', scale=alt.Scale(zero=True)),
            color=alt.Color('Technician:N', title='Technician'),
            tooltip=[
                alt.Tooltip('WeekStart:T', title='Week'),
                alt.Tooltip('Technician:N'),
                alt.Tooltip('Total_Hours:Q', title='Hours', format=',.1f'),
                alt.Tooltip('Completed_Items:Q', title='Items'),
                alt.Tooltip('WOs_Touched:Q', title='WOs'),
            ],
        )

        if bool(st.session_state.get('te_show_40hr', True)):
            rule = alt.Chart(pd.DataFrame({'y':[40.0]})).mark_rule(strokeDash=[6,4], color=get_active_theme()['ACCENT']).encode(y='y:Q')
            st.altair_chart((line + rule).properties(height=320).interactive(), use_container_width=True)
        else:
            st.altair_chart(line.properties(height=320).interactive(), use_container_width=True)

    with tab3:
        st.markdown("<div class='section-title'>Line Item Drilldown (Filtered)</div>", unsafe_allow_html=True)
        drill_cols = [c for c in ['Technician','Work Order','I.D.','Company','Description','CompletedDateTime','EffortHours'] if c in df.columns]
        dview = df[drill_cols].sort_values('CompletedDateTime', ascending=False)
        st.dataframe(dview, hide_index=True, column_config=build_date_column_config(dview))
        download_buttons(dview, base_name='technician_efficiency_line_items', key_prefix='te_drill')
elif page=="📈 WO Efficiency":
    alt.themes.enable("ctpm")
    st.subheader("Work Order Efficiency")
    st.caption("WO outcomes come from All WOs export sub-table line items (Work Order + I.D.). Tech list uses COMPLETE lines only. Date filters default to YTD and do not rely on Receiving dates (field work often has none).")

    if build_efficiency_datasets is None:
        st.error("Missing wo_efficiency.py. Add wo_efficiency.py next to this app file to enable WO Efficiency.")
        st.stop()

    all_wos_path, names_path = locate_efficiency_files(st.session_state['data_file'])
    if not all_wos_path.exists():
        st.error(f"All WOs export not found next to All Data.xlsx. Expected: {all_wos_path}")
        st.stop()
    if not names_path.exists():
        st.error(f"Names_Numbers.xlsx not found next to All Data.xlsx. Expected: {names_path}")
        st.stop()

    datasets = build_efficiency_datasets_from_file(path, sig, str(all_wos_path), str(names_path), skip_threshold=0.30)

    view_field_only = st.toggle("Field Cal only", value=False, help="Filter to Related Event Type = FIELD CALIBRATION")

    wo_rollup = datasets.get('wo_rollup_field', pd.DataFrame()) if view_field_only else datasets.get('wo_rollup', pd.DataFrame())
    tech_summary = datasets.get('tech_summary_field', pd.DataFrame()) if view_field_only else datasets.get('tech_summary', pd.DataFrame())
    tech_drilldown = datasets.get('tech_drilldown_field', pd.DataFrame()) if view_field_only else datasets.get('tech_drilldown', pd.DataFrame())

    # Line-level (one row per item line)
    wo_lines = datasets.get('wo_lines_field', pd.DataFrame()) if view_field_only else datasets.get('wo_lines', pd.DataFrame())

    if wo_rollup is None or wo_rollup.empty:
        st.info("No WO line items parsed from All WOs export.")
        st.stop()

    # Enrich line-level with computed EffectiveDateTime (module helper); safe fallback if helper missing
    if add_effective_date_columns is not None and isinstance(wo_lines, pd.DataFrame) and (not wo_lines.empty):
        try:
            lines_enriched = add_effective_date_columns(wo_lines)
        except Exception:
            lines_enriched = wo_lines.copy()
    else:
        lines_enriched = wo_lines.copy() if isinstance(wo_lines, pd.DataFrame) else pd.DataFrame()

    # Normalize common column names for downstream use
    if 'Equipment_ID' in lines_enriched.columns and 'I.D.' not in lines_enriched.columns:
        lines_enriched = lines_enriched.rename(columns={'Equipment_ID':'I.D.'})

    # ---------------- Filters (Technician / Company / Status / Date range / WO search) ----------------
    today = pd.Timestamp.today().normalize()
    ytd_start = pd.Timestamp(year=int(today.year), month=1, day=1)

    def _split_techs(s: str) -> list[str]:
        s = str(s or '')
        if not s.strip():
            return []
        # Tech list is stored as newline-delimited (legacy exports may include multiple whitespace)
        parts = re.split(r"[\n\r\t]+", s)
        out = []
        for p in parts:
            p = str(p).strip()
            if p:
                out.append(p)
        return out

    # Build tech options from WO_Techs_CompleteOnly
    tech_opts = []
    if 'WO_Techs_CompleteOnly' in wo_rollup.columns:
        for v in wo_rollup['WO_Techs_CompleteOnly'].dropna().astype(str).tolist():
            tech_opts.extend(_split_techs(v))
    tech_opts = sorted(set([t for t in tech_opts if t]))

    comp_opts = sorted(wo_rollup.get('Company', pd.Series(dtype='string')).dropna().astype(str).str.strip().replace('', np.nan).dropna().unique().tolist())

    status_opts = ['Complete','Skipped','Backordered','Incomplete']

    with st.expander('Filters', expanded=False):
        c1, c2, c3, c4 = st.columns([2,2,2,2])
        with c1:
            wo_tech_sel = st.multiselect('Technician (WO tech list)', options=tech_opts, default=[], key='woeff_tech_sel')
        with c2:
            wo_company_sel = st.multiselect('Company', options=comp_opts, default=[], key='woeff_company_sel')
        with c3:
            wo_status_sel = st.multiselect('Status (line items)', options=status_opts, default=status_opts, key='woeff_status_sel')
        with c4:
            wo_query = st.text_input('Work Order search (contains)', value='', key='woeff_wo_query')

        # Date range (YTD default)
        date_val = st.session_state.get('woeff_date_range')
        if not date_val:
            default_range = (ytd_start.date(), today.date())
        else:
            default_range = date_val
        wo_date_range = st.date_input('Date range (YTD default)', value=default_range, key='woeff_date_range')

        only_wos_with_status = st.toggle('Only show WOs that contain selected statuses', value=False, key='woeff_only_wos_with_status')

    # ---------------- Apply filters ----------------
    roll = wo_rollup.copy()

    # Company filter
    if wo_company_sel and 'Company' in roll.columns:
        roll = roll[roll['Company'].astype(str).isin([str(x) for x in wo_company_sel])].copy()

    # Technician filter (WO_Techs_CompleteOnly contains selected tech)
    if wo_tech_sel and 'WO_Techs_CompleteOnly' in roll.columns:
        pat = '|'.join([re.escape(str(t)) for t in wo_tech_sel])
        roll = roll[roll['WO_Techs_CompleteOnly'].astype(str).str.contains(pat, case=False, na=False)].copy()

    # Work Order search
    if wo_query:
        roll = roll[roll.get('Work Order', pd.Series(index=roll.index, dtype='string')).astype(str).str.contains(str(wo_query), case=False, na=False)].copy()

    # Date filter: use WO completed date if present; else use max EffectiveDateTime from line-level
    wo_dates = None
    if isinstance(lines_enriched, pd.DataFrame) and not lines_enriched.empty and 'EffectiveDateTime' in lines_enriched.columns:
        tmp = lines_enriched.copy()
        tmp['EffectiveDateTime'] = pd.to_datetime(tmp['EffectiveDateTime'], errors='coerce')
        wo_dates = tmp.dropna(subset=['Work Order','EffectiveDateTime']).groupby('Work Order', as_index=False)['EffectiveDateTime'].max().rename(columns={'EffectiveDateTime':'WO_AnalysisDateTime'})

    if 'WO_CompletedDateTime' in roll.columns:
        roll['WO_CompletedDateTime'] = pd.to_datetime(roll['WO_CompletedDateTime'], errors='coerce')

    if wo_dates is not None and 'Work Order' in roll.columns:
        roll = roll.merge(wo_dates, on='Work Order', how='left')
    else:
        roll['WO_AnalysisDateTime'] = pd.NaT

    roll['WO_DateForFilter'] = roll.get('WO_CompletedDateTime', pd.Series([pd.NaT]*len(roll))).combine_first(roll.get('WO_AnalysisDateTime', pd.Series([pd.NaT]*len(roll))))

    if isinstance(wo_date_range, (list, tuple)) and len(wo_date_range) == 2 and all(wo_date_range):
        sdt = pd.to_datetime(wo_date_range[0])
        edt = pd.to_datetime(wo_date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        roll = roll[pd.to_datetime(roll['WO_DateForFilter'], errors='coerce').between(sdt, edt)].copy()

    # Apply status filter to line-level for drilldowns and optionally to WO list
    lines_f = lines_enriched.copy() if isinstance(lines_enriched, pd.DataFrame) else pd.DataFrame()
    if not lines_f.empty:
        if 'Status' in lines_f.columns:
            lines_f['Status'] = lines_f['Status'].astype(str).str.strip().str.title()
        if wo_status_sel and 'Status' in lines_f.columns:
            keep = [str(s).title() for s in wo_status_sel]
            lines_status = lines_f[lines_f['Status'].isin(keep)].copy()
        else:
            lines_status = lines_f.copy()

        if 'Work Order' in roll.columns and not roll.empty and 'Work Order' in lines_status.columns:
            lines_status = lines_status[lines_status['Work Order'].astype(str).isin(roll['Work Order'].astype(str).unique().tolist())].copy()

        if only_wos_with_status and not lines_status.empty and 'Work Order' in roll.columns:
            wset = set(lines_status['Work Order'].astype(str).unique().tolist())
            roll = roll[roll['Work Order'].astype(str).isin(wset)].copy()
    else:
        lines_status = pd.DataFrame()

    if roll.empty:
        st.info('No work orders match the selected filters.')
        st.stop()

    # ---------------- KPIs ----------------
    total_wos = int(roll['Work Order'].nunique()) if 'Work Order' in roll.columns else int(len(roll))
    total_items = int(pd.to_numeric(roll.get('Total_Items', 0), errors='coerce').fillna(0).sum())
    avg_skip = float(pd.to_numeric(roll.get('WO_Skipped_%', 0), errors='coerce').fillna(0).mean())

    back_cnt = 0
    if isinstance(lines_f, pd.DataFrame) and not lines_f.empty and 'Status' in lines_f.columns:
        btmp = lines_f.copy()
        if 'Work Order' in roll.columns:
            btmp = btmp[btmp['Work Order'].astype(str).isin(roll['Work Order'].astype(str).unique().tolist())].copy()
        back_cnt = int((btmp['Status'].astype(str).str.strip().str.lower() == 'backordered').sum())

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        brand_metric('Work Orders', total_wos)
    with k2:
        brand_metric('Total Items', total_items)
    with k3:
        brand_metric('Avg Skipped %', f"{avg_skip*100:.1f}%")
    with k4:
        brand_metric('Backordered Items', back_cnt, delta='ALERT' if back_cnt > 0 else None, delta_good=(False if back_cnt > 0 else None))

    if back_cnt > 0:
        st.error('Backordered items detected. Use the Status filter to isolate Backordered and investigate.')
        with st.expander('Backordered items (drilldown)', expanded=True):
            if not lines_f.empty:
                b = lines_f[lines_f['Status'].astype(str).str.strip().str.lower() == 'backordered'].copy()
                if 'Work Order' in roll.columns:
                    b = b[b['Work Order'].astype(str).isin(roll['Work Order'].astype(str).unique().tolist())].copy()
                show_cols = [c for c in ['Work Order','Company','I.D.','Description','Related Event Type','Status','CompletedDateTime','PerformedBy_Name'] if c in b.columns]
                if not b.empty and show_cols:
                    st.dataframe(b[show_cols].sort_values(['Company','Work Order'], na_position='last'), hide_index=True, column_config=build_date_column_config(b))
                    download_buttons(b[show_cols], base_name='backordered_items', key_prefix='woeff_back')
                else:
                    st.caption('No rows to display.')
            else:
                st.caption('No line-level data available for drilldown.')

    tab1, tab2, tab3 = st.tabs(["WO Outcomes", "Tech Drilldown", "Customer Trend (Monthly)"])

    with tab1:
        st.markdown("<div class='section-title'>Work Order Outcomes</div>", unsafe_allow_html=True)

        cols = [c for c in [
            'Company','Work Order','WO_CompletedDateTime','WO_DueDate','WO_DateForFilter',
            'Total_Items','Completed_Items','Skipped_Items','Backordered_Items','Incomplete_Items',
            'WO_Complete_%','WO_Skipped_%','WO_Backordered_%','WO_Incomplete_%','WO_Closed_%','WO_Techs_CompleteOnly'
        ] if c in roll.columns]

        view = roll[cols].copy() if cols else roll.copy()
        if 'WO_Skipped_%' in view.columns:
            sort_cols = ['WO_Skipped_%'] + (['Backordered_Items'] if 'Backordered_Items' in view.columns else [])
            view = view.sort_values(sort_cols, ascending=[False] * len(sort_cols), na_position='last')

        st.dataframe(view, hide_index=True, column_config=build_date_column_config(view))
        download_buttons(view, base_name='wo_efficiency_rollup', key_prefix='wo_eff_roll')

        st.markdown("<div class='section-title'>Outcome Mix (Monthly)</div>", unsafe_allow_html=True)

        if isinstance(lines_f, pd.DataFrame) and not lines_f.empty and 'EffectiveDateTime' in lines_f.columns and 'Status' in lines_f.columns:
            gdf = lines_f.copy()
            if 'Work Order' in roll.columns:
                gdf = gdf[gdf['Work Order'].astype(str).isin(roll['Work Order'].astype(str).unique().tolist())].copy()
            gdf['EffectiveDateTime'] = pd.to_datetime(gdf['EffectiveDateTime'], errors='coerce')
            gdf = gdf.dropna(subset=['EffectiveDateTime']).copy()
            gdf['Month'] = gdf['EffectiveDateTime'].dt.to_period('M').dt.start_time
            gdf['Status'] = gdf['Status'].astype(str).str.strip().str.title()
            gdf = gdf[gdf['Status'].isin(['Complete','Skipped','Backordered','Incomplete'])].copy()

            monthly = gdf.groupby(['Month','Status'], as_index=False).size().rename(columns={'size':'Count'})
            totals = monthly.groupby('Month', as_index=False)['Count'].sum().rename(columns={'Count':'Total'})
            monthly = monthly.merge(totals, on='Month', how='left')
            monthly['Pct'] = (monthly['Count'] / monthly['Total']).where(monthly['Total'] > 0)

            pivot = monthly.pivot_table(index='Month', columns='Status', values='Count', aggfunc='sum', fill_value=0).reset_index()
            if set(['Backordered','Complete','Incomplete','Skipped']).issubset(pivot.columns):
                pivot['Total'] = pivot[['Backordered','Complete','Incomplete','Skipped']].sum(axis=1)
            else:
                pivot['Total'] = pivot.drop(columns=['Month']).sum(axis=1)
            for stn in ['Complete','Skipped','Backordered','Incomplete']:
                if stn in pivot.columns:
                    pivot[f"{stn}%"] = (pivot[stn] / pivot['Total']).where(pivot['Total'] > 0)
            st.dataframe(pivot.sort_values('Month'), hide_index=True, column_config=build_date_column_config(pivot))

            colors = {
                'Complete': '#2C7FB8',
                'Skipped': '#F0027F',
                'Backordered': '#FDB462',
                'Incomplete': '#B3B3B3',
            }

            stacked = alt.Chart(monthly).mark_bar().encode(
                x=alt.X('Month:T', title='Month'),
                y=alt.Y('Pct:Q', title='Percent', axis=alt.Axis(format='%'), stack='normalize'),
                color=alt.Color('Status:N', scale=alt.Scale(domain=list(colors.keys()), range=list(colors.values())), title='Status'),
                tooltip=[
                    alt.Tooltip('Month:T', title='Month'),
                    alt.Tooltip('Status:N'),
                    alt.Tooltip('Pct:Q', format='.1%'),
                    alt.Tooltip('Count:Q', title='Items'),
                ]
            ).properties(height=320)

            denom = alt.Chart(totals).mark_bar(opacity=0.18).encode(
                x=alt.X('Month:T', title='Month'),
                y=alt.Y('Total:Q', title='Total Items'),
                tooltip=[alt.Tooltip('Month:T', title='Month'), alt.Tooltip('Total:Q', title='Total Items')]
            ).properties(height=140)

            st.altair_chart(alt.vconcat(stacked, denom).resolve_scale(x='shared'), use_container_width=True)
        else:
            st.caption('Outcome graphs require line-level EffectiveDateTime (provided by wo_efficiency add_effective_date_columns).')

    with tab2:
        st.markdown("<div class='section-title'>Technician Association (High Skip WOs)</div>", unsafe_allow_html=True)
        if streamlit_tech_drilldown_panel is None:
            st.warning('Streamlit drilldown helpers not available.')
        else:
            streamlit_tech_drilldown_panel(tech_summary, tech_drilldown, threshold=0.30, top_n=15)

    with tab3:
        st.markdown("<div class='section-title'>Customer Trend</div>", unsafe_allow_html=True)
        if compute_skip_trend_per_customer is None or streamlit_customer_skip_trend_panel is None:
            st.warning('Trend helpers not available.')
        else:
            receiving_by_wo = datasets.get('receiving_by_wo', pd.DataFrame())
            try:
                trend_m = compute_skip_trend_per_customer(
                    roll,
                    receiving_by_wo,
                    freq='M',
                    wo_lines_df=lines_enriched if isinstance(lines_enriched, pd.DataFrame) else None,
                )
            except Exception:
                trend_m = pd.DataFrame()

            if trend_m is None or trend_m.empty:
                st.info('No customer trend data available for the selected filters.')
            else:
                streamlit_customer_skip_trend_panel(trend_m)

elif page == "🔐 Admin — Users":
    require_auth(["admin"])  # gate
    st.header("User Management")
    conn = _auth_conn(AUTH_DB_PATH)
    tabs = st.tabs(["Users","Add","Password","Role","Activate/Deactivate","Delete"])
    with tabs[0]:
        users = conn.execute("SELECT username, role, active, created_at, updated_at, last_login_at FROM users ORDER BY username COLLATE NOCASE").fetchall()
        st.dataframe(pd.DataFrame(users, columns=["username","role","active","created_at","updated_at","last_login_at"]), )
    with tabs[1]:
        u = st.text_input("New username", key="add_u"); p = st.text_input("New password", type="password", key="add_p"); r = st.selectbox("Role", ["admin","internal","customer"], index=1)
        if st.button("Create User"):
            if not u or not p: st.error("Username and password are required.")
            else:
                rec = _hash_pw(p); now = _dt.now(_tz.utc).isoformat(timespec="seconds").replace('+00:00','Z')
                try:
                    conn.execute("INSERT INTO users(username,role,password_hash,password_salt,password_iters,active,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?);", (u,r,rec['hash'],rec['salt'],rec['iters'],1,now,now)); conn.commit(); st.success(f"User '{u}' created."); st.rerun()
                except sqlite3.IntegrityError: st.error("That username already exists.")
    with tabs[2]:
        users = [r[0] for r in conn.execute("SELECT username FROM users ORDER BY username COLLATE NOCASE").fetchall()]
        if users:
            sel = st.selectbox("User", users, key="pw_user"); npw = st.text_input("New password", type="password", key="pw_new")
            if st.button("Set New Password"):
                if npw:
                    rec = _hash_pw(npw); now = _dt.now(_tz.utc).isoformat(timespec="seconds").replace('+00:00','Z')
                    conn.execute("UPDATE users SET password_hash=?, password_salt=?, password_iters=?, updated_at=? WHERE username=?;", (rec['hash'],rec['salt'],rec['iters'],now,sel)); conn.commit(); st.success(f"Password updated for '{sel}'.")
                else: st.error("Password cannot be empty.")
        else: st.info("No users.")
    with tabs[3]:
        users = [r[0] for r in conn.execute("SELECT username FROM users ORDER BY username COLLATE NOCASE").fetchall()]
        if users:
            sel = st.selectbox("User", users, key="role_user"); new_role = st.selectbox("New Role", ["admin","internal","customer"], index=1)
            if st.button("Update Role"):
                now = _dt.now(_tz.utc).isoformat(timespec="seconds").replace('+00:00','Z'); conn.execute("UPDATE users SET role=?, updated_at=? WHERE username=?;", (new_role, now, sel)); conn.commit(); st.success(f"Role for '{sel}' is now '{new_role}'.")
        else: st.info("No users.")
    with tabs[4]:
        recs = conn.execute("SELECT username, active FROM users ORDER BY username COLLATE NOCASE").fetchall()
        if recs:
            sel = st.selectbox("User", [r[0] for r in recs], key="act_user"); current = dict(recs).get(sel, 1); label = "Deactivate" if current else "Reactivate"
            if st.button(label):
                now = _dt.now(_tz.utc).isoformat(timespec="seconds").replace('+00:00','Z'); conn.execute("UPDATE users SET active=?, updated_at=? WHERE username=?;", (0 if current else 1, now, sel)); conn.commit(); st.success(f"User '{sel}' {'deactivated' if current else 'reactivated'}. "); st.rerun()
        else: st.info("No users.")
    with tabs[5]:
        users = [r[0] for r in conn.execute("SELECT username FROM users ORDER BY username COLLATE NOCASE").fetchall()]
        if users:
            sel = st.selectbox("User", users, key="del_user")
            if st.button("Delete", type="secondary"):
                if st.session_state.get("auth",{}).get("username","_").lower() == sel.lower(): st.error("You cannot delete your own account while logged in.")
                else:
                    conn.execute("DELETE FROM users WHERE username=?;", (sel,)); conn.commit(); st.success(f"User '{sel}' deleted."); st.rerun()
        else: st.info("No users.")
    # Keep users.db backed up in S3 on every admin page visit (file is tiny)
    _s3_upload(CACHE_DIR / "users.db", "userdata/users.db")

# ---- Weekly Report ----
elif page == "📋 Weekly Report":
    alt.themes.enable("ctpm")
    import sys as _sys
    import tempfile
    import os as _os

    _WEEKLY_REPORT_DIR = str(Path(__file__).parent / "CTPM-Weekly-Report")
    if _WEEKLY_REPORT_DIR not in _sys.path:
        _sys.path.insert(0, _WEEKLY_REPORT_DIR)

    try:
        from utils.parse_estimates import parse_estimates as _wr_parse_estimates
        from utils.parse_wip import parse_wip as _wr_parse_wip
        from utils.parse_historic import build_historic_lookup as _wr_build_historic_lookup
        from utils.parse_cal_events import parse_cal_events as _wr_parse_cal_events
        from utils.parse_events import parse_events_tat as _wr_parse_events_tat
        from utils.build_docs import build_sales_flash as _wr_build_sales_flash, build_weekly_update as _wr_build_weekly_update
        _wr_available = True
    except ImportError as _wr_import_err:
        _wr_available = False
        st.error(f"Weekly Report utils not found. Make sure CTPM-Weekly-Report/ is in the same folder as this app. ({_wr_import_err})")

    if _wr_available:
        tab_tat, tab_gen = st.tabs(["📊 TAT Live View", "📄 Report Generator"])

        # ── Tab 1: TAT live view ──────────────────────────────────────────────
        with tab_tat:
            st.subheader("Turnaround Time — Shop (overview)")
            _wc = wc_all.copy()

            _wlb_opts = {"3 months": 3, "6 months": 6, "12 months": 12, "All time": None}
            _wlb_col, _winfo_col = st.columns([2, 6])
            with _wlb_col:
                _wlb_choice = st.selectbox("Lookback window", list(_wlb_opts.keys()), index=1, key="wr_tat_lookback")
            _wlb_months = _wlb_opts[_wlb_choice]

            _wc_done = _wc[~_wc['In_Shop'].fillna(False)].copy()
            if _wlb_months is not None and 'Recv' in _wc_done.columns:
                _wc_cutoff = pd.Timestamp.today().normalize() - pd.DateOffset(months=_wlb_months)
                _wc_ship_est = pd.to_datetime(_wc_done['Recv'], errors='coerce') + pd.to_timedelta(
                    pd.to_numeric(_wc_done['Days_Total_R2Ship'], errors='coerce').fillna(0), unit='D')
                _wc_done = _wc_done[_wc_ship_est >= _wc_cutoff]

            _wc_plot = charts_df(_wc_done)
            _wc_anomalies = int(pd.Series(_wc.get('TAT_Anomaly', []), dtype='bool').fillna(False).sum()) if hasattr(_wc, 'get') else 0
            with _winfo_col:
                st.caption(f"{len(_wc_plot):,} completed cycles in window  •  {_wc_anomalies:,} anomalies (ship-before-receive) across all data")

            if _wc_plot is None or _wc_plot.empty:
                st.info("No completed WIP cycles in the selected window.")
            else:
                def _wr_med(s):
                    s = pd.to_numeric(s, errors='coerce').dropna()
                    return int(s.median()) if len(s) > 0 else 0

                _k1, _k2, _k3, _k4 = st.columns(4)
                with _k1: st.markdown(f"<div class='kpi'><div class='label'>R→Cal (median)</div><div class='value'>{_wr_med(_wc_plot['Days_R2Cal'])}</div></div>", unsafe_allow_html=True)
                with _k2: st.markdown(f"<div class='kpi'><div class='label'>Cal→QC (median)</div><div class='value'>{_wr_med(_wc_plot['Days_Cal2QC'])}</div></div>", unsafe_allow_html=True)
                with _k3: st.markdown(f"<div class='kpi'><div class='label'>QC→Ship (median)</div><div class='value'>{_wr_med(_wc_plot['Days_QC2Ship'])}</div></div>", unsafe_allow_html=True)
                with _k4: st.markdown(f"<div class='kpi'><div class='label'>R→Ship (median)</div><div class='value'>{_wr_med(_wc_plot['Days_Total_R2Ship'])}</div></div>", unsafe_allow_html=True)

                _m = _wc_plot[['Days_R2Cal','Days_Cal2QC','Days_QC2Ship','Days_Total_R2Ship']].melt(var_name='Stage', value_name='Days').dropna()
                _m['Days'] = pd.to_numeric(_m['Days'], errors='coerce')
                st.altair_chart(
                    alt.Chart(_m).mark_boxplot().encode(
                        x='Stage:N', y='Days:Q', color=alt.value(get_active_theme()['ACCENT'])
                    ).properties(height=280, title='Boxplots by stage').interactive(),
                )

        # ── Tab 2: Report Generator ───────────────────────────────────────────
        with tab_gen:
            st.subheader("Weekly Report Generator")
            st.caption("Upload data files and fill in the form to generate the Sales Flash and Weekly Update Word documents.")

            # File uploads
            with st.expander("📂 Data Files", expanded=True):
                _fc1, _fc2 = st.columns(2)
                with _fc1:
                    _wr_field_est   = st.file_uploader("Field Estimates *(required)*",                    type=["xlsx","xls"], key="wr_field_est")
                    _wr_wip         = st.file_uploader("Current WIP *(required for WIP pricing)*",        type=["xls","xlsx"], key="wr_wip")
                    _wr_cal_events  = st.file_uploader("Shop Cal Events *(required for cal projection)*", type=["xlsx","xls"], key="wr_cal_events")
                with _fc2:
                    _wr_shop_est    = st.file_uploader("Shop Estimates *(optional)*",                     type=["xlsx","xls"], key="wr_shop_est")
                    _wr_historic    = st.file_uploader("Historic Calibrations *(required for WIP pricing)*", type=["xlsx","xls"], key="wr_historic")
                    _wr_events      = st.file_uploader("Events *(required for TAT)*",                     type=["xlsx","xls"], key="wr_events")

            # Report period
            with st.expander("📅 Report Period", expanded=True):
                _rp1, _rp2 = st.columns(2)
                with _rp1:
                    _wr_week_start = st.text_input("Week Start Date", placeholder="e.g. May 5, 2026", key="wr_week_start")
                with _rp2:
                    _wr_week_end   = st.text_input("Week End Date",   placeholder="e.g. May 9, 2026", key="wr_week_end")

            # TAT override inputs
            with st.expander("⏱️ Turnaround Time (TAT)", expanded=True):
                st.caption("Leave at 0 to use values calculated from the Events file.")
                _tc1, _tc2, _tc3 = st.columns(3)
                with _tc1: _wr_tat_r2c = st.number_input("Receipt → Calibration (days)", min_value=0.0, step=0.1, value=0.0, key="wr_tat_r2c")
                with _tc2: _wr_tat_c2q = st.number_input("Calibration → QC (days)",      min_value=0.0, step=0.1, value=0.0, key="wr_tat_c2q")
                with _tc3: _wr_tat_q2s = st.number_input("QC → Shipment (days)",          min_value=0.0, step=0.1, value=0.0, key="wr_tat_q2s")

            # Dynamic list editor helper
            def _wr_list_editor(label, ss_key, item_label="Item"):
                if ss_key not in st.session_state:
                    st.session_state[ss_key] = []
                with st.expander(label, expanded=False):
                    _items = st.session_state[ss_key]
                    _to_del = None
                    for _li, _lv in enumerate(_items):
                        _lc1, _lc2 = st.columns([10, 1])
                        with _lc1:
                            _items[_li] = st.text_input(f"{item_label} {_li+1}", value=_lv,
                                                        key=f"{ss_key}_{_li}", label_visibility="collapsed")
                        with _lc2:
                            if st.button("✕", key=f"{ss_key}_del_{_li}"):
                                _to_del = _li
                    if _to_del is not None:
                        _items.pop(_to_del)
                        st.rerun()
                    if st.button(f"+ Add {item_label}", key=f"{ss_key}_add"):
                        _items.append("")
                        st.rerun()

            _wr_list_editor("🏆 Key Achievements",       "wr_achievements", "Achievement")
            _wr_list_editor("⚠️ Challenges & Solutions", "wr_challenges",   "Challenge")
            _wr_list_editor("💡 Innovations & Initiatives", "wr_innovations", "Innovation")

            _wr_wins    = st.text_area("👥 Customer Wins / Losses",                          key="wr_wins",     height=80)
            _wr_focus   = st.text_area("➡️ Upcoming Week Focus",                             key="wr_focus",    height=80)
            _wr_narr    = st.text_area("✏️ Sales Flash Narrative (optional override)",        key="wr_narr",     height=100,
                                       help="Leave blank to auto-generate from data.")

            # Generate button
            if st.button("📄 Generate Reports", type="primary", key="wr_generate"):
                _wr_gen_warnings = []
                try:
                    _wr_tmpdir = tempfile.mkdtemp(prefix="wr_")

                    def _wr_save(f, prefix):
                        if f is None:
                            return None
                        _ext = _os.path.splitext(f.name)[1] or ".xlsx"
                        _p = _os.path.join(_wr_tmpdir, prefix + _ext)
                        with open(_p, "wb") as _out:
                            _out.write(f.getvalue())
                        return _p

                    _paths = {
                        "field_estimates":      _wr_save(_wr_field_est,  "field_estimates"),
                        "shop_estimates":        _wr_save(_wr_shop_est,   "shop_estimates"),
                        "current_wip":           _wr_save(_wr_wip,        "current_wip"),
                        "historic_calibrations": _wr_save(_wr_historic,   "historic_calibrations"),
                        "shop_cal_events":       _wr_save(_wr_cal_events, "shop_cal_events"),
                        "events":                _wr_save(_wr_events,     "events"),
                    }

                    # Parse files
                    field_data = None
                    if _paths["field_estimates"]:
                        field_data = _wr_parse_estimates(_paths["field_estimates"])
                    else:
                        _wr_gen_warnings.append("Field Estimates not uploaded — revenue totals will be incomplete.")

                    shop_data = None
                    if _paths["shop_estimates"]:
                        shop_data = _wr_parse_estimates(_paths["shop_estimates"])

                    historic_lookup = None
                    if _paths["historic_calibrations"]:
                        historic_lookup = _wr_build_historic_lookup(_paths["historic_calibrations"])
                    else:
                        _wr_gen_warnings.append("Historic Calibrations not uploaded — WIP cost estimates unavailable.")

                    wip_data = None
                    if _paths["current_wip"]:
                        if historic_lookup is None:
                            _wr_gen_warnings.append("WIP uploaded but Historic Calibrations missing — cannot estimate WIP value.")
                        else:
                            wip_data = _wr_parse_wip(_paths["current_wip"], historic_lookup)
                            if wip_data["pricing_confidence"].get("unmatched", 0) > 0:
                                _wr_gen_warnings.append(f"{wip_data['pricing_confidence']['unmatched']} WIP item(s) have no price estimate.")
                    else:
                        _wr_gen_warnings.append("Current WIP not uploaded — WIP value will be $0.")

                    cal_data = None
                    if _paths["shop_cal_events"]:
                        cal_data = _wr_parse_cal_events(_paths["shop_cal_events"])
                    else:
                        _wr_gen_warnings.append("Shop Cal Events not uploaded — calibration projection unavailable.")

                    tat_from_file = None
                    if _paths["events"]:
                        try:
                            tat_from_file = _wr_parse_events_tat(_paths["events"])
                            if tat_from_file.get("sample_size", 0) < 10:
                                _wr_gen_warnings.append(f"TAT based on only {tat_from_file.get('sample_size',0)} instruments — low sample size.")
                        except Exception as _tat_e:
                            _wr_gen_warnings.append(f"TAT parse error: {_tat_e}")
                    else:
                        _wr_gen_warnings.append("Events file not uploaded — TAT will use manual inputs only.")

                    def _wr_tat_val(manual_val, file_key):
                        if manual_val and float(manual_val) > 0:
                            return round(float(manual_val), 1)
                        if tat_from_file:
                            return round(tat_from_file.get(file_key, 0), 1)
                        return 0.0

                    _r2c = _wr_tat_val(_wr_tat_r2c, "receive_to_cal")
                    _c2q = _wr_tat_val(_wr_tat_c2q, "cal_to_qc")
                    _q2s = _wr_tat_val(_wr_tat_q2s, "qc_to_ship")
                    tat = {
                        "receive_to_cal": _r2c,
                        "cal_to_qc":      _c2q,
                        "qc_to_ship":     _q2s,
                        "total":          round(_r2c + _c2q + _q2s, 1),
                        "sample_size":    tat_from_file.get("sample_size", 0) if tat_from_file else 0,
                    }

                    if field_data and field_data["confirmed_count"] < 35:
                        _wr_gen_warnings.append(f"Light field schedule: only {field_data['confirmed_count']} confirmed jobs (threshold is 35).")

                    _fc = field_data["confirmed_total"] if field_data else 0.0
                    _fp = field_data["pending_total"]   if field_data else 0.0
                    _st = (shop_data["confirmed_total"] + shop_data["pending_total"]) if shop_data else 0.0
                    _wv = wip_data["total_value"] if wip_data else 0.0
                    _cp = cal_data["monthly_average"] if cal_data else 0
                    revenue = {
                        "field_confirmed": _fc,
                        "field_pending":   _fp,
                        "shop_total":      _st,
                        "wip_value":       _wv,
                        "cal_projection":  _cp,
                        "total":           _fc + _fp + _st + _wv,
                    }

                    manual_inputs = {
                        "tat":                  tat,
                        "achievements":         [a for a in st.session_state.get("wr_achievements", []) if a.strip()],
                        "challenges":           [c for c in st.session_state.get("wr_challenges",   []) if c.strip()],
                        "innovations":          [v for v in st.session_state.get("wr_innovations",  []) if v.strip()],
                        "customer_wins_losses": _wr_wins,
                        "upcoming_focus":       _wr_focus,
                        "sales_flash_narrative": _wr_narr,
                        "week_start":           _wr_week_start,
                        "week_end":             _wr_week_end,
                    }

                    from datetime import datetime as _wr_dt
                    _report_date = _wr_dt.now()
                    _flash_path  = _os.path.join(_wr_tmpdir, "CTPM_Sales_Flash.docx")
                    _update_path = _os.path.join(_wr_tmpdir, "CTPM_Weekly_Update.docx")

                    _wr_build_sales_flash(_flash_path, _report_date, revenue, field_data, wip_data, cal_data, manual_inputs)
                    _wr_build_weekly_update(_update_path, _report_date, revenue, field_data, wip_data, cal_data, manual_inputs)

                    with open(_flash_path, "rb")  as _f: _flash_bytes  = _f.read()
                    with open(_update_path, "rb") as _f: _update_bytes = _f.read()

                    st.session_state["wr_flash_bytes"]  = _flash_bytes
                    st.session_state["wr_update_bytes"] = _update_bytes
                    st.session_state["wr_gen_warnings"] = _wr_gen_warnings
                    st.session_state["wr_summary"] = {
                        "field_confirmed_count": field_data["confirmed_count"]  if field_data  else 0,
                        "field_pending_count":   field_data["pending_count"]    if field_data  else 0,
                        "field_confirmed_total": _fc,
                        "field_pending_total":   _fp,
                        "wip_items":   wip_data["item_count"] if wip_data else 0,
                        "wip_wos":     wip_data["wo_count"]   if wip_data else 0,
                        "wip_value":   _wv,
                        "cal_projection": _cp,
                        "tat":         tat,
                        "revenue_total": revenue["total"],
                        "pricing_confidence": wip_data["pricing_confidence"] if wip_data else {},
                    }
                    st.session_state["wr_generated"] = True
                    st.rerun()

                except Exception as _wr_exc:
                    import traceback as _tb
                    st.error(f"Error generating reports: {_wr_exc}")
                    st.code(_tb.format_exc())

            # Results (shown after rerun once wr_generated is set)
            if st.session_state.get("wr_generated"):
                _wr_w   = st.session_state.get("wr_gen_warnings", [])
                _wr_sum = st.session_state.get("wr_summary", {})
                _flash_b  = st.session_state.get("wr_flash_bytes")
                _update_b = st.session_state.get("wr_update_bytes")

                st.success("Reports generated!")

                _dl1, _dl2 = st.columns(2)
                with _dl1:
                    if _flash_b:
                        st.download_button("⬇️ Sales Flash (.docx)", data=_flash_b,
                                           file_name="CTPM_Sales_Flash.docx",
                                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                           key="dl_flash")
                with _dl2:
                    if _update_b:
                        st.download_button("⬇️ Weekly Update (.docx)", data=_update_b,
                                           file_name="CTPM_Weekly_Update.docx",
                                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                           key="dl_update")

                if _wr_w:
                    with st.expander(f"⚠️ Warnings ({len(_wr_w)})", expanded=True):
                        for _ww in _wr_w:
                            st.warning(_ww)

                if _wr_sum:
                    st.subheader("Processing Summary")
                    _fmtd = lambda v: f"${v:,.0f}" if v else "$0"
                    _fmtn = lambda v: f"{int(round(v)):,}" if v else "0"
                    _sc1, _sc2 = st.columns(2)
                    with _sc1:
                        st.metric("Field Jobs — Confirmed", f"{_wr_sum.get('field_confirmed_count', 0)} jobs ({_fmtd(_wr_sum.get('field_confirmed_total', 0))})")
                        st.metric("Field Jobs — Pending",   f"{_wr_sum.get('field_pending_count', 0)} jobs ({_fmtd(_wr_sum.get('field_pending_total', 0))})")
                        st.metric("WIP Instruments",        f"{_wr_sum.get('wip_items', 0)} items across {_wr_sum.get('wip_wos', 0)} WOs")
                        st.metric("Estimated WIP Value",    _fmtd(_wr_sum.get("wip_value", 0)))
                    with _sc2:
                        st.metric("Projected Calibrations", _fmtn(_wr_sum.get("cal_projection", 0)))
                        st.metric("Total Revenue",          _fmtd(_wr_sum.get("revenue_total", 0)))
                        _tat_s = _wr_sum.get("tat", {})
                        st.metric("Shop TAT (total)",       f"{_tat_s.get('total', 0)} days (n={_tat_s.get('sample_size', 0)})")
                        _conf = _wr_sum.get("pricing_confidence", {})
                        if _conf:
                            st.caption(f"WIP pricing — ID: {_conf.get('id',0)}  Mfr+Model: {_conf.get('mfr_model',0)}  Type: {_conf.get('type',0)}  Unmatched: {_conf.get('unmatched',0)}")

# ---- Upload Data ----
elif page == "📤 Upload Data":
    require_auth(["admin", "internal"])
    st.header("📤 Upload Data")
    st.caption(
        "Upload each IndySoft export individually. "
        "Files are saved to `.ctpm_cache/uploads/` and cached as Parquet so every subsequent load is fast. "
        "**Events** and **Equipment** are required; the others are optional."
    )

    _upl_dir = CACHE_DIR / "uploads"
    _upl_dir.mkdir(parents=True, exist_ok=True)

    # Metadata for each uploader slot
    _UPL_SLOTS = [
        ("events",       "All Events",       True,  "IndySoft → Reports → Event History export"),
        ("equipment",    "All Equipment",     True,  "IndySoft → Equipment list export"),
        ("wos",          "All Work Orders",   False, "IndySoft → Work Order list export"),
        ("all_companies","All Companies",     False, "IndySoft → Companies/Address list export"),
        ("wip_shop",     "WIP Shop",          False, "IndySoft → Current WIP / In-Shop report"),
    ]

    # --- Status table ---
    _current_sheets = _upl_sheet_paths()
    _status_rows = []
    for _key, _label, _req, _ in _UPL_SLOTS:
        _p = _current_sheets.get(_key)
        if _p and _p.exists():
            _s = _p.stat()
            _status_rows.append({
                "Sheet": _label,
                "Required": "✅ yes" if _req else "optional",
                "Status": "✅ loaded",
                "Size (KB)": f"{_s.st_size / 1024:.0f}",
                "Last updated": datetime.fromtimestamp(_s.st_mtime).strftime("%Y-%m-%d %H:%M"),
            })
        else:
            _status_rows.append({
                "Sheet": _label,
                "Required": "✅ yes" if _req else "optional",
                "Status": "⚠️ missing" if _req else "— not uploaded",
                "Size (KB)": "—",
                "Last updated": "—",
            })
    st.dataframe(_status_rows, use_container_width=True, hide_index=True)

    st.divider()

    # --- Individual uploaders ---
    _any_saved = False
    for _key, _label, _req, _hint in _UPL_SLOTS:
        _tag = "required" if _req else "optional"
        _uf = st.file_uploader(
            f"**{_label}** ({_tag})",
            type=["xlsx", "xls"],
            key=f"upl_{_key}",
            help=_hint,
        )
        if _uf is not None:
            _ext = Path(_uf.name).suffix.lower()
            _dest = _upl_dir / f"{_key}{_ext}"
            # Remove old version with different extension first
            for _old_ext in (".xlsx", ".xls"):
                _old = _upl_dir / f"{_key}{_old_ext}"
                if _old.exists() and _old != _dest:
                    try:
                        _old.unlink()
                    except Exception:
                        pass
            with open(_dest, "wb") as _fh:
                _fh.write(_uf.read())
            _s3_upload(_dest)  # best-effort cloud backup
            st.success(f"✅ {_label} saved ({_dest.stat().st_size / 1024:.0f} KB)")
            _any_saved = True

    if _any_saved:
        # Clear Parquet cache so new data is picked up
        for _pf in PARQUET_DIR.glob("*.parquet"):
            try:
                _pf.unlink()
            except Exception:
                pass
        st.cache_data.clear()
        for _k in ("__aging__", "__wip__", "__tatroll__"):
            st.session_state.pop(_k, None)
        st.info("Cache cleared. Reloading with new data…")
        st.rerun()

    st.divider()

    # --- Parquet cache status + clear button ---
    st.subheader("Parquet Cache")
    _pq_files = sorted(PARQUET_DIR.glob("*.parquet"))
    if _pq_files:
        _pq_total_mb = sum(f.stat().st_size for f in _pq_files) / 1_048_576
        st.caption(f"{len(_pq_files)} cached file(s) — {_pq_total_mb:.1f} MB total")
        st.dataframe(
            [{"File": f.name,
              "Size (KB)": f"{f.stat().st_size / 1024:.0f}",
              "Modified": datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M")}
             for f in _pq_files],
            use_container_width=True, hide_index=True,
        )
        if st.button("🗑️ Clear Parquet Cache", type="secondary"):
            _n = 0
            for _pf in PARQUET_DIR.glob("*.parquet"):
                try:
                    _pf.unlink(); _n += 1
                except Exception:
                    pass
            st.cache_data.clear()
            for _k in ("__aging__", "__wip__", "__tatroll__"):
                st.session_state.pop(_k, None)
            st.success(f"Cleared {_n} Parquet file(s). Next load will re-read from the uploaded files.")
            st.rerun()
    else:
        st.caption("No Parquet cache yet — built automatically on first load after upload.")

    # --- S3 status ---
    st.divider()
    st.subheader("Cloud Storage (S3)")
    if _s3_bucket_name():
        st.success(f"✅ Connected — bucket: `{_s3_bucket_name()}`  \nUploaded files are automatically backed up and restored on restart.")
    else:
        st.info(
            "S3 is not configured. Set the `CTPM_S3_BUCKET` environment variable to an S3 bucket name "
            "to enable automatic cloud backup of uploaded files. "
            "The container will also need an IAM role with `s3:GetObject` / `s3:PutObject` on that bucket."
        )