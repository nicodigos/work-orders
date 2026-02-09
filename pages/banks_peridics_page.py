# pages/banks_periodics_page.py
import os
import tempfile
from pathlib import Path

import msal
import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from openpyxl import load_workbook

# ==========================================
# PAGE CONFIG
# ==========================================
st.set_page_config(page_title="Banks Periodics", layout="wide")
st.title("Banks Periodics")

# ==========================================
# ENV (SharePoint file path + refresh cadence)
# ==========================================
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")

SP_HOSTNAME = os.getenv("SP_HOSTNAME")      # groupcastillo.sharepoint.com
SP_SITE_PATH = os.getenv("SP_SITE_PATH")    # /sites/GroupCastilloTeamSite
SP_DRIVE_NAME = os.getenv("SP_DRIVE_NAME", "Documents")

BANKS_SP_PATH = os.getenv(
    "SP_BANKS_FILE_PATH",
    "General/9359-6633 QUEBEC INC/BGIS/Banks Periodics/2026.xlsx"
)

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All"]

BANKS_REFRESH_SECONDS = 3 * 60 * 60  # 3 hours

# ==========================================
# TOKEN CACHE (disk) - SILENT ONLY (no login UI here)
# ==========================================
def _token_cache_path() -> Path:
    d = Path(tempfile.gettempdir()) / "cnet_reports"
    d.mkdir(exist_ok=True)
    return d / "msal_token_cache.bin"

def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    p = _token_cache_path()
    if p.exists():
        cache.deserialize(p.read_text(encoding="utf-8"))
    return cache

def _save_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        _token_cache_path().write_text(cache.serialize(), encoding="utf-8")

def _msal_app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    if not TENANT_ID or not CLIENT_ID:
        raise RuntimeError("Missing TENANT_ID / CLIENT_ID in environment.")
    return msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

def get_token_silent_only() -> str:
    cache = _load_cache()
    app = _msal_app(cache)

    accounts = app.get_accounts()
    if not accounts:
        raise RuntimeError("Not authenticated. Please connect in the main app (app.py).")

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        _save_cache(cache)
        return result["access_token"]

    raise RuntimeError("Session expired. Please reconnect in the main app (app.py).")

# ==========================================
# GRAPH HELPERS
# ==========================================
def graph_get(url: str, token: str):
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.json()

def graph_download(url: str, token: str) -> bytes:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.content

def resolve_drive_id(token: str) -> str:
    if not SP_HOSTNAME or not SP_SITE_PATH:
        raise RuntimeError("Missing SP_HOSTNAME / SP_SITE_PATH in environment.")

    site = graph_get(f"https://graph.microsoft.com/v1.0/sites/{SP_HOSTNAME}:{SP_SITE_PATH}", token)
    drives = graph_get(f"https://graph.microsoft.com/v1.0/sites/{site['id']}/drives", token)["value"]
    drive = next((d for d in drives if d.get("name") == SP_DRIVE_NAME), drives[0])
    return drive["id"]

@st.cache_data(show_spinner=False, ttl=BANKS_REFRESH_SECONDS)
def download_banks_excel_cached(sp_relative_path: str) -> str:
    token = get_token_silent_only()
    drive_id = resolve_drive_id(token)
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{sp_relative_path}:/content"
    content = graph_download(url, token)

    out_dir = Path(tempfile.gettempdir()) / "cnet_reports"
    out_dir.mkdir(exist_ok=True)
    local = out_dir / Path(sp_relative_path).name
    local.write_bytes(content)
    return str(local)

# ==========================================
# EXCEL HELPERS (VISIBLE SHEETS + HEADER DETECT)
# ==========================================
@st.cache_data(show_spinner=False)
def get_visible_sheet_names(xlsx_path: str) -> list[str]:
    wb = load_workbook(filename=xlsx_path, read_only=True, data_only=True)
    visible = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
    wb.close()
    return visible

def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    s = str(v).strip()
    return s == "" or s.lower() in {"nan", "none"}

@st.cache_data(show_spinner=False)
def detect_header_row(xlsx_path: str, sheet_name: str, scan_rows: int = 80) -> int:
    preview = pd.read_excel(
        xlsx_path,
        sheet_name=sheet_name,
        header=None,
        nrows=scan_rows,
        engine="openpyxl",
    )

    best_i = 0
    best_score = -1.0
    for i in range(len(preview)):
        row = preview.iloc[i].tolist()
        non_blank = [v for v in row if not _is_blank(v)]
        if len(non_blank) < 2:
            continue

        as_str = [str(v).strip() for v in non_blank]
        str_like = sum(1 for v in non_blank if isinstance(v, str))
        uniqueness = len(set(as_str)) / max(1, len(as_str))
        score = (len(non_blank) * 1.5) + (str_like * 2.0) + (uniqueness * 3.0)

        if score > best_score:
            best_score = score
            best_i = i

    return int(best_i)

@st.cache_data(show_spinner=False)
def read_sheet_with_detected_header(xlsx_path: str, sheet_name: str, header_row: int) -> pd.DataFrame:
    df = pd.read_excel(
        xlsx_path,
        sheet_name=sheet_name,
        header=header_row,
        engine="openpyxl",
    )
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df

# ==========================================
# FIXED COLUMNS (BANK + ADDRESS)
# ==========================================
BANK_FALLBACKS = ["bank", "banco"]
ADDRESS_FALLBACKS = ["address", "adresse", "direccion", "dirección", "addr"]

def find_required_col(df: pd.DataFrame, fallbacks: list[str]) -> str | None:
    for c in df.columns:
        cl = str(c).strip().lower()
        if any(k in cl for k in fallbacks):
            return c
    return None

def to_text_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().replace({"": None, "nan": None, "None": None})

# ==========================================
# DONE / PENDING NORMALIZATION (CELL LEVEL)
# ==========================================
DONE_WORDS = {"done", "completed", "complete", "ok", "yes"}
PENDING_WORDS = {"pending", "pendiente", "to do", "todo", "open", "in progress"}
NOT_SCHEDULED_WORDS = {"not scheduled", "n/a", "na", "tbd"}

def normalize_status_cell(v) -> str | None:
    if _is_blank(v):
        return None
    s = str(v).strip().lower()
    if any(w in s for w in DONE_WORDS):
        return "Done"
    if any(w in s for w in PENDING_WORDS):
        return "Pending"
    if any(w in s for w in NOT_SCHEDULED_WORDS):
        return "Not Scheduled"
    return None

# ==========================================
# PIE CHARTS (GREEN / RED)
# ==========================================
PIE_COLORS = {"Done": "#2e7d32", "Pending": "#c62828"}

def make_done_pending_pie(done: int, pending: int, title: str):
    total = done + pending
    if total == 0:
        st.caption(f"{title}: n/a")
        return

    dfp = pd.DataFrame({"Status": ["Done", "Pending"], "Count": [done, pending]})
    fig = px.pie(
        dfp,
        names="Status",
        values="Count",
        hole=0.55,
        color="Status",
        color_discrete_map=PIE_COLORS,
        title=title,
    )
    fig.update_layout(height=170, margin=dict(l=10, r=10, t=35, b=10), showlegend=False)
    fig.update_traces(textposition="inside", textinfo="percent")
    st.plotly_chart(fig, use_container_width=True)

# ==========================================
# TABLE STYLING
#   - Bank: colored if exact match in map
#   - Address: dark gray background, white text
#   - Pending: red
#   - Done: green
#   - Not Scheduled: unchanged
# ==========================================
BANK_STYLES = {
    "TD":   {"bg": "#54B848", "fg": "white"},
    "CIBC": {"bg": "#6f1729", "fg": "white"},
    "NB":   {"bg": "white",   "fg": "red"},
    "RBC":  {"bg": "yellow",  "fg": "blue"},
    "BMO":  {"bg": "blue",    "fg": "white"},
}
ADDRESS_STYLE = "background-color:#2b2b2b; color:white; font-weight:600;"

def cell_style(v, is_bank_col: bool = False, is_addr_col: bool = False) -> str:
    if is_addr_col:
        return ADDRESS_STYLE

    if is_bank_col:
        if v in BANK_STYLES:
            s = BANK_STYLES[v]
            return f"background-color:{s['bg']}; color:{s['fg']}; font-weight:700;"
        return ""

    norm = normalize_status_cell(v)
    if norm == "Pending":
        return "background-color:#ffcdd2; color:#b71c1c; font-weight:600;"
    if norm == "Done":
        return "background-color:#c8e6c9; color:#1b5e20; font-weight:600;"
    return ""

def style_table(df: pd.DataFrame, bank_col: str, addr_col: str):
    def _row_style(row):
        styles = []
        for c in df.columns:
            styles.append(
                cell_style(
                    row.get(c),
                    is_bank_col=(c == bank_col),
                    is_addr_col=(c == addr_col),
                )
            )
        return styles

    return df.style.apply(_row_style, axis=1)

# ==========================================
# LOAD FILE (auto-refresh every 3 hours)
# ==========================================
try:
    with st.spinner("Syncing banks data..."):
        local_path = download_banks_excel_cached(BANKS_SP_PATH)
except Exception as e:
    st.error(str(e))
    st.stop()

EXCEL_PATH_LOCAL = Path(local_path)
if not EXCEL_PATH_LOCAL.exists():
    st.error("Banks cache file missing after download.")
    st.stop()

# ==========================================
# UI
# ==========================================
visible_sheets = get_visible_sheet_names(str(EXCEL_PATH_LOCAL))
if not visible_sheets:
    st.error("No visible sheets found.")
    st.stop()

sheet = st.selectbox("Sheet", options=visible_sheets)

header_row = detect_header_row(str(EXCEL_PATH_LOCAL), sheet)
df_raw = read_sheet_with_detected_header(str(EXCEL_PATH_LOCAL), sheet, header_row)

if df_raw.empty:
    st.info("No data found on this sheet.")
    st.stop()

bank_col = find_required_col(df_raw, BANK_FALLBACKS)
addr_col = find_required_col(df_raw, ADDRESS_FALLBACKS)
if not bank_col or not addr_col:
    st.error("Bank or Address column not found.")
    st.stop()

df_raw[bank_col] = to_text_series(df_raw[bank_col])
df_raw[addr_col] = to_text_series(df_raw[addr_col])

bank_vals = sorted(df_raw[bank_col].dropna().unique().tolist())
addr_vals = sorted(df_raw[addr_col].dropna().unique().tolist())

c1, c2 = st.columns(2)
with c1:
    bank_sel = st.multiselect("Filter: Bank", options=bank_vals, default=[])
with c2:
    addr_sel = st.multiselect("Filter: Address", options=addr_vals, default=[])

# Empty selection means include all
banks_to_use = bank_sel if bank_sel else bank_vals
addrs_to_use = addr_sel if addr_sel else addr_vals

df = df_raw[df_raw[bank_col].isin(banks_to_use) & df_raw[addr_col].isin(addrs_to_use)]

# ==========================================
# PIES (COLUMN ORDER)
# ==========================================
st.subheader("Completion by column (Done vs Pending)")

task_cols = [c for c in df_raw.columns if c not in {bank_col, addr_col}]
for i in range(0, len(task_cols), 6):
    block = task_cols[i : i + 6]
    cols = st.columns(len(block))
    for ui_col, c in zip(cols, block):
        with ui_col:
            ser = df[c] if not df.empty else pd.Series([], dtype=object)
            norm = ser.map(normalize_status_cell)
            done = int((norm == "Done").sum())
            pending = int((norm == "Pending").sum())

            title = str(c)
            if len(title) > 22:
                title = title[:21] + "…"

            make_done_pending_pie(done, pending, title)

# ==========================================
# MATRIX
# ==========================================
st.subheader("Matrix")

if df.empty:
    st.info("No rows match filters.")
else:
    df_show = df.copy()
    for c in df_show.columns:
        df_show[c] = df_show[c].map(lambda v: None if _is_blank(v) else str(v).strip())

    st.dataframe(
        style_table(df_show, bank_col=bank_col, addr_col=addr_col),
        use_container_width=True,
        hide_index=True,
    )
