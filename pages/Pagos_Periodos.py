# pages/pagos_periodos_read_billing_info.py

import os
import re
import tempfile
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import streamlit as st
from utils.ms_graph_excel import (
    download_drive_item_content,
    get_token_silent_or_raise,
    list_children_by_path,
    require_graph_login,
    resolve_drive_id,
)

# ==========================================
# PAGE CONFIG
# ==========================================
st.set_page_config(page_title="One-Shot Period Payments", layout="wide")
st.title("One-Shot Period Payments")
require_graph_login()

# ==========================================
# ENV
# ==========================================
BASE_SP_DIR = os.getenv(
    "SP_PAGOS_PERIODOS_DIR",
    "General/12433087 CANADA INC-MASTER/09-Pagos Periodos",
)

START_YEAR = 2025
START_MONTH = 3

SHEET_ONESHOT = "Billing Info Oneshot"
ONESHOT_MARKER = "Vendor Company"
EXCEL_EXTENSIONS = (".xlsx", ".xlsm")

# ==========================================
# STATUS COLOR REGISTRY
# ==========================================
STATUS_COLOR_MAP = {}
COLOR_PALETTE = [
    "#2E86C1",
    "#28B463",
    "#E74C3C",
    "#F39C12",
    "#8E44AD",
    "#16A085",
    "#D35400",
    "#7D3C98",
]

def get_color_for_status(status: str) -> str:
    status = str(status).strip()
    if status not in STATUS_COLOR_MAP:
        index = len(STATUS_COLOR_MAP) % len(COLOR_PALETTE)
        STATUS_COLOR_MAP[status] = COLOR_PALETTE[index]
    return STATUS_COLOR_MAP[status]

def normalize_status_series(s: pd.Series) -> pd.Series:
    s2 = s.copy()
    s2 = s2.where(~s2.isna(), pd.NA)
    s2 = s2.apply(lambda v: v if pd.isna(v) else str(v).strip())
    s2 = s2.replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA})
    return s2

def is_folder(it: dict) -> bool:
    return "folder" in it

def is_file(it: dict) -> bool:
    return "file" in it

# ==========================================
# HEADER DETECTION
# ==========================================
def _norm_cell(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v).strip()).lower()

def find_header_position(df_raw, marker):
    marker_norm = _norm_cell(marker)
    for r in range(min(120, len(df_raw))):
        for c in range(min(60, len(df_raw.columns))):
            if _norm_cell(df_raw.iat[r, c]) == marker_norm:
                return r, c
    return None, None

def read_oneshot_from_bytes(xlsx_bytes: bytes):
    tmp_dir = Path(tempfile.gettempdir()) / "cnet_reports"
    tmp_dir.mkdir(exist_ok=True)
    tmp_path = tmp_dir / f"pp_{abs(hash(xlsx_bytes))}.xlsx"
    tmp_path.write_bytes(xlsx_bytes)

    try:
        df_raw = pd.read_excel(
            tmp_path,
            sheet_name=SHEET_ONESHOT,
            header=None,
            engine="openpyxl",
            nrows=200,
        )
    except Exception:
        return None

    hr, hc = find_header_position(df_raw, ONESHOT_MARKER)
    if hr is None:
        return None

    df_full = pd.read_excel(
        tmp_path,
        sheet_name=SHEET_ONESHOT,
        header=None,
        engine="openpyxl",
    )

    header_row = (
        df_full.iloc[hr, hc:]
        .astype(str)
        .map(lambda s: re.sub(r"\s+", " ", s.strip()))
    )

    df = df_full.iloc[hr + 1 :, hc:].copy()
    df.columns = [str(c).strip() for c in header_row]

    df = df.loc[:, ~df.columns.duplicated(keep="first")]
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df = df.loc[:, [c for c in df.columns if not _norm_cell(c).startswith("unnamed")]]

    return df

# ==========================================
# FILE LISTING
# ==========================================
@st.cache_data(ttl=600)
def list_excel_files(base_dir: str):
    token = get_token_silent_or_raise(
        "Not authenticated. Please connect in the main app (app.py).",
        "Session expired. Please reconnect in the main app (app.py).",
    )
    drive_id = resolve_drive_id(token)

    results = []
    year_items = list_children_by_path(drive_id, base_dir, token)

    for y in year_items:
        if not is_folder(y) or not str(y["name"]).isdigit():
            continue

        year_int = int(y["name"])
        if year_int < START_YEAR:
            continue

        year_path = f"{base_dir.rstrip('/')}/{year_int}"
        months = list_children_by_path(drive_id, year_path, token)

        for m in months:
            if not is_folder(m):
                continue

            month_path = f"{year_path}/{m['name']}"
            children = list_children_by_path(drive_id, month_path, token)

            for ch in children:
                file_name = str(ch["name"])
                if is_file(ch) and file_name.lower().endswith(EXCEL_EXTENSIONS):
                    results.append(
                        {
                            "drive_id": drive_id,
                            "item_id": ch["id"],
                            "year": year_int,
                            "month": m["name"],
                            "name": file_name,
                            "web_url": ch.get("webUrl"),
                        }
                    )

    return sorted(results, key=lambda x: (x["year"], x["month"], x["name"]))

# ==========================================
# UI
# ==========================================
if st.sidebar.button("Refresh file list", key="pagos_periodos_refresh_file_list"):
    list_excel_files.clear()
    st.rerun()

files = list_excel_files(BASE_SP_DIR)

def route_label(file_item: dict) -> str:
    return f"{file_item['year']}/{file_item['month']}"

all_routes = sorted({route_label(f) for f in files})

selected_routes = st.sidebar.multiselect(
    "Filter by route (Year/Month)",
    options=all_routes,
    default=[],
)

if selected_routes:
    selected_set = set(selected_routes)
    files = [f for f in files if route_label(f) in selected_set]
else:
    files = []

# ==========================================
# RENDER
# ==========================================
token = get_token_silent_or_raise(
    "Not authenticated. Please connect in the main app (app.py).",
    "Session expired. Please reconnect in the main app (app.py).",
)

for f in files:
    title = f"{f['year']} / {f['month']} — {f['name']}"

    file_url = f.get("web_url")
    if file_url:
        sheet_enc = quote(SHEET_ONESHOT)
        deep_link = f"{file_url}?action=edit&activeCell='{sheet_enc}'!A1"

        st.markdown(
            f"""
            <h3 style="margin:0;">
                <a href="{deep_link}" target="_blank"
                style="color: inherit; text-decoration: none;">
                {title}
                </a>
            </h3>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(f"### {title}")

    try:
        xbytes = download_drive_item_content(f["drive_id"], f["item_id"], token)
        df = read_oneshot_from_bytes(xbytes)
    except Exception as e:
        st.error(f"{title}: {e}")
        st.divider()
        continue

    if df is None:
        st.warning("Oneshot sheet not found.")
        st.divider()
        continue

    status_col = next((c for c in df.columns if str(c).strip().lower() == "status"), None)
    if status_col is None:
        st.warning("Column 'Status' not found.")
        st.divider()
        continue

    status_norm = normalize_status_series(df[status_col])
    status_counts = status_norm.dropna().value_counts().sort_index()

    if len(status_counts) > 0:
        cols = st.columns(len(status_counts))
        for i, (status, count) in enumerate(status_counts.items()):
            color = get_color_for_status(status)
            cols[i].markdown(
                f"""
                <div style="padding:0 0 8px 0;">
                    <div style="font-size:13px;color:gray;">{status}</div>
                    <div style="font-size:28px;font-weight:bold;color:{color};">{count}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    with st.expander("View table", expanded=False):

        def color_rows_by_status(row: pd.Series):
            raw = row.get(status_col, pd.NA)
            if pd.isna(raw):
                return [""] * len(row)
            s = str(raw).strip()
            if s == "" or s.lower() == "nan":
                return [""] * len(row)
            bg = get_color_for_status(s)
            return [f"background-color: {bg}; color: white;"] * len(row)

        st.dataframe(
            df.style.apply(color_rows_by_status, axis=1),
            use_container_width=True,
            hide_index=True,
        )

    st.divider()
