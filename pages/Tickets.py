# pages/tickets_page.py
import os
import tempfile
from pathlib import Path

import msal
import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# ==========================================
# PAGE CONFIG
# ==========================================
st.set_page_config(page_title="Tickets", layout="wide")
st.title("Tickets")

# ==========================================
# ENV (SharePoint file path + refresh cadence)
# ==========================================
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")

SP_HOSTNAME = os.getenv("SP_HOSTNAME")      # groupcastillo.sharepoint.com
SP_SITE_PATH = os.getenv("SP_SITE_PATH")    # /sites/GroupCastilloTeamSite
SP_DRIVE_NAME = os.getenv("SP_DRIVE_NAME", "Documents")

TICKETS_SP_PATH = os.getenv(
    "SP_FILE_PATH",
    "General/12433087 CANADA INC-MASTER/21-Work Orders-Complaints-Request/WorkOrders-Complaints-Master-2025-v1.xlsm"
)

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All"]

TICKETS_REFRESH_SECONDS = 30 * 60  # 30 minutes

# ==========================================
# UI CONSTANTS
# ==========================================
SHEETS = {
    "Work Orders": {"sheet": "Work Orders", "status_col": "General Status"},
    "Request": {"sheet": "Request", "status_col": "Status"},
    "Complaints": {"sheet": "Complaints", "status_col": "Status"},
}

PRIORITY_COLORS = {"High": "#d32f2f", "Medium": "#fbc02d", "Low": "#388e3c"}
PRIORITY_COLORS_LIGHT = {"High": "#f28b82", "Medium": "#ffe082", "Low": "#a5d6a7"}

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

@st.cache_data(show_spinner=False, ttl=TICKETS_REFRESH_SECONDS)
def download_tickets_excel_cached(sp_relative_path: str) -> str:
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
# SMALL UI HELPERS
# ==========================================
def thumb_card(message: str, height_px: int = 420):
    st.markdown(
        f"""
        <div style="
            height:{height_px}px;
            display:flex;
            flex-direction:column;
            align-items:center;
            justify-content:center;
            border-radius:18px;
            background:rgba(255,255,255,0.06);
            border:1px solid rgba(255,255,255,0.12);
        ">
            <div style="font-size:96px;">👍</div>
            <div style="font-size:26px;font-weight:700;">{message}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ==========================================
# NORMALIZATION
# ==========================================
def _clean_text(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip()
    return s.replace({"": None, "nan": None, "None": None})

def normalize_priority(s: pd.Series) -> pd.Series:
    raw = _clean_text(s)
    out = []
    for v in raw:
        if v is None:
            out.append(None)
        else:
            vl = v.lower()
            if "high" in vl:
                out.append("High")
            elif "medium" in vl:
                out.append("Medium")
            elif "low" in vl:
                out.append("Low")
            else:
                out.append(None)
    return pd.Series(out, index=s.index)

def normalize_status(s: pd.Series) -> pd.Series:
    raw = _clean_text(s)
    out = []
    for v in raw:
        if v is None:
            out.append(None)
        else:
            vl = v.lower()
            if "closed" in vl:
                out.append("Closed")
            elif "progress" in vl:
                out.append("In Progress")
            elif "open" in vl:
                out.append("Open")
            else:
                out.append("Other")
    return pd.Series(out, index=s.index)

def normalize_assigned_to(s: pd.Series) -> pd.Series:
    return _clean_text(s)

# ==========================================
# FILTERS
# ==========================================
def _prep_common(df: pd.DataFrame, status_col: str) -> pd.DataFrame:
    if status_col not in df.columns:
        return df.iloc[0:0]

    d = df.copy()

    if "Priority" in d.columns:
        d["Priority"] = normalize_priority(d["Priority"])
    else:
        d["Priority"] = None

    d[status_col] = normalize_status(d[status_col])

    if "Assigned To" in d.columns:
        d["Assigned To"] = normalize_assigned_to(d["Assigned To"])
    else:
        d["Assigned To"] = None

    d = d.dropna(subset=["Priority", status_col])
    return d

def filter_not_closed(df: pd.DataFrame, status_col: str) -> pd.DataFrame:
    d = _prep_common(df, status_col)
    return d[d[status_col] != "Closed"]

def filter_closed(df: pd.DataFrame, status_col: str) -> pd.DataFrame:
    d = _prep_common(df, status_col)
    return d[d[status_col] == "Closed"]

# ==========================================
# TABLE STYLING (ROW COLOR BY PRIORITY)
# ==========================================
def style_by_priority(df: pd.DataFrame):
    def row_style(row):
        p = row.get("Priority")
        if p == "High":
            return [f"background-color: {PRIORITY_COLORS_LIGHT['High']}; color:black"] * len(row)
        if p == "Medium":
            return [f"background-color: {PRIORITY_COLORS_LIGHT['Medium']}; color:black"] * len(row)
        if p == "Low":
            return [f"background-color: {PRIORITY_COLORS_LIGHT['Low']}; color:black"] * len(row)
        return [""] * len(row)

    return df.style.apply(row_style, axis=1)

# ==========================================
# CHARTS
# ==========================================
def open_stacked_chart(df: pd.DataFrame, status_col: str, title: str):
    if df.empty:
        thumb_card("0 tickets pendientes")
        return

    g = df.groupby(["Priority", status_col]).size().reset_index(name="Count")
    g["ColorKey"] = g["Priority"] + "|" + g[status_col]
    g["Label"] = g[status_col] + ": " + g["Count"].astype(str)

    color_map = {}
    for p in ["High", "Medium", "Low"]:
        color_map[f"{p}|Open"] = PRIORITY_COLORS[p]
        color_map[f"{p}|In Progress"] = PRIORITY_COLORS_LIGHT[p]
        color_map[f"{p}|Other"] = PRIORITY_COLORS[p]

    fig = px.bar(
        g,
        x="Count",
        y="Priority",
        color="ColorKey",
        orientation="h",
        color_discrete_map=color_map,
        text="Label",
        title=title,
    )
    fig.update_layout(barmode="stack", showlegend=False)
    fig.update_traces(textposition="inside")
    st.plotly_chart(fig, use_container_width=True)

def closed_pie_chart(df: pd.DataFrame, title: str):
    if df.empty:
        thumb_card("0 tickets cerrados")
        return

    g = df.groupby("Priority").size().reset_index(name="Count")
    fig = px.pie(
        g,
        names="Priority",
        values="Count",
        title=title,
        color="Priority",
        color_discrete_map=PRIORITY_COLORS,
        hole=0.35,
    )
    st.plotly_chart(fig, use_container_width=True)

def assigned_to_bars_stacked_by_priority(df_all: pd.DataFrame, title: str):
    if df_all.empty:
        thumb_card("0 tickets", 260)
        return

    g = df_all.groupby(["Assigned To", "Priority"]).size().reset_index(name="Count")
    order = g.groupby("Assigned To")["Count"].sum().sort_values(ascending=False).index
    n_assignees = len(order)

    fig = px.bar(
        g,
        x="Count",
        y="Assigned To",
        color="Priority",
        orientation="h",
        category_orders={"Assigned To": list(order)},
        color_discrete_map=PRIORITY_COLORS,
        title=title,
        text="Count",
    )
    fig.update_layout(
        barmode="stack",
        height=max(320, n_assignees * 48),
        margin=dict(l=140, r=40, t=60, b=40),
    )
    fig.update_traces(textposition="outside", textangle=0, cliponaxis=False)
    st.plotly_chart(fig, use_container_width=True)

def monthly_trend_chart(data_by_sheet: dict[str, pd.DataFrame]):
    rows = []
    for name, df in data_by_sheet.items():
        if "Date of the Work" not in df.columns:
            continue

        d = df.copy()
        d["Date of the Work"] = pd.to_datetime(d["Date of the Work"], errors="coerce")
        d = d.dropna(subset=["Date of the Work"])
        d["Month"] = d["Date of the Work"].dt.to_period("M").dt.to_timestamp()
        g = d.groupby("Month").size().reset_index(name="Count")
        g["Type"] = name
        rows.append(g)

    if not rows:
        return

    allg = pd.concat(rows, ignore_index=True)
    fig = px.line(allg, x="Month", y="Count", color="Type", markers=True, title="Monthly trend")
    st.plotly_chart(fig, use_container_width=True)

# ==========================================
# LOAD DATA (auto-refresh every 30 minutes)
# ==========================================
try:
    with st.spinner("Syncing tickets data..."):
        local_path = download_tickets_excel_cached(TICKETS_SP_PATH)
except Exception as e:
    st.error(str(e))
    st.stop()

EXCEL_PATH = Path(local_path)
if not EXCEL_PATH.exists():
    st.error("Tickets cache file missing after download.")
    st.stop()

data: dict[str, pd.DataFrame] = {}
try:
    for name, meta in SHEETS.items():
        data[name] = pd.read_excel(EXCEL_PATH, sheet_name=meta["sheet"])
except Exception as e:
    st.error(f"Could not read Excel sheets: {e}")
    st.stop()

# ==========================================
# UI ORDER
#   1) Three charts section (Open / Closed / Tables)
#   2) Assignees bar charts (Open / Closed)
#   3) Monthly trend line chart
# ==========================================

# -------------------------------------------------------------------
# 1) THREE CHARTS SECTION
# -------------------------------------------------------------------
st.header("By Type")
tab_3_open, tab_3_closed, tab_3_tables = st.tabs(["Open", "Closed", "Tables (Open)"])

with tab_3_open:
    c1, c2, c3 = st.columns(3)
    for col, name in zip([c1, c2, c3], SHEETS):
        with col:
            st.subheader(name)
            status_col = SHEETS[name]["status_col"]
            df_nc = filter_not_closed(data[name], status_col)
            open_stacked_chart(df_nc, status_col, "By priority")

with tab_3_closed:
    c1, c2, c3 = st.columns(3)
    for col, name in zip([c1, c2, c3], SHEETS):
        with col:
            st.subheader(name)
            status_col = SHEETS[name]["status_col"]
            df_c = filter_closed(data[name], status_col)
            closed_pie_chart(df_c, "By priority")

with tab_3_tables:
    for name in SHEETS:
        st.subheader(f"{name} (Not Closed)")
        status_col = SHEETS[name]["status_col"]
        df_nc = filter_not_closed(data[name], status_col)

        if df_nc.empty:
            st.info("No open tickets.")
        else:
            st.dataframe(style_by_priority(df_nc), use_container_width=True, hide_index=True)

# -------------------------------------------------------------------
# 2) ASSIGNEES BAR CHARTS SECTION
# -------------------------------------------------------------------
st.header("Assignees")
tab_a_open, tab_a_closed = st.tabs(["Open", "Closed"])

with tab_a_open:
    sources_open = st.multiselect(
        "Sources: include",
        options=list(SHEETS.keys()),
        default=list(SHEETS.keys()),
        key="assignees_open_sources",
    )

    open_combined = []
    for name in sources_open:
        status_col = SHEETS[name]["status_col"]
        df_nc = filter_not_closed(data[name], status_col)
        if not df_nc.empty and "Assigned To" in df_nc.columns:
            open_combined.append(df_nc[["Assigned To", "Priority"]])

    df_open_all = pd.concat(open_combined, ignore_index=True) if open_combined else pd.DataFrame()
    assigned_to_bars_stacked_by_priority(df_open_all, "Assignees")

with tab_a_closed:
    sources_closed = st.multiselect(
        "Sources: include",
        options=list(SHEETS.keys()),
        default=list(SHEETS.keys()),
        key="assignees_closed_sources",
    )

    closed_combined = []
    for name in sources_closed:
        status_col = SHEETS[name]["status_col"]
        df_c = filter_closed(data[name], status_col)
        if not df_c.empty and "Assigned To" in df_c.columns:
            closed_combined.append(df_c[["Assigned To", "Priority"]])

    df_closed_all = pd.concat(closed_combined, ignore_index=True) if closed_combined else pd.DataFrame()
    assigned_to_bars_stacked_by_priority(df_closed_all, "Assignees")

# -------------------------------------------------------------------
# 3) TRENDS
# -------------------------------------------------------------------
st.header("Trends")
monthly_trend_chart(data)
