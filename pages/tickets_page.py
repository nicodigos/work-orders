# pages/tickets_page.py
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px

# ==========================================
# PAGE CONFIG
# ==========================================
st.set_page_config(page_title="Tickets", layout="wide")
st.title("Tickets")

# ==========================================
# GET EXCEL PATH FROM MAIN APP
# ==========================================
if "excel_path" not in st.session_state:
    st.error("Excel not loaded. Go to main page first and connect to Microsoft.")
    st.stop()

EXCEL_PATH = Path(st.session_state["excel_path"])

# ==========================================
# EXCEL SHEET CONFIG
# ==========================================
SHEETS = {
    "Work Orders": {"sheet": "Work Orders", "status_col": "General Status"},
    "Request": {"sheet": "Request", "status_col": "Status"},
    "Complaints": {"sheet": "Complaints", "status_col": "Status"},
}

PRIORITY_COLORS = {"High": "#d32f2f", "Medium": "#fbc02d", "Low": "#388e3c"}
PRIORITY_COLORS_LIGHT = {"High": "#f28b82", "Medium": "#ffe082", "Low": "#a5d6a7"}

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
def _clean_text(s):
    s = s.astype(str).str.strip()
    return s.replace({"": None, "nan": None, "None": None})

def normalize_priority(s):
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

def normalize_status(s):
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

def normalize_assigned_to(s):
    return _clean_text(s)

# ==========================================
# FILTERS
# ==========================================
def filter_not_closed(df, status_col):
    if status_col not in df.columns:
        return df.iloc[0:0]

    d = df.copy()
    d["Priority"] = normalize_priority(d["Priority"])
    d[status_col] = normalize_status(d[status_col])

    if "Assigned To" in d.columns:
        d["Assigned To"] = normalize_assigned_to(d["Assigned To"])

    d = d.dropna(subset=["Priority", status_col])
    return d[d[status_col] != "Closed"]

def filter_closed(df, status_col):
    if status_col not in df.columns:
        return df.iloc[0:0]

    d = df.copy()
    d["Priority"] = normalize_priority(d["Priority"])
    d[status_col] = normalize_status(d[status_col])

    if "Assigned To" in d.columns:
        d["Assigned To"] = normalize_assigned_to(d["Assigned To"])

    d = d.dropna(subset=["Priority", status_col])
    return d[d[status_col] == "Closed"]

# ==========================================
# CHARTS
# ==========================================
def open_stacked_chart(df, status_col, title):
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

def closed_pie_chart(df, title):
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

def assigned_to_bars_stacked_by_priority(df_all, title):
    if df_all.empty:
        thumb_card("0 tickets", 260)
        return

    g = df_all.groupby(["Assigned To", "Priority"]).size().reset_index(name="Count")
    order = g.groupby("Assigned To")["Count"].sum().sort_values(ascending=False).index

    fig = px.bar(
        g,
        x="Count",
        y="Assigned To",
        color="Priority",
        orientation="h",
        category_orders={"Assigned To": order},
        color_discrete_map=PRIORITY_COLORS,
        title=title,
        text="Count",
    )
    fig.update_layout(barmode="stack")
    st.plotly_chart(fig, use_container_width=True)

def monthly_trend_chart(data_by_sheet):
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

    allg = pd.concat(rows)
    fig = px.line(allg, x="Month", y="Count", color="Type", markers=True, title="Monthly trend")
    st.plotly_chart(fig, use_container_width=True)

# ==========================================
# LOAD DATA
# ==========================================
data = {}
for name, meta in SHEETS.items():
    data[name] = pd.read_excel(EXCEL_PATH, sheet_name=meta["sheet"])

# ==========================================
# UI
# ==========================================
monthly_trend_chart(data)

tab_open, tab_closed = st.tabs(["Open", "Closed"])

# -------- OPEN --------
with tab_open:
    c1, c2, c3 = st.columns(3)
    for col, name in zip([c1, c2, c3], SHEETS):
        with col:
            st.subheader(name)
            status_col = SHEETS[name]["status_col"]
            df_nc = filter_not_closed(data[name], status_col)
            open_stacked_chart(df_nc, status_col, "By priority")

    st.divider()

    sources = st.multiselect(
        "Assignees: include",
        options=list(SHEETS.keys()),
        default=list(SHEETS.keys()),
    )

    open_combined = []
    for name in sources:
        status_col = SHEETS[name]["status_col"]
        df_nc = filter_not_closed(data[name], status_col)
        if not df_nc.empty and "Assigned To" in df_nc.columns:
            open_combined.append(df_nc[["Assigned To", "Priority"]])

    df_open_all = pd.concat(open_combined) if open_combined else pd.DataFrame()
    assigned_to_bars_stacked_by_priority(df_open_all, "Assignees")

# -------- CLOSED --------
with tab_closed:
    c1, c2, c3 = st.columns(3)
    for col, name in zip([c1, c2, c3], SHEETS):
        with col:
            st.subheader(name)
            status_col = SHEETS[name]["status_col"]
            df_c = filter_closed(data[name], status_col)
            closed_pie_chart(df_c, "By priority")

    st.divider()

    sources = st.multiselect(
        "Assignees: include",
        options=list(SHEETS.keys()),
        default=list(SHEETS.keys()),
        key="closed_sources",
    )

    closed_combined = []
    for name in sources:
        status_col = SHEETS[name]["status_col"]
        df_c = filter_closed(data[name], status_col)
        if not df_c.empty and "Assigned To" in df_c.columns:
            closed_combined.append(df_c[["Assigned To", "Priority"]])

    df_closed_all = pd.concat(closed_combined) if closed_combined else pd.DataFrame()
    assigned_to_bars_stacked_by_priority(df_closed_all, "Assignees")
