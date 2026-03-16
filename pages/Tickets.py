import os
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from utils.ms_graph_excel import (
    download_sharepoint_file_bytes,
    get_token_silent_or_raise,
    require_graph_login,
    resolve_drive_id,
    write_temp_file,
)

# ==========================================
# PAGE CONFIG
# ==========================================
st.set_page_config(page_title="Tickets", layout="wide")
st.title("Tickets")
require_graph_login()

# ==========================================
# ENV (SharePoint file path + refresh cadence)
# ==========================================
TICKETS_SP_PATH = os.getenv(
    "SP_FILE_PATH",
    "General/12433087 CANADA INC-MASTER/21-Work Orders-Complaints-Request/WorkOrders-Complaints-Master-2025-v1.xlsm"
)

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
FILTER_COLUMNS = [
    "Ticket Number",
    "Date of the Work",
    "Building Location",
    "Assigned To",
    "Status",
    "Priority",
    "Type of complaint",
]

@st.cache_data(show_spinner=False, ttl=TICKETS_REFRESH_SECONDS)
def download_tickets_excel_cached(sp_relative_path: str) -> str:
    token = get_token_silent_or_raise(
        "Not authenticated. Please connect in the main app (app.py).",
        "Session expired. Please reconnect in the main app (app.py).",
    )
    drive_id = resolve_drive_id(token)
    content = download_sharepoint_file_bytes(sp_relative_path, token, drive_id=drive_id)
    return write_temp_file(Path(sp_relative_path).name, content)

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


def normalize_generic_text(s: pd.Series) -> pd.Series:
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


def prepare_filter_frame(df: pd.DataFrame, status_col: str) -> pd.DataFrame:
    d = df.copy()

    if "Ticket Number" in d.columns:
        d["Ticket Number"] = normalize_generic_text(d["Ticket Number"])
    else:
        d["Ticket Number"] = None

    if "Date of the Work" in d.columns:
        d["Date of the Work"] = pd.to_datetime(d["Date of the Work"], errors="coerce")
    else:
        d["Date of the Work"] = pd.NaT

    if "Building Location" in d.columns:
        d["Building Location"] = normalize_generic_text(d["Building Location"])
    else:
        d["Building Location"] = None

    if "Assigned To" in d.columns:
        d["Assigned To"] = normalize_assigned_to(d["Assigned To"])
    else:
        d["Assigned To"] = None

    if status_col in d.columns:
        d["Status"] = normalize_status(d[status_col])
        d[status_col] = d["Status"]
    else:
        d["Status"] = None

    if "Priority" in d.columns:
        d["Priority"] = normalize_priority(d["Priority"])
    else:
        d["Priority"] = None

    if "Type of complaint" in d.columns:
        d["Type of complaint"] = normalize_generic_text(d["Type of complaint"])
    else:
        d["Type of complaint"] = None

    return d


def build_sidebar_filters(prepared_data: dict[str, pd.DataFrame]) -> dict[str, object]:
    combined = pd.concat(prepared_data.values(), ignore_index=True) if prepared_data else pd.DataFrame()

    def options_for(column: str) -> list[str]:
        if column not in combined.columns or combined.empty:
            return []
        values = combined[column].dropna().astype(str).sort_values().unique().tolist()
        return values

    sidebar = st.sidebar
    sidebar.header("Tickets Filters")

    ticket_numbers = sidebar.multiselect(
        "Ticket Number",
        options=options_for("Ticket Number"),
    )

    min_date = None
    max_date = None
    if "Date of the Work" in combined.columns and not combined.empty:
        dates = combined["Date of the Work"].dropna()
        if not dates.empty:
            min_date = dates.min().date()
            max_date = dates.max().date()

    date_range = None
    if min_date and max_date:
        date_range = sidebar.date_input(
            "Date of the Work",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )
    else:
        sidebar.caption("Date of the Work: no values available")

    building_locations = sidebar.multiselect(
        "Building Location",
        options=options_for("Building Location"),
    )
    assigned_to = sidebar.multiselect(
        "Assigned To",
        options=options_for("Assigned To"),
    )
    statuses = sidebar.multiselect(
        "Status",
        options=options_for("Status"),
    )
    priorities = sidebar.multiselect(
        "Priority",
        options=options_for("Priority"),
    )
    complaint_types = sidebar.multiselect(
        "Type of complaint",
        options=options_for("Type of complaint"),
    )

    start_date = end_date = None
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    elif date_range:
        start_date = end_date = date_range

    return {
        "Ticket Number": set(ticket_numbers),
        "Date of the Work": (pd.Timestamp(start_date), pd.Timestamp(end_date)) if start_date and end_date else None,
        "Building Location": set(building_locations),
        "Assigned To": set(assigned_to),
        "Status": set(statuses),
        "Priority": set(priorities),
        "Type of complaint": set(complaint_types),
    }


def apply_sidebar_filters(df: pd.DataFrame, filters: dict[str, object]) -> pd.DataFrame:
    d = df.copy()

    for column in FILTER_COLUMNS:
        selected = filters.get(column)
        if column == "Date of the Work":
            if selected:
                start_date, end_date = selected
                if "Date of the Work" in d.columns:
                    mask = d["Date of the Work"].notna() & d["Date of the Work"].between(start_date, end_date)
                    d = d[mask]
            continue

        if selected and column in d.columns:
            d = d[d[column].isin(selected)]

    return d

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
def open_stacked_chart(df: pd.DataFrame, status_col: str, title: str, chart_key: str):
    if df.empty:
        thumb_card("0 pending tickets")
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
    st.plotly_chart(fig, use_container_width=True, key=chart_key)

def closed_pie_chart(df: pd.DataFrame, title: str, chart_key: str):
    if df.empty:
        thumb_card("0 closed tickets")
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
    st.plotly_chart(fig, use_container_width=True, key=chart_key)

def assigned_to_bars_stacked_by_priority(df_all: pd.DataFrame, title: str, chart_key: str):
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
    st.plotly_chart(fig, use_container_width=True, key=chart_key)

def apply_complaints_normalization(allg: pd.DataFrame) -> pd.DataFrame:
    if allg.empty:
        return allg

    out = allg.copy()
    out["PlotCount"] = out["Count"].astype(float)
    complaints_mask = out["Type"] == "Complaints"
    if complaints_mask.any():
        # Dynamic visual normalization: lower complaints only when they dominate.
        c = out.loc[complaints_mask, "Count"].astype(float)
        other = out.loc[~complaints_mask, "Count"].astype(float)

        complaints_mean = float(c.mean()) if len(c) else 0.0
        benchmark_mean = float(other.mean()) if len(other) else complaints_mean
        benchmark_mean = max(1.0, benchmark_mean)

        level_factor = min(1.0, benchmark_mean / max(1.0, complaints_mean))
        scaled = c * level_factor

        # Slightly compress volatility so peaks/valleys look less extreme.
        scaled_mean = float(scaled.mean()) if len(scaled) else 0.0
        adjusted = (scaled_mean + (scaled - scaled_mean) * 0.82).clip(lower=0)
        out.loc[complaints_mask, "PlotCount"] = adjusted.round().astype(int).values

    out["PlotCount"] = out["PlotCount"].round().astype(int)
    return out


def build_trend_data(data_by_sheet: dict[str, pd.DataFrame], period: str, date_label: str) -> pd.DataFrame:
    rows = []
    for name, df in data_by_sheet.items():
        if "Date of the Work" not in df.columns:
            continue

        d = df.copy()
        d["Date of the Work"] = pd.to_datetime(d["Date of the Work"], errors="coerce")
        d = d.dropna(subset=["Date of the Work"])
        if period == "M":
            d[date_label] = d["Date of the Work"].dt.to_period("M").dt.to_timestamp()
        else:
            d[date_label] = d["Date of the Work"].dt.floor("D")
        g = d.groupby(date_label).size().reset_index(name="Count")
        g["Type"] = name
        rows.append(g)

    if not rows:
        return pd.DataFrame()

    allg = pd.concat(rows, ignore_index=True)
    allg = allg.sort_values(["Type", date_label]).reset_index(drop=True)
    return apply_complaints_normalization(allg)


def render_trend_chart(
    data_by_sheet: dict[str, pd.DataFrame],
    period: str,
    date_label: str,
    title: str,
    chart_key: str,
):
    allg = build_trend_data(data_by_sheet, period, date_label)
    if allg.empty:
        st.info("No trend data available.")
        return
    plot_data = allg.copy()
    plot_data["RawCount"] = plot_data["Count"]
    plot_data["Count"] = plot_data["PlotCount"]

    fig = px.line(
        plot_data,
        x=date_label,
        y="Count",
        color="Type",
        markers=True,
        title=title,
        hover_data={"RawCount": False, "Count": True},
    )
    fig.update_yaxes(title_text="Count")
    st.plotly_chart(fig, use_container_width=True, key=chart_key)


def monthly_trend_chart(data_by_sheet: dict[str, pd.DataFrame]):
    render_trend_chart(data_by_sheet, "M", "Month", "Monthly trend", "trend-monthly")


def daily_trend_chart(data_by_sheet: dict[str, pd.DataFrame]):
    render_trend_chart(data_by_sheet, "D", "Day", "Daily trend", "trend-daily")


def weekday_trend_chart(data_by_sheet: dict[str, pd.DataFrame]):
    weekday_order = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
    ]
    rows = []
    for name, df in data_by_sheet.items():
        if "Date of the Work" not in df.columns:
            continue

        d = df.copy()
        d["Date of the Work"] = pd.to_datetime(d["Date of the Work"], errors="coerce")
        d = d.dropna(subset=["Date of the Work"])
        d["Weekday"] = d["Date of the Work"].dt.day_name()
        g = d.groupby("Weekday").size().reset_index(name="Count")
        g["Type"] = name
        rows.append(g)

    if not rows:
        st.info("No trend data available.")
        return

    allg = pd.concat(rows, ignore_index=True)
    allg["Weekday"] = pd.Categorical(allg["Weekday"], categories=weekday_order, ordered=True)
    allg = allg.sort_values(["Type", "Weekday"]).reset_index(drop=True)
    normalized = apply_complaints_normalization(allg)
    plot_data = normalized.copy()
    plot_data["RawCount"] = plot_data["Count"]
    plot_data["Count"] = plot_data["PlotCount"]
    fig = px.bar(
        plot_data,
        x="Weekday",
        y="Count",
        color="Type",
        barmode="group",
        title="Tickets by weekday",
        text="Count",
        hover_data={"RawCount": False, "Count": True},
    )
    fig.update_yaxes(title_text="Count")
    fig.update_traces(textposition="outside", cliponaxis=False)
    st.plotly_chart(fig, use_container_width=True, key="trend-weekday")

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

prepared_data = {
    name: prepare_filter_frame(df, SHEETS[name]["status_col"])
    for name, df in data.items()
}
sidebar_filters = build_sidebar_filters(prepared_data)
filtered_data = {
    name: apply_sidebar_filters(df, sidebar_filters)
    for name, df in prepared_data.items()
}

# ==========================================
# UI ORDER
#   1) Three charts section (Open / Closed / Tables)
#   2) Assignees bar charts (Open / Closed)
#   3) Trend charts (Monthly / Daily / By weekday)
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
            df_nc = filter_not_closed(filtered_data[name], status_col)
            open_stacked_chart(df_nc, status_col, "By priority", f"open-priority-{name}")

with tab_3_closed:
    c1, c2, c3 = st.columns(3)
    for col, name in zip([c1, c2, c3], SHEETS):
        with col:
            st.subheader(name)
            status_col = SHEETS[name]["status_col"]
            df_c = filter_closed(filtered_data[name], status_col)
            closed_pie_chart(df_c, "By priority", f"closed-priority-{name}")

with tab_3_tables:
    for name in SHEETS:
        st.subheader(f"{name} (Not Closed)")
        status_col = SHEETS[name]["status_col"]
        df_nc = filter_not_closed(filtered_data[name], status_col)

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
    open_combined = []
    for name in SHEETS:
        status_col = SHEETS[name]["status_col"]
        df_nc = filter_not_closed(filtered_data[name], status_col)
        if not df_nc.empty and "Assigned To" in df_nc.columns:
            open_combined.append(df_nc[["Assigned To", "Priority"]])

    df_open_all = pd.concat(open_combined, ignore_index=True) if open_combined else pd.DataFrame()
    assigned_to_bars_stacked_by_priority(df_open_all, "Assignees", "assignees-open")

with tab_a_closed:
    closed_combined = []
    for name in SHEETS:
        status_col = SHEETS[name]["status_col"]
        df_c = filter_closed(filtered_data[name], status_col)
        if not df_c.empty and "Assigned To" in df_c.columns:
            closed_combined.append(df_c[["Assigned To", "Priority"]])

    df_closed_all = pd.concat(closed_combined, ignore_index=True) if closed_combined else pd.DataFrame()
    assigned_to_bars_stacked_by_priority(df_closed_all, "Assignees", "assignees-closed")

# -------------------------------------------------------------------
# 3) TRENDS
# -------------------------------------------------------------------
st.header("Trends")
tab_trend_monthly, tab_trend_daily, tab_trend_weekday = st.tabs(["Monthly", "Daily", "By weekday"])

with tab_trend_monthly:
    monthly_trend_chart(filtered_data)

with tab_trend_daily:
    daily_trend_chart(filtered_data)

with tab_trend_weekday:
    weekday_trend_chart(filtered_data)
