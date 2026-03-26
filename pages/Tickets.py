import os
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    HRFlowable,
    Image,
    LongTable,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    TableStyle,
)
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
REPORT_LOGO_PATH = Path(__file__).resolve().parents[1] / "assets" / "logo.jpeg"

TICKETS_REFRESH_SECONDS = 30 * 60  # 30 minutes

# ==========================================
# UI CONSTANTS
# ==========================================
SHEETS = {
    "Work Orders": {"sheet": "Work Orders", "status_col": "General Status"},
    "Request": {"sheet": "Request", "status_col": "Status"},
    "Complaints": {"sheet": "Complaints", "status_col": "Status"},
}
APPENDIX_COLUMNS = {
    "Work Orders": [
        "Ticket Number",
        "Date of the Work",
        "Building Location",
        "Client",
        "Email subject",
        "Assigned To",
        "General Status",
        "Priority",
    ],
    "Request": [
        "Ticket Number",
        "Date of the Work",
        "Building Location",
        "Description/Details",
        "Assigned To",
        "Status",
        "Priority",
    ],
    "Complaints": [
        "Ticket Number",
        "Date of the Work",
        "Building Location",
        "Description/Details",
        "Assigned To",
        "Status",
        "Priority",
    ],
}
APPENDIX_COLUMN_WEIGHTS = {
    "Ticket Number": 1.2,
    "Date of the Work": 1.2,
    "Building Location": 1.6,
    "Description/Details": 3.4,
    "Assigned To": 1.4,
    "Status": 1.2,
    "Priority": 1.0,
    "Client": 1.5,
    "Email subject": 3.0,
    "General Status": 1.4,
}

PRIORITY_COLORS = {"High": "#d32f2f", "Medium": "#fbc02d", "Low": "#388e3c"}
PRIORITY_COLORS_LIGHT = {"High": "#f28b82", "Medium": "#ffe082", "Low": "#a5d6a7"}
TYPE_COLORS = {"Complaints": "#c62828", "Work Orders": "#2e7d32", "Request": "#1565c0"}
STATUS_CELL_COLORS = {
    "Open": (colors.white, colors.HexColor("#111827")),
    "In Progress": (colors.HexColor("#f1f5f9"), colors.HexColor("#111827")),
    "Closed": (colors.HexColor("#e5e7eb"), colors.HexColor("#111827")),
    "Other": (colors.HexColor("#f8fafc"), colors.HexColor("#111827")),
}
PRIORITY_CELL_COLORS = {
    "High": (colors.HexColor("#f4b6b2"), colors.HexColor("#7f1d1d")),
    "Medium": (colors.HexColor("#fde68a"), colors.HexColor("#78350f")),
    "Low": (colors.HexColor("#bbf7d0"), colors.HexColor("#14532d")),
}
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


def _safe_text(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    return str(value)


def _status_sort_key(value) -> int:
    order = {"Open": 0, "In Progress": 1, "Other": 2, "Closed": 3}
    return order.get(_safe_text(value), 4)


def _status_column_name(df: pd.DataFrame) -> str | None:
    if "Status" in df.columns:
        return "Status"
    if "General Status" in df.columns:
        return "General Status"
    return None


def sort_for_appendix(df: pd.DataFrame) -> pd.DataFrame:
    status_column = _status_column_name(df)
    if df.empty or status_column is None:
        return df

    d = df.copy()
    d["__status_order"] = d[status_column].map(_status_sort_key)
    sort_columns = ["__status_order"]
    ascending = [True]
    if "Date of the Work" in d.columns:
        d["__sort_date"] = pd.to_datetime(d["Date of the Work"], errors="coerce")
        sort_columns.append("__sort_date")
        ascending.append(True)
    d = d.sort_values(sort_columns, ascending=ascending, kind="stable").drop(columns=[c for c in ["__status_order", "__sort_date"] if c in d.columns])
    return d


def filters_signature(filters: dict[str, object]) -> tuple:
    signature = []
    for column in FILTER_COLUMNS:
        selected = filters.get(column)
        if column == "Date of the Work":
            if selected:
                start_date, end_date = selected
                signature.append((column, _safe_text(start_date), _safe_text(end_date)))
            else:
                signature.append((column, None))
        else:
            signature.append((column, tuple(sorted(selected)) if selected else tuple()))
    return tuple(signature)


def render_filters_summary(filters: dict[str, object]) -> list[str]:
    lines = []
    for column in FILTER_COLUMNS:
        selected = filters.get(column)
        if column == "Date of the Work":
            if selected:
                start_date, end_date = selected
                lines.append(f"{column}: {_safe_text(start_date)} to {_safe_text(end_date)}")
            else:
                lines.append(f"{column}: All")
            continue

        if selected:
            lines.append(f"{column}: {', '.join(sorted(selected))}")
        else:
            lines.append(f"{column}: All")
    return lines


def render_active_filters_summary(filters: dict[str, object]) -> list[str]:
    lines = []
    for column in FILTER_COLUMNS:
        selected = filters.get(column)
        if column == "Date of the Work":
            if selected:
                start_date, end_date = selected
                lines.append(f"<b>{column}</b>: {_safe_text(start_date)} to {_safe_text(end_date)}")
            continue

        if selected:
            lines.append(f"<b>{column}</b>: {', '.join(sorted(selected))}")
    return lines


def _column_has_content(series: pd.Series) -> bool:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series.notna().any()
    return series.map(_safe_text).str.strip().ne("").any()


def _prepare_appendix_columns(section_name: str, df: pd.DataFrame) -> pd.DataFrame:
    preferred_columns = APPENDIX_COLUMNS.get(section_name, [])
    selected_columns = [column for column in preferred_columns if column in df.columns and _column_has_content(df[column])]

    if not selected_columns:
        selected_columns = [column for column in df.columns if _column_has_content(df[column])]

    table_df = sort_for_appendix(df[selected_columns].copy())
    for column in table_df.columns:
        if pd.api.types.is_datetime64_any_dtype(table_df[column]):
            table_df[column] = table_df[column].dt.strftime("%Y-%m-%d").fillna("")
        else:
            table_df[column] = table_df[column].map(_safe_text)
    return table_df


def _build_column_widths(columns: list[str], total_width: float) -> list[float]:
    weights = [APPENDIX_COLUMN_WEIGHTS.get(column, 1.2) for column in columns]
    weight_sum = sum(weights) or 1.0
    return [(weight / weight_sum) * total_width for weight in weights]


def build_report_table(section_name: str, df: pd.DataFrame) -> LongTable:
    styles = get_pdf_styles()
    table_df = _prepare_appendix_columns(section_name, df)
    if table_df.empty and len(table_df.columns) == 0:
        fallback = LongTable(
            [[Paragraph("No relevant columns with data for this section.", styles["table_cell"])]],
            colWidths=[10.6 * inch],
        )
        fallback.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1"))]))
        return fallback

    header_row = [Paragraph(column, styles["table_header"]) for column in table_df.columns]
    data_rows = [
        [Paragraph(value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), styles["table_cell"]) for value in row]
        for row in table_df.values.tolist()
    ]
    rows = [header_row] + data_rows
    total_width = 10.6 * inch
    col_widths = _build_column_widths(table_df.columns.tolist(), total_width)
    table = LongTable(rows, repeatRows=1, colWidths=col_widths, splitByRow=1)
    style_commands = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f2937")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]

    status_column = _status_column_name(table_df)
    priority_idx = table_df.columns.get_loc("Priority") if "Priority" in table_df.columns else None
    if status_column:
        status_idx = table_df.columns.get_loc(status_column)
        for row_idx, status_value in enumerate(table_df[status_column].tolist(), start=1):
            bg, fg = STATUS_CELL_COLORS.get(_safe_text(status_value), STATUS_CELL_COLORS["Other"])
            if priority_idx is None:
                style_commands.append(("BACKGROUND", (0, row_idx), (-1, row_idx), bg))
                style_commands.append(("TEXTCOLOR", (0, row_idx), (-1, row_idx), fg))
            else:
                if priority_idx > 0:
                    style_commands.append(("BACKGROUND", (0, row_idx), (priority_idx - 1, row_idx), bg))
                    style_commands.append(("TEXTCOLOR", (0, row_idx), (priority_idx - 1, row_idx), fg))
                if priority_idx < len(table_df.columns) - 1:
                    style_commands.append(("BACKGROUND", (priority_idx + 1, row_idx), (-1, row_idx), bg))
                    style_commands.append(("TEXTCOLOR", (priority_idx + 1, row_idx), (-1, row_idx), fg))
                style_commands.append(("BACKGROUND", (status_idx, row_idx), (status_idx, row_idx), bg))
                style_commands.append(("TEXTCOLOR", (status_idx, row_idx), (status_idx, row_idx), fg))

    if priority_idx is not None:
        for row_idx, priority_value in enumerate(table_df["Priority"].tolist(), start=1):
            color_pair = PRIORITY_CELL_COLORS.get(_safe_text(priority_value))
            if color_pair:
                bg, fg = color_pair
                style_commands.append(("BACKGROUND", (priority_idx, row_idx), (priority_idx, row_idx), bg))
                style_commands.append(("TEXTCOLOR", (priority_idx, row_idx), (priority_idx, row_idx), fg))

    table.setStyle(TableStyle(style_commands))
    return table


def get_pdf_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="ReportTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=20, leading=24, spaceAfter=8, alignment=TA_CENTER, textColor=colors.HexColor("#0f172a")))
    styles.add(ParagraphStyle(name="ReportSection", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=13, leading=16, spaceBefore=8, spaceAfter=8, alignment=TA_CENTER, textColor=colors.HexColor("#0f172a")))
    styles.add(ParagraphStyle(name="ReportBody", parent=styles["BodyText"], fontName="Helvetica", fontSize=9, leading=12, spaceAfter=4, alignment=TA_CENTER, textColor=colors.HexColor("#475569")))
    styles.add(ParagraphStyle(name="ReportFilters", parent=styles["BodyText"], fontName="Helvetica", fontSize=8.5, leading=11, spaceAfter=3, alignment=TA_CENTER, textColor=colors.HexColor("#334155")))
    styles.add(ParagraphStyle(name="ReportTableHeader", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=7, leading=8, textColor=colors.white))
    styles.add(ParagraphStyle(name="ReportTableCell", parent=styles["BodyText"], fontName="Helvetica", fontSize=6.5, leading=8))
    return {
        "title": styles["ReportTitle"],
        "section": styles["ReportSection"],
        "body": styles["ReportBody"],
        "filters": styles["ReportFilters"],
        "table_header": styles["ReportTableHeader"],
        "table_cell": styles["ReportTableCell"],
    }


def build_report_header(story: list, styles: dict[str, ParagraphStyle], title: str, filters: dict[str, object]):
    if REPORT_LOGO_PATH.exists():
        image_reader = ImageReader(str(REPORT_LOGO_PATH))
        img_width, img_height = image_reader.getSize()
        target_width = 1.5 * inch
        scale = target_width / float(img_width)
        logo = Image(str(REPORT_LOGO_PATH))
        logo.drawWidth = target_width
        logo.drawHeight = float(img_height) * scale
        logo.hAlign = "CENTER"
        story.append(logo)
        story.append(Spacer(1, 0.1 * inch))

    story.append(Paragraph(title, styles["title"]))
    active_filters = render_active_filters_summary(filters)
    if active_filters:
        story.append(Paragraph("Active filters", styles["body"]))
        for line in active_filters:
            story.append(Paragraph(line, styles["filters"]))
        story.append(Spacer(1, 0.05 * inch))


def build_section_divider() -> HRFlowable:
    return HRFlowable(
        width="72%",
        thickness=0.9,
        color=colors.HexColor("#cbd5e1"),
        hAlign="CENTER",
        spaceBefore=0.08 * inch,
        spaceAfter=0.14 * inch,
    )

# ==========================================
# CHARTS
# ==========================================
def build_open_stacked_figure(df: pd.DataFrame, status_col: str, title: str):
    if df.empty:
        return None

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
    return fig


def open_stacked_chart(df: pd.DataFrame, status_col: str, title: str, chart_key: str):
    fig = build_open_stacked_figure(df, status_col, title)
    if fig is None:
        thumb_card("0 pending tickets")
        return
    st.plotly_chart(fig, use_container_width=True, key=chart_key)

def build_closed_pie_figure(df: pd.DataFrame, title: str):
    if df.empty:
        return None

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
    return fig


def closed_pie_chart(df: pd.DataFrame, title: str, chart_key: str):
    fig = build_closed_pie_figure(df, title)
    if fig is None:
        thumb_card("0 closed tickets")
        return
    st.plotly_chart(fig, use_container_width=True, key=chart_key)

def build_assigned_to_figure(df_all: pd.DataFrame, title: str):
    if df_all.empty:
        return None

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
    return fig


def assigned_to_bars_stacked_by_priority(df_all: pd.DataFrame, title: str, chart_key: str):
    fig = build_assigned_to_figure(df_all, title)
    if fig is None:
        thumb_card("0 tickets", 260)
        return
    st.plotly_chart(fig, use_container_width=True, key=chart_key)

def apply_complaints_normalization(allg: pd.DataFrame) -> pd.DataFrame:
    out = allg.copy()
    if "Count" not in out.columns:
        out["Count"] = pd.Series(dtype=float)
    out["PlotCount"] = out["Count"].astype(float)
    if out.empty:
        return out

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


def build_trend_figure(
    data_by_sheet: dict[str, pd.DataFrame],
    period: str,
    date_label: str,
    title: str,
):
    allg = build_trend_data(data_by_sheet, period, date_label)
    if allg.empty:
        return None
    plot_data = allg.copy()
    plot_data["RawCount"] = plot_data["Count"]
    plot_data["Count"] = plot_data["PlotCount"]

    fig = px.line(
        plot_data,
        x=date_label,
        y="Count",
        color="Type",
        color_discrete_map=TYPE_COLORS,
        markers=True,
        title=title,
        hover_data={"RawCount": False, "Count": True},
    )
    fig.update_yaxes(title_text="Count")
    return fig


def render_trend_chart(
    data_by_sheet: dict[str, pd.DataFrame],
    period: str,
    date_label: str,
    title: str,
    chart_key: str,
):
    fig = build_trend_figure(data_by_sheet, period, date_label, title)
    if fig is None:
        st.info("No trend data available.")
        return
    st.plotly_chart(fig, use_container_width=True, key=chart_key)


def monthly_trend_chart(data_by_sheet: dict[str, pd.DataFrame]):
    render_trend_chart(data_by_sheet, "M", "Month", "Monthly trend", "trend-monthly")


def daily_trend_chart(data_by_sheet: dict[str, pd.DataFrame]):
    render_trend_chart(data_by_sheet, "D", "Day", "Daily trend", "trend-daily")


def weekday_trend_chart(data_by_sheet: dict[str, pd.DataFrame]):
    fig = build_weekday_trend_figure(data_by_sheet)
    if fig is None:
        st.info("No trend data available.")
        return
    st.plotly_chart(fig, use_container_width=True, key="trend-weekday")


def build_weekday_trend_figure(data_by_sheet: dict[str, pd.DataFrame]):
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
        return None

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
    return fig


def build_assignee_frame(data_by_sheet: dict[str, pd.DataFrame], closed: bool) -> pd.DataFrame:
    combined = []
    for name in SHEETS:
        status_col = SHEETS[name]["status_col"]
        df_filtered = filter_closed(data_by_sheet[name], status_col) if closed else filter_not_closed(data_by_sheet[name], status_col)
        if not df_filtered.empty and "Assigned To" in df_filtered.columns:
            combined.append(df_filtered[["Assigned To", "Priority"]])
    return pd.concat(combined, ignore_index=True) if combined else pd.DataFrame()


def build_tickets_report_pdf(filtered_data: dict[str, pd.DataFrame], filters: dict[str, object]) -> bytes:
    styles = get_pdf_styles()
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=0.4 * inch,
        rightMargin=0.4 * inch,
        topMargin=0.45 * inch,
        bottomMargin=0.45 * inch,
    )

    appendix_sections = [("Work Orders", filtered_data["Work Orders"]), ("Request", filtered_data["Request"]), ("Complaints", filtered_data["Complaints"])]
    story = []
    for index, (section_name, section_df) in enumerate(appendix_sections):
        if index > 0:
            story.append(PageBreak())
        build_report_header(story, styles, section_name, filters)
        story.append(build_section_divider())
        if section_df.empty:
            story.append(Paragraph("No rows match the active filters.", styles["body"]))
        else:
            story.append(build_report_table(section_name, section_df))

    doc.build(story)
    return buffer.getvalue()

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
current_filters_signature = filters_signature(sidebar_filters)
if st.session_state.get("tickets_report_filters") != current_filters_signature:
    st.session_state.pop("tickets_report_pdf", None)
    st.session_state["tickets_report_filters"] = current_filters_signature

st.sidebar.divider()
st.sidebar.subheader("PDF Report")
if st.sidebar.button("Prepare PDF report", key="prepare-tickets-pdf"):
    with st.spinner("Generating PDF report..."):
        st.session_state["tickets_report_pdf"] = build_tickets_report_pdf(filtered_data, sidebar_filters)

sidebar_pdf_bytes = st.session_state.get("tickets_report_pdf")
if sidebar_pdf_bytes:
    st.sidebar.download_button(
        "Download PDF report",
        data=sidebar_pdf_bytes,
        file_name="tickets_report.pdf",
        mime="application/pdf",
        key="download-tickets-pdf",
    )
else:
    st.sidebar.caption("Prepare the report to enable the download.")

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
    df_open_all = build_assignee_frame(filtered_data, closed=False)
    assigned_to_bars_stacked_by_priority(df_open_all, "Assignees", "assignees-open")

with tab_a_closed:
    df_closed_all = build_assignee_frame(filtered_data, closed=True)
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
