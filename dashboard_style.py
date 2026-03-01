from __future__ import annotations

import streamlit as st

# ---------------------------------------------------------------------------
# Color Palette
# ---------------------------------------------------------------------------
COLOR_UP = "#CC0000"
COLOR_DOWN = "#0066CC"
COLOR_FLAT = "#888888"
COLOR_PRIMARY = "#2F5496"
COLOR_PRIMARY_LIGHT = "#4472C4"
COLOR_BG_SECTION = "#E9EDF4"
COLOR_BG_LIGHT = "#F8F9FA"
COLOR_BG_STRIPE = "#D6E4F0"
COLOR_BG_CARD = "#FFFFFF"
COLOR_BORDER = "#DEE2E6"
COLOR_TEXT_MUTED = "#888888"

CHART_TEMPLATE = "plotly_white"
CHART_COLORS = [COLOR_PRIMARY, "#E74C3C", "#27AE60", "#F39C12", "#8E44AD",
                "#1ABC9C", "#E67E22", "#3498DB", "#9B59B6", "#2ECC71"]

WEEKDAY_KR = ["월", "화", "수", "목", "금", "토", "일"]

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
CUSTOM_CSS = """
<style>
/* ── Global ── */
html, body, [data-testid="stAppViewContainer"] {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont,
                 'Segoe UI', Roboto, sans-serif;
    color: #1A1A2E;
}
[data-testid="stAppViewContainer"] > .main {
    background: #F7F8FC;
}

/* ── Page title (st.title) ── */
h1, h1 span, h1 div, h1 a {
    font-size: 1.75rem !important;
    font-weight: 700 !important;
    color: #1A1A2E !important;
    letter-spacing: -0.02em;
    padding-bottom: 0 !important;
    margin-bottom: 0 !important;
}

/* ── Page header block (page_header func) ── */
.page-header {
    margin-bottom: 20px;
}
.page-header h2,
.page-header h2 span,
.page-header h2 a {
    font-size: 1.65rem !important;
    font-weight: 700 !important;
    color: #1A1A2E !important;
    letter-spacing: -0.02em;
    margin: 0 0 2px 0;
    padding: 0;
    line-height: 1.3;
}
.page-header .page-subtitle {
    font-size: 0.8rem;
    color: #9CA3AF;
    margin: 0;
    padding: 0;
    font-weight: 400;
    letter-spacing: 0;
}

/* ── KPI Metric Cards ── */
div[data-testid="metric-container"] {
    background: #FFFFFF;
    padding: 12px 16px;
    border-radius: 8px;
    border-left: 3px solid #2F5496;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    transition: box-shadow 0.15s ease;
}
div[data-testid="metric-container"]:hover {
    box-shadow: 0 2px 10px rgba(0,0,0,0.09);
}
div[data-testid="metric-container"] label,
div[data-testid="metric-container"] label p,
div[data-testid="metric-container"] label div,
div[data-testid="metric-container"] label span {
    font-size: 0.72rem !important;
    color: #6B7280 !important;
    font-weight: 500 !important;
    text-transform: uppercase;
    letter-spacing: 0.04em;
}
div[data-testid="metric-container"] [data-testid="stMetricValue"],
div[data-testid="metric-container"] [data-testid="stMetricValue"] > div,
div[data-testid="metric-container"] [data-testid="stMetricValue"] > div > div,
div[data-testid="metric-container"] [data-testid="stMetricValue"] div,
div[data-testid="metric-container"] [data-testid="stMetricValue"] p,
div[data-testid="metric-container"] [data-testid="stMetricValue"] span {
    font-size: 0.95rem !important;
    font-weight: 700 !important;
    color: #111827 !important;
}
div[data-testid="metric-container"] [data-testid="stMetricDelta"],
div[data-testid="metric-container"] [data-testid="stMetricDelta"] div,
div[data-testid="metric-container"] [data-testid="stMetricDelta"] p,
div[data-testid="metric-container"] [data-testid="stMetricDelta"] span,
div[data-testid="metric-container"] [data-testid="stMetricDelta"] svg {
    font-size: 0.72rem !important;
}

/* ── Section Headers (dark professional) ── */
.section-header {
    background: linear-gradient(135deg, #1F3864 0%, #2F5496 100%);
    padding: 10px 16px;
    border-radius: 6px;
    margin: 20px 0 14px 0;
    font-weight: 700;
    color: #FFFFFF;
    font-size: 0.88rem;
    letter-spacing: -0.01em;
    border-left: none;
    box-shadow: 0 2px 4px rgba(31,56,100,0.15);
}

/* ── Sub-section labels ── */
.subsection-label {
    font-size: 0.72rem;
    color: #6B7280;
    font-weight: 600;
    margin: 12px 0 6px 2px;
    letter-spacing: 0.03em;
    padding-left: 6px;
    border-left: 2px solid #D1D5DB;
}

/* ── Indicator Cards (finance app style) ── */
.indicator-card {
    background: #FFFFFF;
    border-radius: 8px;
    overflow: hidden;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04);
    transition: box-shadow 0.15s ease, transform 0.15s ease;
}
.indicator-card:hover {
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    transform: translateY(-1px);
}
.indicator-card__border {
    height: 3px;
    width: 100%;
}
.indicator-card__body {
    padding: 14px 12px 12px 12px;
    text-align: center;
}
.indicator-card__name {
    font-size: 0.7rem;
    color: #6B7280;
    font-weight: 600;
    letter-spacing: 0.03em;
    margin-bottom: 6px;
    text-transform: uppercase;
}
.indicator-card__value {
    font-size: 1.15rem;
    font-weight: 800;
    color: #111827;
    margin-bottom: 4px;
    letter-spacing: -0.02em;
    font-variant-numeric: tabular-nums;
}
.indicator-card__change {
    font-size: 0.73rem;
    font-weight: 600;
    font-variant-numeric: tabular-nums;
}
.indicator-card__prev {
    font-size: 0.65rem;
    color: #9CA3AF;
    margin-top: 4px;
}

/* ── Source attribution ── */
.source-text {
    color: #9CA3AF;
    font-size: 0.75rem;
    text-align: right;
    margin-top: 8px;
    padding-right: 4px;
}

/* ── Summary table ── */
.summary-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.82rem;
    margin-top: 8px;
    background: #FFF;
    border-radius: 6px;
    overflow: hidden;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}
.summary-table thead tr {
    background: #F8F9FA;
    border-bottom: 2px solid #E5E7EB;
}
.summary-table th {
    padding: 8px 10px;
    color: #6B7280;
    font-weight: 600;
    font-size: 0.75rem;
}
.summary-table td {
    padding: 6px 10px;
    font-variant-numeric: tabular-nums;
}
.summary-table tbody tr {
    border-bottom: 1px solid #F0F0F0;
}
.summary-table tbody tr:last-child {
    border-bottom: none;
}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background: #F5F6FA;
    border-right: 1px solid #E5E7EB;
}
section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
    padding-top: 1.2rem;
}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h1 span,
section[data-testid="stSidebar"] h1 a {
    font-size: 1.1rem !important;
    color: #1F3864 !important;
}

/* ── Tabs ── */
button[data-baseweb="tab"] {
    font-size: 0.85rem !important;
    font-weight: 600 !important;
    padding: 8px 16px !important;
    color: #6B7280 !important;
    border-bottom: 2px solid transparent !important;
    transition: color 0.15s ease;
}
button[data-baseweb="tab"]:hover {
    color: #374151 !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #2F5496 !important;
    border-bottom-color: #2F5496 !important;
}

/* ── DataFrames ── */
.stDataFrame {
    font-size: 13px !important;
    border-radius: 4px;
    overflow: hidden;
    border: 1px solid #E5E7EB;
}

/* ── Containers with border ── */
div[data-testid="stVerticalBlock"] > div[data-testid="element-container"] [data-testid="stContainer"] {
    border-color: #E5E7EB !important;
    border-radius: 6px;
}

/* ── Expander ── */
details[data-testid="stExpander"] summary {
    font-weight: 600;
    color: #374151;
    font-size: 0.88rem;
}

/* ── Buttons ── */
button[kind="secondary"] {
    border-radius: 6px !important;
    font-weight: 500 !important;
    font-size: 0.82rem !important;
}

/* ── Hide Streamlit branding ── */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header[data-testid="stHeader"] {background: transparent;}

/* ── Colored text helpers ── */
.up { color: #CC0000; font-weight: 600; }
.down { color: #0066CC; font-weight: 600; }
.flat { color: #9CA3AF; }
.muted { color: #9CA3AF; font-size: 0.8rem; }

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #D1D5DB; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #9CA3AF; }

/* ── Divider ── */
hr { border: none; border-top: 1px solid #E5E7EB; margin: 20px 0; }

/* ── Captions ── */
.stCaption, [data-testid="stCaptionContainer"] {
    color: #9CA3AF !important;
    font-size: 0.78rem !important;
}

/* ── Warning/info boxes ── */
div[data-testid="stAlert"] {
    font-size: 0.85rem;
    border-radius: 6px;
}

/* ── Professional Index Cards (KOSPI/KOSDAQ) ── */
.pro-index {
    background: #FFFFFF;
    border-radius: 10px;
    overflow: hidden;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
    border-top: 4px solid #2F5496;
    transition: box-shadow 0.15s ease, transform 0.15s ease;
}
.pro-index:hover {
    box-shadow: 0 4px 16px rgba(0,0,0,0.12);
    transform: translateY(-2px);
}
.pro-index__header {
    padding: 16px 20px 4px;
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
}
.pro-index__info {
    flex: 1;
    min-width: 0;
}
.pro-index__name {
    font-size: 0.75rem;
    color: #6B7280;
    font-weight: 600;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    margin-bottom: 6px;
}
.pro-index__value {
    font-size: 1.8rem;
    font-weight: 800;
    letter-spacing: -0.02em;
    font-variant-numeric: tabular-nums;
    line-height: 1.2;
}
.pro-index__change {
    font-size: 0.95rem;
    font-weight: 700;
    font-variant-numeric: tabular-nums;
    margin-top: 4px;
    margin-bottom: 4px;
}
.pro-index__ranges {
    padding: 8px 20px 6px;
}
.pro-index__range-row {
    margin-bottom: 8px;
}
.pro-index__range-row:last-child {
    margin-bottom: 0;
}
.pro-index__range-meta {
    display: flex;
    align-items: center;
    margin-bottom: 4px;
}
.pro-index__range-label {
    font-size: 0.6rem;
    color: #9CA3AF;
    font-weight: 600;
    width: 28px;
    flex-shrink: 0;
}
.pro-index__range-labels {
    display: flex;
    justify-content: space-between;
    font-size: 0.6rem;
    color: #9CA3AF;
    font-weight: 500;
    flex: 1;
}
.pro-index__range-track {
    position: relative;
    height: 5px;
    background: #E5E7EB;
    border-radius: 3px;
    overflow: visible;
    margin-left: 28px;
}
.pro-index__range-fill {
    position: absolute;
    top: 0;
    left: 0;
    height: 100%;
    border-radius: 3px;
    opacity: 0.7;
}
.pro-index__range-dot {
    position: absolute;
    top: 50%;
    transform: translate(-50%, -50%);
    width: 11px;
    height: 11px;
    border-radius: 50%;
    border: 2px solid #FFFFFF;
    box-shadow: 0 1px 3px rgba(0,0,0,0.25);
}
.pro-index__prev {
    text-align: center;
    font-size: 0.68rem;
    color: #9CA3AF;
    padding: 6px 20px 12px;
    font-weight: 500;
    font-variant-numeric: tabular-nums;
}
</style>
"""


def inject_css() -> None:
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Formatting Helpers
# ---------------------------------------------------------------------------

def fmt_eok(v: int | float) -> str:
    if v == 0:
        return "-"
    return f"{int(round(v)):,}"


def fmt_jo(v: int | float) -> str:
    jo = v / 1_0000_0000_0000
    return f"{jo:,.1f}조"


def fmt_pct(v: float | None) -> str:
    if v is None:
        return "-"
    sign = "+" if v > 0 else ""
    return f"{sign}{v:.2f}%"


def fmt_ratio(v: float | None) -> str:
    if v is None:
        return "-"
    return f"{v:.1f}"


def color_value(v: float | int | None) -> str:
    if v is None or v == 0:
        return "flat"
    return "up" if v > 0 else "down"


def color_pct_html(v: float | None) -> str:
    if v is None:
        return '<span class="flat">-</span>'
    cls = color_value(v)
    sign = "+" if v > 0 else ""
    return f'<span class="{cls}">{sign}{v:.2f}%</span>'


def section_header(title: str) -> None:
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)


def subsection_label(title: str) -> None:
    st.markdown(f'<p class="subsection-label">{title}</p>', unsafe_allow_html=True)


def page_header(title: str, subtitle: str = "") -> None:
    """Render a unified page header with optional subtitle line."""
    sub_html = f'<p class="page-subtitle">{subtitle}</p>' if subtitle else ""
    st.markdown(
        f'<div class="page-header"><h2>{title}</h2>{sub_html}</div>',
        unsafe_allow_html=True,
    )


def title_date_text(d) -> str:
    return f"{d:%Y/%m/%d}({WEEKDAY_KR[d.weekday()]})"
