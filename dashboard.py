"""Main entry point for the stock market dashboard.

Run with: streamlit run dashboard.py
"""
from __future__ import annotations

from datetime import date, timedelta

import streamlit as st

import plotly.graph_objects as go

from dashboard_data import load_market_overview, load_investor_trends
from dashboard_style import (
    inject_css, title_date_text, section_header, subsection_label, page_header,
    COLOR_UP, COLOR_DOWN, COLOR_FLAT, CHART_TEMPLATE,
)

st.set_page_config(
    page_title="주식 시장 대시보드",
    page_icon="◻",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_css()

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("Stock Dashboard")
    st.divider()

    today = date.today()
    # Default to most recent weekday
    if today.weekday() >= 5:  # Saturday or Sunday
        today = today - timedelta(days=today.weekday() - 4)

    selected_date = st.date_input(
        "조회 일자",
        value=today,
        max_value=date.today(),
        format="YYYY/MM/DD",
    )
    st.session_state["selected_date"] = selected_date
    st.session_state["date_str"] = selected_date.strftime("%Y%m%d")

    st.caption(f"선택: {title_date_text(selected_date)}")

    st.divider()

    if st.button("캐시 초기화", width="stretch"):
        st.cache_data.clear()
        st.toast("캐시가 초기화되었습니다.")

    st.divider()
    st.markdown(
        '<p style="color:#9CA3AF; font-size:0.75rem; margin:0;">'
        'Data: KRX &middot; DART OpenAPI &middot; Yahoo Finance</p>',
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# Home Page Content
# ---------------------------------------------------------------------------
page_header("주식 시장 대시보드", f"{title_date_text(selected_date)} 기준")

# ---------------------------------------------------------------------------
# Market Indices
# ---------------------------------------------------------------------------
overview = load_market_overview()
indices = overview.get("indices", [])
macro = overview.get("macro", [])


def _color_class(pct: float) -> str:
    if pct > 0:
        return "up"
    elif pct < 0:
        return "down"
    return "flat"


def _fmt_value(name: str, val: float) -> str:
    """Format value based on indicator type."""
    if name in ("KOSPI", "KOSDAQ", "NASDAQ", "S&P 500", "다우존스", "니케이225"):
        return f"{val:,.2f}"
    elif name in ("USD/KRW", "USD/JPY"):
        return f"{val:,.2f}"
    elif name in ("금(oz)", "WTI유"):
        return f"${val:,.2f}"
    elif name == "비트코인":
        return f"${val:,.0f}"
    elif name == "VIX":
        return f"{val:.2f}"
    elif name == "미국10Y":
        return f"{val:.3f}%"
    return f"{val:,.2f}"


def _fmt_change(name: str, change: float, pct: float) -> str:
    """Format change string."""
    sign = "+" if change >= 0 else ""
    if name == "미국10Y":
        return f"{sign}{change:.3f} ({sign}{pct:.2f}%)"
    elif name in ("금(oz)", "WTI유"):
        return f"{sign}${change:.2f} ({sign}{pct:.2f}%)"
    elif name == "비트코인":
        return f"{sign}${change:,.0f} ({sign}{pct:.2f}%)"
    return f"{sign}{change:,.2f} ({sign}{pct:.2f}%)"


def _render_indicator_card(item: dict) -> str:
    """Render a single indicator as a finance-app-style card."""
    name = item["name"]
    val = item["value"]
    change = item["change"]
    pct = item["pct"]
    prev_close = item.get("prev_close")
    cls = _color_class(pct)

    # Direction arrow and top border color
    if pct > 0:
        arrow = "▲"
        border_color = COLOR_UP
    elif pct < 0:
        arrow = "▼"
        border_color = COLOR_DOWN
    else:
        arrow = ""
        border_color = "#D1D5DB"

    prev_html = ""
    if prev_close is not None:
        prev_html = (
            f'<div class="indicator-card__prev">'
            f'전일 {_fmt_value(name, prev_close)}'
            f'</div>'
        )

    return f"""
    <div class="indicator-card">
        <div class="indicator-card__border" style="background:{border_color};"></div>
        <div class="indicator-card__body">
            <div class="indicator-card__name">{name}</div>
            <div class="indicator-card__value">{_fmt_value(name, val)}</div>
            <div class="indicator-card__change {cls}">
                {arrow} {_fmt_change(name, change, pct)}
            </div>
            {prev_html}
        </div>
    </div>
    """


# ── Indices Section ──
section_header("주요 지수")

if indices:
    domestic = [i for i in indices if i["name"] in ("KOSPI", "KOSDAQ")]
    international = [i for i in indices if i["name"] not in ("KOSPI", "KOSDAQ")]

    if domestic:
        subsection_label("국내")
        dom_cols = st.columns(len(domestic))
        for col, item in zip(dom_cols, domestic):
            with col:
                st.markdown(_render_indicator_card(item), unsafe_allow_html=True)

    if international:
        subsection_label("해외")
        intl_cols = st.columns(len(international))
        for col, item in zip(intl_cols, international):
            with col:
                st.markdown(_render_indicator_card(item), unsafe_allow_html=True)
else:
    st.info("지수 데이터를 불러올 수 없습니다. (Yahoo Finance 연결 확인)")

# ── Macro Section ──
section_header("매크로 지표")

if macro:
    MACRO_GROUPS = {
        "환율": ["USD/KRW", "USD/JPY"],
        "원자재": ["금(oz)", "WTI유"],
        "시장 심리": ["비트코인", "VIX", "미국10Y"],
    }

    for group_label, group_names in MACRO_GROUPS.items():
        group_items = [m for m in macro if m["name"] in group_names]
        if not group_items:
            continue
        subsection_label(group_label)
        cols = st.columns(len(group_items))
        for col, item in zip(cols, group_items):
            with col:
                st.markdown(_render_indicator_card(item), unsafe_allow_html=True)
else:
    st.info("매크로 데이터를 불러올 수 없습니다.")

st.markdown(
    '<p class="source-text">Yahoo Finance 기준 (전일 대비)</p>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Investor Trading Trends (투자자별 매매동향)
# ---------------------------------------------------------------------------

date_str = st.session_state.get("date_str", selected_date.strftime("%Y%m%d"))
trends = load_investor_trends(date_str)
trend_overview = trends.get("overview", {})
trend_detail = trends.get("detail", {})


def _trend_bar_chart(
    names: list[str], values: list[int], title: str, height: int = 260,
) -> go.Figure:
    """Horizontal bar chart for investor net buy amounts (억원)."""
    values_eok = [int(round(v / 1_0000_0000)) for v in values]
    colors = [COLOR_UP if v > 0 else COLOR_DOWN if v < 0 else COLOR_FLAT for v in values_eok]

    # Dynamic text positioning
    max_abs = max(abs(v) for v in values_eok) if values_eok else 1
    text_pos = [
        "inside" if abs(v) > max_abs * 0.3 else "outside"
        for v in values_eok
    ]
    text_colors = [
        "#FFFFFF" if pos == "inside"
        else (COLOR_UP if v > 0 else COLOR_DOWN if v < 0 else COLOR_FLAT)
        for v, pos in zip(values_eok, text_pos)
    ]

    fig = go.Figure(go.Bar(
        y=names[::-1],
        x=values_eok[::-1],
        orientation="h",
        marker=dict(
            color=colors[::-1],
            line=dict(width=0),
            cornerradius=4,
        ),
        text=[f"{v:+,}억" for v in values_eok[::-1]],
        textposition=text_pos[::-1],
        cliponaxis=False,
        textfont=dict(size=12, color=text_colors[::-1]),
    ))

    fig.update_layout(
        height=height,
        margin=dict(l=10, r=80, t=44, b=10),
        template=CHART_TEMPLATE,
        title=dict(
            text=f"<b>{title}</b>",
            font=dict(size=14, color="#1F3864"),
            x=0.01,
            y=0.95,
        ),
        xaxis=dict(
            showgrid=True,
            gridcolor="rgba(0,0,0,0.04)",
            gridwidth=1,
            zeroline=True,
            zerolinecolor="#9CA3AF",
            zerolinewidth=1.5,
            showticklabels=False,
        ),
        yaxis=dict(
            showgrid=False,
            automargin=True,
            tickfont=dict(size=13, color="#374151"),
        ),
        bargap=0.30,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def _trend_summary_table(overview: dict[str, dict[str, int]]) -> str:
    """HTML summary table for KOSPI/KOSDAQ investor trends."""
    inv_order = ["외국인", "기관", "개인"]
    rows_html = ""
    for inv in inv_order:
        kospi_val = overview.get("KOSPI", {}).get(inv, 0)
        kosdaq_val = overview.get("KOSDAQ", {}).get(inv, 0)
        total_val = kospi_val + kosdaq_val

        def _cell(v: int) -> str:
            eok = int(round(v / 1_0000_0000))
            cls = "up" if eok > 0 else "down" if eok < 0 else "flat"
            arrow = "▲" if eok > 0 else "▼" if eok < 0 else ""
            return (
                f'<td class="{cls}" style="text-align:right;padding:6px 10px;">'
                f'{arrow} {eok:+,}억</td>'
            )

        rows_html += (
            f'<tr>'
            f'<td style="padding:6px 10px;font-weight:600;color:#374151;">{inv}</td>'
            f'{_cell(kospi_val)}{_cell(kosdaq_val)}{_cell(total_val)}'
            f'</tr>'
        )

    return (
        '<table class="summary-table">'
        '<thead><tr>'
        '<th style="text-align:left;">투자자</th>'
        '<th style="text-align:right;">KOSPI</th>'
        '<th style="text-align:right;">KOSDAQ</th>'
        '<th style="text-align:right;">합계</th>'
        '</tr></thead>'
        f'<tbody>{rows_html}</tbody>'
        '</table>'
    )


# ── 투자자별 매매동향 ──
if trend_overview:
    section_header("투자자별 매매동향")

    inv_order = ["외국인", "기관", "개인"]
    col_kospi, col_kosdaq = st.columns(2)

    with col_kospi:
        kospi_data = trend_overview.get("KOSPI", {})
        names = [n for n in inv_order if n in kospi_data]
        vals = [kospi_data[n] for n in names]
        if names:
            fig = _trend_bar_chart(names, vals, "KOSPI", height=230)
            st.plotly_chart(fig, width="stretch")

    with col_kosdaq:
        kosdaq_data = trend_overview.get("KOSDAQ", {})
        names = [n for n in inv_order if n in kosdaq_data]
        vals = [kosdaq_data[n] for n in names]
        if names:
            fig = _trend_bar_chart(names, vals, "KOSDAQ", height=230)
            st.plotly_chart(fig, width="stretch")

    st.markdown(_trend_summary_table(trend_overview), unsafe_allow_html=True)

# ── 기관 세부 순매수 ──
if trend_detail:
    section_header("기관 세부 순매수")

    detail_order = ["연기금", "투자신탁", "사모펀드"]
    col_kospi2, col_kosdaq2 = st.columns(2)

    with col_kospi2:
        kospi_detail = trend_detail.get("KOSPI", {})
        names = [n for n in detail_order if n in kospi_detail]
        vals = [kospi_detail[n] for n in names]
        if names:
            fig = _trend_bar_chart(names, vals, "KOSPI 기관 세부", height=230)
            st.plotly_chart(fig, width="stretch")

    with col_kosdaq2:
        kosdaq_detail = trend_detail.get("KOSDAQ", {})
        names = [n for n in detail_order if n in kosdaq_detail]
        vals = [kosdaq_detail[n] for n in names]
        if names:
            fig = _trend_bar_chart(names, vals, "KOSDAQ 기관 세부", height=230)
            st.plotly_chart(fig, width="stretch")

    st.markdown(
        '<p class="source-text">KRX 기준 (당일 순매수 합계)</p>',
        unsafe_allow_html=True,
    )
