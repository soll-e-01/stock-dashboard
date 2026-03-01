"""Main entry point for the stock market dashboard.

Run with: streamlit run dashboard.py
"""
from __future__ import annotations

from datetime import date, timedelta

import streamlit as st

from dashboard_data import load_market_overview, load_index_detail
from dashboard_style import (
    inject_css, title_date_text, section_header, page_header,
    COLOR_UP, COLOR_DOWN, COLOR_FLAT,
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

    if st.button("캐시 초기화", use_container_width=True):
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


def _svg_sparkline(sparkline_data: list[dict], color: str, width: int = 300, height: int = 65) -> str:
    """Generate an SVG sparkline area chart for embedding in HTML."""
    if not sparkline_data or len(sparkline_data) < 2:
        return ""

    closes = [d["close"] for d in sparkline_data]
    min_v = min(closes)
    max_v = max(closes)
    spread = max_v - min_v or 1

    points = []
    for i, v in enumerate(closes):
        x = i / (len(closes) - 1) * width
        y = height - (v - min_v) / spread * (height - 4) - 2  # 2px padding
        points.append(f"{x:.1f},{y:.1f}")

    polyline_str = " ".join(points)
    area_str = f"0,{height} " + polyline_str + f" {width},{height}"

    r = int(color[1:3], 16)
    g = int(color[3:5], 16)
    b = int(color[5:7], 16)

    return (
        f'<svg width="100%" viewBox="0 0 {width} {height}" preserveAspectRatio="none" '
        f'style="display:block;padding:0 16px;">'
        f'<polygon points="{area_str}" fill="rgba({r},{g},{b},0.08)" />'
        f'<polyline points="{polyline_str}" fill="none" stroke="{color}" '
        f'stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round" />'
        f'</svg>'
    )


def _time_labels_html() -> str:
    """KRX 장중 스파크라인 아래 시간 레이블 (10:00~15:00)."""
    hours = [10, 11, 12, 13, 14, 15]
    labels = []
    for h in hours:
        pct = (h * 60 - 540) / 390 * 100  # 540=09:00, 390min=09:00~15:30
        labels.append(
            f'<span style="position:absolute;left:{pct:.1f}%;'
            f'transform:translateX(-50%);'
            f'font-size:0.5rem;color:#C8C8C8;font-weight:400;">'
            f'{h}:00</span>'
        )
    return (
        '<div style="position:relative;height:12px;margin:0;padding:0 16px;">'
        + "".join(labels)
        + '</div>'
    )


def _render_pro_card(item: dict, detail: dict | None) -> str:
    """Render a complete professional index card as a single HTML block."""
    name = item["name"]
    val = item["value"]
    change = item["change"]
    pct = item["pct"]
    prev_close = item.get("prev_close")
    cls = _color_class(pct)

    if pct > 0:
        arrow = "&#x25B2;"
        border_color = COLOR_UP
    elif pct < 0:
        arrow = "&#x25BC;"
        border_color = COLOR_DOWN
    else:
        arrow = ""
        border_color = "#D1D5DB"

    spark_color = COLOR_UP if pct >= 0 else COLOR_DOWN

    # Header section
    header_html = (
        f'<div class="pro-index__header">'
        f'<div class="pro-index__info">'
        f'<div class="pro-index__name">{name}</div>'
        f'<div class="pro-index__value">{_fmt_value(name, val)}</div>'
        f'<div class="pro-index__change {cls}">'
        f'{arrow} {_fmt_change(name, change, pct)}'
        f'</div>'
        f'</div>'
        f'</div>'
    )

    # SVG sparkline (prefer intraday, fallback to monthly)
    sparkline_html = ""
    intraday = detail.get("sparkline_intraday", []) if detail else []
    if intraday and len(intraday) >= 2:
        sparkline_html = _svg_sparkline(intraday, spark_color) + _time_labels_html()
    elif detail and detail.get("sparkline"):
        sparkline_html = _svg_sparkline(detail["sparkline"], spark_color)

    # Stats grid
    def _stat(label: str, value: str) -> str:
        return (
            f'<div class="pro-index__stat">'
            f'<div class="pro-index__stat-label">{label}</div>'
            f'<div class="pro-index__stat-value">{value}</div>'
            f'</div>'
        )

    if detail:
        open_v = f"{detail['open']:,.2f}"
        high_v = f"{detail['high']:,.2f}"
        low_v = f"{detail['low']:,.2f}"
        vol = detail["volume"]
        if vol >= 1_0000_0000:
            vol_str = f"{vol / 1_0000_0000:.1f}억"
        elif vol >= 1_0000:
            vol_str = f"{vol / 1_0000:.0f}만"
        else:
            vol_str = f"{vol:,}"
    else:
        open_v = high_v = low_v = vol_str = "-"

    prev_str = f"{prev_close:,.2f}" if prev_close is not None else "-"

    stats_html = (
        f'<div class="pro-index__stats">'
        + _stat("시가", open_v)
        + _stat("고가", high_v)
        + _stat("저가", low_v)
        + _stat("거래량", vol_str)
        + _stat("전일종가", prev_str)
        + '</div>'
    )

    # Daily range bar (당일 저가~고가)
    range_html = ""
    if detail:
        day_high = detail["high"]
        day_low = detail["low"]
        spread = day_high - day_low

        if spread <= 0:
            pct_pos = 50.0
        else:
            pct_pos = max(0, min(100, (val - day_low) / spread * 100))

        fill_color = COLOR_UP if pct >= 0 else COLOR_DOWN

        range_html = (
            f'<div class="pro-index__range">'
            f'<div class="pro-index__range-labels">'
            f'<span>저가 {day_low:,.2f}</span>'
            f'<span>고가 {day_high:,.2f}</span>'
            f'</div>'
            f'<div class="pro-index__range-track">'
            f'<div class="pro-index__range-fill" style="width:{pct_pos:.1f}%;background:{fill_color};"></div>'
            f'<div class="pro-index__range-dot" style="left:{pct_pos:.1f}%;background:{fill_color};"></div>'
            f'</div>'
            f'</div>'
        )

    return (
        f'<div class="pro-index" style="border-top-color:{border_color};">'
        + header_html
        + sparkline_html
        + stats_html
        + range_html
        + '</div>'
    )


# ── Domestic Indices (Professional Cards) ──
if indices:
    domestic = [i for i in indices if i["name"] in ("KOSPI", "KOSDAQ")]
    international = [i for i in indices if i["name"] not in ("KOSPI", "KOSDAQ")]

    if domestic:
        index_detail = load_index_detail()
        hero_cols = st.columns(len(domestic))
        for col, item in zip(hero_cols, domestic):
            with col:
                detail = index_detail.get(item["name"])
                card_html = _render_pro_card(item, detail)
                st.markdown(card_html, unsafe_allow_html=True)
        st.markdown("")

    # ── International Indices ──
    section_header("해외 지수")
    if international:
        intl_cols = st.columns(len(international))
        for col, item in zip(intl_cols, international):
            with col:
                st.markdown(_render_indicator_card(item), unsafe_allow_html=True)
else:
    st.info("지수 데이터를 불러올 수 없습니다. (Yahoo Finance 연결 확인)")

# ── Macro Section ──
section_header("매크로 지표")

if macro:
    macro_order = ["USD/KRW", "USD/JPY", "금(oz)", "WTI유", "비트코인", "VIX", "미국10Y"]
    macro_by_name = {m["name"]: m for m in macro}
    ordered_macro = [macro_by_name[n] for n in macro_order if n in macro_by_name]

    # Row 1: 4 items, Row 2: 3 items
    row1 = ordered_macro[:4]
    row2 = ordered_macro[4:]

    if row1:
        cols1 = st.columns(len(row1))
        for col, item in zip(cols1, row1):
            with col:
                st.markdown(_render_indicator_card(item), unsafe_allow_html=True)
    if row2:
        cols2 = st.columns(len(row2))
        for col, item in zip(cols2, row2):
            with col:
                st.markdown(_render_indicator_card(item), unsafe_allow_html=True)
else:
    st.info("매크로 데이터를 불러올 수 없습니다.")

st.markdown(
    '<p class="source-text">Yahoo Finance 기준 (전일 대비)</p>',
    unsafe_allow_html=True,
)
