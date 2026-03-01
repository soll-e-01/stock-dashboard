"""Page 2: Supply/Demand — net buying by investor type."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

_PROJECT_DIR = Path(__file__).parent.parent.resolve()
if str(_PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(_PROJECT_DIR))

from dashboard_data import load_supply_data, load_investor_trends
from dashboard_style import (
    inject_css, section_header, title_date_text, fmt_eok, page_header,
    COLOR_UP, COLOR_DOWN, COLOR_FLAT, COLOR_PRIMARY,
    CHART_TEMPLATE,
)

st.set_page_config(page_title="수급", page_icon="◻", layout="wide")
inject_css()


def _style_signed_col(val):
    """Apply red/blue color to signed numeric cells."""
    if not val or val == "-":
        return "color: #9CA3AF"
    try:
        num = float(val.replace("%", "").replace("+", "").replace(",", ""))
        if num > 0:
            return f"color: {COLOR_UP}; font-weight: 600"
        elif num < 0:
            return f"color: {COLOR_DOWN}; font-weight: 600"
    except (ValueError, AttributeError):
        pass
    return "color: #9CA3AF"

selected_date = st.session_state.get("selected_date")
date_str = st.session_state.get("date_str")

if not selected_date or not date_str:
    st.warning("메인 페이지에서 날짜를 먼저 선택해주세요.")
    st.stop()

page_header("투자자별 수급", title_date_text(selected_date))

# ---------------------------------------------------------------------------
# Load Data
# ---------------------------------------------------------------------------
supply = load_supply_data(date_str)

if not supply:
    st.warning("수급 데이터를 불러올 수 없습니다. 해당 일자의 데이터가 없거나 KRX 접속이 불가합니다.")
    st.stop()

investor_names = list(supply.keys())

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def _trend_bar_chart(
    names: list[str], values: list[int], title: str, height: int = 260,
) -> go.Figure:
    """Horizontal bar chart for investor net buy amounts (억원)."""
    values_eok = [int(round(v / 1_0000_0000)) for v in values]
    colors = [COLOR_UP if v > 0 else COLOR_DOWN if v < 0 else COLOR_FLAT for v in values_eok]

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
        textfont=dict(size=9, color=text_colors[::-1]),
    ))

    fig.update_layout(
        height=height,
        margin=dict(l=10, r=50, t=30, b=4),
        template=CHART_TEMPLATE,
        title=dict(
            text=f"<b>{title}</b>",
            font=dict(size=12, color="#1F3864"),
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
            tickfont=dict(size=10, color="#374151"),
        ),
        bargap=0.20,
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


def _supply_bar_chart(names: list, values: list, title: str) -> go.Figure:
    colors = [COLOR_UP if v > 0 else COLOR_DOWN if v < 0 else COLOR_FLAT for v in values]
    fig = go.Figure(go.Bar(
        y=names[::-1],
        x=values[::-1],
        orientation="h",
        marker_color=colors[::-1],
        text=[f"{v:+,.0f}" for v in values[::-1]],
        textposition="outside",
        cliponaxis=False,
        textfont=dict(size=9),
    ))
    fig.update_layout(
        height=max(360, len(names) * 36),
        margin=dict(l=10, r=70, t=28, b=4),
        template=CHART_TEMPLATE,
        title=dict(text=title, font=dict(size=11, color="#1F3864"), x=0.01),
        xaxis=dict(showgrid=True, gridcolor="#F0F0F0", side="top", tickformat=","),
        yaxis=dict(showgrid=False, automargin=True, tickfont=dict(size=9)),
        bargap=0.25,
        plot_bgcolor="#FAFBFC",
    )
    return fig


# ---------------------------------------------------------------------------
# Tabs: 수급 개요 + 투자자별 상세
# ---------------------------------------------------------------------------
trends = load_investor_trends(date_str)
trend_overview = trends.get("overview", {})
trend_detail = trends.get("detail", {})

# ---------------------------------------------------------------------------
# KPI Cards (탭 밖)
# ---------------------------------------------------------------------------
cols = st.columns(len(investor_names))
for col, inv_name in zip(cols, investor_names):
    rows = supply[inv_name]
    total = sum(r.get("net_buy", 0) for r in rows)
    total_eok = int(round(total / 1_0000_0000))
    col.metric(
        inv_name,
        f"{total_eok:,}억",
        delta=f"{'순매수' if total >= 0 else '순매도'}",
        delta_color="normal" if total >= 0 else "inverse",
    )

st.markdown("---")

tab_names = ["수급 개요"] + investor_names
all_tabs = st.tabs(tab_names)

# ── 수급 개요 탭 ──
with all_tabs[0]:
    # 투자자별 매매동향
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

    # 기관 세부 순매수
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

# ── 투자자 유형별 탭 (사모펀드, 투자신탁, 연기금, 외국인, 개인) ──
for tab, inv_name in zip(all_tabs[1:], investor_names):
    with tab:
        rows = supply[inv_name]
        if not rows:
            st.info(f"{inv_name} 수급 데이터가 없습니다.")
            continue

        df = pd.DataFrame(rows)
        df["net_buy_eok"] = (df["net_buy"] / 1_0000_0000).round(0).astype(int)

        top_buy = df.nlargest(10, "net_buy")
        top_sell = df.nsmallest(10, "net_buy")

        col_buy, col_sell = st.columns(2)

        with col_buy:
            fig = _supply_bar_chart(
                top_buy["name"].tolist(),
                top_buy["net_buy_eok"].tolist(),
                f"{inv_name} 순매수 Top 10 (억원)",
            )
            st.plotly_chart(fig, width="stretch")

        with col_sell:
            fig = _supply_bar_chart(
                top_sell["name"].tolist(),
                top_sell["net_buy_eok"].tolist(),
                f"{inv_name} 순매도 Top 10 (억원)",
            )
            st.plotly_chart(fig, width="stretch")

        # Market cap ratio top 10
        if "ratio" in df.columns:
            section_header(f"{inv_name} 시총대비 순매수 상위")
            df_ratio = df[df["ratio"].notna() & (df["ratio"] != 0)].copy()
            if not df_ratio.empty:
                top_ratio = df_ratio.nlargest(10, "ratio")
                ratio_vals = top_ratio["ratio"].round(4).tolist()
                ratio_names = top_ratio["name"].tolist()
                ratio_colors = [COLOR_UP if v > 0 else COLOR_DOWN if v < 0 else COLOR_FLAT for v in ratio_vals]
                fig = go.Figure(go.Bar(
                    y=ratio_names[::-1],
                    x=ratio_vals[::-1],
                    orientation="h",
                    marker_color=ratio_colors[::-1],
                    text=[f"{v:+.3f}%" for v in ratio_vals[::-1]],
                    textposition="outside",
                    cliponaxis=False,
                    textfont=dict(size=9),
                ))
                fig.update_layout(
                    height=max(360, len(ratio_names) * 36),
                    margin=dict(l=10, r=70, t=28, b=4),
                    template=CHART_TEMPLATE,
                    title=dict(
                        text="시총대비 순매수 비율 Top 10 (%)",
                        font=dict(size=11, color="#1F3864"),
                        x=0.01,
                    ),
                    xaxis=dict(showgrid=True, gridcolor="#F0F0F0", side="top"),
                    yaxis=dict(showgrid=False, automargin=True, tickfont=dict(size=9)),
                    bargap=0.25,
                    plot_bgcolor="#FAFBFC",
                )
                st.plotly_chart(fig, width="stretch")

        # Full data table
        st.markdown("")
        with st.expander(f"{inv_name} 전체 데이터 ({len(df)}종목)", expanded=False):
            display_df = df[["code", "name", "market", "net_buy_eok", "market_cap", "ratio"]].copy()
            display_df.columns = ["종목코드", "종목명", "시장", "순매수(억)", "시총(억)", "비율(%)"]
            display_df["시총(억)"] = display_df["시총(억)"].apply(
                lambda x: f"{int(round(x / 1_0000_0000)):,}" if x else "-"
            )
            display_df["순매수(억)"] = display_df["순매수(억)"].apply(lambda x: f"{x:+,}")
            display_df["비율(%)"] = display_df["비율(%)"].apply(
                lambda x: f"{x:.2f}%" if x is not None else "-"
            )
            display_df = display_df.reset_index(drop=True)
            display_df.index = display_df.index + 1
            styled_df = display_df.style.map(_style_signed_col, subset=["순매수(억)", "비율(%)"])
            st.dataframe(styled_df, width="stretch", height=500, hide_index=False)
