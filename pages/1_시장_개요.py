"""Page 1: Market Overview — prices, rankings, and new highs."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# Ensure project root importable
_PROJECT_DIR = Path(__file__).parent.parent.resolve()
if str(_PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(_PROJECT_DIR))

from dashboard_data import load_price_data, load_high_data
from dashboard_style import (
    inject_css, section_header, title_date_text, fmt_eok, fmt_pct, page_header,
    COLOR_UP, COLOR_DOWN, COLOR_FLAT, COLOR_PRIMARY, COLOR_PRIMARY_LIGHT,
    CHART_TEMPLATE,
)

st.set_page_config(page_title="시장 개요", page_icon="◻", layout="wide")
inject_css()


def _style_pct_col(val):
    """Apply red/blue color to 등락률 cells."""
    if not val or val == "-":
        return "color: #9CA3AF"
    try:
        num = float(val.replace("%", "").replace("+", "").replace(",", ""))
        if "+" in str(val) or num > 0:
            return f"color: {COLOR_UP}; font-weight: 600"
        elif num < 0:
            return f"color: {COLOR_DOWN}; font-weight: 600"
    except (ValueError, AttributeError):
        pass
    return "color: #9CA3AF"

# ---------------------------------------------------------------------------
# Sidebar date (shared state)
# ---------------------------------------------------------------------------
selected_date = st.session_state.get("selected_date")
date_str = st.session_state.get("date_str")

if not selected_date or not date_str:
    st.warning("메인 페이지에서 날짜를 먼저 선택해주세요.")
    st.stop()

page_header("시장 개요", title_date_text(selected_date))

# ---------------------------------------------------------------------------
# Load Data
# ---------------------------------------------------------------------------
prices = load_price_data(date_str)

if not prices:
    st.warning("시세 데이터를 불러올 수 없습니다. 해당 일자의 데이터가 없거나 KRX 접속이 불가합니다.")
    st.info("주가 차트는 '관심종목' 페이지에서 Yahoo Finance 데이터로 확인할 수 있습니다.")
    st.stop()

df = pd.DataFrame(prices)

# ---------------------------------------------------------------------------
# KPI Cards
# ---------------------------------------------------------------------------
limit_up = len(df[df["pct"] >= 29.5]) if "pct" in df.columns else 0
limit_down = len(df[df["pct"] <= -29.5]) if "pct" in df.columns else 0
total_trade_value = df["trade_value"].sum() if "trade_value" in df.columns else 0
total_stocks = len(df)

c1, c2, c3, c4 = st.columns(4)
c1.metric("상한가", f"{limit_up}종목")
c2.metric("하한가", f"{limit_down}종목")
c3.metric("총 거래대금", f"{total_trade_value / 1_0000_0000_0000:.1f}조")
c4.metric("전체 종목수", f"{total_stocks:,}")

st.markdown("---")

# ---------------------------------------------------------------------------
# Rankings
# ---------------------------------------------------------------------------
section_header("종목 랭킹")

tab_pct, tab_trade, tab_ratio = st.tabs(["등락률 상위", "거래대금 상위", "시총대비 거래대금"])


def _bar_chart(names: list, values: list, title: str, fmt: str = ",.0f", color_by_sign: bool = True) -> go.Figure:
    if color_by_sign:
        colors = [COLOR_UP if v > 0 else COLOR_DOWN if v < 0 else COLOR_FLAT for v in values]
    else:
        colors = [COLOR_PRIMARY] * len(values)

    fig = go.Figure(go.Bar(
        y=names[::-1],
        x=values[::-1],
        orientation="h",
        marker_color=colors[::-1],
        text=[f"{v:{fmt}}" for v in values[::-1]],
        textposition="outside",
        cliponaxis=False,
        textfont=dict(size=11, color="#374151"),
    ))
    fig.update_layout(
        height=max(450, len(names) * 36),
        margin=dict(l=10, r=70, t=40, b=5),
        template=CHART_TEMPLATE,
        title=dict(text=title, font=dict(size=13, color="#1F3864", weight=700)),
        xaxis=dict(showgrid=True, gridcolor="#F0F0F0", side="top"),
        yaxis=dict(showgrid=False, automargin=True, tickfont=dict(size=11, color="#374151")),
        bargap=0.28,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig


with tab_pct:
    top_up = df.nlargest(20, "pct")
    col_chart, col_table = st.columns([3, 4])

    with col_chart:
        fig = _bar_chart(
            top_up["name"].tolist(),
            top_up["pct"].tolist(),
            "등락률 상위 20",
            fmt="+.2f",
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_table:
        display_df = top_up[["name", "code", "market", "close", "change", "pct", "market_cap"]].copy()
        display_df.columns = ["종목명", "종목코드", "시장", "종가", "등락(원)", "등락률(%)", "시총(억)"]
        display_df["시총(억)"] = display_df["시총(억)"].apply(lambda x: f"{int(round(x / 1_0000_0000)):,}" if x else "-")
        display_df["종가"] = display_df["종가"].apply(lambda x: f"{x:,}")
        display_df["등락(원)"] = display_df["등락(원)"].apply(lambda x: f"{x:+,}")
        display_df["등락률(%)"] = display_df["등락률(%)"].apply(lambda x: f"{x:+.2f}%")
        display_df = display_df.reset_index(drop=True)
        display_df.index = display_df.index + 1
        styled_df = display_df.style.map(_style_pct_col, subset=["등락률(%)"])
        st.dataframe(styled_df, use_container_width=True, height=600, hide_index=False)


with tab_trade:
    top_trade = df.nlargest(20, "trade_value")
    col_chart, col_table = st.columns([3, 4])

    with col_chart:
        trade_eok = (top_trade["trade_value"] / 1_0000_0000).astype(int)
        fig = _bar_chart(
            top_trade["name"].tolist(),
            trade_eok.tolist(),
            "거래대금 상위 20 (억원)",
            fmt=",.0f",
            color_by_sign=False,
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_table:
        display_df = top_trade[["name", "code", "market", "close", "pct", "trade_value", "market_cap"]].copy()
        display_df.columns = ["종목명", "종목코드", "시장", "종가", "등락률(%)", "거래대금(억)", "시총(억)"]
        display_df["시총(억)"] = display_df["시총(억)"].apply(lambda x: f"{int(round(x / 1_0000_0000)):,}" if x else "-")
        display_df["거래대금(억)"] = display_df["거래대금(억)"].apply(lambda x: f"{int(round(x / 1_0000_0000)):,}" if x else "-")
        display_df["종가"] = display_df["종가"].apply(lambda x: f"{x:,}")
        display_df["등락률(%)"] = display_df["등락률(%)"].apply(lambda x: f"{x:+.2f}%")
        display_df = display_df.reset_index(drop=True)
        display_df.index = display_df.index + 1
        styled_df = display_df.style.map(_style_pct_col, subset=["등락률(%)"])
        st.dataframe(styled_df, use_container_width=True, height=600, hide_index=False)


with tab_ratio:
    # Trade value / market cap ratio
    df_ratio = df[df["market_cap"] > 0].copy()
    df_ratio["ratio"] = df_ratio["trade_value"] / df_ratio["market_cap"] * 100
    top_ratio = df_ratio.nlargest(20, "ratio")

    col_chart, col_table = st.columns([3, 4])

    with col_chart:
        fig = _bar_chart(
            top_ratio["name"].tolist(),
            top_ratio["ratio"].round(2).tolist(),
            "시총대비 거래대금 상위 20 (%)",
            fmt=".2f",
            color_by_sign=False,
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_table:
        display_df = top_ratio[["name", "code", "market", "close", "pct", "trade_value", "market_cap", "ratio"]].copy()
        display_df.columns = ["종목명", "종목코드", "시장", "종가", "등락률(%)", "거래대금(억)", "시총(억)", "비율(%)"]
        display_df["시총(억)"] = display_df["시총(억)"].apply(lambda x: f"{int(round(x / 1_0000_0000)):,}" if x else "-")
        display_df["거래대금(억)"] = display_df["거래대금(억)"].apply(lambda x: f"{int(round(x / 1_0000_0000)):,}" if x else "-")
        display_df["종가"] = display_df["종가"].apply(lambda x: f"{x:,}")
        display_df["등락률(%)"] = display_df["등락률(%)"].apply(lambda x: f"{x:+.2f}%")
        display_df["비율(%)"] = display_df["비율(%)"].apply(lambda x: f"{x:.2f}%")
        display_df = display_df.reset_index(drop=True)
        display_df.index = display_df.index + 1
        styled_df = display_df.style.map(_style_pct_col, subset=["등락률(%)"])
        st.dataframe(styled_df, use_container_width=True, height=600, hide_index=False)


# ---------------------------------------------------------------------------
# New Highs
# ---------------------------------------------------------------------------
st.markdown("---")
section_header("신고가")

highs = load_high_data(date_str)

if not highs:
    st.info("신고가 데이터를 불러올 수 없습니다.")
else:
    high_tabs = st.tabs(list(highs.keys()))
    for tab, (sheet_name, rows) in zip(high_tabs, highs.items()):
        with tab:
            if not rows:
                st.info(f"{sheet_name} 데이터가 없습니다.")
                continue
            hdf = pd.DataFrame(rows)
            display_df = hdf[["code", "name", "market", "market_cap_eok", "pct", "high_price"]].copy()
            display_df.columns = ["종목코드", "종목명", "시장", "시총(억)", "등락률(%)", "고가"]
            display_df["시총(억)"] = display_df["시총(억)"].apply(lambda x: f"{x:,.0f}" if x else "-")
            display_df["등락률(%)"] = display_df["등락률(%)"].apply(lambda x: f"{x:+.2f}%" if x else "-")
            display_df["고가"] = display_df["고가"].apply(lambda x: f"{x:,}" if x else "-")
            display_df = display_df.reset_index(drop=True)
            display_df.index = display_df.index + 1
            styled_df = display_df.style.map(_style_pct_col, subset=["등락률(%)"])
            st.dataframe(styled_df, use_container_width=True, height=min(len(rows) * 35 + 40, 600), hide_index=False)
            st.caption(f"총 {len(rows)}종목")
