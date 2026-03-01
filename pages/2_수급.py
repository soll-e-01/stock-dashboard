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

from dashboard_data import load_supply_data
from dashboard_style import (
    inject_css, section_header, title_date_text, fmt_eok, page_header,
    COLOR_UP, COLOR_DOWN, COLOR_FLAT, COLOR_PRIMARY,
    CHART_TEMPLATE,
)

st.set_page_config(page_title="수급", page_icon="◻", layout="wide")
inject_css()

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

# ---------------------------------------------------------------------------
# KPI Cards — total net buy per investor type
# ---------------------------------------------------------------------------
investor_names = list(supply.keys())
cols = st.columns(len(investor_names))

for col, inv_name in zip(cols, investor_names):
    rows = supply[inv_name]
    total = sum(r.get("net_buy", 0) for r in rows)
    total_eok = int(round(total / 1_0000_0000))
    delta_color = "normal"
    col.metric(
        inv_name,
        f"{total_eok:,}억",
        delta=f"{'순매수' if total >= 0 else '순매도'}",
        delta_color="normal" if total >= 0 else "inverse",
    )

st.markdown("---")

# ---------------------------------------------------------------------------
# Tabs per investor type
# ---------------------------------------------------------------------------

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
        textfont=dict(size=11),
    ))
    fig.update_layout(
        height=max(400, len(names) * 42),
        margin=dict(l=10, r=100, t=40, b=10),
        template=CHART_TEMPLATE,
        title=dict(text=title, font=dict(size=13, color="#1F3864"), x=0.01),
        xaxis=dict(showgrid=True, gridcolor="#F0F0F0", side="top", tickformat=","),
        yaxis=dict(showgrid=False, automargin=True, tickfont=dict(size=12)),
        bargap=0.3,
        plot_bgcolor="#FAFBFC",
    )
    return fig


tabs = st.tabs(investor_names)

for tab, inv_name in zip(tabs, investor_names):
    with tab:
        rows = supply[inv_name]
        if not rows:
            st.info(f"{inv_name} 수급 데이터가 없습니다.")
            continue

        df = pd.DataFrame(rows)

        # Net buy in 억원
        df["net_buy_eok"] = (df["net_buy"] / 1_0000_0000).round(0).astype(int)

        # Top 10 by absolute net buy and by market cap ratio
        top_buy = df.nlargest(10, "net_buy")
        top_sell = df.nsmallest(10, "net_buy")

        col_buy, col_sell = st.columns(2)

        with col_buy:
            fig = _supply_bar_chart(
                top_buy["name"].tolist(),
                top_buy["net_buy_eok"].tolist(),
                f"{inv_name} 순매수 Top 10 (억원)",
            )
            st.plotly_chart(fig, use_container_width=True)

        with col_sell:
            fig = _supply_bar_chart(
                top_sell["name"].tolist(),
                top_sell["net_buy_eok"].tolist(),
                f"{inv_name} 순매도 Top 10 (억원)",
            )
            st.plotly_chart(fig, use_container_width=True)

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
                    textfont=dict(size=11),
                ))
                fig.update_layout(
                    height=max(400, len(ratio_names) * 42),
                    margin=dict(l=10, r=100, t=40, b=10),
                    template=CHART_TEMPLATE,
                    title=dict(
                        text="시총대비 순매수 비율 Top 10 (%)",
                        font=dict(size=13, color="#1F3864"),
                        x=0.01,
                    ),
                    xaxis=dict(showgrid=True, gridcolor="#F0F0F0", side="top"),
                    yaxis=dict(showgrid=False, automargin=True, tickfont=dict(size=12)),
                    bargap=0.3,
                    plot_bgcolor="#FAFBFC",
                )
                st.plotly_chart(fig, use_container_width=True)

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
            st.dataframe(display_df, use_container_width=True, height=500)
