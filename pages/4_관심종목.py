"""Page 4: Watchlist — Equity Research Model dashboard."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

_PROJECT_DIR = Path(__file__).parent.parent.resolve()
if str(_PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(_PROJECT_DIR))

from dashboard_data import (
    get_watchlist,
    get_financial_years,
    load_financial_model_cached,
    dict_to_model,
    load_watchlist_disclosures,
    load_watchlist_krx,
    load_price_history,
    add_watchlist_stock,
    remove_watchlist_stock,
    lookup_corp_code,
    load_naver_valuations,
    load_segment_data,
)
from dashboard_style import (
    inject_css, section_header, title_date_text, fmt_eok, fmt_pct, fmt_ratio,
    color_pct_html, page_header,
    COLOR_UP, COLOR_DOWN, COLOR_FLAT, COLOR_PRIMARY,
    CHART_TEMPLATE,
)

# Re-import model helpers
from daily_watchlist_automation import (
    IS_ACCOUNT_DEFS, BS_ACCOUNT_DEFS, MARGIN_ROWS, INVESTOR_TYPES,
    get_is_value, get_bs_value, get_yoy_value,
    safe_pct, to_eok, calc_valuations, parse_amount,
)

DART_REPORT_URL = "https://dart.fss.or.kr/dsaf001/main.do?rcpNo="

st.set_page_config(page_title="관심종목", page_icon="◻", layout="wide")
inject_css()

selected_date = st.session_state.get("selected_date")
date_str = st.session_state.get("date_str")

if not selected_date or not date_str:
    st.warning("메인 페이지에서 날짜를 먼저 선택해주세요.")
    st.stop()

page_header("관심종목 분석", title_date_text(selected_date))

# ---------------------------------------------------------------------------
# Watchlist Management (Sidebar)
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("---")
    section_header("관심종목 관리")

    # Add stock
    with st.expander("종목 추가", expanded=False):
        add_code = st.text_input("종목코드 (6자리)", placeholder="005930", key="add_stock_code")
        if st.button("조회 및 추가", key="btn_add"):
            if add_code and len(add_code.strip()) == 6:
                with st.spinner("DART에서 종목 정보 조회 중..."):
                    result = lookup_corp_code(add_code.strip())
                if result:
                    success = add_watchlist_stock(
                        result["stock_code"], result["corp_code"], result["name"],
                    )
                    if success:
                        st.success(f"{result['name']} 추가 완료")
                        st.rerun()
                    else:
                        st.warning("이미 등록된 종목입니다.")
                else:
                    st.error("종목을 찾을 수 없습니다.")
            else:
                st.warning("6자리 종목코드를 입력해주세요.")

    # Remove stock
    watchlist = get_watchlist()
    if watchlist:
        with st.expander("종목 제거", expanded=False):
            remove_options = {f"{w['name']} ({w['stock_code']})": w['stock_code'] for w in watchlist}
            remove_choice = st.selectbox(
                "제거할 종목", list(remove_options.keys()), key="remove_stock",
            )
            if st.button("제거", key="btn_remove", type="primary"):
                code_to_remove = remove_options[remove_choice]
                success = remove_watchlist_stock(code_to_remove)
                if success:
                    st.success(f"제거 완료")
                    st.rerun()
                else:
                    st.error("제거 실패")

# ---------------------------------------------------------------------------
# Load watchlist config
# ---------------------------------------------------------------------------
watchlist = get_watchlist()

if not watchlist:
    st.info("관심종목이 비어있습니다. 왼쪽 사이드바에서 종목을 추가해주세요.")
    st.stop()

years = get_financial_years()
years_json = json.dumps(years)

# ---------------------------------------------------------------------------
# Load KRX data for watchlist
# ---------------------------------------------------------------------------
krx_data = load_watchlist_krx(date_str)
price_rows = krx_data.get("price_rows", [])
supply_by_type = krx_data.get("supply_by_type", {})

# Build lookup maps
price_by_code: dict[str, dict] = {}
for row in price_rows:
    code = str(row.get("code", "")).strip()
    price_by_code[code] = row

supply_by_code: dict[str, dict[str, int]] = {}
for inv_name, rows in supply_by_type.items():
    for row in rows:
        code = str(row.get("code", "")).strip()
        if code not in supply_by_code:
            supply_by_code[code] = {}
        supply_by_code[code][inv_name] = int(row.get("net_buy", 0))

# ---------------------------------------------------------------------------
# Load all financial models
# ---------------------------------------------------------------------------
models_dict: dict[str, dict] = {}
for entry in watchlist:
    try:
        md = load_financial_model_cached(
            entry["corp_code"], entry["name"], entry["stock_code"],
            years_json, "CFS",
        )
        models_dict[entry["stock_code"]] = md
    except Exception as e:
        st.warning(f"{entry['name']} 재무 모델 로드 실패: {e}")

# ---------------------------------------------------------------------------
# Summary Comparison Table
# ---------------------------------------------------------------------------
section_header("종목 비교 요약")

summary_data = []
for entry in watchlist:
    stock_code = entry["stock_code"]
    pr = price_by_code.get(stock_code)
    mcap_raw = int(pr.get("market_cap", 0)) if pr else 0
    mcap_eok = int(round(mcap_raw / 1_0000_0000)) if mcap_raw else 0
    close_val = int(pr.get("close", 0)) if pr else 0
    pct_val = float(pr.get("pct", 0)) if pr else 0.0

    summary_data.append({
        "종목명": entry["name"],
        "시총(억)": mcap_eok,
        "주가": close_val,
        "등락률(%)": pct_val,
        "_sort_cap": mcap_eok,
    })

summary_df = pd.DataFrame(summary_data)
summary_df = summary_df.sort_values("_sort_cap", ascending=False).drop(columns=["_sort_cap"]).reset_index(drop=True)
summary_df.index = summary_df.index + 1
summary_df["시총(억)"] = summary_df["시총(억)"].apply(lambda x: f"{x:,}" if x else "-")
summary_df["주가"] = summary_df["주가"].apply(lambda x: f"{x:,}" if x else "-")

def _color_pct(val):
    """등락률 셀에 빨강(상승)/파랑(하락) 색상 적용."""
    if isinstance(val, (int, float)):
        if val > 0:
            return f"color: {COLOR_UP}; font-weight: 600;"
        elif val < 0:
            return f"color: {COLOR_DOWN}; font-weight: 600;"
    return "color: #888;"

summary_styled = (
    summary_df.style
    .format({"등락률(%)": lambda x: f"{x:+.2f}%" if isinstance(x, (int, float)) else x})
    .map(_color_pct, subset=["등락률(%)"])
)
st.dataframe(summary_styled, use_container_width=True)

st.markdown("---")

# ---------------------------------------------------------------------------
# Individual Stock Detail
# ---------------------------------------------------------------------------
stock_names = [f"{e['name']} ({e['stock_code']})" for e in watchlist]
selected_idx = st.selectbox("종목 선택", range(len(stock_names)), format_func=lambda i: stock_names[i])

entry = watchlist[selected_idx]
stock_code = entry["stock_code"]
corp_code = entry["corp_code"]
corp_name = entry["name"]

md = models_dict.get(stock_code)
if not md:
    st.error(f"{corp_name} 재무 모델을 불러올 수 없습니다.")
    st.stop()

model = dict_to_model(md)
price_row = price_by_code.get(stock_code)
supply_data = supply_by_code.get(stock_code)

st.markdown("---")

# ── Price Chart (Candlestick) ──
section_header("주가 차트")

chart_period = st.radio(
    "기간", ["1개월", "3개월", "6개월", "1년"],
    index=3, horizontal=True, key="chart_period",
)
period_map = {"1개월": "1mo", "3개월": "3mo", "6개월": "6mo", "1년": "1y"}
yf_period = period_map[chart_period]

# Determine market suffix
market_suffix = "KS"  # default KOSPI
if price_row:
    mkt = str(price_row.get("market", ""))
    if "KOSDAQ" in mkt or "코스닥" in mkt:
        market_suffix = "KQ"

history = load_price_history(stock_code, market_suffix, yf_period)

if history:
    hdf = pd.DataFrame(history)
    fig = go.Figure()

    # Candlestick chart
    fig.add_trace(go.Candlestick(
        x=hdf["date"],
        open=hdf["open"],
        high=hdf["high"],
        low=hdf["low"],
        close=hdf["close"],
        increasing_line_color=COLOR_UP,
        decreasing_line_color=COLOR_DOWN,
        increasing_fillcolor=COLOR_UP,
        decreasing_fillcolor=COLOR_DOWN,
        increasing_line_width=0.8,
        decreasing_line_width=0.8,
        name="주가",
    ))

    # 20-day moving average
    if len(hdf) >= 20:
        hdf["ma20"] = hdf["close"].rolling(window=20).mean()
        fig.add_trace(go.Scatter(
            x=hdf["date"], y=hdf["ma20"],
            mode="lines",
            line=dict(color="#F59E0B", width=1.2, dash="dot"),
            name="MA20",
        ))
    # 60-day moving average
    if len(hdf) >= 60:
        hdf["ma60"] = hdf["close"].rolling(window=60).mean()
        fig.add_trace(go.Scatter(
            x=hdf["date"], y=hdf["ma60"],
            mode="lines",
            line=dict(color="#8B5CF6", width=1.2, dash="dot"),
            name="MA60",
        ))

    # Volume bar chart on secondary y-axis
    vol_colors = [
        "rgba(204,0,0,0.25)" if hdf.iloc[i]["close"] >= hdf.iloc[i]["open"]
        else "rgba(0,102,204,0.25)"
        for i in range(len(hdf))
    ]
    fig.add_trace(go.Bar(
        x=hdf["date"],
        y=hdf["volume"],
        marker_color=vol_colors,
        marker_line_width=0,
        name="거래량",
        yaxis="y2",
        showlegend=False,
    ))

    fig.update_layout(
        height=500,
        margin=dict(l=0, r=0, t=40, b=0),
        template=CHART_TEMPLATE,
        title=dict(
            text=f"{corp_name} ({stock_code})",
            font=dict(size=14, color="#1F3864", weight=700),
            x=0.01, xanchor="left",
        ),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1,
            font=dict(size=10, color="#6B7280"),
            bgcolor="rgba(0,0,0,0)",
        ),
        xaxis=dict(
            rangeslider=dict(visible=False),
            showgrid=False,
            showline=True, linewidth=1, linecolor="#E5E7EB",
            tickfont=dict(size=10, color="#9CA3AF"),
        ),
        yaxis=dict(
            title="", showgrid=True, gridcolor="#F3F4F6", gridwidth=0.5,
            side="right",
            tickfont=dict(size=10, color="#6B7280"),
            zeroline=False,
        ),
        yaxis2=dict(
            title="", overlaying="y", side="left",
            showgrid=False, range=[0, max(hdf["volume"]) * 5],
            showticklabels=False,
        ),
        xaxis_rangebreaks=[dict(bounds=["sat", "mon"])],
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True)

    # Quick stats
    if len(hdf) >= 2:
        latest = hdf.iloc[-1]
        first = hdf.iloc[0]
        period_change = latest["close"] - first["close"]
        period_pct = (period_change / first["close"]) * 100
        high_max = hdf["high"].max()
        low_min = hdf["low"].min()
        avg_vol = hdf["volume"].mean()

        sc1, sc2, sc3, sc4 = st.columns(4)
        sc1.metric("기간 수익률", f"{period_pct:+.1f}%", delta=f"{period_change:+,.0f}원", delta_color="normal")
        sc2.metric("최고가", f"{high_max:,.0f}원")
        sc3.metric("최저가", f"{low_min:,.0f}원")
        sc4.metric("평균 거래량", f"{avg_vol:,.0f}")
else:
    st.info("주가 차트 데이터를 불러올 수 없습니다. (Yahoo Finance 연결 확인)")

st.markdown("---")

# ── Price Overview + Valuation ──
col_price, col_val = st.columns(2)

with col_price:
    section_header("시세 개요")
    if price_row:
        close = int(price_row.get("close", 0))
        change = int(price_row.get("change", 0))
        pct_val = float(price_row.get("pct", 0))
        mktcap = int(round(int(price_row.get("market_cap", 0)) / 1_0000_0000))

        pc1, pc2 = st.columns(2)
        pc1.metric("종가", f"{close:,}원", delta=f"{change:+,}원 ({pct_val:+.2f}%)", delta_color="normal")
        pc2.metric("시가총액", f"{mktcap:,}억")
    else:
        st.info("KRX 시세 데이터 없음")

with col_val:
    section_header("밸류에이션")
    vals = calc_valuations(model, price_row)

    # TTM EPS 계산
    from daily_watchlist_automation import calc_ttm
    eps_ttm = calc_ttm(model, "eps")
    controlling_ttm = calc_ttm(model, "controlling")

    # 주가 기준 직접 PER 계산
    close_price = int(price_row.get("close", 0)) if price_row else 0
    per_from_price = None
    if close_price > 0 and eps_ttm and eps_ttm != 0:
        per_from_price = close_price / eps_ttm

    # 네이버 금융 컨센서스
    naver_vals = load_naver_valuations(stock_code)

    # ── Compact valuation HTML card ──
    def _val_item(label: str, value: str, sub: str = "", highlight: bool = False) -> str:
        bg = "background:#EAF0FB;" if highlight else ""
        sub_html = f'<div style="font-size:0.65rem;color:#9CA3AF;margin-top:1px;">{sub}</div>' if sub else ""
        return (
            f'<div style="flex:1;min-width:0;padding:6px 8px;{bg}'
            f'border-right:1px solid #F0F0F0;text-align:center;">'
            f'<div style="font-size:0.62rem;color:#9CA3AF;font-weight:500;'
            f'text-transform:uppercase;letter-spacing:0.03em;margin-bottom:2px;">{label}</div>'
            f'<div style="font-size:0.82rem;font-weight:700;color:#111827;">{value}</div>'
            f'{sub_html}'
            f'</div>'
        )

    per_str = f"{vals['PER']:.1f}" if vals["PER"] is not None else "-"
    pbr_str = f"{vals['PBR']:.2f}" if vals["PBR"] is not None else "-"
    psr_str = f"{vals['PSR']:.1f}" if vals["PSR"] is not None else "-"
    ev_str = f"{vals['EV/EBITDA']:.1f}" if vals["EV/EBITDA"] is not None else "-"
    eps_str = f"{eps_ttm:,}" if eps_ttm else "-"
    fwd_per_str = f"{naver_vals['forward_per']:.1f}" if naver_vals.get("forward_per") else "-"
    fwd_eps_str = f"{int(naver_vals['forward_eps']):,}" if naver_vals.get("forward_eps") else "-"
    ind_per_str = f"{naver_vals['industry_per']:.1f}" if naver_vals.get("industry_per") else "-"
    div_str = f"{naver_vals['dividend_yield']:.2f}%" if naver_vals.get("dividend_yield") else "-"

    val_html = (
        '<div style="background:#FFF;border:1px solid #E5E7EB;border-radius:8px;overflow:hidden;">'
        # Row 1: Trailing
        '<div style="padding:4px 10px 2px;background:#F8F9FA;">'
        '<span style="font-size:0.62rem;font-weight:600;color:#6B7280;letter-spacing:0.03em;">TRAILING (DART 실적 기준)</span>'
        '</div>'
        '<div style="display:flex;border-bottom:1px solid #F0F0F0;">'
        + _val_item("PER", per_str)
        + _val_item("PBR", pbr_str)
        + _val_item("PSR", psr_str)
        + _val_item("EV/EBITDA", ev_str)
        + _val_item("EPS(TTM)", f"{eps_str}원")
        + '</div>'
        # Row 2: Forward
        '<div style="padding:4px 10px 2px;background:#F0F5FF;">'
        '<span style="font-size:0.62rem;font-weight:600;color:#2F5496;letter-spacing:0.03em;">FORWARD (컨센서스)</span>'
        '</div>'
        '<div style="display:flex;">'
        + _val_item("추정 PER", fwd_per_str, "", True)
        + _val_item("추정 EPS", f"{fwd_eps_str}원", "", True)
        + _val_item("업종 PER", ind_per_str, "", True)
        + _val_item("배당수익률", div_str, "", True)
        + '</div>'
        '</div>'
    )
    st.markdown(val_html, unsafe_allow_html=True)
    st.caption(
        "Trailing: DART 공시 실적(최근4분기) | "
        "Forward: 네이버금융(에프앤가이드 컨센서스)"
    )

# ── Segment Revenue ──
section_header("사업부문별 매출")

seg_year = max(years) if years else 2025
segments = load_segment_data(corp_code, seg_year)
if not segments:
    # 직전 연도로 재시도
    segments = load_segment_data(corp_code, seg_year - 1)
    if segments:
        seg_year = seg_year - 1

if segments:
    col_seg_chart, col_seg_table = st.columns([1, 1])

    with col_seg_chart:
        seg_names = [s["name"] for s in segments]
        seg_pcts = [s["pct"] for s in segments]

        fig_seg = go.Figure(go.Pie(
            labels=seg_names,
            values=seg_pcts,
            hole=0.45,
            textinfo="label+percent",
            textfont=dict(size=11),
            marker=dict(
                colors=["#2F5496", "#E74C3C", "#27AE60", "#F39C12", "#8E44AD",
                         "#1ABC9C", "#E67E22", "#3498DB", "#9B59B6", "#2ECC71"],
            ),
            insidetextorientation="auto",
        ))
        fig_seg.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=30, b=10),
            template=CHART_TEMPLATE,
            title=dict(
                text=f"매출 비중 ({seg_year}년)",
                font=dict(size=13, color="#1F3864"),
                x=0.5, xanchor="center",
            ),
            legend=dict(
                font=dict(size=10, color="#374151"),
                orientation="h", yanchor="bottom", y=-0.15,
                xanchor="center", x=0.5,
            ),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_seg, use_container_width=True)

    with col_seg_table:
        from dashboard_style import fmt_jo
        seg_table = []
        for s in segments:
            rev_eok = int(round(s["revenue"] / 1_0000_0000))
            seg_table.append({
                "사업부문": s["name"],
                "매출액(억)": f"{rev_eok:,}",
                "비중(%)": f"{s['pct']:.1f}%",
            })
        seg_df = pd.DataFrame(seg_table)
        seg_df.index = seg_df.index + 1
        st.dataframe(seg_df, use_container_width=True, hide_index=True,
                      height=min(len(seg_table) * 35 + 40, 320))
        st.caption(f"출처: DART {seg_year}년 사업보고서")
else:
    st.info("사업부문별 매출 데이터가 없습니다. (DART 사업보고서에 세그먼트 정보가 없는 종목)")

# ── Segment Revenue History (매출비중 변화) ──
with st.expander("사업부문 매출 추이 (2023년~)", expanded=False):
    from datetime import date as _date
    _current_year = _date.today().year
    seg_history: dict[int, list] = {}
    for _yr in range(2023, _current_year + 1):
        _segs = load_segment_data(corp_code, _yr)
        if _segs:
            seg_history[_yr] = _segs

    if seg_history and len(seg_history) >= 1:
        # 모든 부문명 수집 (순서 보존)
        all_seg_names: list[str] = []
        seen_names: set[str] = set()
        for yr in sorted(seg_history.keys()):
            for s in seg_history[yr]:
                if s["name"] not in seen_names:
                    all_seg_names.append(s["name"])
                    seen_names.add(s["name"])

        chart_years = sorted(seg_history.keys())
        year_strs = [str(y) for y in chart_years]
        seg_colors = ["#2F5496", "#E74C3C", "#27AE60", "#F39C12", "#8E44AD",
                      "#1ABC9C", "#E67E22", "#3498DB", "#9B59B6", "#2ECC71"]

        # ── 1) 매출액 추이 ──
        st.markdown(
            '<p style="font-weight:700;color:#1F3864;font-size:0.95rem;margin:12px 0 4px;">'
            '매출액 추이</p>',
            unsafe_allow_html=True,
        )

        # 매출액 Grouped Bar Chart
        fig_rev = go.Figure()
        for si, seg_name in enumerate(all_seg_names):
            rev_vals = []
            for yr in chart_years:
                yr_segs = {s["name"]: s for s in seg_history[yr]}
                if seg_name in yr_segs:
                    rev_vals.append(int(round(yr_segs[seg_name]["revenue"] / 1_0000_0000)))
                else:
                    rev_vals.append(0)
            fig_rev.add_trace(go.Bar(
                x=year_strs,
                y=rev_vals,
                name=seg_name,
                marker_color=seg_colors[si % len(seg_colors)],
                text=[f"{v:,}" if v else "" for v in rev_vals],
                textposition="outside",
                textfont=dict(size=9, color="#6B7280"),
                cliponaxis=False,
            ))
        fig_rev.update_layout(
            barmode="group",
            height=340,
            margin=dict(l=0, r=0, t=10, b=0),
            template=CHART_TEMPLATE,
            yaxis=dict(
                title="억원", titlefont=dict(size=10, color="#9CA3AF"),
                showgrid=True, gridcolor="#F3F4F6", gridwidth=0.5,
                tickfont=dict(size=9, color="#9CA3AF"),
            ),
            xaxis=dict(showgrid=False, tickfont=dict(size=11, color="#374151")),
            legend=dict(
                orientation="h", yanchor="bottom", y=-0.22,
                xanchor="center", x=0.5, font=dict(size=10),
            ),
            bargap=0.25, bargroupgap=0.08,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_rev, use_container_width=True)

        # 매출액 테이블
        rev_rows = []
        for seg_name in all_seg_names:
            row: dict[str, Any] = {"사업부문": seg_name}
            prev_val = None
            for yr in chart_years:
                yr_segs = {s["name"]: s for s in seg_history[yr]}
                if seg_name in yr_segs:
                    rev_eok = int(round(yr_segs[seg_name]["revenue"] / 1_0000_0000))
                    row[f"{yr}"] = f"{rev_eok:,}" if rev_eok else "-"
                    if prev_val and rev_eok:
                        yoy = (rev_eok - prev_val) / abs(prev_val) * 100
                        row[f"{yr} YoY"] = f"{yoy:+.1f}%"
                    else:
                        row[f"{yr} YoY"] = ""
                    prev_val = rev_eok if rev_eok else None
                else:
                    row[f"{yr}"] = "-"
                    row[f"{yr} YoY"] = ""
                    prev_val = None
            rev_rows.append(row)
        rev_df = pd.DataFrame(rev_rows)
        # YoY 열 스타일링
        def _style_yoy(val):
            if isinstance(val, str) and val.startswith("+"):
                return f"color:{COLOR_UP};font-size:0.8em;font-weight:600;"
            elif isinstance(val, str) and val.startswith("-"):
                return f"color:{COLOR_DOWN};font-size:0.8em;font-weight:600;"
            return "color:#9CA3AF;font-size:0.8em;"
        yoy_cols = [c for c in rev_df.columns if "YoY" in c]
        styled_rev = rev_df.style.map(_style_yoy, subset=yoy_cols) if yoy_cols else rev_df.style
        st.dataframe(styled_rev, use_container_width=True, hide_index=True,
                     height=min(len(rev_rows) * 35 + 40, 320))

        st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

        # ── 2) 매출비중 추이 ──
        st.markdown(
            '<p style="font-weight:700;color:#1F3864;font-size:0.95rem;margin:12px 0 4px;">'
            '매출비중 추이</p>',
            unsafe_allow_html=True,
        )

        # 비중 Stacked Bar Chart
        if len(chart_years) >= 2:
            fig_pct = go.Figure()
            for si, seg_name in enumerate(all_seg_names):
                pcts = []
                for yr in chart_years:
                    yr_segs = {s["name"]: s for s in seg_history[yr]}
                    pcts.append(yr_segs[seg_name]["pct"] if seg_name in yr_segs else 0)
                fig_pct.add_trace(go.Bar(
                    x=year_strs,
                    y=pcts,
                    name=seg_name,
                    marker_color=seg_colors[si % len(seg_colors)],
                    text=[f"{p:.1f}%" if p else "" for p in pcts],
                    textposition="inside",
                    insidetextanchor="middle",
                    textfont=dict(size=10, color="white", weight="bold"),
                ))
            fig_pct.update_layout(
                barmode="stack",
                height=340,
                margin=dict(l=0, r=0, t=10, b=0),
                template=CHART_TEMPLATE,
                yaxis=dict(
                    title="%", titlefont=dict(size=10, color="#9CA3AF"),
                    showgrid=True, gridcolor="#F3F4F6", gridwidth=0.5,
                    tickfont=dict(size=9, color="#9CA3AF"),
                    range=[0, 105],
                ),
                xaxis=dict(showgrid=False, tickfont=dict(size=11, color="#374151")),
                legend=dict(
                    orientation="h", yanchor="bottom", y=-0.22,
                    xanchor="center", x=0.5, font=dict(size=10),
                ),
                bargap=0.35,
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_pct, use_container_width=True)

        # 비중 테이블
        pct_rows = []
        for seg_name in all_seg_names:
            row: dict[str, Any] = {"사업부문": seg_name}
            first_pct = None
            last_pct = None
            for yr in chart_years:
                yr_segs = {s["name"]: s for s in seg_history[yr]}
                if seg_name in yr_segs:
                    pct = yr_segs[seg_name]["pct"]
                    row[f"{yr}"] = f"{pct:.1f}%"
                    if first_pct is None:
                        first_pct = pct
                    last_pct = pct
                else:
                    row[f"{yr}"] = "-"
            # 전체 기간 변화 (pp)
            if first_pct is not None and last_pct is not None and len(chart_years) >= 2:
                diff = last_pct - first_pct
                row["변화(pp)"] = f"{diff:+.1f}"
            else:
                row["변화(pp)"] = ""
            pct_rows.append(row)
        pct_df = pd.DataFrame(pct_rows)
        def _style_pp(val):
            if isinstance(val, str) and val.startswith("+"):
                return f"color:{COLOR_UP};font-weight:600;"
            elif isinstance(val, str) and val.startswith("-"):
                return f"color:{COLOR_DOWN};font-weight:600;"
            return "color:#9CA3AF;"
        styled_pct = pct_df.style.map(_style_pp, subset=["변화(pp)"]) if "변화(pp)" in pct_df.columns else pct_df.style
        st.dataframe(styled_pct, use_container_width=True, hide_index=True,
                     height=min(len(pct_rows) * 35 + 40, 320))

        st.caption("출처: DART 사업보고서 (2023년~) | 단위: 억원")
    else:
        st.info("매출비중 변화 데이터가 충분하지 않습니다.")

# ── Supply ──
section_header("투자자별 수급 (순매수)")

if supply_data:
    inv_names = list(supply_data.keys())
    inv_values = [int(round(supply_data[k] / 1_000_000)) for k in inv_names]

    colors = [COLOR_UP if v > 0 else COLOR_DOWN if v < 0 else COLOR_FLAT for v in inv_values]
    fig = go.Figure(go.Bar(
        x=inv_names,
        y=inv_values,
        marker_color=colors,
        marker_line_width=0,
        text=[f"{v:+,}" for v in inv_values],
        textposition="outside",
        cliponaxis=False,
        textfont=dict(size=10, color="#374151", weight="bold"),
    ))
    fig.update_layout(
        height=280,
        margin=dict(l=0, r=0, t=10, b=0),
        template=CHART_TEMPLATE,
        yaxis=dict(
            title="", showgrid=True, gridcolor="#F3F4F6", gridwidth=0.5,
            tickfont=dict(size=9, color="#9CA3AF"),
            zeroline=True, zerolinecolor="#E5E7EB", zerolinewidth=1,
        ),
        xaxis=dict(
            showgrid=False,
            tickfont=dict(size=10, color="#374151"),
        ),
        bargap=0.4,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("수급 데이터 없음")

# ── Income Statement ──
section_header("손익계산서")

if model.periods:
    # Separate annual and quarterly periods (sorted chronologically)
    fy_periods = sorted([p for p in model.periods if p.quarter == "FY"], key=lambda p: p.year)
    q_periods = sorted(
        [p for p in model.periods if p.quarter != "FY"],
        key=lambda p: (p.year, p.quarter),
    )
    year_sets = sorted(set(p.year for p in q_periods))

    # Column name helpers
    def _fy_col(p) -> str:
        return str(p.year)

    def _q_col(p) -> str:
        return f"{str(p.year)[2:]}.{p.quarter}"

    # ── Shared builder functions ──
    def _build_is_rows(periods, col_fn):
        cols = [col_fn(p) for p in periods]
        rows = []
        for acct_def in IS_ACCOUNT_DEFS:
            key = acct_def["key"]
            is_eps = key == "eps"
            row = {"계정": acct_def["display"]}
            for p, cn in zip(periods, cols):
                val = get_is_value(model, key, p)
                if is_eps:
                    row[cn] = f"{val:,}" if val else ""
                else:
                    row[cn] = f"{to_eok(val):,}" if val else ""
            rows.append(row)
            for mr in MARGIN_ROWS:
                if mr["after"] == key:
                    mrow = {"계정": mr["label"]}
                    for p, cn in zip(periods, cols):
                        if mr["type"] == "margin":
                            rev = get_is_value(model, "revenue", p)
                            val = get_is_value(model, mr["of"], p)
                            pct = safe_pct(val, rev)
                            mrow[cn] = f"{pct:.1f}" if pct is not None else ""
                        elif mr["type"] == "yoy":
                            yoy = get_yoy_value(model, mr["of"], p)
                            mrow[cn] = f"{yoy:+.1f}" if yoy is not None else ""
                    rows.append(mrow)
        return pd.DataFrame(rows), cols

    def _build_bs_rows(periods, col_fn):
        cols = [col_fn(p) for p in periods]
        rows = []
        for acct_def in BS_ACCOUNT_DEFS:
            key = acct_def["key"]
            row = {"계정": acct_def["display"]}
            for p, cn in zip(periods, cols):
                val = get_bs_value(model, key, p)
                row[cn] = f"{to_eok(val):,}" if val else ""
            rows.append(row)
        # 유동비율
        ratio_row = {"계정": "  유동비율(%)"}
        for p, cn in zip(periods, cols):
            ca = get_bs_value(model, "current_assets", p)
            cl = get_bs_value(model, "current_liabilities", p)
            pct = safe_pct(ca, cl)
            ratio_row[cn] = f"{pct:.1f}" if pct is not None else ""
        rows.append(ratio_row)
        # 부채비율
        debt_row = {"계정": "  부채비율(%)"}
        for p, cn in zip(periods, cols):
            tl = get_bs_value(model, "total_liabilities", p)
            te = get_bs_value(model, "total_equity", p)
            pct = safe_pct(tl, te)
            debt_row[cn] = f"{pct:.1f}" if pct is not None else ""
        rows.append(debt_row)
        return pd.DataFrame(rows), cols

    def _build_prof_rows(periods, col_fn):
        cols = [col_fn(p) for p in periods]
        prof_defs = [
            ("ROE(%)", "net_income", "total_equity"),
            ("ROA(%)", "net_income", "total_assets"),
            ("OPM(%)", "operating", "revenue"),
            ("NPM(%)", "net_income", "revenue"),
            ("매출원가율(%)", "cogs", "revenue"),
        ]
        rows = []
        for label, num_key, den_key in prof_defs:
            row = {"지표": label}
            for p, cn in zip(periods, cols):
                numerator = get_is_value(model, num_key, p)
                if den_key in ("total_equity", "total_assets"):
                    denominator = get_bs_value(model, den_key, p)
                else:
                    denominator = get_is_value(model, den_key, p)
                pct = safe_pct(numerator, denominator)
                row[cn] = f"{pct:.1f}" if pct is not None else ""
            rows.append(row)
        return pd.DataFrame(rows)

    # ── Style helpers ──
    def _style_margin_only(row):
        """Style: margin/yoy rows get muted italic."""
        label_col = "계정" if "계정" in row.index else "지표"
        is_margin = str(row.get(label_col, "")).startswith("  ")
        return [
            "color:#888;font-style:italic;font-size:0.85em;" if is_margin else ""
            for _ in row.index
        ]

    def _style_q_yearly_stripe(row):
        """Quarterly tab: alternating year backgrounds + margin styling."""
        styles = []
        label_col = "계정" if "계정" in row.index else "지표"
        is_margin = str(row.get(label_col, "")).startswith("  ")
        for col_name in row.index:
            base = "color:#888;font-style:italic;font-size:0.85em;" if is_margin else ""
            if col_name not in (label_col,):
                for i, yr in enumerate(year_sets):
                    if col_name.startswith(str(yr)[2:] + "."):
                        if i % 2 == 1:
                            base += "background-color:#F5F5F5;"
                        break
            styles.append(base)
        return styles

    # ── IS Tabs ──
    is_tab_q, is_tab_a = st.tabs(["📅 분기별", "📊 연간"])

    with is_tab_q:
        if q_periods:
            is_df_q, _ = _build_is_rows(q_periods, _q_col)
            styled_q = is_df_q.style.apply(_style_q_yearly_stripe, axis=1)
            st.dataframe(styled_q, use_container_width=True, hide_index=True,
                         height=min(len(is_df_q) * 35 + 40, 700))
            st.caption("단위: 억원 (주당이익 제외) | 연도별 음영 구분")
        else:
            st.info("분기 데이터 없음")

    with is_tab_a:
        if fy_periods:
            is_df_a, _ = _build_is_rows(fy_periods, _fy_col)
            styled_a = is_df_a.style.apply(_style_margin_only, axis=1)
            st.dataframe(styled_a, use_container_width=True, hide_index=True,
                         height=min(len(is_df_a) * 35 + 40, 700))
            st.caption("단위: 억원 (주당이익 제외)")
        else:
            st.info("연간 데이터 없음")

    # ── Balance Sheet + Profitability ──
    section_header("재무상태표 / 수익성")
    bs_tab_q, bs_tab_a = st.tabs(["📅 분기별", "📊 연간"])

    with bs_tab_q:
        if q_periods:
            col_bs_q, col_prof_q = st.columns(2)
            with col_bs_q:
                bs_df_q, _ = _build_bs_rows(q_periods, _q_col)
                styled_bs_q = bs_df_q.style.apply(_style_q_yearly_stripe, axis=1)
                st.dataframe(styled_bs_q, use_container_width=True, hide_index=True,
                             height=min(len(bs_df_q) * 35 + 40, 500))
                st.caption("단위: 억원")
            with col_prof_q:
                prof_df_q = _build_prof_rows(q_periods, _q_col)
                styled_prof_q = prof_df_q.style.apply(_style_q_yearly_stripe, axis=1)
                st.dataframe(styled_prof_q, use_container_width=True, hide_index=True,
                             height=min(len(prof_df_q) * 35 + 40, 300))

    with bs_tab_a:
        if fy_periods:
            col_bs_a, col_prof_a = st.columns(2)
            with col_bs_a:
                bs_df_a, _ = _build_bs_rows(fy_periods, _fy_col)
                styled_bs_a = bs_df_a.style.apply(_style_margin_only, axis=1)
                st.dataframe(styled_bs_a, use_container_width=True, hide_index=True,
                             height=min(len(bs_df_a) * 35 + 40, 500))
                st.caption("단위: 억원")
            with col_prof_a:
                prof_df_a = _build_prof_rows(fy_periods, _fy_col)
                styled_prof_a = prof_df_a.style.apply(_style_margin_only, axis=1)
                st.dataframe(styled_prof_a, use_container_width=True, hide_index=True,
                             height=min(len(prof_df_a) * 35 + 40, 300))

# ── Recent Disclosures ──
section_header("최근 공시")

disc_list = load_watchlist_disclosures(corp_code, date_str)

if disc_list:
    disc_data = []
    for d in disc_list:
        rcept_no = d.get("rcept_no", "")
        disc_data.append({
            "일자": d.get("rcept_dt", ""),
            "공시명": d.get("report_nm", ""),
            "DART 링크": f"{DART_REPORT_URL}{rcept_no}" if rcept_no else "",
        })
    disc_df = pd.DataFrame(disc_data)
    disc_df.index = disc_df.index + 1
    st.dataframe(
        disc_df,
        use_container_width=True,
        column_config={
            "DART 링크": st.column_config.LinkColumn("DART 링크", display_text="보기"),
        },
        height=min(len(disc_list) * 35 + 40, 400),
    )
else:
    st.info("최근 공시 데이터가 없습니다.")
