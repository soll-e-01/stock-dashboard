"""Page 3: DART Disclosures — market-wide and watchlist."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import streamlit as st

_PROJECT_DIR = Path(__file__).parent.parent.resolve()
if str(_PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(_PROJECT_DIR))

from dashboard_data import load_disclosures, load_price_data
from dashboard_style import (
    inject_css, section_header, title_date_text, color_pct_html, page_header,
    COLOR_UP, COLOR_DOWN,
)

PBLNTF_TY_DISPLAY = {
    "A": "정기공시", "B": "주요사항", "C": "발행공시", "D": "지분공시",
    "E": "기타공시", "F": "외부감사", "G": "펀드공시", "H": "자산유동화",
    "I": "거래소공시", "J": "공정위공시",
}

CORP_CLS_DISPLAY = {"Y": "유가", "K": "코스닥", "N": "코넥스", "E": "기타"}

DART_REPORT_URL = "https://dart.fss.or.kr/dsaf001/main.do?rcpNo="

st.set_page_config(page_title="공시", page_icon="◻", layout="wide")
inject_css()

selected_date = st.session_state.get("selected_date")
date_str = st.session_state.get("date_str")

if not selected_date or not date_str:
    st.warning("메인 페이지에서 날짜를 먼저 선택해주세요.")
    st.stop()

page_header("DART 공시", title_date_text(selected_date))

# ---------------------------------------------------------------------------
# Load Data
# ---------------------------------------------------------------------------
data = load_disclosures(date_str)

market_rows = data.get("market_rows", [])
watchlist_rows = data.get("watchlist_rows", [])
summaries = data.get("summaries", [])
by_sheet = data.get("by_sheet", {})

# Load price data for 등락률 lookup
_price_list = load_price_data(date_str)
_price_by_code: dict[str, dict] = {}
_price_by_name: dict[str, dict] = {}
if _price_list:
    for _pr in _price_list:
        _code = str(_pr.get("code", "")).strip()
        _name = str(_pr.get("name", "")).strip()
        if _code:
            _price_by_code[_code] = _pr
        if _name:
            _price_by_name[_name] = _pr


def _get_pct(stock_code: str, corp_name: str) -> float | None:
    """Look up 등락률 by stock_code or corp_name. Returns float or None."""
    pr = _price_by_code.get(stock_code.strip()) if stock_code else None
    if not pr:
        pr = _price_by_name.get(corp_name.strip()) if corp_name else None
    if pr:
        return float(pr.get("pct", 0))
    return None


def _remark_with_pct(rm: str, stock_code: str, corp_name: str) -> str:
    """Combine 비고 text with colored 등락률."""
    pct = _get_pct(stock_code, corp_name)
    if pct is None:
        return rm or ""
    sign = "+" if pct > 0 else ""
    pct_str = f"{sign}{pct:.2f}%"
    parts = []
    if pct_str:
        parts.append(pct_str)
    if rm:
        parts.append(rm)
    return " | ".join(parts) if parts else ""


def _fmt_pct_val(stock_code: str, corp_name: str) -> str:
    """Return formatted 등락률 string for a separate column."""
    pct = _get_pct(stock_code, corp_name)
    if pct is None:
        return "-"
    sign = "+" if pct > 0 else ""
    return f"{sign}{pct:.2f}%"


def _style_pct_col(val):
    """Apply color styling to 등락률 cell."""
    if not val or val == "-":
        return "color: #9CA3AF"
    try:
        num = float(val.replace("%", "").replace("+", ""))
        if num > 0:
            return f"color: {COLOR_UP}; font-weight: 600"
        elif num < 0:
            return f"color: {COLOR_DOWN}; font-weight: 600"
    except (ValueError, AttributeError):
        pass
    return "color: #9CA3AF"

# ---------------------------------------------------------------------------
# KPI Cards
# ---------------------------------------------------------------------------
summary_map = {s["display_name"]: s for s in summaries}
cols = st.columns(4)

kpi_names = ["주요사항", "지분공시", "공정공시", "정기공시"]
for col, name in zip(cols, kpi_names):
    s = summary_map.get(name)
    count = s["total_count"] if s else 0
    col.metric(name, f"{count}건")

total = sum(s.get("total_count", 0) for s in summaries)
st.caption(f"전체 공시 합계: **{total}건**")

st.markdown("---")

# ---------------------------------------------------------------------------
# Watchlist Disclosures
# ---------------------------------------------------------------------------
section_header("관심종목 공시 현황")

if watchlist_rows:
    with st.container(border=True):
        wl_data = []
        for r in watchlist_rows:
            wl_data.append({
                "종목코드": r["stock_code"],
                "종목명": r["corp_name"],
                "공시유형": PBLNTF_TY_DISPLAY.get(r["pblntf_ty"], r["pblntf_ty"]),
                "세부분류": r["subcategory"],
                "보고서명": r["report_nm"],
                "DART 링크": f"{DART_REPORT_URL}{r['rcept_no']}",
            })
        wl_df = pd.DataFrame(wl_data)
        wl_df.index = wl_df.index + 1
        st.dataframe(
            wl_df,
            use_container_width=True,
            column_config={
                "DART 링크": st.column_config.LinkColumn("DART 링크", display_text="보기"),
            },
        )
        st.caption(f"총 {len(watchlist_rows)}건")
else:
    st.info("해당일 관심종목 공시가 없습니다.")

st.markdown("---")

# ---------------------------------------------------------------------------
# Market Disclosures by Category
# ---------------------------------------------------------------------------
section_header("전체 시장 공시")

sheet_tabs = st.tabs(list(by_sheet.keys()) if by_sheet else ["주요사항", "지분공시", "공정공시"])

for tab, sheet_name in zip(sheet_tabs, by_sheet.keys() if by_sheet else []):
    with tab:
        rows = by_sheet.get(sheet_name, [])
        if not rows:
            st.info(f"{sheet_name} 공시가 없습니다.")
            continue

        # Summary bar
        s = summary_map.get(sheet_name)
        if s and s.get("subcategory_counts"):
            sorted_subs = sorted(s["subcategory_counts"].items(), key=lambda x: -x[1])
            chips = " | ".join(f"**{k}** {v}건" for k, v in sorted_subs[:8])
            st.markdown(f"세부 분류: {chips}")

        # Data table
        table_data = []
        for r in rows:
            table_data.append({
                "시장": CORP_CLS_DISPLAY.get(r["corp_cls"], r["corp_cls"]),
                "종목명": r["corp_name"],
                "등락률": _fmt_pct_val(r.get("stock_code", ""), r["corp_name"]),
                "세부분류": r["subcategory"],
                "보고서명": r["report_nm"],
                "DART 링크": f"{DART_REPORT_URL}{r['rcept_no']}",
            })

        tdf = pd.DataFrame(table_data)

        # Search filter
        search = st.text_input(
            f"{sheet_name} 검색", key=f"search_{sheet_name}",
            placeholder="종목명, 보고서명 등으로 검색...",
        )
        if search:
            mask = (
                tdf["종목명"].str.contains(search, case=False, na=False)
                | tdf["보고서명"].str.contains(search, case=False, na=False)
                | tdf["세부분류"].str.contains(search, case=False, na=False)
            )
            tdf = tdf[mask]

        styled_tdf = tdf.style.map(_style_pct_col, subset=["등락률"])
        st.dataframe(
            styled_tdf,
            use_container_width=True,
            height=min(len(tdf) * 35 + 40, 600),
            column_config={
                "DART 링크": st.column_config.LinkColumn("DART 링크", display_text="보기"),
            },
            hide_index=True,
        )
        st.caption(f"총 {len(tdf)}건")
