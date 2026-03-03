"""Data fetching layer for the dashboard.

Wraps existing API clients from the automation scripts with Streamlit caching.
All heavy API calls are cached with TTL to avoid redundant requests.
"""
from __future__ import annotations

import json
import logging
import os
import shutil
import sys
import time
from datetime import date, timedelta
from pathlib import Path
from typing import Any

import streamlit as st

# Google Sheets (optional - Streamlit Cloud 환경에서 관심종목 영구 저장용)
try:
    import gspread
    from google.oauth2.service_account import Credentials as _GCredentials
    _GSHEETS_AVAILABLE = True
except ImportError:
    _GSHEETS_AVAILABLE = False

# Fix curl_cffi SSL cert issue on paths with non-ASCII characters (e.g. Korean)
if "CURL_CA_BUNDLE" not in os.environ:
    try:
        import certifi
        _cert_src = Path(certifi.where())
        if any(ord(c) > 127 for c in str(_cert_src)):
            _cert_dst = Path(os.environ.get("TEMP", "/tmp")) / "cacert.pem"
            if not _cert_dst.exists() or _cert_dst.stat().st_size != _cert_src.stat().st_size:
                shutil.copy2(_cert_src, _cert_dst)
            os.environ["CURL_CA_BUNDLE"] = str(_cert_dst)
    except Exception:
        pass

# Ensure project root is on sys.path so we can import from existing scripts
_PROJECT_DIR = Path(__file__).parent.resolve()
if str(_PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(_PROJECT_DIR))

# Import from existing automation scripts
from daily_krx_automation import (
    INVESTOR_SHEETS,
    build_client,
    fetch_with_retry,
    normalize_field_map,
    DEFAULT_PRICE_FIELD_MAP,
    DEFAULT_SUPPLY_FIELD_MAP,
    DEFAULT_HIGH_FIELD_MAP,
    parse_price_rows,
    parse_supply_rows,
    parse_high_rows,
    as_text_code,
    as_int,
    as_eok_from_won,
    load_json,
)

from daily_dart_automation import (
    DartClient,
    DisclosureRow,
    CategorySummary,
    PBLNTF_TY_DISPLAY,
    SHEET_TYPE_MAP,
    parse_disclosure_rows,
    build_summaries,
    group_by_sheet,
)

from daily_watchlist_automation import (
    DartFinancialClient,
    FinancialModel,
    PeriodLabel,
    IS_ACCOUNT_DEFS,
    BS_ACCOUNT_DEFS,
    MARGIN_ROWS,
    INVESTOR_TYPES,
    build_financial_model,
    calc_valuations,
    get_is_value,
    get_bs_value,
    get_yoy_value,
    safe_pct,
    to_eok,
    fetch_segment_data,
)


# ---------------------------------------------------------------------------
# Config Loading
# ---------------------------------------------------------------------------

def _load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _apply_secrets(cfg: dict[str, Any]) -> dict[str, Any]:
    """Override config values with st.secrets when available (for cloud deploy)."""
    try:
        secrets = st.secrets
        # DART API key
        if "dart" in secrets:
            dart_sec = secrets["dart"]
            if "api_key" in dart_sec:
                cfg.setdefault("dart", {})
                if not cfg["dart"].get("api_key"):
                    cfg["dart"]["api_key"] = dart_sec["api_key"]

        # KRX login credentials
        if "krx" in secrets:
            krx_sec = secrets["krx"]
            cfg.setdefault("krx", {})
            cfg["krx"].setdefault("login", {})
            if "mbrId" in krx_sec and not cfg["krx"]["login"].get("mbrId"):
                cfg["krx"]["login"]["mbrId"] = krx_sec["mbrId"]
            if "pw" in krx_sec and not cfg["krx"]["login"].get("pw"):
                cfg["krx"]["login"]["pw"] = krx_sec["pw"]
    except Exception:
        pass

    return cfg


@st.cache_data
def load_dart_config() -> dict[str, Any]:
    path = _PROJECT_DIR / "config.dart.json"
    fallback = _PROJECT_DIR / "config.dart.example.json"
    if path.exists():
        cfg = _load_json(path)
    elif fallback.exists():
        cfg = _load_json(fallback)
    else:
        return {}
    return _apply_secrets(cfg)


# ---------------------------------------------------------------------------
# KRX Data (KRX 직접 API 기반)
# ---------------------------------------------------------------------------


def _krx_client_and_config():
    """KRX 클라이언트 생성 + 세션 초기화 + 로그인."""
    logger = logging.getLogger("dashboard")
    config_path = _PROJECT_DIR / "config.krx.json"
    fallback_path = _PROJECT_DIR / "config.krx.example.json"
    if config_path.exists():
        cfg = load_json(config_path)
        logger.warning("[KRX] config.krx.json 사용")
    elif fallback_path.exists():
        cfg = load_json(fallback_path)
        logger.warning("[KRX] config.krx.example.json 사용 (폴백)")
    else:
        raise FileNotFoundError("KRX config not found")
    cfg = _apply_secrets(cfg)
    login_cfg = cfg.get("krx", {}).get("login", {})
    has_creds = bool(login_cfg.get("mbrId")) and bool(login_cfg.get("pw"))
    logger.warning(f"[KRX] 로그인 자격증명: {'있음' if has_creds else '없음'}")
    client = build_client(cfg)
    init_url = cfg["krx"].get(
        "login_page_url",
        "https://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd",
    )
    try:
        client.session.get(init_url, timeout=client.timeout_sec)
        logger.warning("[KRX] 세션 초기화 성공")
    except Exception as e:
        logger.warning(f"[KRX] 세션 초기화 실패: {e}")
    try:
        client.login()
        logger.warning("[KRX] 로그인 성공")
    except Exception as e:
        logger.warning(f"[KRX] 로그인 실패 (비로그인 모드로 계속): {e}")
    return client, cfg


@st.cache_data(show_spinner="KRX 시세 데이터 조회 중...")
def load_price_data(date_str: str) -> list[dict[str, Any]]:
    """KRX API로 전종목 시세 + 시가총액 데이터 조회."""
    try:
        client, cfg = _krx_client_and_config()
        datasets = cfg.get("datasets", {})
        price_cfg = datasets.get("price_all", {})
        price_fm = normalize_field_map(
            price_cfg.get("field_map", {}), DEFAULT_PRICE_FIELD_MAP,
        )
        raw = fetch_with_retry(client, price_cfg, date_str, "시세")
        rows = parse_price_rows(raw, price_fm)
        return [{
            "code": r.code, "name": r.name, "market": r.market,
            "sector": r.sector, "close": r.close, "change": r.change,
            "pct": r.pct, "open": r.open_price, "high": r.high,
            "low": r.low, "volume": r.volume, "trade_value": r.trade_value,
            "market_cap": r.market_cap, "listed_shares": r.listed_shares,
        } for r in rows]
    except Exception as e:
        print(f"[dashboard] 시세 데이터 로드 실패: {e}")
        return []


@st.cache_data(show_spinner="KRX 수급 데이터 조회 중...")
def load_supply_data(date_str: str) -> dict[str, list[dict[str, Any]]]:
    """KRX API로 투자자 유형별 종목 순매수 데이터 조회."""
    try:
        client, cfg = _krx_client_and_config()
    except Exception as e:
        print(f"[dashboard] KRX 클라이언트 생성 실패: {e}")
        return {}

    datasets = cfg.get("datasets", {})
    supply_cfgs = datasets.get("supply", {})
    supply_fm = normalize_field_map({}, DEFAULT_SUPPLY_FIELD_MAP)

    # 시가총액/시장구분 보강을 위한 시세 데이터 (캐시됨)
    price_data = load_price_data(date_str)
    mcap_map = {r["code"]: r["market_cap"] for r in price_data}
    mkt_map = {r["code"]: r["market"] for r in price_data}

    result: dict[str, list[dict[str, Any]]] = {}
    for sheet_name in INVESTOR_SHEETS:
        inv_cfg = supply_cfgs.get(sheet_name)
        if not inv_cfg:
            result[sheet_name] = []
            continue
        try:
            raw = fetch_with_retry(client, inv_cfg, date_str, f"수급-{sheet_name}")
            for row in raw:
                code = as_text_code(
                    row.get("ISU_SRT_CD", row.get("종목코드", ""))
                )
                if not any(f in row for f in supply_fm["market_cap"]):
                    row["시가총액"] = str(mcap_map.get(code, 0))
                if not any(f in row for f in supply_fm["market"]):
                    row["시장구분"] = mkt_map.get(code, "")
            rows = parse_supply_rows(raw, supply_fm)
            result[sheet_name] = [{
                "code": r.code, "name": r.name, "market": r.market,
                "net_buy": r.net_buy, "market_cap": r.market_cap,
                "ratio": r.ratio,
            } for r in rows]
        except Exception:
            result[sheet_name] = []
        time.sleep(0.5)
    return result


@st.cache_data(show_spinner="투자자별 매매동향 집계 중...")
def load_investor_trends(date_str: str) -> dict[str, Any]:
    """투자자 유형별 순매수를 코스피/코스닥 시장별로 집계.

    Returns dict:
        overview: {KOSPI: {외국인: amount, 기관: amount, 개인: amount},
                   KOSDAQ: {...}}
        detail:   {KOSPI: {사모펀드: amount, 투자신탁: amount, 연기금: amount},
                   KOSDAQ: {...}}
    """
    supply = load_supply_data(date_str)
    if not supply:
        return {"overview": {}, "detail": {}}

    # 기관 = 사모펀드 + 투자신탁 + 연기금
    institutional_types = {"사모펀드", "투자신탁", "연기금"}

    overview: dict[str, dict[str, int]] = {
        "KOSPI": {"외국인": 0, "기관": 0, "개인": 0},
        "KOSDAQ": {"외국인": 0, "기관": 0, "개인": 0},
    }
    detail: dict[str, dict[str, int]] = {
        "KOSPI": {},
        "KOSDAQ": {},
    }

    for inv_name, rows in supply.items():
        for market_key in ("KOSPI", "KOSDAQ"):
            total = sum(
                r.get("net_buy", 0) for r in rows
                if r.get("market", "") == market_key
            )

            # overview 집계
            if inv_name in institutional_types:
                overview[market_key]["기관"] += total
            elif inv_name == "외국인":
                overview[market_key]["외국인"] = total
            elif inv_name == "개인":
                overview[market_key]["개인"] = total

            # 기관 세부 (사모, 투신, 연기금)
            if inv_name in institutional_types:
                detail[market_key][inv_name] = total

    return {"overview": overview, "detail": detail}


@st.cache_data(show_spinner="KRX 신고가 데이터 조회 중...")
def load_high_data(date_str: str) -> dict[str, list[dict[str, Any]]]:
    """KRX API로 신고가 종목 조회."""
    try:
        client, cfg = _krx_client_and_config()
    except Exception as e:
        print(f"[dashboard] KRX 클라이언트 생성 실패: {e}")
        return {}

    datasets = cfg.get("datasets", {})
    high_cfgs = datasets.get("highs", {})
    high_fm = normalize_field_map({}, DEFAULT_HIGH_FIELD_MAP)
    market_cap_unit = str(high_cfgs.get("market_cap_unit", "won"))

    # 보강용 시세 룩업맵
    price_data = load_price_data(date_str)
    mcap_map = {r["code"]: r["market_cap"] for r in price_data}
    mkt_map = {r["code"]: r["market"] for r in price_data}

    result: dict[str, list[dict[str, Any]]] = {}
    for sheet_name, sheet_cfg in high_cfgs.items():
        if not isinstance(sheet_cfg, dict):
            continue
        if "request_params" not in sheet_cfg and "otp_params" not in sheet_cfg:
            continue
        try:
            raw = fetch_with_retry(client, sheet_cfg, date_str, f"신고가-{sheet_name}")
            filtered = [
                row for row in raw
                if as_int(row.get("TDD_CLSPRC", "0")) > 0
                and as_int(row.get("HGST_ADJ_CLSPRC", "0")) > 0
                and as_int(row.get("TDD_CLSPRC", "0"))
                >= as_int(row.get("HGST_ADJ_CLSPRC", "0"))
            ]
            raw = filtered if filtered else raw
            for row in raw:
                code = as_text_code(
                    row.get("ISU_CD", row.get("ISU_SRT_CD", row.get("종목코드", "")))
                )
                if not any(f in row for f in high_fm["market_cap"]):
                    row["시가총액"] = str(mcap_map.get(code, 0))
                if not any(f in row for f in high_fm["market"]):
                    row["시장구분"] = mkt_map.get(code, "")
            rows = parse_high_rows(raw, high_fm, market_cap_unit)
            result[sheet_name] = [{
                "code": r.code, "name": r.name, "market": r.market,
                "market_cap_eok": r.market_cap_eok, "pct": r.pct,
                "high_price": r.high_price,
            } for r in rows]
        except Exception:
            result[sheet_name] = []
        time.sleep(0.5)

    # 신고가 API 실패 시 시세 데이터 기반 폴백
    if not result:
        candidates = []
        for p in price_data:
            if p["pct"] <= 0 or p["high"] <= 0 or p["close"] <= 0:
                continue
            if p["close"] < p["high"] * 0.99:
                continue
            candidates.append({
                "code": p["code"], "name": p["name"], "market": p["market"],
                "market_cap_eok": as_eok_from_won(p["market_cap"]),
                "pct": p["pct"], "high_price": p["high"],
            })
        candidates.sort(key=lambda x: x["pct"], reverse=True)
        result["당일 강세 종목"] = candidates

    return result


# ---------------------------------------------------------------------------
# DART Disclosures
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="DART 공시 데이터 조회 중...")
def load_disclosures(date_str: str) -> dict[str, Any]:
    """Load market-wide and watchlist disclosures.

    Returns dict with keys:
        market_rows, watchlist_rows, summaries, by_sheet
    """
    cfg = load_dart_config()
    if not cfg:
        return {"market_rows": [], "watchlist_rows": [], "summaries": [], "by_sheet": {}}

    dart_cfg = cfg.get("dart", {})
    api_key = dart_cfg.get("api_key", "")
    if not api_key:
        return {"market_rows": [], "watchlist_rows": [], "summaries": [], "by_sheet": {}}

    classification = cfg.get("classification", {})
    pblntf_types = dart_cfg.get("pblntf_types", ["A", "B", "D", "E", "I"])
    watchlist_cfg = get_watchlist()

    client = DartClient(
        api_key=api_key,
        base_url=dart_cfg.get("base_url", "https://opendart.fss.or.kr/api"),
        timeout_sec=int(dart_cfg.get("timeout_sec", 30)),
        page_size=int(dart_cfg.get("page_size", 100)),
        request_delay=float(dart_cfg.get("request_delay_sec", 0.5)),
    )

    # Market-wide
    raw_market = client.fetch_all_types(
        bgn_de=date_str, end_de=date_str, pblntf_types=pblntf_types,
    )
    market_rows = parse_disclosure_rows(raw_market, classification)

    # Watchlist
    watchlist_rows: list[DisclosureRow] = []
    if watchlist_cfg:
        raw_wl = client.fetch_watchlist(
            bgn_de=date_str, end_de=date_str, watchlist=watchlist_cfg,
        )
        watchlist_rows = parse_disclosure_rows(raw_wl, classification)

    summaries = build_summaries(market_rows, SHEET_TYPE_MAP)
    by_sheet = group_by_sheet(market_rows, SHEET_TYPE_MAP)

    # Convert to serializable dicts
    def _row_to_dict(r: DisclosureRow) -> dict[str, Any]:
        return {
            "corp_code": r.corp_code, "corp_name": r.corp_name,
            "stock_code": r.stock_code, "corp_cls": r.corp_cls,
            "report_nm": r.report_nm, "rcept_no": r.rcept_no,
            "flr_nm": r.flr_nm, "rcept_dt": r.rcept_dt,
            "rm": r.rm, "pblntf_ty": r.pblntf_ty,
            "subcategory": r.subcategory,
            "subcategory_priority": r.subcategory_priority,
        }

    def _summary_to_dict(s: CategorySummary) -> dict[str, Any]:
        return {
            "pblntf_ty": s.pblntf_ty,
            "display_name": s.display_name,
            "total_count": s.total_count,
            "subcategory_counts": s.subcategory_counts,
        }

    return {
        "market_rows": [_row_to_dict(r) for r in market_rows],
        "watchlist_rows": [_row_to_dict(r) for r in watchlist_rows],
        "summaries": [_summary_to_dict(s) for s in summaries],
        "by_sheet": {
            k: [_row_to_dict(r) for r in v]
            for k, v in by_sheet.items()
        },
    }


# ---------------------------------------------------------------------------
# Watchlist / Financial Model
# ---------------------------------------------------------------------------

def _use_gsheets() -> bool:
    """Google Sheets 사용 가능 여부 확인."""
    if not _GSHEETS_AVAILABLE:
        return False
    try:
        return "gcp_service_account" in st.secrets and "gsheets" in st.secrets
    except Exception:
        return False


@st.cache_resource
def _get_gsheet():
    """Google Sheets 워크시트 연결 (캐시됨)."""
    creds = _GCredentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    gc = gspread.authorize(creds)
    spreadsheet_id = st.secrets["gsheets"]["spreadsheet_id"]
    sh = gc.open_by_key(spreadsheet_id)
    return sh.worksheet("watchlist")


def get_watchlist() -> list[dict[str, str]]:
    """관심종목 목록 반환. Google Sheets 우선, 없으면 config.dart.json 폴백."""
    if _use_gsheets():
        try:
            ws = _get_gsheet()
            records = ws.get_all_records()
            return [
                {
                    "stock_code": str(r.get("stock_code", "")).strip().zfill(6),
                    "corp_code": str(r.get("corp_code", "")).strip(),
                    "name": str(r.get("name", "")).strip(),
                }
                for r in records
                if r.get("stock_code")
            ]
        except Exception as e:
            st.warning(f"Google Sheets 읽기 실패, config 폴백: {e}")
    cfg = load_dart_config()
    return cfg.get("watchlist", [])


def get_financial_years() -> list[int]:
    cfg = load_dart_config()
    wo = cfg.get("watchlist_output", {})
    current_year = date.today().year
    return wo.get("financial_years", [current_year - 1, current_year])


@st.cache_data(show_spinner="재무 모델 구축 중...")
def load_financial_model_cached(
    corp_code: str,
    corp_name: str,
    stock_code: str,
    years_json: str,
    fs_div: str,
) -> dict[str, Any]:
    """Build financial model and return as serializable dict."""
    cfg = load_dart_config()
    dart_cfg = cfg.get("dart", {})
    api_key = dart_cfg.get("api_key", "")

    client = DartFinancialClient(
        api_key=api_key,
        timeout_sec=int(dart_cfg.get("timeout_sec", 30)),
        delay=float(dart_cfg.get("request_delay_sec", 0.5)),
    )

    years = json.loads(years_json)
    model = build_financial_model(
        client, corp_code, corp_name, stock_code, years, fs_div
    )

    # Convert to serializable dict
    periods_list = [
        {"year": p.year, "quarter": p.quarter, "reprt_code": p.reprt_code}
        for p in model.periods
    ]

    return {
        "corp_name": model.corp_name,
        "stock_code": model.stock_code,
        "periods": periods_list,
        "is_data": model.is_data,
        "bs_data": model.bs_data,
    }


def dict_to_model(d: dict[str, Any]) -> FinancialModel:
    """Convert cached dict back to FinancialModel."""
    model = FinancialModel(
        corp_name=d["corp_name"],
        stock_code=d["stock_code"],
    )
    model.periods = [
        PeriodLabel(year=p["year"], quarter=p["quarter"], reprt_code=p["reprt_code"])
        for p in d["periods"]
    ]
    model.is_data = d["is_data"]
    model.bs_data = d["bs_data"]
    return model


@st.cache_data(show_spinner="공시 조회 중...")
def load_watchlist_disclosures(
    corp_code: str, date_str: str,
) -> list[dict[str, str]]:
    cfg = load_dart_config()
    dart_cfg = cfg.get("dart", {})
    api_key = dart_cfg.get("api_key", "")
    if not api_key:
        return []

    client = DartFinancialClient(
        api_key=api_key,
        timeout_sec=int(dart_cfg.get("timeout_sec", 30)),
        delay=float(dart_cfg.get("request_delay_sec", 0.5)),
    )

    target = date.today()
    try:
        from datetime import datetime
        target = datetime.strptime(date_str, "%Y%m%d").date()
    except Exception:
        pass

    bgn_de = (target - timedelta(days=30)).strftime("%Y%m%d")
    end_de = date_str
    return client.fetch_disclosures(corp_code, bgn_de, end_de, page_count=10)


@st.cache_data(show_spinner="KRX 관심종목 시세/수급 조회 중...")
def load_watchlist_krx(date_str: str) -> dict[str, Any]:
    """KRX 직접 API로 관심종목의 시세/수급 데이터 조회."""
    watchlist = get_watchlist()
    watchlist_codes = {w["stock_code"] for w in watchlist}
    if not watchlist_codes:
        return {"price_rows": [], "supply_by_type": {}}

    # 전종목 시세에서 관심종목만 필터 (캐시됨)
    all_prices = load_price_data(date_str)
    price_rows = [r for r in all_prices if r["code"] in watchlist_codes]

    # 전종목 수급에서 관심종목만 필터 (캐시됨)
    all_supply = load_supply_data(date_str)
    supply_by_type: dict[str, list[dict[str, Any]]] = {}
    for inv_name, rows in all_supply.items():
        supply_by_type[inv_name] = [r for r in rows if r["code"] in watchlist_codes]

    return {
        "price_rows": price_rows,
        "supply_by_type": supply_by_type,
    }


# ---------------------------------------------------------------------------
# Historical Price Data (yfinance)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="주가 차트 데이터 조회 중...")
def load_price_history(
    stock_code: str, market: str = "KS", period: str = "1y",
) -> list[dict[str, Any]]:
    """Fetch OHLCV history from Yahoo Finance.

    Args:
        stock_code: 6-digit KRX stock code (e.g. '005930')
        market: 'KS' for KOSPI, 'KQ' for KOSDAQ
        period: yfinance period string ('3mo', '6mo', '1y', '2y', '5y')
    """
    import yfinance as yf

    ticker = f"{stock_code}.{market}"
    try:
        data = yf.download(ticker, period=period, progress=False, auto_adjust=True)
        if data.empty:
            return []
        # Handle MultiIndex columns from yfinance
        if hasattr(data.columns, 'levels'):
            data.columns = data.columns.get_level_values(0)
        records = []
        for dt, row in data.iterrows():
            records.append({
                "date": dt.strftime("%Y-%m-%d"),
                "open": float(row["Open"]),
                "high": float(row["High"]),
                "low": float(row["Low"]),
                "close": float(row["Close"]),
                "volume": int(row["Volume"]),
            })
        return records
    except Exception:
        return []


# ---------------------------------------------------------------------------
# Watchlist Management (Google Sheets 기반 - Streamlit Cloud 영구 저장)
# ---------------------------------------------------------------------------

def _config_path() -> Path:
    return _PROJECT_DIR / "config.dart.json"


def add_watchlist_stock(stock_code: str, corp_code: str, name: str) -> bool:
    """관심종목 추가. Google Sheets 우선, 없으면 config.dart.json 폴백."""
    if _use_gsheets():
        try:
            ws = _get_gsheet()
            existing_codes = ws.col_values(1)[1:]  # 헤더(1행) 제외
            if stock_code in existing_codes:
                return False
            ws.append_row([stock_code, corp_code, name])
            _get_gsheet.clear()
            load_watchlist_krx.clear()
            return True
        except Exception as e:
            st.error(f"Google Sheets 종목 추가 실패: {e}")
            return False
    path = _config_path()
    if not path.exists():
        return False
    cfg = _load_json(path)
    watchlist = cfg.get("watchlist", [])
    if any(w.get("stock_code") == stock_code for w in watchlist):
        return False
    watchlist.append({"stock_code": stock_code, "corp_code": corp_code, "name": name})
    cfg["watchlist"] = watchlist
    with path.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    load_dart_config.clear()
    load_watchlist_krx.clear()
    return True


def remove_watchlist_stock(stock_code: str) -> bool:
    """관심종목 제거. Google Sheets 우선, 없으면 config.dart.json 폴백."""
    if _use_gsheets():
        try:
            ws = _get_gsheet()
            cell = ws.find(stock_code, in_column=1)
            if not cell:
                return False
            ws.delete_rows(cell.row)
            _get_gsheet.clear()
            load_watchlist_krx.clear()
            return True
        except Exception as e:
            st.error(f"Google Sheets 종목 제거 실패: {e}")
            return False
    path = _config_path()
    if not path.exists():
        return False
    cfg = _load_json(path)
    watchlist = cfg.get("watchlist", [])
    new_watchlist = [w for w in watchlist if w.get("stock_code") != stock_code]
    if len(new_watchlist) == len(watchlist):
        return False
    cfg["watchlist"] = new_watchlist
    with path.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    load_dart_config.clear()
    load_watchlist_krx.clear()
    return True


@st.cache_data(show_spinner="DART corp_code 조회 중...")
def lookup_corp_code(stock_code: str) -> dict[str, str] | None:
    """Look up corp_code from DART corpCode.xml for a given stock_code."""
    import io
    import zipfile
    from xml.etree import ElementTree as ET

    cfg = load_dart_config()
    dart_cfg = cfg.get("dart", {})
    api_key = dart_cfg.get("api_key", "")
    if not api_key:
        return None

    import requests
    try:
        resp = requests.get(
            "https://opendart.fss.or.kr/api/corpCode.xml",
            params={"crtfc_key": api_key},
            timeout=60,
        )
        resp.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            xml_name = zf.namelist()[0]
            root = ET.fromstring(zf.read(xml_name))

        target = stock_code.strip().zfill(6)
        for item in root.iter("list"):
            sc = (item.findtext("stock_code") or "").strip()
            if sc == target:
                cc = (item.findtext("corp_code") or "").strip()
                cn = (item.findtext("corp_name") or "").strip()
                return {"stock_code": sc, "corp_code": cc, "name": cn}
        return None
    except Exception:
        return None


@st.cache_data(ttl=3600, show_spinner="종목명 검색 중...")
def search_corp_by_name(query: str) -> list[dict[str, str]]:
    """Search DART corpCode.xml by company name (substring match).

    Returns up to 20 matches with non-empty stock_code (listed companies only).
    """
    import io
    import zipfile
    from xml.etree import ElementTree as ET

    cfg = load_dart_config()
    dart_cfg = cfg.get("dart", {})
    api_key = dart_cfg.get("api_key", "")
    if not api_key:
        return []

    import requests
    try:
        resp = requests.get(
            "https://opendart.fss.or.kr/api/corpCode.xml",
            params={"crtfc_key": api_key},
            timeout=60,
        )
        resp.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            xml_name = zf.namelist()[0]
            root = ET.fromstring(zf.read(xml_name))

        query_lower = query.strip().lower()
        results: list[dict[str, str]] = []
        for item in root.iter("list"):
            sc = (item.findtext("stock_code") or "").strip()
            if not sc or sc == "000000":
                continue
            cn = (item.findtext("corp_name") or "").strip()
            if query_lower in cn.lower():
                cc = (item.findtext("corp_code") or "").strip()
                results.append({"stock_code": sc, "corp_code": cc, "name": cn})
                if len(results) >= 20:
                    break
        return results
    except Exception:
        return []


@st.cache_data(ttl=600, show_spinner=False)
def load_index_history(period: str = "1y") -> list[dict[str, Any]]:
    """Fetch KOSPI index daily history from Yahoo Finance for benchmark comparison."""
    import yfinance as yf

    try:
        data = yf.download("^KS11", period=period, progress=False, auto_adjust=True)
        if data.empty:
            return []
        if hasattr(data.columns, "levels"):
            data.columns = data.columns.get_level_values(0)
        return [
            {"date": dt.strftime("%Y-%m-%d"), "close": float(row["Close"])}
            for dt, row in data.iterrows()
        ]
    except Exception:
        return []


# ---------------------------------------------------------------------------
# Naver Finance: Forward PER/EPS (Consensus)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="네이버 금융 밸류에이션 조회 중...")
def load_naver_valuations(stock_code: str) -> dict[str, Any]:
    """Scrape PER, EPS, forward PER/EPS, PBR, BPS, dividend yield from Naver Finance."""
    import requests
    from bs4 import BeautifulSoup

    result: dict[str, Any] = {
        "per": None, "eps": None,
        "forward_per": None, "forward_eps": None,
        "pbr": None, "bps": None,
        "dividend_yield": None,
        "industry_per": None,
    }

    url = f"https://finance.naver.com/item/main.nhn?code={stock_code}"
    try:
        resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        def _parse_num(elem_id: str) -> float | None:
            em = soup.find("em", id=elem_id)
            if em:
                txt = em.get_text().strip().replace(",", "")
                try:
                    return float(txt)
                except ValueError:
                    return None
            return None

        result["per"] = _parse_num("_per")
        result["eps"] = _parse_num("_eps")
        result["forward_per"] = _parse_num("_cns_per")
        result["forward_eps"] = _parse_num("_cns_eps")
        result["pbr"] = _parse_num("_pbr")
        result["bps"] = _parse_num("_bps")

        # Dividend yield
        for dt in soup.find_all("em", id="_dvr"):
            txt = dt.get_text().strip().replace("%", "").replace(",", "")
            try:
                result["dividend_yield"] = float(txt)
            except ValueError:
                pass

        # Industry PER
        per_table = soup.find("table", summary="동일업종 PER 정보")
        if per_table:
            em = per_table.find("em")
            if em:
                txt = em.get_text().strip().replace(",", "")
                try:
                    result["industry_per"] = float(txt)
                except ValueError:
                    pass

        return result
    except Exception:
        return result


# ---------------------------------------------------------------------------
# DART Segment Revenue
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="사업부문 매출 조회 중...")
def load_segment_data(
    corp_code: str, year: int, reprt_code: str = "11011",
) -> list[dict[str, Any]] | None:
    """DART에서 사업부문별 매출 데이터 조회.

    1차: 재무제표 API (기존 방식)
    2차: 보고서 본문 HTML에서 부문별 매출 테이블 파싱

    Args:
        reprt_code: "11013"=Q1, "11012"=반기, "11014"=Q3, "11011"=연간
    """
    cfg = load_dart_config()
    dart_cfg = cfg.get("dart", {})
    api_key = dart_cfg.get("api_key", "")
    if not api_key:
        return None

    client = DartFinancialClient(
        api_key=api_key,
        timeout_sec=int(dart_cfg.get("timeout_sec", 30)),
        delay=float(dart_cfg.get("request_delay_sec", 0.5)),
    )

    # 1차: 재무제표 API
    result = fetch_segment_data(client, corp_code, year, reprt_code)
    if result:
        return result

    # 2차: 보고서 본문에서 파싱
    return _parse_segment_from_report(api_key, corp_code, year, reprt_code)


def _parse_segment_from_report(
    api_key: str, corp_code: str, year: int, reprt_code: str = "11011",
) -> list[dict[str, Any]] | None:
    """DART 보고서 ZIP에서 부문별 매출 테이블 파싱.

    지원 테이블 형태:
    1) 제품현황형 — '사업부문|매출유형|품목|매출액(비율)' 형태 (한섬, LG 등)
    2) 가로형 — 헤더에 부문명, 매출 행에 숫자 (삼성전자 등 대기업)
    3) 세로형 — 부문명이 행 방향, 당기/전기 열 방향

    Args:
        reprt_code: "11013"=Q1, "11012"=반기, "11014"=Q3, "11011"=연간
    """
    import re
    import requests
    import zipfile
    import io
    from bs4 import BeautifulSoup
    import warnings

    try:
        from bs4 import XMLParsedAsHTMLWarning
        warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
    except ImportError:
        pass

    # ── 공통 유틸리티 ──

    def _detect_unit_multiplier(text: str) -> int:
        """텍스트에서 단위를 감지하여 원 단위로 변환하는 승수 반환."""
        unit_patterns = [
            (r"단위\s*[:\s=]\s*백만\s*원", 1_000_000),
            (r"단위\s*[:\s=]\s*천\s*원", 1_000),
            (r"단위\s*[:\s=]\s*억\s*원", 100_000_000),
            (r"단위\s*[:\s=]\s*원", 1),
            (r"\(백만원\)", 1_000_000),
            (r"\(천원\)", 1_000),
            (r"\(억원\)", 100_000_000),
            (r"\(백만\s*원\)", 1_000_000),
            (r"\(천\s*원\)", 1_000),
        ]
        for pattern, multiplier in unit_patterns:
            if re.search(pattern, text):
                return multiplier
        return 1_000_000

    def _table_to_grid(table) -> list[list[str]]:
        """HTML table → 2D 텍스트 grid (rowspan/colspan 처리)."""
        trs = table.find_all("tr")
        if not trs:
            return []
        max_cols = 0
        for tr in trs:
            cc = sum(int(c.get("colspan", 1)) for c in tr.find_all(["td", "th"]))
            max_cols = max(max_cols, cc)
        if max_cols == 0:
            return []

        grid: list[list[str | None]] = [[None] * max_cols for _ in range(len(trs))]
        for ri, tr in enumerate(trs):
            ci = 0
            for cell in tr.find_all(["td", "th"]):
                while ci < max_cols and grid[ri][ci] is not None:
                    ci += 1
                if ci >= max_cols:
                    break
                txt = cell.get_text().strip().replace("\n", " ").replace("\r", "")
                txt = re.sub(r"\s+", " ", txt)
                rs = int(cell.get("rowspan", 1))
                cs = int(cell.get("colspan", 1))
                for dr in range(rs):
                    for dc in range(cs):
                        r, c = ri + dr, ci + dc
                        if r < len(trs) and c < max_cols:
                            grid[r][c] = txt
                ci += cs

        return [[(v or "") for v in row] for row in grid]

    def _extract_number(s: str) -> int | None:
        """다양한 형태의 숫자 추출.

        - '732,884(71.29)' → 732884
        - '(28,515)' → -28515
        - '△28,515' → -28515
        - '-28,515' → -28515
        - '1,234 백만원' → 1234
        - '-' → None
        """
        s = s.strip()
        if not s or s in ("-", "—", "―", "‐"):
            return None
        # 음수 표기 감지
        neg = bool(re.match(r"^\s*[\(（△\-]", s))
        # 괄호 안의 비율 부분 제거: '732,884(71.29)' → '732,884'
        # 단, 전체가 괄호로 둘러싸인 경우(음수)는 보존
        if "(" in s and not s.lstrip().startswith("("):
            s = s.split("(")[0]
        # 숫자와 콤마, 소수점만 남기기
        cleaned = re.sub(r"[^\d.,]", "", s)
        cleaned = cleaned.strip(".")
        if not cleaned:
            return None
        try:
            val = int(float(cleaned.replace(",", "")))
        except ValueError:
            return None
        return -val if neg else val

    # ── 스킵 판단 ──

    _SKIP_EXACT = {"합계", "소계", "계", "총계", "총합계", "전사합계", "연결합계"}
    _SKIP_KEYWORDS = [
        "내부거래", "연결조정", "부문간제거", "부문간거래",
        "내부매출", "상계", "중복제거", "부문간",
    ]

    # 재무제표 테이블 식별 키워드 — 이런 키워드가 포함된 테이블은 세그먼트 테이블이 아님
    _FS_TABLE_KEYWORDS = [
        "유동자산", "비유동자산", "자산총계", "부채총계", "자본총계",
        "유동부채", "비유동부채", "이익잉여금", "자본금",
        "재무상태표", "손익계산서", "현금흐름표", "자본변동표",
        "포괄손익", "영업활동", "투자활동", "재무활동",
        "당기순이익", "당기순손실", "기본주당", "희석주당",
        "매출원가", "판매비와관리비", "판관비",
        "감가상각", "무형자산상각", "대손상각",
        "법인세", "법인세비용",
        "공정가치", "상각후원가", "측정금융",
        "기초잔액", "기말잔액",  # 자본변동표
        "총자산회전율", "자기자본", "부채비율",
        "제조원가명세서",
    ]

    # 재무제표 행 항목 (첫 열) — 이런 항목이 많으면 재무제표
    _FS_ROW_ITEMS = {
        "유동자산", "비유동자산", "자산총계",
        "유동부채", "비유동부채", "부채총계",
        "자본총계", "이익잉여금", "자본금",
        "매출액", "매출원가", "매출총이익",
        "영업이익", "영업손실", "당기순이익", "당기순손실",
        "판매비와관리비",
    }

    def _is_skip_name(name: str) -> bool:
        ns = re.sub(r"\s+", "", name)
        if ns in _SKIP_EXACT:
            return True
        if any(kw in ns for kw in _SKIP_KEYWORDS):
            return True
        # 세그먼트 이름이 너무 길면 (설명문/수식 등) 스킵
        if len(ns) > 25:
            return True
        # 숫자/기호만으로 된 이름 스킵
        if re.match(r"^[\d,.%()（）\-\s]+$", name.strip()):
            return True
        return False

    def _is_financial_statement_table(grid: list[list[str]], ttext: str) -> bool:
        """테이블이 재무제표(BS/IS/CF)인지 판단.

        단, 세그먼트 관련 키워드가 포함된 테이블은 제외하지 않음.
        """
        ttext_ns = ttext.replace(" ", "")

        # 세그먼트 관련 키워드가 있으면 → 재무제표가 아닌 세그먼트 표일 가능성
        seg_keywords = ["부문", "사업부문", "세그먼트", "사업영역", "DX", "DS", "SDC"]
        has_seg = any(kw in ttext_ns for kw in seg_keywords)
        if has_seg:
            return False

        # 재무제표 키워드가 4개 이상 (세그먼트 키워드 없을 때만) → 재무제표
        fs_hit = sum(1 for kw in _FS_TABLE_KEYWORDS if kw in ttext_ns)
        if fs_hit >= 4:
            return True

        # 첫 열 값 중 재무제표 항목이 많으면 → 재무제표
        first_col_items = set()
        for row in grid:
            if row:
                item = re.sub(r"\s+", "", row[0])
                first_col_items.add(item)
        fs_row_hit = sum(1 for item in first_col_items if item in _FS_ROW_ITEMS)
        if fs_row_hit >= 3:
            return True

        return False

    # ── 헤더 키워드 패턴 ──

    _SEG_HDR_PATTERNS = [
        "사업부문", "사업 부문", "부문명", "사업영역", "사업구분",
        "부 문", "세그먼트", "사업분야", "부문구분",
    ]

    def _is_seg_header(text: str) -> bool:
        ns = text.replace(" ", "")
        return any(p.replace(" ", "") in ns for p in _SEG_HDR_PATTERNS)

    def _is_rev_header(text: str) -> bool:
        """매출액 열 헤더인지 판단 (매출원가 등 제외)."""
        ns = text.replace(" ", "")
        # 제외 키워드
        exclude = ["원가", "이익", "비용", "총이익", "총손실", "총액", "채권", "채무",
                    "자산", "부채", "상각", "손실", "세금", "법인세"]
        if any(kw in ns for kw in exclude):
            return False
        return "매출액" in ns or ns == "매출" or "매출금액" in ns or "매출실적" in ns

    _TYPE_HDR_PATTERNS = ["매출유형", "매출 유형", "유형"]

    def _is_type_header(text: str) -> bool:
        ns = text.replace(" ", "")
        return any(p.replace(" ", "") in ns for p in _TYPE_HDR_PATTERNS)

    # ── 당기 열 식별 ──

    def _find_current_period_col(
        hdr_rows: list[list[str]], candidate_cols: list[int], target_year: int,
    ) -> int | None:
        """여러 매출 열 후보 중 '당기' (가장 최근 기간) 열을 식별.

        전략:
        1) 'target_year' 또는 '당기' 키워드가 헤더에 있는 열
        2) '제XX기' 에서 숫자가 가장 큰 열
        3) 첫 번째 매출 열 (기본)
        """
        if not candidate_cols:
            return None
        if len(candidate_cols) == 1:
            return candidate_cols[0]

        # 모든 헤더 행에서 열별 텍스트 수집
        col_texts: dict[int, str] = {}
        for col_idx in candidate_cols:
            parts = []
            for hdr in hdr_rows:
                if col_idx < len(hdr):
                    parts.append(hdr[col_idx])
            col_texts[col_idx] = " ".join(parts)

        # 1) '당기' 키워드 or 대상 연도
        for col_idx in candidate_cols:
            txt = col_texts[col_idx]
            if "당기" in txt or "당 기" in txt:
                return col_idx
            if str(target_year) in txt:
                return col_idx

        # 2) 제XX기 → 숫자가 가장 큰 것 = 당기
        period_nums: list[tuple[int, int]] = []
        for col_idx in candidate_cols:
            txt = col_texts[col_idx]
            m = re.search(r"제\s*(\d+)\s*기", txt)
            if m:
                period_nums.append((int(m.group(1)), col_idx))
        if period_nums:
            period_nums.sort(key=lambda x: x[0], reverse=True)
            return period_nums[0][1]

        # 3) 기본: 첫 번째 매출 열
        return candidate_cols[0]

    # ── 제품현황형 테이블 파싱 ──

    def _try_product_table(grid: list[list[str]], target_year: int) -> list[dict[str, Any]] | None:
        """'주요 제품 등의 현황' 형태 테이블 파싱.

        헤더 패턴: 사업부문 | 매출유형 | 품목 | 매출액(비율) …
        값 형태: '732,884(71.29)'
        """
        if len(grid) < 3:
            return None

        # 헤더 행 탐색 (처음 7행까지)
        header_idx = None
        for i, row in enumerate(grid[:7]):
            joined = " ".join(row)
            has_seg = any(_is_seg_header(cell) for cell in row)
            has_type = any(_is_type_header(cell) for cell in row)
            has_rev = any(_is_rev_header(cell) for cell in row)
            if has_seg or has_type or (has_rev and ("부문" in joined or "품목" in joined)):
                header_idx = i
                break
        if header_idx is None:
            return None

        # 서브헤더 행 감지 (2~3행에 걸치는 헤더)
        hdr_rows = [grid[header_idx]]
        for offset in range(1, 3):
            next_idx = header_idx + offset
            if next_idx >= len(grid):
                break
            next_row = grid[next_idx]
            next_joined = " ".join(next_row)
            is_sub = (
                any(_is_rev_header(cell) for cell in next_row)
                or "비율" in next_joined
                or re.search(r"제\s?\d+\s?기", next_joined)
                or "당기" in next_joined or "전기" in next_joined
            )
            # 서브헤더에 숫자 데이터가 많으면 데이터 행
            num_count = sum(1 for c in next_row if _extract_number(c) is not None)
            if is_sub and num_count <= len(next_row) // 2:
                hdr_rows.append(next_row)
            else:
                break

        # 열 인덱스 식별
        seg_col = type_col = None
        rev_cols: list[int] = []

        # 모든 헤더 행에서 탐색
        for hdr in hdr_rows:
            for ci, h in enumerate(hdr):
                if _is_seg_header(h) and seg_col is None:
                    seg_col = ci
                elif _is_type_header(h) and type_col is None:
                    type_col = ci
                elif _is_rev_header(h) and ci not in rev_cols:
                    rev_cols.append(ci)

        # 당기 매출 열 선택
        rev_col = _find_current_period_col(hdr_rows, rev_cols, target_year)
        if rev_col is None:
            return None

        # 데이터 시작 행
        data_start = header_idx + len(hdr_rows)
        # 추가 기수/단위 행 스킵
        while data_start < len(grid):
            row_joined = " ".join(grid[data_start])
            row_stripped = row_joined.replace(" ", "")
            if (
                re.search(r"제\s?\d+\s?기|당\s?기|전\s?기", row_joined)
                and not re.search(r"\d{3,}", row_joined.replace(",", ""))
            ) or "단위" in row_stripped:
                data_start += 1
            else:
                break

        # 데이터 행 파싱
        segments_by_name: dict[str, int] = {}
        current_segment = ""
        for ri in range(data_start, len(grid)):
            row = grid[ri]
            if len(row) <= rev_col:
                continue

            # 사업부문 열 값 (빈 값이면 이전 부문 유지 = rowspan 처리)
            if seg_col is not None and seg_col < len(row):
                sname = row[seg_col].strip()
                if sname and not _is_skip_name(sname):
                    current_segment = sname

            # 합계/소계 행 스킵
            row_text = " ".join(row)
            if re.search(r"합\s*계|소\s*계|총\s*계", row_text):
                continue

            # 매출액 추출
            rev = _extract_number(row[rev_col])
            if rev is None or rev <= 0:
                continue

            # 그룹 키 결정
            group_key = current_segment
            if seg_col is not None:
                group_key = current_segment or "기타"
            elif type_col is not None and type_col < len(row):
                group_key = row[type_col].strip() or "기타"
            else:
                # 사업부문/매출유형 열이 없으면 품목 열 사용
                for ci in range(len(row)):
                    if ci == rev_col:
                        continue
                    cell_val = row[ci].strip()
                    if cell_val and _extract_number(cell_val) is None and not _is_skip_name(cell_val):
                        group_key = cell_val
                        break

            if _is_skip_name(group_key):
                continue
            segments_by_name[group_key] = segments_by_name.get(group_key, 0) + rev

        # 사업부문이 1개뿐이면 → 매출유형 또는 품목별로 재분류
        if len(segments_by_name) <= 1:
            alt_col = type_col
            # type_col 없으면 품목 열 탐색
            if alt_col is None:
                for hdr in hdr_rows:
                    for ci, h in enumerate(hdr):
                        hn = h.replace(" ", "")
                        if ci != seg_col and ci != rev_col and ("품목" in hn or "제품" in hn or "서비스" in hn):
                            alt_col = ci
                            break
                    if alt_col is not None:
                        break
            if alt_col is not None:
                segments_by_name = {}
                for ri in range(data_start, len(grid)):
                    row = grid[ri]
                    if len(row) <= rev_col:
                        continue
                    row_text = " ".join(row)
                    if re.search(r"합\s*계|소\s*계|총\s*계", row_text):
                        continue
                    rev = _extract_number(row[rev_col])
                    if rev is None or rev <= 0:
                        continue
                    tname = row[alt_col].strip() if alt_col < len(row) else ""
                    if not tname or _is_skip_name(tname):
                        continue
                    segments_by_name[tname] = segments_by_name.get(tname, 0) + rev

        if len(segments_by_name) < 2:
            return None

        total = sum(segments_by_name.values())
        result = [
            {"name": n, "revenue": v, "pct": v / total * 100 if total else 0}
            for n, v in segments_by_name.items()
        ]
        return sorted(result, key=lambda x: x["revenue"], reverse=True)

    # ── 가로형 세그먼트 테이블 파싱 ──

    def _try_horizontal_segment(grid: list[list[str]], target_year: int) -> list[dict[str, Any]] | None:
        """가로형 테이블 파싱 (삼성전자 형태).

        헤더: 구분 | DX부문 | DS부문 | SDC | Harman | 내부거래 | 계
        매출:      | 174..  | 111..  | 29.. | 14..  | (28..) | 300..
        """
        if len(grid) < 2:
            return None

        # 매출 행 찾기 (여러 개일 수 있음 → 당기 매출 행 선택)
        rev_candidates: list[int] = []
        for ri, row in enumerate(grid):
            for cell in row:
                cn = cell.replace(" ", "")
                if _is_rev_header(cn):
                    rev_candidates.append(ri)
                    break
        if not rev_candidates:
            return None

        # 매출 행이 여러 개면 당기/최근 것 선택
        rev_ri = rev_candidates[0]
        if len(rev_candidates) > 1:
            for ri in rev_candidates:
                row_text = " ".join(grid[ri])
                if "당기" in row_text or str(target_year) in row_text:
                    rev_ri = ri
                    break

        # 헤더 행 찾기 (매출 행 위에서 부문명이 있는 행)
        hdr_ri = None
        for ri in range(rev_ri - 1, -1, -1):
            text_cells = [c for c in grid[ri] if c.strip() and not re.match(r"^[\d,.()\-△]+$", c.replace(" ", ""))]
            if len(text_cells) >= 2:
                hdr_ri = ri
                break
        if hdr_ri is None:
            return None

        headers = grid[hdr_ri]
        rev_cells = grid[rev_ri]

        # 제외할 헤더 키워드
        bad_headers = {
            "금액", "비중", "비율", "구분", "계", "합계", "매출액", "매출",
            "전기", "당기", "내부거래", "조정", "연결조정", "총계", "소계",
            "단위", "영업이익", "매출원가", "총합계", "전사",
            # 재무제표 항목 (가로형 오인식 방지)
            "유동자산", "비유동자산", "자산총계", "부채총계", "자본총계",
            "매출총이익", "판매비", "관리비", "영업손실", "당기순이익",
            "현금흐름", "기타수익", "기타비용", "법인세", "이자수익",
            "지배기업", "비지배지분", "주당이익", "기본주당",
        }
        period_pat = re.compile(r"제\s?\d+\s?기|20\d{2}")

        segments: list[dict[str, Any]] = []
        for j in range(len(headers)):
            name = headers[j].strip()
            name_ns = name.replace(" ", "")
            if not name:
                continue
            if any(bw in name_ns for bw in bad_headers):
                continue
            if _is_skip_name(name):
                continue
            if period_pat.search(name):
                continue

            if j < len(rev_cells):
                rev = _extract_number(rev_cells[j])
                if rev is not None and rev > 0:
                    segments.append({"name": name, "revenue": rev})

        if len(segments) < 2:
            return None
        total = sum(s["revenue"] for s in segments)
        for s in segments:
            s["pct"] = s["revenue"] / total * 100 if total else 0
        return sorted(segments, key=lambda x: x["revenue"], reverse=True)

    # ── 세로형 세그먼트 테이블 파싱 ──

    def _try_vertical_segment(grid: list[list[str]], target_year: int) -> list[dict[str, Any]] | None:
        """세로형 테이블: 행에 부문명, 열에 매출액 (당기/전기).

        구분     | 당기    | 전기
        DX부문   | 174,000 | 160,000
        DS부문   | 111,000 | 80,000
        """
        if len(grid) < 3:
            return None

        # 헤더 행 찾기 (첫 5행 내에서 '구분' or '부문' + '당기'/'매출'/'금액' 패턴)
        hdr_ri = None
        for ri, row in enumerate(grid[:5]):
            joined = " ".join(row)
            has_label = any(
                cell.replace(" ", "") in ("구분", "부문", "사업부문", "부문명", "사업영역", "세그먼트")
                or _is_seg_header(cell)
                for cell in row
            )
            has_value = (
                "당기" in joined or "금액" in joined
                or "매출" in joined or str(target_year) in joined
                or re.search(r"제\s?\d+\s?기", joined)
            )
            if has_label and has_value:
                hdr_ri = ri
                break
        if hdr_ri is None:
            return None

        headers = grid[hdr_ri]

        # 라벨 열 (부문명), 값 열 (당기 매출) 식별
        label_col = None
        value_candidates: list[int] = []
        for ci, h in enumerate(headers):
            hn = h.replace(" ", "")
            if hn in ("구분", "부문", "사업부문", "부문명", "사업영역", "세그먼트") or _is_seg_header(h):
                label_col = ci
            elif "당기" in hn or _is_rev_header(h) or "금액" in hn:
                value_candidates.append(ci)
            elif str(target_year) in h:
                value_candidates.append(ci)
            elif re.search(r"제\s*(\d+)\s*기", h):
                value_candidates.append(ci)

        if label_col is None:
            label_col = 0
        val_col = _find_current_period_col([headers], value_candidates, target_year)
        if val_col is None:
            # 첫 번째 숫자 열 시도
            for ci in range(len(headers)):
                if ci == label_col:
                    continue
                # 데이터 행에 숫자가 있는 열
                for ri in range(hdr_ri + 1, min(hdr_ri + 3, len(grid))):
                    if ci < len(grid[ri]) and _extract_number(grid[ri][ci]) is not None:
                        val_col = ci
                        break
                if val_col is not None:
                    break
        if val_col is None:
            return None

        # 데이터 행 파싱
        segments: list[dict[str, Any]] = []
        for ri in range(hdr_ri + 1, len(grid)):
            row = grid[ri]
            if len(row) <= max(label_col, val_col):
                continue
            name = row[label_col].strip()
            if not name or _is_skip_name(name):
                continue
            row_text = " ".join(row)
            if re.search(r"합\s*계|소\s*계|총\s*계", row_text):
                continue
            rev = _extract_number(row[val_col])
            if rev is not None and rev > 0:
                segments.append({"name": name, "revenue": rev})

        if len(segments) < 2:
            return None
        total = sum(s["revenue"] for s in segments)
        for s in segments:
            s["pct"] = s["revenue"] / total * 100 if total else 0
        return sorted(segments, key=lambda x: x["revenue"], reverse=True)

    # ── 테이블 컨텍스트 점수 계산 ──

    def _context_score(table) -> int:
        """테이블 주변 텍스트에서 매출비중 관련 컨텍스트 점수 계산.

        높은 점수 = 사업부문별 매출 테이블일 가능성이 높음.
        """
        score = 0
        # 테이블 내부 텍스트
        ttext = table.get_text()
        if "사업부문" in ttext or "사업 부문" in ttext:
            score += 3
        if "매출유형" in ttext:
            score += 2
        if "부문" in ttext:
            score += 1

        # 주변 요소 탐색 (이전 3개 형제)
        context_parts = []
        prev = table.find_previous_sibling()
        for _ in range(3):
            if prev is None:
                break
            context_parts.append(prev.get_text())
            prev = prev.find_previous_sibling()
        context = " ".join(context_parts)

        context_keywords = [
            ("주요 제품", 5), ("주요제품", 5),
            ("매출실적", 4), ("매출 실적", 4),
            ("사업부문별", 4), ("사업 부문별", 4),
            ("부문별 매출", 4),
            ("매출 및 비율", 3), ("매출및비율", 3),
            ("매출현황", 3), ("매출 현황", 3),
            ("세그먼트", 3),
            ("제품 등의 현황", 3), ("제품등의현황", 3),
        ]
        for kw, pts in context_keywords:
            if kw in context or kw.replace(" ", "") in context.replace(" ", ""):
                score += pts

        return score

    # ── 단위 감지 (테이블 + 주변) ──

    def _resolve_multiplier(table, file_multiplier: int) -> int:
        ttext = table.get_text()
        m = _detect_unit_multiplier(ttext)
        if m != 1_000_000:
            return m
        # 직전 형제에서 탐색
        for prev in _iter_prev_siblings(table, 3):
            detected = _detect_unit_multiplier(prev.get_text())
            if detected != 1_000_000:
                return detected
        return file_multiplier

    def _iter_prev_siblings(elem, count: int):
        prev = elem.find_previous_sibling()
        for _ in range(count):
            if prev is None:
                break
            yield prev
            prev = prev.find_previous_sibling()

    # ── 메인 로직 ──

    # reprt_code별 보고서명 키워드 매핑
    _REPORT_KEYWORD_MAP = {
        "11011": "사업보고서",
        "11012": "반기보고서",
        "11013": "분기보고서",  # Q1
        "11014": "분기보고서",  # Q3
    }
    target_keyword = _REPORT_KEYWORD_MAP.get(reprt_code, "사업보고서")

    # Q1 vs Q3 구분 (둘 다 "분기보고서")
    # Q1: 보고서명에 "03" 또는 rcept_dt 4~6월
    # Q3: 보고서명에 "09" 또는 rcept_dt 10~12월
    def _is_target_quarter(report_nm: str, rcept_dt: str) -> bool:
        if reprt_code not in ("11013", "11014"):
            return True  # 사업보고서/반기보고서는 구분 필요 없음
        rn = report_nm.replace(" ", "")
        dt_month = int(rcept_dt[4:6]) if len(rcept_dt) >= 6 else 0
        if reprt_code == "11013":  # Q1
            # 보고서명에 "03" 포함 or 제출일 4~6월
            return (".03)" in rn or "03월" in rn or "1분기" in rn
                    or 4 <= dt_month <= 6)
        else:  # "11014" — Q3
            return (".09)" in rn or "09월" in rn or "3분기" in rn
                    or 10 <= dt_month <= 12)

    try:
        # 1. 보고서 검색
        resp = requests.get(
            "https://opendart.fss.or.kr/api/list.json",
            params={
                "crtfc_key": api_key,
                "corp_code": corp_code,
                "bgn_de": f"{year}0101",
                "end_de": f"{year + 1}1231",
                "pblntf_ty": "A",
                "page_count": 60,
            },
            timeout=30,
        )
        data = resp.json()
        if data.get("status") != "000":
            return None

        rcept_no = None
        reports = data.get("list", [])

        # 1차: 해당 키워드 보고서 (정정 제외)
        for item in reports:
            rn = item.get("report_nm", "")
            rcept_dt = item.get("rcept_dt", "")
            if target_keyword in rn and "정정" not in rn:
                if _is_target_quarter(rn, rcept_dt):
                    if f"{year}." in rn or f"{year}년" in rn or f"({year}" in rn:
                        rcept_no = item["rcept_no"]
                        break
        # 2차: 정정 보고서도 허용
        if not rcept_no:
            for item in reports:
                rn = item.get("report_nm", "")
                rcept_dt = item.get("rcept_dt", "")
                if target_keyword in rn:
                    if _is_target_quarter(rn, rcept_dt):
                        if f"{year}." in rn or f"{year}년" in rn or f"({year}" in rn:
                            rcept_no = item["rcept_no"]
                            break
        # 3차: 보고서명 형식이 비표준인 경우 (연도 검증 포함)
        if not rcept_no:
            for item in reports:
                rn = item.get("report_nm", "")
                rcept_dt = item.get("rcept_dt", "")
                if target_keyword in rn and _is_target_quarter(rn, rcept_dt):
                    found_years = re.findall(r'\d{4}', rn)
                    if not found_years or any(int(y) == year for y in found_years):
                        rcept_no = item["rcept_no"]
                        break
        if not rcept_no:
            return None

        # 2. 문서 ZIP 다운로드
        resp2 = requests.get(
            "https://opendart.fss.or.kr/api/document.xml",
            params={"crtfc_key": api_key, "rcept_no": rcept_no},
            timeout=60,
        )
        if resp2.status_code != 200:
            return None

        zf = zipfile.ZipFile(io.BytesIO(resp2.content))

        # 3. 모든 후보 테이블 수집 후 점수 기반 선택
        candidates: list[tuple[int, list[dict[str, Any]]]] = []
        found_single_segment = False  # 100% 단일부문 증거 추적

        # 비세그먼트 테이블 제외 키워드 (특수관계자, 종속기업 등)
        _NON_SEG_TABLE_KEYWORDS = [
            "특수관계자", "관계기업", "종속기업", "공동기업",
            "지배기업", "비지배지분", "연결대상",
            "채권", "채무", "대여금", "차입금",
            "공정가치측정", "상각후원가", "리스",
        ]

        for fname in zf.namelist():
            raw_bytes = zf.read(fname)
            # DART XML files may be UTF-8 or EUC-KR(CP949) encoded
            raw = None
            for enc in ("utf-8", "cp949", "euc-kr"):
                try:
                    raw = raw_bytes.decode(enc)
                    break
                except (UnicodeDecodeError, LookupError):
                    continue
            if raw is None:
                raw = raw_bytes.decode("utf-8", errors="ignore")
            soup = BeautifulSoup(raw, "html.parser")
            file_multiplier = _detect_unit_multiplier(raw)

            for table in soup.find_all("table"):
                ttext = table.get_text()

                # 단일부문이면 조기 종료
                if "단일부문" in ttext or "단일 부문" in ttext:
                    return None

                # 최소 조건: 매출 키워드 + 숫자
                if "매출" not in ttext:
                    continue
                if not re.search(r"\d{3,}", ttext.replace(",", "")):
                    continue

                # 비세그먼트 테이블 제외 (특수관계자, 종속기업 거래 등)
                ttext_ns = ttext.replace(" ", "")
                non_seg_hits = sum(
                    1 for kw in _NON_SEG_TABLE_KEYWORDS if kw in ttext_ns
                )
                if non_seg_hits >= 2:
                    continue

                grid = _table_to_grid(table)
                if not grid or len(grid) < 2:
                    continue

                # 재무제표 테이블 제외
                if _is_financial_statement_table(grid, ttext):
                    continue

                # 100% 단일부문 테이블 감지: "반도체 부문 | 32,765,719 | 100.0%"
                # 세그먼트 관련 키워드가 있는 테이블에서
                # 100% 패턴 + 데이터 행 1개 → 이 테이블은 단일부문 (스킵)
                has_seg_context = any(
                    kw in ttext_ns for kw in ["부문", "사업부문", "세그먼트"]
                )
                if has_seg_context and re.search(r"100\.?0?\s*%", ttext):
                    # 합계/소계 행 제외하고 데이터 행이 1개면 단일부문
                    data_rows = [
                        row for row in grid[1:]
                        if row and not re.search(r"합\s*계|소\s*계|총\s*계", " ".join(row))
                        and any(_extract_number(c) is not None for c in row)
                    ]
                    if len(data_rows) <= 1:
                        found_single_segment = True
                        continue  # 이 테이블만 스킵, 함수 전체 중단하지 않음

                table_multiplier = _resolve_multiplier(table, file_multiplier)

                ctx_score = _context_score(table)
                result = None

                # (A) 제품현황형 → 가장 우선
                any_seg_hdr = any(
                    any(_is_seg_header(cell) or _is_type_header(cell) for cell in row)
                    for row in grid[:7]
                )
                if any_seg_hdr:
                    result = _try_product_table(grid, year)
                    if result:
                        ctx_score += 10  # 보너스

                # (B) 가로형
                if result is None and ("부문" in ttext or "사업" in ttext):
                    result = _try_horizontal_segment(grid, year)
                    if result:
                        ctx_score += 5

                # (C) 세로형
                if result is None:
                    result = _try_vertical_segment(grid, year)
                    if result:
                        ctx_score += 3

                if result and len(result) >= 2:
                    # 결과에서 불량 세그먼트 필터링
                    result = [
                        s for s in result
                        if not _is_skip_name(s["name"])
                        and s["revenue"] > 0
                    ]
                    if len(result) >= 2:
                        for seg in result:
                            seg["revenue"] = seg["revenue"] * table_multiplier
                        candidates.append((ctx_score, result))

        if not candidates:
            return None

        # 100% 단일부문 증거가 있는 경우, 높은 품질 후보만 허용
        # (낮은 점수의 후보는 비세그먼트 테이블 오인식일 가능성 높음)
        if found_single_segment:
            high_quality = [(s, r) for s, r in candidates if s >= 8]
            if not high_quality:
                return None
            candidates = high_quality

        # 점수가 높은 것 우선, 동점이면 세그먼트 수가 많은 것
        candidates.sort(key=lambda x: (x[0], len(x[1])), reverse=True)
        best = candidates[0][1]

        # 최종 검증: 모든 세그먼트명이 동일하면 무효
        unique_names = {s["name"] for s in best}
        if len(unique_names) <= 1:
            # 같은 이름이 반복되면 차순위 후보 확인
            for _, candidate in candidates[1:]:
                cand_names = {s["name"] for s in candidate}
                if len(cand_names) >= 2:
                    best = candidate
                    break
            else:
                return None

        # 비중 재계산 (단위 변환 후)
        total = sum(s["revenue"] for s in best)
        for s in best:
            s["pct"] = s["revenue"] / total * 100 if total else 0
        return sorted(best, key=lambda x: x["revenue"], reverse=True)

    except Exception:
        return None


@st.cache_data(show_spinner="매출비중 변화 데이터 조회 중...")
def load_segment_history(
    corp_code: str, start_year: int = 2023,
) -> dict[int, list[dict[str, Any]]]:
    """연도별 사업부문 매출 데이터 조회 (start_year~현재).

    Returns: {year: [{"name": ..., "revenue": ..., "pct": ...}, ...]}
    """
    from datetime import date as _date
    current_year = _date.today().year
    result: dict[int, list[dict[str, Any]]] = {}
    for year in range(start_year, current_year + 1):
        segments = load_segment_data(corp_code, year)
        if segments:
            result[year] = segments
    return result


# ---------------------------------------------------------------------------
# Market Overview (Indices + Macro)
# ---------------------------------------------------------------------------

@st.cache_data(ttl=600, show_spinner="시장 지표 조회 중...")
def load_market_overview() -> dict[str, Any]:
    """Fetch major market indices and macro indicators via yfinance.

    Returns dict with 'indices' and 'macro' keys, each a list of dicts:
        {"name", "value", "change", "pct", "prev_close"}
    """
    import yfinance as yf

    TICKERS = {
        "indices": [
            ("KOSPI", "^KS11"),
            ("KOSDAQ", "^KQ11"),
            ("NASDAQ", "^IXIC"),
            ("S&P 500", "^GSPC"),
            ("다우존스", "^DJI"),
            ("니케이225", "^N225"),
        ],
        "macro": [
            ("USD/KRW", "KRW=X"),
            ("USD/JPY", "JPY=X"),
            ("금(oz)", "GC=F"),
            ("WTI유", "CL=F"),
            ("비트코인", "BTC-USD"),
            ("VIX", "^VIX"),
            ("미국10Y", "^TNX"),
        ],
    }

    def _fetch(name: str, symbol: str) -> dict[str, Any] | None:
        try:
            ticker = yf.Ticker(symbol)
            hist = ticker.history(period="5d")
            if hist.empty or len(hist) < 2:
                return None
            latest = hist.iloc[-1]
            prev = hist.iloc[-2]
            close_val = float(latest["Close"])
            prev_close = float(prev["Close"])
            change = close_val - prev_close
            pct = (change / prev_close * 100) if prev_close != 0 else 0.0

            # USD/JPY special: yfinance returns JPY per USD reciprocal
            if symbol == "JPY=X":
                close_val = close_val
                change = close_val - prev_close
                pct = (change / prev_close * 100) if prev_close != 0 else 0.0

            return {
                "name": name,
                "value": close_val,
                "change": change,
                "pct": pct,
                "prev_close": prev_close,
            }
        except Exception:
            return None

    result: dict[str, list[dict[str, Any]]] = {"indices": [], "macro": []}

    from concurrent.futures import ThreadPoolExecutor, as_completed

    # Flatten all fetch tasks
    tasks: list[tuple[str, str, str]] = []  # (category, name, symbol)
    for category, items in TICKERS.items():
        for name, symbol in items:
            tasks.append((category, name, symbol))

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = {
            executor.submit(_fetch, name, symbol): (cat, name)
            for cat, name, symbol in tasks
        }
        fetched: dict[str, list[tuple[int, dict]]] = {"indices": [], "macro": []}
        task_order = {name: i for i, (_, name, _) in enumerate(tasks)}
        for future in as_completed(futures):
            cat, name = futures[future]
            data = future.result()
            if data:
                fetched[cat].append((task_order[name], data))

    # Maintain original order
    for cat in fetched:
        fetched[cat].sort(key=lambda x: x[0])
        result[cat] = [d for _, d in fetched[cat]]

    return result


# ---------------------------------------------------------------------------
# Index Detail (KOSPI/KOSDAQ sparkline + OHLCV + 52-week range)
# ---------------------------------------------------------------------------

@st.cache_data(ttl=600, show_spinner=False)
def load_index_detail() -> dict[str, dict[str, Any]]:
    """Fetch extended data for KOSPI/KOSDAQ: OHLCV, 1-month sparkline, 52-week range."""
    import yfinance as yf
    from concurrent.futures import ThreadPoolExecutor

    SYMBOLS = {"KOSPI": "^KS11", "KOSDAQ": "^KQ11"}

    def _fetch_detail(name: str, symbol: str) -> tuple[str, dict[str, Any] | None]:
        try:
            ticker = yf.Ticker(symbol)

            # 1-month daily for sparkline
            hist_1mo = ticker.history(period="1mo")
            if hist_1mo.empty:
                return name, None

            latest = hist_1mo.iloc[-1]

            sparkline = [
                {"date": str(row.Index.date()), "close": round(float(row.Close), 2)}
                for row in hist_1mo.itertuples()
                if row.Close == row.Close  # skip NaN
            ]

            # Intraday 5-min data for sparkline with time labels
            intraday_sparkline: list[dict[str, Any]] = []
            try:
                hist_intraday = ticker.history(period="1d", interval="5m")
                if not hist_intraday.empty:
                    for row in hist_intraday.itertuples():
                        if row.Close == row.Close:  # skip NaN
                            intraday_sparkline.append({
                                "time": row.Index.strftime("%H:%M"),
                                "close": round(float(row.Close), 2),
                            })
            except Exception:
                pass

            # 1-year for 52-week range
            hist_1y = ticker.history(period="1y")
            if hist_1y.empty:
                week52_high = float(latest["High"])
                week52_low = float(latest["Low"])
            else:
                week52_high = float(hist_1y["High"].max())
                week52_low = float(hist_1y["Low"].min())

            return name, {
                "open": round(float(latest["Open"]), 2),
                "high": round(float(latest["High"]), 2),
                "low": round(float(latest["Low"]), 2),
                "volume": int(latest["Volume"]),
                "week52_high": round(week52_high, 2),
                "week52_low": round(week52_low, 2),
                "sparkline": sparkline,
                "sparkline_intraday": intraday_sparkline,
            }
        except Exception:
            return name, None

    result: dict[str, dict[str, Any]] = {}

    with ThreadPoolExecutor(max_workers=2) as executor:
        futures = [
            executor.submit(_fetch_detail, name, symbol)
            for name, symbol in SYMBOLS.items()
        ]
        for future in futures:
            name, data = future.result()
            if data:
                result[name] = data

    return result
