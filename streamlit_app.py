"""
SME Expansion Readiness Scorer — Streamlit Web App
Features:
  - Manual financial statement entry OR Excel upload (auto-detection)
  - Country Risk Score management with live World Bank API refresh
"""

import streamlit as st
import plotly.graph_objects as go
import numpy as np
import pandas as pd
import requests
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime as _parse_rfc_date
from typing import Optional, Dict, List
from datetime import datetime, timezone

# ─────────────────────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SME Expansion Readiness Scorer",
    page_icon="📊",
    layout="wide"
)

# ─────────────────────────────────────────────────────────────
#  CONSTANTS — WEIGHTS
# ─────────────────────────────────────────────────────────────
CRS_WEIGHTS: Dict[str, float] = {
    "macro_stability":       0.25,
    "institutional_quality": 0.25,
    "credit_risk":           0.25,
    "market_attractiveness": 0.25,
}

# ─────────────────────────────────────────────────────────────
#  BASELINE COUNTRY DATA
# ─────────────────────────────────────────────────────────────
COUNTRY_DATA_BASELINE: Dict[str, Dict[str, float]] = {
    "Australia":   {"macro_stability": 80, "institutional_quality": 90, "credit_risk": 88, "market_attractiveness": 82},
    "Bangladesh":  {"macro_stability": 55, "institutional_quality": 38, "credit_risk": 38, "market_attractiveness": 60},
    "Cambodia":    {"macro_stability": 55, "institutional_quality": 35, "credit_risk": 35, "market_attractiveness": 50},
    "China":       {"macro_stability": 70, "institutional_quality": 52, "credit_risk": 72, "market_attractiveness": 95},
    "Hong Kong":   {"macro_stability": 65, "institutional_quality": 78, "credit_risk": 80, "market_attractiveness": 78},
    "India":       {"macro_stability": 65, "institutional_quality": 55, "credit_risk": 55, "market_attractiveness": 90},
    "Indonesia":   {"macro_stability": 62, "institutional_quality": 50, "credit_risk": 52, "market_attractiveness": 80},
    "Malaysia":    {"macro_stability": 65, "institutional_quality": 65, "credit_risk": 65, "market_attractiveness": 68},
    "Myanmar":     {"macro_stability": 22, "institutional_quality": 18, "credit_risk": 20, "market_attractiveness": 42},
    "New Zealand": {"macro_stability": 78, "institutional_quality": 92, "credit_risk": 88, "market_attractiveness": 62},
    "Pakistan":    {"macro_stability": 28, "institutional_quality": 35, "credit_risk": 25, "market_attractiveness": 55},
    "Philippines": {"macro_stability": 58, "institutional_quality": 48, "credit_risk": 50, "market_attractiveness": 68},
    "Singapore":   {"macro_stability": 85, "institutional_quality": 95, "credit_risk": 92, "market_attractiveness": 78},
    "South Korea": {"macro_stability": 75, "institutional_quality": 80, "credit_risk": 82, "market_attractiveness": 80},
    "Sri Lanka":   {"macro_stability": 35, "institutional_quality": 42, "credit_risk": 28, "market_attractiveness": 48},
    "Taiwan":      {"macro_stability": 72, "institutional_quality": 78, "credit_risk": 80, "market_attractiveness": 75},
    "Thailand":    {"macro_stability": 60, "institutional_quality": 55, "credit_risk": 58, "market_attractiveness": 65},
    "Vietnam":     {"macro_stability": 65, "institutional_quality": 42, "credit_risk": 48, "market_attractiveness": 72},
}

COUNTRY_LIST = sorted(COUNTRY_DATA_BASELINE.keys())

# World Bank country codes (None = no WB membership, use baseline)
WORLD_BANK_CODES: Dict[str, Optional[str]] = {
    "Australia": "AUS", "Bangladesh": "BGD", "Cambodia": "KHM",
    "China": "CHN", "Hong Kong": "HKG", "India": "IND",
    "Indonesia": "IDN", "Malaysia": "MYS", "Myanmar": "MMR",
    "New Zealand": "NZL", "Pakistan": "PAK", "Philippines": "PHL",
    "Singapore": "SGP", "South Korea": "KOR", "Sri Lanka": "LKA",
    "Taiwan": None,   # not a WB member — baseline only
    "Thailand": "THA", "Vietnam": "VNM",
}

# ─────────────────────────────────────────────────────────────
#  EXCEL LABEL ALIASES
#  Keys must match session_state field names used in number_inputs
# ─────────────────────────────────────────────────────────────
IS_LABELS: Dict[str, List[str]] = {
    "rev":     ["total revenue", "net revenue", "net sales", "revenue", "turnover", "sales"],
    "cogs":    ["cost of goods sold", "cost of sales", "cost of revenue", "cogs", "direct costs"],
    "da":      ["depreciation and amortis", "depreciation & amortis", "depreciation and amortiz",
                "depreciation & amortiz", "d&a", "depreciation"],
    "ebit":    ["operating income", "operating profit", "income from operations", "ebit"],
    "int_exp": ["interest expense", "finance costs", "net interest expense", "interest charges"],
    "tax":     ["income tax expense", "provision for income taxes", "income tax", "tax expense", "taxes"],
    "ni":      ["net income", "net profit", "profit after tax", "net earnings", "pat"],
}

BS_LABELS: Dict[str, List[str]] = {
    "cash":            ["cash and cash equivalents", "cash & cash equivalents", "cash and equivalents", "cash"],
    "ar":              ["accounts receivable", "trade receivables", "trade debtors", "receivables", "debtors"],
    "inventory":       ["inventories", "inventory", "stock"],
    "other_ca":        ["other current assets", "prepaid expenses", "other current"],
    "total_assets":    ["total assets"],
    "ap":              ["accounts payable", "trade payables", "trade creditors", "payables", "creditors"],
    "short_term_debt": ["short-term debt", "short term debt", "current portion of long",
                        "current maturities", "current borrowings"],
    "long_term_debt":  ["long-term debt", "long term debt", "non-current borrowings",
                        "long term borrowings", "long-term borrowings"],
    "total_equity":    ["total stockholders", "total shareholders", "stockholders equity",
                        "shareholders equity", "shareholders' equity", "total equity"],
}

CF_LABELS: Dict[str, List[str]] = {
    "ocf":   ["net cash from operating", "net cash provided by operating", "cash flows from operating",
              "operating cash flow", "cash from operations"],
    "capex": ["purchases of property", "purchase of property", "additions to property",
              "capital expenditure", "capital expenditures", "capex"],
}

# ─────────────────────────────────────────────────────────────
#  SESSION STATE INITIALISATION
# ─────────────────────────────────────────────────────────────
if "country_data" not in st.session_state:
    st.session_state.country_data = {k: dict(v) for k, v in COUNTRY_DATA_BASELINE.items()}
if "crs_source" not in st.session_state:
    st.session_state.crs_source = "Baseline estimates (2024)"
if "crs_last_updated" not in st.session_state:
    st.session_state.crs_last_updated = None

# Initialise all financial input keys to 0.0 so widgets can read them safely
_IS_FIELDS = [f"{f}{y}" for f in IS_LABELS for y in (1, 2, 3)]
_BS_FIELDS = list(BS_LABELS.keys()) + ["total_assets"]
_CF_FIELDS = list(CF_LABELS.keys())
for _k in _IS_FIELDS + _BS_FIELDS + _CF_FIELDS:
    if _k not in st.session_state:
        st.session_state[_k] = 0.0

# ─────────────────────────────────────────────────────────────
#  SCORING FUNCTIONS
# ─────────────────────────────────────────────────────────────
def piecewise(val: float, pts: List) -> float:
    pts = sorted(pts, key=lambda p: p[0])
    if val <= pts[0][0]:  return float(pts[0][1])
    if val >= pts[-1][0]: return float(pts[-1][1])
    for i in range(len(pts) - 1):
        x0, s0 = pts[i]; x1, s1 = pts[i + 1]
        if x0 <= val <= x1:
            return s0 + (val - x0) / (x1 - x0) * (s1 - s0)
    return 50.0

def clamp(v: float) -> float:
    return max(0.0, min(100.0, v))

def score_revenue_quality(revenues: List[float]) -> float:
    rev = [r for r in revenues if r and r > 0]
    if len(rev) < 2: return 50.0
    n = len(rev) - 1
    cagr = (rev[-1] / rev[0]) ** (1 / n) - 1
    gr = [(rev[i] - rev[i - 1]) / rev[i - 1] for i in range(1, len(rev))]
    base = piecewise(cagr * 100, [(-20, 5), (0, 25), (5, 45), (10, 65), (20, 85), (30, 100)])
    penalty = min(20, np.std(gr) / (abs(np.mean(gr)) + 0.001) * 15) if len(gr) > 1 else 0
    return clamp(base - penalty)

def score_unit_economics(gm_pct: float) -> float:
    return clamp(piecewise(gm_pct, [(0, 5), (10, 20), (20, 38), (35, 58), (50, 72), (65, 88), (80, 100)]))

def score_operating_leverage(rev_t: float, rev_t1: float, ebit_t: float, ebit_t1: float) -> float:
    if not all([rev_t, rev_t1]) or rev_t1 == 0 or rev_t == rev_t1: return 50.0
    pct_rev = (rev_t - rev_t1) / abs(rev_t1)
    if abs(ebit_t1) < 0.001 * abs(rev_t1): return 50.0
    pct_ebit = (ebit_t - ebit_t1) / abs(ebit_t1)
    if pct_rev == 0: return 50.0
    ol = pct_ebit / pct_rev
    if pct_rev > 0:
        return clamp(piecewise(ol, [(-2, 5), (0, 25), (0.5, 50), (1, 65), (1.5, 80), (2.5, 95), (4, 100)]))
    else:
        return clamp(piecewise(ol, [(3, 5), (2, 20), (1, 40), (0.5, 60), (0, 75), (-1, 85)]))

def score_capital_efficiency(roic_pct: float, capex_rev_pct: float, ccc_days: float) -> float:
    rs = piecewise(roic_pct,      [(-20, 5), (0, 18), (5, 38), (10, 58), (15, 73), (25, 88), (40, 100)])
    cs = piecewise(capex_rev_pct, [(0, 95), (3, 82), (8, 65), (15, 45), (25, 25), (40, 10)])
    ds = piecewise(ccc_days,      [(-30, 100), (0, 88), (30, 72), (60, 52), (90, 35), (120, 20), (180, 8)])
    return clamp(0.5 * rs + 0.25 * cs + 0.25 * ds)

def score_balance_sheet(nd_ebitda: float, interest_coverage: Optional[float]) -> float:
    nd = piecewise(nd_ebitda, [(-2, 100), (0, 90), (1, 78), (2, 63), (3, 47), (4, 30), (5, 15), (7, 5)])
    ic = piecewise(interest_coverage,
                   [(0, 5), (1, 18), (2, 38), (3, 58), (5, 73), (8, 87), (12, 100)]) if interest_coverage else 90
    return clamp(0.5 * nd + 0.5 * ic)

# ─────────────────────────────────────────────────────────────
#  CHART FUNCTIONS
# ─────────────────────────────────────────────────────────────
def _score_color(v: float) -> str:
    return "#e74c3c" if v < 35 else "#f39c12" if v < 55 else "#27ae60" if v >= 75 else "#2980b9"

def gauge_chart(score: float, title: str) -> go.Figure:
    color = _score_color(score)
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=score,
        title={"text": title, "font": {"size": 16}},
        number={"suffix": " / 100", "font": {"size": 28}},
        gauge={
            "axis": {"range": [0, 100]},
            "bar": {"color": color},
            "steps": [
                {"range": [0,  35], "color": "#fadbd8"},
                {"range": [35, 55], "color": "#fdebd0"},
                {"range": [55, 75], "color": "#d6eaf8"},
                {"range": [75, 100], "color": "#d5f5e3"},
            ],
        }
    ))
    fig.update_layout(height=260, margin=dict(t=60, b=10, l=30, r=30))
    return fig

def bar_chart(labels: List[str], values: List[float], title: str) -> go.Figure:
    colors = [_score_color(v) for v in values]
    fig = go.Figure(go.Bar(
        x=values, y=labels, orientation="h",
        marker_color=colors,
        text=[f"{v:.1f}" for v in values],
        textposition="outside",
    ))
    fig.update_layout(
        title=title,
        xaxis=dict(range=[0, 115], title="Score (0–100)"),
        height=300,
        margin=dict(t=50, b=20, l=20, r=60),
    )
    return fig

# ─────────────────────────────────────────────────────────────
#  EXCEL PARSING HELPERS
# ─────────────────────────────────────────────────────────────
def _to_float(val) -> Optional[float]:
    """Convert an Excel cell to float, handling commas and parenthetical negatives."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    s = str(val).strip().replace(",", "").replace(" ", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except ValueError:
        return None

def _find_row(df: pd.DataFrame, aliases: List[str]) -> List[float]:
    """
    Search column 0 of df for any alias (case-insensitive partial match).
    Returns list of numeric values from that row (columns 1+), left to right.
    """
    label_col = df.iloc[:, 0].astype(str).str.strip().str.lower()
    for alias in aliases:
        mask = label_col.str.contains(alias.lower(), na=False, regex=False)
        if mask.any():
            row = df.loc[mask.idxmax()]
            return [v for v in (_to_float(c) for c in row.iloc[1:]) if v is not None]
    return []

def _best_sheet(xl: pd.ExcelFile, label_dict: Dict) -> Optional[pd.DataFrame]:
    """Return the sheet with the most matching labels."""
    best_df, best_n = None, 0
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet, header=None)
            n = sum(1 for aliases in label_dict.values() if _find_row(df, aliases))
            if n > best_n:
                best_n, best_df = n, df
        except Exception:
            continue
    return best_df if best_n > 0 else None

def parse_is(xl: pd.ExcelFile) -> Dict[str, float]:
    """Parse income statement from the best-matching sheet. Returns flat field dict."""
    df = _best_sheet(xl, IS_LABELS)
    if df is None:
        return {}
    result: Dict[str, float] = {}
    for field, aliases in IS_LABELS.items():
        vals = _find_row(df, aliases)
        # Assume columns ordered oldest→newest; take last 3
        vals = vals[-3:] if len(vals) >= 3 else vals
        for yr_idx, col_val in zip(range(3 - len(vals), 3), vals):
            result[f"{field}{yr_idx + 1}"] = col_val
    return result

def parse_bs(xl: pd.ExcelFile) -> Dict[str, float]:
    """Parse balance sheet (most recent column only)."""
    df = _best_sheet(xl, BS_LABELS)
    if df is None:
        return {}
    result: Dict[str, float] = {}
    for field, aliases in BS_LABELS.items():
        vals = _find_row(df, aliases)
        if vals:
            result[field] = vals[-1]   # rightmost = most recent
    return result

def parse_cf(xl: pd.ExcelFile) -> Dict[str, float]:
    """Parse cash flow statement (most recent column only)."""
    df = _best_sheet(xl, CF_LABELS)
    if df is None:
        return {}
    result: Dict[str, float] = {}
    for field, aliases in CF_LABELS.items():
        vals = _find_row(df, aliases)
        if vals:
            result[field] = abs(vals[-1])   # capex stored as positive
    return result

def apply_parsed(parsed: Dict[str, float]) -> None:
    """Write parsed values into session_state (picked up by number_inputs on rerun)."""
    for k, v in parsed.items():
        st.session_state[k] = float(v)

def _excel_preview(parsed: Dict[str, float], label_map: Dict, years: bool = True) -> pd.DataFrame:
    """Build a human-readable preview DataFrame from parsed values."""
    rows = []
    for field, aliases in label_map.items():
        label = aliases[0].title()
        if years:
            rows.append({
                "Field": label,
                "Year 1": parsed.get(f"{field}1", "—"),
                "Year 2": parsed.get(f"{field}2", "—"),
                "Year 3": parsed.get(f"{field}3", "—"),
            })
        else:
            rows.append({"Field": label, "Value": parsed.get(field, "—")})
    return pd.DataFrame(rows)

# ─────────────────────────────────────────────────────────────
#  WORLD BANK API
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=86_400, show_spinner=False)
def _wb(country_code: str, indicator: str) -> Optional[float]:
    """Fetch the most-recent World Bank indicator value (24h cache)."""
    url = (f"https://api.worldbank.org/v2/country/{country_code}"
           f"/indicator/{indicator}?format=json&mrv=5&per_page=10")
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        payload = r.json()
        if len(payload) < 2 or not payload[1]:
            return None
        vals = [e["value"] for e in payload[1] if e.get("value") is not None]
        return vals[0] if vals else None
    except Exception:
        return None

def _norm_gov(v: Optional[float]) -> Optional[float]:
    """Governance indicator (-2.5 → +2.5) → 0–100."""
    return clamp((v + 2.5) / 5.0 * 100) if v is not None else None

def wb_crs(country: str) -> Optional[Dict[str, float]]:
    """
    Compute CRS sub-scores from live World Bank data.
    Returns None if country has no WB code or all fetches fail.
    """
    code = WORLD_BANK_CODES.get(country)
    if not code:
        return None

    # Macro Stability
    gdp_g = _wb(code, "NY.GDP.MKTP.KD.ZG")
    inf   = _wb(code, "FP.CPI.TOTL.ZG")
    gdp_s = piecewise(gdp_g if gdp_g is not None else 3.0,
                      [(-5, 10), (0, 30), (3, 55), (6, 75), (9, 90), (12, 100)])
    inf_s = piecewise(abs(inf) if inf is not None else 5.0,
                      [(0, 100), (2, 92), (5, 72), (10, 45), (20, 20), (50, 5)])
    macro = clamp(0.6 * gdp_s + 0.4 * inf_s)

    # Institutional Quality — Rule of Law, Gov Effectiveness, Corruption Control
    rl = _norm_gov(_wb(code, "RL.EST"))
    ge = _norm_gov(_wb(code, "GE.EST"))
    cc = _norm_gov(_wb(code, "CC.EST"))
    gov_vals = [v for v in [rl, ge, cc] if v is not None]
    instit = float(np.mean(gov_vals)) if gov_vals else 50.0

    # Credit Risk proxy — Political Stability + Rule of Law + Regulatory Quality
    pv = _norm_gov(_wb(code, "PV.EST"))
    rq = _norm_gov(_wb(code, "RQ.EST"))
    cr_vals = [v for v in [pv, rl, rq] if v is not None]
    credit = float(np.mean(cr_vals)) if cr_vals else 50.0

    # Market Attractiveness — GDP size (log-scaled) + growth
    gdp_usd = _wb(code, "NY.GDP.MKTP.CD")
    if gdp_usd and gdp_usd > 0:
        gdp_log = np.log10(gdp_usd)
        size_s  = piecewise(gdp_log, [(9, 15), (10, 35), (11, 55), (11.5, 65),
                                       (12, 80), (12.7, 93), (13.5, 100)])
    else:
        size_s = 50.0
    grow_s = piecewise(gdp_g if gdp_g is not None else 3.0,
                       [(-5, 10), (0, 30), (3, 55), (6, 75), (9, 90)])
    market = clamp(0.7 * size_s + 0.3 * grow_s)

    return {
        "macro_stability":       round(macro,  1),
        "institutional_quality": round(instit, 1),
        "credit_risk":           round(credit, 1),
        "market_attractiveness": round(market, 1),
    }

# ─────────────────────────────────────────────────────────────
#  NEWS FEED
# ─────────────────────────────────────────────────────────────
NEWS_TOPICS: Dict[str, str] = {
    "Economy":       "economy GDP growth inflation interest rates",
    "Finance":       "banking finance capital markets monetary policy",
    "Trade":         "trade exports imports tariffs supply chain",
    "Investment":    "foreign investment FDI infrastructure development",
    "Business":      "business corporate earnings SME startup",
}

@st.cache_data(ttl=1800, show_spinner=False)
def fetch_country_news(country: str, topic_query: str) -> List[Dict]:
    """
    Fetch latest news for a country+topic via Google News RSS.
    Results are cached for 30 minutes.
    """
    raw_query = f"{country} {topic_query}"
    encoded   = raw_query.replace(" ", "+")
    url = f"https://news.google.com/rss/search?q={encoded}&hl=en-US&gl=US&ceid=US:en"
    articles: List[Dict] = []
    try:
        r = requests.get(url, timeout=10,
                         headers={"User-Agent": "Mozilla/5.0 (compatible; SMEScorer/1.0)"})
        r.raise_for_status()
        root    = ET.fromstring(r.content)
        channel = root.find("channel")
        if channel is None:
            return []
        for item in channel.findall("item")[:20]:
            title = (item.findtext("title") or "").strip()
            link  = (item.findtext("link")  or "").strip()
            pub   = (item.findtext("pubDate") or "")
            src_el = item.find("source")
            source = (src_el.text or "").strip() if src_el is not None else "Unknown source"

            # Resolve age
            try:
                dt        = _parse_rfc_date(pub)
                now       = datetime.now(timezone.utc)
                age_hours = (now - dt).total_seconds() / 3600
                date_str  = f"{dt.day} {dt.strftime('%b %Y, %H:%M')} UTC"
            except Exception:
                age_hours = 999
                date_str  = pub[:22] if pub else "—"

            if title and link:
                articles.append({
                    "title":     title,
                    "link":      link,
                    "source":    source,
                    "date":      date_str,
                    "age_hours": age_hours,
                })
    except Exception:
        pass
    return articles

# ─────────────────────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("📊 SME Readiness Scorer")
    st.markdown("---")
    company_name  = st.text_input("Company Name", value="My Company", key="company_name")
    target_country = st.selectbox("Target Country", COUNTRY_LIST, key="target_country")
    currency      = st.selectbox("Currency",
                                 ["SGD", "USD", "MYR", "AUD", "INR", "PHP", "IDR", "THB", "VND", "Other"],
                                 key="currency")
    values_unit   = st.selectbox("Values in", ["Thousands", "Millions", "Billions"], key="values_unit")

    st.markdown("---")
    st.markdown("#### Country Risk Score")
    _cd = st.session_state.country_data[target_country]
    _crs_val = sum(_cd[k] * CRS_WEIGHTS[k] for k in CRS_WEIGHTS)
    st.metric("Overall CRS", f"{_crs_val:.1f} / 100")
    for _lbl, _key in zip(
        ["Macro Stability", "Institutional Quality", "Credit Risk", "Market Attractiveness"],
        ["macro_stability", "institutional_quality", "credit_risk", "market_attractiveness"],
    ):
        _s = _cd[_key]
        _icon = "🟢" if _s >= 70 else "🟡" if _s >= 50 else "🔴"
        st.write(f"{_icon} **{_lbl}**: {_s:.0f}")
    st.caption(f"Source: {st.session_state.crs_source}")

# ─────────────────────────────────────────────────────────────
#  MAIN TABS
# ─────────────────────────────────────────────────────────────
tab_is, tab_bs, tab_cf, tab_res, tab_crs = st.tabs([
    "📈 Income Statement",
    "🏦 Balance Sheet",
    "💸 Cash Flow Statement",
    "🎯 Your Results",
    "📰 Country News",
])

# ══════════════════════════════════════════════════════════════
#  TAB 1 — INCOME STATEMENT
# ══════════════════════════════════════════════════════════════
with tab_is:
    st.header("Income Statement")

    with st.expander("📤 Upload via Excel — auto-fill from file", expanded=False):
        st.markdown(
            "Upload an Excel workbook containing your income statement. "
            "The app scans all sheets for row labels such as *Revenue*, *COGS*, *EBIT*, etc., "
            "and pre-fills the form below. Review and correct any values before calculating."
        )
        is_file = st.file_uploader(
            "Income Statement (.xlsx / .xls)",
            type=["xlsx", "xls"],
            key="is_file",
        )
        if is_file:
            with st.spinner("Reading Excel…"):
                try:
                    xl = pd.ExcelFile(is_file)
                    parsed_is = parse_is(xl)
                    if parsed_is:
                        st.success(f"Detected {len(parsed_is)} field(s). Review below, then click Apply.")
                        st.dataframe(_excel_preview(parsed_is, IS_LABELS, years=True),
                                     use_container_width=True, hide_index=True)
                        if st.button("✅ Apply to form", key="apply_is"):
                            apply_parsed(parsed_is)
                            st.rerun()
                    else:
                        st.warning(
                            "No income statement data detected. "
                            "Ensure row labels (e.g. 'Revenue', 'COGS', 'EBIT') appear in the first column."
                        )
                except Exception as exc:
                    st.error(f"Could not read file: {exc}")

    st.markdown(
        "Enter or review figures for up to **three financial years**. "
        "Year 2 and Year 3 (Latest) revenue are required for scoring."
    )
    col_y1, col_y2, col_y3 = st.columns(3)

    with col_y1:
        st.subheader("Year 1 (Oldest)")
        rev1     = st.number_input("Revenue",                     value=float(st.session_state.rev1),     min_value=0.0, key="rev1")
        cogs1    = st.number_input("COGS",                        value=float(st.session_state.cogs1),    min_value=0.0, key="cogs1")
        da1      = st.number_input("Depreciation & Amortisation", value=float(st.session_state.da1),      min_value=0.0, key="da1")
        ebit1    = st.number_input("EBIT",                        value=float(st.session_state.ebit1),                  key="ebit1")
        int_exp1 = st.number_input("Interest Expense",            value=float(st.session_state.int_exp1), min_value=0.0, key="int_exp1")
        tax1     = st.number_input("Income Tax",                  value=float(st.session_state.tax1),     min_value=0.0, key="tax1")
        ni1      = st.number_input("Net Income",                  value=float(st.session_state.ni1),                    key="ni1")

    with col_y2:
        st.subheader("Year 2")
        rev2     = st.number_input("Revenue",                     value=float(st.session_state.rev2),     min_value=0.0, key="rev2")
        cogs2    = st.number_input("COGS",                        value=float(st.session_state.cogs2),    min_value=0.0, key="cogs2")
        da2      = st.number_input("Depreciation & Amortisation", value=float(st.session_state.da2),      min_value=0.0, key="da2")
        ebit2    = st.number_input("EBIT",                        value=float(st.session_state.ebit2),                  key="ebit2")
        int_exp2 = st.number_input("Interest Expense",            value=float(st.session_state.int_exp2), min_value=0.0, key="int_exp2")
        tax2     = st.number_input("Income Tax",                  value=float(st.session_state.tax2),     min_value=0.0, key="tax2")
        ni2      = st.number_input("Net Income",                  value=float(st.session_state.ni2),                    key="ni2")

    with col_y3:
        st.subheader("Year 3 (Latest)")
        rev3     = st.number_input("Revenue",                     value=float(st.session_state.rev3),     min_value=0.0, key="rev3")
        cogs3    = st.number_input("COGS",                        value=float(st.session_state.cogs3),    min_value=0.0, key="cogs3")
        da3      = st.number_input("Depreciation & Amortisation", value=float(st.session_state.da3),      min_value=0.0, key="da3")
        ebit3    = st.number_input("EBIT",                        value=float(st.session_state.ebit3),                  key="ebit3")
        int_exp3 = st.number_input("Interest Expense",            value=float(st.session_state.int_exp3), min_value=0.0, key="int_exp3")
        tax3     = st.number_input("Income Tax",                  value=float(st.session_state.tax3),     min_value=0.0, key="tax3")
        ni3      = st.number_input("Net Income",                  value=float(st.session_state.ni3),                    key="ni3")

    st.markdown("---")
    st.markdown("#### Auto-calculated (Year 3)")
    ac1, ac2, ac3 = st.columns(3)
    gp3     = rev3 - cogs3
    gm3     = (gp3 / rev3 * 100) if rev3 > 0 else 0.0
    ebitda3 = ebit3 + da3
    ac1.metric("Gross Profit", f"{gp3:,.2f}")
    ac2.metric("Gross Margin %", f"{gm3:.1f}%")
    ac3.metric("EBITDA", f"{ebitda3:,.2f}")
    st.info("💡 All values should be in the same currency and unit selected in the sidebar.")

# ══════════════════════════════════════════════════════════════
#  TAB 2 — BALANCE SHEET
# ══════════════════════════════════════════════════════════════
with tab_bs:
    st.header("Balance Sheet")

    with st.expander("📤 Upload via Excel — auto-fill from file", expanded=False):
        st.markdown(
            "Upload an Excel workbook containing your balance sheet. "
            "The app scans for labels such as *Cash*, *Accounts Receivable*, "
            "*Total Equity*, etc. and pre-fills the form."
        )
        bs_file = st.file_uploader(
            "Balance Sheet (.xlsx / .xls)",
            type=["xlsx", "xls"],
            key="bs_file",
        )
        if bs_file:
            with st.spinner("Reading Excel…"):
                try:
                    xl = pd.ExcelFile(bs_file)
                    parsed_bs = parse_bs(xl)
                    if parsed_bs:
                        st.success(f"Detected {len(parsed_bs)} field(s). Review below, then click Apply.")
                        st.dataframe(_excel_preview(parsed_bs, BS_LABELS, years=False),
                                     use_container_width=True, hide_index=True)
                        if st.button("✅ Apply to form", key="apply_bs"):
                            apply_parsed(parsed_bs)
                            st.rerun()
                    else:
                        st.warning(
                            "No balance sheet data detected. "
                            "Ensure row labels appear in the first column."
                        )
                except Exception as exc:
                    st.error(f"Could not read file: {exc}")

    st.markdown("Enter figures for the **most recent year-end**.")
    bs_col1, bs_col2 = st.columns(2)

    with bs_col1:
        st.subheader("Assets")
        cash       = st.number_input("Cash & Equivalents",    value=float(st.session_state.cash),       min_value=0.0, key="cash")
        ar         = st.number_input("Accounts Receivable",   value=float(st.session_state.ar),         min_value=0.0, key="ar")
        inventory  = st.number_input("Inventory",             value=float(st.session_state.inventory),  min_value=0.0, key="inventory")
        other_ca   = st.number_input("Other Current Assets",  value=float(st.session_state.other_ca),   min_value=0.0, key="other_ca")
        total_assets = st.number_input("Total Assets",        value=float(st.session_state.total_assets), min_value=0.0, key="total_assets")

    with bs_col2:
        st.subheader("Liabilities & Equity")
        ap              = st.number_input("Accounts Payable",       value=float(st.session_state.ap),             min_value=0.0, key="ap")
        short_term_debt = st.number_input("Short-term Debt",        value=float(st.session_state.short_term_debt), min_value=0.0, key="short_term_debt")
        long_term_debt  = st.number_input("Long-term Debt",         value=float(st.session_state.long_term_debt),  min_value=0.0, key="long_term_debt")
        total_equity    = st.number_input("Total Shareholders' Equity", value=float(st.session_state.total_equity), key="total_equity")

    st.markdown("---")
    st.markdown("#### Auto-calculated")
    bs_ac1, bs_ac2, bs_ac3 = st.columns(3)
    total_debt_val = short_term_debt + long_term_debt
    net_debt_val   = total_debt_val - cash
    invested_cap   = total_debt_val + total_equity - cash
    bs_ac1.metric("Total Debt",      f"{total_debt_val:,.2f}")
    bs_ac2.metric("Net Debt",        f"{net_debt_val:,.2f}")
    bs_ac3.metric("Invested Capital", f"{invested_cap:,.2f}")
    st.info("💡 All values should be in the same currency and unit selected in the sidebar.")

# ══════════════════════════════════════════════════════════════
#  TAB 3 — CASH FLOW STATEMENT
# ══════════════════════════════════════════════════════════════
with tab_cf:
    st.header("Cash Flow Statement")

    with st.expander("📤 Upload via Excel — auto-fill from file", expanded=False):
        st.markdown(
            "Upload an Excel workbook containing your cash flow statement. "
            "The app looks for *Operating Cash Flow* and *Capital Expenditures* rows."
        )
        cf_file = st.file_uploader(
            "Cash Flow Statement (.xlsx / .xls)",
            type=["xlsx", "xls"],
            key="cf_file",
        )
        if cf_file:
            with st.spinner("Reading Excel…"):
                try:
                    xl = pd.ExcelFile(cf_file)
                    parsed_cf = parse_cf(xl)
                    if parsed_cf:
                        st.success(f"Detected {len(parsed_cf)} field(s). Review below, then click Apply.")
                        st.dataframe(_excel_preview(parsed_cf, CF_LABELS, years=False),
                                     use_container_width=True, hide_index=True)
                        if st.button("✅ Apply to form", key="apply_cf"):
                            apply_parsed(parsed_cf)
                            st.rerun()
                    else:
                        st.warning(
                            "No cash flow data detected. "
                            "Ensure row labels appear in the first column."
                        )
                except Exception as exc:
                    st.error(f"Could not read file: {exc}")

    st.markdown("Enter figures for the **most recent financial year**.")
    cf_col1, cf_col2 = st.columns(2)

    with cf_col1:
        ocf   = st.number_input("Operating Cash Flow",          value=float(st.session_state.ocf),   key="ocf",
                                 help="Net cash generated from operating activities.")
        capex = st.number_input("Capital Expenditures (Capex)", value=float(st.session_state.capex), min_value=0.0, key="capex",
                                 help="Enter as a positive number (cash outflow).")

    with cf_col2:
        st.markdown("#### Auto-calculated")
        fcf = ocf - capex
        delta_txt  = "Positive FCF ✓" if fcf >= 0 else "Negative FCF ✗"
        delta_col  = "normal" if fcf >= 0 else "inverse"
        st.metric("Free Cash Flow", f"{fcf:,.2f}", delta=delta_txt, delta_color=delta_col)

    st.info("💡 All values should be in the same currency and unit selected in the sidebar.")

# ══════════════════════════════════════════════════════════════
#  TAB 4 — RESULTS
# ══════════════════════════════════════════════════════════════
with tab_res:
    st.header("Your Results")

    if rev3 <= 0 or rev2 <= 0:
        st.warning(
            "⚠️ **Missing required inputs.** "
            "Please enter at least **Year 2** and **Year 3 (Latest) Revenue** "
            "in the Income Statement tab to compute your scores."
        )
    else:
        # ── Compute sub-scores ──────────────────────────────────
        revenues = [r for r in [rev1, rev2, rev3] if r > 0]
        rq_score = score_revenue_quality(revenues)

        gm_pct   = (rev3 - cogs3) / rev3 * 100 if rev3 > 0 else 0
        ue_score = score_unit_economics(gm_pct)

        ol_score = score_operating_leverage(rev3, rev2, ebit3, ebit2)

        ebt3_calc  = ebit3 - int_exp3
        tax_rate   = (tax3 / ebt3_calc) if ebt3_calc > 0 and tax3 > 0 else 0.20
        nopat      = ebit3 * (1 - tax_rate)
        roic_pct   = (nopat / invested_cap * 100) if invested_cap > 0 else 0
        capex_rev  = (capex / rev3 * 100) if rev3 > 0 else 0
        dso        = (ar        / rev3  * 365) if rev3  > 0 and ar        > 0 else 30
        dio        = (inventory / cogs3 * 365) if cogs3 > 0 and inventory > 0 else 30
        dpo        = (ap        / cogs3 * 365) if cogs3 > 0 and ap        > 0 else 30
        ccc        = dso + dio - dpo
        ce_score   = score_capital_efficiency(roic_pct, capex_rev, ccc)

        ebitda3_calc = ebit3 + da3
        nd_ebitda    = (net_debt_val / ebitda3_calc) if ebitda3_calc > 0 else (5 if net_debt_val > 0 else -1)
        ic_ratio     = (ebit3 / int_exp3) if int_exp3 > 0 else None
        bs_score     = score_balance_sheet(nd_ebitda, ic_ratio)

        fes = rq_score * 0.25 + ue_score * 0.20 + ol_score * 0.20 + ce_score * 0.20 + bs_score * 0.15

        cd_res = st.session_state.country_data[target_country]
        crs    = sum(cd_res[k] * CRS_WEIGHTS[k] for k in CRS_WEIGHTS)
        ers    = fes * (crs / 100)

        if ers >= 75:
            readiness_label = "★★★ Highly Ready";     readiness_color = "#27ae60"
        elif ers >= 55:
            readiness_label = "★★ Moderately Ready";  readiness_color = "#2980b9"
        elif ers >= 35:
            readiness_label = "★ Developing";         readiness_color = "#f39c12"
        else:
            readiness_label = "✗ Not Ready";           readiness_color = "#e74c3c"

        rev_pos     = [r for r in [rev1, rev2, rev3] if r > 0]
        n_cagr      = len(rev_pos) - 1
        rev_cagr    = ((rev_pos[-1] / rev_pos[0]) ** (1 / n_cagr) - 1) * 100 if n_cagr > 0 else 0.0
        ebit_margin = (ebit3 / rev3 * 100) if rev3 > 0 else 0

        # ── Heading ─────────────────────────────────────────
        st.markdown(f"### {company_name} → {target_country}")
        st.markdown("---")

        g_col, m_col, r_col = st.columns([2, 1, 1])

        with g_col:
            st.plotly_chart(gauge_chart(ers, "Expansion Readiness Score (ERS)"),
                            use_container_width=True)
        with m_col:
            st.metric("ERS",                f"{ers:.1f} / 100")
            st.metric("FES (Firm Score)",   f"{fes:.1f} / 100")
            st.metric("CRS (Country Score)", f"{crs:.1f} / 100")
        with r_col:
            st.markdown("#### Readiness Level")
            st.markdown(
                f"""<div style="background:{readiness_color}22;
                    border-left:6px solid {readiness_color};
                    padding:20px; border-radius:8px; margin-top:10px;">
                    <h2 style="color:{readiness_color}; margin:0;">{readiness_label}</h2>
                </div>""",
                unsafe_allow_html=True,
            )

        st.markdown("---")

        # ── Bar charts ───────────────────────────────────────
        bc1, bc2 = st.columns(2)
        with bc1:
            fes_labels = ["Revenue Quality", "Unit Economics", "Operating Leverage",
                          "Capital Efficiency", "Balance Sheet Strength"]
            fes_values = [rq_score, ue_score, ol_score, ce_score, bs_score]
            st.plotly_chart(bar_chart(fes_labels, fes_values, "FES Sub-score Breakdown"),
                            use_container_width=True)
        with bc2:
            crs_bar_labels = ["Macro Stability", "Institutional Quality",
                              "Credit Risk", "Market Attractiveness"]
            crs_bar_values = [cd_res["macro_stability"], cd_res["institutional_quality"],
                              cd_res["credit_risk"],     cd_res["market_attractiveness"]]
            st.plotly_chart(bar_chart(crs_bar_labels, crs_bar_values,
                                      f"CRS Sub-score Breakdown — {target_country}"),
                            use_container_width=True)

        st.markdown("---")

        # ── Key financial metrics table ───────────────────────
        st.markdown("#### Key Financial Metrics")
        ic_display = f"{ic_ratio:.2f}x" if ic_ratio is not None else "N/A (no interest expense)"
        df_metrics = pd.DataFrame({
            "Metric": ["Revenue CAGR (%)", "Gross Margin (%)", "EBIT Margin (%)", "ROIC (%)",
                       "Capex / Revenue (%)", "Cash Conversion Cycle (days)",
                       "Net Debt / EBITDA (x)", "Interest Coverage (x)"],
            "Value":  [f"{rev_cagr:.1f}%", f"{gm_pct:.1f}%", f"{ebit_margin:.1f}%", f"{roic_pct:.1f}%",
                       f"{capex_rev:.1f}%", f"{ccc:.1f} days",
                       f"{nd_ebitda:.2f}x", ic_display],
        })
        st.dataframe(df_metrics, use_container_width=True, hide_index=True)

        st.markdown("---")

        # ── Improvement cards ────────────────────────────────
        improvement_items = [
            ("Revenue Quality",        rq_score,
             "Grow revenue consistently year-over-year. Aim for a higher share of recurring revenue "
             "and reduce volatility to improve your CAGR profile."),
            ("Unit Economics",         ue_score,
             f"Your gross margin of {gm_pct:.1f}% leaves room to improve. "
             "Reduce COGS, renegotiate supplier contracts, or shift to higher-margin products/services. "
             "Target above 35%."),
            ("Operating Leverage",     ol_score,
             "Work on growing EBIT faster than revenue by controlling fixed costs. "
             "Positive operating leverage means each incremental revenue dollar drops more to profit."),
            ("Capital Efficiency",     ce_score,
             f"ROIC of {roic_pct:.1f}% and a CCC of {ccc:.0f} days suggest efficiency gains are available. "
             "Collect receivables faster, optimise inventory, and review capex allocation."),
            ("Balance Sheet Strength", bs_score,
             f"Net Debt/EBITDA of {nd_ebitda:.1f}x indicates leverage risk. "
             "Aim for below 2x before expanding. Ensure interest coverage stays comfortably above 3x."),
        ]

        low_scores = [(n, s, a) for n, s, a in improvement_items if s < 60]
        if low_scores:
            st.markdown("#### What Needs Improvement")
            for name, score, advice in low_scores:
                col = _score_color(score)
                st.markdown(
                    f"""<div style="border-left:6px solid {col}; background:{col}11;
                        padding:14px 18px; border-radius:6px; margin-bottom:12px;">
                        <strong style="color:{col}; font-size:16px;">{name}</strong>
                        <span style="float:right; font-weight:bold; color:{col};">{score:.1f} / 100</span>
                        <br/><span style="color:#555; font-size:14px;">{advice}</span>
                    </div>""",
                    unsafe_allow_html=True,
                )
        else:
            st.success("✅ All sub-scores are 60 or above — strong financial health across the board!")

        st.markdown("---")

        # ── Download report ──────────────────────────────────
        report = f"""SME EXPANSION READINESS REPORT
================================
Company   : {company_name}
Country   : {target_country}
Currency  : {currency} ({values_unit})
Generated : {datetime.now().strftime("%Y-%m-%d %H:%M")}

SCORES
------
Expansion Readiness Score (ERS) : {ers:.2f} / 100
Firm Expansion Score (FES)       : {fes:.2f} / 100
Country Risk Score (CRS)         : {crs:.2f} / 100
Readiness Level                  : {readiness_label}

FES SUB-SCORES
--------------
Revenue Quality         : {rq_score:.2f}  (weight 25%)
Unit Economics          : {ue_score:.2f}  (weight 20%)
Operating Leverage      : {ol_score:.2f}  (weight 20%)
Capital Efficiency      : {ce_score:.2f}  (weight 20%)
Balance Sheet Strength  : {bs_score:.2f}  (weight 15%)

CRS SUB-SCORES — {target_country}
{'-' * (20 + len(target_country))}
Macro Stability         : {cd_res["macro_stability"]}
Institutional Quality   : {cd_res["institutional_quality"]}
Credit Risk             : {cd_res["credit_risk"]}
Market Attractiveness   : {cd_res["market_attractiveness"]}
Data source             : {st.session_state.crs_source}

KEY FINANCIAL METRICS
---------------------
Revenue CAGR            : {rev_cagr:.1f}%
Gross Margin            : {gm_pct:.1f}%
EBIT Margin             : {ebit_margin:.1f}%
ROIC                    : {roic_pct:.1f}%
Capex / Revenue         : {capex_rev:.1f}%
Cash Conversion Cycle   : {ccc:.1f} days
Net Debt / EBITDA       : {nd_ebitda:.2f}x
Interest Coverage       : {ic_display}

FORMULA
-------
ERS = FES x (CRS / 100)
    = {fes:.2f} x ({crs:.2f} / 100)
    = {ers:.2f}
"""
        st.download_button(
            "📥 Download Report",
            data=report.strip(),
            file_name=f"{company_name.replace(' ', '_')}_ERS_Report.txt",
            mime="text/plain",
        )

# ══════════════════════════════════════════════════════════════
#  TAB 5 — COUNTRY RISK SCORES
# ══════════════════════════════════════════════════════════════
with tab_crs:
    st.header("📰 Country Economic & Financial News")
    st.markdown(
        "Browse the latest financial and economic news for any APAC country. "
        "Articles are fetched live from Google News and refreshed every 30 minutes."
    )

    # ── Controls row ─────────────────────────────────────────
    ctrl1, ctrl2, ctrl3, ctrl4 = st.columns([2, 2, 1, 1])

    with ctrl1:
        news_country = st.selectbox(
            "Country",
            COUNTRY_LIST,
            index=COUNTRY_LIST.index(target_country),
            key="news_country",
        )
    with ctrl2:
        news_topic = st.selectbox(
            "Topic",
            list(NEWS_TOPICS.keys()),
            key="news_topic",
        )
    with ctrl3:
        custom_query = st.text_input("Custom search term", placeholder="e.g. inflation", key="news_custom")
    with ctrl4:
        st.markdown("&nbsp;", unsafe_allow_html=True)  # vertical alignment spacer
        do_refresh = st.button("🔄 Refresh", key="news_refresh")

    if do_refresh:
        fetch_country_news.clear()

    # Build final query
    query_extra = custom_query.strip() if custom_query.strip() else NEWS_TOPICS[news_topic]

    # ── Fetch ────────────────────────────────────────────────
    with st.spinner(f"Fetching {news_topic} news for {news_country}…"):
        articles = fetch_country_news(news_country, query_extra)

    if not articles:
        st.warning(
            "⚠️ No articles returned. This may be a network issue or the query returned no results. "
            "Try a different topic or check your internet connection."
        )
    else:
        st.caption(
            f"Showing **{len(articles)} articles** · {news_country} · {news_topic} · "
            f"Cached for 30 min — click Refresh for latest"
        )
        st.markdown("---")

        # ── Article grid (2 columns) ─────────────────────────
        left_col, right_col = st.columns(2)
        col_cycle = [left_col, right_col]

        for i, art in enumerate(articles):
            age = art["age_hours"]
            if age < 6:
                border = "#27ae60"; badge = "🟢 Breaking"
            elif age < 24:
                border = "#2ecc71"; badge = "🟢 Today"
            elif age < 72:
                border = "#2980b9"; badge = "🔵 This week"
            else:
                border = "#95a5a6"; badge = "⚪ Older"

            with col_cycle[i % 2]:
                st.markdown(
                    f"""
                    <div style="
                        border-left: 4px solid {border};
                        background: #f8f9fa;
                        padding: 14px 18px;
                        border-radius: 6px;
                        margin-bottom: 14px;
                    ">
                        <a href="{art['link']}" target="_blank" style="
                            text-decoration: none;
                            color: #1a252f;
                            font-weight: 600;
                            font-size: 14px;
                            line-height: 1.5;
                        ">{art['title']}</a>
                        <br/><br/>
                        <span style="color: #7f8c8d; font-size: 12px;">
                            📰 {art['source']} &nbsp;·&nbsp; 🕐 {art['date']} &nbsp;·&nbsp; {badge}
                        </span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        st.markdown("---")
        st.caption("News sourced from Google News RSS. Articles link directly to their original publishers.")
