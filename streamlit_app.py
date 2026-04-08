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
#  STRATEGY DASHBOARD — STATIC DATA
# ─────────────────────────────────────────────────────────────

COUNTRY_CURRENCIES: Dict[str, str] = {
    "Australia": "AUD", "Bangladesh": "BDT", "Cambodia": "KHR",
    "China": "CNY", "Hong Kong": "HKD", "India": "INR",
    "Indonesia": "IDR", "Malaysia": "MYR", "Myanmar": "MMK",
    "New Zealand": "NZD", "Pakistan": "PKR", "Philippines": "PHP",
    "Singapore": "SGD", "South Korea": "KRW", "Sri Lanka": "LKR",
    "Taiwan": "TWD", "Thailand": "THB", "Vietnam": "VND",
}

FX_VS_SGD: Dict[str, float] = {
    "Australia": 1.12, "Bangladesh": 86.0, "Cambodia": 3050.0,
    "China": 5.35, "Hong Kong": 5.85, "India": 62.0,
    "Indonesia": 11500.0, "Malaysia": 3.45, "Myanmar": 1580.0,
    "New Zealand": 1.22, "Pakistan": 210.0, "Philippines": 42.0,
    "Singapore": 1.0, "South Korea": 565.0, "Sri Lanka": 235.0,
    "Taiwan": 23.5, "Thailand": 26.5, "Vietnam": 18900.0,
}

REGULATORY_DATA: Dict[str, Dict] = {
    "Australia": {
        "corp_tax": "30% (25% for SMEs under AUD 50M)",
        "gst_vat": "10% GST",
        "import_duty": "0–5% (SAFTA free trade with SG)",
        "withholding_tax": "30% dividends (reduced under DTA)",
        "key_licenses": ["Australian Business Number (ABN)", "ASIC company registration", "GST registration (revenue > AUD 75K)", "Industry-specific: APRA, TGA"],
        "fdi_restrictions": "FIRB review for investments > AUD 310M",
        "risk_level": "Low", "risk_score": 85,
        "key_risks": ["FIRB approval process delays", "State-level licensing variations", "ACCC consumer law compliance"],
        "dta_with_sg": True,
    },
    "Bangladesh": {
        "corp_tax": "27.5% (listed) / 32.5% (unlisted)",
        "gst_vat": "15% VAT",
        "import_duty": "5–25% + supplementary duty",
        "withholding_tax": "20% dividends",
        "key_licenses": ["RJSC company registration", "Trade License from City Corporation", "Tax ID Number (TIN)", "Export/Import Registration Certificate"],
        "fdi_restrictions": "100% FDI allowed in most sectors; restricted in telecom, banking",
        "risk_level": "High", "risk_score": 38,
        "key_risks": ["BDT not freely convertible", "Bureaucratic delays and corruption risk", "Political instability"],
        "dta_with_sg": True,
    },
    "Cambodia": {
        "corp_tax": "20%",
        "gst_vat": "10% VAT",
        "import_duty": "0–35%",
        "withholding_tax": "14% dividends",
        "key_licenses": ["MOC company registration", "Tax registration (GDT)", "QIP status for CDC incentives", "Sector-specific permits"],
        "fdi_restrictions": "100% FDI allowed (except land ownership)",
        "risk_level": "High", "risk_score": 35,
        "key_risks": ["USD-dollarised — limited KHR hedging options", "Weak IP and contract enforcement", "Political concentration risk"],
        "dta_with_sg": False,
    },
    "China": {
        "corp_tax": "25% (15% for high-tech enterprises)",
        "gst_vat": "13% / 9% / 6% VAT (multi-rate)",
        "import_duty": "Avg ~7.5% (varies by HS code)",
        "withholding_tax": "10% dividends (reduced under DTA)",
        "key_licenses": ["SAMR business registration (WFOE/JV)", "Business license + org code", "CFDA/MIIT for regulated sectors", "GACC import/export license"],
        "fdi_restrictions": "Negative list sectors restricted; WFOE structure common",
        "risk_level": "Medium", "risk_score": 60,
        "key_risks": ["Capital controls on repatriation", "Regulatory complexity and local partner risk", "Geopolitical and compliance risk"],
        "dta_with_sg": True,
    },
    "Hong Kong": {
        "corp_tax": "16.5% (8.25% on first HKD 2M profits)",
        "gst_vat": "None",
        "import_duty": "0% (free port)",
        "withholding_tax": "0% dividends",
        "key_licenses": ["Companies Registry CR registration", "Business Registration Certificate", "SFC license (financial services)", "Import/export declaration only"],
        "fdi_restrictions": "Minimal — highly open economy",
        "risk_level": "Medium", "risk_score": 72,
        "key_risks": ["Political risk post-2020 NSL", "HKD peg long-term sustainability", "Regulatory convergence with mainland China"],
        "dta_with_sg": True,
    },
    "India": {
        "corp_tax": "22% base (25.17% effective with surcharge)",
        "gst_vat": "5–28% GST (multi-slab)",
        "import_duty": "10–20% BCD + IGST",
        "withholding_tax": "10% dividends (reduced under DTA)",
        "key_licenses": ["MCA21 company incorporation", "PAN and TAN registration", "GST registration (revenue > INR 20L)", "FEMA/RBI approval for FDI", "Sector-specific: SEBI, IRDAI"],
        "fdi_restrictions": "FDI under automatic or government approval route by sector",
        "risk_level": "Medium", "risk_score": 58,
        "key_risks": ["Multi-layer GST and tax complexity", "Bureaucratic approval delays", "FEMA repatriation regulations"],
        "dta_with_sg": True,
    },
    "Indonesia": {
        "corp_tax": "22%",
        "gst_vat": "11% PPN",
        "import_duty": "0–40%",
        "withholding_tax": "10% dividends",
        "key_licenses": ["NIB via OSS portal", "KITAS for foreign directors", "BPJS social security", "OJK/BPOM for regulated sectors"],
        "fdi_restrictions": "PT PMA: minimum IDR 10B capital; positive investment list applies",
        "risk_level": "Medium", "risk_score": 55,
        "key_risks": ["Local content requirements", "Complex import/export procedures", "Omnibus Law implementation gaps"],
        "dta_with_sg": True,
    },
    "Malaysia": {
        "corp_tax": "24% (17% for SMEs on first MYR 600K)",
        "gst_vat": "10% SST services / 5–10% sales tax",
        "import_duty": "0–25%",
        "withholding_tax": "No WHT on dividends (single-tier system)",
        "key_licenses": ["SSM company registration", "MOF registration (government contracts)", "BNM license (finance)", "Halal certification if applicable"],
        "fdi_restrictions": "Generally open; equity restrictions in certain sectors",
        "risk_level": "Low-Medium", "risk_score": 65,
        "key_risks": ["Bumiputera equity requirements in some sectors", "Political transition uncertainty", "SST compliance complexity"],
        "dta_with_sg": True,
    },
    "Myanmar": {
        "corp_tax": "22%",
        "gst_vat": "5% Commercial Tax",
        "import_duty": "0–40%",
        "withholding_tax": "10% dividends",
        "key_licenses": ["DICA registration", "MIC permit for FDI", "MEB banking approval for transfers"],
        "fdi_restrictions": "Significant restrictions; international sanctions apply",
        "risk_level": "Very High", "risk_score": 18,
        "key_risks": ["International sanctions (US, EU)", "Military governance and political instability", "Currency inconvertibility and capital controls"],
        "dta_with_sg": True,
    },
    "New Zealand": {
        "corp_tax": "28%",
        "gst_vat": "15% GST",
        "import_duty": "0–5%",
        "withholding_tax": "15% dividends (reduced under DTA)",
        "key_licenses": ["Companies Office (NZBN)", "IRD number + GST registration", "FMA license (finance)", "MPI (food/agri)"],
        "fdi_restrictions": "OIO consent required for sensitive NZ assets/land",
        "risk_level": "Low", "risk_score": 88,
        "key_risks": ["Small market size limits scale", "Overseas Investment Act compliance", "High logistics costs from distance"],
        "dta_with_sg": True,
    },
    "Pakistan": {
        "corp_tax": "29%",
        "gst_vat": "18% GST",
        "import_duty": "3–35%",
        "withholding_tax": "15% dividends",
        "key_licenses": ["SECP company registration", "National Tax Number (NTN)", "SRB/PRA sales tax registration", "BOI investment approval"],
        "fdi_restrictions": "FDI allowed; restrictions in defence, media, real estate",
        "risk_level": "Very High", "risk_score": 28,
        "key_risks": ["PKR depreciation and capital controls", "IMF program dependency", "Security and political risk"],
        "dta_with_sg": True,
    },
    "Philippines": {
        "corp_tax": "25% (20% for SMEs with net income ≤ PHP 5M)",
        "gst_vat": "12% VAT",
        "import_duty": "0–30%",
        "withholding_tax": "25% dividends (15% reduced rate possible)",
        "key_licenses": ["SEC company registration", "BIR tax registration", "LGU business permit", "PEZA/BOI for incentives"],
        "fdi_restrictions": "60/40 foreign equity limits in key sectors",
        "risk_level": "Medium", "risk_score": 52,
        "key_risks": ["60/40 foreign equity restriction in key sectors", "Multiple layers: SEC, BIR, LGU", "Natural disaster risk"],
        "dta_with_sg": True,
    },
    "Singapore": {
        "corp_tax": "17% (startup tax exemption available)",
        "gst_vat": "9% GST",
        "import_duty": "0% (except alcohol, tobacco, fuel)",
        "withholding_tax": "N/A (home market)",
        "key_licenses": ["ACRA registration (same-day)", "GST registration if revenue > SGD 1M", "MAS license (finance)", "NEA license (food/environment)"],
        "fdi_restrictions": "None — most open economy in APAC",
        "risk_level": "Very Low", "risk_score": 95,
        "key_risks": ["High operating costs (rent, salaries)", "Small domestic market size"],
        "dta_with_sg": True,
    },
    "South Korea": {
        "corp_tax": "9–24% progressive",
        "gst_vat": "10% VAT",
        "import_duty": "0–8% (SGKFTA reduced rates)",
        "withholding_tax": "20% dividends (reduced under DTA)",
        "key_licenses": ["MOCI business registration", "KSIC classification", "FSC license (finance)", "KCGS environmental compliance"],
        "fdi_restrictions": "Generally open; OPIC notification required for some industries",
        "risk_level": "Low", "risk_score": 80,
        "key_risks": ["Language and cultural barriers", "Chaebols dominate key sectors", "North Korea geopolitical risk"],
        "dta_with_sg": True,
    },
    "Sri Lanka": {
        "corp_tax": "30%",
        "gst_vat": "18% VAT",
        "import_duty": "0–30% + PAL + CESS levies",
        "withholding_tax": "15% dividends",
        "key_licenses": ["Registrar of Companies", "BOI approval for FDI", "CBSL for financial sector", "CUSTOMS import/export license"],
        "fdi_restrictions": "BOI approval required; some sector restrictions",
        "risk_level": "High", "risk_score": 32,
        "key_risks": ["Post-debt crisis economic fragility", "IMF program conditionality", "LKR volatility and capital account restrictions"],
        "dta_with_sg": True,
    },
    "Taiwan": {
        "corp_tax": "20%",
        "gst_vat": "5% VAT",
        "import_duty": "0–25%",
        "withholding_tax": "21% dividends",
        "key_licenses": ["MOEA company registration", "NTBSA tax registration", "FSC license (finance)", "BSMI product certification"],
        "fdi_restrictions": "Negative list approach; open to most FDI",
        "risk_level": "Low-Medium", "risk_score": 75,
        "key_risks": ["Cross-strait geopolitical risk", "No DTA with Singapore", "Strategic sector restrictions"],
        "dta_with_sg": False,
    },
    "Thailand": {
        "corp_tax": "20%",
        "gst_vat": "7% VAT",
        "import_duty": "0–80% (TAFTA reduced rates for SG)",
        "withholding_tax": "10% dividends",
        "key_licenses": ["DBD company registration", "Revenue Department tax reg", "BOI certificate for incentives", "FDA/TFDA license (food/pharma)"],
        "fdi_restrictions": "Foreign Business Act: restricted sectors across 3 lists",
        "risk_level": "Medium", "risk_score": 58,
        "key_risks": ["Foreign Business Act sector restrictions", "Historical political risk (coups)", "FBA legal liability for foreign directors"],
        "dta_with_sg": True,
    },
    "Vietnam": {
        "corp_tax": "20% (10–17% for priority sectors)",
        "gst_vat": "10% VAT (8% reduced rate)",
        "import_duty": "0–35% (VSFTA reduced rates for SG)",
        "withholding_tax": "0% dividends (repatriation conditions apply)",
        "key_licenses": ["MPI/DPI registration", "IRC (Investment Registration Certificate)", "ERC (Enterprise Registration Certificate)", "SBV/MOH for regulated sectors"],
        "fdi_restrictions": "Vietnam–Singapore FTA; conditional sector list applies",
        "risk_level": "Medium", "risk_score": 50,
        "key_risks": ["IP protection enforcement gaps", "Rapid and unpredictable regulatory changes", "VND capital repatriation procedures"],
        "dta_with_sg": True,
    },
}

DBS_PRODUCTS: Dict[str, Dict] = {
    "DBS BusinessTerm Loan": {
        "rate_pa": 4.25,
        "max_loan": 500_000,
        "tenure_min": 12,
        "tenure_max": 60,
        "description": "Fixed-rate term loan for business expansion and working capital needs. Competitive rates with flexible repayment.",
        "features": ["Fixed interest rate — no surprises", "Tenure: 1–5 years", "No collateral for amounts ≤ SGD 500K", "Fast approval within 3–5 business days"],
        "structure": "Amortising",
        "color": "#1a5276",
        "icon": "🏦",
    },
    "DBS SME Working Capital Loan (EFS)": {
        "rate_pa": 4.75,
        "max_loan": 500_000,
        "tenure_min": 12,
        "tenure_max": 60,
        "description": "Enterprise Financing Scheme co-funded by Enterprise Singapore. Government bears 70% of default risk — lower barrier for SMEs.",
        "features": ["Government-backed (EnterpriseSG)", "Lower collateral requirement", "Open to companies ≥ 30% local shareholding", "Loan up to SGD 500,000"],
        "structure": "Amortising",
        "color": "#1e8449",
        "icon": "🏛️",
    },
    "DBS Trade Finance Line": {
        "rate_pa": 5.50,
        "max_loan": 1_000_000,
        "tenure_min": 3,
        "tenure_max": 24,
        "description": "Revolving short-term trade finance facility for import/export transactions, inventory financing, and overseas supplier payments.",
        "features": ["Multi-currency drawdown (SGD, USD, EUR, CNY)", "Revolving — redraw as you repay", "Supports LC, TR, invoice financing", "Up to SGD 1,000,000"],
        "structure": "Interest-only (revolving)",
        "color": "#784212",
        "icon": "🚢",
    },
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
if "mc_results" not in st.session_state:
    st.session_state.mc_results = None

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
tab_is, tab_bs, tab_cf, tab_res, tab_crs, tab_strat, tab_mc = st.tabs([
    "📈 Income Statement",
    "🏦 Balance Sheet",
    "💸 Cash Flow Statement",
    "🎯 Your Results",
    "📰 Country News",
    "🌏 Strategy Dashboard",
    "🎲 Monte Carlo NPV",
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

# ══════════════════════════════════════════════════════════════
#  TAB 6 — STRATEGY DASHBOARD
# ══════════════════════════════════════════════════════════════
with tab_strat:
    st.header(f"Strategy Dashboard — Expanding to {target_country}")
    st.caption("Actionable expansion strategy, FX risk management, financing, and regulatory guidance tailored to your selected market.")

    sd1, sd2, sd3, sd4 = st.tabs([
        "🗺️  Expansion Strategy",
        "🔄  Hedging Solutions",
        "💰  Capital Solutions",
        "📋  Regulatory Risk",
    ])

    # ──────────────────────────────────────────────────────────
    #  SD-1 — EXPANSION STRATEGY
    # ──────────────────────────────────────────────────────────
    with sd1:
        _crs_strat = sum(
            st.session_state.country_data[target_country][k] * CRS_WEIGHTS[k]
            for k in CRS_WEIGHTS
        )
        _ma_score = st.session_state.country_data[target_country]["market_attractiveness"]

        # ── Entry Mode Recommendation ──────────────────────
        st.subheader("Entry Mode Recommendation")
        if _crs_strat >= 72:
            entry_mode = "Wholly-Owned Subsidiary (WOS)"
            entry_color = "#27ae60"
            entry_icon = "🏢"
            entry_rationale = (
                f"{target_country} has a high Country Risk Score ({_crs_strat:.0f}/100), indicating strong "
                "institutional quality and macro stability. A wholly-owned subsidiary maximises control, "
                "IP protection, and profit repatriation flexibility."
            )
            entry_checklist = [
                "Register a local entity (see Regulatory tab for specific requirements)",
                "Appoint a local director if required by jurisdiction",
                "Set up a dedicated local bank account (consider DBS multi-currency)",
                "Establish local HR and payroll compliance from day one",
            ]
        elif _crs_strat >= 52:
            entry_mode = "Joint Venture (JV) with Local Partner"
            entry_color = "#2980b9"
            entry_icon = "🤝"
            entry_rationale = (
                f"{target_country}'s Country Risk Score is {_crs_strat:.0f}/100 — moderate. A JV with a "
                "reputable local partner reduces regulatory and operational risk while leveraging local "
                "market knowledge and relationships."
            )
            entry_checklist = [
                "Conduct thorough partner due diligence (financial + reputational)",
                "Define equity split, governance rights, and exit clauses in the JV agreement",
                "Agree on IP ownership and licensing terms upfront",
                "Structure inter-company pricing carefully for tax efficiency",
            ]
        elif _crs_strat >= 36:
            entry_mode = "Franchise / Master License Agreement"
            entry_color = "#f39c12"
            entry_icon = "📄"
            entry_rationale = (
                f"{target_country} has an elevated risk profile (CRS: {_crs_strat:.0f}/100). A franchise or "
                "master license limits capital exposure while enabling market penetration through a "
                "trusted local operator."
            )
            entry_checklist = [
                "Draft a robust franchise agreement with performance KPIs and auditing rights",
                "Protect IP and brand standards contractually with penalty clauses",
                "Retain periodic site inspection and termination rights",
                "Structure royalty fees to account for local withholding tax",
            ]
        else:
            entry_mode = "Distributor / Agent Model"
            entry_color = "#e74c3c"
            entry_icon = "📦"
            entry_rationale = (
                f"{target_country}'s risk score ({_crs_strat:.0f}/100) signals significant political, "
                "macro, or institutional risk. A distributor or commission-agent model limits capital "
                "at risk while testing market demand with minimal exposure."
            )
            entry_checklist = [
                "Appoint a vetted local distributor with clear exclusivity and termination terms",
                "Include IP reversion rights if distributor relationship ends",
                "Cap consignment inventory exposure",
                "Consider trade credit insurance (e.g. Coface, Euler Hermes)",
            ]

        em_c1, em_c2 = st.columns([1, 2])
        with em_c1:
            st.markdown(
                f"""<div style="background:{entry_color}22; border:2px solid {entry_color};
                    border-radius:12px; padding:24px; text-align:center; height:100%;">
                    <div style="font-size:52px;">{entry_icon}</div>
                    <div style="font-size:17px; font-weight:700; color:{entry_color}; margin-top:10px;">
                        {entry_mode}
                    </div>
                    <div style="font-size:12px; color:#666; margin-top:6px;">
                        Recommended · CRS {_crs_strat:.0f}/100
                    </div>
                </div>""",
                unsafe_allow_html=True,
            )
        with em_c2:
            st.markdown(f"**Rationale:** {entry_rationale}")
            st.markdown("**Action Checklist:**")
            for item in entry_checklist:
                st.markdown(f"- ✅ {item}")

        st.markdown("---")

        # ── Interactive Pricing Model ──────────────────────
        st.subheader("Interactive Pricing Model")
        st.caption("Adjust the sliders to explore how unit price affects demand, revenue, and gross profit.")

        _cost_pct_default = int((cogs3 / rev3 * 100)) if rev3 > 0 and cogs3 > 0 else 60

        pm_c1, pm_c2 = st.columns([1, 2])
        with pm_c1:
            st.markdown("**Pricing Parameters**")
            unit_price = st.slider(
                "Unit Selling Price (SGD)", min_value=10, max_value=500,
                value=100, step=5, key="pm_price",
            )
            unit_cost_pct = st.slider(
                "Unit Cost (% of current price)", min_value=10, max_value=90,
                value=_cost_pct_default, step=1, key="pm_cost_pct",
                help="Auto-populated from your COGS/Revenue if entered above.",
            )
            market_size = st.slider(
                "Market Size (units / month)", min_value=100, max_value=50_000,
                value=5_000, step=100, key="pm_market_size",
            )
            price_elasticity = st.slider(
                "Price Elasticity of Demand", min_value=-3.0, max_value=-0.2,
                value=-1.2, step=0.1, key="pm_elasticity",
                help="How sensitive demand is to price. More negative = more price-sensitive.",
            )

        with pm_c2:
            prices = np.linspace(max(5.0, unit_price * 0.3), unit_price * 2.5, 200)
            unit_cost_sgd = unit_price * (unit_cost_pct / 100)
            demand_curve  = np.clip(market_size * (prices / unit_price) ** price_elasticity, 0, market_size * 4)
            supply_curve  = np.clip(market_size * (prices / unit_price) ** 0.8, 0, market_size * 4)
            gp_curve      = (prices - unit_cost_sgd) * demand_curve

            diff = demand_curve - supply_curve
            eq_idx   = int(np.argmin(np.abs(diff)))
            eq_price = prices[eq_idx]
            opt_idx  = int(np.argmax(gp_curve))
            opt_price = prices[opt_idx]

            curr_demand  = float(market_size)
            curr_revenue = unit_price * curr_demand
            curr_gm      = (1 - unit_cost_pct / 100) * 100

            fig_pricing = go.Figure()
            fig_pricing.add_trace(go.Scatter(
                x=prices, y=demand_curve, name="Demand",
                line=dict(color="#2980b9", width=2.5),
            ))
            fig_pricing.add_trace(go.Scatter(
                x=prices, y=supply_curve, name="Supply",
                line=dict(color="#27ae60", width=2.5),
            ))
            fig_pricing.add_vline(
                x=unit_price, line_dash="dash", line_color="#e74c3c", line_width=2,
                annotation_text=f"Current SGD {unit_price}", annotation_position="top right",
            )
            fig_pricing.add_vline(
                x=eq_price, line_dash="dot", line_color="#f39c12", line_width=1.5,
                annotation_text=f"Equilibrium SGD {eq_price:.0f}", annotation_position="top left",
            )
            fig_pricing.add_vline(
                x=opt_price, line_dash="dot", line_color="#8e44ad", line_width=1.5,
                annotation_text=f"Max Profit SGD {opt_price:.0f}", annotation_position="bottom right",
            )
            fig_pricing.update_layout(
                title="Demand & Supply Curve",
                xaxis_title="Unit Price (SGD)",
                yaxis_title="Quantity (units / month)",
                height=330,
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                margin=dict(t=50, b=40, l=50, r=30),
            )
            st.plotly_chart(fig_pricing, use_container_width=True)

        mc1, mc2, mc3, mc4 = st.columns(4)
        mc1.metric("Est. Monthly Demand", f"{curr_demand:,.0f} units")
        mc2.metric("Est. Monthly Revenue", f"SGD {curr_revenue:,.0f}")
        mc3.metric("Gross Margin at Price", f"{curr_gm:.1f}%")
        mc4.metric("Profit-Maximising Price", f"SGD {opt_price:.0f}")

        st.markdown("---")

        # ── Supplier & Distribution ────────────────────────
        st.subheader("Supplier & Distribution Strategy")
        sup_c1, sup_c2 = st.columns(2)

        with sup_c1:
            st.markdown("**Supplier Recommendations**")
            if _ma_score >= 70:
                sup_items = [
                    ("Local manufacturing / contract production", "Reduces import duties and shortens lead times"),
                    ("Regional hub sourcing (SG or MY)", "Lower FX risk with a tighter supply chain"),
                    ("Dual-source strategy", "Mitigates single-supplier concentration risk"),
                ]
            else:
                sup_items = [
                    ("Source from Singapore or a third-country supplier", "Avoids high local import complexity"),
                    ("Use bonded warehouse / FTZ if available", "Defers duty payments and aids cash flow"),
                    ("Retain SG-based fulfilment backup", "Hedge against local supply disruption"),
                ]
            for title, desc in sup_items:
                st.markdown(
                    f"""<div style="border-left:4px solid #2980b9; background:#ebf5fb;
                        padding:10px 14px; border-radius:4px; margin-bottom:8px;">
                        <strong style="font-size:14px;">{title}</strong><br/>
                        <span style="color:#555; font-size:13px;">{desc}</span>
                    </div>""",
                    unsafe_allow_html=True,
                )

        with sup_c2:
            st.markdown("**Distribution Channels**")
            if _crs_strat >= 65:
                dist_channels = [
                    ("Direct Sales Force", "High control; suitable for B2B and enterprise contracts"),
                    ("E-commerce / Digital Marketplace", "Low overhead; ideal for consumer products"),
                    ("Retail Partnerships", "Local brand credibility and physical reach"),
                ]
            else:
                dist_channels = [
                    ("Exclusive Local Distributor", "Reduces compliance burden and local legal exposure"),
                    ("Online-first via regional platforms", "Lower fixed cost with scalable reach"),
                    ("Strategic alliance / co-marketing", "Share costs, leverage existing local network"),
                ]
            for ch, desc in dist_channels:
                st.markdown(
                    f"""<div style="border-left:4px solid #27ae60; background:#eafaf1;
                        padding:10px 14px; border-radius:4px; margin-bottom:8px;">
                        <strong style="font-size:14px;">{ch}</strong><br/>
                        <span style="color:#555; font-size:13px;">{desc}</span>
                    </div>""",
                    unsafe_allow_html=True,
                )

    # ──────────────────────────────────────────────────────────
    #  SD-2 — HEDGING SOLUTIONS
    # ──────────────────────────────────────────────────────────
    with sd2:
        _tgt_ccy = COUNTRY_CURRENCIES.get(target_country, "USD")
        _fx_rate  = FX_VS_SGD.get(target_country, 1.0)

        st.subheader(f"FX Risk Management — SGD / {_tgt_ccy}")

        hdr_c1, hdr_c2, hdr_c3 = st.columns(3)
        hdr_c1.metric("Currency Pair", f"SGD / {_tgt_ccy}")
        hdr_c2.metric("Indicative Rate (1 SGD =)", f"{_fx_rate:,.4g} {_tgt_ccy}")
        hdr_c3.metric(f"Indicative Rate (1 {_tgt_ccy} =)", f"SGD {1/_fx_rate:.5f}" if _fx_rate else "N/A")
        st.caption("Rates are indicative reference values for scenario modelling only. Always verify live rates before transacting.")

        st.markdown("---")
        st.markdown("#### FX Scenario Analysis")

        hfx_c1, hfx_c2 = st.columns([1, 2])
        with hfx_c1:
            fx_exposure = st.number_input(
                "Annual FX Exposure (SGD equivalent)",
                min_value=0.0,
                value=float(rev3) * 0.5 if rev3 > 0 else 100_000.0,
                step=10_000.0, format="%.0f", key="fx_exposure",
                help="Estimated annual revenue or payables denominated in the target currency.",
            )
            hedge_ratio = st.slider(
                "Hedging Ratio (%)", min_value=0, max_value=100,
                value=50, step=5, key="hedge_ratio",
                help="Proportion of FX exposure covered by hedging instruments.",
            )
            hedge_cost_bps = st.slider(
                "Hedging Cost (basis points p.a.)", min_value=10, max_value=200,
                value=60, step=10, key="hedge_cost_bps",
                help="Approximate annual cost of hedging (forward premium or option premium).",
            )

        with hfx_c2:
            scenarios = [-20, -15, -10, -5, 0, 5, 10, 15, 20]
            unhedged   = [fx_exposure * s / 100 for s in scenarios]
            hedged     = [
                fx_exposure * s / 100 * (1 - hedge_ratio / 100) - fx_exposure * (hedge_cost_bps / 10_000)
                for s in scenarios
            ]
            sc_labels  = [f"{s:+d}%" for s in scenarios]
            uh_colors  = ["#e74c3c" if v < 0 else "#27ae60" for v in unhedged]
            h_colors   = ["#c0392b" if v < 0 else "#1e8449" for v in hedged]

            fig_fx = go.Figure()
            fig_fx.add_trace(go.Bar(
                x=sc_labels, y=unhedged, name="Unhedged P&L (SGD)",
                marker_color=uh_colors, opacity=0.75,
            ))
            fig_fx.add_trace(go.Bar(
                x=sc_labels, y=hedged,
                name=f"Hedged P&L ({hedge_ratio}% cover, {hedge_cost_bps} bps cost)",
                marker_color=h_colors,
            ))
            fig_fx.update_layout(
                title=f"FX P&L Impact — SGD/{_tgt_ccy} Rate Scenarios",
                xaxis_title=f"SGD/{_tgt_ccy} Rate Change (%)",
                yaxis_title="P&L Impact (SGD)",
                barmode="group", height=340,
                margin=dict(t=50, b=40, l=55, r=20),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
            )
            st.plotly_chart(fig_fx, use_container_width=True)

        st.markdown("---")
        st.markdown("#### Hedging Instrument Options")

        hed_c1, hed_c2, hed_c3 = st.columns(3)
        instruments = [
            {
                "name": "FX Forward Contract",
                "icon": "📅",
                "color": "#1a5276",
                "desc": "Lock in today's exchange rate for a future transaction. Eliminates both upside and downside FX risk entirely.",
                "pros": ["Certainty of cash flows", "No upfront premium required", "Available in most currency pairs"],
                "cons": ["Loses upside if rate moves favourably", "Inflexible once locked in"],
                "best_for": "Large, predictable transactions (e.g. regular supplier payments)",
            },
            {
                "name": "FX Vanilla Options (Put/Call)",
                "icon": "⚖️",
                "color": "#1e8449",
                "desc": "Buy the right (not obligation) to exchange at a fixed rate. Protects downside while retaining upside.",
                "pros": ["Retains upside FX gain", "Flexible — let it expire if not needed", "Customisable strike and tenor"],
                "cons": ["Option premium is an upfront cost", "More complex to manage"],
                "best_for": "When you want downside protection but expect a favourable rate move",
            },
            {
                "name": "Natural Hedging",
                "icon": "🌿",
                "color": "#6c3483",
                "desc": "Match revenues and costs in the same currency to reduce net FX exposure without financial instruments.",
                "pros": ["No hedging cost", "Operationally intuitive", "Reduces ongoing management burden"],
                "cons": ["Requires a local cost base (staff, suppliers)", "Rarely achieves 100% offset"],
                "best_for": "Businesses establishing local operations with a significant local cost base",
            },
        ]
        for col, inst in zip([hed_c1, hed_c2, hed_c3], instruments):
            pros_html = "".join(f"<li>✅ {p}</li>" for p in inst["pros"])
            cons_html = "".join(f"<li>⚠️ {c}</li>" for c in inst["cons"])
            col.markdown(
                f"""<div style="border:2px solid {inst['color']}; border-radius:10px; padding:18px;">
                    <div style="font-size:36px; text-align:center;">{inst['icon']}</div>
                    <h4 style="color:{inst['color']}; text-align:center; font-size:15px; margin:8px 0;">
                        {inst['name']}
                    </h4>
                    <p style="font-size:12px; color:#444; min-height:56px;">{inst['desc']}</p>
                    <ul style="font-size:12px; padding-left:16px; margin:4px 0;">{pros_html}</ul>
                    <ul style="font-size:12px; padding-left:16px; margin:4px 0;">{cons_html}</ul>
                    <div style="background:{inst['color']}22; border-radius:6px; padding:8px;
                        font-size:12px; margin-top:10px;">
                        <strong>Best for:</strong> {inst['best_for']}
                    </div>
                </div>""",
                unsafe_allow_html=True,
            )

    # ──────────────────────────────────────────────────────────
    #  SD-3 — CAPITAL SOLUTIONS
    # ──────────────────────────────────────────────────────────
    with sd3:
        st.subheader("DBS Financing Solutions for SME Expansion")
        st.caption("Select a product below to explore rates, payment structure, and a full monthly repayment schedule.")

        # ── Product Summary Cards ──────────────────────────
        cap_c1, cap_c2, cap_c3 = st.columns(3)
        for col, pname in zip([cap_c1, cap_c2, cap_c3], DBS_PRODUCTS):
            p = DBS_PRODUCTS[pname]
            feat_html = "".join(f"<li style='font-size:12px; margin-bottom:3px;'>{f}</li>" for f in p["features"])
            col.markdown(
                f"""<div style="border:2px solid {p['color']}; border-radius:10px; padding:16px; height:100%;">
                    <div style="font-size:32px; text-align:center;">{p['icon']}</div>
                    <h4 style="color:{p['color']}; text-align:center; font-size:14px; margin:8px 0;">{pname}</h4>
                    <p style="font-size:12px; color:#444; min-height:52px;">{p['description']}</p>
                    <div style="background:{p['color']}22; border-radius:4px; padding:7px;
                        text-align:center; margin:8px 0;">
                        <strong style="color:{p['color']}; font-size:15px;">{p['rate_pa']:.2f}% p.a.</strong>
                        <span style="font-size:11px; color:#666;"> · Max SGD {p['max_loan']:,}</span><br/>
                        <span style="font-size:11px; color:#666;">{p['structure']}</span>
                    </div>
                    <ul style="padding-left:16px; margin:0;">{feat_html}</ul>
                </div>""",
                unsafe_allow_html=True,
            )

        st.markdown("---")
        st.markdown("#### Loan Calculator")

        calc_c1, calc_c2 = st.columns([1, 2])
        with calc_c1:
            selected_product = st.selectbox(
                "Select DBS Product", list(DBS_PRODUCTS.keys()), key="dbs_product_select",
            )
            _p = DBS_PRODUCTS[selected_product]
            loan_amount = st.number_input(
                "Loan Amount (SGD)",
                min_value=10_000.0, max_value=float(_p["max_loan"]),
                value=min(200_000.0, float(_p["max_loan"])),
                step=10_000.0, key="dbs_loan_amount",
            )
            loan_tenure = st.slider(
                "Tenure (months)",
                min_value=_p["tenure_min"], max_value=_p["tenure_max"],
                value=min(36, _p["tenure_max"]), step=1, key="dbs_tenure",
            )
            st.markdown(f"**Structure:** {_p['structure']}")
            st.markdown(f"**Rate:** {_p['rate_pa']:.2f}% p.a. ({_p['rate_pa']/12:.3f}% / month)")

        with calc_c2:
            monthly_rate = _p["rate_pa"] / 100 / 12
            months_range = list(range(1, loan_tenure + 1))

            if _p["structure"] == "Amortising":
                if monthly_rate > 0:
                    monthly_payment = (
                        loan_amount * monthly_rate * (1 + monthly_rate) ** loan_tenure
                        / ((1 + monthly_rate) ** loan_tenure - 1)
                    )
                else:
                    monthly_payment = loan_amount / loan_tenure

                # Build schedule
                sched_raw: List[Dict] = []
                bal = loan_amount
                for m in months_range:
                    int_m  = bal * monthly_rate
                    pri_m  = monthly_payment - int_m
                    bal    = max(0.0, bal - pri_m)
                    sched_raw.append({"m": m, "pri": pri_m, "int": int_m, "bal": bal})

                total_payment  = monthly_payment * loan_tenure
                total_interest = total_payment - loan_amount

                km1, km2, km3 = st.columns(3)
                km1.metric("Monthly Payment",     f"SGD {monthly_payment:,.2f}")
                km2.metric("Total Interest Paid", f"SGD {total_interest:,.2f}")
                km3.metric("Total Cost of Loan",  f"SGD {total_payment:,.2f}")

                fig_amort = go.Figure()
                fig_amort.add_trace(go.Bar(
                    x=months_range,
                    y=[r["pri"] for r in sched_raw],
                    name="Principal",
                    marker_color=_p["color"],
                ))
                fig_amort.add_trace(go.Bar(
                    x=months_range,
                    y=[r["int"] for r in sched_raw],
                    name="Interest",
                    marker_color="#aed6f1",
                ))
                fig_amort.update_layout(
                    title="Monthly Payment Breakdown",
                    xaxis_title="Month", yaxis_title="SGD",
                    barmode="stack", height=300,
                    margin=dict(t=40, b=30, l=50, r=20),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02),
                )
                st.plotly_chart(fig_amort, use_container_width=True)

                display_n = min(12, loan_tenure)
                st.markdown(f"**Repayment Schedule (first {display_n} months)**")
                st.dataframe(
                    pd.DataFrame([{
                        "Month": r["m"],
                        "Monthly Payment (SGD)": f"{monthly_payment:,.2f}",
                        "Principal (SGD)": f"{r['pri']:,.2f}",
                        "Interest (SGD)": f"{r['int']:,.2f}",
                        "Remaining Balance (SGD)": f"{r['bal']:,.2f}",
                    } for r in sched_raw[:display_n]]),
                    use_container_width=True, hide_index=True,
                )
                if loan_tenure > 12:
                    st.caption(f"Showing first 12 of {loan_tenure} months.")

            else:  # Interest-only revolving
                monthly_interest = loan_amount * monthly_rate
                total_interest   = monthly_interest * loan_tenure

                km1, km2, km3 = st.columns(3)
                km1.metric("Monthly Interest",          f"SGD {monthly_interest:,.2f}")
                km2.metric("Total Interest (tenure)",   f"SGD {total_interest:,.2f}")
                km3.metric("Bullet Principal at End",   f"SGD {loan_amount:,.2f}")

                bullet_vals = [0.0] * loan_tenure
                bullet_vals[-1] = loan_amount

                fig_io = go.Figure()
                fig_io.add_trace(go.Bar(
                    x=months_range,
                    y=[monthly_interest] * loan_tenure,
                    name="Monthly Interest",
                    marker_color=_p["color"],
                ))
                fig_io.add_trace(go.Bar(
                    x=months_range, y=bullet_vals,
                    name="Principal (Bullet Repayment)",
                    marker_color="#d35400",
                ))
                fig_io.update_layout(
                    title="Monthly Payment Breakdown (Interest-Only + Bullet)",
                    xaxis_title="Month", yaxis_title="SGD",
                    barmode="stack", height=300,
                    margin=dict(t=40, b=30, l=50, r=20),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02),
                )
                st.plotly_chart(fig_io, use_container_width=True)

                display_n = min(12, loan_tenure)
                st.markdown(f"**Repayment Schedule (first {display_n} months)**")
                st.dataframe(
                    pd.DataFrame([{
                        "Month": m,
                        "Interest Payment (SGD)": f"{monthly_interest:,.2f}",
                        "Principal Repayment (SGD)": f"{loan_amount:,.2f}" if m == loan_tenure else "—",
                        "Outstanding Balance (SGD)": f"{loan_amount:,.2f}",
                    } for m in range(1, display_n + 1)]),
                    use_container_width=True, hide_index=True,
                )

    # ──────────────────────────────────────────────────────────
    #  SD-4 — REGULATORY RISK
    # ──────────────────────────────────────────────────────────
    with sd4:
        _reg = REGULATORY_DATA.get(target_country, {})
        if not _reg:
            st.warning(f"No regulatory data available for {target_country}.")
        else:
            st.subheader(f"Regulatory Landscape — {target_country}")

            r_score = _reg.get("risk_score", 50)
            r_level = _reg.get("risk_level", "Medium")
            r_col   = "#27ae60" if r_score >= 70 else "#f39c12" if r_score >= 45 else "#e74c3c"

            # ── Summary Banner ─────────────────────────────
            rh1, rh2, rh3, rh4 = st.columns(4)
            rh1.markdown(
                f"""<div style="background:{r_col}22; border:2px solid {r_col};
                    border-radius:8px; padding:16px; text-align:center;">
                    <div style="font-size:32px; font-weight:700; color:{r_col};">{r_score}</div>
                    <div style="font-size:12px; color:#555;">Regulatory Risk Score</div>
                    <div style="font-weight:600; color:{r_col}; font-size:14px; margin-top:4px;">
                        {r_level} Risk
                    </div>
                </div>""",
                unsafe_allow_html=True,
            )
            rh2.metric("Corporate Tax",   _reg.get("corp_tax", "—"))
            rh3.metric("GST / VAT",       _reg.get("gst_vat", "—"))
            rh4.metric(
                "DTA with Singapore",
                "Yes ✅" if _reg.get("dta_with_sg") else "No ❌",
                help="Double Taxation Agreement reduces withholding taxes and clarifies tax residency.",
            )

            st.markdown("---")
            reg_c1, reg_c2 = st.columns(2)

            with reg_c1:
                st.markdown("#### 📦 Import / Export & Tax Rules")
                st.markdown(
                    f"""<div style="background:#fdfefe; border:1px solid #d5d8dc;
                        border-radius:8px; padding:16px; font-size:13px;">
                    <table style="width:100%; border-collapse:collapse;">
                        <tr style="border-bottom:1px solid #eee;">
                            <td style="color:#666; padding:6px 0; width:45%;">Import Duties</td>
                            <td><strong>{_reg.get('import_duty','—')}</strong></td>
                        </tr>
                        <tr style="border-bottom:1px solid #eee;">
                            <td style="color:#666; padding:6px 0;">Withholding Tax</td>
                            <td><strong>{_reg.get('withholding_tax','—')}</strong></td>
                        </tr>
                        <tr>
                            <td style="color:#666; padding:6px 0;">FDI Restrictions</td>
                            <td><strong>{_reg.get('fdi_restrictions','—')}</strong></td>
                        </tr>
                    </table>
                    </div>""",
                    unsafe_allow_html=True,
                )

                st.markdown("#### ⚠️ Key Regulatory Risks")
                for risk in _reg.get("key_risks", []):
                    st.markdown(
                        f"""<div style="border-left:4px solid #e74c3c; background:#fdedec;
                            padding:8px 12px; border-radius:4px; margin-bottom:6px; font-size:13px;">
                            ⚠️ {risk}
                        </div>""",
                        unsafe_allow_html=True,
                    )

            with reg_c2:
                st.markdown("#### 🪪 Licensing Requirements")
                for lic in _reg.get("key_licenses", []):
                    st.markdown(
                        f"""<div style="border-left:4px solid #2980b9; background:#ebf5fb;
                            padding:8px 12px; border-radius:4px; margin-bottom:6px; font-size:13px;">
                            📋 {lic}
                        </div>""",
                        unsafe_allow_html=True,
                    )

                st.markdown("#### 📊 Risk Dimension Overview")
                _cd_reg = st.session_state.country_data[target_country]
                radar_cats = [
                    "Macro Stability", "Institutional Quality", "Credit Risk",
                    "Market Attractiveness", "Regulatory Score",
                ]
                radar_vals = [
                    _cd_reg["macro_stability"],
                    _cd_reg["institutional_quality"],
                    _cd_reg["credit_risk"],
                    _cd_reg["market_attractiveness"],
                    r_score,
                ]
                # Close the polygon
                radar_vals_c = radar_vals + [radar_vals[0]]
                radar_cats_c = radar_cats + [radar_cats[0]]

                # Convert hex to rgba for transparent fill
                _r, _g, _b = int(r_col[1:3], 16), int(r_col[3:5], 16), int(r_col[5:7], 16)
                _fillcolor = f"rgba({_r}, {_g}, {_b}, 0.2)"

                fig_radar = go.Figure(go.Scatterpolar(
                    r=radar_vals_c,
                    theta=radar_cats_c,
                    fill="toself",
                    line_color=r_col,
                    fillcolor=_fillcolor,
                    name=target_country,
                ))
                fig_radar.update_layout(
                    polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
                    showlegend=False,
                    height=310,
                    margin=dict(t=20, b=20, l=30, r=30),
                )
                st.plotly_chart(fig_radar, use_container_width=True)

# ══════════════════════════════════════════════════════════════
#  TAB 7 — MONTE CARLO NPV SIMULATION
# ══════════════════════════════════════════════════════════════
with tab_mc:
    st.header("Monte Carlo NPV Simulation")
    st.caption(
        "Simulate thousands of possible expansion scenarios using your financial data and the selected "
        "country's risk profile. Includes a real option to abandon the project at a chosen year."
    )

    if rev3 <= 0:
        st.warning(
            "⚠️ Please enter at least **Year 3 Revenue** in the Income Statement tab "
            "before running simulations."
        )
    else:
        # ── Derive historical parameters from submitted financials ──────
        _mc_revs   = [r for r in [rev1, rev2, rev3] if r > 0]
        _mc_n      = len(_mc_revs) - 1
        _hist_cagr = ((_mc_revs[-1] / _mc_revs[0]) ** (1 / _mc_n) - 1) * 100 if _mc_n > 0 else 5.0
        _ebit_mg   = (ebit3 / rev3 * 100) if rev3 > 0 else 15.0
        _ebt_calc  = ebit3 - int_exp3
        _tax_r     = (tax3 / _ebt_calc) if _ebt_calc > 0 and tax3 > 0 else 0.20
        _capex_r   = (capex / rev3 * 100) if rev3 > 0 else 5.0
        _nwc_r     = ((ar + inventory - ap) / rev3 * 100) if rev3 > 0 else 15.0

        # WACC estimate from balance sheet
        _debt_mc = short_term_debt + long_term_debt
        _eq_mc   = total_equity if total_equity > 0 else rev3 * 0.5
        _tot_cap = max(_debt_mc + _eq_mc, 1.0)
        _ke      = 9.5   # CAPM: rf 3.5% + beta 1.1 × ERP 5.5%
        _kd      = 5.0 * (1 - _tax_r)
        _wacc_est = float(np.clip(_ke * _eq_mc / _tot_cap + _kd * _debt_mc / _tot_cap, 5.0, 25.0))

        # Country risk premium derived from CRS
        _crs_mc = sum(st.session_state.country_data[target_country][k] * CRS_WEIGHTS[k] for k in CRS_WEIGHTS)
        _crp    = max(0.0, (70.0 - _crs_mc) / 100 * 12.0)   # 0% at CRS≥70, up to ~12% at CRS=0

        # ── Parameter panel (left) / Context panel (right) ─────────────
        par_c, ctx_c = st.columns([2, 3])

        with par_c:
            st.markdown("#### Simulation Parameters")

            with st.expander("📐 Model Setup", expanded=True):
                initial_inv = st.number_input(
                    "Initial Investment (SGD)", min_value=0.0,
                    value=float(capex) * 3 if capex > 0 else float(rev3) * 0.5,
                    step=10_000.0, key="mc_init_inv",
                    help="Total capital outlay at Year 0 for the expansion.",
                )
                horizon_yrs = st.slider("Projection Horizon (years)", 3, 15, 7, key="mc_horizon")
                n_sims = st.selectbox(
                    "Number of Simulations",
                    [500, 1_000, 2_000, 5_000, 10_000], index=2, key="mc_n_sims",
                )

            with st.expander("📈 Revenue & Margin Assumptions", expanded=True):
                growth_mean = st.slider(
                    "Mean Revenue Growth (% p.a.)", -10.0, 30.0,
                    float(round(float(np.clip(_hist_cagr, -10.0, 30.0)), 1)), 0.5,
                    key="mc_gmean", help=f"Auto-populated from historical CAGR: {_hist_cagr:.1f}%",
                )
                growth_std = st.slider(
                    "Growth Uncertainty σ (%)", 1.0, 20.0, 8.0, 0.5, key="mc_gstd",
                    help="Wider = more volatile annual revenue growth.",
                )
                margin_mean = st.slider(
                    "Mean EBIT Margin (% of revenue)", -10.0, 40.0,
                    float(round(float(np.clip(_ebit_mg, -10.0, 40.0)), 1)), 0.5,
                    key="mc_mmean", help=f"Auto-populated from Year 3 EBIT margin: {_ebit_mg:.1f}%",
                )
                margin_std = st.slider(
                    "Margin Uncertainty σ (%)", 0.5, 15.0, 4.0, 0.5, key="mc_mstd",
                )

            with st.expander("💰 Discount Rate", expanded=True):
                wacc_input = st.slider(
                    "WACC (%)", 5.0, 25.0,
                    float(round(_wacc_est, 1)), 0.5, key="mc_wacc",
                    help=f"Estimated from balance sheet: {_wacc_est:.1f}%",
                )
                st.markdown(
                    f"""<div style="background:#ebf5fb; border:1px solid #2980b9;
                        border-radius:6px; padding:10px; font-size:13px; margin-top:4px;">
                        <strong>Country Risk Premium (CRP):</strong> +{_crp:.1f}%<br/>
                        <strong>Effective Discount Rate:</strong> {wacc_input + _crp:.2f}%<br/>
                        <span style="color:#666;">{target_country} CRS: {_crs_mc:.0f}/100 —
                        higher risk → higher hurdle rate</span>
                    </div>""",
                    unsafe_allow_html=True,
                )

            with st.expander("🚪 Abandonment Option (Real Option)", expanded=True):
                abandon_year = st.slider(
                    "Earliest Year to Abandon  (0 = no option)",
                    0, max(1, horizon_yrs - 1), 0, 1, key="mc_abandon",
                    help="At this year, management can choose to exit and recover salvage value.",
                )
                if abandon_year > 0:
                    salvage_pct = st.slider(
                        "Salvage Recovery (% of initial investment)", 0, 80, 30, 5,
                        key="mc_salvage",
                        help="Capital recovered on liquidation if you abandon.",
                    )
                else:
                    salvage_pct = 0
                    st.caption("Set abandon year > 0 to activate the real option.")

        with ctx_c:
            st.markdown("#### Auto-detected from Your Financial Data")
            ad1, ad2, ad3 = st.columns(3)
            ad1.metric("Base Revenue (Yr 3)", f"SGD {rev3:,.0f}")
            ad1.metric("Historical CAGR",     f"{_hist_cagr:.1f}%")
            ad2.metric("EBIT Margin",         f"{_ebit_mg:.1f}%")
            ad2.metric("Effective Tax Rate",  f"{_tax_r*100:.1f}%")
            ad3.metric("Capex / Revenue",     f"{_capex_r:.1f}%")
            ad3.metric("Est. WACC",           f"{_wacc_est:.1f}%")

            st.markdown("#### Country Risk Context")
            _cd_mc = st.session_state.country_data[target_country]
            _crs_vals = [_cd_mc[k] for k in ["macro_stability","institutional_quality","credit_risk","market_attractiveness"]]
            _crs_lbls = ["Macro Stability","Institutional Quality","Credit Risk","Mkt Attractiveness"]
            fig_ctx = go.Figure(go.Bar(
                x=_crs_vals, y=_crs_lbls, orientation="h",
                marker_color=[_score_color(v) for v in _crs_vals],
                text=[f"{v:.0f}" for v in _crs_vals], textposition="outside",
            ))
            fig_ctx.update_layout(
                title=f"{target_country}  ·  CRS {_crs_mc:.0f}/100  ·  CRP +{_crp:.1f}%",
                xaxis=dict(range=[0, 115]),
                height=230, margin=dict(t=45, b=10, l=140, r=60),
            )
            st.plotly_chart(fig_ctx, use_container_width=True)

            if abandon_year > 0:
                salvage_sgd = initial_inv * salvage_pct / 100
                st.markdown(
                    f"""<div style="background:#fef9e7; border:1px solid #f39c12;
                        border-radius:6px; padding:12px; font-size:13px;">
                        <strong>Abandonment Option Active</strong><br/>
                        Exit point: <strong>end of Year {abandon_year}</strong><br/>
                        Salvage recovery: <strong>{salvage_pct}% → SGD {salvage_sgd:,.0f}</strong><br/>
                        The simulation computes the <em>real option value</em> this flexibility provides.
                    </div>""",
                    unsafe_allow_html=True,
                )

        # ── Full-width run button ───────────────────────────────────────
        st.markdown("---")
        run_mc_btn = st.button(
            f"▶️  Run {n_sims:,} Monte Carlo Simulations  →  {target_country}",
            type="primary", use_container_width=True, key="mc_run_btn",
        )

        # ── Execute simulation ─────────────────────────────────────────
        if run_mc_btn:
            with st.spinner(f"Simulating {n_sims:,} expansion paths for {target_country}…"):
                _r_eff = (wacc_input + _crp) / 100
                rng    = np.random.default_rng(42)

                # Simulate growth rates and EBIT margins: shape (n_sims, horizon_yrs)
                sim_growths = rng.normal(growth_mean / 100, growth_std / 100, (n_sims, horizon_yrs))
                sim_margins = np.clip(
                    rng.normal(margin_mean / 100, margin_std / 100, (n_sims, horizon_yrs)),
                    -0.5, 0.5,
                )

                # Revenue paths
                rev_paths = rev3 * np.cumprod(1 + sim_growths, axis=1)

                # Revenue delta for NWC changes
                rev_prev_mat = np.column_stack([np.full(n_sims, rev3), rev_paths[:, :-1]])
                delta_rev    = rev_paths - rev_prev_mat

                # Cash flow components (broadcast da_series across all sims)
                da_series   = np.array([da3 * (1.02 ** t) for t in range(1, horizon_yrs + 1)])
                nopat_paths = rev_paths * sim_margins * (1 - _tax_r)
                nwc_out     = (_nwc_r / 100) * np.maximum(0.0, delta_rev)
                maint_capex = (_capex_r / 100) * rev_paths
                cf_paths    = nopat_paths + da_series - nwc_out - maint_capex

                # Discount all cash flows
                disc_f   = np.array([1.0 / (1 + _r_eff) ** t for t in range(1, horizon_yrs + 1)])
                disc_cfs = cf_paths * disc_f   # (n_sims, horizon_yrs)

                # NPV without abandonment option
                npv_base = -initial_inv + disc_cfs.sum(axis=1)

                # NPV with abandonment option (real option)
                if abandon_year > 0 and abandon_year < horizon_yrs:
                    pv_before_ab = disc_cfs[:, :abandon_year].sum(axis=1)
                    pv_after_ab  = disc_cfs[:, abandon_year:].sum(axis=1)
                    salvage_pv   = (initial_inv * salvage_pct / 100) / (1 + _r_eff) ** abandon_year
                    # At the option year: take max(continue, abandon)
                    npv_option = -initial_inv + pv_before_ab + np.maximum(pv_after_ab, salvage_pv)
                else:
                    npv_option = npv_base.copy()

                st.session_state.mc_results = {
                    "npv_base":    npv_base,
                    "npv_option":  npv_option,
                    "rev_paths":   rev_paths,
                    "cf_paths":    cf_paths,
                    "sim_growths": sim_growths,
                    "sim_margins": sim_margins,
                    "horizon":     horizon_yrs,
                    "abandon_year": abandon_year,
                    "r_eff":       _r_eff,
                    "initial_inv": initial_inv,
                    "n_sims":      n_sims,
                    "params": {
                        "growth_mean": growth_mean, "growth_std": growth_std,
                        "margin_mean": margin_mean, "margin_std": margin_std,
                        "wacc": wacc_input, "crp": _crp, "salvage_pct": salvage_pct,
                    },
                }

        # ── Display results ────────────────────────────────────────────
        if st.session_state.mc_results is not None:
            res   = st.session_state.mc_results
            npv0  = res["npv_base"]
            npv1  = res["npv_option"]
            n_s   = res["n_sims"]
            ab_yr = res["abandon_year"]
            h_yr  = res["horizon"]
            r_e   = res["r_eff"]
            i_inv = res["initial_inv"]
            rp    = res["rev_paths"]
            prm   = res["params"]

            # Pre-compute key stats
            m0  = float(np.mean(npv0));    m1  = float(np.mean(npv1))
            sd0 = float(np.std(npv0));     sd1 = float(np.std(npv1))
            pos0 = float((npv0 > 0).mean() * 100)
            pos1 = float((npv1 > 0).mean() * 100)
            var0 = float(np.percentile(npv0, 5))
            med0 = float(np.median(npv0))
            rov  = m1 - m0   # real option value

            st.markdown("---")
            st.markdown(
                f"### Results — {target_country}  ·  {n_s:,} simulations  ·  "
                f"{h_yr}-year horizon  ·  Discount rate {r_e*100:.2f}%"
            )

            # ── Summary metric cards ───────────────────────────────────
            sm1, sm2, sm3, sm4, sm5 = st.columns(5)
            sm1.metric("Mean NPV",      f"SGD {m0:,.0f}")
            sm2.metric("Median NPV",    f"SGD {med0:,.0f}")
            sm3.metric("Volatility σ",  f"SGD {sd0:,.0f}")
            sm4.metric("P(NPV > 0)",    f"{pos0:.1f}%")
            sm5.metric("VaR (5th pct)", f"SGD {var0:,.0f}", help="5% of simulations produce NPV below this value.")

            if ab_yr > 0:
                st.markdown("**With Abandonment Option:**")
                oa1, oa2, oa3, oa4, oa5 = st.columns(5)
                oa1.metric("Mean NPV (option)",   f"SGD {m1:,.0f}",
                           delta=f"+SGD {rov:,.0f}" if rov >= 0 else f"SGD {rov:,.0f}")
                oa2.metric("Median NPV (option)",  f"SGD {float(np.median(npv1)):,.0f}")
                oa3.metric("Volatility σ (option)", f"SGD {sd1:,.0f}",
                           delta=f"{sd1-sd0:,.0f}", delta_color="inverse")
                oa4.metric("P(NPV > 0, option)",   f"{pos1:.1f}%")
                oa5.metric("Real Option Value",     f"SGD {rov:,.0f}",
                           help="Value added purely by having the flexibility to abandon.")

            st.markdown("---")

            # ── NPV Distribution histogram + percentile sidebar ────────
            hist_c, pct_c = st.columns([3, 1])
            with hist_c:
                fig_hist = go.Figure()
                fig_hist.add_trace(go.Histogram(
                    x=npv0, nbinsx=100, name="Without Abandonment Option",
                    marker_color="#2980b9", opacity=0.65,
                ))
                if ab_yr > 0:
                    fig_hist.add_trace(go.Histogram(
                        x=npv1, nbinsx=100, name="With Abandonment Option",
                        marker_color="#27ae60", opacity=0.65,
                    ))
                fig_hist.add_vline(
                    x=0, line_dash="dash", line_color="#e74c3c", line_width=2,
                    annotation_text="Break-even", annotation_position="top right",
                )
                fig_hist.add_vline(
                    x=m0, line_dash="dot", line_color="#2980b9", line_width=2,
                    annotation_text=f"Mean: SGD {m0:,.0f}", annotation_position="top left",
                )
                fig_hist.update_layout(
                    title=f"NPV Distribution  ·  {n_s:,} Simulations  ·  {target_country}",
                    xaxis_title="NPV (SGD)", yaxis_title="Frequency",
                    barmode="overlay", height=390,
                    margin=dict(t=50, b=40, l=55, r=30),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02),
                )
                st.plotly_chart(fig_hist, use_container_width=True)

            with pct_c:
                st.markdown("**Percentiles**")
                for q in [5, 10, 25, 50, 75, 90, 95]:
                    v = float(np.percentile(npv0, q))
                    c_p = "#27ae60" if v > 0 else "#e74c3c"
                    st.markdown(
                        f"""<div style="background:{c_p}22; border-left:4px solid {c_p};
                            padding:5px 9px; border-radius:4px; margin-bottom:5px; font-size:12px;">
                            <strong>{q}th:</strong> SGD {v:,.0f}
                        </div>""",
                        unsafe_allow_html=True,
                    )

            st.markdown("---")

            # ── Revenue Fan Chart ──────────────────────────────────────
            yr_ax = list(range(1, h_yr + 1))
            p5, p25, p50, p75, p95 = [np.percentile(rp, q, axis=0) for q in [5, 25, 50, 75, 95]]

            fig_fan = go.Figure()
            fig_fan.add_trace(go.Scatter(
                x=yr_ax + yr_ax[::-1], y=list(p95) + list(p5[::-1]),
                fill="toself", fillcolor="rgba(41,128,185,0.10)",
                line=dict(color="rgba(0,0,0,0)"), name="5th–95th Pct", showlegend=True,
            ))
            fig_fan.add_trace(go.Scatter(
                x=yr_ax + yr_ax[::-1], y=list(p75) + list(p25[::-1]),
                fill="toself", fillcolor="rgba(41,128,185,0.28)",
                line=dict(color="rgba(0,0,0,0)"), name="25th–75th Pct", showlegend=True,
            ))
            fig_fan.add_trace(go.Scatter(
                x=yr_ax, y=p50, name="Median", line=dict(color="#2980b9", width=2.5),
            ))
            fig_fan.add_hline(
                y=rev3, line_dash="dot", line_color="#95a5a6", line_width=1.5,
                annotation_text=f"Base SGD {rev3:,.0f}", annotation_position="right",
            )
            if ab_yr > 0:
                fig_fan.add_vline(
                    x=ab_yr, line_dash="dash", line_color="#e74c3c", line_width=1.5,
                    annotation_text=f"Option Yr {ab_yr}", annotation_position="top right",
                )
            fig_fan.update_layout(
                title="Simulated Revenue Paths  (Fan Chart)",
                xaxis_title="Year", yaxis_title="Revenue (SGD)",
                height=300, margin=dict(t=50, b=40, l=60, r=90),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
            )
            st.plotly_chart(fig_fan, use_container_width=True)

            st.markdown("---")

            # ── Sensitivity Tornado ────────────────────────────────────
            st.markdown("#### Sensitivity Tornado  —  ΔNPV vs Base Case")

            def _det_npv(g_m, m_m, r_d, n_yr):
                """Deterministic NPV given mean inputs (for tornado analysis)."""
                npv_d = float(-i_inv)
                rev_d = float(rev3)
                for t in range(1, max(1, n_yr) + 1):
                    rev_d = rev_d * (1.0 + g_m)
                    cf_d  = (rev_d * m_m * (1.0 - _tax_r)
                             + float(da3)
                             - float(_nwc_r) / 100.0 * max(0.0, rev_d * g_m)
                             - float(_capex_r) / 100.0 * rev_d)
                    npv_d += cf_d / (1.0 + r_d) ** t
                return npv_d

            g0 = prm["growth_mean"] / 100;  gs = prm["growth_std"] / 100
            m0_p = prm["margin_mean"] / 100; ms = prm["margin_std"] / 100
            base_det = _det_npv(g0, m0_p, r_e, h_yr)

            raw_scenarios = [
                (f"Revenue Growth +{prm['growth_std']:.0f}% (1σ)", _det_npv(g0 + gs, m0_p, r_e, h_yr)),
                (f"Revenue Growth −{prm['growth_std']:.0f}% (1σ)", _det_npv(g0 - gs, m0_p, r_e, h_yr)),
                (f"EBIT Margin +{prm['margin_std']:.0f}% (1σ)",    _det_npv(g0, m0_p + ms, r_e, h_yr)),
                (f"EBIT Margin −{prm['margin_std']:.0f}% (1σ)",    _det_npv(g0, m0_p - ms, r_e, h_yr)),
                ("Discount Rate +2%",                               _det_npv(g0, m0_p, r_e + 0.02, h_yr)),
                ("Discount Rate −2%",                               _det_npv(g0, m0_p, max(r_e - 0.02, 0.01), h_yr)),
                ("Horizon +3 years",                                _det_npv(g0, m0_p, r_e, h_yr + 3)),
            ]
            if h_yr > 3:
                raw_scenarios.append(("Horizon −3 years", _det_npv(g0, m0_p, r_e, h_yr - 3)))

            raw_scenarios.sort(key=lambda x: abs(x[1] - base_det), reverse=True)
            t_lbls   = [s[0] for s in raw_scenarios]
            t_deltas = [s[1] - base_det for s in raw_scenarios]
            t_cols   = ["#27ae60" if d >= 0 else "#e74c3c" for d in t_deltas]
            t_texts  = [f"SGD {s[1]:,.0f}" for s in raw_scenarios]

            fig_torn = go.Figure(go.Bar(
                y=t_lbls, x=t_deltas, orientation="h",
                marker_color=t_cols, text=t_texts, textposition="outside",
            ))
            fig_torn.add_vline(x=0, line_color="#555", line_width=1)
            fig_torn.update_layout(
                title=f"Base Case (deterministic) NPV: SGD {base_det:,.0f}",
                xaxis_title="ΔNPV vs Base Case (SGD)",
                height=370, margin=dict(t=50, b=40, l=225, r=110),
            )
            st.plotly_chart(fig_torn, use_container_width=True)

            st.markdown("---")

            # ── Full statistics table ──────────────────────────────────
            with st.expander("📊 Full Statistics Table", expanded=False):
                pq = [1, 5, 10, 25, 50, 75, 90, 95, 99]
                stat_rows = [
                    ("Mean NPV",           f"SGD {m0:,.0f}",                    f"SGD {m1:,.0f}"),
                    ("Median NPV",         f"SGD {med0:,.0f}",                  f"SGD {float(np.median(npv1)):,.0f}"),
                    ("Standard Deviation", f"SGD {sd0:,.0f}",                   f"SGD {sd1:,.0f}"),
                    ("Minimum",            f"SGD {float(npv0.min()):,.0f}",     f"SGD {float(npv1.min()):,.0f}"),
                    ("Maximum",            f"SGD {float(npv0.max()):,.0f}",     f"SGD {float(npv1.max()):,.0f}"),
                    ("P(NPV > 0)",         f"{pos0:.1f}%",                      f"{pos1:.1f}%"),
                    ("Real Option Value",  "—",                                 f"SGD {rov:,.0f}"),
                ] + [
                    (f"{q}th Percentile",
                     f"SGD {float(np.percentile(npv0, q)):,.0f}",
                     f"SGD {float(np.percentile(npv1, q)):,.0f}")
                    for q in pq
                ]
                st.dataframe(
                    pd.DataFrame(stat_rows, columns=["Statistic", "Without Option", "With Abandonment Option"]),
                    use_container_width=True, hide_index=True,
                )
