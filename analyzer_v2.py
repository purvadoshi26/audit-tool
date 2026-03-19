"""
analyzer_v2.py
==============
VERSION 2 — Smart Audit Analytical Procedures Engine
Fixes: real-world P&L parsing, unit auto-detection, subtotal row removal
New:   Going Concern (ISA 570), Benford's Law (ISA 240), IFRS/Ind AS dual standard

Author: Purva Doshi's Audit AP Tool v2.0
"""

import pandas as pd
import numpy as np
import math
from collections import Counter


# ══════════════════════════════════════════════════════════════
# SECTION 1 — STANDARDS CONFIG
# Toggle between IFRS and Ind AS — changes labels, references, flags
# ══════════════════════════════════════════════════════════════

STANDARDS = {
    "IFRS": {
        "cover_subtitle":   "IFRS — ISA 520 Analytical Procedures",
        "ratio_ref":        "IAS 1 / IFRS",
        "revenue_label":    "Revenue (IFRS 15)",
        "profit_label":     "Profit for the Year",
        "provision_ref":    "IAS 37",
        "impairment_ref":   "IAS 36",
        "going_concern_ref":"ISA 570",
        "fraud_ref":        "ISA 240",
        "lease_note":       "IFRS 16 — Lease liabilities should be separated from other borrowings.",
        "revenue_note":     "Apply IFRS 15 five-step model: identify contract, PO, price, allocate, recognise.",
        "std_label":        "IFRS",
        "doc_ref":          "IAS 1.85 / ISA 520.5",
    },
    "Ind AS": {
        "cover_subtitle":   "Ind AS — SA 520 Analytical Procedures",
        "ratio_ref":        "Ind AS 1 / Schedule III",
        "revenue_label":    "Revenue from Operations (Ind AS 115)",
        "profit_label":     "Profit / (Loss) for the Year",
        "provision_ref":    "Ind AS 37",
        "impairment_ref":   "Ind AS 36",
        "going_concern_ref":"SA 570",
        "fraud_ref":        "SA 240",
        "lease_note":       "Ind AS 116 — Lease liabilities to be recognised on balance sheet.",
        "revenue_note":     "Ind AS 115 five-step recognition model — verify contract modifications and variable consideration.",
        "std_label":        "Ind AS",
        "doc_ref":          "Ind AS 1.85 / SA 520.5",
    },
}


# ══════════════════════════════════════════════════════════════
# SECTION 2 — KEYWORD MAP (expanded for both IFRS and Ind AS)
# ══════════════════════════════════════════════════════════════

KEYWORD_MAP = {
    "current_assets": [
        "cash", "bank", "debtor", "trade receivable", "sundry receivable",
        "inventory", "stock", "prepaid", "advance paid", "tds receivable",
        "gst receivable", "input credit", "deposit", "short term investment",
        "current asset", "other current asset", "loan and advance",
    ],
    "fixed_assets": [
        "fixed asset", "plant", "machinery", "equipment", "vehicle",
        "furniture", "building", "land", "computer", "intangible",
        "capital wip", "cwip", "goodwill", "trademark", "patent",
        "right of use", "rou asset", "property plant", "ppe",
    ],
    "current_liabilities": [
        "creditor", "trade payable", "sundry creditor", "outstanding",
        "advance received", "tds payable", "gst payable", "tax payable",
        "provision", "short term loan", "bank overdraft", "od limit",
        "current liability", "other current liability",
    ],
    "long_term_liabilities": [
        "term loan", "long term loan", "debenture", "bond", "mortgage",
        "long term borrowing", "ncd", "long term liability", "lease liability",
        "borrowing", "long-term",
    ],
    "equity": [
        "share capital", "capital", "reserve", "surplus", "retained",
        "equity", "owner", "partner capital", "proprietor", "p&l appropriation",
        "other comprehensive income", "oci",
    ],
    "revenue": [
        "revenue from operation", "revenue from contract", "net revenue",
        "sales", "turnover", "service income", "domestic sales",
        "export sales", "other income", "income from operation",
    ],
    "cogs": [
        "purchase of", "purchase of stock", "purchase of consumable",
        "material cost", "direct cost", "cost of goods", "cost of revenue",
        "freight inward", "direct expense", "raw material",
        "changes in inventor", "change in inventor",  # handles "Changes in inventories"
    ],
    "interest_expense": [
        "finance cost", "interest expense", "bank charges",
        "interest on loan", "interest on od", "financial charge",
        "interest and finance",
    ],
    "expenses": [
        "employee benefit", "salaries", "wages", "rent", "power",
        "electricity", "repair", "maintenance", "advertising",
        "professional fee", "audit fee", "depreciation", "amortis",
        "insurance", "travelling", "printing", "postage", "telephone",
        "commission", "security", "cleaning", "freight outward",
        "other expense", "administrative", "selling",
    ],
}

# Rows to SKIP — subtotals, headers, blank labels
SKIP_KEYWORDS = [
    "total", "sub-total", "subtotal", "grand total", "net total",
    "loss before", "profit before", "loss after", "profit after",
    "profit for the year", "loss for the year",
    "comprehensive income", "comprehensive loss",
    "tax expense", "total income", "total expense",
    "earnings per", "eps", "basic", "diluted",
    "other comprehensive",
]


def should_skip_row(account_name: str) -> bool:
    """Return True if this row is a subtotal/header/computed row to exclude."""
    name = str(account_name).lower().strip()
    if not name or name == "nan":
        return True
    for kw in SKIP_KEYWORDS:
        if kw in name:
            return True
    return False


def auto_classify(account_name: str) -> str:
    """Classify account from name. Returns category string or 'unclassified'."""
    name_lower = str(account_name).lower().strip()
    priority = [
        "interest_expense", "cogs", "current_assets", "fixed_assets",
        "current_liabilities", "long_term_liabilities", "equity", "revenue", "expenses",
    ]
    for cat in priority:
        for kw in KEYWORD_MAP[cat]:
            if kw in name_lower:
                return cat
    return "unclassified"


# ══════════════════════════════════════════════════════════════
# SECTION 3 — SMART LOADER
# Handles: real P&L format, "-" values, Notes col, unit detection
# ══════════════════════════════════════════════════════════════

def detect_unit_scale(df: pd.DataFrame, cy_col: str, py_col: str) -> dict:
    """
    Detect if amounts are in units, lakhs, or crores.
    Returns a dict with scale info and suggested materiality threshold.

    Logic:
    - If max value > 10,000,000  → units (₹) → threshold ₹5,00,000
    - If max value > 10,000      → likely lakhs (₹ L) → threshold 5 (₹5L)
    - If max value < 10,000      → likely crores (₹ Cr) → threshold 0.5 (₹50L)
    """
    all_vals = pd.concat([df[cy_col], df[py_col]]).dropna().abs()
    max_val = all_vals.max() if len(all_vals) > 0 else 0

    if max_val > 10_000_000:
        return {"label": "₹ (Units)", "threshold": 500_000, "scale": "units"}
    elif max_val > 10_000:
        return {"label": "₹ Lakhs", "threshold": 5.0, "scale": "lakhs"}
    elif max_val > 100:
        return {"label": "₹ Crores", "threshold": 0.50, "scale": "crores"}
    else:
        return {"label": "₹ Crores", "threshold": 0.25, "scale": "crores"}


def clean_amount(val) -> float:
    """Convert any amount value to float. Handles '-', None, text."""
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return 0.0
    s = str(val).strip()
    if s in ["-", "–", "—", "", "nan", "None", "N/A", "NA"]:
        return 0.0
    # Remove ₹, commas, spaces
    s = s.replace("₹", "").replace(",", "").replace(" ", "")
    # Handle parentheses as negative: (123) → -123
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except ValueError:
        return 0.0


def load_trial_balance(filepath: str) -> tuple[pd.DataFrame, dict]:
    """
    Load trial balance or P&L statement — handles both formats.

    Accepts:
    - Format A: Account Name | Category | CY Amount (₹) | PY Amount (₹)
    - Format B: Account Name | Notes | CY Amount (₹) | PY Amount (₹)  (real P&L)
    - Format C: Account Name | CY Amount (₹) | PY Amount (₹)           (no category)

    Returns:
        (DataFrame, unit_info_dict)
    """
    try:
        if filepath.endswith(".csv"):
            df = pd.read_csv(filepath)
        else:
            df = pd.read_excel(filepath)
    except Exception as e:
        raise ValueError(f"Cannot read file: {e}")

    df.columns = [str(c).strip() for c in df.columns]

    # ── Find CY and PY columns ──
    cy_col = py_col = None
    for col in df.columns:
        cl = col.lower()
        if "cy" in cl or "current" in cl:
            cy_col = col
        if "py" in cl or "prior" in cl or "previous" in cl:
            py_col = col

    if cy_col is None or py_col is None:
        # Try positional: last two numeric columns
        num_cols = [c for c in df.columns if pd.to_numeric(
            df[c].astype(str).str.replace(r'[₹,\-\(\) ]', '', regex=True), errors='coerce'
        ).notna().sum() > len(df) * 0.4]
        if len(num_cols) >= 2:
            cy_col, py_col = num_cols[-2], num_cols[-1]
        else:
            raise ValueError(
                "Cannot find CY and PY amount columns. "
                "Please name them 'CY Amount (₹)' and 'PY Amount (₹)'."
            )

    # ── Find Account Name column ──
    name_col = None
    for col in df.columns:
        if "account" in col.lower() or "name" in col.lower() or "description" in col.lower() or "particular" in col.lower():
            name_col = col
            break
    if name_col is None:
        name_col = df.columns[0]  # First column

    # ── Build clean DataFrame ──
    df_clean = pd.DataFrame()
    df_clean["Account Name"] = df[name_col].astype(str).str.strip()

    # ── Clean amounts ──
    df_clean["CY Amount (₹)"] = df[cy_col].apply(clean_amount)
    df_clean["PY Amount (₹)"] = df[py_col].apply(clean_amount)

    # ── Drop subtotal / header / blank rows ──
    df_clean = df_clean[~df_clean["Account Name"].apply(should_skip_row)].copy()
    df_clean = df_clean.reset_index(drop=True)

    # ── Handle Category column ──
    cat_col = None
    for col in df.columns:
        cl = col.lower()
        if "categor" in cl or "type" in cl or "head" in cl:
            cat_col = col
            break

    if cat_col and cat_col in df.columns:
        # Align index after row drops
        df_with_cat = df[[name_col, cat_col, cy_col, py_col]].copy()
        df_with_cat.columns = ["Account Name", "Category", "CY", "PY"]
        df_with_cat["Account Name"] = df_with_cat["Account Name"].astype(str).str.strip()
        df_with_cat = df_with_cat[~df_with_cat["Account Name"].apply(should_skip_row)]
        df_clean["Category"] = df_with_cat["Category"].values[:len(df_clean)]
    else:
        df_clean["Category"] = ""

    # ── Auto-classify blanks ──
    mask = df_clean["Category"].isna() | (df_clean["Category"].astype(str).str.strip() == "") | (df_clean["Category"].astype(str).str.lower() == "nan")
    df_clean.loc[mask, "Category"] = df_clean.loc[mask, "Account Name"].apply(auto_classify)
    df_clean["Category"] = df_clean["Category"].astype(str).str.lower().str.strip()

    # ── Detect unit scale ──
    unit_info = detect_unit_scale(df_clean, "CY Amount (₹)", "PY Amount (₹)")

    return df_clean, unit_info


# ══════════════════════════════════════════════════════════════
# SECTION 4 — FINANCIAL AGGREGATION
# ══════════════════════════════════════════════════════════════

def aggregate_financials(df: pd.DataFrame) -> dict:
    """Sum amounts by category and derive key financial line items."""
    categories = [
        "current_assets", "fixed_assets",
        "current_liabilities", "long_term_liabilities",
        "equity", "revenue", "cogs",
        "interest_expense", "expenses",
    ]
    result = {}
    for cat in categories:
        sub = df[df["Category"] == cat]
        result[cat] = {
            "cy": sub["CY Amount (₹)"].sum(),
            "py": sub["PY Amount (₹)"].sum(),
        }

    # Derived totals
    result["total_assets"] = {
        "cy": result["current_assets"]["cy"] + result["fixed_assets"]["cy"],
        "py": result["current_assets"]["py"] + result["fixed_assets"]["py"],
    }
    result["total_liabilities"] = {
        "cy": result["current_liabilities"]["cy"] + result["long_term_liabilities"]["cy"],
        "py": result["current_liabilities"]["py"] + result["long_term_liabilities"]["py"],
    }
    result["gross_profit"] = {
        "cy": result["revenue"]["cy"] - result["cogs"]["cy"],
        "py": result["revenue"]["py"] - result["cogs"]["py"],
    }
    result["ebit"] = {
        "cy": result["gross_profit"]["cy"] - result["expenses"]["cy"],
        "py": result["gross_profit"]["py"] - result["expenses"]["py"],
    }
    result["net_profit"] = {
        "cy": result["ebit"]["cy"] - result["interest_expense"]["cy"],
        "py": result["ebit"]["py"] - result["interest_expense"]["py"],
    }
    result["working_capital"] = {
        "cy": result["current_assets"]["cy"] - result["current_liabilities"]["cy"],
        "py": result["current_assets"]["py"] - result["current_liabilities"]["py"],
    }
    return result


# ══════════════════════════════════════════════════════════════
# SECTION 5 — RATIO CALCULATION (standard-aware)
# ══════════════════════════════════════════════════════════════

def safe_div(n, d, default=None):
    if d == 0 or d is None or (isinstance(d, float) and math.isnan(d)):
        return default
    return n / d


def calculate_ratios(agg: dict, standard: str = "Ind AS") -> list:
    """Calculate ratios. standard = 'IFRS' or 'Ind AS'"""
    cfg = STANDARDS[standard]
    ratios = []

    def add(name, formula, cy_val, py_val, flag_fn, note):
        cy = round(cy_val, 2) if cy_val is not None else None
        py = round(py_val, 2) if py_val is not None else None
        chg = round(((cy - py) / abs(py)) * 100, 1) if (cy is not None and py and py != 0) else None
        try:
            flag = flag_fn(cy, py, chg)
        except Exception:
            flag = "⚪ N/A"
        ratios.append({
            "Ratio / Metric": name, "Formula": formula,
            "Current Year": cy, "Prior Year": py,
            "YoY Change (%)": chg, "Risk Flag": flag,
            "Auditor's Note": note,
        })

    rev_cy  = agg["revenue"]["cy"];        rev_py  = agg["revenue"]["py"]
    gp_cy   = agg["gross_profit"]["cy"];   gp_py   = agg["gross_profit"]["py"]
    np_cy   = agg["net_profit"]["cy"];     np_py   = agg["net_profit"]["py"]
    ebit_cy = agg["ebit"]["cy"];           ebit_py = agg["ebit"]["py"]
    ca_cy   = agg["current_assets"]["cy"]; ca_py   = agg["current_assets"]["py"]
    cl_cy   = agg["current_liabilities"]["cy"]; cl_py = agg["current_liabilities"]["py"]
    fa_cy   = agg["fixed_assets"]["cy"];   fa_py   = agg["fixed_assets"]["py"]
    ta_cy   = agg["total_assets"]["cy"];   ta_py   = agg["total_assets"]["py"]
    eq_cy   = agg["equity"]["cy"];         eq_py   = agg["equity"]["py"]
    lt_cy   = agg["long_term_liabilities"]["cy"]
    lt_py   = agg["long_term_liabilities"]["py"]
    int_cy  = agg["interest_expense"]["cy"]; int_py = agg["interest_expense"]["py"]
    cogs_cy = agg["cogs"]["cy"];           cogs_py = agg["cogs"]["py"]
    opex_cy = agg["expenses"]["cy"]

    # PROFITABILITY
    add("Gross Profit Margin (%)", f"Gross Profit / {cfg['revenue_label']} × 100",
        safe_div(gp_cy, rev_cy, 0)*100, safe_div(gp_py, rev_py, 0)*100,
        lambda cy, py, chg: "🔴 HIGH — Margin fell >5pp" if (cy and py and cy-py < -5) else
                             "🟡 MONITOR — Margin fell 2–5pp" if (cy and py and cy-py < -2) else "🟢 OK",
        f"Margin decline → rising COGS, pricing pressure, or revenue understatement. Ref: {cfg['revenue_note']}")

    add("Net Profit Margin (%)", f"{cfg['profit_label']} / Revenue × 100",
        safe_div(np_cy, rev_cy, 0)*100, safe_div(np_py, rev_py, 0)*100,
        lambda cy, py, chg: "🔴 HIGH — Loss-making entity" if (cy is not None and cy < 0) else
                             "🔴 HIGH — Margin fell >5pp" if (cy and py and cy-py < -5) else
                             "🟡 MONITOR — Decline >3pp" if (cy and py and cy-py < -3) else "🟢 OK",
        "Sustained losses → going concern flag. Check interest burden and unusual provisions.")

    add("Revenue Growth (%)", "(CY Revenue − PY Revenue) / PY Revenue × 100",
        rev_cy, rev_py,
        lambda cy, py, chg: "🔴 HIGH — Revenue fell >20%" if (chg and chg < -20) else
                             "🟡 MONITOR — Unusual growth/fall >20%" if (chg and abs(chg) > 20) else "🟢 OK",
        f"Investigate cut-off, returns, credit notes. Ref: {cfg['revenue_note']}")

    add("COGS % of Revenue", "COGS / Revenue × 100",
        safe_div(cogs_cy, rev_cy, 0)*100, safe_div(cogs_py, rev_py, 0)*100,
        lambda cy, py, chg: "🔴 HIGH — COGS ratio jumped >10pp" if (cy and py and cy-py > 10) else
                             "🟡 MONITOR — COGS ratio changed >5pp" if (cy and py and abs(cy-py) > 5) else "🟢 OK",
        "Rising COGS% → cost inflation, waste, incorrect inventory valuation, or fictitious purchases.")

    add("Operating Expense Ratio (%)", "Operating Expenses / Revenue × 100",
        safe_div(opex_cy, rev_cy, 0)*100 if rev_cy > 0 else None,
        safe_div(agg["expenses"]["py"], rev_py, 0)*100 if rev_py > 0 else None,
        lambda cy, py, chg: "🔴 HIGH — Expense jumped >5pp of revenue" if (cy and py and cy-py > 5) else
                             "🟡 MONITOR — Increase 2–5pp" if (cy and py and cy-py > 2) else "🟢 OK",
        "Identify specific cost lines driving increase. Check for fictitious expenses or cut-off errors.")

    # LIQUIDITY
    add("Current Ratio", "Current Assets / Current Liabilities",
        safe_div(ca_cy, cl_cy), safe_div(ca_py, cl_py),
        lambda cy, py, chg: "🔴 HIGH — Below 1.0 (insolvency risk)" if (cy and cy < 1.0) else
                             "🟡 MONITOR — Below 1.5" if (cy and cy < 1.5) else "🟢 OK",
        f"Ratio <1 → {cfg['going_concern_ref']} going concern trigger. Verify receivable recoverability.")

    add("Working Capital", "Current Assets − Current Liabilities",
        agg["working_capital"]["cy"], agg["working_capital"]["py"],
        lambda cy, py, chg: "🔴 HIGH — Negative working capital" if (cy is not None and cy < 0) else
                             "🟡 MONITOR — Fell >25%" if (chg and chg < -25) else "🟢 OK",
        "Negative WC → cannot meet short-term obligations. Check overdraft facilities and creditor terms.")

    # LEVERAGE
    total_debt_cy = agg["total_liabilities"]["cy"]
    total_debt_py = agg["total_liabilities"]["py"]
    add("Debt-to-Equity Ratio", "Total Liabilities / Equity",
        safe_div(total_debt_cy, eq_cy), safe_div(total_debt_py, eq_py),
        lambda cy, py, chg: "🔴 HIGH — D/E >3 (high leverage)" if (cy and cy > 3) else
                             "🟡 MONITOR — D/E >2" if (cy and cy > 2) else "🟢 OK",
        f"High leverage → covenant breach risk. Review loan agreements. Ref: {cfg['lease_note']}")

    add("Interest Coverage Ratio (EBIT / Interest)", "EBIT / Interest Expense",
        safe_div(ebit_cy, int_cy), safe_div(ebit_py, int_py),
        lambda cy, py, chg: "🔴 HIGH — Cannot cover interest (ICR <1)" if (cy is not None and cy < 1) else
                             "🟡 MONITOR — Low coverage 1–2x" if (cy is not None and cy < 2) else "🟢 OK",
        f"ICR <1 → {cfg['going_concern_ref']} trigger. Entity cannot service debt from operations.")

    # EFFICIENCY
    add("Asset Turnover", f"Revenue / Total Assets",
        safe_div(rev_cy, ta_cy), safe_div(rev_py, ta_py),
        lambda cy, py, chg: "🟡 MONITOR — Turnover fell >20%" if (chg and chg < -20) else "🟢 OK",
        f"Declining turnover → underutilised assets or revenue shortfall. Verify both. {cfg['impairment_ref']} impairment test if AT drops sharply.")

    add("Return on Equity (%)", f"{cfg['profit_label']} / Equity × 100",
        safe_div(np_cy, eq_cy, 0)*100 if eq_cy else None,
        safe_div(np_py, eq_py, 0)*100 if eq_py else None,
        lambda cy, py, chg: "🔴 HIGH — Negative ROE (loss-making)" if (cy is not None and cy < 0) else
                             "🟡 MONITOR — ROE fell >10pp" if (cy and py and cy-py < -10) else "🟢 OK",
        "Negative ROE signals equity erosion. Verify dividend payouts and reserves movement.")

    add("Debtors Days (est.)", "Est. Debtors / Revenue × 365",
        safe_div(ca_cy * 0.45, rev_cy, 0)*365 if rev_cy else None,
        safe_div(ca_py * 0.45, rev_py, 0)*365 if rev_py else None,
        lambda cy, py, chg: "🔴 HIGH — Debtor days jumped >30 days" if (cy and py and cy-py > 30) else
                             "🟡 MONITOR — Increase 15–30 days" if (cy and py and cy-py > 15) else "🟢 OK",
        "Increasing debtor days → bad debt risk, fictitious debtors, or aggressive revenue recognition.")

    add("Creditors Days (est.)", "Est. Creditors / COGS × 365",
        safe_div(cl_cy * 0.55, cogs_cy, 0)*365 if cogs_cy else None,
        safe_div(cl_py * 0.55, cogs_py, 0)*365 if cogs_py else None,
        lambda cy, py, chg: "🟡 MONITOR — Creditor days jumped >30" if (cy and py and cy-py > 30) else "🟢 OK",
        "Rising creditor days → cash flow stress or disputed payables. Obtain creditor confirmations.")

    add("Equity to Total Assets (%)", "Equity / Total Assets × 100",
        safe_div(eq_cy, ta_cy, 0)*100 if ta_cy else None,
        safe_div(eq_py, ta_py, 0)*100 if ta_py else None,
        lambda cy, py, chg: "🔴 HIGH — Equity <20% of assets" if (cy is not None and 0 < cy < 20) else
                             "🔴 HIGH — Negative equity (technical insolvency)" if (cy is not None and cy < 0) else
                             "🟢 OK",
        "Negative/eroded equity → technical insolvency. Verify capital adequacy and regulatory compliance.")

    return ratios


# ══════════════════════════════════════════════════════════════
# SECTION 6 — LINE ITEM VARIANCE
# ══════════════════════════════════════════════════════════════

def line_item_variance(df: pd.DataFrame, threshold_pct: float = 20.0,
                       threshold_abs: float = 50000.0) -> pd.DataFrame:
    df = df.copy()
    df["Change (₹)"] = df["CY Amount (₹)"] - df["PY Amount (₹)"]

    def calc_pct(row):
        py = row["PY Amount (₹)"]
        cy = row["CY Amount (₹)"]
        if py == 0:
            return None if cy == 0 else 999.0
        return round(((cy - py) / abs(py)) * 100, 1)

    df["Change (%)"] = df.apply(calc_pct, axis=1)

    def flag(row):
        chg_pct = row["Change (%)"]
        chg_abs = abs(row["Change (₹)"])
        if chg_pct is None:
            return "⚪ New Item"
        if chg_pct == 999.0:
            return "🟡 New Account — Verify"
        if chg_abs < threshold_abs:
            return "⚪ Immaterial"
        if abs(chg_pct) >= threshold_pct * 2:
            return "🔴 HIGH — Investigate"
        if abs(chg_pct) >= threshold_pct:
            return "🟡 MODERATE — Explain"
        return "🟢 Within Threshold"

    df["Variance Flag"] = df.apply(flag, axis=1)

    priority = {
        "🔴 HIGH — Investigate": 0, "🟡 MODERATE — Explain": 1,
        "🟡 New Account — Verify": 2, "⚪ New Item": 3,
        "🟢 Within Threshold": 4, "⚪ Immaterial": 5,
    }
    df["_sort"] = df["Variance Flag"].map(priority).fillna(99)
    df = df.sort_values("_sort").drop(columns=["_sort"])
    return df


# ══════════════════════════════════════════════════════════════
# SECTION 7 — GOING CONCERN ASSESSMENT (ISA 570 / SA 570)
# ══════════════════════════════════════════════════════════════

def going_concern_assessment(agg: dict, ratios: list, standard: str = "Ind AS") -> dict:
    """
    ISA 570 / SA 570 going concern assessment.
    Scores 10 indicators and returns an overall risk level.

    Returns a dict with:
      - indicators: list of {label, status, finding, reference}
      - overall_risk: 'LOW' | 'MODERATE' | 'HIGH' | 'CRITICAL'
      - score: int (0–10, higher = more concern)
      - conclusion: str
    """
    cfg = STANDARDS[standard]
    indicators = []
    score = 0

    def ratio_val(name):
        for r in ratios:
            if r["Ratio / Metric"] == name:
                return r["Current Year"]
        return None

    # Convenience
    np_cy   = agg["net_profit"]["cy"]
    np_py   = agg["net_profit"]["py"]
    wc      = agg["working_capital"]["cy"]
    rev_cy  = agg["revenue"]["cy"]
    rev_py  = agg["revenue"]["py"]
    eq_cy   = agg["equity"]["cy"]
    int_cy  = agg["interest_expense"]["cy"]
    ebit_cy = agg["ebit"]["cy"]
    lt_cy   = agg["long_term_liabilities"]["cy"]
    ca_cy   = agg["current_assets"]["cy"]
    cl_cy   = agg["current_liabilities"]["cy"]

    # ── Indicator 1: Net loss in current year ──
    if np_cy < 0:
        score += 2
        indicators.append({
            "Indicator": "Net loss in current year",
            "Status": "🔴 CONCERN",
            "Finding": f"Entity reported a loss of {abs(np_cy):,.2f}. Losses reduce equity and cash reserves.",
            "Reference": f"{cfg['going_concern_ref']} para A2(a)",
        })
    elif np_cy < np_py * 0.5 and np_py > 0:
        score += 1
        indicators.append({
            "Indicator": "Profit declined >50% vs prior year",
            "Status": "🟡 MONITOR",
            "Finding": f"Profit fell sharply from {np_py:,.2f} to {np_cy:,.2f}.",
            "Reference": f"{cfg['going_concern_ref']} para A2(a)",
        })
    else:
        indicators.append({
            "Indicator": "Net loss in current year",
            "Status": "🟢 CLEAR",
            "Finding": "Entity is profitable.",
            "Reference": f"{cfg['going_concern_ref']} para A2(a)",
        })

    # ── Indicator 2: Current ratio < 1 ──
    cr = safe_div(ca_cy, cl_cy)
    if cr is not None and cr < 1.0:
        score += 2
        indicators.append({
            "Indicator": "Current Ratio < 1 (liquidity failure)",
            "Status": "🔴 CONCERN",
            "Finding": f"Current ratio = {cr:.2f}. Cannot meet short-term obligations from current assets.",
            "Reference": f"{cfg['going_concern_ref']} para A2(b)",
        })
    elif cr is not None and cr < 1.5:
        score += 1
        indicators.append({
            "Indicator": "Current Ratio < 1.5 (low liquidity)",
            "Status": "🟡 MONITOR",
            "Finding": f"Current ratio = {cr:.2f}. Limited liquidity buffer.",
            "Reference": f"{cfg['going_concern_ref']} para A2(b)",
        })
    else:
        indicators.append({
            "Indicator": "Current Ratio < 1 (liquidity failure)",
            "Status": "🟢 CLEAR",
            "Finding": f"Current ratio = {f'{cr:.2f}' if cr is not None else 'N/A'}. Adequate liquidity.",
            "Reference": f"{cfg['going_concern_ref']} para A2(b)",
        })

    # ── Indicator 3: Negative working capital ──
    if wc < 0:
        score += 1
        indicators.append({
            "Indicator": "Negative working capital",
            "Status": "🔴 CONCERN",
            "Finding": f"Working capital = {wc:,.2f}. Net current liabilities indicate funding gap.",
            "Reference": f"{cfg['going_concern_ref']} para A2(b)",
        })
    else:
        indicators.append({
            "Indicator": "Negative working capital",
            "Status": "🟢 CLEAR",
            "Finding": f"Working capital = {wc:,.2f}. Positive.",
            "Reference": f"{cfg['going_concern_ref']} para A2(b)",
        })

    # ── Indicator 4: Interest coverage < 1 ──
    icr = safe_div(ebit_cy, int_cy)
    if int_cy > 0:
        if icr is not None and icr < 1.0:
            score += 2
            indicators.append({
                "Indicator": "Interest coverage < 1 (cannot service debt)",
                "Status": "🔴 CONCERN",
                "Finding": f"ICR = {icr:.2f}. EBIT insufficient to cover finance costs — debt default risk.",
                "Reference": f"{cfg['going_concern_ref']} para A2(c)",
            })
        elif icr is not None and icr < 2.0:
            score += 1
            indicators.append({
                "Indicator": "Interest coverage < 2 (low debt service)",
                "Status": "🟡 MONITOR",
                "Finding": f"ICR = {icr:.2f}. Low debt service capacity.",
                "Reference": f"{cfg['going_concern_ref']} para A2(c)",
            })
        else:
            indicators.append({
                "Indicator": "Interest coverage < 1 (cannot service debt)",
                "Status": "🟢 CLEAR",
                "Finding": f"ICR = {f'{icr:.2f}' if icr is not None else 'N/A'}. Adequate interest coverage.",
                "Reference": f"{cfg['going_concern_ref']} para A2(c)",
            })
    else:
        indicators.append({
            "Indicator": "Interest coverage < 1 (cannot service debt)",
            "Status": "⚪ N/A",
            "Finding": "No interest expense recorded — not applicable.",
            "Reference": f"{cfg['going_concern_ref']} para A2(c)",
        })

    # ── Indicator 5: Revenue declining ──
    if rev_py > 0:
        rev_chg = (rev_cy - rev_py) / abs(rev_py) * 100
        if rev_chg < -20:
            score += 1
            indicators.append({
                "Indicator": "Significant revenue decline (>20%)",
                "Status": "🔴 CONCERN",
                "Finding": f"Revenue fell {abs(rev_chg):.1f}% from {rev_py:,.2f} to {rev_cy:,.2f}.",
                "Reference": f"{cfg['going_concern_ref']} para A2(d)",
            })
        else:
            indicators.append({
                "Indicator": "Significant revenue decline (>20%)",
                "Status": "🟢 CLEAR",
                "Finding": f"Revenue change: {rev_chg:+.1f}%. Within acceptable range.",
                "Reference": f"{cfg['going_concern_ref']} para A2(d)",
            })
    else:
        indicators.append({
            "Indicator": "Significant revenue decline (>20%)",
            "Status": "⚪ N/A",
            "Finding": "No PY revenue to compare.",
            "Reference": f"{cfg['going_concern_ref']} para A2(d)",
        })

    # ── Indicator 6: Equity erosion ──
    if eq_cy < 0:
        score += 2
        indicators.append({
            "Indicator": "Negative equity (technical insolvency)",
            "Status": "🔴 CONCERN",
            "Finding": f"Equity = {eq_cy:,.2f}. Liabilities exceed assets — technically insolvent.",
            "Reference": f"{cfg['going_concern_ref']} para A2(e)",
        })
    elif eq_cy == 0:
        indicators.append({
            "Indicator": "Negative equity (technical insolvency)",
            "Status": "⚪ N/A",
            "Finding": "No equity data available (possibly P&L-only input).",
            "Reference": f"{cfg['going_concern_ref']} para A2(e)",
        })
    else:
        indicators.append({
            "Indicator": "Negative equity (technical insolvency)",
            "Status": "🟢 CLEAR",
            "Finding": f"Equity = {eq_cy:,.2f}. Positive.",
            "Reference": f"{cfg['going_concern_ref']} para A2(e)",
        })

    # ── Overall risk determination ──
    if score >= 6:
        overall_risk = "CRITICAL"
        conclusion = (
            f"CRITICAL GOING CONCERN RISK — Multiple financial distress indicators present. "
            f"Under {cfg['going_concern_ref']}, the auditor must obtain sufficient appropriate evidence "
            f"that management's use of the going concern basis is appropriate. "
            f"Consider whether a 'material uncertainty' paragraph is required in the audit report."
        )
    elif score >= 4:
        overall_risk = "HIGH"
        conclusion = (
            f"HIGH GOING CONCERN RISK — Significant concerns exist. "
            f"Under {cfg['going_concern_ref']}, perform additional procedures: review cash flow forecasts, "
            f"assess management's plans to address the risks, review post-balance sheet events."
        )
    elif score >= 2:
        overall_risk = "MODERATE"
        conclusion = (
            f"MODERATE GOING CONCERN RISK — Some indicators present. "
            f"Under {cfg['going_concern_ref']}, obtain management's written assessment of going concern "
            f"and review supporting evidence for the 12-month outlook."
        )
    else:
        overall_risk = "LOW"
        conclusion = (
            f"LOW GOING CONCERN RISK — No significant indicators identified. "
            f"Document that going concern has been considered under {cfg['going_concern_ref']} "
            f"and no material uncertainty exists."
        )

    return {
        "indicators": indicators,
        "overall_risk": overall_risk,
        "score": score,
        "conclusion": conclusion,
    }


# ══════════════════════════════════════════════════════════════
# SECTION 8 — BENFORD'S LAW ANALYSIS (ISA 240)
# ══════════════════════════════════════════════════════════════

def benfords_law_analysis(amounts: list, label: str = "Transactions") -> dict:
    """
    Benford's Law first-digit analysis for fraud detection.

    Benford's Law: In naturally occurring financial data, the first digit
    is NOT uniformly distributed. Digit 1 appears ~30% of the time,
    digit 9 only ~4.6%. Deviations from this pattern can signal manipulation.

    Reference: ISA 240 — Auditor's responsibilities re fraud.

    Parameters:
        amounts : list of numeric transaction amounts
        label   : name of the dataset (for display)

    Returns:
        dict with observed %, expected %, chi-square stat, risk flag, findings
    """
    # Expected Benford distribution (%)
    benford_expected = {
        d: round(math.log10(1 + 1/d) * 100, 2)
        for d in range(1, 10)
    }

    # Extract first digits from all non-zero positive amounts
    first_digits = []
    for amt in amounts:
        try:
            v = abs(float(amt))
            if v <= 0:
                continue
            # Get first digit: convert to string, strip "0."
            s = f"{v:.10f}".lstrip("0").replace(".", "")
            if s:
                d = int(s[0])
                if 1 <= d <= 9:
                    first_digits.append(d)
        except (ValueError, TypeError):
            continue

    if len(first_digits) < 20:
        return {
            "label": label,
            "n": len(first_digits),
            "sufficient": False,
            "message": f"Only {len(first_digits)} usable amounts. Need at least 20 for Benford's analysis.",
            "risk_flag": "⚪ Insufficient Data",
        }

    n = len(first_digits)
    counts = Counter(first_digits)

    # Build result per digit
    digit_data = []
    chi_sq = 0.0

    for d in range(1, 10):
        observed_count = counts.get(d, 0)
        observed_pct   = round(observed_count / n * 100, 2)
        expected_pct   = benford_expected[d]
        expected_count = expected_pct / 100 * n
        deviation      = round(observed_pct - expected_pct, 2)
        deviation_pct  = round(abs(deviation / expected_pct) * 100, 1) if expected_pct else 0

        # Chi-square contribution for this digit
        chi_sq += ((observed_count - expected_count) ** 2) / expected_count if expected_count > 0 else 0

        digit_data.append({
            "Digit": d,
            "Observed Count": observed_count,
            "Observed (%)": observed_pct,
            "Expected (%)": expected_pct,
            "Deviation (pp)": deviation,
            "Deviation (%)": deviation_pct,
        })

    chi_sq = round(chi_sq, 2)

    # Critical values for chi-square with 8 df:
    # p=0.10 → 13.36, p=0.05 → 15.51, p=0.01 → 20.09
    if chi_sq > 20.09:
        risk_flag = "🔴 HIGH — Statistically significant deviation (p<0.01)"
        interpretation = (
            f"Chi-square = {chi_sq:.2f} exceeds the critical value of 20.09 (p<0.01). "
            "The digit distribution significantly deviates from Benford's Law. "
            "This is a strong indicator of data manipulation, number rounding, or fabrication. "
            "Perform targeted substantive procedures on amounts with overrepresented digits."
        )
    elif chi_sq > 15.51:
        risk_flag = "🟡 MODERATE — Possible anomaly (p<0.05)"
        interpretation = (
            f"Chi-square = {chi_sq:.2f} exceeds 15.51 (p<0.05). "
            "Moderate deviation from Benford's Law — possible manipulation or data quality issue. "
            "Review transactions starting with overrepresented digits."
        )
    elif chi_sq > 13.36:
        risk_flag = "🟡 LOW-MODERATE — Slight anomaly (p<0.10)"
        interpretation = (
            f"Chi-square = {chi_sq:.2f} exceeds 13.36 (p<0.10). "
            "Slight statistical deviation — note and consider additional sampling."
        )
    else:
        risk_flag = "🟢 PASS — Consistent with Benford's Law"
        interpretation = (
            f"Chi-square = {chi_sq:.2f} is within expected range. "
            "No statistically significant deviation from Benford's Law. "
            "This provides some comfort over the completeness and accuracy of the dataset."
        )

    return {
        "label": label,
        "n": n,
        "sufficient": True,
        "digit_data": digit_data,
        "chi_square": chi_sq,
        "risk_flag": risk_flag,
        "interpretation": interpretation,
        "benford_expected": benford_expected,
        "message": "",
    }


# ══════════════════════════════════════════════════════════════
# SECTION 9 — SUMMARY
# ══════════════════════════════════════════════════════════════

def generate_summary(agg: dict, df: pd.DataFrame, gc: dict, standard: str = "Ind AS") -> dict:
    flagged_high = len(df[df.get("Variance Flag", pd.Series(dtype=str)) == "🔴 HIGH — Investigate"]) if "Variance Flag" in df.columns else 0
    flagged_mod  = len(df[df.get("Variance Flag", pd.Series(dtype=str)) == "🟡 MODERATE — Explain"]) if "Variance Flag" in df.columns else 0
    unclassified = len(df[df["Category"] == "unclassified"])

    return {
        "Standard":                     STANDARDS[standard]["std_label"],
        "Total Accounts in TB":         len(df),
        "Flagged — HIGH Risk":          flagged_high,
        "Flagged — MODERATE":           flagged_mod,
        "Unclassified Accounts":        unclassified,
        "Going Concern Risk":           gc["overall_risk"],
        "CY Revenue":                   agg["revenue"]["cy"],
        "PY Revenue":                   agg["revenue"]["py"],
        "CY Net Profit":                agg["net_profit"]["cy"],
        "PY Net Profit":                agg["net_profit"]["py"],
        "CY Total Assets":              agg["total_assets"]["cy"],
        "PY Total Assets":              agg["total_assets"]["py"],
        "CY Working Capital":           agg["working_capital"]["cy"],
        "CY Gross Profit":              agg["gross_profit"]["cy"],
    }
