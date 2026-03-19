"""
app_v2.py
=========
VERSION 2 — Streamlit UI for Audit AP Tool
Improvements: standard toggle, Benford's tab, st.cache_data for speed, better UX

Run: streamlit run app_v2.py
"""

import streamlit as st
import pandas as pd
import os, tempfile, math

# ── Page config — must be FIRST ──────────────────────────────
st.set_page_config(
    page_title="Audit AP Tool v2 — Purva Doshi",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Lazy imports (speeds up startup) ─────────────────────────
@st.cache_data(show_spinner=False)
def _import_modules():
    from analyzer_v2 import (
        load_trial_balance, aggregate_financials, calculate_ratios,
        line_item_variance, going_concern_assessment,
        benfords_law_analysis, generate_summary,
    )
    from reporter_v2 import generate_report
    return {
        "load_trial_balance": load_trial_balance,
        "aggregate_financials": aggregate_financials,
        "calculate_ratios": calculate_ratios,
        "line_item_variance": line_item_variance,
        "going_concern_assessment": going_concern_assessment,
        "benfords_law_analysis": benfords_law_analysis,
        "generate_summary": generate_summary,
        "generate_report": generate_report,
    }

# ── CSS ───────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #F8FAFC; }
    [data-testid="stSidebar"] { background-color: #0F2B4C; }
    [data-testid="stSidebar"] * { color: #E2E8F0 !important; }
    [data-testid="stSidebar"] .stSelectbox label { color: #94A3B8 !important; }
    [data-testid="metric-container"] {
        background: white; border: 1px solid #E2E8F0;
        border-radius: 8px; padding: 12px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    h1, h2, h3 { font-family: Calibri, sans-serif !important; }
    .stDownloadButton > button {
        background-color: #0F2B4C !important;
        color: white !important; border-radius: 6px;
        font-weight: 700; padding: 10px 28px;
    }
    .stAlert { border-radius: 8px; }
    .risk-badge { font-size: 14px; font-weight: 700; padding: 4px 12px; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)


# ── Sidebar ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📋 Audit AP Tool")
    st.markdown("**Version 2.0**")
    st.markdown("---")

    st.markdown("### 🏢 Engagement")
    client_name = st.text_input("Client Name", value="ABC Pvt Ltd")
    period      = st.text_input("Period / Year End", value="FY 2024-25 (31 March 2025)")

    st.markdown("---")
    st.markdown("### 📐 Reporting Standard")
    standard = st.selectbox(
        "Select Standard",
        ["Ind AS", "IFRS"],
        help="Ind AS → SA 520 references (India). IFRS → ISA 520 references (global)."
    )
    st.caption("Affects: ratio notes, going concern references, cover sheet labels")

    st.markdown("---")
    st.markdown("### ⚙️ Materiality Settings")
    auto_threshold = st.checkbox("Auto-detect materiality threshold", value=True,
                                  help="Recommended. Tool detects if amounts are in ₹, Lakhs, or Crores and scales threshold.")
    if not auto_threshold:
        threshold_pct = st.slider("Variance Threshold (%)", 10, 50, 20, 5)
        threshold_abs = st.number_input("Absolute Threshold", 0.01, 1_000_000.0, 50000.0)
    else:
        threshold_pct = 20
        threshold_abs = None  # will be set after unit detection

    st.markdown("---")
    st.markdown("### 📖 Standards")
    st.markdown("""
    `ISA/SA 520` Analytical Procedures  
    `ISA/SA 570` Going Concern  
    `ISA/SA 240` Fraud (Benford's)  
    `ISA/SA 315` Risk Assessment  
    `IFRS 15 / Ind AS 115` Revenue  
    `IAS 36 / Ind AS 36` Impairment
    """)
    st.markdown("---")
    st.caption("Purva Doshi | ACCA Finalist | v2.0")


# ── Main ──────────────────────────────────────────────────────
st.title("📋 Audit Analytical Procedures Tool")
st.markdown(
    f"**{standard} | ISA/SA 520 · 570 · 240** — Upload a Trial Balance to generate a full workpaper."
)
st.markdown("---")


# ── Tabs ──────────────────────────────────────────────────────
tab_main, tab_benford, tab_help = st.tabs([
    "📊 Trial Balance & Workpaper",
    "🔍 Benford's Law (Fraud)",
    "❓ Help & Format Guide",
])


# ══════════════════════════════════════════════════════════════
# TAB 1 — MAIN ANALYSIS
# ══════════════════════════════════════════════════════════════

with tab_main:
    col_up, col_tip = st.columns([2, 1])

    with col_up:
        st.markdown("### 📁 Upload Trial Balance")
        uploaded_file = st.file_uploader(
            "Upload your Excel file (.xlsx)",
            type=["xlsx"],
            help="Needs: Account Name, CY Amount (₹), PY Amount (₹). Category optional.",
            key="tb_upload",
        )

    with col_tip:
        st.info("""
        **Required columns:**
        - `Account Name`
        - `CY Amount (₹)`
        - `PY Amount (₹)`

        **Auto-handled:**
        - Subtotal rows are skipped
        - "-" treated as zero
        - Units (₹/Lakhs/Crores) auto-detected
        - Category auto-detected if blank
        """)

    # ── Sample download ──
    with st.expander("⬇️ Download sample Trial Balance file"):
        sample_data = {
            "Account Name": [
                "Cash & Bank", "Sundry Debtors", "Inventory / Stock", "Advance Paid", "TDS Receivable",
                "Fixed Assets (Net)", "Capital WIP",
                "Sundry Creditors", "Short Term Loan", "GST Payable",
                "Term Loan (Long Term)",
                "Share Capital", "Reserves & Surplus",
                "Revenue from Operations", "Other Income",
                "Purchases", "Changes in Inventories",
                "Employee Benefits Expense", "Depreciation & Amortisation",
                "Other Expenses", "Finance Costs",
            ],
            "Category": [
                "current_assets", "current_assets", "current_assets", "current_assets", "current_assets",
                "fixed_assets", "fixed_assets",
                "current_liabilities", "current_liabilities", "current_liabilities",
                "long_term_liabilities",
                "equity", "equity",
                "revenue", "revenue",
                "cogs", "cogs",
                "expenses", "expenses", "expenses", "interest_expense",
            ],
            "CY Amount (₹)": [
                5.00, 12.00, 8.00, 1.50, 0.80,
                30.00, 2.00,
                6.00, 2.00, 1.20,
                15.00,
                20.00, 12.00,
                85.00, 15.00,
                52.00, 1.80,
                8.00, 3.00, 9.54, 1.80,
            ],
            "PY Amount (₹)": [
                4.20, 9.50, 7.00, 1.00, 0.65,
                32.00, 0.80,
                5.00, 1.50, 0.95,
                18.00,
                20.00, 8.20,
                72.00, 9.00,
                44.00, -1.26,
                7.20, 2.80, 13.19, 2.10,
            ],
        }
        sample_df = pd.DataFrame(sample_data)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            sample_df.to_excel(tmp.name, index=False)
            tmp_path = tmp.name
        with open(tmp_path, "rb") as f:
            st.download_button(
                "⬇️ Download Sample (₹ Crores format)",
                data=f,
                file_name="sample_trial_balance_crores.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.caption("Amounts are in ₹ Crores. The tool will auto-detect this and scale the threshold.")

    # ── PROCESS ──
    if uploaded_file is not None:
        st.markdown("---")

        mods = _import_modules()
        load_tb    = mods["load_trial_balance"]
        agg_fn     = mods["aggregate_financials"]
        ratio_fn   = mods["calculate_ratios"]
        var_fn     = mods["line_item_variance"]
        gc_fn      = mods["going_concern_assessment"]
        bf_fn      = mods["benfords_law_analysis"]
        summary_fn = mods["generate_summary"]
        report_fn  = mods["generate_report"]

        # Save uploaded file
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_in:
            tmp_in.write(uploaded_file.read())
            input_path = tmp_in.name

        with st.spinner("Loading and parsing trial balance..."):
            try:
                df, unit_info = load_tb(input_path)
            except ValueError as e:
                st.error(f"❌ {e}")
                st.stop()

        unit_label = unit_info["label"]
        detected_threshold = unit_info["threshold"]
        if auto_threshold:
            threshold_abs = detected_threshold

        st.success(
            f"✅ Loaded **{len(df)} accounts** | Detected units: **{unit_label}** | "
            f"Materiality threshold: **{threshold_abs:,.2f}**"
        )

        # Show if any accounts auto-classified
        unclassified = df[df["Category"] == "unclassified"]
        if len(unclassified) > 0:
            with st.expander(f"⚠️ {len(unclassified)} accounts could not be auto-classified — click to review"):
                st.dataframe(unclassified[["Account Name", "CY Amount (₹)", "PY Amount (₹)"]])
                st.caption("Add a 'Category' column to your input file and assign the correct category for these accounts.")

        with st.spinner("Running analytical procedures..."):
            agg     = agg_fn(df)
            ratios  = ratio_fn(agg, standard=standard)
            df_var  = var_fn(df, threshold_pct=threshold_pct, threshold_abs=threshold_abs)
            gc      = gc_fn(agg, ratios, standard=standard)
            summary = summary_fn(agg, df_var, gc, standard=standard)

        # ── Dashboard ──
        st.markdown("### 📈 Results Dashboard")

        def fmt(v, dec=2):
            if v is None: return "N/A"
            try:
                return f"{float(v):,.{dec}f}"
            except: return str(v)

        def delta(cy, py):
            if py and py != 0:
                return f"{((cy-py)/abs(py)*100):+.1f}%"
            return None

        r1c1, r1c2, r1c3, r1c4 = st.columns(4)
        r1c1.metric(f"Revenue (CY) [{unit_label}]",
                    fmt(agg["revenue"]["cy"]),
                    delta(agg["revenue"]["cy"], agg["revenue"]["py"]))
        r1c2.metric(f"Net Profit (CY) [{unit_label}]",
                    fmt(agg["net_profit"]["cy"]),
                    delta(agg["net_profit"]["cy"], agg["net_profit"]["py"]))
        r1c3.metric(f"Gross Profit Margin",
                    f"{(agg['gross_profit']['cy']/agg['revenue']['cy']*100):.1f}%" if agg["revenue"]["cy"] else "N/A",
                    None)
        r1c4.metric(f"Working Capital [{unit_label}]",
                    fmt(agg["working_capital"]["cy"]),
                    "⚠️ NEGATIVE" if agg["working_capital"]["cy"] < 0 else None)

        r2c1, r2c2, r2c3, r2c4 = st.columns(4)
        gc_colors = {"LOW": "🟢", "MODERATE": "🟡", "HIGH": "🟠", "CRITICAL": "🔴"}
        r2c1.metric("Going Concern Risk",
                    f"{gc_colors.get(gc['overall_risk'], '')} {gc['overall_risk']}")
        r2c2.metric("🔴 HIGH Risk Items",  summary.get("Flagged — HIGH Risk", 0))
        r2c3.metric("🟡 MODERATE Items",   summary.get("Flagged — MODERATE", 0))
        r2c4.metric("⚪ Unclassified A/c", summary.get("Unclassified Accounts", 0))

        # ── Tabs within results ──
        res1, res2, res3, res4 = st.tabs([
            "📐 Ratios", "📋 All Accounts", "🚩 Flagged", "🏥 Going Concern"
        ])

        def colour_flag_df(df_show):
            def hl(val):
                if "🔴" in str(val): return "background-color: #FEE2E2"
                if "🟡" in str(val): return "background-color: #FEF9C3"
                if "🟢" in str(val): return "background-color: #DCFCE7"
                return ""
            return df_show.style.applymap(hl, subset=["Risk Flag"] if "Risk Flag" in df_show.columns else
                                          ["Variance Flag"] if "Variance Flag" in df_show.columns else [])

        with res1:
            disp_r = pd.DataFrame(ratios)[[
                "Ratio / Metric", "Current Year", "Prior Year", "YoY Change (%)", "Risk Flag"
            ]].copy()
            for c in ["Current Year", "Prior Year"]:
                disp_r[c] = disp_r[c].apply(lambda x: f"{x:.2f}" if isinstance(x, float) else x)
            disp_r["YoY Change (%)"] = disp_r["YoY Change (%)"].apply(
                lambda x: f"{x:+.1f}%" if isinstance(x, float) else x)
            st.dataframe(colour_flag_df(disp_r), use_container_width=True, height=420)

        with res2:
            disp_v = df_var[["Account Name","Category","CY Amount (₹)","PY Amount (₹)","Change (₹)","Change (%)","Variance Flag"]].copy()
            for c in ["CY Amount (₹)","PY Amount (₹)","Change (₹)"]:
                disp_v[c] = disp_v[c].apply(lambda x: f"{x:,.2f}" if isinstance(x,(int,float)) else x)
            disp_v["Change (%)"] = disp_v["Change (%)"].apply(
                lambda x: f"{x:+.1f}%" if isinstance(x, float) and x != 999.0 else ("New" if x == 999.0 else x))
            st.dataframe(colour_flag_df(disp_v), use_container_width=True, height=480)

        with res3:
            flagged = df_var[df_var["Variance Flag"].str.contains("🔴|🟡", na=False)]
            if flagged.empty:
                st.success("🎉 No flagged items — all accounts within threshold.")
            else:
                st.warning(f"**{len(flagged)} accounts** need attention.")
                disp_f = flagged[["Account Name","Category","CY Amount (₹)","PY Amount (₹)","Change (₹)","Change (%)","Variance Flag"]].copy()
                for c in ["CY Amount (₹)","PY Amount (₹)","Change (₹)"]:
                    disp_f[c] = disp_f[c].apply(lambda x: f"{x:,.2f}" if isinstance(x,(int,float)) else x)
                disp_f["Change (%)"] = disp_f["Change (%)"].apply(
                    lambda x: f"{x:+.1f}%" if isinstance(x,float) and x != 999.0 else ("New" if x==999.0 else x))
                st.dataframe(colour_flag_df(disp_f), use_container_width=True)

        with res4:
            risk_color_map = {"CRITICAL":"#DC2626","HIGH":"#EA580C","MODERATE":"#D97706","LOW":"#16A34A"}
            bg = risk_color_map.get(gc["overall_risk"], "#1F2D3D")
            st.markdown(
                f'<div style="background:{bg};color:white;padding:16px;border-radius:8px;font-size:18px;font-weight:700;">'
                f'Overall Risk: {gc["overall_risk"]}  |  Score: {gc["score"]}/10</div>',
                unsafe_allow_html=True
            )
            st.markdown(f"> {gc['conclusion']}")
            st.markdown("#### Indicators")
            gc_df = pd.DataFrame(gc["indicators"])
            def hl_gc(val):
                if "🔴" in str(val): return "background-color: #FEE2E2"
                if "🟡" in str(val): return "background-color: #FEF9C3"
                if "🟢" in str(val): return "background-color: #DCFCE7"
                return ""
            st.dataframe(gc_df.style.applymap(hl_gc, subset=["Status"]),
                         use_container_width=True)

        # ── Generate workpaper ──
        st.markdown("---")
        st.markdown("### 📥 Generate Excel Workpaper")

        # Get Benford data if available from session state
        bf_result = st.session_state.get("bf_result", {
            "sufficient": False,
            "message": "No transaction data uploaded. Use the Benford's Law tab to add this.",
            "risk_flag": "⚪ Not Analysed",
        })

        if st.button("🔄 Generate Full 7-Sheet Workpaper", type="primary"):
            with st.spinner("Building workpaper — formatting all 7 sheets..."):
                out_path = os.path.join(tempfile.gettempdir(), "AP_Workpaper_v2.xlsx")
                report_fn(
                    df_variance=df_var,
                    ratios=ratios,
                    summary=summary,
                    gc=gc,
                    bf=bf_result,
                    output_path=out_path,
                    client_name=client_name,
                    period=period,
                    standard=standard,
                    unit_label=unit_label,
                )

            st.success("✅ Workpaper ready!")
            with open(out_path, "rb") as f:
                st.download_button(
                    f"⬇️ Download AP Workpaper — {client_name}",
                    data=f,
                    file_name=f"AP_Workpaper_v2_{client_name.replace(' ','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            st.info("""
            **7-sheet workpaper contains:**
            1. Cover & Summary  — engagement details + risk overview
            2. Ratio Analysis   — 15 ratios with standard-specific notes
            3. Variance         — all accounts with YoY movement
            4. Flagged Items    — HIGH & MODERATE risk only
            5. Going Concern    — ISA/SA 570 assessment (6 indicators)
            6. Benford's Law    — ISA/SA 240 fraud indicator (if data uploaded)
            7. Instructions     — format guide and audit references
            """)


# ══════════════════════════════════════════════════════════════
# TAB 2 — BENFORD'S LAW
# ══════════════════════════════════════════════════════════════

with tab_benford:
    st.markdown("### 🔍 Benford's Law Analysis — ISA/SA 240 Fraud Risk Indicator")
    st.markdown("""
    > **What is Benford's Law?**
    > In naturally occurring financial data, the first digit is NOT random.
    > Digit **1** appears ~30% of the time, digit **9** only ~4.6%.
    > If your transactions show a different pattern, it may indicate **manipulation, rounding, or fabrication**.
    > This is a real procedure used by Big 4 forensic teams.
    """)

    col_bf1, col_bf2 = st.columns([2, 1])

    with col_bf1:
        bf_file = st.file_uploader(
            "Upload transaction amounts file (.xlsx)",
            type=["xlsx"],
            help="One column named 'Amount' with individual transaction amounts. Min 20, ideally 500+.",
            key="bf_upload",
        )

    with col_bf2:
        st.info("""
        **Input format:**
        Excel with one column named `Amount`

        **Good sources:**
        - Journal entries
        - Purchase invoices
        - Payment vouchers
        - GL ledger lines
        """)

    # Sample download for Benford
    with st.expander("⬇️ Download sample transaction file"):
        import random, math as _math
        random.seed(42)
        # Generate realistic-looking Benford-compliant amounts using log-uniform distribution
        benford_amounts = [round(10 ** random.uniform(2, 6), 2) for _ in range(300)]
        bf_sample = pd.DataFrame({"Amount": benford_amounts})
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            bf_sample.to_excel(tmp.name, index=False)
            bf_tmp = tmp.name
        with open(bf_tmp, "rb") as f:
            st.download_button(
                "⬇️ Download Sample Transaction File",
                data=f, file_name="sample_transactions_benford.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    if bf_file is not None:
        mods = _import_modules()
        bf_fn = mods["benfords_law_analysis"]

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_bf:
            tmp_bf.write(bf_file.read())
            bf_path = tmp_bf.name

        bf_df = pd.read_excel(bf_path)
        # Find amount column
        amt_col = None
        for c in bf_df.columns:
            if "amount" in c.lower() or "value" in c.lower():
                amt_col = c
                break
        if amt_col is None:
            amt_col = bf_df.columns[0]

        amounts = bf_df[amt_col].tolist()

        with st.spinner("Running Benford's Law analysis..."):
            bf_result = bf_fn(amounts, label=f"{bf_file.name} ({amt_col})")

        # Store in session for workpaper
        st.session_state["bf_result"] = bf_result

        st.markdown(f"### Result: {bf_result.get('risk_flag', '')}")
        if bf_result.get("sufficient"):
            st.markdown(f"> {bf_result['interpretation']}")
            st.markdown(f"**Chi-Square: {bf_result['chi_square']:.2f}**  |  "
                        f"Critical: p<0.10 → 13.36 | p<0.05 → 15.51 | p<0.01 → 20.09")

            bf_disp = pd.DataFrame(bf_result["digit_data"])

            def hl_bf(val):
                v = abs(float(val)) if isinstance(val, (int, float)) else 0
                if v > 5: return "background-color: #FEE2E2"
                if v > 2: return "background-color: #FEF9C3"
                return "background-color: #DCFCE7"

            st.dataframe(
                bf_disp.style.applymap(hl_bf, subset=["Deviation (pp)"]),
                use_container_width=True, height=350
            )
            st.caption("🔴 = deviation >5pp | 🟡 = 2–5pp | 🟢 = within expected range")
        else:
            st.warning(bf_result.get("message", "Insufficient data"))

    else:
        st.info("Upload a transaction amounts file above to run Benford's Law analysis.")
        st.markdown("""
        **Why Benford's Law matters in interviews:**
        > *"I implemented Benford's Law analysis in my audit tool — it flags unusual digit patterns
        in transaction amounts that could indicate fabrication or manipulation, as referenced in ISA/SA 240."*

        This shows you understand forensic audit concepts beyond the standard analytical procedures syllabus.
        """)


# ══════════════════════════════════════════════════════════════
# TAB 3 — HELP
# ══════════════════════════════════════════════════════════════

with tab_help:
    st.markdown("### ❓ Help & Format Guide")

    with st.expander("📁 Trial Balance Input Format", expanded=True):
        st.markdown("""
        | Column | Required? | Example |
        |--------|-----------|---------|
        | `Account Name` | ✅ Yes | Cash & Bank |
        | `CY Amount (₹)` | ✅ Yes | 500000 or 5.00 (Crores) |
        | `PY Amount (₹)` | ✅ Yes | 420000 |
        | `Category` | ❌ Optional | current_assets |

        **Valid Category values:**
        `current_assets` · `fixed_assets` · `current_liabilities` · `long_term_liabilities`
        · `equity` · `revenue` · `cogs` · `expenses` · `interest_expense`

        **Handled automatically:**
        - Subtotal rows ("Total Income", "Total Expenses", etc.) → **auto-skipped**
        - "-" values → treated as zero
        - Units (₹, Lakhs, Crores) → **auto-detected**, threshold scaled
        - Category blank → **auto-classified** from account name
        """)

    with st.expander("📐 IFRS vs Ind AS — What changes?"):
        st.markdown("""
        | Feature | IFRS | Ind AS |
        |---------|------|--------|
        | Standards reference | ISA 520 / IAS 1 | SA 520 / Ind AS 1 |
        | Revenue standard | IFRS 15 | Ind AS 115 |
        | Going concern | ISA 570 | SA 570 |
        | Fraud procedures | ISA 240 | SA 240 |
        | Lease recognition | IFRS 16 | Ind AS 116 |
        | Impairment | IAS 36 | Ind AS 36 |
        | Provisions | IAS 37 | Ind AS 37 |

        For ratio calculation, both standards produce identical results.
        The difference is in the audit references cited in the workpaper and the notes attached to each ratio flag.
        """)

    with st.expander("🔍 Benford's Law — Simple Explanation"):
        st.markdown("""
        **What it is:** A mathematical law that says first digits in natural data follow a specific pattern:
        - Digit 1 should appear ~30% of the time
        - Digit 9 should appear only ~4.6% of the time

        **Why auditors use it:** If someone is making up numbers (fraud), they tend to use digits too evenly,
        or avoid certain digits. Benford's Law catches this.

        **Reference:** ISA/SA 240 — Auditor's responsibilities for fraud

        **How to use in the tool:**
        1. Export a ledger of individual transactions from Tally (journal entries, purchase invoices)
        2. Upload as Excel with an 'Amount' column
        3. The tool runs the chi-square test and flags unusual patterns
        """)

    with st.expander("🚀 How to run the tool — step by step"):
        st.markdown("""
        ```bash
        # 1. Create a folder and put all files in it
        mkdir audit_ap_tool_v2
        cd audit_ap_tool_v2

        # 2. Install required libraries
        pip install -r requirements.txt

        # 3. Run
        streamlit run app_v2.py

        # 4. Browser opens at http://localhost:8501
        ```

        **Files needed:**
        - `app_v2.py` — the app (this file)
        - `analyzer_v2.py` — core analysis logic
        - `reporter_v2.py` — Excel output generator
        - `requirements.txt` — list of libraries
        """)

    with st.expander("💬 Interview talking points"):
        st.markdown(f"""
        > *"I built a Python audit tool that automates ISA/SA 520 analytical procedures with
        dual IFRS and Ind AS support."*

        > *"It computes 15 financial ratios, auto-detects unit scale (lakhs/crores), and flags
        material year-on-year variances based on a materiality threshold."*

        > *"The tool runs an ISA/SA 570 going concern assessment across 6 indicators —
        scoring risk as Low, Moderate, High, or Critical."*

        > *"I also implemented Benford's Law (ISA/SA 240) for fraud detection on transaction-level data."*

        > *"The output is a formatted 7-sheet Excel workpaper — the kind actually filed in a Big 4 audit folder.
        I used this on the Akin Chemicals statutory audit."*
        """)
