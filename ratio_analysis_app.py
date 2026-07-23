"""
Financial Ratio Analysis Tool — for student learning
------------------------------------------------------
Students manually enter balance sheet & income statement figures.
The app computes 18 standard ratios across four categories and
gives a plain-language interpretation for each one.

Run with:  streamlit run ratio_analysis_app.py
"""

import streamlit as st
import pandas as pd

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Financial Ratio Analysis Tool",
    page_icon="📊",
    layout="wide",
)

# ============================================================
# HELPERS
# ============================================================

def safe_div(numerator, denominator):
    """Divide safely; return None if denominator is zero/blank."""
    if denominator in (0, None):
        return None
    return numerator / denominator


def fmt(value, suffix="", decimals=2):
    """Format a ratio value for display, handling None gracefully."""
    if value is None:
        return "N/A"
    return f"{value:,.{decimals}f}{suffix}"


def render_ratio(col, name, formula_text, value, level, message, value_suffix=""):
    """Render one ratio as a metric + colored interpretation box."""
    with col:
        st.metric(name, fmt(value, value_suffix))
        st.caption(f"Formula: {formula_text}")
        if level == "good":
            st.success(message)
        elif level == "ok":
            st.info(message)
        elif level == "warning":
            st.warning(message)
        else:
            st.error(message)
        st.write("")  # spacing


NA_MSG = "Can't be calculated — check that the denominator in the formula above isn't zero or blank."

# ============================================================
# INTERPRETATION FUNCTIONS
# Each returns (level, message). level in {good, ok, warning, na}
# Thresholds are common textbook rules of thumb — real benchmarks
# vary by industry, so treat these as a starting point for discussion.
# ============================================================

def interp_current_ratio(v):
    if v is None:
        return "na", NA_MSG
    if v >= 2:
        return "good", "Comfortably above 2, meaning current assets cover current liabilities more than twice over. Strong short-term liquidity, though very high values can also mean idle cash or excess inventory."
    if v >= 1:
        return "ok", "Between 1 and 2 — current assets cover current liabilities, but the safety margin is moderate. Worth comparing against industry peers."
    return "warning", "Below 1 — current liabilities exceed current assets. This is a liquidity warning sign; the company may struggle to meet short-term obligations."


def interp_quick_ratio(v):
    if v is None:
        return "na", NA_MSG
    if v >= 1:
        return "good", "At or above 1 — the company can cover current liabilities without relying on selling inventory. Solid immediate liquidity."
    if v >= 0.7:
        return "ok", "Slightly below 1 — liquid assets almost cover current liabilities. Reasonable, but leaves little room for error."
    return "warning", "Well below 1 — the company depends heavily on inventory or other slow-moving assets to meet short-term debts."


def interp_cash_ratio(v):
    if v is None:
        return "na", NA_MSG
    if v >= 0.5:
        return "good", "High cash coverage of current liabilities — very conservative liquidity position. Could also suggest cash is not being put to productive use."
    if v >= 0.2:
        return "ok", "Moderate cash coverage — a reasonable cash buffer against current liabilities."
    return "warning", "Low cash coverage — the company holds little cash relative to what it owes in the short term."


def interp_gross_margin(v):
    if v is None:
        return "na", NA_MSG
    if v >= 0.40:
        return "good", "A high gross margin suggests strong pricing power or low production costs relative to sales."
    if v >= 0.20:
        return "ok", "A moderate gross margin — typical of many manufacturing and retail businesses."
    return "warning", "A low gross margin leaves little cushion after production costs — worth checking cost control and pricing strategy."


def interp_operating_margin(v):
    if v is None:
        return "na", NA_MSG
    if v >= 0.15:
        return "good", "Strong operating margin — core business operations are generating healthy profit before interest and tax."
    if v >= 0.05:
        return "ok", "Positive but modest operating margin — operating expenses are consuming a large share of gross profit."
    return "warning", "Weak or negative operating margin — operating costs are eating heavily into profitability."


def interp_net_margin(v):
    if v is None:
        return "na", NA_MSG
    if v >= 0.10:
        return "good", "A healthy net margin — for every dollar of sales, a solid portion becomes bottom-line profit."
    if v >= 0.02:
        return "ok", "A thin but positive net margin — profitable, but with limited buffer against rising costs."
    return "warning", "Very low or negative net margin — the company is barely profitable (or is losing money) after all expenses."


def interp_roa(v):
    if v is None:
        return "na", NA_MSG
    if v >= 0.10:
        return "good", "The company is generating strong profit relative to the assets it owns — efficient use of the asset base."
    if v >= 0.03:
        return "ok", "Modest returns on assets — assets are generating some profit, but there may be room to use them more efficiently."
    return "warning", "Low or negative ROA — assets are not generating much profit relative to their size."


def interp_roe(v):
    if v is None:
        return "na", NA_MSG
    if v >= 0.15:
        return "good", "Strong return for shareholders — the company is generating healthy profit relative to equity invested."
    if v >= 0.05:
        return "ok", "Positive but modest return to shareholders — acceptable, but compare against the return investors could get elsewhere."
    return "warning", "Low or negative ROE — shareholders are earning little (or losing value) on their investment in the company."


def interp_asset_turnover(v):
    if v is None:
        return "na", NA_MSG
    if v >= 1.0:
        return "good", "The company generates at least a dollar of sales for every dollar of assets — efficient asset use."
    if v >= 0.5:
        return "ok", "Moderate efficiency in using assets to generate sales. Typical for capital-intensive industries."
    return "warning", "Low asset turnover — the company is generating relatively little sales from its asset base, which may indicate underused or excess assets."


def interp_inventory_turnover(v):
    if v is None:
        return "na", NA_MSG
    if v >= 8:
        return "good", "Inventory is selling and being replaced quickly — efficient inventory management."
    if v >= 3:
        return "ok", "Reasonable inventory turnover — goods move at a moderate pace."
    return "warning", "Low inventory turnover — inventory is sitting for a long time before being sold, tying up cash and raising obsolescence risk."


def interp_dio(v):
    if v is None:
        return "na", NA_MSG
    if v <= 45:
        return "good", "Inventory is held for a relatively short period before sale — efficient inventory cycle."
    if v <= 90:
        return "ok", "A moderate number of days to sell through inventory."
    return "warning", "Inventory sits for a long time before being sold — this ties up working capital."


def interp_receivables_turnover(v):
    if v is None:
        return "na", NA_MSG
    if v >= 10:
        return "good", "Receivables are collected quickly — effective credit and collections management."
    if v >= 4:
        return "ok", "Moderate collection speed on credit sales."
    return "warning", "Receivables turn over slowly — customers are taking a long time to pay, which can strain cash flow."


def interp_dso(v):
    if v is None:
        return "na", NA_MSG
    if v <= 45:
        return "good", "Customers pay relatively quickly after a credit sale — healthy collections."
    if v <= 90:
        return "ok", "A moderate collection period — worth monitoring against the company's credit terms."
    return "warning", "A long collection period — cash is tied up in receivables for a long time, which can create cash flow strain."


def interp_debt_ratio(v):
    if v is None:
        return "na", NA_MSG
    if v <= 0.4:
        return "good", "A relatively low proportion of assets is financed by debt — conservative financing and lower financial risk."
    if v <= 0.6:
        return "ok", "A moderate reliance on debt financing — common for many established companies."
    return "warning", "A high proportion of assets is financed by debt — this raises financial risk and sensitivity to interest rate or earnings shocks."


def interp_debt_to_equity(v):
    if v is None:
        return "na", NA_MSG
    if v <= 1.0:
        return "good", "Debt is less than or equal to equity — a relatively conservative capital structure."
    if v <= 2.0:
        return "ok", "Debt moderately exceeds equity — leverage is elevated but not extreme; depends heavily on the industry norm."
    return "warning", "Debt significantly exceeds equity — high financial leverage increases risk to shareholders and creditors alike."


def interp_equity_multiplier(v):
    if v is None:
        return "na", NA_MSG
    if v <= 2.0:
        return "good", "Assets are financed mostly by equity rather than debt — lower financial leverage."
    if v <= 3.0:
        return "ok", "A moderate level of financial leverage — a meaningful portion of assets is debt-financed."
    return "warning", "High financial leverage — a large share of assets is financed by debt relative to equity, amplifying both potential returns and risk."


def interp_interest_coverage(v):
    if v is None:
        return "na", NA_MSG
    if v >= 5:
        return "good", "Operating profit covers interest expense many times over — little risk of default on interest payments."
    if v >= 1.5:
        return "ok", "Operating profit covers interest, but the cushion is moderate — a downturn in earnings could pressure debt payments."
    return "warning", "Operating profit barely covers (or doesn't cover) interest expense — a serious warning sign for solvency."


# ============================================================
# SIDEBAR — instructions
# ============================================================
with st.sidebar:
    st.header("📘 How to use this tool")
    st.markdown(
        """
        1. Go to the **Enter Data** tab and type in the figures from
           a balance sheet and income statement (sample values are
           pre-filled so you can explore first).
        2. Click **Calculate Ratios**.
        3. Browse the **Liquidity**, **Profitability**, **Efficiency**,
           and **Solvency** tabs to see each ratio with its formula
           and a plain-language interpretation.
        4. Check the **Summary Dashboard** for a one-page overview,
           and download it as a CSV if you'd like to keep a record.

        *Note:* Interpretation thresholds are general classroom
        rules of thumb. Real-world benchmarks vary a lot by industry —
        use these as a starting point for discussion, not a strict rule.
        """
    )
    company_name = st.text_input("Company name (optional)", value="Sample Company Inc.")
    period_label = st.text_input("Period label (optional)", value="Year Ended Dec 31")

st.title("📊 Financial Ratio Analysis Tool")
st.caption("Enter your own numbers, then explore what each ratio means.")

tab_input, tab_liq, tab_prof, tab_eff, tab_solv, tab_summary = st.tabs(
    ["📥 Enter Data", "💧 Liquidity", "💰 Profitability", "⚙️ Efficiency", "🏦 Solvency", "📋 Summary Dashboard"]
)

# ============================================================
# TAB: ENTER DATA
# ============================================================
with tab_input:
    st.subheader("Balance Sheet & Income Statement Inputs")
    st.write(
        "Enter figures for a single period. All totals (current assets, "
        "total assets, current liabilities, total liabilities) are "
        "calculated automatically from the line items below."
    )

    with st.form("data_form"):
        col_bs, col_is = st.columns(2)

        with col_bs:
            st.markdown("### 🏦 Balance Sheet")
            st.markdown("**Current Assets**")
            cash = st.number_input("Cash & Cash Equivalents", value=15000.0, step=500.0)
            receivables = st.number_input("Accounts Receivable", value=25000.0, step=500.0)
            inventory = st.number_input("Inventory", value=30000.0, step=500.0)
            other_ca = st.number_input("Other Current Assets", value=5000.0, step=500.0)

            st.markdown("**Non-Current Assets**")
            fixed_assets = st.number_input("Net Fixed / Non-Current Assets", value=120000.0, step=1000.0)

            st.markdown("**Current Liabilities**")
            payables = st.number_input("Accounts Payable", value=18000.0, step=500.0)
            other_cl = st.number_input("Other Current Liabilities (e.g. short-term notes)", value=7000.0, step=500.0)

            st.markdown("**Non-Current Liabilities & Equity**")
            long_term_debt = st.number_input("Long-Term Debt", value=50000.0, step=1000.0)
            equity = st.number_input("Shareholders' Equity", value=120000.0, step=1000.0)

        with col_is:
            st.markdown("### 💵 Income Statement")
            sales = st.number_input("Net Sales / Revenue", value=250000.0, step=1000.0)
            cogs = st.number_input("Cost of Goods Sold (COGS)", value=150000.0, step=1000.0)
            operating_expenses = st.number_input("Operating Expenses (SG&A, excl. interest & tax)", value=60000.0, step=1000.0)
            interest_expense = st.number_input("Interest Expense", value=5000.0, step=250.0)
            tax_expense = st.number_input("Tax Expense", value=8000.0, step=250.0)

            st.markdown("---")
            st.caption(
                "Gross Profit, Operating Income (EBIT), and Net Income "
                "are calculated automatically:\n\n"
                "Gross Profit = Sales − COGS\n\n"
                "EBIT = Gross Profit − Operating Expenses\n\n"
                "Net Income = EBIT − Interest − Tax"
            )

        submitted = st.form_submit_button("Calculate Ratios", type="primary", use_container_width=True)

    # Store to session state (also runs on first load so tabs aren't empty)
    if submitted or "data" not in st.session_state:
        total_current_assets = cash + receivables + inventory + other_ca
        total_assets = total_current_assets + fixed_assets
        total_current_liabilities = payables + other_cl
        total_liabilities = total_current_liabilities + long_term_debt

        gross_profit = sales - cogs
        ebit = gross_profit - operating_expenses
        ebt = ebit - interest_expense
        net_income = ebt - tax_expense

        st.session_state["data"] = dict(
            cash=cash, receivables=receivables, inventory=inventory, other_ca=other_ca,
            fixed_assets=fixed_assets, total_current_assets=total_current_assets, total_assets=total_assets,
            payables=payables, other_cl=other_cl, total_current_liabilities=total_current_liabilities,
            long_term_debt=long_term_debt, total_liabilities=total_liabilities, equity=equity,
            sales=sales, cogs=cogs, operating_expenses=operating_expenses,
            interest_expense=interest_expense, tax_expense=tax_expense,
            gross_profit=gross_profit, ebit=ebit, ebt=ebt, net_income=net_income,
        )

    d = st.session_state["data"]

    st.markdown("### Quick Check")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Current Assets", f"{d['total_current_assets']:,.0f}")
    c1.metric("Total Assets", f"{d['total_assets']:,.0f}")
    c2.metric("Total Current Liabilities", f"{d['total_current_liabilities']:,.0f}")
    c2.metric("Total Liabilities", f"{d['total_liabilities']:,.0f}")
    c3.metric("Net Income", f"{d['net_income']:,.0f}")
    c3.metric("EBIT (Operating Income)", f"{d['ebit']:,.0f}")

    balance_check = d["total_assets"] - (d["total_liabilities"] + d["equity"])
    if abs(balance_check) < 0.01:
        st.success("✅ Balance sheet balances: Total Assets = Total Liabilities + Equity")
    else:
        st.warning(
            f"⚠️ Balance sheet does not balance. Assets minus (Liabilities + Equity) = "
            f"{balance_check:,.2f}. Consider adjusting the Equity figure so the books balance."
        )

# Pull data for use in the ratio tabs below
d = st.session_state["data"]

# ============================================================
# COMPUTE ALL RATIOS
# ============================================================
current_ratio = safe_div(d["total_current_assets"], d["total_current_liabilities"])
quick_ratio = safe_div(d["total_current_assets"] - d["inventory"], d["total_current_liabilities"])
cash_ratio = safe_div(d["cash"], d["total_current_liabilities"])

gross_margin = safe_div(d["gross_profit"], d["sales"])
operating_margin = safe_div(d["ebit"], d["sales"])
net_margin = safe_div(d["net_income"], d["sales"])
roa = safe_div(d["net_income"], d["total_assets"])
roe = safe_div(d["net_income"], d["equity"])

asset_turnover = safe_div(d["sales"], d["total_assets"])
inventory_turnover = safe_div(d["cogs"], d["inventory"])
dio = safe_div(365, inventory_turnover) if inventory_turnover else None
receivables_turnover = safe_div(d["sales"], d["receivables"])
dso = safe_div(365, receivables_turnover) if receivables_turnover else None

debt_ratio = safe_div(d["total_liabilities"], d["total_assets"])
debt_to_equity = safe_div(d["total_liabilities"], d["equity"])
equity_multiplier = safe_div(d["total_assets"], d["equity"])
interest_coverage = safe_div(d["ebit"], d["interest_expense"])

# ============================================================
# TAB: LIQUIDITY
# ============================================================
with tab_liq:
    st.subheader("💧 Liquidity Ratios")
    st.write("Liquidity ratios measure whether a company can meet its short-term obligations.")
    c1, c2, c3 = st.columns(3)
    render_ratio(c1, "Current Ratio", "Current Assets ÷ Current Liabilities", current_ratio, *interp_current_ratio(current_ratio))
    render_ratio(c2, "Quick Ratio (Acid-Test)", "(Current Assets − Inventory) ÷ Current Liabilities", quick_ratio, *interp_quick_ratio(quick_ratio))
    render_ratio(c3, "Cash Ratio", "Cash ÷ Current Liabilities", cash_ratio, *interp_cash_ratio(cash_ratio))

    nwc = d["total_current_assets"] - d["total_current_liabilities"]
    st.info(f"**Net Working Capital** = Current Assets − Current Liabilities = **{nwc:,.0f}**. "
            f"{'Positive working capital means short-term assets exceed short-term obligations.' if nwc >= 0 else 'Negative working capital means short-term obligations exceed short-term assets — a liquidity concern.'}")

# ============================================================
# TAB: PROFITABILITY
# ============================================================
with tab_prof:
    st.subheader("💰 Profitability Ratios")
    st.write("Profitability ratios measure how well a company converts sales and assets into profit.")
    c1, c2, c3 = st.columns(3)
    render_ratio(c1, "Gross Profit Margin", "Gross Profit ÷ Sales", gross_margin, *interp_gross_margin(gross_margin), value_suffix="")
    render_ratio(c2, "Operating Profit Margin", "EBIT ÷ Sales", operating_margin, *interp_operating_margin(operating_margin))
    render_ratio(c3, "Net Profit Margin", "Net Income ÷ Sales", net_margin, *interp_net_margin(net_margin))
    c4, c5 = st.columns(2)
    render_ratio(c4, "Return on Assets (ROA)", "Net Income ÷ Total Assets", roa, *interp_roa(roa))
    render_ratio(c5, "Return on Equity (ROE)", "Net Income ÷ Shareholders' Equity", roe, *interp_roe(roe))

# ============================================================
# TAB: EFFICIENCY
# ============================================================
with tab_eff:
    st.subheader("⚙️ Efficiency (Activity) Ratios")
    st.write("Efficiency ratios measure how well a company uses its assets to generate sales.")
    c1, c2 = st.columns(2)
    render_ratio(c1, "Asset Turnover", "Sales ÷ Total Assets", asset_turnover, *interp_asset_turnover(asset_turnover), value_suffix="x")
    render_ratio(c2, "Inventory Turnover", "COGS ÷ Inventory", inventory_turnover, *interp_inventory_turnover(inventory_turnover), value_suffix="x")
    c3, c4 = st.columns(2)
    render_ratio(c3, "Days Inventory Outstanding", "365 ÷ Inventory Turnover", dio, *interp_dio(dio), value_suffix=" days")
    render_ratio(c4, "Receivables Turnover", "Sales ÷ Accounts Receivable", receivables_turnover, *interp_receivables_turnover(receivables_turnover), value_suffix="x")
    c5, _ = st.columns(2)
    render_ratio(c5, "Days Sales Outstanding (DSO)", "365 ÷ Receivables Turnover", dso, *interp_dso(dso), value_suffix=" days")

# ============================================================
# TAB: SOLVENCY
# ============================================================
with tab_solv:
    st.subheader("🏦 Solvency (Leverage) Ratios")
    st.write("Solvency ratios measure long-term financial risk and reliance on debt financing.")
    c1, c2 = st.columns(2)
    render_ratio(c1, "Debt Ratio", "Total Liabilities ÷ Total Assets", debt_ratio, *interp_debt_ratio(debt_ratio))
    render_ratio(c2, "Debt-to-Equity Ratio", "Total Liabilities ÷ Equity", debt_to_equity, *interp_debt_to_equity(debt_to_equity), value_suffix="x")
    c3, c4 = st.columns(2)
    render_ratio(c3, "Equity Multiplier", "Total Assets ÷ Equity", equity_multiplier, *interp_equity_multiplier(equity_multiplier), value_suffix="x")
    render_ratio(c4, "Interest Coverage Ratio", "EBIT ÷ Interest Expense", interest_coverage, *interp_interest_coverage(interest_coverage), value_suffix="x")

# ============================================================
# TAB: SUMMARY DASHBOARD
# ============================================================
with tab_summary:
    st.subheader("📋 Summary Dashboard")
    st.caption(f"{company_name} — {period_label}")

    rows = [
        ("Liquidity", "Current Ratio", current_ratio, interp_current_ratio(current_ratio)[0]),
        ("Liquidity", "Quick Ratio", quick_ratio, interp_quick_ratio(quick_ratio)[0]),
        ("Liquidity", "Cash Ratio", cash_ratio, interp_cash_ratio(cash_ratio)[0]),
        ("Profitability", "Gross Profit Margin", gross_margin, interp_gross_margin(gross_margin)[0]),
        ("Profitability", "Operating Profit Margin", operating_margin, interp_operating_margin(operating_margin)[0]),
        ("Profitability", "Net Profit Margin", net_margin, interp_net_margin(net_margin)[0]),
        ("Profitability", "Return on Assets", roa, interp_roa(roa)[0]),
        ("Profitability", "Return on Equity", roe, interp_roe(roe)[0]),
        ("Efficiency", "Asset Turnover", asset_turnover, interp_asset_turnover(asset_turnover)[0]),
        ("Efficiency", "Inventory Turnover", inventory_turnover, interp_inventory_turnover(inventory_turnover)[0]),
        ("Efficiency", "Days Inventory Outstanding", dio, interp_dio(dio)[0]),
        ("Efficiency", "Receivables Turnover", receivables_turnover, interp_receivables_turnover(receivables_turnover)[0]),
        ("Efficiency", "Days Sales Outstanding", dso, interp_dso(dso)[0]),
        ("Solvency", "Debt Ratio", debt_ratio, interp_debt_ratio(debt_ratio)[0]),
        ("Solvency", "Debt-to-Equity", debt_to_equity, interp_debt_to_equity(debt_to_equity)[0]),
        ("Solvency", "Equity Multiplier", equity_multiplier, interp_equity_multiplier(equity_multiplier)[0]),
        ("Solvency", "Interest Coverage", interest_coverage, interp_interest_coverage(interest_coverage)[0]),
    ]

    level_label = {"good": "✅ Strong", "ok": "🟡 Moderate", "warning": "🔴 Weak", "na": "⚪ N/A"}
    df = pd.DataFrame(rows, columns=["Category", "Ratio", "Value", "Level"])
    df["Assessment"] = df["Level"].map(level_label)
    df_display = df[["Category", "Ratio", "Value", "Assessment"]].copy()
    df_display["Value"] = df_display["Value"].apply(lambda v: "N/A" if v is None else round(v, 3))

    st.dataframe(df_display, use_container_width=True, hide_index=True)

    counts = df["Level"].value_counts()
    cc1, cc2, cc3, cc4 = st.columns(4)
    cc1.metric("✅ Strong", int(counts.get("good", 0)))
    cc2.metric("🟡 Moderate", int(counts.get("ok", 0)))
    cc3.metric("🔴 Weak", int(counts.get("warning", 0)))
    cc4.metric("⚪ Not calculable", int(counts.get("na", 0)))

    csv = df_display.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download Summary as CSV",
        data=csv,
        file_name=f"{company_name.replace(' ', '_')}_ratio_summary.csv",
        mime="text/csv",
        use_container_width=True,
    )

    st.caption(
        "Reminder for students: these interpretations use general rules of thumb. "
        "Always compare a company's ratios against its own history and against "
        "industry peers, since 'good' and 'bad' vary a lot by sector."
    )
