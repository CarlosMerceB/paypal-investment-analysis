"""
PayPal (PYPL) - SEC EDGAR XBRL Data Extraction
================================================
Fills the FY2019-FY2021 gap that yfinance doesn't cover.
Uses SEC's Company Facts API (free, no API key needed).

Run AFTER 01_extract_paypal_data.py

Required: pip install requests pandas --break-system-packages
"""

import requests
import pandas as pd
import json
import os
from datetime import datetime

# =============================================================================
# CONFIGURATION
# =============================================================================
SEC_CIK = "0001633917"  # PayPal CIK (zero-padded to 10 digits)
SEC_HEADERS = {"User-Agent": "InvestmentAnalysis carlos@example.com"}
TARGET_YEARS = [2019, 2020, 2021]
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "raw")
PROCESSED_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "processed")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(PROCESSED_DIR, exist_ok=True)

# =============================================================================
# XBRL TAG MAPPING
# =============================================================================
# Maps SEC XBRL tags to our standardized line item names.
# PayPal files under us-gaap taxonomy.
# Some tags have changed over the years, so we list alternatives.

INCOME_STATEMENT_TAGS = {
    "Total Revenue": [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "Revenues",
        "RevenueFromContractWithCustomerNet",
    ],
    "Cost of Revenue": [
        "CostOfRevenue",
        "CostOfGoodsAndServicesSold",
    ],
    "Gross Profit": [
        "GrossProfit",
    ],
    "Operating Expenses": [
        "OperatingExpenses",
    ],
    "Operating Income": [
        "OperatingIncomeLoss",
    ],
    "Interest Income": [
        "InvestmentIncomeInterest",
        "InterestIncomeOther",
    ],
    "Interest Expense": [
        "InterestExpense",
    ],
    "Income Before Taxes": [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxes",
    ],
    "Income Tax Expense": [
        "IncomeTaxExpenseBenefit",
    ],
    "Net Income": [
        "NetIncomeLoss",
    ],
    "EPS Diluted": [
        "EarningsPerShareDiluted",
    ],
    "EPS Basic": [
        "EarningsPerShareBasic",
    ],
    "Shares Outstanding (Diluted)": [
        "WeightedAverageNumberOfDilutedSharesOutstanding",
    ],
}

BALANCE_SHEET_TAGS = {
    "Cash & Cash Equivalents": [
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsAndShortTermInvestments",
    ],
    "Short-Term Investments": [
        "ShortTermInvestments",
        "AvailableForSaleSecuritiesDebtSecuritiesCurrent",
    ],
    "Accounts Receivable": [
        "AccountsReceivableNetCurrent",
        "AccountsReceivableNet",
    ],
    "Total Current Assets": [
        "AssetsCurrent",
    ],
    "Property & Equipment Net": [
        "PropertyPlantAndEquipmentNet",
    ],
    "Goodwill": [
        "Goodwill",
    ],
    "Intangible Assets": [
        "IntangibleAssetsNetExcludingGoodwill",
        "FiniteLivedIntangibleAssetsNet",
    ],
    "Total Assets": [
        "Assets",
    ],
    "Accounts Payable": [
        "AccountsPayableCurrent",
    ],
    "Short-Term Debt": [
        "ShortTermBorrowings",
        "DebtCurrent",
    ],
    "Total Current Liabilities": [
        "LiabilitiesCurrent",
    ],
    "Long-Term Debt": [
        "LongTermDebtNoncurrent",
        "LongTermDebt",
    ],
    "Total Liabilities": [
        "Liabilities",
    ],
    "Total Stockholders Equity": [
        "StockholdersEquity",
    ],
    "Retained Earnings": [
        "RetainedEarningsAccumulatedDeficit",
    ],
    "Additional Paid-In Capital": [
        "AdditionalPaidInCapital",
        "AdditionalPaidInCapitalCommonStock",
    ],
    "Treasury Stock": [
        "TreasuryStockValue",
    ],
}

CASH_FLOW_TAGS = {
    "Cash from Operations": [
        "NetCashProvidedByUsedInOperatingActivities",
    ],
    "Depreciation & Amortization": [
        "DepreciationDepletionAndAmortization",
        "DepreciationAndAmortization",
    ],
    "Stock-Based Compensation": [
        "ShareBasedCompensation",
        "StockBasedCompensation",
    ],
    "Capital Expenditures": [
        "PaymentsToAcquirePropertyPlantAndEquipment",
        "CapitalExpenditure",
    ],
    "Cash from Investing": [
        "NetCashProvidedByUsedInInvestingActivities",
    ],
    "Share Repurchases": [
        "PaymentsForRepurchaseOfCommonStock",
    ],
    "Cash from Financing": [
        "NetCashProvidedByUsedInFinancingActivities",
    ],
    "Net Change in Cash": [
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
        "CashAndCashEquivalentsPeriodIncreaseDecrease",
    ],
}


# =============================================================================
# EXTRACTION LOGIC
# =============================================================================
def fetch_company_facts():
    """Fetch all XBRL facts for PayPal from SEC EDGAR."""
    url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{SEC_CIK}.json"
    print(f"Fetching SEC EDGAR XBRL data from:\n  {url}\n")

    response = requests.get(url, headers=SEC_HEADERS, timeout=30)
    response.raise_for_status()
    data = response.json()

    # Save full response for reference
    with open(os.path.join(OUTPUT_DIR, "sec_xbrl_full.json"), "w") as f:
        json.dump(data, f)
    print(f"  ✓ Full XBRL data saved ({len(json.dumps(data)) / 1e6:.1f} MB)")

    return data


def extract_annual_value(facts_data, xbrl_tags, target_fy, taxonomy="us-gaap"):
    """
    Extract annual value for a specific fiscal year from XBRL facts.
    Tries multiple tags in order of preference.
    Filters for 10-K filings and full-year periods only.
    """
    facts = facts_data.get("facts", {}).get(taxonomy, {})

    for tag in xbrl_tags:
        if tag not in facts:
            continue

        units_data = facts[tag].get("units", {})

        # Try USD first, then USD/shares for EPS, then shares
        for unit_key in ["USD", "USD/shares", "shares"]:
            if unit_key not in units_data:
                continue

            entries = units_data[unit_key]

            for entry in entries:
                # Filter: annual 10-K filings only
                form = entry.get("form", "")
                if form != "10-K":
                    continue

                # Filter: full fiscal year (not quarterly)
                # Annual entries typically have fp="FY" or a ~365 day frame
                fp = entry.get("fp", "")
                frame = entry.get("frame", "")

                # Match by fiscal year end date
                end_date = entry.get("end", "")
                if not end_date:
                    continue

                try:
                    end_dt = datetime.strptime(end_date, "%Y-%m-%d")
                except ValueError:
                    continue

                # PayPal's FY ends Dec 31
                if end_dt.year == target_fy and end_dt.month == 12:
                    # Prefer FY period, skip quarterly
                    if fp == "FY" or "Q" not in fp:
                        return entry.get("val")

    return None


def build_statement(facts_data, tag_mapping, statement_name):
    """Build a complete financial statement for target years."""
    print(f"\n  Building {statement_name}...")
    results = {}

    for year in TARGET_YEARS:
        results[f"FY{year}"] = {}
        for item_name, tags in tag_mapping.items():
            value = extract_annual_value(facts_data, tags, year)
            results[f"FY{year}"][item_name] = value

            if value is not None:
                # Convert to millions for display (except EPS and shares)
                if "EPS" in item_name:
                    display = f"${value:.2f}"
                elif "Shares" in item_name:
                    display = f"{value/1e6:.0f}M shares"
                else:
                    display = f"${value/1e6:,.1f}M"
            else:
                display = "MISSING"

        found = sum(1 for v in results[f"FY{year}"].values() if v is not None)
        total = len(tag_mapping)
        print(f"    FY{year}: {found}/{total} items found")

    df = pd.DataFrame(results).T
    df.index.name = "Fiscal Year"
    return df


def validate_and_report(is_df, bs_df, cf_df):
    """Validate extracted data and print summary."""
    print(f"\n{'='*60}")
    print(f"  VALIDATION SUMMARY")
    print(f"{'='*60}")

    # Revenue check
    if "Total Revenue" in is_df.columns:
        print("\n  Revenue (from SEC 10-K filings):")
        for idx, row in is_df.iterrows():
            rev = row.get("Total Revenue")
            if pd.notna(rev):
                print(f"    {idx}: ${rev/1e9:.2f}B")
            else:
                print(f"    {idx}: MISSING — check SEC EDGAR manually")

    # Balance sheet check
    if all(col in bs_df.columns for col in ["Total Assets", "Total Liabilities", "Total Stockholders Equity"]):
        print("\n  Balance Sheet Check (A = L + E):")
        for idx, row in bs_df.iterrows():
            ta = row.get("Total Assets")
            tl = row.get("Total Liabilities")
            eq = row.get("Total Stockholders Equity")
            if all(pd.notna(v) for v in [ta, tl, eq]):
                diff = abs(ta - (tl + eq))
                status = "✓ BALANCED" if diff < 1e6 else f"⚠ DIFF: ${diff/1e6:.1f}M"
                print(f"    {idx}: {status}")
            else:
                print(f"    {idx}: ⚠ Missing components")

    # Operating cash flow check
    if "Cash from Operations" in cf_df.columns:
        print("\n  Operating Cash Flow:")
        for idx, row in cf_df.iterrows():
            ocf = row.get("Cash from Operations")
            if pd.notna(ocf):
                print(f"    {idx}: ${ocf/1e9:.2f}B")

    # Count missing values
    total_items = is_df.size + bs_df.size + cf_df.size
    missing = is_df.isna().sum().sum() + bs_df.isna().sum().sum() + cf_df.isna().sum().sum()
    completeness = (1 - missing / total_items) * 100
    print(f"\n  Overall Completeness: {completeness:.0f}% ({total_items - missing}/{total_items} values found)")


# =============================================================================
# MERGE WITH EXISTING YFINANCE DATA
# =============================================================================
def merge_with_yfinance(sec_is, sec_bs, sec_cf):
    """Combine SEC EDGAR (FY2019-2021) with yfinance (FY2022-2025) data."""
    print(f"\n{'='*60}")
    print(f"  MERGING SEC EDGAR + YFINANCE DATA")
    print(f"{'='*60}\n")

    for stmt_name, sec_df, yf_file in [
        ("income_statement", sec_is, "income_statement_USD_millions.csv"),
        ("balance_sheet", sec_bs, "balance_sheet_USD_millions.csv"),
        ("cash_flow", sec_cf, "cash_flow_USD_millions.csv"),
    ]:
        yf_path = os.path.join(PROCESSED_DIR, yf_file)
        if not os.path.exists(yf_path):
            print(f"  ⚠ {yf_file} not found — skipping merge for {stmt_name}")
            continue

        yf_df = pd.read_csv(yf_path, index_col=0)

        # Convert SEC data to millions to match yfinance processed format
        sec_millions = sec_df.copy()
        for col in sec_millions.columns:
            if "EPS" not in col and "Shares" not in col:
                sec_millions[col] = sec_millions[col].apply(
                    lambda x: round(x / 1e6, 2) if pd.notna(x) else None
                )
            elif "Shares" in col:
                sec_millions[col] = sec_millions[col].apply(
                    lambda x: round(x / 1e6, 2) if pd.notna(x) else None
                )

        # Rename SEC index to match yfinance format (year only)
        sec_millions.index = [idx.replace("FY", "") for idx in sec_millions.index]

        # Combine: SEC years first, then yfinance years
        sec_millions.index = sec_millions.index.astype(str)
        yf_df.index = yf_df.index.astype(str)
        combined = pd.concat([sec_millions, yf_df], axis=0)
        combined.index = combined.index.astype(str)
        combined = combined.sort_index()
        combined = combined[~combined.index.duplicated(keep='last')]  # yfinance takes priority if overlap

        output_path = os.path.join(PROCESSED_DIR, f"combined_{stmt_name}_USD_millions.csv")
        combined.to_csv(output_path)
        print(f"  ✓ {stmt_name}: {len(combined)} years merged → {output_path}")
        print(f"    Years: {', '.join(combined.index.tolist())}")

    print(f"\n  Merged files saved to: {PROCESSED_DIR}/")


# =============================================================================
# MAIN
# =============================================================================
def main():
    print(f"\n{'#'*60}")
    print(f"  SEC EDGAR XBRL EXTRACTION — PayPal (PYPL)")
    print(f"  Target: FY{TARGET_YEARS[0]}-FY{TARGET_YEARS[-1]}")
    print(f"  Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'#'*60}")

    # Step 1: Fetch all XBRL facts
    facts_data = fetch_company_facts()

    # Step 2: Build statements for missing years
    is_df = build_statement(facts_data, INCOME_STATEMENT_TAGS, "Income Statement")
    bs_df = build_statement(facts_data, BALANCE_SHEET_TAGS, "Balance Sheet")
    cf_df = build_statement(facts_data, CASH_FLOW_TAGS, "Cash Flow Statement")

    # Step 3: Save SEC-sourced data
    is_df.to_csv(os.path.join(OUTPUT_DIR, "sec_income_statement.csv"))
    bs_df.to_csv(os.path.join(OUTPUT_DIR, "sec_balance_sheet.csv"))
    cf_df.to_csv(os.path.join(OUTPUT_DIR, "sec_cash_flow.csv"))
    print(f"\n  ✓ SEC data saved to {OUTPUT_DIR}/")

    # Step 4: Validate
    validate_and_report(is_df, bs_df, cf_df)

    # Step 5: Merge with yfinance data
    merge_with_yfinance(is_df, bs_df, cf_df)

    print(f"\n{'#'*60}")
    print(f"  EXTRACTION COMPLETE")
    print(f"  Next: Cross-check key figures against 10-K filings")
    print(f"  SEC EDGAR: https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK={SEC_CIK}&type=10-K")
    print(f"{'#'*60}\n")


if __name__ == "__main__":
    main()
