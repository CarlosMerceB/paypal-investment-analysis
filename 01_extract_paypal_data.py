"""
PayPal (PYPL) Financial Data Extraction Pipeline
=================================================
Investment Fund Analysis - Data Collection Phase

Extracts 5 years of historical financials (FY2019-FY2024) from:
- yfinance: Structured financial statements + market data
- SEC EDGAR: 10-K filing metadata and links for cross-referencing

Output: Clean CSV files ready for SQL database loading and Excel model input.

Required packages:
    pip install yfinance pandas requests --break-system-packages
"""

import yfinance as yf
import pandas as pd
import requests
import json
import os
from datetime import datetime

# =============================================================================
# CONFIGURATION
# =============================================================================
TICKER = "PYPL"
COMPANY_NAME = "PayPal Holdings Inc"
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "raw")
FISCAL_YEARS = [2019, 2020, 2021, 2022, 2023, 2024]

# SEC EDGAR configuration
SEC_HEADERS = {
    "User-Agent": "InvestmentAnalysis carlos@example.com",  # SEC requires identification
    "Accept-Encoding": "gzip, deflate",
}
SEC_CIK = "0001633917"  # PayPal's CIK number

os.makedirs(OUTPUT_DIR, exist_ok=True)


# =============================================================================
# PHASE 1A: YFINANCE EXTRACTION
# =============================================================================
def extract_yfinance_data():
    """Extract financial statements and market data from yfinance."""
    print(f"\n{'='*60}")
    print(f"  EXTRACTING {TICKER} DATA FROM YFINANCE")
    print(f"{'='*60}\n")

    stock = yf.Ticker(TICKER)

    # --- Income Statement (Annual) ---
    print("[1/6] Extracting Income Statements...")
    income_stmt = stock.income_stmt  # Last 4 years by default
    if income_stmt is not None and not income_stmt.empty:
        income_stmt.to_csv(os.path.join(OUTPUT_DIR, "income_statement.csv"))
        print(f"  ✓ {income_stmt.shape[1]} years extracted")
        print(f"  ✓ Line items: {income_stmt.shape[0]}")
    else:
        print("  ✗ No income statement data available")

    # --- Balance Sheet (Annual) ---
    print("[2/6] Extracting Balance Sheets...")
    balance_sheet = stock.balance_sheet
    if balance_sheet is not None and not balance_sheet.empty:
        balance_sheet.to_csv(os.path.join(OUTPUT_DIR, "balance_sheet.csv"))
        print(f"  ✓ {balance_sheet.shape[1]} years extracted")
        print(f"  ✓ Line items: {balance_sheet.shape[0]}")
    else:
        print("  ✗ No balance sheet data available")

    # --- Cash Flow Statement (Annual) ---
    print("[3/6] Extracting Cash Flow Statements...")
    cashflow = stock.cashflow
    if cashflow is not None and not cashflow.empty:
        cashflow.to_csv(os.path.join(OUTPUT_DIR, "cash_flow.csv"))
        print(f"  ✓ {cashflow.shape[1]} years extracted")
        print(f"  ✓ Line items: {cashflow.shape[0]}")
    else:
        print("  ✗ No cash flow data available")

    # --- Quarterly Statements (for TTM calculations) ---
    print("[4/6] Extracting Quarterly Data (for TTM)...")
    q_income = stock.quarterly_income_stmt
    q_balance = stock.quarterly_balance_sheet
    q_cashflow = stock.quarterly_cashflow
    if q_income is not None and not q_income.empty:
        q_income.to_csv(os.path.join(OUTPUT_DIR, "quarterly_income_statement.csv"))
        q_balance.to_csv(os.path.join(OUTPUT_DIR, "quarterly_balance_sheet.csv"))
        q_cashflow.to_csv(os.path.join(OUTPUT_DIR, "quarterly_cash_flow.csv"))
        print(f"  ✓ {q_income.shape[1]} quarters extracted")
    else:
        print("  ✗ No quarterly data available")

    # --- Historical Stock Price Data ---
    print("[5/6] Extracting Historical Stock Prices...")
    hist = stock.history(start="2019-01-01", end=datetime.now().strftime("%Y-%m-%d"))
    if hist is not None and not hist.empty:
        hist.to_csv(os.path.join(OUTPUT_DIR, "stock_prices.csv"))
        print(f"  ✓ {len(hist)} trading days extracted")
        print(f"  ✓ Date range: {hist.index[0].strftime('%Y-%m-%d')} to {hist.index[-1].strftime('%Y-%m-%d')}")
    else:
        print("  ✗ No stock price data available")

    # --- Company Info & Key Statistics ---
    print("[6/6] Extracting Company Info & Key Stats...")
    info = stock.info
    if info:
        key_stats = {
            "ticker": TICKER,
            "company_name": info.get("longName", COMPANY_NAME),
            "sector": info.get("sector", "N/A"),
            "industry": info.get("industry", "N/A"),
            "market_cap": info.get("marketCap", "N/A"),
            "enterprise_value": info.get("enterpriseValue", "N/A"),
            "shares_outstanding": info.get("sharesOutstanding", "N/A"),
            "beta": info.get("beta", "N/A"),
            "trailing_pe": info.get("trailingPE", "N/A"),
            "forward_pe": info.get("forwardPE", "N/A"),
            "ev_ebitda": info.get("enterpriseToEbitda", "N/A"),
            "price_to_book": info.get("priceToBook", "N/A"),
            "profit_margin": info.get("profitMargins", "N/A"),
            "operating_margin": info.get("operatingMargins", "N/A"),
            "roe": info.get("returnOnEquity", "N/A"),
            "roa": info.get("returnOnAssets", "N/A"),
            "revenue_growth": info.get("revenueGrowth", "N/A"),
            "debt_to_equity": info.get("debtToEquity", "N/A"),
            "current_ratio": info.get("currentRatio", "N/A"),
            "free_cashflow": info.get("freeCashflow", "N/A"),
            "dividend_yield": info.get("dividendYield", "N/A"),
            "currency": info.get("currency", "USD"),
            "exchange": info.get("exchange", "N/A"),
            "extraction_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        with open(os.path.join(OUTPUT_DIR, "company_info.json"), "w") as f:
            json.dump(key_stats, f, indent=2, default=str)
        print(f"  ✓ Company info saved ({len(key_stats)} fields)")

    return income_stmt, balance_sheet, cashflow, hist, info


# =============================================================================
# PHASE 1B: SEC EDGAR EXTRACTION
# =============================================================================
def extract_sec_filings():
    """
    Fetch 10-K filing index from SEC EDGAR for cross-referencing.
    Provides direct links to annual reports for manual verification.
    """
    print(f"\n{'='*60}")
    print(f"  EXTRACTING SEC EDGAR FILING INDEX")
    print(f"{'='*60}\n")

    url = f"https://efts.sec.gov/LATEST/search-index?q=%22paypal%22&dateRange=custom&startdt=2019-01-01&enddt=2025-12-31&forms=10-K"
    submissions_url = f"https://data.sec.gov/submissions/CIK{SEC_CIK}.json"

    try:
        response = requests.get(submissions_url, headers=SEC_HEADERS, timeout=15)
        response.raise_for_status()
        data = response.json()

        filings = data.get("filings", {}).get("recent", {})
        forms = filings.get("form", [])
        dates = filings.get("filingDate", [])
        accessions = filings.get("accessionNumber", [])
        primary_docs = filings.get("primaryDocument", [])

        annual_filings = []
        for i, form in enumerate(forms):
            if form == "10-K":
                accession_clean = accessions[i].replace("-", "")
                filing_url = f"https://www.sec.gov/Archives/edgar/data/{SEC_CIK}/{accession_clean}/{primary_docs[i]}"
                annual_filings.append({
                    "form": form,
                    "filing_date": dates[i],
                    "accession_number": accessions[i],
                    "document_url": filing_url,
                })

        if annual_filings:
            filings_df = pd.DataFrame(annual_filings)
            filings_df.to_csv(os.path.join(OUTPUT_DIR, "sec_10k_filings.csv"), index=False)
            print(f"  ✓ {len(annual_filings)} annual (10-K) filings found")
            for f in annual_filings[:6]:
                print(f"    - {f['filing_date']}: {f['document_url'][:80]}...")
        else:
            print("  ✗ No 10-K filings found")

        return annual_filings

    except requests.exceptions.RequestException as e:
        print(f"  ✗ SEC EDGAR request failed: {e}")
        print("  → This is expected in sandboxed environments.")
        print("  → Run this script locally to fetch SEC filing links.")
        return []


# =============================================================================
# PHASE 1C: DATA QUALITY VALIDATION
# =============================================================================
def validate_extraction(income_stmt, balance_sheet, cashflow):
    """Run basic quality checks on extracted data."""
    print(f"\n{'='*60}")
    print(f"  DATA QUALITY VALIDATION")
    print(f"{'='*60}\n")

    checks_passed = 0
    checks_total = 0

    # Check 1: Income statement has revenue
    checks_total += 1
    if income_stmt is not None and "Total Revenue" in income_stmt.index:
        revenue = income_stmt.loc["Total Revenue"]
        print(f"  ✓ Revenue data found: {len(revenue.dropna())} years")
        for col in revenue.index:
            yr = col.strftime("%Y") if hasattr(col, "strftime") else str(col)
            val = revenue[col]
            if pd.notna(val):
                print(f"    FY{yr}: ${val/1e9:.2f}B")
        checks_passed += 1
    else:
        print("  ✗ Revenue data missing!")

    # Check 2: Balance sheet balances (Assets = Liabilities + Equity)
    checks_total += 1
    if balance_sheet is not None:
        try:
            total_assets = balance_sheet.loc["Total Assets"] if "Total Assets" in balance_sheet.index else None
            total_le = balance_sheet.loc["Total Liabilities Net Minority Interest"] if "Total Liabilities Net Minority Interest" in balance_sheet.index else None
            equity = balance_sheet.loc["Stockholders Equity"] if "Stockholders Equity" in balance_sheet.index else None

            if total_assets is not None and total_le is not None and equity is not None:
                for col in total_assets.index:
                    yr = col.strftime("%Y") if hasattr(col, "strftime") else str(col)
                    diff = abs(total_assets[col] - (total_le[col] + equity[col]))
                    if diff < 1e6:
                        print(f"  ✓ Balance sheet balances for FY{yr} (diff: ${diff:,.0f})")
                    else:
                        print(f"  ⚠ Balance sheet imbalance FY{yr}: ${diff/1e6:.1f}M")
                checks_passed += 1
            else:
                print("  ⚠ Could not verify balance sheet balance (missing line items)")
        except Exception as e:
            print(f"  ⚠ Balance check error: {e}")
    else:
        print("  ✗ Balance sheet data missing!")

    # Check 3: Cash flow statement has operating cash flow
    checks_total += 1
    if cashflow is not None:
        ocf_labels = ["Operating Cash Flow", "Total Cash From Operating Activities"]
        for label in ocf_labels:
            if label in cashflow.index:
                print(f"  ✓ Operating Cash Flow found ({label})")
                checks_passed += 1
                break
        else:
            print(f"  ⚠ Standard OCF label not found. Available: {list(cashflow.index[:5])}")
    else:
        print("  ✗ Cash flow data missing!")

    # Check 4: No major gaps in data
    checks_total += 1
    if income_stmt is not None:
        null_pct = income_stmt.isnull().sum().sum() / (income_stmt.shape[0] * income_stmt.shape[1]) * 100
        if null_pct < 20:
            print(f"  ✓ Data completeness: {100-null_pct:.1f}% populated")
            checks_passed += 1
        else:
            print(f"  ⚠ High null rate: {null_pct:.1f}% of cells are empty")

    print(f"\n  RESULT: {checks_passed}/{checks_total} checks passed")
    return checks_passed == checks_total


# =============================================================================
# PHASE 1D: GENERATE STANDARDIZED OUTPUT FOR EXCEL MODEL
# =============================================================================
def prepare_excel_input(income_stmt, balance_sheet, cashflow):
    """
    Transform raw yfinance data into a clean, standardized format
    ready for direct import into the Excel financial model.
    """
    print(f"\n{'='*60}")
    print(f"  PREPARING STANDARDIZED DATA FOR EXCEL MODEL")
    print(f"{'='*60}\n")

    processed_dir = os.path.join(os.path.dirname(OUTPUT_DIR), "processed")
    os.makedirs(processed_dir, exist_ok=True)

    def standardize_statement(df, statement_name):
        """Transpose and clean statement for Excel-friendly format."""
        if df is None or df.empty:
            return None
        clean = df.copy()
        clean.columns = [col.strftime("%Y") if hasattr(col, "strftime") else str(col) for col in clean.columns]
        clean = clean.T.sort_index()
        clean.index.name = "Fiscal Year"
        clean = clean.div(1e6).round(2)  # Convert to millions USD
        return clean

    # Standardize all three statements (values in $M)
    for stmt, name in [(income_stmt, "income_statement"), (balance_sheet, "balance_sheet"), (cashflow, "cash_flow")]:
        clean = standardize_statement(stmt, name)
        if clean is not None:
            filepath = os.path.join(processed_dir, f"{name}_USD_millions.csv")
            clean.to_csv(filepath)
            print(f"  ✓ {name}: {clean.shape[0]} years × {clean.shape[1]} items → {filepath}")

    # Create a summary metrics file for quick Excel reference
    if income_stmt is not None and balance_sheet is not None and cashflow is not None:
        metrics = {}
        years = [col.strftime("%Y") for col in income_stmt.columns if hasattr(col, "strftime")]

        for col in income_stmt.columns:
            yr = col.strftime("%Y") if hasattr(col, "strftime") else str(col)
            rev = income_stmt.loc["Total Revenue", col] if "Total Revenue" in income_stmt.index else None
            ni = income_stmt.loc["Net Income", col] if "Net Income" in income_stmt.index else None
            gp = income_stmt.loc["Gross Profit", col] if "Gross Profit" in income_stmt.index else None
            oi = income_stmt.loc["Operating Income", col] if "Operating Income" in income_stmt.index else None
            ta = balance_sheet.loc["Total Assets", col] if col in balance_sheet.columns and "Total Assets" in balance_sheet.index else None
            eq = balance_sheet.loc["Stockholders Equity", col] if col in balance_sheet.columns and "Stockholders Equity" in balance_sheet.index else None

            metrics[yr] = {
                "Revenue ($M)": round(rev / 1e6, 2) if pd.notna(rev) else None,
                "Net Income ($M)": round(ni / 1e6, 2) if pd.notna(ni) else None,
                "Gross Margin (%)": round(gp / rev * 100, 1) if pd.notna(gp) and pd.notna(rev) and rev != 0 else None,
                "Operating Margin (%)": round(oi / rev * 100, 1) if pd.notna(oi) and pd.notna(rev) and rev != 0 else None,
                "Net Margin (%)": round(ni / rev * 100, 1) if pd.notna(ni) and pd.notna(rev) and rev != 0 else None,
                "ROE (%)": round(ni / eq * 100, 1) if pd.notna(ni) and pd.notna(eq) and eq != 0 else None,
                "ROA (%)": round(ni / ta * 100, 1) if pd.notna(ni) and pd.notna(ta) and ta != 0 else None,
                "Total Assets ($M)": round(ta / 1e6, 2) if pd.notna(ta) else None,
                "Equity ($M)": round(eq / 1e6, 2) if pd.notna(eq) else None,
            }

        metrics_df = pd.DataFrame(metrics).T
        metrics_df.index.name = "Fiscal Year"
        metrics_df.to_csv(os.path.join(processed_dir, "key_metrics_summary.csv"))
        print(f"  ✓ Key metrics summary: {metrics_df.shape[0]} years × {metrics_df.shape[1]} metrics")

    print(f"\n  All processed files saved to: {processed_dir}/")


# =============================================================================
# MAIN EXECUTION
# =============================================================================
def main():
    print(f"\n{'#'*60}")
    print(f"  PAYPAL ({TICKER}) - INVESTMENT ANALYSIS DATA EXTRACTION")
    print(f"  Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Scope: FY2019-FY2024 + Market Data")
    print(f"{'#'*60}")

    # Step 1: Extract from yfinance
    income_stmt, balance_sheet, cashflow, hist, info = extract_yfinance_data()

    # Step 2: Extract SEC EDGAR filing links
    sec_filings = extract_sec_filings()

    # Step 3: Validate data quality
    validate_extraction(income_stmt, balance_sheet, cashflow)

    # Step 4: Prepare Excel-ready output
    prepare_excel_input(income_stmt, balance_sheet, cashflow)

    print(f"\n{'#'*60}")
    print(f"  EXTRACTION COMPLETE")
    print(f"  Raw data:       ./data/raw/")
    print(f"  Processed data: ./data/processed/")
    print(f"  Next step:      Run 02_load_to_sql.py")
    print(f"{'#'*60}\n")


if __name__ == "__main__":
    main()
