"""
PayPal (PYPL) - SQL Database Loader
====================================
Creates SQLite database from schema and loads processed CSV data
into the star schema fact tables.

Run AFTER extraction scripts (01_extract, 01b, 01c).
Requires: combined_*_USD_millions.csv files in data/processed/

Usage: python 02_load_to_sql.py
"""

import sqlite3
import pandas as pd
import os
import json
from datetime import datetime

# =============================================================================
# CONFIGURATION
# =============================================================================
BASE_DIR = os.path.dirname(os.path.dirname(__file__))
SCHEMA_PATH = os.path.join(BASE_DIR, "sql", "schema.sql")
DB_PATH = os.path.join(BASE_DIR, "data", "paypal_analysis.db")
PROCESSED_DIR = os.path.join(BASE_DIR, "data", "processed")
RAW_DIR = os.path.join(BASE_DIR, "data", "raw")

# =============================================================================
# COLUMN MAPPING: CSV column names → dim_line_item.item_name
# =============================================================================
# yfinance and SEC use different column names than our schema.
# This maps them to our standardized names.

INCOME_STMT_MAP = {
    # CSV column name → our dim_line_item name
    "Total Revenue": "Total Revenue",
    "Gross Profit": "Gross Profit",
    "Cost Of Revenue": "Cost of Revenue",
    "Operating Income": "Operating Income",
    "Operating Expense": "Total Operating Expenses",
    "Interest Income": "Interest Income",
    "Interest Expense": "Interest Expense",
    "Pretax Income": "Income Before Taxes",
    "Tax Provision": "Income Tax Expense",
    "Net Income": "Net Income",
    "Basic EPS": "EPS Basic",
    "Diluted EPS": "EPS Diluted",
    "Diluted Average Shares": "Shares Outstanding (Diluted)",
    "Selling General And Administration": "General & Administrative",
    "Research And Development": "Technology & Development",
    # SEC EDGAR names (from 01b extraction)
    "EPS Basic": "EPS Basic",
    "EPS Diluted": "EPS Diluted",
    "Shares Outstanding (Diluted)": "Shares Outstanding (Diluted)",
    "Income Before Taxes": "Income Before Taxes",
    "Income Tax Expense": "Income Tax Expense",
    "Operating Expenses": "Total Operating Expenses",
}

BALANCE_SHEET_MAP = {
    "Cash & Cash Equivalents": "Cash & Cash Equivalents",
    "Cash And Cash Equivalents": "Cash & Cash Equivalents",
    "Cash Cash Equivalents And Short Term Investments": "Cash & Cash Equivalents",
    "Short-Term Investments": "Short-Term Investments",
    "Other Short Term Investments": "Short-Term Investments",
    "Accounts Receivable": "Accounts Receivable",
    "Receivables": "Accounts Receivable",
    "Total Current Assets": "Total Current Assets",
    "Current Assets": "Total Current Assets",
    "Property & Equipment Net": "Property & Equipment Net",
    "Net PPE": "Property & Equipment Net",
    "Goodwill": "Goodwill",
    "Intangible Assets": "Intangible Assets",
    "Other Intangible Assets": "Intangible Assets",
    "Total Assets": "Total Assets",
    "Accounts Payable": "Accounts Payable",
    "Payables": "Accounts Payable",
    "Short-Term Debt": "Short-Term Debt",
    "Current Debt": "Short-Term Debt",
    "Total Current Liabilities": "Total Current Liabilities",
    "Current Liabilities": "Total Current Liabilities",
    "Long-Term Debt": "Long-Term Debt",
    "Long Term Debt": "Long-Term Debt",
    "Total Liabilities": "Total Liabilities",
    "Total Liabilities Net Minority Interest": "Total Liabilities",
    "Total Stockholders Equity": "Total Stockholders Equity",
    "Stockholders Equity": "Total Stockholders Equity",
    "Common Stock Equity": "Total Stockholders Equity",
    "Retained Earnings": "Retained Earnings",
    "Retained Earnings Accumulated Deficit": "Retained Earnings",
    "Additional Paid-In Capital": "Additional Paid-In Capital",
    "Additional Paid In Capital": "Additional Paid-In Capital",
    "Treasury Stock": "Treasury Stock",
    "Treasury Shares Number": "Treasury Stock",
}

CASH_FLOW_MAP = {
    "Operating Cash Flow": "Cash from Operations",
    "Cash from Operations": "Cash from Operations",
    "Depreciation And Amortization": "Depreciation & Amortization",
    "Depreciation & Amortization": "Depreciation & Amortization",
    "Stock Based Compensation": "Stock-Based Compensation",
    "Stock-Based Compensation": "Stock-Based Compensation",
    "Capital Expenditure": "Capital Expenditures",
    "Capital Expenditures": "Capital Expenditures",
    "Investing Cash Flow": "Cash from Investing",
    "Cash from Investing": "Cash from Investing",
    "Financing Cash Flow": "Cash from Financing",
    "Cash from Financing": "Cash from Financing",
    "Repurchase Of Capital Stock": "Share Repurchases",
    "Share Repurchases": "Share Repurchases",
    "Changes In Cash": "Net Change in Cash",
    "Net Change in Cash": "Net Change in Cash",
    "Change In Working Capital": "Changes in Working Capital",
    "Changes in Working Capital": "Changes in Working Capital",
    "Free Cash Flow": "Free Cash Flow",
    "Issuance Of Debt": "Debt Issuance",
    "Repayment Of Debt": "Debt Repayment",
}


# =============================================================================
# DATABASE SETUP
# =============================================================================
def create_database():
    """Create SQLite database from schema file."""
    print(f"\n{'='*60}")
    print(f"  CREATING DATABASE")
    print(f"{'='*60}\n")

    # Remove existing DB to start fresh
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
        print(f"  Removed existing database")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    with open(SCHEMA_PATH, "r") as f:
        schema_sql = f.read()

    cursor.executescript(schema_sql)
    conn.commit()

    # Verify tables created
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = [row[0] for row in cursor.fetchall()]
    print(f"  ✓ Database created: {DB_PATH}")
    print(f"  ✓ Tables: {', '.join(tables)}")

    # Verify dimension data
    for dim_table in ["dim_period", "dim_scenario", "dim_line_item", "dim_ratio"]:
        cursor.execute(f"SELECT COUNT(*) FROM {dim_table}")
        count = cursor.fetchone()[0]
        print(f"  ✓ {dim_table}: {count} rows")

    conn.close()
    return True


# =============================================================================
# DATA LOADING
# =============================================================================
def get_period_id(cursor, fiscal_year):
    """Look up period_id for a given fiscal year."""
    cursor.execute(
        "SELECT period_id FROM dim_period WHERE fiscal_year = ? AND quarter IS NULL",
        (int(fiscal_year),)
    )
    result = cursor.fetchone()
    return result[0] if result else None


def get_line_item_id(cursor, statement_type, item_name):
    """Look up line_item_id for a given statement type and item name."""
    cursor.execute(
        "SELECT line_item_id FROM dim_line_item WHERE statement_type = ? AND item_name = ?",
        (statement_type, item_name)
    )
    result = cursor.fetchone()
    return result[0] if result else None


def get_scenario_id(cursor, scenario_name="Actual"):
    """Look up scenario_id."""
    cursor.execute(
        "SELECT scenario_id FROM dim_scenario WHERE scenario_name = ?",
        (scenario_name,)
    )
    result = cursor.fetchone()
    return result[0] if result else None


def load_statement(conn, csv_path, column_map, statement_type):
    """Load a financial statement CSV into fact_financials."""
    if not os.path.exists(csv_path):
        print(f"  ⚠ File not found: {csv_path}")
        return 0

    df = pd.read_csv(csv_path, index_col=0)
    df.index = df.index.astype(str)
    cursor = conn.cursor()
    actual_id = get_scenario_id(cursor, "Actual")
    loaded = 0
    skipped = 0
    unmapped = set()

    for year_str in df.index:
        # Extract just the year (handle "2024-12-31" or "2024" formats)
        year = int(year_str[:4])
        period_id = get_period_id(cursor, year)
        if period_id is None:
            continue

        for csv_col in df.columns:
            # Find mapping
            std_name = column_map.get(csv_col)
            if std_name is None:
                unmapped.add(csv_col)
                continue

            line_item_id = get_line_item_id(cursor, statement_type, std_name)
            if line_item_id is None:
                continue

            value = df.loc[year_str, csv_col]
            if pd.isna(value):
                continue

            try:
                cursor.execute(
                    """INSERT OR REPLACE INTO fact_financials 
                       (period_id, line_item_id, scenario_id, amount, source)
                       VALUES (?, ?, ?, ?, ?)""",
                    (period_id, line_item_id, actual_id, float(value),
                     f"{'10-K' if year <= 2021 else 'yfinance'} FY{year}")
                )
                loaded += 1
            except Exception as e:
                print(f"    Error loading {std_name} FY{year}: {e}")
                skipped += 1

    conn.commit()
    print(f"    Loaded: {loaded} | Skipped: {skipped}")
    if unmapped:
        print(f"    Unmapped columns (not in schema, OK to ignore): {len(unmapped)}")
    return loaded


def load_stock_prices(conn):
    """Load historical stock price data."""
    csv_path = os.path.join(RAW_DIR, "stock_prices.csv")
    if not os.path.exists(csv_path):
        print(f"  ⚠ Stock prices file not found")
        return 0

    df = pd.read_csv(csv_path)
    cursor = conn.cursor()
    loaded = 0

    # Identify date and price columns
    date_col = df.columns[0]  # Usually 'Date' or index

    for _, row in df.iterrows():
        try:
            trade_date = str(row[date_col])[:10]  # YYYY-MM-DD
            cursor.execute(
                """INSERT OR REPLACE INTO fact_stock_price
                   (trade_date, open_price, high_price, low_price, close_price, volume)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (trade_date,
                 row.get("Open"), row.get("High"), row.get("Low"),
                 row.get("Close"), row.get("Volume"))
            )
            loaded += 1
        except Exception:
            continue

    conn.commit()
    return loaded


def calculate_ratios(conn):
    """Calculate financial ratios from loaded statement data."""
    print(f"\n  Calculating financial ratios...")
    cursor = conn.cursor()
    actual_id = get_scenario_id(cursor, "Actual")

    # Get all periods
    cursor.execute("SELECT period_id, fiscal_year FROM dim_period WHERE quarter IS NULL")
    periods = cursor.fetchall()

    def get_value(period_id, statement_type, item_name):
        cursor.execute(
            """SELECT ff.amount FROM fact_financials ff
               JOIN dim_line_item dli ON ff.line_item_id = dli.line_item_id
               WHERE ff.period_id = ? AND ff.scenario_id = ?
               AND dli.statement_type = ? AND dli.item_name = ?""",
            (period_id, actual_id, statement_type, item_name)
        )
        result = cursor.fetchone()
        return result[0] if result else None

    def get_ratio_id(ratio_name):
        cursor.execute("SELECT ratio_id FROM dim_ratio WHERE ratio_name = ?", (ratio_name,))
        result = cursor.fetchone()
        return result[0] if result else None

    def insert_ratio(period_id, ratio_name, value):
        ratio_id = get_ratio_id(ratio_name)
        if ratio_id and value is not None:
            cursor.execute(
                """INSERT OR REPLACE INTO fact_ratios (period_id, ratio_id, scenario_id, value)
                   VALUES (?, ?, ?, ?)""",
                (period_id, ratio_id, actual_id, value)
            )
            return True
        return False

    ratios_loaded = 0
    prev_revenue = None

    for period_id, fy in sorted(periods, key=lambda x: x[1]):
        rev = get_value(period_id, "income_statement", "Total Revenue")
        gp = get_value(period_id, "income_statement", "Gross Profit")
        oi = get_value(period_id, "income_statement", "Operating Income")
        ni = get_value(period_id, "income_statement", "Net Income")
        ta = get_value(period_id, "balance_sheet", "Total Assets")
        eq = get_value(period_id, "balance_sheet", "Total Stockholders Equity")
        tca = get_value(period_id, "balance_sheet", "Total Current Assets")
        tcl = get_value(period_id, "balance_sheet", "Total Current Liabilities")
        cash = get_value(period_id, "balance_sheet", "Cash & Cash Equivalents")
        ltd = get_value(period_id, "balance_sheet", "Long-Term Debt")
        std = get_value(period_id, "balance_sheet", "Short-Term Debt")
        ie = get_value(period_id, "income_statement", "Interest Expense")
        da = get_value(period_id, "cash_flow", "Depreciation & Amortization")
        cfo = get_value(period_id, "cash_flow", "Cash from Operations")
        capex = get_value(period_id, "cash_flow", "Capital Expenditures")

        # Profitability ratios
        if rev and rev != 0:
            if gp is not None:
                if insert_ratio(period_id, "Gross Margin", gp / rev * 100): ratios_loaded += 1
            if oi is not None:
                if insert_ratio(period_id, "Operating Margin", oi / rev * 100): ratios_loaded += 1
            if ni is not None:
                if insert_ratio(period_id, "Net Margin", ni / rev * 100): ratios_loaded += 1
            if oi is not None and da is not None:
                ebitda = oi + abs(da)
                if insert_ratio(period_id, "EBITDA Margin", ebitda / rev * 100): ratios_loaded += 1

        if ni and eq and eq != 0:
            if insert_ratio(period_id, "ROE", ni / eq * 100): ratios_loaded += 1
        if ni and ta and ta != 0:
            if insert_ratio(period_id, "ROA", ni / ta * 100): ratios_loaded += 1

        # Liquidity ratios
        if tca and tcl and tcl != 0:
            if insert_ratio(period_id, "Current Ratio", tca / tcl): ratios_loaded += 1

        # Leverage ratios
        total_debt = (ltd or 0) + (std or 0)
        if total_debt > 0 and eq and eq != 0:
            if insert_ratio(period_id, "Debt to Equity", total_debt / eq): ratios_loaded += 1
        if ta and ta != 0:
            if insert_ratio(period_id, "Debt to Assets", total_debt / ta): ratios_loaded += 1
        if ie and ie != 0 and oi:
            if insert_ratio(period_id, "Interest Coverage", abs(oi / ie)): ratios_loaded += 1

        # Efficiency
        if rev and ta and ta != 0:
            if insert_ratio(period_id, "Asset Turnover", rev / ta): ratios_loaded += 1

        # Growth (needs previous year)
        if prev_revenue and prev_revenue != 0 and rev:
            if insert_ratio(period_id, "Revenue Growth", (rev - prev_revenue) / abs(prev_revenue) * 100):
                ratios_loaded += 1
        prev_revenue = rev

        # FCF-related
        if cfo and capex:
            fcf = cfo - abs(capex)
            if insert_ratio(period_id, "FCF Yield", fcf / rev * 100 if rev else None): ratios_loaded += 1

    conn.commit()
    print(f"  ✓ {ratios_loaded} ratio values calculated and loaded")
    return ratios_loaded


# =============================================================================
# SUMMARY REPORT
# =============================================================================
def print_summary(conn):
    """Print database summary."""
    cursor = conn.cursor()

    print(f"\n{'='*60}")
    print(f"  DATABASE SUMMARY")
    print(f"{'='*60}\n")

    for table in ["fact_financials", "fact_ratios", "fact_stock_price"]:
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        count = cursor.fetchone()[0]
        print(f"  {table}: {count} rows")

    # Show revenue by year as sanity check
    print(f"\n  Revenue by Year (sanity check):")
    cursor.execute("""
        SELECT dp.period_label, ff.amount, ff.source
        FROM fact_financials ff
        JOIN dim_period dp ON ff.period_id = dp.period_id
        JOIN dim_line_item dli ON ff.line_item_id = dli.line_item_id
        JOIN dim_scenario ds ON ff.scenario_id = ds.scenario_id
        WHERE dli.item_name = 'Total Revenue' AND ds.scenario_name = 'Actual'
        ORDER BY dp.fiscal_year
    """)
    for row in cursor.fetchall():
        print(f"    {row[0]}: ${row[1]:,.0f}M (source: {row[2]})")

    # Show key ratios for most recent year
    print(f"\n  Key Ratios (latest year):")
    cursor.execute("""
        SELECT dr.ratio_name, fr.value, dr.format_type
        FROM fact_ratios fr
        JOIN dim_ratio dr ON fr.ratio_id = dr.ratio_id
        JOIN dim_period dp ON fr.period_id = dp.period_id
        WHERE dp.fiscal_year = (SELECT MAX(fiscal_year) FROM dim_period WHERE period_type = 'actual')
        ORDER BY dr.ratio_category, dr.ratio_name
    """)
    for row in cursor.fetchall():
        if row[2] == "percentage":
            print(f"    {row[0]}: {row[1]:.1f}%")
        elif row[2] == "multiple":
            print(f"    {row[0]}: {row[1]:.1f}x")
        else:
            print(f"    {row[0]}: {row[1]:.2f}")


# =============================================================================
# MAIN
# =============================================================================
def main():
    print(f"\n{'#'*60}")
    print(f"  PAYPAL (PYPL) - DATABASE LOADER")
    print(f"  Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'#'*60}")

    # Step 1: Create database
    create_database()

    # Step 2: Load financial statements
    conn = sqlite3.connect(DB_PATH)

    print(f"\n{'='*60}")
    print(f"  LOADING FINANCIAL STATEMENTS")
    print(f"{'='*60}")

    print(f"\n  [1/4] Income Statement...")
    load_statement(
        conn,
        os.path.join(PROCESSED_DIR, "combined_income_statement_USD_millions.csv"),
        INCOME_STMT_MAP,
        "income_statement"
    )

    print(f"\n  [2/4] Balance Sheet...")
    load_statement(
        conn,
        os.path.join(PROCESSED_DIR, "combined_balance_sheet_USD_millions.csv"),
        BALANCE_SHEET_MAP,
        "balance_sheet"
    )

    print(f"\n  [3/4] Cash Flow Statement...")
    load_statement(
        conn,
        os.path.join(PROCESSED_DIR, "combined_cash_flow_USD_millions.csv"),
        CASH_FLOW_MAP,
        "cash_flow"
    )

    print(f"\n  [4/4] Stock Prices...")
    loaded = load_stock_prices(conn)
    print(f"    Loaded: {loaded} trading days")

    # Step 3: Calculate ratios
    calculate_ratios(conn)

    # Step 4: Summary
    print_summary(conn)

    conn.close()

    print(f"\n{'#'*60}")
    print(f"  DATABASE READY")
    print(f"  File: {DB_PATH}")
    print(f"  Open in SQLite Browser to verify")
    print(f"  Next step: Build Excel financial model")
    print(f"{'#'*60}\n")


if __name__ == "__main__":
    main()
