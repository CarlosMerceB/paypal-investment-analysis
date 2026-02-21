"""
Patch FY2021 Balance Sheet gaps with data from PayPal 10-K filing.
Source: PayPal FY2021 10-K, Consolidated Balance Sheet (SEC EDGAR)

Run from scripts/ folder AFTER 01b_extract_sec_edgar.py
"""

import pandas as pd
import os

PROCESSED_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "processed")
BS_FILE = os.path.join(PROCESSED_DIR, "combined_balance_sheet_USD_millions.csv")

# FY2021 data from 10-K (values in USD Millions)
FY2021_PATCH = {
    "Cash & Cash Equivalents": 5197.0,
    "Short-Term Investments": 109.0,
    "Accounts Receivable": 12723.0,
    "Total Current Assets": 18029.0,
    "Property & Equipment Net": 1909.0,
    "Goodwill": 11454.0,
    "Intangible Assets": 1332.0,
    "Total Assets": 75803.0,
    "Accounts Payable": 197.0,
    "Short-Term Debt": 0.0,
    "Total Current Liabilities": 43029.0,
    "Long-Term Debt": 8049.0,
    "Total Liabilities": 54076.0,
    "Total Stockholders Equity": 21727.0,
    "Retained Earnings": 16535.0,
    # Additional items from the 10-K notes
    "Current Accrued Expenses": 3755.0,
    "Income Tax Payable": 236.0,
    "Non Current Deferred Taxes Liabilities": 2998.0,
}

print(f"Loading: {BS_FILE}")
df = pd.read_csv(BS_FILE, index_col=0)
df.index = df.index.astype(str)

patched = 0
skipped = 0
for col_name, value in FY2021_PATCH.items():
    if col_name in df.columns:
        old_val = df.loc["2021", col_name]
        df.loc["2021", col_name] = value
        status = "FILLED" if pd.isna(old_val) else f"UPDATED ({old_val} → {value})"
        print(f"  ✓ {col_name}: ${value:,.0f}M [{status}]")
        patched += 1
    else:
        print(f"  ⚠ Column '{col_name}' not found in CSV — skipped")
        skipped += 1

df.to_csv(BS_FILE)
print(f"\n  Patched: {patched} values | Skipped: {skipped}")

# Verify balance
ta = df.loc["2021", "Total Assets"]
tl = df.loc["2021", "Total Liabilities"]
eq = df.loc["2021", "Total Stockholders Equity"]
if all(pd.notna(v) for v in [ta, tl, eq]):
    diff = abs(ta - (tl + eq))
    print(f"  Balance check FY2021: Assets=${ta:,.0f}M = Liab=${tl:,.0f}M + Equity=${eq:,.0f}M (diff: ${diff:,.0f}M)")
    print(f"  {'✓ BALANCED' if diff < 1 else '⚠ IMBALANCE'}")

print(f"\n  Saved to: {BS_FILE}")
