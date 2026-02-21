# PayPal (PYPL) — Investment Analysis & Financial Model

End-to-end equity research project: data extraction, 3-statement financial model, DCF valuation, scenario analysis, and interactive dashboard.

Built as a portfolio demonstration of financial modeling, Python automation, SQL data management, and business intelligence skills.

## Project Overview

PayPal is at an inflection point. After Q4 2025 earnings missed expectations, the board replaced CEO Alex Chriss with HP's Enrique Lores. The stock dropped ~50% to $40, yet the company generates $7B+ in annual free cash flow and is aggressively buying back shares. This project asks: **what is PayPal actually worth?**

**Key output:** Probability-weighted target price of **~$68**, derived from Bull ($139), Base ($90), and Bear ($31) DCF scenarios — aligning with analyst consensus (~$62) and implying ~69% upside from current levels.

## Project Structure

```
paypal-investment-analysis/
│
├── data/                           # Raw financial data
│   ├── annual_financials.csv       # Income Statement (FY2019-2025)
│   ├── annual_balance_sheet.csv    # Balance Sheet (FY2019-2025)
│   ├── annual_cash_flow.csv        # Cash Flow Statement (FY2019-2025)
│   └── stock_data.csv              # Historical price data
│
├── scripts/                        # Python automation
│   ├── 01_extract_data.py          # yfinance API → CSV extraction
│   ├── 02_load_to_sql.py           # CSV → SQL Server database
│   ├── 03_build_excel_model.py     # Cover + Assumptions tabs
│   ├── 03b_build_income_statement.py
│   ├── 03c_build_balance_sheet.py
│   ├── 03d_build_cash_flow.py
│   ├── 03e_build_ratios_dcf.py     # Financial Ratios + DCF valuation
│   └── 03f_build_scenarios_memo.py # Scenario Analysis + Investment Memo
│
├── model/
│   └── PYPL_Financial_Model.xlsx   # Complete 8-tab financial model
│
├── sql/
│   └── schema.sql                  # Database schema definition
│
├── .gitignore
└── README.md
```

## Financial Model (Excel)

The model contains 8 fully linked tabs with industry-standard formatting (blue = inputs, black = formulas, green = cross-sheet references, yellow = key assumptions).

| Tab | Description |
|-----|-------------|
| **Cover** | Navigation page with formatting legend |
| **Assumptions** | All drivers in one place — revenue growth, margins, CapEx, WACC components |
| **Income Statement** | Revenue through EPS, fully driven by Assumptions tab |
| **Balance Sheet** | Assets, liabilities, equity with balance check = $0 |
| **Cash Flow** | CFO, CFI, CFF with FCF calculation; closes the 3-statement loop |
| **Ratios** | Profitability, returns, leverage, efficiency, and growth metrics |
| **DCF** | WACC via CAPM, 3-year FCF projection, Gordon Growth terminal value, sensitivity table |
| **Scenarios** | Bull/Base/Bear with probability-weighted target price |
| **Investment Memo** | Equity research format: thesis, risks, catalysts, recommendation |

### 3-Statement Loop

The model is fully circular: Net Income flows from the Income Statement to both the Cash Flow Statement (starting point for CFO) and the Balance Sheet (Retained Earnings). Working capital changes from the Balance Sheet drive operating cash flow adjustments. Ending Cash from the Cash Flow Statement links back to the Balance Sheet, closing the loop. A change to any single assumption cascades through all three statements.

### Key Assumptions (Post-Q4 2025 Update)

The model was updated mid-project after PayPal's Q4 2025 earnings miss:

| Metric | Pre-Q4 | Post-Q4 | Rationale |
|--------|--------|---------|-----------|
| Revenue Growth '26E | 7.1% | 3.0% | Management guided flat total margins, weak branded checkout |
| Operating Margin '26E | 19.5% | 18.0% | No operating leverage during CEO transition |
| Beta | 1.25 | 1.40 | Elevated: CEO change, competitive pressure, guidance uncertainty |
| Terminal Growth Rate | 2.5% | 1.5% | Conservative: below GDP given competitive risk |
| CapEx | ~$530M | ~$900M | Management guided ~$1B investment cycle |

## Scenario Analysis

| Scenario | Probability | Revenue Growth | Op. Margin '28E | FCF '28E | WACC | TGR | Implied Price |
|----------|-------------|----------------|-----------------|----------|------|-----|---------------|
| **Bull** | 15% | 6.0% avg | 22.0% | $8.5B | 8.5% | 2.5% | ~$139 |
| **Base** | 35% | 4.0% avg | 19.5% | ~$7.7B | 9.9% | 1.5% | ~$90 |
| **Bear** | 50% | 1.5% avg | 15.5% | $3.5B | 11.5% | 1.0% | ~$31 |
| **Weighted** | 100% | — | — | — | — | — | **~$68** |

The 50% Bear Case weight reflects genuine uncertainty: branded checkout erosion, third CEO in three years, and withdrawn 2027 guidance.

## Valuation Methodology

**Discounted Cash Flow (FCFF)** with 3-year explicit forecast (FY2026E–2028E) and Gordon Growth terminal value.

- **WACC:** 9.9% (base case). Cost of Equity via CAPM: Rf 4.2% + Beta 1.40 × ERP 5.5% = 11.9%. After-tax Cost of Debt: 3.5%. Capital weights: 70% equity / 30% debt.
- **Terminal Value:** FCF₂₀₂₈ × (1 + 1.5%) / (9.9% − 1.5%). Terminal value represents ~81% of base case enterprise value — typical for a mature company.
- **Equity Bridge:** Enterprise Value − Net Debt ($1.9B) = Equity Value ÷ Diluted Shares = Implied Share Price.

## Technology Stack

| Tool | Use Case |
|------|----------|
| **Python** | Data extraction (yfinance API), SQL loading, Excel model generation (openpyxl) |
| **SQL Server** | Structured storage of 7 years of financial data |
| **Excel** | 3-statement financial model with live formulas and cross-sheet linking |
| **Power BI** | Executive dashboard *(Phase 4 — in progress)* |

## Data Sources

- SEC EDGAR: 10-K annual filings (FY2019–2025)
- yfinance API: Historical financial data and stock prices
- PayPal Q4 2025 Earnings Release (February 3, 2026)
- Analyst consensus estimates via public financial data providers

## What I Learned

Building this project taught me that **assumptions drive 90% of the valuation, not the formulas.** The same model produced a $178 implied price with pre-Q4 assumptions and ~$90 after updating for earnings reality. The mechanical skill is building a model that balances and flows correctly. The analytical skill is knowing which assumptions to challenge and why.

The scenario analysis bridges the gap between fundamental DCF value (what the business is worth if things go right) and market price (what investors are willing to pay given uncertainty). Neither number is "wrong" — they answer different questions.

## Status

- [x] Phase 1: Data extraction (Python + yfinance API)
- [x] Phase 2: SQL database loading
- [x] Phase 3: Excel financial model (3-statement + DCF + scenarios)
- [ ] Phase 4: Power BI dashboard
- [ ] Phase 5: Final documentation and presentation

## Disclaimer

This project is for educational and portfolio demonstration purposes only. It does not constitute investment advice. The author has no position in PYPL. All projections are forward-looking estimates subject to significant uncertainty.
