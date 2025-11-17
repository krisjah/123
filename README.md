# Beer Brewery Financial Model - Build Scripts

Complete Python-based builder for a 10-year, 3-statement financial model for a beer brewery and distribution business.

## Overview

This model includes:
- **19 worksheets**: From assumptions to 3-statement integration to KPI dashboards
- **120 months**: Monthly projections over 10 years
- **3 scenarios**: Base, Upside, Downside
- **Full integration**: P&L, Cash Flow, Balance Sheet with circular reference handling
- **60+ KPIs**: Volume, profitability, efficiency, working capital, leverage, returns

## Requirements

```bash
pip install openpyxl
```

Python 3.7+ required.

## Quick Start

### Option 1: Run All Scripts at Once

```bash
python run_all.py
```

This executes all build scripts in sequence and creates the complete model.

### Option 2: Run Scripts Step-by-Step

Run each script in order:

```bash
python 00_setup_workbook.py          # Creates blank workbook
python 01_control_assumptions.py     # Builds control & assumptions sheets
python 02_volume_pricing.py          # Builds volume & revenue logic
python 03_operations_costs.py        # Builds COGS & operating costs
python 04_capital_funding.py         # Builds capex, debt, equity, working capital
python 05_statements.py              # Builds P&L, Cash Flow, Balance Sheet
python 06_kpis_checks.py             # Builds KPI dashboard & error checks
python 07_cover_docs.py              # Builds cover, documentation, change log
python 08_formatting.py              # Applies formatting, colors, protection
```

## Output

**File generated**: `Brewery_Financial_Model.xlsx`

## Script Descriptions

| Script | Purpose | Sheets Built |
|--------|---------|--------------|
| `00_setup_workbook.py` | Creates blank workbook structure | All 19 sheets (empty) |
| `01_control_assumptions.py` | Scenario control & global assumptions | Control, Assumptions_Global, Product_Matrix |
| `02_volume_pricing.py` | Volume forecasts & revenue calculations | Volume_Assumptions, Pricing_Revenue |
| `03_operations_costs.py` | Production, COGS, logistics, SG&A | Production_Capacity, COGS_Buildup, Logistics_Distribution, SG&A |
| `04_capital_funding.py` | Capex, debt, equity, working capital | Capex_Schedule, Working_Capital, Debt_Schedule, Equity |
| `05_statements.py` | Integrated financial statements | P&L, CashFlow, BalanceSheet |
| `06_kpis_checks.py` | KPIs and model validation | KPI_Dashboard, Checks_Errors |
| `07_cover_docs.py` | Documentation and navigation | Cover, Documentation, Change_Log |
| `08_formatting.py` | Final formatting and protection | All sheets (formatting only) |
| `run_all.py` | Master script to run all steps | N/A (orchestrator) |

## Key Features

### Scenario Management
- **Base**: Central case (3% volume growth, 2.5% price escalation)
- **Upside**: Strong demand (5% volume growth, 3.5% pricing)
- **Downside**: Weak demand (1% volume growth, excise shock in Year 3)

Change scenario in `Control` sheet dropdown - entire model updates instantly.

### Product Structure
- **5 brands**: SkyBrew (lager), AllDark (stout), GoldCoast (pale ale), CityLight (light), HarborIPA (IPA)
- **3 pack types**: Bottles, Cans, Kegs
- **2 channels**: OnTrade (bars/restaurants), OffTrade (retail/supermarkets)
- **30 SKUs**: 5 brands × 3 packs × 2 channels (note: some combinations may be zero, e.g., kegs in OffTrade)

### Financial Structure
- **Revenue**: Net of discounts, excise duty shown separately
- **COGS**: Materials, packaging, labour, utilities, overhead
- **OpEx**: Excise, logistics, SG&A (A&P as % of revenue)
- **Capex**: Brewery equipment, packaging lines, warehouse, IT
- **Financing**: Term loan ($10M, 7yr, 6.5%), equity injections
- **Working Capital**: Inventory (66 days), receivables (36 days), payables (45 days)

### Integration
- **Circular references**: Handled via Excel iterative calculation (interest depends on debt depends on cash)
- **Balance sheet balances**: Every period Assets = Liabilities + Equity
- **Cash flow reconciles**: Operating + Investing + Financing = Change in Cash

### Quality Control
- **40+ automated checks**: Balance sheet balance, capacity constraints, liquidity, covenants
- **Data validation**: Scenario dropdown, input ranges
- **Color coding**: Blue (inputs), Grey (links), White (calcs), Yellow (outputs), Red (errors)
- **Protection**: Formulas locked, inputs unlocked

## Using the Model

1. **Open** `Brewery_Financial_Model.xlsx`
2. **Enable iterative calculation**: File → Options → Formulas → Enable iterative calculation (100 iterations, 0.001 max change)
3. **Select scenario**: Go to `Control` sheet, choose Base/Upside/Downside
4. **Review/edit inputs**: All BLUE cells are editable (assumptions, volumes, pricing, costs)
5. **Check for errors**: Go to `Checks_Errors` sheet, verify "ALL CHECKS PASS"
6. **Review outputs**: P&L, CashFlow, BalanceSheet, KPI_Dashboard

## Customization

### Adding New Brands
1. Edit `Product_Matrix` sheet: add row to Brand List
2. Edit `Volume_Assumptions`: add rows for new brand × pack × channel combinations
3. Edit `Pricing_Revenue`: add corresponding pricing rows
4. Update summary formulas (SUM/SUMIF ranges)

### Changing Projection Period
1. `Control` sheet: change "Projection period (months)" (default: 120)
2. Re-run affected scripts (or extend/contract columns manually)

### Adding Capex
1. `Capex_Schedule` sheet: add row with description, month, amount, useful life
2. Depreciation calculates automatically

## Troubleshooting

### Balance Sheet Doesn't Balance
- Check `Checks_Errors` sheet for imbalance amount
- Verify iterative calculation is enabled
- Check for broken formulas in Equity or Debt schedules

### Negative Cash
- Review `Cash Flow` statement Month-by-month
- Likely causes: insufficient initial equity, large capex without financing, working capital build
- Solution: Increase equity injection in `Equity` sheet or add/draw more debt

### #REF! Errors
- Likely cause: sheet names changed or cells deleted
- Re-run the script that builds the affected sheet

### Model is Slow
- Large models with 120 months × complex formulas can take 5-10 seconds to recalculate
- Ensure calculation mode is "Automatic" (not Manual)
- Close other Excel workbooks

## Model Design Summary

Full design documentation provided in Steps 1-8:
- **Step 1**: Scope & requirements (Q&A)
- **Step 2**: Architecture (19 sheets, data flow)
- **Step 3**: Assumptions & drivers (8 blocks: volume, pricing, COGS, logistics, SG&A, capex, WC, financing)
- **Step 4**: Timeline, seasonality, scenarios
- **Step 5**: Sheet-by-sheet build plan
- **Step 6**: 3-statement integration
- **Step 7**: KPIs & dashboards
- **Step 8**: Quality control & documentation

## Support

For questions or issues:
- Review `Documentation` sheet in the model
- Check `Change_Log` for modification history
- Refer to design documentation (Steps 1-8 above)

## License

This model is provided as-is for use by the requestor. Modify as needed for your brewery business.

---

**Created**: November 2025
**Version**: 1.0
**Model Type**: 3-Statement Financial Model (P&L, Cash Flow, Balance Sheet)
**Industry**: Beer Brewery & Distribution (Australia/New Zealand)
