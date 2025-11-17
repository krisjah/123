# 3-Statement Financial Model with Control Accounts

## Overview

This repository contains a fully integrated 3-statement financial model built in Excel following professional financial modeling best practices. The model features **control accounts for every Balance Sheet line item**, ensuring complete auditability and transparency.

## Model Structure

### Files

- **3_Statement_Financial_Model.xlsx** - The complete financial model
- **build_financial_model.py** - Python script to regenerate the model

### Sheets

1. **Inputs** - All assumptions and drivers (color-coded for easy identification)
2. **Workings** - All control accounts tracking balance sheet movements
3. **Income Statement** - Monthly P&L statement
4. **Balance Sheet** - Monthly balance sheet (all items linked to control accounts)
5. **Cash Flow Statement** - Indirect method cash flow statement

## Key Features

### Control Account Architecture

Every Balance Sheet line item is driven by its own control account that tracks:
- **Opening Balance** - Balance at start of period
- **Increases** - All movements that increase the account
- **Decreases** - All movements that reduce the account
- **Closing Balance** = Opening + Increases - Decreases

### Control Accounts Implemented

#### Assets
1. **Cash Control Account**
   - Tracks all cash inflows (customer collections, borrowing, equity)
   - Tracks all cash outflows (supplier payments, opex, capex, interest, debt repayment, dividends)

2. **Accounts Receivable Control**
   - Opening AR + Credit Sales - Cash Collections = Closing AR
   - Collections based on AR Days assumption

3. **Inventory Control**
   - Opening + Purchases - COGS Withdrawal + Adjustments = Closing
   - Purchases calculated to maintain inventory days

4. **Prepayments Control**
   - Opening + New Prepayments - Amortization = Closing
   - Maintains prepayment months assumption

5. **Fixed Assets (Gross) Control**
   - Opening + CapEx - Disposals = Closing

6. **Accumulated Depreciation Control**
   - Opening + Depreciation Expense = Closing
   - Straight-line depreciation based on gross assets

#### Liabilities
7. **Accounts Payable Control**
   - Opening + Purchases on Credit - Cash Payments = Closing
   - Payments based on AP Days assumption

8. **Accrued Expenses Control**
   - Opening + Expenses Incurred - Cash Paid = Closing
   - Accrual rate applied to operating expenses

9. **Deferred Revenue Control**
   - Opening + Cash Received in Advance - Revenue Recognized = Closing
   - Based on deferred revenue months assumption

10. **Debt Control**
    - Opening + New Borrowing + Interest Accrued - Interest Paid - Principal Repayment = Closing
    - Monthly interest calculation

#### Equity
11. **Share Capital Control**
    - Opening + New Share Capital Issued = Closing

12. **Retained Earnings Control**
    - Opening + Net Income - Dividends = Closing
    - Automatically flows from Income Statement

## Financial Statements

### Income Statement
- **Revenue** - Grows monthly based on growth rate
- **COGS** - As % of revenue, linked to inventory control
- **Operating Expenses** - Salaries, rent, marketing, other
- **Depreciation** - From accumulated depreciation control
- **Interest Expense** - From debt control
- **Net Income** - Flows to retained earnings

### Balance Sheet
- **All line items** pulled from control account closing balances
- **Built-in balance check** - Assets must equal Liabilities + Equity
- **Fully reconciles** with control accounts

### Cash Flow Statement (Indirect Method)
- **Operating Activities**
  - Starts with Net Income
  - Adds back depreciation
  - Adjusts for working capital changes (AR, inventory, prepayments, AP, accruals, deferred revenue)

- **Investing Activities**
  - Capital expenditure
  - Asset disposals

- **Financing Activities**
  - New borrowing
  - Debt repayment
  - Interest paid
  - Share capital raised
  - Dividends paid

- **Built-in cash reconciliation check** - Must match Balance Sheet cash

## Key Assumptions (Inputs Sheet)

### Revenue
- Base revenue: $100,000 (Month 1)
- Monthly growth rate: 5%

### Cost Structure
- COGS: 40% of revenue
- Operating expenses: Salaries, rent, marketing, other

### Working Capital
- AR Days: 45
- Inventory Days: 30
- AP Days: 30
- Prepayments: 1 month
- Accrued expenses: 15% of OpEx
- Deferred revenue: 1 month

### Capital
- Initial CapEx: $50,000 (Month 1)
- Annual depreciation rate: 20%

### Financing
- Initial debt: $200,000 (Month 1)
- Annual interest rate: 6%
- Monthly debt repayment: $5,000
- Initial equity: $100,000 (Month 1)

### Opening Balances
- Opening cash: $50,000
- All other balances: $0

## Color Coding

- **Dark Blue** - Headers and section titles
- **Light Orange** - Input cells (user can modify)
- **Light Blue** - Calculated cells (formulas)
- **Medium Blue** - Subtotals and section totals
- **Gold** - Validation checks (should always be zero)

## Model Standards

### Best Practices Implemented

1. **No hardcoding in calculation blocks** - All inputs in dedicated Inputs sheet
2. **Clear separation** - Inputs, Workings, and Outputs in separate sheets
3. **Transparent formulas** - All calculations can be traced and audited
4. **Control accounts** - Every balance sheet item has full movement tracking
5. **Full integration** - All three statements fully linked and reconciled
6. **Built-in checks** - Balance sheet balance check and cash reconciliation check
7. **Consistent structure** - Monthly timeline across all sheets
8. **Indirect cash flow** - Derived from P&L and balance sheet movements, not hardcoded

### Formula Principles

- **Absolute references** for inputs (e.g., $C$12)
- **Relative references** for timeline calculations
- **Direct links** between control accounts and financial statements
- **No circular references** - All calculations flow in one direction

## How to Use

### Regenerating the Model

If you need to modify the model structure:

```bash
python build_financial_model.py
```

This will generate a fresh `3_Statement_Financial_Model.xlsx` file.

### Modifying Assumptions

1. Open `3_Statement_Financial_Model.xlsx`
2. Go to the **Inputs** sheet
3. Modify any cells with **light orange** background (input cells)
4. All calculations will update automatically

### Key Scenarios to Model

- Change revenue growth rate
- Adjust working capital days (AR, inventory, AP)
- Modify capital expenditure plans
- Test different financing scenarios (debt vs. equity)
- Analyze impact of dividend policies

## Validation Checks

The model includes two critical validation checks:

1. **Balance Sheet Check** (Balance Sheet row 45)
   - Formula: Total Assets - (Total Liabilities + Total Equity)
   - **Should always be 0**
   - If not zero, there's an error in the balance sheet linkages

2. **Cash Reconciliation Check** (Cash Flow Statement row 36)
   - Formula: Cash Flow Statement closing cash - Balance Sheet cash
   - **Should always be 0**
   - If not zero, there's a mismatch between the cash flow and balance sheet

## Model Timeline

- **Projection period**: 12 months
- **Start date**: January 2025
- **Frequency**: Monthly
- **All periods** use consistent calculation methodology

## Technical Details

### Dependencies

- Python 3.x
- openpyxl library (for generating Excel files)

### Installation

```bash
pip install openpyxl
```

## Reconciliation Flow

```
Income Statement Net Income
    ↓
Retained Earnings Control Account
    ↓
Balance Sheet Retained Earnings
    ↓
Cash Flow Statement (Net Income starting point)
    ↓
Cash Control Account
    ↓
Balance Sheet Cash
```

## Control Account Example: Accounts Receivable

```
Opening AR Balance (from previous period)
+ Credit Sales (from Revenue)
- Cash Received from Customers (based on AR Days)
= Closing AR Balance (shown on Balance Sheet)
```

The cash received flows into the Cash Control Account, which then flows to the Balance Sheet and Cash Flow Statement.

## Support & Modifications

To modify the model structure:
1. Edit `build_financial_model.py`
2. Adjust row numbers, formulas, or add new control accounts
3. Regenerate the Excel file

The Python script is fully documented and modular, making it easy to:
- Add new line items
- Modify control account logic
- Change timeline length
- Add additional sheets

## License

This financial model is provided as-is for educational and business purposes.

---

**Built with professional financial modeling standards**
**All statements reconcile through control accounts**
**Fully auditable and transparent**
