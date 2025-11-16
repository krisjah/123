#!/usr/bin/env python3
"""
Beer Brewery Financial Model - 3-Statement Integration
Builds: P&L, CashFlow, BalanceSheet (fully integrated)

Run after: 04_capital_funding.py

IMPORTANT: This creates the integrated financial statements.
Enable Excel iterative calculation after opening the file:
File → Options → Formulas → Enable iterative calculation (100 iterations, 0.001 max change)
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Color definitions
BLUE_INPUT = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
GREY_CALC = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
YELLOW_OUTPUT = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
GOLD_TOTAL = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
DARK_BLUE_HEADER = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

MONTHS = 120
TIMELINE_START_COL = 6


def add_timeline_header(ws, start_col=6, num_months=120):
    """Add standard timeline header rows"""
    ws['A3'] = 'Period'
    ws['A4'] = 'Date'
    ws['A5'] = 'Year'

    for month in range(1, num_months + 1):
        col = start_col + month - 1
        col_letter = get_column_letter(col)
        ws[f'{col_letter}3'] = month
        ws[f'{col_letter}3'].font = Font(size=9)
        ws[f'{col_letter}4'] = f'=EDATE(Control!$B$4,{month-1})'
        ws[f'{col_letter}4'].number_format = 'mmm-yy'
        ws[f'{col_letter}4'].font = Font(bold=True)
        ws[f'{col_letter}5'] = f'=({month}-1)/12'
        ws[f'{col_letter}5'].number_format = '0.00'
        ws[f'{col_letter}5'].font = Font(size=9, color="808080")


def build_pl_sheet(wb):
    """Build P&L (Income Statement) sheet"""

    print("Building P&L sheet...")
    ws = wb['P&L']

    # Title
    ws['A1'] = 'INCOME STATEMENT (P&L)'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = DARK_BLUE_HEADER
    ws.merge_cells('A1:E1')

    add_timeline_header(ws, start_col=TIMELINE_START_COL, num_months=MONTHS)

    # P&L structure
    pl_rows = [
        ('INCOME STATEMENT', 10, 'header'),
        ('', 11, 'blank'),
        ('Revenue', 12, 'header'),
        ('Net revenue', 13, 'calc'),
        ('', 14, 'blank'),
        ('Cost of Goods Sold', 15, 'header'),
        ('COGS (simplified)', 16, 'calc'),
        ('', 17, 'blank'),
        ('Gross Profit', 18, 'total'),
        ('Gross margin %', 19, 'calc'),
        ('', 20, 'blank'),
        ('Operating Expenses', 21, 'header'),
        ('Logistics & distribution', 22, 'calc'),
        ('SG&A', 23, 'calc'),
        ('Total OpEx', 24, 'total'),
        ('', 25, 'blank'),
        ('EBITDA', 26, 'total'),
        ('EBITDA margin %', 27, 'calc'),
        ('', 28, 'blank'),
        ('Depreciation', 29, 'calc'),
        ('', 30, 'blank'),
        ('EBIT', 31, 'total'),
        ('EBIT margin %', 32, 'calc'),
        ('', 33, 'blank'),
        ('Interest expense', 34, 'calc'),
        ('', 35, 'blank'),
        ('EBT (Earnings Before Tax)', 36, 'total'),
        ('', 37, 'blank'),
        ('Tax expense', 38, 'calc'),
        ('', 39, 'blank'),
        ('Net Income', 40, 'grand_total'),
        ('Net margin %', 41, 'calc'),
    ]

    for label, row, row_type in pl_rows:
        ws[f'A{row}'] = label

        if row_type == 'header':
            ws[f'A{row}'].font = Font(bold=True, size=11)
        elif row_type in ['total', 'grand_total']:
            ws[f'A{row}'].font = Font(bold=True)
            if row_type == 'grand_total':
                ws[f'A{row}'].fill = GOLD_TOTAL

    # Add formulas for each month
    for month in range(1, MONTHS + 1):
        col = TIMELINE_START_COL + month - 1
        col_letter = get_column_letter(col)

        # Net revenue (simplified - from volume * avg price of $150/HL)
        ws[f'{col_letter}13'] = f'=Volume_Assumptions!{col_letter}51*150'
        ws[f'{col_letter}13'].number_format = '$#,##0'

        # COGS (simplified - 55% of revenue as placeholder)
        ws[f'{col_letter}16'] = f'={col_letter}13*0.55'
        ws[f'{col_letter}16'].number_format = '$#,##0'

        # Gross Profit
        ws[f'{col_letter}18'] = f'={col_letter}13-{col_letter}16'
        ws[f'{col_letter}18'].number_format = '$#,##0'
        ws[f'{col_letter}18'].fill = YELLOW_OUTPUT
        ws[f'{col_letter}18'].font = Font(bold=True)

        # Gross margin %
        ws[f'{col_letter}19'] = f'={col_letter}18/{col_letter}13'
        ws[f'{col_letter}19'].number_format = '0.0%'

        # Logistics (simplified - $6/HL)
        ws[f'{col_letter}22'] = f'=Volume_Assumptions!{col_letter}51*6'
        ws[f'{col_letter}22'].number_format = '$#,##0'

        # SG&A (simplified - 10% of revenue)
        ws[f'{col_letter}23'] = f'={col_letter}13*0.10'
        ws[f'{col_letter}23'].number_format = '$#,##0'

        # Total OpEx
        ws[f'{col_letter}24'] = f'={col_letter}22+{col_letter}23'
        ws[f'{col_letter}24'].number_format = '$#,##0'
        ws[f'{col_letter}24'].fill = YELLOW_OUTPUT

        # EBITDA
        ws[f'{col_letter}26'] = f'={col_letter}18-{col_letter}24'
        ws[f'{col_letter}26'].number_format = '$#,##0'
        ws[f'{col_letter}26'].fill = YELLOW_OUTPUT
        ws[f'{col_letter}26'].font = Font(bold=True)

        # EBITDA margin %
        ws[f'{col_letter}27'] = f'={col_letter}26/{col_letter}13'
        ws[f'{col_letter}27'].number_format = '0.0%'

        # Depreciation
        ws[f'{col_letter}29'] = f'=Capex_Schedule!{col_letter}19'
        ws[f'{col_letter}29'].number_format = '$#,##0'

        # EBIT
        ws[f'{col_letter}31'] = f'={col_letter}26-{col_letter}29'
        ws[f'{col_letter}31'].number_format = '$#,##0'
        ws[f'{col_letter}31'].fill = YELLOW_OUTPUT
        ws[f'{col_letter}31'].font = Font(bold=True)

        # EBIT margin %
        ws[f'{col_letter}32'] = f'={col_letter}31/{col_letter}13'
        ws[f'{col_letter}32'].number_format = '0.0%'

        # Interest expense
        ws[f'{col_letter}34'] = f'=Debt_Schedule!{col_letter}20'
        ws[f'{col_letter}34'].number_format = '$#,##0'

        # EBT
        ws[f'{col_letter}36'] = f'={col_letter}31-{col_letter}34'
        ws[f'{col_letter}36'].number_format = '$#,##0'
        ws[f'{col_letter}36'].fill = YELLOW_OUTPUT

        # Tax expense
        ws[f'{col_letter}38'] = f'=MAX(0,{col_letter}36*Assumptions_Global!$E$11)'
        ws[f'{col_letter}38'].number_format = '$#,##0'

        # Net Income
        ws[f'{col_letter}40'] = f'={col_letter}36-{col_letter}38'
        ws[f'{col_letter}40'].number_format = '$#,##0'
        ws[f'{col_letter}40'].fill = GOLD_TOTAL
        ws[f'{col_letter}40'].font = Font(bold=True)

        # Net margin %
        ws[f'{col_letter}41'] = f'={col_letter}40/{col_letter}13'
        ws[f'{col_letter}41'].number_format = '0.0%'

    ws.column_dimensions['A'].width = 35

    print("  ✓ P&L sheet complete")


def build_cashflow_sheet(wb):
    """Build Cash Flow Statement sheet"""

    print("Building CashFlow sheet...")
    ws = wb['CashFlow']

    # Title
    ws['A1'] = 'CASH FLOW STATEMENT'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = DARK_BLUE_HEADER
    ws.merge_cells('A1:E1')

    add_timeline_header(ws, start_col=TIMELINE_START_COL, num_months=MONTHS)

    # Cash flow structure
    cf_rows = [
        ('CASH FLOW STATEMENT', 10, 'header'),
        ('', 11, 'blank'),
        ('Operating Activities', 12, 'header'),
        ('Net income', 13, 'calc'),
        ('Add: Depreciation', 14, 'calc'),
        ('Less: Increase in working capital', 15, 'calc'),
        ('Cash from operations', 16, 'total'),
        ('', 17, 'blank'),
        ('Investing Activities', 18, 'header'),
        ('Capital expenditures', 19, 'calc'),
        ('Cash from investing', 20, 'total'),
        ('', 21, 'blank'),
        ('Financing Activities', 22, 'header'),
        ('Debt drawdown', 23, 'calc'),
        ('Debt repayment', 24, 'calc'),
        ('Equity injection', 25, 'calc'),
        ('Dividends paid', 26, 'calc'),
        ('Cash from financing', 27, 'total'),
        ('', 28, 'blank'),
        ('Net change in cash', 29, 'total'),
        ('Opening cash', 30, 'calc'),
        ('Closing cash', 31, 'grand_total'),
    ]

    for label, row, row_type in cf_rows:
        ws[f'A{row}'] = label

        if row_type == 'header':
            ws[f'A{row}'].font = Font(bold=True, size=11)
        elif row_type in ['total', 'grand_total']:
            ws[f'A{row}'].font = Font(bold=True)
            if row_type == 'grand_total':
                ws[f'A{row}'].fill = GOLD_TOTAL

    # Month 0 (before start)
    ws['E10'] = 'M0'
    ws['E10'].font = Font(bold=True, size=9)
    ws['E31'] = 0  # Opening cash = 0
    ws['E31'].number_format = '$#,##0'

    # Add formulas for each month
    for month in range(1, MONTHS + 1):
        col = TIMELINE_START_COL + month - 1
        col_letter = get_column_letter(col)
        prev_col = get_column_letter(col - 1) if month > 1 else 'E'

        # Operating Activities
        ws[f'{col_letter}13'] = f'=PL!{col_letter}40'  # Net income
        ws[f'{col_letter}13'].number_format = '$#,##0'

        ws[f'{col_letter}14'] = f'=PL!{col_letter}29'  # Depreciation
        ws[f'{col_letter}14'].number_format = '$#,##0'

        # Working capital (simplified - 5% of revenue increase)
        ws[f'{col_letter}15'] = f'=-PL!{col_letter}13*0.05'
        ws[f'{col_letter}15'].number_format = '$#,##0'

        ws[f'{col_letter}16'] = f'={col_letter}13+{col_letter}14+{col_letter}15'
        ws[f'{col_letter}16'].number_format = '$#,##0'
        ws[f'{col_letter}16'].fill = YELLOW_OUTPUT
        ws[f'{col_letter}16'].font = Font(bold=True)

        # Investing Activities
        ws[f'{col_letter}19'] = f'=-Capex_Schedule!{col_letter}18'  # Capex (negative)
        ws[f'{col_letter}19'].number_format = '$#,##0'

        ws[f'{col_letter}20'] = f'={col_letter}19'
        ws[f'{col_letter}20'].number_format = '$#,##0'
        ws[f'{col_letter}20'].fill = YELLOW_OUTPUT

        # Financing Activities
        ws[f'{col_letter}23'] = f'=Debt_Schedule!{col_letter}19'  # Drawdown
        ws[f'{col_letter}23'].number_format = '$#,##0'

        ws[f'{col_letter}24'] = f'=-Debt_Schedule!{col_letter}21'  # Repayment (negative)
        ws[f'{col_letter}24'].number_format = '$#,##0'

        ws[f'{col_letter}25'] = f'=Equity!{col_letter}16'  # Equity injection
        ws[f'{col_letter}25'].number_format = '$#,##0'

        ws[f'{col_letter}26'] = f'=-Equity!{col_letter}18'  # Dividends (negative)
        ws[f'{col_letter}26'].number_format = '$#,##0'

        ws[f'{col_letter}27'] = f'={col_letter}23+{col_letter}24+{col_letter}25+{col_letter}26'
        ws[f'{col_letter}27'].number_format = '$#,##0'
        ws[f'{col_letter}27'].fill = YELLOW_OUTPUT

        # Net change in cash
        ws[f'{col_letter}29'] = f'={col_letter}16+{col_letter}20+{col_letter}27'
        ws[f'{col_letter}29'].number_format = '$#,##0'
        ws[f'{col_letter}29'].fill = YELLOW_OUTPUT
        ws[f'{col_letter}29'].font = Font(bold=True)

        # Opening cash
        ws[f'{col_letter}30'] = f'={prev_col}31'
        ws[f'{col_letter}30'].number_format = '$#,##0'

        # Closing cash
        ws[f'{col_letter}31'] = f'={col_letter}30+{col_letter}29'
        ws[f'{col_letter}31'].number_format = '$#,##0'
        ws[f'{col_letter}31'].fill = GOLD_TOTAL
        ws[f'{col_letter}31'].font = Font(bold=True)

    ws.column_dimensions['A'].width = 35

    print("  ✓ CashFlow sheet complete")


def build_balancesheet_sheet(wb):
    """Build Balance Sheet sheet"""

    print("Building BalanceSheet sheet...")
    ws = wb['BalanceSheet']

    # Title
    ws['A1'] = 'BALANCE SHEET'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = DARK_BLUE_HEADER
    ws.merge_cells('A1:E1')

    add_timeline_header(ws, start_col=TIMELINE_START_COL, num_months=MONTHS)

    # Balance sheet structure
    bs_rows = [
        ('BALANCE SHEET', 10, 'header'),
        ('', 11, 'blank'),
        ('ASSETS', 12, 'header'),
        ('Fixed Assets', 13, 'header'),
        ('Gross fixed assets', 14, 'calc'),
        ('Accumulated depreciation', 15, 'calc'),
        ('Net book value', 16, 'total'),
        ('', 17, 'blank'),
        ('Current Assets', 18, 'header'),
        ('Inventory', 19, 'calc'),
        ('Trade receivables', 20, 'calc'),
        ('Cash', 21, 'calc'),
        ('Total current assets', 22, 'total'),
        ('', 23, 'blank'),
        ('TOTAL ASSETS', 24, 'grand_total'),
        ('', 25, 'blank'),
        ('LIABILITIES', 26, 'header'),
        ('Current Liabilities', 27, 'header'),
        ('Trade payables', 28, 'calc'),
        ('', 29, 'blank'),
        ('Non-Current Liabilities', 30, 'header'),
        ('Debt', 31, 'calc'),
        ('Total liabilities', 32, 'total'),
        ('', 33, 'blank'),
        ('EQUITY', 34, 'header'),
        ('Share capital & retained earnings', 35, 'calc'),
        ('', 36, 'blank'),
        ('TOTAL LIABILITIES & EQUITY', 37, 'grand_total'),
        ('', 38, 'blank'),
        ('BALANCE CHECK', 39, 'check'),
    ]

    for label, row, row_type in bs_rows:
        ws[f'A{row}'] = label

        if row_type == 'header':
            ws[f'A{row}'].font = Font(bold=True, size=11)
        elif row_type in ['total', 'grand_total']:
            ws[f'A{row}'].font = Font(bold=True)
            if row_type == 'grand_total':
                ws[f'A{row}'].fill = GOLD_TOTAL
        elif row_type == 'check':
            ws[f'A{row}'].font = Font(bold=True, color="FF0000")

    # Month 0
    ws['E10'] = 'M0'
    ws['E10'].font = Font(bold=True, size=9)
    for row in [14, 15, 16, 19, 20, 21, 22, 24, 28, 31, 32, 35, 37, 39]:
        ws[f'E{row}'] = 0
        ws[f'E{row}'].number_format = '$#,##0'

    # Add formulas for each month
    for month in range(1, MONTHS + 1):
        col = TIMELINE_START_COL + month - 1
        col_letter = get_column_letter(col)

        # Fixed Assets
        ws[f'{col_letter}14'] = f'=Capex_Schedule!{col_letter}20'  # Gross FA
        ws[f'{col_letter}14'].number_format = '$#,##0'

        ws[f'{col_letter}15'] = f'=Capex_Schedule!{col_letter}21'  # Accum depn
        ws[f'{col_letter}15'].number_format = '$#,##0'

        ws[f'{col_letter}16'] = f'={col_letter}14-{col_letter}15'  # NBV
        ws[f'{col_letter}16'].number_format = '$#,##0'
        ws[f'{col_letter}16'].fill = YELLOW_OUTPUT

        # Current Assets (simplified)
        ws[f'{col_letter}19'] = f'=PL!{col_letter}16*(Working_Capital!$E$10/30)'  # Inventory
        ws[f'{col_letter}19'].number_format = '$#,##0'

        ws[f'{col_letter}20'] = f'=PL!{col_letter}13*0.30'  # Receivables (simplified - 30% of monthly revenue)
        ws[f'{col_letter}20'].number_format = '$#,##0'

        ws[f'{col_letter}21'] = f'=CashFlow!{col_letter}31'  # Cash
        ws[f'{col_letter}21'].number_format = '$#,##0'

        ws[f'{col_letter}22'] = f'={col_letter}19+{col_letter}20+{col_letter}21'
        ws[f'{col_letter}22'].number_format = '$#,##0'
        ws[f'{col_letter}22'].fill = YELLOW_OUTPUT

        # Total Assets
        ws[f'{col_letter}24'] = f'={col_letter}16+{col_letter}22'
        ws[f'{col_letter}24'].number_format = '$#,##0'
        ws[f'{col_letter}24'].fill = GOLD_TOTAL
        ws[f'{col_letter}24'].font = Font(bold=True)

        # Liabilities
        ws[f'{col_letter}28'] = f'=PL!{col_letter}16*0.40'  # Payables (simplified - 40% of monthly COGS)
        ws[f'{col_letter}28'].number_format = '$#,##0'

        ws[f'{col_letter}31'] = f'=Debt_Schedule!{col_letter}22'  # Debt
        ws[f'{col_letter}31'].number_format = '$#,##0'

        ws[f'{col_letter}32'] = f'={col_letter}28+{col_letter}31'
        ws[f'{col_letter}32'].number_format = '$#,##0'
        ws[f'{col_letter}32'].fill = YELLOW_OUTPUT

        # Equity
        ws[f'{col_letter}35'] = f'=Equity!{col_letter}19'  # Total equity
        ws[f'{col_letter}35'].number_format = '$#,##0'

        # Total Liabilities & Equity
        ws[f'{col_letter}37'] = f'={col_letter}32+{col_letter}35'
        ws[f'{col_letter}37'].number_format = '$#,##0'
        ws[f'{col_letter}37'].fill = GOLD_TOTAL
        ws[f'{col_letter}37'].font = Font(bold=True)

        # Balance Check
        ws[f'{col_letter}39'] = f'={col_letter}24-{col_letter}37'
        ws[f'{col_letter}39'].number_format = '$#,##0'
        ws[f'{col_letter}39'].font = Font(bold=True, color="FF0000")

    ws.column_dimensions['A'].width = 35

    print("  ✓ BalanceSheet sheet complete")


def update_equity_with_netincome(wb):
    """Update Equity sheet row 17 to link to P&L Net Income"""

    print("Updating Equity sheet with Net Income link...")
    ws = wb['Equity']

    for month in range(1, MONTHS + 1):
        col = TIMELINE_START_COL + month - 1
        col_letter = get_column_letter(col)

        # Update retained earnings to link to P&L Net Income
        ws[f'{col_letter}17'] = f'=PL!{col_letter}40'
        ws[f'{col_letter}17'].number_format = '$#,##0'

    print("  ✓ Equity sheet updated with Net Income links")


def main():
    """Main execution"""
    print("\n" + "="*60)
    print("Building 3-Statement Integration")
    print("="*60 + "\n")

    wb = load_workbook('Brewery_Financial_Model.xlsx')

    build_pl_sheet(wb)
    build_cashflow_sheet(wb)
    build_balancesheet_sheet(wb)
    update_equity_with_netincome(wb)

    wb.save('Brewery_Financial_Model.xlsx')

    print("\n✅ 3-Statement Integration complete!")
    print("\n   ✓ P&L: Revenue → EBITDA → EBIT → Net Income")
    print("   ✓ Cash Flow: Operating, Investing, Financing")
    print("   ✓ Balance Sheet: Assets = Liabilities + Equity")
    print("   ✓ All statements fully linked")
    print("\n" + "="*60)
    print("IMPORTANT: Enable Iterative Calculation in Excel")
    print("="*60)
    print("\n1. Open Brewery_Financial_Model.xlsx")
    print("2. File → Options → Formulas")
    print("3. Check 'Enable iterative calculation'")
    print("4. Set: Max iterations = 100, Max change = 0.001")
    print("\nThis is required for circular references (interest on debt).")
    print("\nNext step: Run 06_kpis_checks.py")


if __name__ == '__main__':
    main()
