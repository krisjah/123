#!/usr/bin/env python3
"""
Beer Brewery Financial Model - Capital & Funding
Builds: Capex_Schedule, Working_Capital, Debt_Schedule, Equity sheets

Run after: 03_operations_costs.py
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Color definitions
BLUE_INPUT = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
GREY_CALC = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
YELLOW_OUTPUT = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
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


def build_capex_schedule_sheet(wb):
    """Build Capex_Schedule sheet"""

    print("Building Capex_Schedule sheet...")
    ws = wb['Capex_Schedule']

    # Title
    ws['A1'] = 'CAPEX SCHEDULE'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = DARK_BLUE_HEADER
    ws.merge_cells('A1:F1')

    add_timeline_header(ws, start_col=TIMELINE_START_COL, num_months=MONTHS)

    # Capex plan
    ws['A8'] = 'Capital Expenditure Plan'
    ws['A8'].font = Font(bold=True, size=11)

    headers = ['Description', 'Capex Month', 'Amount ($)', 'Useful Life (years)', 'Residual %', 'Category']
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=9, column=col_idx)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = DARK_BLUE_HEADER

    capex_items = [
        ('Initial brewhouse', 1, 8000000, 15, 0.10, 'Brewery'),
        ('Initial packaging line', 1, 2500000, 10, 0.10, 'Packaging'),
        ('Warehouse', 1, 1500000, 25, 0.10, 'Warehouse'),
        ('IT systems', 1, 300000, 5, 0.00, 'IT'),
        ('Vehicles & kegs', 1, 500000, 7, 0.10, 'Other'),
    ]

    row = 10
    for desc, month, amount, life, resid, category in capex_items:
        ws[f'A{row}'] = desc
        ws[f'B{row}'] = month
        ws[f'C{row}'] = amount
        ws[f'C{row}'].number_format = '$#,##0'
        ws[f'C{row}'].fill = BLUE_INPUT
        ws[f'D{row}'] = life
        ws[f'D{row}'].fill = BLUE_INPUT
        ws[f'E{row}'] = resid
        ws[f'E{row}'].number_format = '0%'
        ws[f'E{row}'].fill = BLUE_INPUT
        ws[f'F{row}'] = category
        row += 1

    # Summary calculations
    summary_row = row + 2
    ws[f'A{summary_row}'] = 'MONTHLY SUMMARY'
    ws[f'A{summary_row}'].font = Font(bold=True, size=11)
    ws[f'A{summary_row}'].fill = YELLOW_OUTPUT

    ws[f'A{summary_row+1}'] = 'Total capex (month)'
    ws[f'A{summary_row+2}'] = 'Total depreciation (month)'
    ws[f'A{summary_row+3}'] = 'Cumulative capex'
    ws[f'A{summary_row+4}'] = 'Cumulative depreciation'
    ws[f'A{summary_row+5}'] = 'Net book value'

    # Formulas for each month
    for month in range(1, MONTHS + 1):
        col = TIMELINE_START_COL + month - 1
        col_letter = get_column_letter(col)

        # Total capex this month
        capex_sum = '+'.join([f'IF(B{r}={month},C{r},0)' for r in range(10, 15)])
        ws[f'{col_letter}{summary_row+1}'] = f'={capex_sum}'
        ws[f'{col_letter}{summary_row+1}'].number_format = '$#,##0'

        # Total depreciation (sum of all active assets)
        # Simplified: each asset depreciates from its capex month onwards
        depn_formulas = []
        for r in range(10, 15):
            # Monthly depreciation = (Amount * (1-Residual%)) / (Life * 12)
            # Only if month >= capex month
            depn_formulas.append(f'IF({month}>=B{r},(C{r}*(1-E{r}))/(D{r}*12),0)')

        ws[f'{col_letter}{summary_row+2}'] = f'={"+".join(depn_formulas)}'
        ws[f'{col_letter}{summary_row+2}'].number_format = '$#,##0'

        # Cumulative capex
        if month == 1:
            ws[f'{col_letter}{summary_row+3}'] = f'={col_letter}{summary_row+1}'
        else:
            prev_col = get_column_letter(col - 1)
            ws[f'{col_letter}{summary_row+3}'] = f'={prev_col}{summary_row+3}+{col_letter}{summary_row+1}'
        ws[f'{col_letter}{summary_row+3}'].number_format = '$#,##0'

        # Cumulative depreciation
        if month == 1:
            ws[f'{col_letter}{summary_row+4}'] = f'={col_letter}{summary_row+2}'
        else:
            prev_col = get_column_letter(col - 1)
            ws[f'{col_letter}{summary_row+4}'] = f'={prev_col}{summary_row+4}+{col_letter}{summary_row+2}'
        ws[f'{col_letter}{summary_row+4}'].number_format = '$#,##0'

        # NBV
        ws[f'{col_letter}{summary_row+5}'] = f'={col_letter}{summary_row+3}-{col_letter}{summary_row+4}'
        ws[f'{col_letter}{summary_row+5}'].number_format = '$#,##0'
        ws[f'{col_letter}{summary_row+5}'].fill = YELLOW_OUTPUT

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 15

    print("  ✓ Capex_Schedule sheet complete")


def build_working_capital_sheet(wb):
    """Build Working_Capital sheet"""

    print("Building Working_Capital sheet...")
    ws = wb['Working_Capital']

    # Title
    ws['A1'] = 'WORKING CAPITAL'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = DARK_BLUE_HEADER
    ws.merge_cells('A1:E1')

    add_timeline_header(ws, start_col=TIMELINE_START_COL, num_months=MONTHS)

    # Assumptions
    ws['A8'] = 'Working Capital Days'
    ws['A8'].font = Font(bold=True, size=11)

    headers = ['', 'Base', 'Upside', 'Downside', 'Active']
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=9, column=col_idx)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = DARK_BLUE_HEADER

    wc_assumptions = [
        ('Inventory days (total)', 66, 60, 72),
        ('DSO - OnTrade', 45, 42, 48),
        ('DSO - OffTrade', 30, 28, 32),
        ('DPO (supplier terms)', 45, 50, 40),
    ]

    row = 10
    for label, base, upside, downside in wc_assumptions:
        ws[f'A{row}'] = label
        ws[f'B{row}'] = base
        ws[f'C{row}'] = upside
        ws[f'D{row}'] = downside
        ws[f'E{row}'] = f'=INDEX(B{row}:D{row},1,Control!$B$10)'

        for col in ['B', 'C', 'D']:
            ws[f'{col}{row}'].fill = BLUE_INPUT
        ws[f'E{row}'].fill = GREY_CALC
        row += 1

    # Note: Simplified WC calculations (will be enhanced in statements)
    ws['A15'] = 'Note: Detailed WC calculations integrated in Balance Sheet'
    ws['A15'].font = Font(italic=True, size=9)

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 15

    print("  ✓ Working_Capital sheet complete")


def build_debt_schedule_sheet(wb):
    """Build Debt_Schedule sheet"""

    print("Building Debt_Schedule sheet...")
    ws = wb['Debt_Schedule']

    # Title
    ws['A1'] = 'DEBT SCHEDULE'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = DARK_BLUE_HEADER
    ws.merge_cells('A1:E1')

    add_timeline_header(ws, start_col=TIMELINE_START_COL, num_months=MONTHS)

    # Term loan assumptions
    ws['A8'] = 'Term Loan'
    ws['A8'].font = Font(bold=True, size=11)

    ws['A9'] = 'Facility size'
    ws['B9'] = 10000000
    ws['B9'].number_format = '$#,##0'
    ws['B9'].fill = BLUE_INPUT

    ws['A10'] = 'Drawdown month'
    ws['B10'] = 1
    ws['B10'].fill = BLUE_INPUT

    ws['A11'] = 'Interest rate (% p.a.)'
    ws['B11'] = 0.065
    ws['B11'].number_format = '0.0%'
    ws['B11'].fill = BLUE_INPUT

    ws['A12'] = 'Term (years)'
    ws['B12'] = 7
    ws['B12'].fill = BLUE_INPUT

    ws['A13'] = 'Repayment frequency'
    ws['B13'] = 'Quarterly'

    ws['A14'] = 'Principal per repayment'
    ws['B14'] = '=$B$9/($B$12*4)'
    ws['B14'].number_format = '$#,##0'
    ws['B14'].font = Font(bold=True)

    # Monthly schedule
    ws['A17'] = 'MONTHLY SCHEDULE'
    ws['A17'].font = Font(bold=True, size=11)
    ws['A17'].fill = YELLOW_OUTPUT

    schedule_labels = [
        'Opening balance',
        'Drawdown',
        'Interest expense',
        'Principal repayment',
        'Closing balance'
    ]

    for idx, label in enumerate(schedule_labels, start=18):
        ws[f'A{idx}'] = label

    # Add Month 0 (before start)
    ws['E17'] = 'M0'
    ws['E17'].font = Font(bold=True, size=9)

    ws['E18'] = 0  # Opening = 0
    ws['E19'] = 0  # Drawdown = 0
    ws['E20'] = 0  # Interest = 0
    ws['E21'] = 0  # Repayment = 0
    ws['E22'] = 0  # Closing = 0

    for r in range(18, 23):
        ws[f'E{r}'].number_format = '$#,##0'

    # Formulas for each month
    for month in range(1, MONTHS + 1):
        col = TIMELINE_START_COL + month - 1
        col_letter = get_column_letter(col)
        prev_col = get_column_letter(col - 1) if month > 1 else 'E'

        # Opening balance
        ws[f'{col_letter}18'] = f'={prev_col}22'
        ws[f'{col_letter}18'].number_format = '$#,##0'

        # Drawdown
        ws[f'{col_letter}19'] = f'=IF({month}=$B$10,$B$9,0)'
        ws[f'{col_letter}19'].number_format = '$#,##0'

        # Interest (on average balance)
        ws[f'{col_letter}20'] = f'=AVERAGE({col_letter}18,{col_letter}22)*($B$11/12)'
        ws[f'{col_letter}20'].number_format = '$#,##0'

        # Principal repayment (quarterly: months 3, 6, 9, ...)
        ws[f'{col_letter}21'] = f'=IF(MOD({month},3)=0,$B$14,0)'
        ws[f'{col_letter}21'].number_format = '$#,##0'

        # Closing balance
        ws[f'{col_letter}22'] = f'={col_letter}18+{col_letter}19-{col_letter}21'
        ws[f'{col_letter}22'].number_format = '$#,##0'
        ws[f'{col_letter}22'].fill = YELLOW_OUTPUT

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 18

    print("  ✓ Debt_Schedule sheet complete")


def build_equity_sheet(wb):
    """Build Equity sheet"""

    print("Building Equity sheet...")
    ws = wb['Equity']

    # Title
    ws['A1'] = 'EQUITY'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = DARK_BLUE_HEADER
    ws.merge_cells('A1:E1')

    add_timeline_header(ws, start_col=TIMELINE_START_COL, num_months=MONTHS)

    # Equity injection plan
    ws['A8'] = 'Equity Injections'
    ws['A8'].font = Font(bold=True, size=11)

    headers = ['Description', 'Month', 'Amount ($)']
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=9, column=col_idx)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = DARK_BLUE_HEADER

    equity_injections = [
        ('Initial equity', 1, 40000000),
    ]

    row = 10
    for desc, month, amount in equity_injections:
        ws[f'A{row}'] = desc
        ws[f'B{row}'] = month
        ws[f'B{row}'].fill = BLUE_INPUT
        ws[f'C{row}'] = amount
        ws[f'C{row}'].number_format = '$#,##0'
        ws[f'C{row}'].fill = BLUE_INPUT
        row += 1

    # Roll-forward
    ws['A14'] = 'EQUITY ROLL-FORWARD'
    ws['A14'].font = Font(bold=True, size=11)
    ws['A14'].fill = YELLOW_OUTPUT

    rollforward_labels = [
        'Opening equity',
        'Equity injection',
        'Retained earnings (Net Income)',
        'Dividends',
        'Closing equity'
    ]

    for idx, label in enumerate(rollforward_labels, start=15):
        ws[f'A{idx}'] = label

    # Month 0
    ws['E14'] = 'M0'
    ws['E14'].font = Font(bold=True, size=9)
    ws['E15'] = 0  # Opening = 0
    ws['E16'] = 0  # Injection = 0
    ws['E17'] = 0  # Retained = 0
    ws['E18'] = 0  # Dividends = 0
    ws['E19'] = 0  # Closing = 0

    for r in range(15, 20):
        ws[f'E{r}'].number_format = '$#,##0'

    # Formulas for each month
    for month in range(1, MONTHS + 1):
        col = TIMELINE_START_COL + month - 1
        col_letter = get_column_letter(col)
        prev_col = get_column_letter(col - 1) if month > 1 else 'E'

        # Opening
        ws[f'{col_letter}15'] = f'={prev_col}19'
        ws[f'{col_letter}15'].number_format = '$#,##0'

        # Injection
        injection_sum = '+'.join([f'IF(B{r}={month},C{r},0)' for r in range(10, row)])
        ws[f'{col_letter}16'] = f'={injection_sum}'
        ws[f'{col_letter}16'].number_format = '$#,##0'

        # Retained earnings (will link to P&L Net Income when P&L is built)
        ws[f'{col_letter}17'] = 0  # Placeholder - will be =P&L!NetIncome
        ws[f'{col_letter}17'].number_format = '$#,##0'

        # Dividends
        ws[f'{col_letter}18'] = 0  # No dividends per requirements
        ws[f'{col_letter}18'].number_format = '$#,##0'

        # Closing
        ws[f'{col_letter}19'] = f'={col_letter}15+{col_letter}16+{col_letter}17-{col_letter}18'
        ws[f'{col_letter}19'].number_format = '$#,##0'
        ws[f'{col_letter}19'].fill = YELLOW_OUTPUT

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 18

    print("  ✓ Equity sheet complete")


def main():
    """Main execution"""
    print("\n" + "="*60)
    print("Building Capital & Funding Sheets")
    print("="*60 + "\n")

    wb = load_workbook('Brewery_Financial_Model.xlsx')

    build_capex_schedule_sheet(wb)
    build_working_capital_sheet(wb)
    build_debt_schedule_sheet(wb)
    build_equity_sheet(wb)

    wb.save('Brewery_Financial_Model.xlsx')

    print("\n✅ Capital & Funding sheets complete!")
    print("\n   - Capex schedule with depreciation")
    print("   - Working capital assumptions")
    print("   - Debt schedule (term loan with quarterly repayments)")
    print("   - Equity injections and roll-forward")
    print("\nNext step: Run 05_statements.py")


if __name__ == '__main__':
    main()
