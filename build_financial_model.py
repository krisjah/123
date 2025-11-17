#!/usr/bin/env python3
"""
3-Statement Financial Model Generator
Creates a fully integrated financial model with control accounts for all balance sheet items
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime

def create_financial_model():
    """Create a comprehensive 3-statement financial model"""
    wb = Workbook()

    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Create all sheets
    inputs_sheet = wb.create_sheet("Inputs", 0)
    workings_sheet = wb.create_sheet("Workings", 1)
    income_statement = wb.create_sheet("Income Statement", 2)
    balance_sheet = wb.create_sheet("Balance Sheet", 3)
    cash_flow = wb.create_sheet("Cash Flow Statement", 4)

    # Define colors
    color_header = "366092"  # Dark blue
    color_input = "FDE9D9"   # Light orange
    color_calc = "DCE6F1"    # Light blue
    color_total = "B8CCE4"   # Medium blue

    # Define timeline (12 months)
    months = 12
    start_date = datetime.date(2025, 1, 1)

    # ====================
    # INPUTS SHEET
    # ====================
    print("Building Inputs sheet...")
    build_inputs_sheet(inputs_sheet, months, start_date, color_header, color_input)

    # ====================
    # WORKINGS SHEET
    # ====================
    print("Building Workings sheet...")
    build_workings_sheet(workings_sheet, months, color_header, color_calc, color_total)

    # ====================
    # INCOME STATEMENT
    # ====================
    print("Building Income Statement...")
    build_income_statement(income_statement, months, color_header, color_calc, color_total)

    # ====================
    # BALANCE SHEET
    # ====================
    print("Building Balance Sheet...")
    build_balance_sheet(balance_sheet, months, color_header, color_calc, color_total)

    # ====================
    # CASH FLOW STATEMENT
    # ====================
    print("Building Cash Flow Statement...")
    build_cash_flow_statement(cash_flow, months, color_header, color_calc, color_total)

    # Save workbook
    filename = "3_Statement_Financial_Model.xlsx"
    wb.save(filename)
    print(f"Financial model saved as {filename}")
    return filename

def build_inputs_sheet(ws, months, start_date, color_header, color_input):
    """Build the Inputs sheet with all assumptions"""

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 15

    # Set up month columns
    for i in range(months):
        col = get_column_letter(i + 3)
        ws.column_dimensions[col].width = 12

    # Header
    ws['A1'] = "FINANCIAL MODEL INPUTS"
    ws['A1'].font = Font(bold=True, size=14)

    # Timeline row
    ws['A3'] = "Timeline"
    ws['A3'].font = Font(bold=True)
    ws['B3'] = "Unit"
    ws['B3'].font = Font(bold=True)
    apply_fill(ws['A3'], color_header)
    apply_fill(ws['B3'], color_header)

    for i in range(months):
        col = get_column_letter(i + 3)
        month_date = add_months(start_date, i)
        ws[f'{col}3'] = month_date.strftime('%b-%y')
        ws[f'{col}3'].font = Font(bold=True)
        apply_fill(ws[f'{col}3'], color_header)

    row = 5

    # REVENUE ASSUMPTIONS
    ws[f'A{row}'] = "REVENUE ASSUMPTIONS"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws[f'A{row}'] = "Monthly Revenue Growth"
    ws[f'B{row}'] = "%"
    ws[f'C{row}'] = 0.05  # 5% monthly growth
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 1

    ws[f'A{row}'] = "Base Revenue (Month 1)"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 100000
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Revenue"
    ws[f'B{row}'] = "$"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = f'=C{row-1}'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row}*(1+$C${row-2})'
    row += 2

    # COST OF GOODS SOLD
    ws[f'A{row}'] = "COST OF GOODS SOLD"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws[f'A{row}'] = "COGS % of Revenue"
    ws[f'B{row}'] = "%"
    ws[f'C{row}'] = 0.40  # 40%
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 2

    # OPERATING EXPENSES
    ws[f'A{row}'] = "OPERATING EXPENSES"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws[f'A{row}'] = "Salaries & Wages"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 25000
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 1

    ws[f'A{row}'] = "Rent"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 5000
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 1

    ws[f'A{row}'] = "Marketing"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 8000
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 1

    ws[f'A{row}'] = "Other Operating Expenses"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 3000
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 2

    # WORKING CAPITAL ASSUMPTIONS
    ws[f'A{row}'] = "WORKING CAPITAL"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws[f'A{row}'] = "AR Days"
    ws[f'B{row}'] = "days"
    ws[f'C{row}'] = 45
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 1

    ws[f'A{row}'] = "Inventory Days"
    ws[f'B{row}'] = "days"
    ws[f'C{row}'] = 30
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 1

    ws[f'A{row}'] = "AP Days"
    ws[f'B{row}'] = "days"
    ws[f'C{row}'] = 30
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 1

    ws[f'A{row}'] = "Prepayments (months)"
    ws[f'B{row}'] = "months"
    ws[f'C{row}'] = 1
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 1

    ws[f'A{row}'] = "Accrued Expenses % of OpEx"
    ws[f'B{row}'] = "%"
    ws[f'C{row}'] = 0.15
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 1

    ws[f'A{row}'] = "Deferred Revenue (months)"
    ws[f'B{row}'] = "months"
    ws[f'C{row}'] = 1
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 2

    # CAPITAL EXPENDITURE
    ws[f'A{row}'] = "CAPITAL EXPENDITURE"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws[f'A{row}'] = "Monthly CapEx"
    ws[f'B{row}'] = "$"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = 50000
        else:
            ws[f'{col}{row}'] = 0
        apply_fill(ws[f'{col}{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Depreciation Rate (annual)"
    ws[f'B{row}'] = "%"
    ws[f'C{row}'] = 0.20  # 20% annual
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 1

    ws[f'A{row}'] = "Monthly Depreciation Rate"
    ws[f'B{row}'] = "%"
    ws[f'C{row}'] = '=C{}/12'.format(row-1)
    row += 2

    # DEBT
    ws[f'A{row}'] = "DEBT"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws[f'A{row}'] = "Interest Rate (annual)"
    ws[f'B{row}'] = "%"
    ws[f'C{row}'] = 0.06
    apply_fill(ws[f'C{row}'], color_input)
    ws.merge_cells(f'C{row}:N{row}')
    row += 1

    ws[f'A{row}'] = "Monthly Interest Rate"
    ws[f'B{row}'] = "%"
    ws[f'C{row}'] = '=C{}/12'.format(row-1)
    row += 1

    ws[f'A{row}'] = "New Borrowing"
    ws[f'B{row}'] = "$"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = 200000
        else:
            ws[f'{col}{row}'] = 0
        apply_fill(ws[f'{col}{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Debt Repayment"
    ws[f'B{row}'] = "$"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = 5000
        apply_fill(ws[f'{col}{row}'], color_input)
    row += 2

    # EQUITY
    ws[f'A{row}'] = "EQUITY"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws[f'A{row}'] = "Share Capital Injection"
    ws[f'B{row}'] = "$"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = 100000
        else:
            ws[f'{col}{row}'] = 0
        apply_fill(ws[f'{col}{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Dividends Paid"
    ws[f'B{row}'] = "$"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = 0
        apply_fill(ws[f'{col}{row}'], color_input)
    row += 2

    # OPENING BALANCES
    ws[f'A{row}'] = "OPENING BALANCES"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws[f'A{row}'] = "Opening Cash"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 50000
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Opening AR"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 0
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Opening Inventory"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 0
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Opening Prepayments"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 0
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Opening Fixed Assets (Gross)"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 0
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Opening Accumulated Depreciation"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 0
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Opening AP"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 0
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Opening Accrued Expenses"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 0
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Opening Deferred Revenue"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 0
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Opening Debt"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 0
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Opening Share Capital"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 0
    apply_fill(ws[f'C{row}'], color_input)
    row += 1

    ws[f'A{row}'] = "Opening Retained Earnings"
    ws[f'B{row}'] = "$"
    ws[f'C{row}'] = 0
    apply_fill(ws[f'C{row}'], color_input)

def build_workings_sheet(ws, months, color_header, color_calc, color_total):
    """Build the Workings sheet with all control accounts"""

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 15

    for i in range(months):
        col = get_column_letter(i + 3)
        ws.column_dimensions[col].width = 12

    # Header
    ws['A1'] = "CONTROL ACCOUNTS - WORKINGS"
    ws['A1'].font = Font(bold=True, size=14)

    # Timeline
    ws['A3'] = "Timeline"
    ws['A3'].font = Font(bold=True)
    apply_fill(ws['A3'], color_header)

    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}3'] = f'=Inputs!{col}3'
        ws[f'{col}3'].font = Font(bold=True)
        apply_fill(ws[f'{col}3'], color_header)

    row = 5

    # ===============================
    # CASH CONTROL ACCOUNT
    # ===============================
    ws[f'A{row}'] = "CASH CONTROL ACCOUNT"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$59'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+11}'  # Closing balance from previous month
    row += 1

    ws[f'A{row}'] = "Cash Inflows:"
    ws[f'A{row}'].font = Font(italic=True)
    row += 1

    ws[f'A{row}'] = "  Cash from Customers"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}32'  # Links to AR control account
    row += 1

    ws[f'A{row}'] = "  New Borrowing"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$52'
    row += 1

    ws[f'A{row}'] = "  Share Capital Injection"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$57'
    row += 1

    ws[f'A{row}'] = "Total Inflows"
    ws[f'A{row}'].font = Font(bold=True)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=SUM({col}{row-3}:{col}{row-1})'
    row += 1

    ws[f'A{row}'] = "Cash Outflows:"
    ws[f'A{row}'].font = Font(italic=True)
    row += 1

    ws[f'A{row}'] = "  Payments to Suppliers"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}69'  # Links to AP control account
    row += 1

    ws[f'A{row}'] = "  Operating Expenses Paid"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}90'  # Links to accrued expenses
    row += 1

    ws[f'A{row}'] = "  Capital Expenditure"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$44'
    row += 1

    ws[f'A{row}'] = "  Interest Paid"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}118'  # Links to debt control account
    row += 1

    ws[f'A{row}'] = "  Debt Repayment"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$53'
    row += 1

    ws[f'A{row}'] = "  Dividends Paid"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$58'
    row += 1

    ws[f'A{row}'] = "Total Outflows"
    ws[f'A{row}'].font = Font(bold=True)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=SUM({col}{row-6}:{col}{row-1})'
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-13}+{col}{row-7}-{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # ACCOUNTS RECEIVABLE CONTROL
    # ===============================
    ws[f'A{row}'] = "ACCOUNTS RECEIVABLE CONTROL"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$60'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+3}'
    row += 1

    ws[f'A{row}'] = "Credit Sales (Revenue)"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$9'
    row += 1

    ws[f'A{row}'] = "Cash Received from Customers"
    for i in range(months):
        col = get_column_letter(i + 3)
        # AR collection based on AR days
        ws[f'{col}{row}'] = f'=({col}{row-2}*30)/(Inputs!$C$23+30)'
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-3}+{col}{row-2}-{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # INVENTORY CONTROL
    # ===============================
    ws[f'A{row}'] = "INVENTORY CONTROL"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$61'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+4}'
    row += 1

    ws[f'A{row}'] = "Purchases"
    for i in range(months):
        col = get_column_letter(i + 3)
        # Purchases to maintain inventory days
        ws[f'{col}{row}'] = f'=(Inputs!{col}$9*Inputs!$C$12+Inputs!$C$24*30)/30'
    row += 1

    ws[f'A{row}'] = "COGS Withdrawal"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$9*Inputs!$C$12'
    row += 1

    ws[f'A{row}'] = "Adjustments"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = 0
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-4}+{col}{row-3}-{col}{row-2}+{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # PREPAYMENTS CONTROL
    # ===============================
    ws[f'A{row}'] = "PREPAYMENTS CONTROL"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$62'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+3}'
    row += 1

    ws[f'A{row}'] = "New Prepayments (cash paid)"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=(Inputs!$C$17+Inputs!$C$18+Inputs!$C$19)*Inputs!$C$26'
    row += 1

    ws[f'A{row}'] = "Amortization (expense)"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-2}+{col}{row-1}'
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-3}+{col}{row-2}-{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # ACCOUNTS PAYABLE CONTROL
    # ===============================
    ws[f'A{row}'] = "ACCOUNTS PAYABLE CONTROL"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$65'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+3}'
    row += 1

    ws[f'A{row}'] = "Purchases on Credit"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}41'  # Links to Inventory purchases
    row += 1

    ws[f'A{row}'] = "Cash Payments to Suppliers"
    for i in range(months):
        col = get_column_letter(i + 3)
        # Payment based on AP days
        ws[f'{col}{row}'] = f'=({col}{row-2}*30)/(Inputs!$C$25+30)'
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-3}+{col}{row-2}-{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # ACCRUED EXPENSES CONTROL
    # ===============================
    ws[f'A{row}'] = "ACCRUED EXPENSES CONTROL"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$66'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+3}'
    row += 1

    ws[f'A{row}'] = "Operating Expenses Incurred"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=(Inputs!$C$15+Inputs!$C$17+Inputs!$C$18+Inputs!$C$19)*Inputs!$C$27'
    row += 1

    ws[f'A{row}'] = "Cash Paid for OpEx"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-2}+{col}{row-1}'
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-3}+{col}{row-2}-{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # DEFERRED REVENUE CONTROL
    # ===============================
    ws[f'A{row}'] = "DEFERRED REVENUE CONTROL"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$67'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+3}'
    row += 1

    ws[f'A{row}'] = "Cash Received in Advance"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$9*Inputs!$C$28'
    row += 1

    ws[f'A{row}'] = "Revenue Recognized"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-2}+{col}{row-1}'
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-3}+{col}{row-2}-{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # FIXED ASSETS CONTROL (GROSS)
    # ===============================
    ws[f'A{row}'] = "FIXED ASSETS - GROSS CONTROL"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$63'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+3}'
    row += 1

    ws[f'A{row}'] = "Capital Expenditure"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$44'
    row += 1

    ws[f'A{row}'] = "Disposals"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = 0
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-3}+{col}{row-2}-{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # ACCUMULATED DEPRECIATION
    # ===============================
    ws[f'A{row}'] = "ACCUMULATED DEPRECIATION CONTROL"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$64'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+2}'
    row += 1

    ws[f'A{row}'] = "Depreciation Expense"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-6}*Inputs!$C$47'
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-2}+{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # DEBT CONTROL
    # ===============================
    ws[f'A{row}'] = "DEBT CONTROL"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$68'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+4}'
    row += 1

    ws[f'A{row}'] = "New Borrowing"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$52'
    row += 1

    ws[f'A{row}'] = "Interest Accrued"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-2}*Inputs!$C$51'
    row += 1

    ws[f'A{row}'] = "Interest Paid"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-1}'
    row += 1

    ws[f'A{row}'] = "Principal Repayment"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$53'
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-5}+{col}{row-4}+{col}{row-3}-{col}{row-2}-{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # SHARE CAPITAL CONTROL
    # ===============================
    ws[f'A{row}'] = "SHARE CAPITAL CONTROL"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$69'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+2}'
    row += 1

    ws[f'A{row}'] = "New Share Capital"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$57'
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-2}+{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # RETAINED EARNINGS CONTROL
    # ===============================
    ws[f'A{row}'] = "RETAINED EARNINGS CONTROL"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Opening Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$70'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+3}'
    row += 1

    ws[f'A{row}'] = "Net Income"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=\'Income Statement\'!{col}20'  # Links to Net Income
    row += 1

    ws[f'A{row}'] = "Dividends"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$58'
    row += 1

    ws[f'A{row}'] = "Closing Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-3}+{col}{row-2}-{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)

def build_income_statement(ws, months, color_header, color_calc, color_total):
    """Build the Income Statement"""

    ws.column_dimensions['A'].width = 30

    for i in range(months):
        col = get_column_letter(i + 3)
        ws.column_dimensions[col].width = 12

    # Header
    ws['A1'] = "INCOME STATEMENT"
    ws['A1'].font = Font(bold=True, size=14)

    # Timeline
    ws['A3'] = "Period"
    ws['A3'].font = Font(bold=True)
    apply_fill(ws['A3'], color_header)

    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}3'] = f'=Inputs!{col}3'
        ws[f'{col}3'].font = Font(bold=True)
        apply_fill(ws[f'{col}3'], color_header)

    row = 5

    # Revenue
    ws[f'A{row}'] = "Revenue"
    ws[f'A{row}'].font = Font(bold=True)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$9'
    row += 1

    # COGS
    ws[f'A{row}'] = "Cost of Goods Sold"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}42'  # Links to inventory COGS
    row += 1

    # Gross Profit
    ws[f'A{row}'] = "Gross Profit"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_calc)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-2}-{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_calc)
    row += 2

    # Operating Expenses
    ws[f'A{row}'] = "Operating Expenses:"
    ws[f'A{row}'].font = Font(bold=True)
    row += 1

    ws[f'A{row}'] = "Salaries & Wages"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!$C$15'
    row += 1

    ws[f'A{row}'] = "Rent"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!$C$17'
    row += 1

    ws[f'A{row}'] = "Marketing"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!$C$18'
    row += 1

    ws[f'A{row}'] = "Other Operating Expenses"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!$C$19'
    row += 1

    ws[f'A{row}'] = "Depreciation"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}112'  # Links to depreciation
    row += 1

    ws[f'A{row}'] = "Total Operating Expenses"
    ws[f'A{row}'].font = Font(bold=True)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=SUM({col}{row-5}:{col}{row-1})'
    row += 2

    # EBIT
    ws[f'A{row}'] = "EBIT"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_calc)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-9}-{col}{row-2}'
        apply_fill(ws[f'{col}{row}'], color_calc)
    row += 2

    # Interest Expense
    ws[f'A{row}'] = "Interest Expense"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}117'  # Links to debt interest
    row += 1

    # EBT
    ws[f'A{row}'] = "Earnings Before Tax"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_calc)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-3}-{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_calc)
    row += 2

    # Tax (assuming 0% for simplicity, can add tax rate to inputs)
    ws[f'A{row}'] = "Income Tax"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = 0
    row += 2

    # Net Income
    ws[f'A{row}'] = "NET INCOME"
    ws[f'A{row}'].font = Font(bold=True, size=12)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-4}-{col}{row-2}'
        apply_fill(ws[f'{col}{row}'], color_total)

def build_balance_sheet(ws, months, color_header, color_calc, color_total):
    """Build the Balance Sheet"""

    ws.column_dimensions['A'].width = 30

    for i in range(months):
        col = get_column_letter(i + 3)
        ws.column_dimensions[col].width = 12

    # Header
    ws['A1'] = "BALANCE SHEET"
    ws['A1'].font = Font(bold=True, size=14)

    # Timeline
    ws['A3'] = "Period Ending"
    ws['A3'].font = Font(bold=True)
    apply_fill(ws['A3'], color_header)

    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}3'] = f'=Inputs!{col}3'
        ws[f'{col}3'].font = Font(bold=True)
        apply_fill(ws[f'{col}3'], color_header)

    row = 5

    # ===============================
    # ASSETS
    # ===============================
    ws[f'A{row}'] = "ASSETS"
    ws[f'A{row}'].font = Font(bold=True, size=12, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Current Assets:"
    ws[f'A{row}'].font = Font(bold=True)
    row += 1

    ws[f'A{row}'] = "Cash"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}18'  # Cash closing balance
    row += 1

    ws[f'A{row}'] = "Accounts Receivable"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}32'  # AR closing balance
    row += 1

    ws[f'A{row}'] = "Inventory"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}45'  # Inventory closing balance
    row += 1

    ws[f'A{row}'] = "Prepayments"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}54'  # Prepayments closing balance
    row += 1

    ws[f'A{row}'] = "Total Current Assets"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_calc)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=SUM({col}{row-4}:{col}{row-1})'
        apply_fill(ws[f'{col}{row}'], color_calc)
    row += 2

    ws[f'A{row}'] = "Non-Current Assets:"
    ws[f'A{row}'].font = Font(bold=True)
    row += 1

    ws[f'A{row}'] = "Fixed Assets (Gross)"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}105'  # Fixed assets gross
    row += 1

    ws[f'A{row}'] = "Accumulated Depreciation"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=-Workings!{col}112'  # Accumulated depreciation (negative)
    row += 1

    ws[f'A{row}'] = "Fixed Assets (Net)"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_calc)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-2}+{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_calc)
    row += 2

    ws[f'A{row}'] = "TOTAL ASSETS"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-9}+{col}{row-2}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # LIABILITIES
    # ===============================
    ws[f'A{row}'] = "LIABILITIES"
    ws[f'A{row}'].font = Font(bold=True, size=12, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Current Liabilities:"
    ws[f'A{row}'].font = Font(bold=True)
    row += 1

    ws[f'A{row}'] = "Accounts Payable"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}69'  # AP closing balance
    row += 1

    ws[f'A{row}'] = "Accrued Expenses"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}90'  # Accrued expenses closing balance
    row += 1

    ws[f'A{row}'] = "Deferred Revenue"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}97'  # Deferred revenue closing balance
    row += 1

    ws[f'A{row}'] = "Total Current Liabilities"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_calc)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=SUM({col}{row-3}:{col}{row-1})'
        apply_fill(ws[f'{col}{row}'], color_calc)
    row += 2

    ws[f'A{row}'] = "Non-Current Liabilities:"
    ws[f'A{row}'].font = Font(bold=True)
    row += 1

    ws[f'A{row}'] = "Long-term Debt"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}123'  # Debt closing balance
    row += 1

    ws[f'A{row}'] = "Total Non-Current Liabilities"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_calc)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_calc)
    row += 2

    ws[f'A{row}'] = "TOTAL LIABILITIES"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-8}+{col}{row-2}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # ===============================
    # EQUITY
    # ===============================
    ws[f'A{row}'] = "EQUITY"
    ws[f'A{row}'].font = Font(bold=True, size=12, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Share Capital"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}128'  # Share capital closing balance
    row += 1

    ws[f'A{row}'] = "Retained Earnings"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}134'  # Retained earnings closing balance
    row += 1

    ws[f'A{row}'] = "TOTAL EQUITY"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=SUM({col}{row-2}:{col}{row-1})'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # Total Liabilities + Equity
    ws[f'A{row}'] = "TOTAL LIABILITIES + EQUITY"
    ws[f'A{row}'].font = Font(bold=True, size=11)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-12}+{col}{row-2}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # Balance Check
    ws[f'A{row}'] = "BALANCE CHECK (should be 0)"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], "FFD700")  # Gold color
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-18}-{col}{row-2}'
        apply_fill(ws[f'{col}{row}'], "FFD700")

def build_cash_flow_statement(ws, months, color_header, color_calc, color_total):
    """Build the Cash Flow Statement (Indirect Method)"""

    ws.column_dimensions['A'].width = 35

    for i in range(months):
        col = get_column_letter(i + 3)
        ws.column_dimensions[col].width = 12

    # Header
    ws['A1'] = "CASH FLOW STATEMENT (INDIRECT METHOD)"
    ws['A1'].font = Font(bold=True, size=14)

    # Timeline
    ws['A3'] = "Period"
    ws['A3'].font = Font(bold=True)
    apply_fill(ws['A3'], color_header)

    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}3'] = f'=Inputs!{col}3'
        ws[f'{col}3'].font = Font(bold=True)
        apply_fill(ws[f'{col}3'], color_header)

    row = 5

    # ===============================
    # OPERATING ACTIVITIES
    # ===============================
    ws[f'A{row}'] = "CASH FROM OPERATING ACTIVITIES"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Net Income"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=\'Income Statement\'!{col}20'
    row += 1

    ws[f'A{row}'] = "Adjustments:"
    ws[f'A{row}'].font = Font(italic=True)
    row += 1

    ws[f'A{row}'] = "  Depreciation"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Workings!{col}112'
    row += 1

    ws[f'A{row}'] = "Changes in Working Capital:"
    ws[f'A{row}'].font = Font(italic=True)
    row += 1

    ws[f'A{row}'] = "  (Increase)/Decrease in AR"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = f'=Inputs!$C$60-\'Balance Sheet\'!{col}8'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'=\'Balance Sheet\'!{prev_col}8-\'Balance Sheet\'!{col}8'
    row += 1

    ws[f'A{row}'] = "  (Increase)/Decrease in Inventory"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = f'=Inputs!$C$61-\'Balance Sheet\'!{col}9'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'=\'Balance Sheet\'!{prev_col}9-\'Balance Sheet\'!{col}9'
    row += 1

    ws[f'A{row}'] = "  (Increase)/Decrease in Prepayments"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = f'=Inputs!$C$62-\'Balance Sheet\'!{col}10'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'=\'Balance Sheet\'!{prev_col}10-\'Balance Sheet\'!{col}10'
    row += 1

    ws[f'A{row}'] = "  Increase/(Decrease) in AP"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = f'=\'Balance Sheet\'!{col}24-Inputs!$C$65'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'=\'Balance Sheet\'!{col}24-\'Balance Sheet\'!{prev_col}24'
    row += 1

    ws[f'A{row}'] = "  Increase/(Decrease) in Accrued Exp"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = f'=\'Balance Sheet\'!{col}25-Inputs!$C$66'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'=\'Balance Sheet\'!{col}25-\'Balance Sheet\'!{prev_col}25'
    row += 1

    ws[f'A{row}'] = "  Increase/(Decrease) in Deferred Rev"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = f'=\'Balance Sheet\'!{col}26-Inputs!$C$67'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'=\'Balance Sheet\'!{col}26-\'Balance Sheet\'!{prev_col}26'
    row += 1

    ws[f'A{row}'] = "Net Cash from Operating Activities"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_calc)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=SUM({col}{row-11}:{col}{row-1})'
        apply_fill(ws[f'{col}{row}'], color_calc)
    row += 2

    # ===============================
    # INVESTING ACTIVITIES
    # ===============================
    ws[f'A{row}'] = "CASH FROM INVESTING ACTIVITIES"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "Capital Expenditure"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=-Inputs!{col}$44'
    row += 1

    ws[f'A{row}'] = "Proceeds from Asset Disposals"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = 0
    row += 1

    ws[f'A{row}'] = "Net Cash from Investing Activities"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_calc)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=SUM({col}{row-2}:{col}{row-1})'
        apply_fill(ws[f'{col}{row}'], color_calc)
    row += 2

    # ===============================
    # FINANCING ACTIVITIES
    # ===============================
    ws[f'A{row}'] = "CASH FROM FINANCING ACTIVITIES"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FFFFFF")
    apply_fill(ws[f'A{row}'], color_header)
    row += 1

    ws[f'A{row}'] = "New Borrowing"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$52'
    row += 1

    ws[f'A{row}'] = "Debt Repayment"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=-Inputs!{col}$53'
    row += 1

    ws[f'A{row}'] = "Interest Paid"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=-Workings!{col}118'
    row += 1

    ws[f'A{row}'] = "Share Capital Raised"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=Inputs!{col}$57'
    row += 1

    ws[f'A{row}'] = "Dividends Paid"
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=-Inputs!{col}$58'
    row += 1

    ws[f'A{row}'] = "Net Cash from Financing Activities"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_calc)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'=SUM({col}{row-5}:{col}{row-1})'
        apply_fill(ws[f'{col}{row}'], color_calc)
    row += 2

    # ===============================
    # NET CHANGE IN CASH
    # ===============================
    ws[f'A{row}'] = "Net Increase/(Decrease) in Cash"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-16}+{col}{row-9}+{col}{row-2}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    ws[f'A{row}'] = "Opening Cash Balance"
    for i in range(months):
        col = get_column_letter(i + 3)
        if i == 0:
            ws[f'{col}{row}'] = '=Inputs!$C$59'
        else:
            prev_col = get_column_letter(i + 2)
            ws[f'{col}{row}'] = f'={prev_col}{row+1}'
    row += 1

    ws[f'A{row}'] = "Closing Cash Balance"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], color_total)
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-3}+{col}{row-1}'
        apply_fill(ws[f'{col}{row}'], color_total)
    row += 2

    # Cash Reconciliation Check
    ws[f'A{row}'] = "CASH CHECK (should be 0)"
    ws[f'A{row}'].font = Font(bold=True)
    apply_fill(ws[f'A{row}'], "FFD700")
    for i in range(months):
        col = get_column_letter(i + 3)
        ws[f'{col}{row}'] = f'={col}{row-1}-\'Balance Sheet\'!{col}7'
        apply_fill(ws[f'{col}{row}'], "FFD700")

def apply_fill(cell, color):
    """Apply fill color to a cell"""
    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

def add_months(start_date, months):
    """Add months to a date"""
    month = start_date.month - 1 + months
    year = start_date.year + month // 12
    month = month % 12 + 1
    day = min(start_date.day, [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month-1])
    return datetime.date(year, month, day)

if __name__ == "__main__":
    print("Building 3-Statement Financial Model...")
    print("=" * 60)
    create_financial_model()
    print("=" * 60)
    print("Model complete!")
