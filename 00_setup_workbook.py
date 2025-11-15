#!/usr/bin/env python3
"""
Beer Brewery Financial Model - Setup Script
Creates the base workbook structure with all 19 sheets

Run this first before any other scripts.

Requirements:
    pip install openpyxl
"""

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

def create_workbook():
    """Create new workbook with all required sheets in correct order"""

    print("Creating brewery financial model workbook...")

    # Create workbook
    wb = Workbook()

    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Define all sheets in order
    sheet_names = [
        'Cover',
        'Control',
        'Assumptions_Global',
        'Product_Matrix',
        'Volume_Assumptions',
        'Pricing_Revenue',
        'Production_Capacity',
        'COGS_Buildup',
        'Logistics_Distribution',
        'SG&A',
        'Capex_Schedule',
        'Working_Capital',
        'Debt_Schedule',
        'Equity',
        'P&L',
        'CashFlow',
        'BalanceSheet',
        'KPI_Dashboard',
        'Checks_Errors',
        'Documentation',
        'Change_Log'
    ]

    # Create all sheets
    for sheet_name in sheet_names:
        wb.create_sheet(sheet_name)
        print(f"  ✓ Created sheet: {sheet_name}")

    # Save workbook
    filename = 'Brewery_Financial_Model.xlsx'
    wb.save(filename)
    print(f"\n✅ Workbook created: {filename}")
    print(f"   Sheets: {len(sheet_names)}")
    print(f"\nNext step: Run 01_control_assumptions.py")

    return filename

if __name__ == '__main__':
    create_workbook()
