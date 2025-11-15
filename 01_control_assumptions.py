#!/usr/bin/env python3
"""
Beer Brewery Financial Model - Control & Assumptions
Builds: Control, Assumptions_Global, Product_Matrix sheets

Run after: 00_setup_workbook.py
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import date

# Color definitions
BLUE_INPUT = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
GREY_CALC = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
DARK_BLUE_HEADER = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

def build_control_sheet(wb):
    """Build Control sheet with scenario selector"""

    print("Building Control sheet...")
    ws = wb['Control']

    # Title
    ws['A1'] = 'CONTROL PANEL'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = DARK_BLUE_HEADER
    ws.merge_cells('A1:D1')

    # Model Timeline section
    ws['A3'] = 'Model Timeline'
    ws['A3'].font = Font(bold=True, size=12)

    ws['A4'] = 'Model start date'
    ws['B4'] = date(2025, 1, 1)  # Default start date
    ws['B4'].number_format = 'dd-mmm-yyyy'
    ws['B4'].fill = BLUE_INPUT
    ws['C4'] = '← INPUT'
    ws['C4'].font = Font(italic=True)

    ws['A5'] = 'Projection period (months)'
    ws['B5'] = 120
    ws['B5'].fill = BLUE_INPUT

    ws['A6'] = 'Model end date'
    ws['B6'] = '=EDATE(B4,B5-1)'
    ws['B6'].number_format = 'mmm-yyyy'
    ws['B6'].fill = GREY_CALC
    ws['C6'] = '=DATE formula'
    ws['C6'].font = Font(italic=True)

    # Scenario Selection section
    ws['A8'] = 'Scenario Selection'
    ws['A8'].font = Font(bold=True, size=12)

    ws['A9'] = 'Active scenario'
    ws['B9'] = 'Base'
    ws['B9'].fill = BLUE_INPUT
    ws['C9'] = '← DROPDOWN'
    ws['C9'].font = Font(italic=True)

    # Add data validation for scenario
    dv = DataValidation(type="list", formula1='"Base,Upside,Downside"', allow_blank=False)
    dv.add(ws['B9'])
    ws.add_data_validation(dv)

    ws['A10'] = 'Scenario index'
    ws['B10'] = '=MATCH(B9,{"Base";"Upside";"Downside"},0)'
    ws['B10'].fill = GREY_CALC
    ws['C10'] = 'Returns 1/2/3'
    ws['C10'].font = Font(italic=True)

    # Model Settings section
    ws['A12'] = 'Model Settings'
    ws['A12'].font = Font(bold=True, size=12)

    ws['A13'] = 'Enable Excel iterations?'
    ws['B13'] = 'YES'
    ws['B13'].font = Font(bold=True)
    ws['C13'] = 'Manual check'
    ws['C13'].font = Font(italic=True)
    ws['D13'] = 'File→Options→Formulas'
    ws['D13'].font = Font(italic=True, size=9)

    ws['A14'] = 'Max iterations'
    ws['B14'] = 100

    ws['A15'] = 'Max change'
    ws['B15'] = 0.001
    ws['B15'].number_format = '0.000'

    # Set column widths
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 25

    print("  ✓ Control sheet complete")


def build_assumptions_global_sheet(wb):
    """Build Assumptions_Global sheet"""

    print("Building Assumptions_Global sheet...")
    ws = wb['Assumptions_Global']

    # Title
    ws['A1'] = 'GLOBAL ASSUMPTIONS'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = DARK_BLUE_HEADER
    ws.merge_cells('A1:E1')

    # Header row
    ws['A3'] = 'Inflation & Escalation'
    ws['A3'].font = Font(bold=True, size=11)

    headers = ['Assumption', 'Base', 'Upside', 'Downside', 'Active']
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = DARK_BLUE_HEADER
        cell.font = Font(bold=True, color="FFFFFF")

    # Inflation assumptions
    assumptions = [
        ('CPI (general inflation) % p.a.', 0.025, 0.030, 0.020),
        ('Wage inflation % p.a.', 0.030, 0.035, 0.025),
        ('Materials cost inflation % p.a.', 0.025, 0.035, 0.020),
        ('Packaging cost inflation % p.a.', 0.030, 0.040, 0.025),
        ('Utilities inflation % p.a.', 0.040, 0.050, 0.030),
    ]

    row = 5
    for label, base, upside, downside in assumptions:
        ws[f'A{row}'] = label
        ws[f'B{row}'] = base
        ws[f'C{row}'] = upside
        ws[f'D{row}'] = downside
        ws[f'E{row}'] = f'=INDEX(B{row}:D{row},1,Control!$B$10)'

        # Format percentages
        for col in ['B', 'C', 'D', 'E']:
            ws[f'{col}{row}'].number_format = '0.0%'

        # Color coding
        for col in ['B', 'C', 'D']:
            ws[f'{col}{row}'].fill = BLUE_INPUT
        ws[f'E{row}'].fill = GREY_CALC

        row += 1

    # Tax & Regulatory section
    ws[f'A{row}'] = ''
    row += 1
    ws[f'A{row}'] = 'Tax & Regulatory'
    ws[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    tax_assumptions = [
        ('Corporate tax rate %', 0.30, 0.30, 0.30),
        ('Excise duty ($/LAL)', 58.26, 58.26, 58.26),
        ('Excise shock (Year 3+)', 0.00, 0.00, 0.10),
        ('VAT/GST rate %', 0.10, 0.10, 0.10),
    ]

    for label, base, upside, downside in tax_assumptions:
        ws[f'A{row}'] = label
        ws[f'B{row}'] = base
        ws[f'C{row}'] = upside
        ws[f'D{row}'] = downside
        ws[f'E{row}'] = f'=INDEX(B{row}:D{row},1,Control!$B$10)'

        # Format based on type
        if '%' in label:
            for col in ['B', 'C', 'D', 'E']:
                ws[f'{col}{row}'].number_format = '0.0%'
        else:
            for col in ['B', 'C', 'D', 'E']:
                ws[f'{col}{row}'].number_format = '$#,##0.00'

        # Color coding
        for col in ['B', 'C', 'D']:
            ws[f'{col}{row}'].fill = BLUE_INPUT
        ws[f'E{row}'].fill = GREY_CALC

        row += 1

    # Discount Rate section
    ws[f'A{row}'] = ''
    row += 1
    ws[f'A{row}'] = 'Discount Rate (for NPV)'
    ws[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws[f'A{row}'] = 'WACC / Discount rate %'
    ws[f'B{row}'] = 0.12
    ws[f'C{row}'] = 0.11
    ws[f'D{row}'] = 0.13
    ws[f'E{row}'] = f'=INDEX(B{row}:D{row},1,Control!$B$10)'

    for col in ['B', 'C', 'D', 'E']:
        ws[f'{col}{row}'].number_format = '0.0%'
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].fill = BLUE_INPUT
    ws[f'E{row}'].fill = GREY_CALC

    # Set column widths
    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15

    print("  ✓ Assumptions_Global sheet complete")


def build_product_matrix_sheet(wb):
    """Build Product_Matrix reference sheet"""

    print("Building Product_Matrix sheet...")
    ws = wb['Product_Matrix']

    # Title
    ws['A1'] = 'PRODUCT MATRIX'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = DARK_BLUE_HEADER
    ws.merge_cells('A1:E1')

    # Brand List section
    ws['A3'] = 'Brand List'
    ws['A3'].font = Font(bold=True, size=11)

    brand_headers = ['Brand ID', 'Brand Name', 'Brand Type', 'ABV %', 'Recipe Type']
    for col, header in enumerate(brand_headers, start=1):
        cell = ws.cell(row=4, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = DARK_BLUE_HEADER
        cell.font = Font(bold=True, color="FFFFFF")

    brands = [
        (1, 'SkyBrew', 'Lager', 0.045, 'Core'),
        (2, 'AllDark', 'Stout', 0.058, 'Craft'),
        (3, 'GoldCoast', 'Pale Ale', 0.052, 'Craft'),
        (4, 'CityLight', 'Light Lager', 0.035, 'Core'),
        (5, 'HarborIPA', 'IPA', 0.062, 'Premium'),
    ]

    for row_idx, (bid, name, btype, abv, recipe) in enumerate(brands, start=5):
        ws[f'A{row_idx}'] = bid
        ws[f'B{row_idx}'] = name
        ws[f'C{row_idx}'] = btype
        ws[f'D{row_idx}'] = abv
        ws[f'D{row_idx}'].number_format = '0.0%'
        ws[f'E{row_idx}'] = recipe

    # Pack Type List section
    ws['A11'] = 'Pack Type List'
    ws['A11'].font = Font(bold=True, size=11)

    pack_headers = ['Pack ID', 'Pack Type', 'HL per unit', 'Packaging', 'Returnable?']
    for col, header in enumerate(pack_headers, start=1):
        cell = ws.cell(row=12, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = DARK_BLUE_HEADER
        cell.font = Font(bold=True, color="FFFFFF")

    packs = [
        (1, 'Bottles', 0.00033, 'Glass 330ml', 'No'),
        (2, 'Cans', 0.00037, 'Aluminum 375ml', 'No'),
        (3, 'Kegs', 0.50, 'Stainless 50L', 'Yes'),
    ]

    for row_idx, (pid, ptype, hl, pkg, ret) in enumerate(packs, start=13):
        ws[f'A{row_idx}'] = pid
        ws[f'B{row_idx}'] = ptype
        ws[f'C{row_idx}'] = hl
        ws[f'C{row_idx}'].number_format = '0.00000'
        ws[f'D{row_idx}'] = pkg
        ws[f'E{row_idx}'] = ret

    # Channel List section
    ws['A17'] = 'Channel List'
    ws['A17'].font = Font(bold=True, size=11)

    channel_headers = ['Channel ID', 'Channel Name', 'Customer Type']
    for col, header in enumerate(channel_headers, start=1):
        cell = ws.cell(row=18, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = DARK_BLUE_HEADER
        cell.font = Font(bold=True, color="FFFFFF")

    channels = [
        (1, 'OnTrade', 'Bars, Restaurants, Pubs'),
        (2, 'OffTrade', 'Retail, Supermarkets, Bottle shops'),
    ]

    for row_idx, (cid, cname, ctype) in enumerate(channels, start=19):
        ws[f'A{row_idx}'] = cid
        ws[f'B{row_idx}'] = cname
        ws[f'C{row_idx}'] = ctype

    # Set column widths
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 20

    print("  ✓ Product_Matrix sheet complete")


def main():
    """Main execution"""
    print("\n" + "="*60)
    print("Building Control & Assumptions Sheets")
    print("="*60 + "\n")

    # Load workbook
    wb = load_workbook('Brewery_Financial_Model.xlsx')

    # Build sheets
    build_control_sheet(wb)
    build_assumptions_global_sheet(wb)
    build_product_matrix_sheet(wb)

    # Save
    wb.save('Brewery_Financial_Model.xlsx')

    print("\n✅ Control & Assumptions sheets complete!")
    print("\nNext step: Run 02_volume_pricing.py")


if __name__ == '__main__':
    main()
