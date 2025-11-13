#!/usr/bin/env python3
"""
Beer Brewery Financial Model Builder
Builds a comprehensive 10-year, monthly financial model for a beer brewery
with integrated Income Statement, Cash Flow, and Balance Sheet
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import calendar

# Color scheme
COLOR_INPUT = "FFFFE0"  # Light yellow
COLOR_CALC = "FFFFFF"   # White
COLOR_OUTPUT = "E0F0FF" # Light blue
COLOR_HEADER = "4472C4"  # Dark blue
FONT_INPUT = Font(color="0000FF", bold=False)  # Blue
FONT_CALC = Font(color="000000", bold=False)   # Black
FONT_OUTPUT = Font(color="000000", bold=True)  # Black bold
FONT_HEADER = Font(color="FFFFFF", bold=True)  # White bold

def create_border(style='thin'):
    """Create cell border"""
    side = Side(style=style, color="000000")
    return Border(left=side, right=side, top=side, bottom=side)

def set_column_width(ws, col, width):
    """Set column width"""
    ws.column_dimensions[get_column_letter(col)].width = width

def format_header(cell):
    """Format header cell"""
    cell.font = FONT_HEADER
    cell.fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type="solid")
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = create_border()

def format_input(cell):
    """Format input cell"""
    cell.font = FONT_INPUT
    cell.fill = PatternFill(start_color=COLOR_INPUT, end_color=COLOR_INPUT, fill_type="solid")
    cell.border = create_border()

def format_calc(cell):
    """Format calculation cell"""
    cell.font = FONT_CALC
    cell.border = create_border('thin')

def format_output(cell):
    """Format output cell"""
    cell.font = FONT_OUTPUT
    cell.fill = PatternFill(start_color=COLOR_OUTPUT, end_color=COLOR_OUTPUT, fill_type="solid")
    cell.border = create_border()

print("Starting Brewery Financial Model build...")
print("=" * 80)

# Create workbook
wb = Workbook()
wb.remove(wb.active)  # Remove default sheet

# Model parameters
START_DATE = datetime(2025, 1, 1)
NUM_PERIODS = 120  # 10 years monthly
END_DATE = START_DATE + relativedelta(months=NUM_PERIODS-1)

print(f"Model period: {START_DATE.strftime('%b %Y')} to {END_DATE.strftime('%b %Y')}")
print(f"Number of periods: {NUM_PERIODS} months")
print()

# ============================================================================
# SHEET 1: CONTROL
# ============================================================================
print("Creating sheet: CONTROL")
ws_control = wb.create_sheet("CONTROL", 0)

ws_control['A1'] = "BEER BREWERY FINANCIAL MODEL"
ws_control['A1'].font = Font(size=16, bold=True, color="4472C4")
ws_control.merge_cells('A1:D1')

ws_control['A2'] = "CONTROL PANEL"
ws_control['A2'].font = Font(size=14, bold=True)
ws_control.merge_cells('A2:D2')

ws_control['A4'] = "Scenario:"
ws_control['B4'] = "Base"
format_input(ws_control['B4'])
ws_control['A4'].font = Font(bold=True)

ws_control['A5'] = "Scenario ID:"
ws_control['B5'] = '=MATCH(B4,{"Base","Upside","Downside"},0)'
format_calc(ws_control['B5'])
ws_control['A5'].font = Font(bold=True)

ws_control['A7'] = "Reporting Date:"
ws_control['B7'] = datetime(2025, 6, 1)
ws_control['B7'].number_format = 'dd-mmm-yyyy'
format_input(ws_control['B7'])
ws_control['A7'].font = Font(bold=True)

ws_control['A9'] = "TOGGLES:"
ws_control['A9'].font = Font(bold=True, underline='single')

ws_control['A10'] = "Enable Contract Brewing?"
ws_control['B10'] = "Yes"
format_input(ws_control['B10'])

ws_control['A11'] = "Enable Excise Shock (Year 5)?"
ws_control['B11'] = "Yes"
format_input(ws_control['B11'])

ws_control['A12'] = "Enable Capacity Expansion?"
ws_control['B12'] = "Yes"
format_input(ws_control['B12'])

ws_control['A14'] = "MODEL VERSION:"
ws_control['A14'].font = Font(bold=True, underline='single')
ws_control['A15'] = "Version:"
ws_control['B15'] = "1.0"
ws_control['A16'] = "Last Updated:"
ws_control['B16'] = datetime.now().strftime('%d-%b-%Y')
ws_control['A17'] = "Updated By:"
ws_control['B17'] = "Model Builder"

set_column_width(ws_control, 1, 30)
set_column_width(ws_control, 2, 20)

# Define names
wb.define_name('ScenarioID', '=CONTROL!$B$5')
wb.define_name('ReportingDate', '=CONTROL!$B$7')
wb.define_name('Toggle_ContractBrewing', '=CONTROL!$B$10')
wb.define_name('Toggle_ExciseShock', '=CONTROL!$B$11')
wb.define_name('Toggle_CapacityExpansion', '=CONTROL!$B$12')

# ============================================================================
# SHEET 2: TIMELINE
# ============================================================================
print("Creating sheet: TIMELINE")
ws_timeline = wb.create_sheet("TIMELINE", 1)

ws_timeline['A1'] = "MODEL TIMELINE"
ws_timeline['A1'].font = Font(size=14, bold=True)
ws_timeline.merge_cells('A1:E1')

ws_timeline['A3'] = "Model Start Date:"
ws_timeline['B3'] = START_DATE
ws_timeline['B3'].number_format = 'dd-mmm-yyyy'
format_input(ws_timeline['B3'])

ws_timeline['A4'] = "Number of Periods:"
ws_timeline['B4'] = NUM_PERIODS
format_input(ws_timeline['B4'])

ws_timeline['A5'] = "Frequency:"
ws_timeline['B5'] = "Monthly"

# Create date array
ws_timeline['A7'] = "Date"
ws_timeline['A8'] = "Period #"
ws_timeline['A9'] = "Month"
ws_timeline['A10'] = "Quarter"
ws_timeline['A11'] = "FY"
ws_timeline['A12'] = "FY-End Flag"
ws_timeline['A13'] = "Year #"

for col in ['A7', 'A8', 'A9', 'A10', 'A11', 'A12', 'A13']:
    format_header(ws_timeline[col])

# Generate timeline
current_date = START_DATE
for period in range(NUM_PERIODS):
    col_num = period + 2  # Start from column B
    col_letter = get_column_letter(col_num)

    # Date
    ws_timeline[f'{col_letter}7'] = current_date
    ws_timeline[f'{col_letter}7'].number_format = 'mmm-yy'

    # Period #
    ws_timeline[f'{col_letter}8'] = period + 1

    # Month
    ws_timeline[f'{col_letter}9'] = current_date.strftime('%b')

    # Quarter
    quarter = (current_date.month - 1) // 3 + 1
    ws_timeline[f'{col_letter}10'] = f'Q{quarter}'

    # FY (assuming Jul-Jun fiscal year for AU/NZ)
    if current_date.month >= 7:
        fy_year = current_date.year + 1
    else:
        fy_year = current_date.year
    ws_timeline[f'{col_letter}11'] = f'FY{fy_year}'

    # FY-End Flag (1 if June, 0 otherwise)
    ws_timeline[f'{col_letter}12'] = 1 if current_date.month == 6 else 0

    # Year # (1-10)
    year_num = (period // 12) + 1
    ws_timeline[f'{col_letter}13'] = year_num

    # Format
    for row in range(7, 14):
        cell = ws_timeline[f'{col_letter}{row}']
        format_calc(cell)
        cell.alignment = Alignment(horizontal='center')

    current_date = current_date + relativedelta(months=1)

set_column_width(ws_timeline, 1, 20)
for col in range(2, NUM_PERIODS + 2):
    set_column_width(ws_timeline, col, 10)

# Freeze panes
ws_timeline.freeze_panes = 'B8'

# Define named range for dates
wb.define_name('Timeline_Dates', f'=TIMELINE!$B$7:${get_column_letter(NUM_PERIODS+1)}$7')
wb.define_name('Timeline_Periods', f'=TIMELINE!$B$8:${get_column_letter(NUM_PERIODS+1)}$8')

# ============================================================================
# SHEET 3: SKU_Master
# ============================================================================
print("Creating sheet: SKU_Master")
ws_sku = wb.create_sheet("SKU_Master", 2)

ws_sku['A1'] = "SKU MASTER DATA"
ws_sku['A1'].font = Font(size=14, bold=True)
ws_sku.merge_cells('A1:H1')

# Headers
headers = ['SKU ID', 'Brand', 'Pack Type', 'Size (mL)', 'ABV %', 'On-Trade %', 'Off-Trade %', 'Pack Category']
for col, header in enumerate(headers, 1):
    cell = ws_sku.cell(row=3, column=col, value=header)
    format_header(cell)

# SKU data
skus = [
    ['SKU-01', 'SkyBrew', 'Bottle', 330, 0.045, 0.30, 0.70, 'Packaged'],
    ['SKU-02', 'SkyBrew', 'Can 6-pack', 375, 0.045, 0.10, 0.90, 'Packaged'],
    ['SKU-03', 'SkyBrew', 'Keg', 50000, 0.045, 0.95, 0.05, 'Draught'],
    ['SKU-04', 'AllDark', 'Bottle', 330, 0.058, 0.40, 0.60, 'Packaged'],
    ['SKU-05', 'AllDark', 'Can 4-pack', 375, 0.058, 0.20, 0.80, 'Packaged'],
    ['SKU-06', 'AllDark', 'Keg', 50000, 0.058, 0.95, 0.05, 'Draught'],
    ['SKU-07', 'Craft IPA', 'Bottle', 330, 0.062, 0.50, 0.50, 'Packaged'],
    ['SKU-08', 'Craft IPA', 'Can 4-pack', 375, 0.062, 0.30, 0.70, 'Packaged'],
    ['SKU-09', 'Session', 'Can 6-pack', 375, 0.035, 0.10, 0.90, 'Packaged'],
    ['SKU-10', 'Premium Pilsner', 'Bottle', 330, 0.050, 0.40, 0.60, 'Packaged'],
]

for row, sku_data in enumerate(skus, 4):
    for col, value in enumerate(sku_data, 1):
        cell = ws_sku.cell(row=row, column=col, value=value)
        if col in [5, 6, 7]:  # Percentage columns
            cell.number_format = '0.0%'
        format_input(cell)

# Set column widths
widths = [12, 18, 18, 12, 10, 12, 12, 15]
for col, width in enumerate(widths, 1):
    set_column_width(ws_sku, col, width)

# Convert to table
# Note: openpyxl doesn't have full table support, so we'll add as reference
print("  - Added 10 SKUs")

# ============================================================================
# SHEET 4: Assumptions_Macro
# ============================================================================
print("Creating sheet: Assumptions_Macro")
ws_macro = wb.create_sheet("Assumptions_Macro", 3)

ws_macro['A1'] = "MACRO & TAX ASSUMPTIONS"
ws_macro['A1'].font = Font(size=14, bold=True)
ws_macro.merge_cells('A1:E1')

ws_macro['A3'] = "SCENARIO ASSUMPTIONS"
ws_macro['A3'].font = Font(bold=True, underline='single')

# Scenario header
ws_macro['A4'] = "Driver"
ws_macro['B4'] = "Base"
ws_macro['C4'] = "Upside"
ws_macro['D4'] = "Downside"
for col in ['A4', 'B4', 'C4', 'D4']:
    format_header(ws_macro[col])

# Inflation rates
assumptions = [
    ['Materials Inflation % p.a.', 0.030, 0.020, 0.050],
    ['Packaging Inflation % p.a.', 0.025, 0.020, 0.045],
    ['Labour Inflation % p.a.', 0.035, 0.030, 0.050],
    ['CPI (Excise Indexation) % p.a.', 0.025, 0.020, 0.035],
    ['FX Rate AUD/NZD', 1.08, 1.05, 1.12],
]

row = 5
for assumption in assumptions:
    ws_macro[f'A{row}'] = assumption[0]
    ws_macro[f'A{row}'].font = Font(bold=True)
    ws_macro[f'B{row}'] = assumption[1]
    ws_macro[f'C{row}'] = assumption[2]
    ws_macro[f'D{row}'] = assumption[3]

    # Format percentage cells
    if '%' in assumption[0]:
        for col in ['B', 'C', 'D']:
            ws_macro[f'{col}{row}'].number_format = '0.0%'

    format_input(ws_macro[f'B{row}'])
    format_input(ws_macro[f'C{row}'])
    format_input(ws_macro[f'D{row}'])
    row += 1

# Excise shock
ws_macro[f'A{row}'] = "Excise Shock Year 5 (%)"
ws_macro[f'A{row}'].font = Font(bold=True)
ws_macro[f'B{row}'] = 0.00
ws_macro[f'C{row}'] = 0.00
ws_macro[f'D{row}'] = 0.10
for col in ['B', 'C', 'D']:
    ws_macro[f'{col}{row}'].number_format = '0.0%'
    format_input(ws_macro[f'{col}{row}'])

row += 2
ws_macro[f'A{row}'] = "TAX & EXCISE RATES"
ws_macro[f'A{row}'].font = Font(bold=True, underline='single')

row += 1
ws_macro[f'A{row}'] = "Corporate Tax Rate"
ws_macro[f'B{row}'] = 0.30
ws_macro[f'B{row}'].number_format = '0.0%'
format_input(ws_macro[f'B{row}'])

row += 1
ws_macro[f'A{row}'] = "Excise Rate - AU Draught (AUD/L alcohol)"
ws_macro[f'B{row}'] = 36.32
ws_macro[f'B{row}'].number_format = '#,##0.00'
format_input(ws_macro[f'B{row}'])

row += 1
ws_macro[f'A{row}'] = "Excise Rate - AU Packaged >3.5% (AUD/L)"
ws_macro[f'B{row}'] = 54.72
ws_macro[f'B{row}'].number_format = '#,##0.00'
format_input(ws_macro[f'B{row}'])

row += 1
ws_macro[f'A{row}'] = "Excise Rate - AU Packaged ≤3.5% (AUD/L)"
ws_macro[f'B{row}'] = 36.32
ws_macro[f'B{row}'].number_format = '#,##0.00'
format_input(ws_macro[f'B{row}'])

row += 1
ws_macro[f'A{row}'] = "Excise Rate - NZ All Types (NZD/L alcohol)"
ws_macro[f'B{row}'] = 30.23
ws_macro[f'B{row}'].number_format = '#,##0.00'
format_input(ws_macro[f'B{row}'])

set_column_width(ws_macro, 1, 35)
set_column_width(ws_macro, 2, 15)
set_column_width(ws_macro, 3, 15)
set_column_width(ws_macro, 4, 15)

print("  - Added macro assumptions and excise rates")

# ============================================================================
# SHEET 5: Assumptions_Commercial
# ============================================================================
print("Creating sheet: Assumptions_Commercial")
ws_comm = wb.create_sheet("Assumptions_Commercial", 4)

ws_comm['A1'] = "COMMERCIAL ASSUMPTIONS"
ws_comm['A1'].font = Font(size=14, bold=True)
ws_comm.merge_cells('A1:E1')

ws_comm['A3'] = "VOLUME GROWTH RATES (% p.a.)"
ws_comm['A3'].font = Font(bold=True, underline='single')

ws_comm['A4'] = "Market"
ws_comm['B4'] = "Base"
ws_comm['C4'] = "Upside"
ws_comm['D4'] = "Downside"
for col in ['A4', 'B4', 'C4', 'D4']:
    format_header(ws_comm[col])

growth_rates = [
    ['Australia', 0.040, 0.065, 0.010],
    ['New Zealand', 0.030, 0.050, 0.000],
    ['Exports', 0.150, 0.200, 0.080],
]

row = 5
for rate in growth_rates:
    ws_comm[f'A{row}'] = rate[0]
    ws_comm[f'A{row}'].font = Font(bold=True)
    for col_idx, val in enumerate(rate[1:], 2):
        cell = ws_comm.cell(row=row, column=col_idx, value=val)
        cell.number_format = '0.0%'
        format_input(cell)
    row += 1

row += 1
ws_comm[f'A{row}'] = "PRICE ESCALATION RATES (% p.a.)"
ws_comm[f'A{row}'].font = Font(bold=True, underline='single')

row += 1
ws_comm[f'A{row}'] = "Channel"
ws_comm[f'B{row}'] = "Base"
ws_comm[f'C{row}'] = "Upside"
ws_comm[f'D{row}'] = "Downside"
for col in [f'A{row}', f'B{row}', f'C{row}', f'D{row}']:
    format_header(ws_comm[col])

row += 1
price_rates = [
    ['On-Trade', 0.035, 0.050, 0.015],
    ['Off-Trade', 0.020, 0.035, 0.005],
]

for price in price_rates:
    ws_comm[f'A{row}'] = price[0]
    ws_comm[f'A{row}'].font = Font(bold=True)
    for col_idx, val in enumerate(price[1:], 2):
        cell = ws_comm.cell(row=row, column=col_idx, value=val)
        cell.number_format = '0.0%'
        format_input(cell)
    row += 1

row += 1
ws_comm[f'A{row}'] = "A&P SPEND (% of Net Revenue)"
ws_comm[f'A{row}'].font = Font(bold=True, underline='single')

row += 1
ws_comm[f'A{row}'] = "A&P %"
ws_comm[f'B{row}'] = 0.12
ws_comm[f'C{row}'] = 0.14
ws_comm[f'D{row}'] = 0.09
for col in ['B', 'C', 'D']:
    ws_comm[f'{col}{row}'].number_format = '0.0%'
    format_input(ws_comm[f'{col}{row}'])

row += 2
ws_comm[f'A{row}'] = "SEASONALITY INDICES (Monthly)"
ws_comm[f'A{row}'].font = Font(bold=True, underline='single')

row += 1
ws_comm[f'A{row}'] = "Month"
ws_comm[f'B{row}'] = "Index"
format_header(ws_comm[f'A{row}'])
format_header(ws_comm[f'B{row}'])

seasonality = [
    ['Jan', 1.15], ['Feb', 1.10], ['Mar', 1.05], ['Apr', 0.95],
    ['May', 0.85], ['Jun', 0.80], ['Jul', 0.80], ['Aug', 0.85],
    ['Sep', 0.95], ['Oct', 1.05], ['Nov', 1.10], ['Dec', 1.35]
]

row += 1
for month, index in seasonality:
    ws_comm[f'A{row}'] = month
    ws_comm[f'B{row}'] = index
    ws_comm[f'B{row}'].number_format = '0.00'
    format_input(ws_comm[f'B{row}'])
    row += 1

# Add average check
ws_comm[f'A{row}'] = "Average (must = 1.00)"
ws_comm[f'B{row}'] = f'=AVERAGE(B{row-12}:B{row-1})'
ws_comm[f'B{row}'].number_format = '0.00'
ws_comm[f'B{row}'].font = Font(bold=True)
format_calc(ws_comm[f'B{row}'])

set_column_width(ws_comm, 1, 35)
for col in range(2, 5):
    set_column_width(ws_comm, col, 15)

print("  - Added commercial assumptions including seasonality")

# Continue building remaining sheets...
# Due to length constraints, I'll create a modular approach

print("\nPhase 1 complete: Core structure and assumptions sheets created")
print("=" * 80)

# Save the workbook
output_file = "/home/user/123/Brewery_Financial_Model_v1.0.xlsx"
wb.save(output_file)
print(f"\nWorkbook saved to: {output_file}")
print(f"Total sheets created: {len(wb.sheetnames)}")
print("Sheets:", ", ".join(wb.sheetnames))
