#!/usr/bin/env python3
"""
Beer Brewery Financial Model - Master Build Script
Runs all build scripts in sequence

Usage: python run_all.py
"""

import sys
import subprocess
from pathlib import Path

def run_script(script_name):
    """Run a Python script and handle errors"""
    print(f"\n{'='*60}")
    print(f"Running: {script_name}")
    print('='*60)

    try:
        result = subprocess.run([sys.executable, script_name], check=True, capture_output=False)
        print(f"✅ {script_name} completed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {script_name} failed with error")
        print(f"   Error: {e}")
        return False
    except FileNotFoundError:
        print(f"⚠️  {script_name} not found - skipping")
        return None  # None means script doesn't exist yet


def main():
    print("\n" + "="*60)
    print("BEER BREWERY FINANCIAL MODEL - FULL BUILD")
    print("="*60)
    print("\nThis will create a complete 10-year, 3-statement financial model")
    print("Estimated time: 1-2 minutes\n")

    # Check for openpyxl
    try:
        import openpyxl
        print("✓ openpyxl library found")
    except ImportError:
        print("❌ openpyxl not found!")
        print("\nPlease install: pip install openpyxl")
        sys.exit(1)

    # List of scripts to run in order
    scripts = [
        '00_setup_workbook.py',
        '01_control_assumptions.py',
        '02_volume_pricing.py',
        '03_operations_costs.py',
        '04_capital_funding.py',
        '05_statements.py',
        '06_kpis_checks.py',
        '07_cover_docs.py',
        '08_formatting.py',
    ]

    # Check which scripts exist
    existing_scripts = [s for s in scripts if Path(s).exists()]
    missing_scripts = [s for s in scripts if not Path(s).exists()]

    print(f"\nFound {len(existing_scripts)} of {len(scripts)} build scripts")

    if missing_scripts:
        print(f"\n⚠️  Missing scripts (will be skipped):")
        for script in missing_scripts:
            print(f"   - {script}")
        print("\n   These scripts are still being developed.")
        print("   The model will be partially built with available scripts.\n")

    input("Press ENTER to continue...")

    # Run existing scripts
    success_count = 0
    skip_count = 0

    for script in scripts:
        result = run_script(script)
        if result is True:
            success_count += 1
        elif result is None:
            skip_count += 1
        else:
            print(f"\n❌ Build stopped due to error in {script}")
            print("   Fix the error and run again")
            sys.exit(1)

    # Summary
    print("\n" + "="*60)
    print("BUILD COMPLETE!")
    print("="*60)
    print(f"\n✅ Successfully ran: {success_count} scripts")
    if skip_count > 0:
        print(f"⚠️  Skipped: {skip_count} scripts (not yet created)")

    print(f"\n📄 Output file: Brewery_Financial_Model.xlsx")

    print("\n" + "="*60)
    print("NEXT STEPS")
    print("="*60)
    print("\n1. Open: Brewery_Financial_Model.xlsx")
    print("2. Enable iterative calculation:")
    print("   File → Options → Formulas → Enable iterative calculation")
    print("   (Set: 100 iterations, 0.001 max change)")
    print("\n3. Select scenario: Go to Control sheet, choose Base/Upside/Downside")
    print("\n4. Review outputs: P&L, CashFlow, BalanceSheet, KPI_Dashboard")
    print("\n5. Check for errors: Checks_Errors sheet")

    if skip_count > 0:
        print("\n" + "="*60)
        print("NOTE: Model is PARTIALLY BUILT")
        print("="*60)
        print(f"\n{skip_count} scripts were not yet created.")
        print("You have a working foundation with:")
        print("  ✓ Control & scenarios")
        print("  ✓ Volume assumptions (120 months)")
        print("  ✓ Pricing structure")
        print("\nRemaining scripts (operations, statements, KPIs) will be")
        print("created in the next iteration.")

    print("\n")


if __name__ == '__main__':
    main()
