"""CLI entry point for xls-extract.

Usage:
    python -m xls_extract workbook.xlsx -o ./output
    xls-extract workbook.xlsx -o ./output
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def main() -> int:
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        prog="xls-extract",
        description="Extract data from Excel workbooks and generate reports",
    )
    parser.add_argument(
        "input",
        help="Path to Excel file (.xlsx, .xlsm, .xlsb)",
    )
    parser.add_argument(
        "-o", "--output",
        help="Output directory for reports (default: <filename>_analysis/)",
    )
    parser.add_argument(
        "--no-screenshots",
        action="store_true",
        help="Skip screenshot capture (screenshots are Windows-only anyway)",
    )
    parser.add_argument(
        "--data-only",
        action="store_true",
        help="Extract data only, skip report generation",
    )

    args = parser.parse_args()

    # Validate input file
    file_path = Path(args.input)
    if not file_path.exists():
        print(f"Error: File not found: {file_path}")
        return 1

    if file_path.suffix.lower() not in (".xlsx", ".xlsm", ".xlsb"):
        print(f"Error: Not an Excel file: {file_path}")
        return 1

    # Determine output directory
    output_dir = Path(args.output) if args.output else file_path.parent / f"{file_path.stem}_analysis"

    try:
        if args.data_only:
            # Data extraction only
            from . import analyze

            print(f"Analyzing: {file_path.name}")
            result = analyze(file_path)

            print(f"\nExtraction complete:")
            print(f"  Sheets: {len(result.sheets)}")
            print(f"  Formulas: {len(result.formulas)}")
            print(f"  Named Ranges: {len(result.named_ranges)}")
            print(f"  Charts: {len(result.charts)}")
            print(f"  Tables: {len(result.tables)}")
            print(f"  Pivot Tables: {len(result.pivot_tables)}")

            if result.vba_modules:
                print(f"  VBA Modules: {len(result.vba_modules)}")
            if result.power_queries:
                print(f"  Power Queries: {len(result.power_queries)}")

        else:
            # Full analysis with reports
            from . import analyze_and_report

            result = analyze_and_report(
                file_path=file_path,
                output_dir=output_dir,
                capture_screenshots=not args.no_screenshots,
            )

            print(f"\nExtraction complete:")
            print(f"  Sheets: {len(result.sheets)}")
            print(f"  Formulas: {len(result.formulas)}")
            print(f"  Named Ranges: {len(result.named_ranges)}")
            print(f"  Charts: {len(result.charts)}")
            print(f"  Tables: {len(result.tables)}")
            print(f"  Pivot Tables: {len(result.pivot_tables)}")

            if result.vba_modules:
                print(f"  VBA Modules: {len(result.vba_modules)}")
            if result.power_queries:
                print(f"  Power Queries: {len(result.power_queries)}")
            if result.screenshots:
                print(f"  Screenshots: {len(result.screenshots)}")

            print(f"\nOutput: {output_dir}")
            print(f"  HTML Report: {output_dir}/index.html")
            print(f"  Markdown: {output_dir}/README.md")

        return 0

    except KeyboardInterrupt:
        print("\nCancelled.")
        return 130
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
