"""Excel Analyzer Skill - Claude Code integration for xls-extract.

This skill wraps xls-extract and adds AI-powered enrichment to the analysis.
The factual extraction, reports, and screenshots are handled by xls-extract.
This skill adds LLM-powered insights, explanations, and annotations.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from xls_extract import analyze_and_report, AnalysisOptions


def main() -> int:
    """Main entry point for the skill."""
    parser = argparse.ArgumentParser(
        description="Analyze Excel workbooks and generate documentation"
    )
    parser.add_argument(
        "input",
        help="Path to Excel file (.xlsx, .xlsm)"
    )
    parser.add_argument(
        "-o", "--output",
        help="Output directory for reports (default: <filename>_analysis/)"
    )
    parser.add_argument(
        "--no-screenshots",
        action="store_true",
        help="Skip screenshot capture"
    )

    args = parser.parse_args()

    # Validate input file
    file_path = Path(args.input)
    if not file_path.exists():
        print(f"Error: File not found: {file_path}")
        return 1

    if file_path.suffix.lower() not in ('.xlsx', '.xlsm', '.xlsb'):
        print(f"Error: Not an Excel file: {file_path}")
        return 1

    # Determine output directory
    output_dir = Path(args.output) if args.output else Path.cwd() / f"{file_path.stem}_analysis"

    try:
        # Run xls-extract for all factual analysis
        result = analyze_and_report(
            file_path=file_path,
            output_dir=output_dir,
            capture_screenshots=not args.no_screenshots,
        )

        # Print summary
        print("\n" + "=" * 60)
        print(f"ANALYSIS COMPLETE: {result.file_name}")
        print("=" * 60)
        print(f"\nSheets: {len(result.sheets)}")
        print(f"Formulas: {len(result.formulas)}")
        print(f"Named Ranges: {len(result.named_ranges)}")

        if result.vba_modules:
            print(f"VBA Modules: {len(result.vba_modules)}")
        if result.power_queries:
            print(f"Power Queries: {len(result.power_queries)}")
        if result.has_dax:
            print(f"DAX/Power Pivot: Detected")
        if result.screenshots:
            print(f"Screenshots: {len(result.screenshots)}")

        print(f"\nReports saved to: {output_dir}")
        print(f"  - HTML: {output_dir}/index.html")
        print(f"  - Markdown: {output_dir}/README.md")

        return 0

    except KeyboardInterrupt:
        print("\nCancelled.")
        return 130
    except Exception as e:
        print(f"Error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
