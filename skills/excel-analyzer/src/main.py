"""Excel Analyzer Skill - Claude Code integration for xls-extract.

This skill:
1. Runs xls-extract for factual extraction (reports, screenshots)
2. Guides Claude to provide AI-powered insights on the results
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from xls_extract import analyze_and_report, WorkbookAnalysis


def print_ai_analysis_prompt(result: WorkbookAnalysis, output_dir: Path) -> None:
    """Print guidance for Claude to provide AI analysis."""

    print("\n" + "=" * 70)
    print("AI ANALYSIS - Please analyze the extracted data")
    print("=" * 70)

    print("""
Now read the generated files and provide AI-powered insights:

## Files to Read

1. **Start with summary**: Read `{output_dir}/summary.md`
2. **Sheet details**: Read `{output_dir}/sheets/*.md` (focus on main sheets)
3. **View screenshots**: Look at `{output_dir}/screenshots/*.png`
""".format(output_dir=output_dir))

    # Add VBA prompt if macros exist
    if result.vba_modules:
        print(f"4. **VBA Code**: Read `{output_dir}/vba/*.md` - {len(result.vba_modules)} modules found")

    # Add Power Query prompt if present
    if result.power_queries:
        print(f"5. **Power Query**: Read `{output_dir}/power_query/*.md` - {len(result.power_queries)} queries found")

    print("""
## Analysis to Provide

### 1. Workbook Purpose
- What is the business purpose of this workbook?
- Who are the likely users?
- What decisions does it support?

### 2. Sheet-by-Sheet Analysis
For each **main sheet** (skip `pbi-*`, `ref-*`, `config-*` data sheets):
- **Purpose**: What business question does it answer?
- **Key Inputs**: Dropdowns, filters, date ranges, parameters
- **Data Layout**: What do rows/columns represent?
- **Key Outputs**: Important calculated values or summaries
- **Charts**: What visualizations exist and what do they show?

### 3. Data Flow & Architecture
- How do sheets connect to each other?
- What is the data flow? (inputs → calculations → outputs)
- Create a simple diagram if helpful:
  ```
  [Input Sheet] → [Calculations] → [Summary] → [Charts]
  ```

### 4. Complexity Analysis
- What are the most complex formulas?
- Any clever techniques worth noting?
- What would be hard to maintain?
""")

    # Add specific prompts based on what was found
    if result.error_cells:
        print(f"""### 5. Issues Found
- **{len(result.error_cells)} error cells** detected - explain what's broken and suggest fixes
""")

    if result.external_refs:
        print(f"""### 6. External Dependencies
- **{len(result.external_refs)} external references** found - assess risks of broken links
""")

    if result.vba_modules:
        print("""### 7. VBA Macro Analysis
- What do the macros do?
- Any security concerns?
- Are they essential or could they be removed?
""")

    print("""### 8. Recommendations
- What could be improved?
- What should users be careful about?
- Any modernization suggestions? (e.g., VLOOKUP → XLOOKUP)
""")

    print("=" * 70)
    print("Please read the files above and provide your analysis.")
    print("=" * 70 + "\n")


def main() -> int:
    """Main entry point for the skill."""
    parser = argparse.ArgumentParser(
        description="Analyze Excel workbooks with AI-powered insights"
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
    output_dir = Path(args.output) if args.output else file_path.parent / f"{file_path.stem}_analysis"

    try:
        # Step 1: Run xls-extract for factual analysis
        print("=" * 70)
        print("STEP 1: Extracting workbook data (xls-extract)")
        print("=" * 70 + "\n")

        result = analyze_and_report(
            file_path=file_path,
            output_dir=output_dir,
            capture_screenshots=not args.no_screenshots,
        )

        # Print extraction summary
        print("\n" + "-" * 40)
        print("Extraction Summary:")
        print("-" * 40)
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
        if result.connections:
            print(f"  Data Connections: {len(result.connections)}")
        if result.has_dax:
            print(f"  DAX/Power Pivot: Detected")
        if result.error_cells:
            print(f"  Error Cells: {len(result.error_cells)}")
        if result.external_refs:
            print(f"  External Refs: {len(result.external_refs)}")
        if result.screenshots:
            print(f"  Screenshots: {len(result.screenshots)}")

        print(f"\nOutput directory: {output_dir}")
        print(f"  - HTML Report: {output_dir}/index.html")
        print(f"  - Markdown: {output_dir}/README.md")

        # Step 2: Prompt Claude for AI analysis
        print_ai_analysis_prompt(result, output_dir)

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
