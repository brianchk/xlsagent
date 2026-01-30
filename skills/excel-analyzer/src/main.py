"""Excel Analyzer Skill - Claude Code integration for xls-extract.

This skill:
1. Runs xls-extract for factual extraction (reports, screenshots)
2. Guides Claude to provide AI-powered insights on the results

Can also work with pre-existing extraction output (e.g., from Windows with screenshots)
by using the --existing flag.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xls_extract import WorkbookAnalysis


def print_ai_analysis_prompt_from_result(result: "WorkbookAnalysis", output_dir: Path) -> None:
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


def print_ai_analysis_prompt_from_dir(output_dir: Path) -> None:
    """Print guidance for Claude based on existing output directory."""

    print("\n" + "=" * 70)
    print("AI ANALYSIS - Please analyze the extracted data")
    print("=" * 70)

    print(f"""
Using existing extraction output from: {output_dir}

## Files to Read

1. **Start with summary**: Read `{output_dir}/summary.md`
2. **Sheet details**: Read `{output_dir}/sheets/*.md` (focus on main sheets)
3. **View screenshots**: Look at `{output_dir}/screenshots/*.png`
""")

    # Check what exists and add to prompts
    vba_dir = output_dir / "vba"
    if vba_dir.exists() and any(vba_dir.glob("*.md")):
        vba_count = len(list(vba_dir.glob("*.md"))) - 1  # exclude _index.md
        print(f"4. **VBA Code**: Read `{output_dir}/vba/*.md` - {vba_count} modules found")

    pq_dir = output_dir / "power_query"
    if pq_dir.exists() and any(pq_dir.glob("*.md")):
        pq_count = len(list(pq_dir.glob("*.md"))) - 1  # exclude _index.md
        print(f"5. **Power Query**: Read `{output_dir}/power_query/*.md` - {pq_count} queries found")

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

### 5. Issues & Risks
- Check `{output_dir}/issues/` for errors and external references
- Explain what's broken and suggest fixes
- Assess risks of external dependencies

### 6. VBA Analysis (if present)
- What do the macros do?
- Any security concerns?
- Are they essential or could they be removed?

### 7. Recommendations
- What could be improved?
- What should users be careful about?
- Any modernization suggestions? (e.g., VLOOKUP → XLOOKUP)
""".format(output_dir=output_dir))

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
        nargs="?",
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
    parser.add_argument(
        "--existing",
        metavar="DIR",
        help="Use existing extraction output directory (skip extraction, AI analysis only)"
    )

    args = parser.parse_args()

    # Mode 1: Use existing output directory (AI analysis only)
    if args.existing:
        output_dir = Path(args.existing)
        if not output_dir.exists():
            print(f"Error: Output directory not found: {output_dir}")
            return 1

        # Verify it looks like valid extraction output
        if not (output_dir / "summary.md").exists():
            print(f"Error: Not a valid extraction output (missing summary.md): {output_dir}")
            return 1

        print("=" * 70)
        print("USING EXISTING EXTRACTION OUTPUT")
        print("=" * 70)
        print(f"\nOutput directory: {output_dir}")

        # Check what's available
        sheets_dir = output_dir / "sheets"
        screenshots_dir = output_dir / "screenshots"
        if sheets_dir.exists():
            sheet_count = len(list(sheets_dir.glob("*.md"))) - 1  # exclude _index.md
            print(f"  Sheets: {sheet_count}")
        if screenshots_dir.exists():
            screenshot_count = len(list(screenshots_dir.glob("*.png")))
            print(f"  Screenshots: {screenshot_count}")

        # Print AI analysis prompt
        print_ai_analysis_prompt_from_dir(output_dir)
        return 0

    # Mode 2: Full extraction + AI analysis
    if not args.input:
        parser.error("Either provide an input file or use --existing DIR")

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
        # Import here so --existing mode works without xls-extract installed
        from xls_extract import analyze_and_report

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
        print_ai_analysis_prompt_from_result(result, output_dir)

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
