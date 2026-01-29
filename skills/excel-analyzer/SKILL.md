# Excel Workbook Analyzer

Comprehensively analyze local Excel workbooks, producing rich HTML reports for human review and agent-optimized markdown files for AI consumption.

## Usage

```
/analyze-excel /path/to/workbook.xlsx
```

With custom output directory:
```
/analyze-excel /path/to/workbook.xlsx -o ./my-output
```

Skip screenshots:
```
/analyze-excel /path/to/workbook.xlsx --no-screenshots
```

## What It Does

1. **Extracts** comprehensive metadata and content:
   - All sheets (visible, hidden, very hidden)
   - Formulas with classification (dynamic array, LAMBDA, lookup, etc.)
   - Named ranges and LAMBDA function definitions
   - VBA macros (for .xlsm files)
   - Power Query M code
   - Conditional formatting rules
   - Data validations (dropdowns, constraints)
   - Pivot tables and charts
   - Structured tables (ListObjects)
   - AutoFilters
   - Form controls and shapes
   - Data connections and external references
   - Comments (classic and threaded)
   - Hyperlinks
   - Protection settings
   - Print settings
   - Error cells (#REF!, #NAME?, etc.)
2. **Captures screenshots** of each sheet via Desktop Excel (xlwings)
3. **Generates** a comprehensive HTML report and agent-optimized markdown files

## Output

Creates an output directory with:
- `report.html` - Interactive HTML report with navigation and search
- `README.md` - Agent entry point
- `summary.md` - Quick facts for AI context
- `sheets/` - Per-sheet deep dives
- `formulas/` - Formula analysis
- `vba/` - VBA code extraction
- `power_query/` - M code extraction
- `features/` - Feature-specific documentation
- `issues/` - Error cells and broken references
- `screenshots/` - Visual captures of each sheet
- `artifacts/` - Raw extracted files (.bas, .m, etc.)

## Requirements

- Python 3.11+
- **Screenshots (Windows only):**
  - Microsoft Excel installed
  - pywin32 and pillow packages (installed automatically)
- **macOS:** Screenshots not supported due to Excel automation limitations. Analysis still works fully.

## First Run

The skill will automatically set up its virtual environment and install dependencies on first run.

## Limitations

- **Screenshots (macOS)**: Not supported due to Excel automation limitations (security dialogs cannot be suppressed). Use `--no-screenshots` or run on Windows.
- **DAX/Power Pivot**: Can detect presence but cannot fully extract DAX formulas (proprietary format).
- **Very Hidden Sheets**: Documented but screenshots not available (requires VBA to unhide).
- **ActiveX Controls**: Limited extraction due to complex OLE embedding.

## CLI Usage

Can also be run directly from command line:

```bash
cd skills/excel-analyzer
source .venv/bin/activate  # or .venv\Scripts\activate on Windows
python -m src.main /path/to/workbook.xlsx -o ./output
```

## Examples

As a Claude skill:
```
/analyze-excel ~/Downloads/quarterly-report.xlsx
/analyze-excel /Users/brian/Documents/budget.xlsm -o ./budget-analysis
/analyze-excel ./data.xlsx --no-screenshots
```

As CLI:
```bash
python -m src.main quarterly-report.xlsx
python -m src.main budget.xlsm -o ./budget-analysis
python -m src.main data.xlsx --no-screenshots
```
