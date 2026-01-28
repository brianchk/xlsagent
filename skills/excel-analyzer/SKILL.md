# Excel Workbook Analyzer

Comprehensively analyze Excel workbooks from SharePoint URLs, producing rich HTML reports for human review and agent-optimized markdown files for AI consumption.

## Usage

```
/analyze-excel <sharepoint-url>
```

Or with a local file:
```
/analyze-excel /path/to/workbook.xlsx
```

## What It Does

1. **Downloads** the Excel file from SharePoint (handles M365 SSO authentication)
2. **Extracts** comprehensive metadata and content:
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
3. **Captures screenshots** of each sheet via Excel Online
4. **Generates** a comprehensive HTML report and agent-optimized markdown files

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
- uv package manager
- Playwright browsers installed

## First Run

The skill will automatically set up its virtual environment and install dependencies on first run.

For SharePoint URLs, you'll be prompted to authenticate via browser on first use. Sessions are cached for subsequent runs.

## Limitations

- **DAX/Power Pivot**: Can detect presence but cannot fully extract DAX formulas (proprietary format). Screenshots of Data Model view are captured when possible.
- **Very Hidden Sheets**: Documented but cannot be unhidden via Excel Online (requires VBA).
- **ActiveX Controls**: Limited extraction due to complex OLE embedding.

## Examples

```
/analyze-excel https://company.sharepoint.com/:x:/r/sites/Finance/Reports.xlsx
/analyze-excel ~/Downloads/quarterly-report.xlsx
```
