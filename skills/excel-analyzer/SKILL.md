# Excel Workbook Analyzer

Comprehensively analyze Excel workbooks by combining factual extraction (via `xls-extract`) with AI-powered insights.

## Usage

```
/analyze-excel /path/to/workbook.xlsx
```

With custom output directory:
```
/analyze-excel /path/to/workbook.xlsx -o ./my-output
```

Skip screenshots (required on macOS):
```
/analyze-excel /path/to/workbook.xlsx --no-screenshots
```

## How This Skill Works

This skill has two distinct phases:

### Phase 1: Factual Extraction (Python/xls-extract)

The `xls-extract` library handles all deterministic extraction:
- Sheets, formulas, named ranges
- VBA macros, Power Query M code
- Charts, pivot tables, tables
- Conditional formatting, data validation
- Screenshots (Windows only)
- HTML report and Markdown documentation

Run the extraction:
```bash
cd ~/.claude/skills/excel-analyzer
source .venv/bin/activate  # .venv\Scripts\activate on Windows
python -m src.main "<file_path>" -o "<output_dir>"
```

### Phase 2: AI Analysis (Claude)

After extraction, Claude reads the generated files and provides insights that require understanding and interpretation:

1. **Read the generated files**:
   - `<output_dir>/summary.md` - Start here
   - `<output_dir>/sheets/*.md` - Sheet details
   - `<output_dir>/screenshots/*.png` - Visual layout
   - `<output_dir>/vba/*.md` - VBA code (if present)
   - `<output_dir>/formulas/*.md` - Formula analysis

2. **Provide AI-powered analysis**:

#### Workbook Purpose & Context
- What is the business purpose?
- Who are the likely users?
- What decisions does it support?

#### Sheet-by-Sheet Analysis
For main sheets (skip `pbi-*`, `ref-*`, `config-*` data sheets):
- **Purpose**: What business question does it answer?
- **Key Inputs**: Dropdowns, filters, parameters users can change
- **Data Layout**: What rows/columns represent
- **Key Outputs**: Important calculations and summaries
- **Visualizations**: What charts show and why

#### Architecture & Data Flow
- How sheets connect to each other
- Data flow diagram (inputs → calculations → outputs)
- Dependencies between components

#### Complexity Analysis
- Most complex formulas explained in plain English
- Clever techniques worth noting
- Areas that would be hard to maintain

#### Issues & Risks
- Error cells explained with fix suggestions
- External reference risks
- Security concerns (macros, external links)
- Data quality issues

#### VBA Analysis (if present)
- What each macro does
- Security assessment
- Whether macros are essential

#### Recommendations
- Improvements and modernization suggestions
- User warnings and gotchas
- Migration considerations

## Output Structure

The extraction creates:
```
<output_dir>/
├── index.html           # Interactive HTML report
├── README.md            # Entry point for Claude
├── summary.md           # Quick facts
├── sheets/              # Per-sheet details
│   ├── _index.md
│   └── <sheet>.md
├── formulas/            # Formula analysis
├── vba/                 # VBA code (if .xlsm)
├── power_query/         # M code (if present)
├── features/            # CF, validations, etc.
├── issues/              # Errors, external refs
├── screenshots/         # Visual captures
│   ├── <sheet>_full.png
│   └── <sheet>_detail.png
└── workbook/            # Workbook-wide items
    ├── vba.html
    └── connections.html
```

## Example Session

When user runs `/analyze-excel /path/to/budget.xlsx`:

**Step 1** - Run extraction:
```bash
cd ~/.claude/skills/excel-analyzer && source .venv/bin/activate
python -m src.main "/path/to/budget.xlsx" -o "/path/to/budget_analysis"
```

**Step 2** - Claude reads and analyzes:
```
Reading summary.md...
Reading sheets/Input.md, sheets/Calculations.md, sheets/Summary.md...
Viewing screenshots...
```

**Step 3** - Claude provides insights:
> "This workbook is a departmental budget tracker used by finance teams to monitor monthly spending against plan.
>
> **Key sheets:**
> - **Input**: Where users enter monthly actuals via data validation dropdowns
> - **Calculations**: SUMIFS formulas aggregate by department and category
> - **Summary**: Executive dashboard with YTD variance analysis
>
> **Data flow:**
> ```
> [Input] → [Calculations] → [Summary] → [Charts]
>            ↑
>       [ref-budget]
> ```
>
> **Complex formula on Calculations!E15:**
> This SUMIFS with multiple criteria calculates YTD spend by department...
>
> **Issues found:**
> - 3 #REF! errors in row 45 - likely caused by deleted rows
> - External link to `\\server\data.xlsx` may break if moved
>
> **Recommendations:**
> - Replace VLOOKUP formulas with XLOOKUP for better maintainability
> - The manual copy-paste step could be automated with Power Query"

## Requirements

- Python 3.11+
- `xls-extract` library (installed automatically)
- **Screenshots**: Windows only with Desktop Excel

## Setup

First-time setup:
```bash
cd ~/.claude/skills/excel-analyzer
python -m venv .venv
source .venv/bin/activate  # .venv\Scripts\activate on Windows
pip install -e .
```

For screenshots on Windows:
```bash
pip install "xls-extract[screenshots]"
```

## Limitations

- **Screenshots**: Windows only (macOS not supported)
- **DAX/Power Pivot**: Can detect but not fully extract
- **Very Hidden Sheets**: Documented but not screenshotted
- **ActiveX Controls**: Limited extraction
