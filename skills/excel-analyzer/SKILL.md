# Excel Workbook Analyzer

Comprehensively analyze Excel workbooks by combining technical extraction with AI-powered insights.

## Usage

```
/analyze-excel /path/to/workbook.xlsx
```

With custom output directory:
```
/analyze-excel /path/to/workbook.xlsx -o ./my-output
```

Skip screenshots (or required on macOS):
```
/analyze-excel /path/to/workbook.xlsx --no-screenshots
```

## How This Skill Works

This skill combines **deterministic technical extraction** with **AI-powered analysis**:

### Step 1: Technical Extraction (Python CLI)
Run the extraction tool to get raw technical data:
```bash
cd ~/.claude/skills/excel-analyzer
source .venv/bin/activate  # .venv\Scripts\activate on Windows
python -m src.main "<file_path>" -o "<output_dir>"
```

### Step 2: AI Analysis (Claude)
After extraction completes, read and analyze the generated files:

1. **Read the summary**: `<output_dir>/summary.md`
2. **Read sheet details**: `<output_dir>/sheets/*.md`
3. **Read VBA code** (if present): `<output_dir>/vba/*.md`
4. **Read formulas**: `<output_dir>/formulas/*.md`
5. **View screenshots** (if available): `<output_dir>/screenshots/*.png`

### Step 3: Generate AI Insights
Based on the technical extracts, provide:

#### 1. Workbook Purpose
- What is the overall business purpose of this workbook?
- Who are the likely users?
- What decisions does it support?

#### 2. Sheet-by-Sheet Analysis
For each sheet, explain:
- Its purpose and role in the workbook
- What data it contains or calculates
- How it relates to other sheets

#### 3. Architecture & Data Flow
- How do sheets connect to each other?
- What is the data flow (inputs → calculations → outputs)?
- Draw a simple ASCII diagram if helpful

#### 4. Complexity Analysis
- What are the most complex formulas or logic?
- Are there any impressive or clever techniques?
- What would be hard to understand or maintain?

#### 5. Issues & Risks
- Are there any errors (#REF!, #NAME?, etc.)?
- Potential bugs or fragile logic?
- Security concerns (external links, macros)?
- Data quality issues?

#### 6. Recommendations
- What could be improved?
- What should users be careful about?
- Migration considerations (if moving to a web app)

## Output Structure

The Python tool creates:
```
<output_dir>/
├── report.html          # Interactive HTML report
├── README.md            # Entry point
├── summary.md           # Quick facts (read this first)
├── sheets/              # Per-sheet technical details
│   ├── _index.md
│   └── <sheet-name>.md
├── formulas/            # Formula analysis
├── vba/                 # VBA module code (if .xlsm)
├── power_query/         # M code (if present)
├── features/            # Conditional formatting, validations, etc.
├── issues/              # Errors and broken references
├── screenshots/         # Visual captures (Windows only)
└── artifacts/           # Raw extracted files
```

## Example Workflow

When user runs `/analyze-excel /path/to/budget.xlsx`:

1. **Run extraction**:
   ```bash
   cd ~/.claude/skills/excel-analyzer && source .venv/bin/activate
   python -m src.main "/path/to/budget.xlsx" -o "/path/to/budget_analysis"
   ```

2. **Read key files**:
   - Read `summary.md` for overview
   - Read each `sheets/*.md` for details
   - View screenshots to understand layout

3. **Provide analysis** covering:
   - "This workbook is a departmental budget tracker that..."
   - "The 'Input' sheet collects monthly actuals, which flow to..."
   - "Data flows: Input → Calculations → Summary → Charts"
   - "The SUMIFS formula on row 45 is complex because..."
   - "Risk: External link to `\\server\data.xlsx` may break if..."
   - "Recommendation: The manual copy-paste step could be automated..."

## Requirements

- Python 3.11+
- Virtual environment set up (run `setup-windows.bat` on Windows)
- **Screenshots**: Windows only (macOS not supported due to Excel automation limitations)

## First Run

If the virtual environment doesn't exist:
```bash
cd ~/.claude/skills/excel-analyzer
python -m venv .venv
source .venv/bin/activate  # .venv\Scripts\activate on Windows
pip install -e ".[dev]"
pip install pywin32 pillow  # Windows only, for screenshots
```

## Limitations

- **Screenshots (macOS)**: Not supported. Use `--no-screenshots` flag.
- **DAX/Power Pivot**: Can detect but not fully extract (proprietary format).
- **Very Hidden Sheets**: Documented but not screenshotted.
- **ActiveX Controls**: Limited extraction.
