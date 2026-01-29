# xls-extract

A comprehensive Python library for extracting structured data from Excel workbooks (.xlsx, .xlsm).

## Features

**xls-extract** parses Excel files and returns structured Python objects containing:

| Category | What's Extracted |
|----------|------------------|
| **Sheets** | Names, visibility (visible/hidden/very hidden), dimensions, tab colors |
| **Formulas** | All formulas with category classification, cleaned syntax, external references |
| **Named Ranges** | Global and sheet-scoped names, including LAMBDA function definitions |
| **VBA Macros** | Module code, procedure names, module types (requires oletools) |
| **Power Query** | M code for all queries, load settings |
| **Data Connections** | ODBC, OLEDB, web queries, DAX queries |
| **Pivot Tables** | Structure, source data, row/column/data/filter fields |
| **Charts** | Type, title, data ranges, position |
| **Tables** | Structured tables (ListObjects) with columns and styles |
| **Conditional Formatting** | All rule types (color scales, data bars, icon sets, formulas, etc.) |
| **Data Validation** | Dropdowns, constraints, custom formulas, error messages |
| **Comments** | Classic comments and threaded comments with replies |
| **Hyperlinks** | Cell hyperlinks with targets and display text |
| **Controls** | Form controls and buttons with macro assignments |
| **Protection** | Sheet and workbook protection status |
| **Errors** | Cells containing #REF!, #NAME?, #VALUE!, etc. |
| **External References** | Links to other workbooks |

## Installation

### From GitHub (private repo)

```bash
# Using SSH (recommended for developers)
pip install git+ssh://git@github.com/brianchk/xls-extract.git

# Using personal access token
pip install git+https://${GITHUB_TOKEN}@github.com/brianchk/xls-extract.git

# Specific version
pip install git+ssh://git@github.com/brianchk/xls-extract.git@v0.1.0
```

### In requirements.txt

```
git+ssh://git@github.com/brianchk/xls-extract.git@main
```

### In pyproject.toml

```toml
dependencies = [
    "xls-extract @ git+ssh://git@github.com/brianchk/xls-extract.git@main",
]
```

## Quick Start

```python
from xls_extract import analyze

# Analyze an Excel file
result = analyze("financial_report.xlsx")

# Access extracted data
print(f"Sheets: {[s.name for s in result.sheets]}")
print(f"Formulas: {len(result.formulas)}")
print(f"VBA Modules: {len(result.vba_modules)}")

# Iterate over formulas
for formula in result.formulas:
    print(f"{formula.location.address}: {formula.formula_clean}")
    print(f"  Category: {formula.category.value}")

# Check for VBA macros
if result.vba_modules:
    for module in result.vba_modules:
        print(f"Module: {module.name} ({module.module_type})")
        print(f"  Procedures: {module.procedures}")

# Get Power Query M code
for query in result.power_queries:
    print(f"Query: {query.name}")
    print(query.formula)
```

## API Reference

### Main Function

#### `analyze(file_path, options=None) -> WorkbookAnalysis`

Analyzes an Excel workbook and returns structured data.

**Parameters:**
- `file_path` (str | Path): Path to the Excel file (.xlsx or .xlsm)
- `options` (AnalysisOptions | None): Optional configuration for extraction

**Returns:**
- `WorkbookAnalysis`: Dataclass containing all extracted data

**Example:**
```python
from xls_extract import analyze, AnalysisOptions

# Basic usage
result = analyze("workbook.xlsx")

# With options
options = AnalysisOptions(
    extract_vba=True,
    extract_power_query=True,
    include_formula_values=False,
)
result = analyze("workbook.xlsm", options)
```

### Data Models

All extracted data is returned as typed dataclasses for easy access and IDE support.

#### `WorkbookAnalysis`

The main result object containing all extracted data:

```python
@dataclass
class WorkbookAnalysis:
    file_path: Path
    file_name: str
    file_size: int

    # Core content
    sheets: list[SheetInfo]
    formulas: list[FormulaInfo]
    named_ranges: list[NamedRangeInfo]

    # Features
    conditional_formats: list[ConditionalFormatInfo]
    data_validations: list[DataValidationInfo]
    pivot_tables: list[PivotTableInfo]
    charts: list[ChartInfo]
    tables: list[TableInfo]
    filters: list[AutoFilterInfo]

    # Code and queries
    vba_modules: list[VBAModuleInfo]
    power_queries: list[PowerQueryInfo]
    connections: list[DataConnectionInfo]

    # Other
    comments: list[CommentInfo]
    hyperlinks: list[HyperlinkInfo]
    controls: list[ControlInfo]
    error_cells: list[ErrorCellInfo]
    external_refs: list[ExternalRefInfo]

    # Protection
    workbook_protection: WorkbookProtectionInfo | None
    sheet_protections: list[SheetProtectionInfo]
```

#### `SheetInfo`

Information about a worksheet:

```python
@dataclass
class SheetInfo:
    name: str
    index: int
    visibility: SheetVisibility  # VISIBLE, HIDDEN, VERY_HIDDEN
    row_count: int
    col_count: int
    has_data: bool
    has_formulas: bool
    has_charts: bool
    has_pivots: bool
    has_tables: bool
    tab_color: str | None  # Hex color code
```

#### `FormulaInfo`

Information about a formula:

```python
@dataclass
class FormulaInfo:
    location: CellReference
    formula: str              # Raw formula as stored
    formula_clean: str        # Cleaned (e.g., _xlfn.XLOOKUP -> XLOOKUP)
    category: FormulaCategory # LOOKUP, DYNAMIC_ARRAY, LAMBDA, etc.
    is_array_formula: bool
    references_external: bool
    external_refs: list[str]
```

#### `VBAModuleInfo`

Information about a VBA module:

```python
@dataclass
class VBAModuleInfo:
    name: str
    module_type: str  # "Standard", "Class", "ThisWorkbook", "Sheet"
    code: str
    line_count: int
    procedures: list[str]
```

### Formula Categories

Formulas are automatically categorized:

| Category | Examples |
|----------|----------|
| `SIMPLE` | Basic arithmetic, cell references |
| `LOOKUP` | VLOOKUP, XLOOKUP, INDEX/MATCH |
| `DYNAMIC_ARRAY` | FILTER, SORT, UNIQUE, SEQUENCE |
| `LAMBDA` | LAMBDA definitions and calls |
| `AGGREGATE` | SUM, SUMIF, COUNTIF, AVERAGEIF |
| `VOLATILE` | NOW, TODAY, RAND, INDIRECT |
| `TEXT` | CONCAT, LEFT, RIGHT, MID |
| `DATE_TIME` | DATE, YEAR, MONTH, NETWORKDAYS |
| `LOGICAL` | IF, IFS, SWITCH, AND, OR |
| `FINANCIAL` | NPV, IRR, PMT, FV |
| `STATISTICAL` | AVERAGE, MEDIAN, STDEV |
| `EXTERNAL` | References other workbooks |

## Advanced Usage

### Selective Extraction

Extract only what you need for better performance:

```python
from xls_extract import analyze, AnalysisOptions

options = AnalysisOptions(
    extract_formulas=True,
    extract_vba=False,        # Skip VBA extraction
    extract_power_query=False, # Skip Power Query
    extract_charts=True,
    extract_pivots=True,
)

result = analyze("large_workbook.xlsx", options)
```

### Working with Large Files

For very large files, consider processing sheets individually:

```python
from xls_extract import open_workbook

with open_workbook("huge_file.xlsx") as wb:
    # Get sheet names without loading everything
    print(wb.sheet_names)

    # Extract specific sheet only
    sheet_data = wb.extract_sheet("Summary")
```

### Error Handling

```python
from xls_extract import analyze, ExtractionError

try:
    result = analyze("workbook.xlsx")
except FileNotFoundError:
    print("File not found")
except ExtractionError as e:
    print(f"Extraction failed: {e.message}")
    print(f"Partial results available: {e.partial_result}")
```

## Formula Translation

Excel stores modern functions with internal prefixes. xls-extract automatically translates these:

| Stored As | Displayed As |
|-----------|--------------|
| `_xlfn.XLOOKUP` | `XLOOKUP` |
| `_xlfn._xlpm.LAMBDA` | `LAMBDA` |
| `_xlfn.FILTER` | `FILTER` |
| `_xlfn.SORT` | `SORT` |
| `_xlfn.UNIQUE` | `UNIQUE` |
| `_xlfn.LET` | `LET` |
| `_xlfn.SEQUENCE` | `SEQUENCE` |

## Requirements

- Python 3.11+
- openpyxl >= 3.1.0
- oletools >= 0.60 (for VBA extraction)
- lxml >= 5.0.0 (for Power Query extraction)

## License

MIT License - see [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
