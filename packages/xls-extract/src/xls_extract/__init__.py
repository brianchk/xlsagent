"""
xls-extract: Comprehensive Excel workbook data extraction library.

This library extracts structured data from Excel workbooks (.xlsx, .xlsm),
including formulas, VBA macros, Power Query, pivot tables, charts, and more.
It also generates HTML reports, Markdown documentation, and screenshots.

Basic usage (data only):
    >>> from xls_extract import analyze
    >>> result = analyze("workbook.xlsx")
    >>> print(result.sheets)
    >>> print(result.formulas)

Full analysis with reports:
    >>> from xls_extract import analyze_and_report
    >>> result = analyze_and_report("workbook.xlsx", "./output")
    >>> # Creates: output/index.html, output/README.md, output/screenshots/
"""

from .analyze import analyze, analyze_and_report, open_workbook, AnalysisOptions
from .models import (
    # Main result
    WorkbookAnalysis,
    # Enums
    SheetVisibility,
    FormulaCategory,
    ErrorType,
    CFRuleType,
    # Core models
    CellReference,
    SheetInfo,
    FormulaInfo,
    NamedRangeInfo,
    # Features
    ConditionalFormatInfo,
    DataValidationInfo,
    PivotTableInfo,
    ChartInfo,
    TableInfo,
    AutoFilterInfo,
    # Code and queries
    VBAModuleInfo,
    PowerQueryInfo,
    DataConnectionInfo,
    # Other
    CommentInfo,
    HyperlinkInfo,
    ControlInfo,
    ErrorCellInfo,
    ExternalRefInfo,
    ScreenshotInfo,
    # Protection
    WorkbookProtectionInfo,
    SheetProtectionInfo,
    PrintSettingsInfo,
    # Errors
    ExtractionError,
    ExtractionWarning,
)

__version__ = "0.1.0"
__author__ = "Brian Chan"

__all__ = [
    # Main API
    "analyze",
    "analyze_and_report",
    "open_workbook",
    "AnalysisOptions",
    # Main result
    "WorkbookAnalysis",
    # Enums
    "SheetVisibility",
    "FormulaCategory",
    "ErrorType",
    "CFRuleType",
    # Core models
    "CellReference",
    "SheetInfo",
    "FormulaInfo",
    "NamedRangeInfo",
    # Features
    "ConditionalFormatInfo",
    "DataValidationInfo",
    "PivotTableInfo",
    "ChartInfo",
    "TableInfo",
    "AutoFilterInfo",
    # Code and queries
    "VBAModuleInfo",
    "PowerQueryInfo",
    "DataConnectionInfo",
    # Other
    "CommentInfo",
    "HyperlinkInfo",
    "ControlInfo",
    "ErrorCellInfo",
    "ExternalRefInfo",
    "ScreenshotInfo",
    # Protection
    "WorkbookProtectionInfo",
    "SheetProtectionInfo",
    "PrintSettingsInfo",
    # Errors
    "ExtractionError",
    "ExtractionWarning",
]
