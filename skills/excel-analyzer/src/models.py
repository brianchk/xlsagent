"""Data models for Excel analysis results."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional


class SheetVisibility(Enum):
    VISIBLE = "visible"
    HIDDEN = "hidden"
    VERY_HIDDEN = "very_hidden"


class FormulaCategory(Enum):
    SIMPLE = "simple"
    LOOKUP = "lookup"
    DYNAMIC_ARRAY = "dynamic_array"
    LAMBDA = "lambda"
    AGGREGATE = "aggregate"
    VOLATILE = "volatile"
    ARRAY_LEGACY = "array_legacy"
    TEXT = "text"
    DATE_TIME = "date_time"
    LOGICAL = "logical"
    FINANCIAL = "financial"
    MATH = "math"
    STATISTICAL = "statistical"
    ERROR_HANDLING = "error_handling"
    EXTERNAL = "external"


class ErrorType(Enum):
    REF = "#REF!"
    NAME = "#NAME?"
    VALUE = "#VALUE!"
    DIV = "#DIV/0!"
    NULL = "#NULL!"
    NUM = "#NUM!"
    NA = "#N/A"
    CALC = "#CALC!"
    SPILL = "#SPILL!"
    GETTING_DATA = "#GETTING_DATA"


class CFRuleType(Enum):
    COLOR_SCALE = "color_scale"
    DATA_BAR = "data_bar"
    ICON_SET = "icon_set"
    CELL_IS = "cell_is"
    FORMULA = "formula"
    TOP_BOTTOM = "top_bottom"
    ABOVE_AVERAGE = "above_average"
    DUPLICATE = "duplicate"
    UNIQUE = "unique"
    TEXT_CONTAINS = "text_contains"
    DATE_OCCURRING = "date_occurring"
    BLANK = "blank"
    NOT_BLANK = "not_blank"
    ERROR = "error"
    NOT_ERROR = "not_error"


@dataclass
class CellReference:
    """Reference to a cell location."""
    sheet: str
    cell: str
    row: int
    col: int

    @property
    def address(self) -> str:
        return f"'{self.sheet}'!{self.cell}"


@dataclass
class SheetInfo:
    """Information about a worksheet."""
    name: str
    index: int
    visibility: SheetVisibility
    used_range: str | None = None
    row_count: int = 0
    col_count: int = 0
    has_data: bool = False
    has_formulas: bool = False
    has_charts: bool = False
    has_pivots: bool = False
    has_tables: bool = False
    has_comments: bool = False
    has_conditional_formatting: bool = False
    has_data_validation: bool = False
    has_hyperlinks: bool = False
    has_merged_cells: bool = False
    merged_cell_ranges: list[str] = field(default_factory=list)
    tab_color: str | None = None


@dataclass
class FormulaInfo:
    """Information about a formula."""
    location: CellReference
    formula: str
    formula_clean: str  # With _xlfn. prefixes translated
    category: FormulaCategory
    result_value: Any = None
    is_array_formula: bool = False
    spill_range: str | None = None
    references_external: bool = False
    external_refs: list[str] = field(default_factory=list)


@dataclass
class NamedRangeInfo:
    """Information about a named range or LAMBDA function."""
    name: str
    value: str  # The formula or range reference
    scope: str | None = None  # Sheet name if local, None if global
    is_lambda: bool = False
    comment: str | None = None
    hidden: bool = False


@dataclass
class ConditionalFormatInfo:
    """Information about a conditional formatting rule."""
    range: str
    rule_type: CFRuleType
    priority: int
    formula: str | None = None
    operator: str | None = None
    values: list[Any] = field(default_factory=list)
    stop_if_true: bool = False
    description: str = ""


@dataclass
class DataValidationInfo:
    """Information about data validation rules."""
    range: str
    type: str  # list, whole, decimal, date, time, textLength, custom
    operator: str | None = None
    formula1: str | None = None
    formula2: str | None = None
    allow_blank: bool = True
    show_dropdown: bool = True
    show_input_message: bool = False
    input_title: str | None = None
    input_message: str | None = None
    show_error_message: bool = False
    error_title: str | None = None
    error_message: str | None = None
    error_style: str | None = None  # stop, warning, information


@dataclass
class PivotTableInfo:
    """Information about a pivot table."""
    name: str
    sheet: str
    location: str
    source_range: str | None = None
    source_connection: str | None = None
    row_fields: list[str] = field(default_factory=list)
    column_fields: list[str] = field(default_factory=list)
    data_fields: list[str] = field(default_factory=list)
    filter_fields: list[str] = field(default_factory=list)
    cache_id: int | None = None


@dataclass
class ChartInfo:
    """Information about a chart."""
    name: str
    sheet: str
    chart_type: str
    title: str | None = None
    data_range: str | None = None
    position: str | None = None


@dataclass
class TableInfo:
    """Information about a structured table (ListObject)."""
    name: str
    sheet: str
    range: str
    display_name: str
    columns: list[str] = field(default_factory=list)
    has_totals_row: bool = False
    has_header_row: bool = True
    style_name: str | None = None


@dataclass
class AutoFilterInfo:
    """Information about AutoFilter settings."""
    sheet: str
    range: str
    column_filters: dict[int, dict] = field(default_factory=dict)


@dataclass
class VBAModuleInfo:
    """Information about a VBA module."""
    name: str
    module_type: str  # Standard, Class, ThisWorkbook, Sheet
    code: str
    line_count: int = 0
    procedures: list[str] = field(default_factory=list)


@dataclass
class PowerQueryInfo:
    """Information about a Power Query."""
    name: str
    formula: str  # The M code
    description: str | None = None
    load_enabled: bool = True
    result_type: str | None = None


@dataclass
class DataConnectionInfo:
    """Information about a data connection."""
    name: str
    connection_type: str  # ODBC, OLEDB, Web, etc.
    connection_string: str | None = None
    command_text: str | None = None
    description: str | None = None


@dataclass
class CommentInfo:
    """Information about a cell comment."""
    location: CellReference
    author: str | None = None
    text: str = ""
    is_threaded: bool = False
    replies: list["CommentInfo"] = field(default_factory=list)


@dataclass
class HyperlinkInfo:
    """Information about a hyperlink."""
    location: CellReference
    target: str
    display_text: str | None = None
    tooltip: str | None = None
    is_external: bool = False


@dataclass
class ControlInfo:
    """Information about a form control or shape."""
    name: str
    sheet: str
    control_type: str  # Button, CheckBox, ComboBox, etc.
    position: str | None = None
    linked_cell: str | None = None
    macro: str | None = None
    text: str | None = None


@dataclass
class ProtectionInfo:
    """Information about protection settings."""
    workbook_protected: bool = False
    workbook_structure: bool = False
    workbook_windows: bool = False
    sheets: dict[str, dict] = field(default_factory=dict)  # sheet -> protection details


@dataclass
class PrintSettingsInfo:
    """Information about print settings."""
    sheet: str
    print_area: str | None = None
    print_titles_rows: str | None = None
    print_titles_cols: str | None = None
    page_breaks_row: list[int] = field(default_factory=list)
    page_breaks_col: list[int] = field(default_factory=list)
    orientation: str = "portrait"
    paper_size: str | None = None
    fit_to_page: bool = False
    fit_to_width: int | None = None
    fit_to_height: int | None = None


@dataclass
class ErrorCellInfo:
    """Information about a cell containing an error."""
    location: CellReference
    error_type: ErrorType
    formula: str | None = None


@dataclass
class ExternalRefInfo:
    """Information about an external workbook reference."""
    source_cell: CellReference
    target_workbook: str
    target_sheet: str | None = None
    target_range: str | None = None
    is_broken: bool = False


@dataclass
class ScreenshotInfo:
    """Information about a captured screenshot."""
    sheet: str
    path: Path
    width: int = 0
    height: int = 0
    captured_at: str | None = None


@dataclass
class ExtractionError:
    """An error that occurred during extraction."""
    extractor: str
    message: str
    details: str | None = None


@dataclass
class ExtractionWarning:
    """A warning from extraction (partial success)."""
    extractor: str
    message: str
    details: str | None = None


@dataclass
class WorkbookAnalysis:
    """Complete analysis results for a workbook."""
    file_path: Path
    file_name: str
    file_size: int
    is_macro_enabled: bool

    # Sheet information
    sheets: list[SheetInfo] = field(default_factory=list)

    # Formulas
    formulas: list[FormulaInfo] = field(default_factory=list)
    named_ranges: list[NamedRangeInfo] = field(default_factory=list)

    # Features
    conditional_formats: list[ConditionalFormatInfo] = field(default_factory=list)
    data_validations: list[DataValidationInfo] = field(default_factory=list)
    pivot_tables: list[PivotTableInfo] = field(default_factory=list)
    charts: list[ChartInfo] = field(default_factory=list)
    tables: list[TableInfo] = field(default_factory=list)
    auto_filters: list[AutoFilterInfo] = field(default_factory=list)
    controls: list[ControlInfo] = field(default_factory=list)
    connections: list[DataConnectionInfo] = field(default_factory=list)
    comments: list[CommentInfo] = field(default_factory=list)
    hyperlinks: list[HyperlinkInfo] = field(default_factory=list)
    protection: ProtectionInfo | None = None
    print_settings: list[PrintSettingsInfo] = field(default_factory=list)

    # VBA/Macros
    vba_modules: list[VBAModuleInfo] = field(default_factory=list)
    vba_project_name: str | None = None

    # Power Query
    power_queries: list[PowerQueryInfo] = field(default_factory=list)

    # Issues
    error_cells: list[ErrorCellInfo] = field(default_factory=list)
    external_refs: list[ExternalRefInfo] = field(default_factory=list)

    # Screenshots
    screenshots: list[ScreenshotInfo] = field(default_factory=list)

    # Metadata
    has_dax: bool = False
    dax_detection_note: str | None = None

    # Errors and warnings from extraction
    errors: list[ExtractionError] = field(default_factory=list)
    warnings: list[ExtractionWarning] = field(default_factory=list)
