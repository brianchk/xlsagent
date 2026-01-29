"""
Data models for Excel workbook analysis results.

This module contains all the dataclasses used to represent extracted data
from Excel workbooks. All models are immutable-friendly and fully typed
for excellent IDE support.

Example:
    >>> from xls_extract import analyze
    >>> result = analyze("workbook.xlsx")
    >>> for sheet in result.sheets:
    ...     print(f"{sheet.name}: {sheet.row_count} rows")
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any


# =============================================================================
# Enums
# =============================================================================


class SheetVisibility(Enum):
    """Visibility state of a worksheet.

    Attributes:
        VISIBLE: Sheet is visible in the workbook.
        HIDDEN: Sheet is hidden but can be unhidden via Excel UI.
        VERY_HIDDEN: Sheet is hidden and can only be unhidden via VBA.
    """

    VISIBLE = "visible"
    HIDDEN = "hidden"
    VERY_HIDDEN = "very_hidden"


class FormulaCategory(Enum):
    """Category of a formula based on its primary function.

    Used to classify formulas for analysis and reporting. A formula
    is categorized based on its most significant function.

    Attributes:
        SIMPLE: Basic arithmetic or cell references (e.g., =A1+B1)
        LOOKUP: Lookup functions (VLOOKUP, XLOOKUP, INDEX/MATCH)
        DYNAMIC_ARRAY: Dynamic array functions (FILTER, SORT, UNIQUE)
        LAMBDA: LAMBDA function definitions or calls
        AGGREGATE: Aggregation functions (SUM, SUMIF, COUNTIF)
        VOLATILE: Volatile functions that recalculate (NOW, TODAY, RAND)
        TEXT: Text manipulation functions (CONCAT, LEFT, MID)
        DATE_TIME: Date/time functions (DATE, YEAR, NETWORKDAYS)
        LOGICAL: Logical functions (IF, IFS, SWITCH, AND, OR)
        FINANCIAL: Financial functions (NPV, IRR, PMT)
        MATH: Mathematical functions (ROUND, ABS, SQRT)
        STATISTICAL: Statistical functions (AVERAGE, MEDIAN, STDEV)
        ERROR_HANDLING: Error handling (IFERROR, ISERROR)
        EXTERNAL: References external workbooks
        ARRAY_LEGACY: Legacy CSE array formulas ({=...})
    """

    SIMPLE = "simple"
    LOOKUP = "lookup"
    DYNAMIC_ARRAY = "dynamic_array"
    LAMBDA = "lambda"
    AGGREGATE = "aggregate"
    VOLATILE = "volatile"
    TEXT = "text"
    DATE_TIME = "date_time"
    LOGICAL = "logical"
    FINANCIAL = "financial"
    MATH = "math"
    STATISTICAL = "statistical"
    ERROR_HANDLING = "error_handling"
    EXTERNAL = "external"
    ARRAY_LEGACY = "array_legacy"


class ErrorType(Enum):
    """Types of Excel cell errors.

    Attributes:
        REF: Invalid cell reference (#REF!)
        NAME: Unrecognized formula name (#NAME?)
        VALUE: Wrong type of argument (#VALUE!)
        DIV: Division by zero (#DIV/0!)
        NULL: Incorrect range operator (#NULL!)
        NUM: Invalid numeric value (#NUM!)
        NA: Value not available (#N/A)
        CALC: Calculation error (#CALC!)
        SPILL: Spill range blocked (#SPILL!)
        GETTING_DATA: Data still loading (#GETTING_DATA)
    """

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
    """Types of conditional formatting rules.

    Attributes:
        COLOR_SCALE: 2-color or 3-color gradient scale
        DATA_BAR: In-cell data bars
        ICON_SET: Icon sets (arrows, traffic lights, etc.)
        CELL_IS: Cell value comparison (equals, between, etc.)
        FORMULA: Custom formula-based rule
        TOP_BOTTOM: Top/bottom N or N%
        ABOVE_AVERAGE: Above/below average
        DUPLICATE: Highlight duplicates
        UNIQUE: Highlight unique values
        TEXT_CONTAINS: Text contains/begins/ends with
        DATE_OCCURRING: Date-based rules
        BLANK: Highlight blank cells
        NOT_BLANK: Highlight non-blank cells
        ERROR: Highlight error cells
        NOT_ERROR: Highlight non-error cells
    """

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


# =============================================================================
# Core Models
# =============================================================================


@dataclass
class CellReference:
    """Reference to a specific cell location in the workbook.

    Attributes:
        sheet: Name of the worksheet containing the cell.
        cell: Cell address in A1 notation (e.g., "A1", "B2").
        row: 1-based row number.
        col: 1-based column number.

    Example:
        >>> ref = CellReference(sheet="Sheet1", cell="B5", row=5, col=2)
        >>> print(ref.address)
        'Sheet1'!B5
    """

    sheet: str
    cell: str
    row: int
    col: int

    @property
    def address(self) -> str:
        """Full address including sheet name (e.g., 'Sheet1'!A1)."""
        return f"'{self.sheet}'!{self.cell}"


@dataclass
class SheetInfo:
    """Information about a worksheet.

    Attributes:
        name: Name of the sheet as shown on the tab.
        index: 0-based position in the workbook.
        visibility: Whether the sheet is visible, hidden, or very hidden.
        used_range: Address of the used range (e.g., "A1:Z100").
        row_count: Number of rows in the used range.
        col_count: Number of columns in the used range.
        has_data: Whether the sheet contains any data.
        has_formulas: Whether the sheet contains formulas.
        has_charts: Whether the sheet contains embedded charts.
        has_pivots: Whether the sheet contains pivot tables.
        has_tables: Whether the sheet contains structured tables.
        has_comments: Whether the sheet contains comments.
        has_conditional_formatting: Whether the sheet has CF rules.
        has_data_validation: Whether the sheet has validation rules.
        has_hyperlinks: Whether the sheet contains hyperlinks.
        has_merged_cells: Whether the sheet has merged cells.
        merged_cell_ranges: List of merged cell ranges.
        tab_color: Hex color of the sheet tab (e.g., "#FF0000").
    """

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
    """Information about a formula in the workbook.

    The formula_clean field contains the formula with Excel's internal
    prefixes translated to human-readable form (e.g., _xlfn.XLOOKUP
    becomes XLOOKUP).

    Attributes:
        location: Cell location of the formula.
        formula: Raw formula as stored in the file.
        formula_clean: Formula with internal prefixes translated.
        category: Classification of the formula type.
        result_value: Cached result value (if available).
        is_array_formula: Whether this is a legacy CSE array formula.
        spill_range: Range that this dynamic array spills to.
        references_external: Whether formula references other workbooks.
        external_refs: List of external workbook references.

    Example:
        >>> for f in result.formulas:
        ...     if f.category == FormulaCategory.LOOKUP:
        ...         print(f"{f.location.cell}: {f.formula_clean}")
    """

    location: CellReference
    formula: str
    formula_clean: str
    category: FormulaCategory
    result_value: Any = None
    is_array_formula: bool = False
    spill_range: str | None = None
    references_external: bool = False
    external_refs: list[str] = field(default_factory=list)


@dataclass
class NamedRangeInfo:
    """Information about a named range or named formula.

    Named ranges can be global (workbook-scoped) or local (sheet-scoped).
    LAMBDA functions are also stored as named ranges.

    Attributes:
        name: The defined name.
        value: The formula or range reference.
        scope: Sheet name if local scope, None if global.
        is_lambda: Whether this is a LAMBDA function definition.
        comment: Optional comment/description.
        hidden: Whether the name is hidden from the UI.

    Example:
        >>> for nr in result.named_ranges:
        ...     if nr.is_lambda:
        ...         print(f"LAMBDA: {nr.name}")
        ...         print(f"  {nr.value}")
    """

    name: str
    value: str
    scope: str | None = None
    is_lambda: bool = False
    comment: str | None = None
    hidden: bool = False


# =============================================================================
# Feature Models
# =============================================================================


@dataclass
class ConditionalFormatInfo:
    """Information about a conditional formatting rule.

    Attributes:
        range: Cell range the rule applies to.
        rule_type: Type of conditional formatting rule.
        priority: Rule priority (lower = higher priority).
        formula: Formula for formula-based rules.
        operator: Comparison operator for cell_is rules.
        values: Threshold values for scales/bars/icons.
        stop_if_true: Whether to stop evaluating if this rule matches.
        description: Human-readable description of the rule.
    """

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
    """Information about data validation rules.

    Attributes:
        range: Cell range the validation applies to.
        type: Validation type (list, whole, decimal, date, etc.).
        operator: Comparison operator (between, equal, etc.).
        formula1: First formula/value for validation.
        formula2: Second formula/value (for between operator).
        allow_blank: Whether blank values are allowed.
        show_dropdown: Whether to show dropdown for list validation.
        show_input_message: Whether to show input message.
        input_title: Title of the input message.
        input_message: Body of the input message.
        show_error_message: Whether to show error on invalid input.
        error_title: Title of the error message.
        error_message: Body of the error message.
        error_style: Error style (stop, warning, information).
    """

    range: str
    type: str
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
    error_style: str | None = None


@dataclass
class PivotTableInfo:
    """Information about a pivot table.

    Attributes:
        name: Name of the pivot table.
        sheet: Sheet containing the pivot table.
        location: Cell address where pivot table starts.
        source_range: Source data range (if range-based).
        source_connection: Connection name (if connection-based).
        row_fields: Fields in the row area.
        column_fields: Fields in the column area.
        data_fields: Fields in the values area.
        filter_fields: Fields in the filter area.
        cache_id: Pivot cache ID for cross-referencing.
    """

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
    """Information about an embedded chart.

    Attributes:
        name: Name of the chart object.
        sheet: Sheet containing the chart.
        chart_type: Type of chart (Bar, Line, Pie, etc.).
        title: Chart title text.
        data_range: Source data range for the chart.
        position: Position description in the sheet.
    """

    name: str
    sheet: str
    chart_type: str
    title: str | None = None
    data_range: str | None = None
    position: str | None = None


@dataclass
class TableInfo:
    """Information about a structured table (ListObject).

    Attributes:
        name: Internal name of the table.
        sheet: Sheet containing the table.
        range: Cell range of the table.
        display_name: Display name of the table.
        columns: List of column names.
        has_totals_row: Whether the table has a totals row.
        has_header_row: Whether the table has a header row.
        style_name: Name of the table style applied.
    """

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
    """Information about AutoFilter settings on a range.

    Attributes:
        sheet: Sheet containing the filter.
        range: Range the filter applies to.
        column_filters: Dict mapping column index to filter criteria.
    """

    sheet: str
    range: str
    column_filters: dict[int, dict[str, Any]] = field(default_factory=dict)


# =============================================================================
# Code and Query Models
# =============================================================================


@dataclass
class VBAModuleInfo:
    """Information about a VBA module.

    Attributes:
        name: Name of the module.
        module_type: Type of module (Standard, Class, ThisWorkbook, Sheet).
        code: Full source code of the module.
        line_count: Number of lines of code.
        procedures: List of Sub/Function names in the module.

    Example:
        >>> for module in result.vba_modules:
        ...     print(f"{module.name} ({module.module_type})")
        ...     for proc in module.procedures:
        ...         print(f"  - {proc}")
    """

    name: str
    module_type: str
    code: str
    line_count: int = 0
    procedures: list[str] = field(default_factory=list)


@dataclass
class PowerQueryInfo:
    """Information about a Power Query (M code).

    Attributes:
        name: Name of the query.
        formula: The M code for the query.
        description: Optional description of the query.
        load_enabled: Whether the query loads to the data model.
        result_type: Expected result type of the query.

    Example:
        >>> for query in result.power_queries:
        ...     print(f"Query: {query.name}")
        ...     print(query.formula)
    """

    name: str
    formula: str
    description: str | None = None
    load_enabled: bool = True
    result_type: str | None = None


@dataclass
class DataConnectionInfo:
    """Information about a data connection.

    Connections can be ODBC, OLEDB, web queries, or Power Pivot
    connections with DAX queries.

    Attributes:
        name: Name of the connection.
        connection_type: Type (ODBC, OLEDB, Web, etc.).
        connection_string: Full connection string.
        command_text: SQL or other command text.
        command_type: Type of command (SQL, DAX, Table, etc.).
        description: Optional description.
        is_dax: Whether this contains a DAX query.
        dax_query: The DAX query if applicable.
        connection_id: Internal connection ID.
        used_by_pivot_caches: List of pivot caches using this connection.
    """

    name: str
    connection_type: str
    connection_string: str | None = None
    command_text: str | None = None
    command_type: str | None = None
    description: str | None = None
    is_dax: bool = False
    dax_query: str | None = None
    connection_id: str | None = None
    used_by_pivot_caches: list[str] = field(default_factory=list)


# =============================================================================
# Other Content Models
# =============================================================================


@dataclass
class CommentInfo:
    """Information about a cell comment.

    Supports both classic comments and modern threaded comments.

    Attributes:
        location: Cell containing the comment.
        author: Author of the comment.
        text: Comment text content.
        is_threaded: Whether this is a threaded comment.
        replies: List of reply comments (for threaded comments).
    """

    location: CellReference
    author: str | None = None
    text: str = ""
    is_threaded: bool = False
    replies: list["CommentInfo"] = field(default_factory=list)


@dataclass
class HyperlinkInfo:
    """Information about a hyperlink.

    Attributes:
        location: Cell containing the hyperlink.
        target: URL or cell reference target.
        display_text: Text displayed in the cell.
        tooltip: Hover tooltip text.
        is_external: Whether the link points outside the workbook.
    """

    location: CellReference
    target: str
    display_text: str | None = None
    tooltip: str | None = None
    is_external: bool = False


@dataclass
class ControlInfo:
    """Information about a form control or ActiveX control.

    Attributes:
        name: Name of the control.
        sheet: Sheet containing the control.
        control_type: Type (Button, CheckBox, ComboBox, etc.).
        position: Position in the sheet.
        linked_cell: Cell linked to the control value.
        macro: Assigned macro name.
        text: Text/caption of the control.
    """

    name: str
    sheet: str
    control_type: str
    position: str | None = None
    linked_cell: str | None = None
    macro: str | None = None
    text: str | None = None


@dataclass
class ErrorCellInfo:
    """Information about a cell containing an error value.

    Attributes:
        location: Cell containing the error.
        error_type: Type of error (#REF!, #NAME?, etc.).
        formula: Formula that produced the error (if any).
    """

    location: CellReference
    error_type: ErrorType
    formula: str | None = None


@dataclass
class ExternalRefInfo:
    """Information about a reference to another workbook.

    Attributes:
        source_cell: Cell containing the external reference.
        target_workbook: Name/path of the referenced workbook.
        target_sheet: Referenced sheet name (if specified).
        target_range: Referenced range (if specified).
        is_broken: Whether the reference cannot be resolved.
    """

    source_cell: CellReference
    target_workbook: str
    target_sheet: str | None = None
    target_range: str | None = None
    is_broken: bool = False


# =============================================================================
# Screenshot Model (for consumers that capture screenshots)
# =============================================================================


@dataclass
class ScreenshotInfo:
    """Information about a captured screenshot.

    This model is used by consumers of xls-extract that capture
    screenshots of the workbook (e.g., via desktop Excel automation).

    Attributes:
        sheet: Name of the sheet captured.
        path: File path where the screenshot is saved.
        width: Width of the image in pixels.
        height: Height of the image in pixels.
        captured_at: ISO timestamp when captured.
        is_chart: Whether this is a chart screenshot (vs. sheet).
        chart_name: Name of the chart if is_chart is True.
    """

    sheet: str
    path: Path
    width: int = 0
    height: int = 0
    captured_at: str | None = None
    is_chart: bool = False
    chart_name: str | None = None


# =============================================================================
# Protection Models
# =============================================================================


@dataclass
class WorkbookProtectionInfo:
    """Information about workbook-level protection.

    Attributes:
        is_protected: Whether the workbook is protected.
        protect_structure: Whether structure changes are blocked.
        protect_windows: Whether window changes are blocked.
    """

    is_protected: bool = False
    protect_structure: bool = False
    protect_windows: bool = False


@dataclass
class SheetProtectionInfo:
    """Information about sheet-level protection.

    Attributes:
        sheet: Name of the protected sheet.
        is_protected: Whether the sheet is protected.
        allow_select_locked: Whether selecting locked cells is allowed.
        allow_select_unlocked: Whether selecting unlocked cells is allowed.
        allow_format_cells: Whether formatting cells is allowed.
        allow_format_columns: Whether formatting columns is allowed.
        allow_format_rows: Whether formatting rows is allowed.
        allow_insert_columns: Whether inserting columns is allowed.
        allow_insert_rows: Whether inserting rows is allowed.
        allow_insert_hyperlinks: Whether inserting hyperlinks is allowed.
        allow_delete_columns: Whether deleting columns is allowed.
        allow_delete_rows: Whether deleting rows is allowed.
        allow_sort: Whether sorting is allowed.
        allow_filter: Whether filtering is allowed.
        allow_pivot_tables: Whether pivot tables are allowed.
    """

    sheet: str
    is_protected: bool = False
    allow_select_locked: bool = True
    allow_select_unlocked: bool = True
    allow_format_cells: bool = False
    allow_format_columns: bool = False
    allow_format_rows: bool = False
    allow_insert_columns: bool = False
    allow_insert_rows: bool = False
    allow_insert_hyperlinks: bool = False
    allow_delete_columns: bool = False
    allow_delete_rows: bool = False
    allow_sort: bool = False
    allow_filter: bool = False
    allow_pivot_tables: bool = False


@dataclass
class PrintSettingsInfo:
    """Information about print settings for a sheet.

    Attributes:
        sheet: Name of the sheet.
        print_area: Defined print area range.
        print_titles_rows: Rows to repeat at top of each page.
        print_titles_cols: Columns to repeat at left of each page.
        page_breaks_row: Row numbers with manual page breaks.
        page_breaks_col: Column numbers with manual page breaks.
        orientation: Page orientation (portrait/landscape).
        paper_size: Paper size name.
        fit_to_page: Whether to fit to page.
        fit_to_width: Number of pages wide to fit to.
        fit_to_height: Number of pages tall to fit to.
    """

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


# =============================================================================
# Error Handling
# =============================================================================


@dataclass
class ExtractionError:
    """An error that occurred during extraction.

    Non-fatal errors are collected here rather than raising exceptions,
    allowing partial extraction results to be returned.

    Attributes:
        extractor: Name of the extractor that failed.
        message: Human-readable error message.
        details: Additional technical details.
    """

    extractor: str
    message: str
    details: str | None = None


@dataclass
class ExtractionWarning:
    """A warning from extraction indicating partial success.

    Attributes:
        extractor: Name of the extractor that generated the warning.
        message: Human-readable warning message.
        details: Additional details.
    """

    extractor: str
    message: str
    details: str | None = None


# =============================================================================
# Main Result Model
# =============================================================================


@dataclass
class WorkbookAnalysis:
    """Complete analysis results for an Excel workbook.

    This is the main result object returned by the analyze() function.
    It contains all extracted data organized by category.

    Attributes:
        file_path: Path to the analyzed file.
        file_name: Name of the file.
        file_size: Size in bytes.
        is_macro_enabled: Whether the file can contain macros (.xlsm).

    Example:
        >>> result = analyze("report.xlsx")
        >>> print(f"File: {result.file_name}")
        >>> print(f"Sheets: {len(result.sheets)}")
        >>> print(f"Formulas: {len(result.formulas)}")
        >>> if result.vba_modules:
        ...     print("Contains VBA macros!")
    """

    # File info
    file_path: Path
    file_name: str
    file_size: int
    is_macro_enabled: bool

    # Sheet information
    sheets: list[SheetInfo] = field(default_factory=list)

    # Formulas and names
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

    # Protection
    workbook_protection: WorkbookProtectionInfo | None = None
    sheet_protections: list[SheetProtectionInfo] = field(default_factory=list)
    print_settings: list[PrintSettingsInfo] = field(default_factory=list)

    # Code and queries
    vba_modules: list[VBAModuleInfo] = field(default_factory=list)
    vba_project_name: str | None = None
    power_queries: list[PowerQueryInfo] = field(default_factory=list)

    # Issues
    error_cells: list[ErrorCellInfo] = field(default_factory=list)
    external_refs: list[ExternalRefInfo] = field(default_factory=list)

    # DAX/Power Pivot detection
    has_dax: bool = False
    dax_detection_note: str | None = None

    # Screenshots (populated by consumers that capture them)
    screenshots: list[ScreenshotInfo] = field(default_factory=list)

    # Extraction status
    errors: list[ExtractionError] = field(default_factory=list)
    warnings: list[ExtractionWarning] = field(default_factory=list)

    @property
    def has_vba(self) -> bool:
        """Whether the workbook contains VBA code."""
        return len(self.vba_modules) > 0

    @property
    def has_power_query(self) -> bool:
        """Whether the workbook contains Power Query."""
        return len(self.power_queries) > 0

    @property
    def has_external_refs(self) -> bool:
        """Whether the workbook references other workbooks."""
        return len(self.external_refs) > 0

    @property
    def has_errors(self) -> bool:
        """Whether extraction encountered any errors."""
        return len(self.errors) > 0

    @property
    def visible_sheets(self) -> list[SheetInfo]:
        """List of visible sheets only."""
        return [s for s in self.sheets if s.visibility == SheetVisibility.VISIBLE]

    @property
    def hidden_sheets(self) -> list[SheetInfo]:
        """List of hidden and very hidden sheets."""
        return [s for s in self.sheets if s.visibility != SheetVisibility.VISIBLE]
