"""Tests for data models."""

from __future__ import annotations

from pathlib import Path

import pytest

from src.models import (
    CellReference,
    SheetInfo,
    SheetVisibility,
    FormulaInfo,
    FormulaCategory,
    NamedRangeInfo,
    WorkbookAnalysis,
    ErrorType,
    CFRuleType,
)


class TestCellReference:
    """Tests for CellReference."""

    def test_address_property(self):
        ref = CellReference(sheet="Sheet1", cell="A1", row=1, col=1)
        assert ref.address == "'Sheet1'!A1"

    def test_address_with_special_chars(self):
        ref = CellReference(sheet="My Sheet", cell="B2", row=2, col=2)
        assert ref.address == "'My Sheet'!B2"


class TestSheetInfo:
    """Tests for SheetInfo."""

    def test_default_values(self):
        info = SheetInfo(
            name="Test",
            index=0,
            visibility=SheetVisibility.VISIBLE,
        )
        assert info.row_count == 0
        assert info.col_count == 0
        assert info.has_data is False
        assert info.has_formulas is False
        assert info.merged_cell_ranges == []

    def test_visibility_enum(self):
        assert SheetVisibility.VISIBLE.value == "visible"
        assert SheetVisibility.HIDDEN.value == "hidden"
        assert SheetVisibility.VERY_HIDDEN.value == "very_hidden"


class TestFormulaInfo:
    """Tests for FormulaInfo."""

    def test_formula_categories(self):
        assert FormulaCategory.SIMPLE.value == "simple"
        assert FormulaCategory.LAMBDA.value == "lambda"
        assert FormulaCategory.DYNAMIC_ARRAY.value == "dynamic_array"

    def test_external_refs_default(self):
        info = FormulaInfo(
            location=CellReference("Sheet1", "A1", 1, 1),
            formula="=SUM(A1:A10)",
            formula_clean="=SUM(A1:A10)",
            category=FormulaCategory.AGGREGATE,
        )
        assert info.external_refs == []
        assert info.references_external is False


class TestNamedRangeInfo:
    """Tests for NamedRangeInfo."""

    def test_lambda_detection_field(self):
        # Regular named range
        regular = NamedRangeInfo(
            name="MyRange",
            value="Sheet1!$A$1:$A$10",
            is_lambda=False,
        )
        assert regular.is_lambda is False

        # LAMBDA function
        lambda_func = NamedRangeInfo(
            name="Double",
            value="LAMBDA(x, x*2)",
            is_lambda=True,
        )
        assert lambda_func.is_lambda is True


class TestWorkbookAnalysis:
    """Tests for WorkbookAnalysis."""

    def test_default_lists_are_empty(self):
        analysis = WorkbookAnalysis(
            file_path=Path("/test.xlsx"),
            file_name="test.xlsx",
            file_size=1000,
            is_macro_enabled=False,
        )
        assert analysis.sheets == []
        assert analysis.formulas == []
        assert analysis.vba_modules == []
        assert analysis.errors == []
        assert analysis.warnings == []

    def test_macro_enabled_detection(self):
        xlsx = WorkbookAnalysis(
            file_path=Path("/test.xlsx"),
            file_name="test.xlsx",
            file_size=1000,
            is_macro_enabled=False,
        )
        assert xlsx.is_macro_enabled is False

        xlsm = WorkbookAnalysis(
            file_path=Path("/test.xlsm"),
            file_name="test.xlsm",
            file_size=1000,
            is_macro_enabled=True,
        )
        assert xlsm.is_macro_enabled is True


class TestEnums:
    """Tests for enum values."""

    def test_error_types(self):
        assert ErrorType.REF.value == "#REF!"
        assert ErrorType.NAME.value == "#NAME?"
        assert ErrorType.VALUE.value == "#VALUE!"
        assert ErrorType.DIV.value == "#DIV/0!"
        assert ErrorType.NA.value == "#N/A"

    def test_cf_rule_types(self):
        assert CFRuleType.COLOR_SCALE.value == "color_scale"
        assert CFRuleType.DATA_BAR.value == "data_bar"
        assert CFRuleType.ICON_SET.value == "icon_set"
        assert CFRuleType.CELL_IS.value == "cell_is"
        assert CFRuleType.FORMULA.value == "formula"
