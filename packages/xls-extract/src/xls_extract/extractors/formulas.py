"""Formula extractor with classification and cleanup."""

from __future__ import annotations

import re
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell

from ..models import CellReference, FormulaCategory, FormulaInfo
from .base import BaseExtractor


class FormulaExtractor(BaseExtractor):
    """Extracts and classifies all formulas in the workbook."""

    name = "formulas"

    # Prefix translations for modern Excel functions
    PREFIX_TRANSLATIONS = {
        "_xlfn.": "",
        "_xlpm.": "",
        "ANCHORARRAY(": "",  # Will handle specially
    }

    # Function patterns for classification
    DYNAMIC_ARRAY_FUNCTIONS = {
        "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY",
        "XLOOKUP", "XMATCH", "LET", "LAMBDA", "MAP", "REDUCE", "SCAN",
        "MAKEARRAY", "BYROW", "BYCOL", "ISOMITTED", "CHOOSECOLS", "CHOOSEROWS",
        "DROP", "TAKE", "EXPAND", "VSTACK", "HSTACK", "TOROW", "TOCOL",
        "WRAPROWS", "WRAPCOLS", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER",
    }

    LOOKUP_FUNCTIONS = {
        "VLOOKUP", "HLOOKUP", "LOOKUP", "INDEX", "MATCH", "XLOOKUP", "XMATCH",
        "OFFSET", "INDIRECT", "CHOOSE", "GETPIVOTDATA",
    }

    VOLATILE_FUNCTIONS = {
        "NOW", "TODAY", "RAND", "RANDBETWEEN", "INDIRECT", "OFFSET", "INFO",
        "CELL", "RANDARRAY",
    }

    AGGREGATE_FUNCTIONS = {
        "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "COUNT", "COUNTA", "COUNTIF",
        "COUNTIFS", "COUNTBLANK", "AVERAGE", "AVERAGEIF", "AVERAGEIFS",
        "MIN", "MINIFS", "MAX", "MAXIFS", "AGGREGATE", "SUBTOTAL",
    }

    TEXT_FUNCTIONS = {
        "CONCATENATE", "CONCAT", "TEXTJOIN", "LEFT", "RIGHT", "MID", "LEN",
        "FIND", "SEARCH", "SUBSTITUTE", "REPLACE", "TRIM", "CLEAN", "UPPER",
        "LOWER", "PROPER", "TEXT", "VALUE", "FIXED", "DOLLAR", "CHAR", "CODE",
        "REPT", "EXACT", "T", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER",
    }

    DATE_TIME_FUNCTIONS = {
        "DATE", "DATEVALUE", "TIME", "TIMEVALUE", "NOW", "TODAY", "YEAR",
        "MONTH", "DAY", "HOUR", "MINUTE", "SECOND", "WEEKDAY", "WEEKNUM",
        "ISOWEEKNUM", "NETWORKDAYS", "WORKDAY", "EDATE", "EOMONTH", "DATEDIF",
    }

    LOGICAL_FUNCTIONS = {
        "IF", "IFS", "SWITCH", "AND", "OR", "NOT", "XOR", "TRUE", "FALSE",
        "IFERROR", "IFNA", "ISERROR", "ISNA", "ISBLANK", "ISNUMBER", "ISTEXT",
        "ISLOGICAL", "ISREF", "ISERR", "ISEVEN", "ISODD", "ISFORMULA",
    }

    FINANCIAL_FUNCTIONS = {
        "PMT", "IPMT", "PPMT", "FV", "PV", "NPV", "IRR", "MIRR", "XNPV", "XIRR",
        "RATE", "NPER", "SLN", "SYD", "DB", "DDB", "VDB",
    }

    MATH_FUNCTIONS = {
        "ABS", "SIGN", "ROUND", "ROUNDUP", "ROUNDDOWN", "CEILING", "FLOOR",
        "INT", "TRUNC", "MOD", "POWER", "SQRT", "EXP", "LN", "LOG", "LOG10",
        "PRODUCT", "QUOTIENT", "RAND", "RANDBETWEEN", "PI", "DEGREES", "RADIANS",
        "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "ATAN2",
    }

    ERROR_HANDLING_FUNCTIONS = {
        "IFERROR", "IFNA", "ISERROR", "ISNA", "ISERR", "ERROR.TYPE",
    }

    def extract(self) -> list[FormulaInfo]:
        """Extract all formulas from the workbook.

        Returns:
            List of FormulaInfo objects
        """
        formulas = []

        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            if not isinstance(sheet, Worksheet):
                continue

            sheet_formulas = self._extract_sheet_formulas(sheet, sheet_name)
            formulas.extend(sheet_formulas)

        return formulas

    def _extract_sheet_formulas(self, sheet: Worksheet, sheet_name: str) -> list[FormulaInfo]:
        """Extract formulas from a single sheet."""
        formulas = []

        try:
            for row_idx, row in enumerate(sheet.iter_rows(), start=1):
                for cell in row:
                    if self._is_formula_cell(cell):
                        formula_info = self._create_formula_info(cell, sheet_name)
                        if formula_info:
                            formulas.append(formula_info)
        except Exception:
            # Sheet may be malformed or chartsheet
            pass

        return formulas

    def _is_formula_cell(self, cell: Cell) -> bool:
        """Check if cell contains a formula."""
        if cell.value is None:
            return False
        if isinstance(cell.value, str) and cell.value.startswith("="):
            return True
        # Check data_type for formula
        try:
            return cell.data_type == "f"
        except Exception:
            return False

    def _create_formula_info(self, cell: Cell, sheet_name: str) -> FormulaInfo | None:
        """Create FormulaInfo from a cell."""
        try:
            formula = str(cell.value) if cell.value else ""
            if not formula.startswith("="):
                return None

            # Clean up formula (translate prefixes)
            formula_clean = self._clean_formula(formula)

            # Skip empty formulas (just "=" or whitespace after =)
            if formula_clean.strip() == "=" or len(formula_clean.strip()) <= 1:
                return None

            # Classify formula
            category = self._classify_formula(formula_clean)

            # Check for array formula
            is_array = self._is_array_formula(cell)

            # Check for external references
            external_refs = self._extract_external_refs(formula)

            return FormulaInfo(
                location=CellReference(
                    sheet=sheet_name,
                    cell=cell.coordinate,
                    row=cell.row,
                    col=cell.column,
                ),
                formula=formula,
                formula_clean=formula_clean,
                category=category,
                is_array_formula=is_array,
                references_external=len(external_refs) > 0,
                external_refs=external_refs,
            )
        except Exception:
            return None

    def _clean_formula(self, formula: str) -> str:
        """Clean formula by translating internal prefixes to user-friendly names."""
        cleaned = formula

        # Remove _xlfn. and _xlpm. prefixes
        cleaned = re.sub(r"_xlfn\.", "", cleaned)
        cleaned = re.sub(r"_xlpm\.", "", cleaned)

        # Convert ANCHORARRAY to spill notation
        # ANCHORARRAY(A1) -> A1#
        cleaned = re.sub(r"ANCHORARRAY\(([^)]+)\)", r"\1#", cleaned)

        return cleaned

    def _classify_formula(self, formula: str) -> FormulaCategory:
        """Classify a formula based on the functions it uses."""
        formula_upper = formula.upper()

        # Extract function names from formula
        functions = set(re.findall(r"([A-Z][A-Z0-9_.]+)\s*\(", formula_upper))

        # Check for LAMBDA (highest priority)
        if "LAMBDA" in functions:
            return FormulaCategory.LAMBDA

        # Check for dynamic array functions
        if functions & self.DYNAMIC_ARRAY_FUNCTIONS:
            return FormulaCategory.DYNAMIC_ARRAY

        # Check for array formula syntax
        if formula.startswith("{=") or "CSE" in str(functions):
            return FormulaCategory.ARRAY_LEGACY

        # Check other categories
        if functions & self.LOOKUP_FUNCTIONS:
            return FormulaCategory.LOOKUP

        if functions & self.VOLATILE_FUNCTIONS:
            return FormulaCategory.VOLATILE

        if functions & self.AGGREGATE_FUNCTIONS:
            return FormulaCategory.AGGREGATE

        if functions & self.ERROR_HANDLING_FUNCTIONS:
            return FormulaCategory.ERROR_HANDLING

        if functions & self.TEXT_FUNCTIONS:
            return FormulaCategory.TEXT

        if functions & self.DATE_TIME_FUNCTIONS:
            return FormulaCategory.DATE_TIME

        if functions & self.LOGICAL_FUNCTIONS:
            return FormulaCategory.LOGICAL

        if functions & self.FINANCIAL_FUNCTIONS:
            return FormulaCategory.FINANCIAL

        if functions & self.MATH_FUNCTIONS:
            return FormulaCategory.MATH

        # Check for external references
        if "[" in formula and "]" in formula:
            return FormulaCategory.EXTERNAL

        return FormulaCategory.SIMPLE

    def _is_array_formula(self, cell: Cell) -> bool:
        """Check if cell contains an array formula."""
        try:
            # Check for CSE array formula
            if hasattr(cell, "value") and isinstance(cell.value, str):
                if cell.value.startswith("{=") and cell.value.endswith("}"):
                    return True
            # Check for array formula attribute
            if hasattr(cell, "array_formula") and cell.array_formula:
                return True
        except Exception:
            pass
        return False

    def _extract_external_refs(self, formula: str) -> list[str]:
        """Extract external workbook references from formula."""
        refs = []

        # Pattern for external references: [WorkbookName.xlsx]SheetName!Range
        # or 'C:\path\[file.xlsx]Sheet'!A1
        patterns = [
            r"\[([^\]]+\.xlsx?)\]",  # [filename.xlsx]
            r"'[^']*\[([^\]]+\.xlsx?)\][^']*'",  # 'path\[file.xlsx]sheet'
        ]

        for pattern in patterns:
            matches = re.findall(pattern, formula, re.IGNORECASE)
            refs.extend(matches)

        return list(set(refs))
