"""Error cell extractor."""

from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet

from ..models import CellReference, ErrorCellInfo, ErrorType
from .base import BaseExtractor


class ErrorExtractor(BaseExtractor):
    """Extracts cells containing Excel errors."""

    name = "errors"

    # Map error strings to ErrorType enum
    ERROR_MAP = {
        "#REF!": ErrorType.REF,
        "#NAME?": ErrorType.NAME,
        "#VALUE!": ErrorType.VALUE,
        "#DIV/0!": ErrorType.DIV,
        "#NULL!": ErrorType.NULL,
        "#NUM!": ErrorType.NUM,
        "#N/A": ErrorType.NA,
        "#CALC!": ErrorType.CALC,
        "#SPILL!": ErrorType.SPILL,
        "#GETTING_DATA": ErrorType.GETTING_DATA,
    }

    def extract(self) -> list[ErrorCellInfo]:
        """Extract all cells containing errors.

        Returns:
            List of ErrorCellInfo objects
        """
        errors = []

        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            if not isinstance(sheet, Worksheet):
                continue

            sheet_errors = self._extract_sheet_errors(sheet, sheet_name)
            errors.extend(sheet_errors)

        return errors

    def _extract_sheet_errors(self, sheet: Worksheet, sheet_name: str) -> list[ErrorCellInfo]:
        """Extract error cells from a sheet."""
        errors = []

        try:
            for row in sheet.iter_rows():
                for cell in row:
                    error_info = self._check_cell_for_error(cell, sheet_name)
                    if error_info:
                        errors.append(error_info)
        except Exception:
            pass

        return errors

    def _check_cell_for_error(self, cell, sheet_name: str) -> ErrorCellInfo | None:
        """Check if a cell contains an error."""
        try:
            value = cell.value

            if value is None:
                return None

            # Check if the value is an error string
            value_str = str(value).upper()

            for error_str, error_type in self.ERROR_MAP.items():
                if value_str == error_str:
                    # Try to get the formula that caused the error
                    formula = None
                    try:
                        # In some cases, the formula is stored separately
                        if hasattr(cell, "data_type") and cell.data_type == "f":
                            formula = str(cell.value)
                    except Exception:
                        pass

                    return ErrorCellInfo(
                        location=CellReference(
                            sheet=sheet_name,
                            cell=cell.coordinate,
                            row=cell.row,
                            col=cell.column,
                        ),
                        error_type=error_type,
                        formula=formula,
                    )

            # Also check if cell data type indicates error
            if hasattr(cell, "data_type") and cell.data_type == "e":
                return ErrorCellInfo(
                    location=CellReference(
                        sheet=sheet_name,
                        cell=cell.coordinate,
                        row=cell.row,
                        col=cell.column,
                    ),
                    error_type=self._get_error_type(value_str),
                    formula=None,
                )

        except Exception:
            pass

        return None

    def _get_error_type(self, error_str: str) -> ErrorType:
        """Get ErrorType from error string."""
        error_str_upper = error_str.upper()

        for key, error_type in self.ERROR_MAP.items():
            if key in error_str_upper:
                return error_type

        # Default to VALUE error if unknown
        return ErrorType.VALUE
