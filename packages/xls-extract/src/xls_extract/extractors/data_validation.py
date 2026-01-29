"""Data validation extractor."""

from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet

from ..models import DataValidationInfo
from .base import BaseExtractor


class DataValidationExtractor(BaseExtractor):
    """Extracts data validation rules from all sheets."""

    name = "data_validations"

    def extract(self) -> list[DataValidationInfo]:
        """Extract all data validation rules.

        Returns:
            List of DataValidationInfo objects
        """
        validations = []

        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            if not isinstance(sheet, Worksheet):
                continue

            sheet_validations = self._extract_sheet_validations(sheet, sheet_name)
            validations.extend(sheet_validations)

        return validations

    def _extract_sheet_validations(
        self, sheet: Worksheet, sheet_name: str
    ) -> list[DataValidationInfo]:
        """Extract data validations from a sheet."""
        validations = []

        try:
            for dv in sheet.data_validations.dataValidation:
                info = self._create_validation_info(dv, sheet_name)
                if info:
                    validations.append(info)
        except Exception:
            pass

        return validations

    def _create_validation_info(self, dv, sheet_name: str) -> DataValidationInfo | None:
        """Create DataValidationInfo from a data validation object."""
        try:
            # Get the range(s) this validation applies to
            ranges = []
            if hasattr(dv, "sqref") and dv.sqref:
                ranges = [str(r) for r in dv.sqref.ranges]

            range_str = ", ".join(ranges) if ranges else ""
            if sheet_name:
                range_str = f"'{sheet_name}'!{range_str}" if range_str else sheet_name

            return DataValidationInfo(
                range=range_str,
                type=dv.type or "any",
                operator=dv.operator,
                formula1=dv.formula1,
                formula2=dv.formula2,
                allow_blank=dv.allow_blank if dv.allow_blank is not None else True,
                show_dropdown=dv.showDropDown != True,  # Inverted in Excel
                show_input_message=dv.showInputMessage or False,
                input_title=dv.promptTitle,
                input_message=dv.prompt,
                show_error_message=dv.showErrorMessage or False,
                error_title=dv.errorTitle,
                error_message=dv.error,
                error_style=dv.errorStyle,
            )
        except Exception:
            return None
