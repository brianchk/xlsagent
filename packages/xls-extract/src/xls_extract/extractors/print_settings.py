"""Print settings extractor."""

from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet

from ..models import PrintSettingsInfo
from .base import BaseExtractor


class PrintSettingsExtractor(BaseExtractor):
    """Extracts print settings from all sheets."""

    name = "print_settings"

    def extract(self) -> list[PrintSettingsInfo]:
        """Extract print settings from all sheets.

        Returns:
            List of PrintSettingsInfo objects
        """
        settings = []

        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            if not isinstance(sheet, Worksheet):
                continue

            sheet_settings = self._extract_sheet_settings(sheet, sheet_name)
            if sheet_settings:
                settings.append(sheet_settings)

        return settings

    def _extract_sheet_settings(self, sheet: Worksheet, sheet_name: str) -> PrintSettingsInfo | None:
        """Extract print settings from a sheet."""
        try:
            info = PrintSettingsInfo(sheet=sheet_name)

            # Get print area
            if sheet.print_area:
                info.print_area = sheet.print_area

            # Get print titles
            if sheet.print_title_rows:
                info.print_titles_rows = sheet.print_title_rows

            if sheet.print_title_cols:
                info.print_titles_cols = sheet.print_title_cols

            # Get page breaks
            if hasattr(sheet, "row_breaks") and sheet.row_breaks:
                info.page_breaks_row = [
                    brk.id for brk in sheet.row_breaks.brk if hasattr(brk, "id")
                ]

            if hasattr(sheet, "col_breaks") and sheet.col_breaks:
                info.page_breaks_col = [
                    brk.id for brk in sheet.col_breaks.brk if hasattr(brk, "id")
                ]

            # Get page setup
            if sheet.page_setup:
                setup = sheet.page_setup

                if hasattr(setup, "orientation"):
                    info.orientation = setup.orientation or "portrait"

                if hasattr(setup, "paperSize"):
                    info.paper_size = self._get_paper_size_name(setup.paperSize)

                if hasattr(setup, "fitToPage"):
                    info.fit_to_page = setup.fitToPage or False

                if hasattr(setup, "fitToWidth"):
                    info.fit_to_width = setup.fitToWidth

                if hasattr(setup, "fitToHeight"):
                    info.fit_to_height = setup.fitToHeight

            # Only return if there are meaningful settings
            has_settings = (
                info.print_area or
                info.print_titles_rows or
                info.print_titles_cols or
                info.page_breaks_row or
                info.page_breaks_col or
                info.fit_to_page
            )

            return info if has_settings else None

        except Exception:
            return None

    def _get_paper_size_name(self, paper_size: int | None) -> str | None:
        """Convert paper size code to name."""
        if paper_size is None:
            return None

        # Common paper sizes
        paper_sizes = {
            1: "Letter (8.5 x 11 in)",
            2: "Letter Small (8.5 x 11 in)",
            3: "Tabloid (11 x 17 in)",
            4: "Ledger (17 x 11 in)",
            5: "Legal (8.5 x 14 in)",
            6: "Statement (5.5 x 8.5 in)",
            7: "Executive (7.25 x 10.5 in)",
            8: "A3 (297 x 420 mm)",
            9: "A4 (210 x 297 mm)",
            10: "A4 Small (210 x 297 mm)",
            11: "A5 (148 x 210 mm)",
        }

        return paper_sizes.get(paper_size, f"Custom ({paper_size})")
