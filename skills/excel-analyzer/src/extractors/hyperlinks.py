"""Hyperlink extractor."""

from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet

from ..models import CellReference, HyperlinkInfo
from .base import BaseExtractor


class HyperlinkExtractor(BaseExtractor):
    """Extracts hyperlinks from all sheets."""

    name = "hyperlinks"

    def extract(self) -> list[HyperlinkInfo]:
        """Extract all hyperlinks.

        Returns:
            List of HyperlinkInfo objects
        """
        hyperlinks = []

        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            if not isinstance(sheet, Worksheet):
                continue

            sheet_hyperlinks = self._extract_sheet_hyperlinks(sheet, sheet_name)
            hyperlinks.extend(sheet_hyperlinks)

        return hyperlinks

    def _extract_sheet_hyperlinks(self, sheet: Worksheet, sheet_name: str) -> list[HyperlinkInfo]:
        """Extract hyperlinks from a sheet."""
        hyperlinks = []

        try:
            # Get hyperlinks from sheet.hyperlinks collection
            for hyperlink in sheet.hyperlinks:
                ref = hyperlink.ref if hasattr(hyperlink, "ref") else ""
                target = hyperlink.target or ""

                # Determine if external
                is_external = self._is_external_link(target)

                # Get display text from the cell
                display_text = None
                try:
                    cell = sheet[ref] if ref else None
                    if cell:
                        display_text = str(cell.value) if cell.value else None
                except Exception:
                    pass

                hyperlinks.append(HyperlinkInfo(
                    location=CellReference(
                        sheet=sheet_name,
                        cell=ref,
                        row=self._get_row_from_ref(ref),
                        col=self._get_col_from_ref(ref),
                    ),
                    target=target,
                    display_text=display_text,
                    tooltip=hyperlink.tooltip if hasattr(hyperlink, "tooltip") else None,
                    is_external=is_external,
                ))
        except Exception:
            pass

        # Also check cells directly for hyperlink property
        try:
            for row in sheet.iter_rows():
                for cell in row:
                    if hasattr(cell, "hyperlink") and cell.hyperlink:
                        # Check if we already have this one
                        existing = any(
                            h.location.sheet == sheet_name and h.location.cell == cell.coordinate
                            for h in hyperlinks
                        )
                        if not existing:
                            target = cell.hyperlink.target or ""
                            hyperlinks.append(HyperlinkInfo(
                                location=CellReference(
                                    sheet=sheet_name,
                                    cell=cell.coordinate,
                                    row=cell.row,
                                    col=cell.column,
                                ),
                                target=target,
                                display_text=str(cell.value) if cell.value else None,
                                tooltip=getattr(cell.hyperlink, "tooltip", None),
                                is_external=self._is_external_link(target),
                            ))
        except Exception:
            pass

        return hyperlinks

    def _is_external_link(self, target: str) -> bool:
        """Determine if a hyperlink target is external."""
        if not target:
            return False

        external_prefixes = ["http://", "https://", "ftp://", "mailto:", "file://"]
        return any(target.lower().startswith(prefix) for prefix in external_prefixes)

    def _get_row_from_ref(self, ref: str) -> int:
        """Extract row number from cell reference."""
        import re
        match = re.search(r"(\d+)", ref)
        return int(match.group(1)) if match else 0

    def _get_col_from_ref(self, ref: str) -> int:
        """Extract column number from cell reference."""
        import re
        match = re.search(r"([A-Z]+)", ref.upper())
        if not match:
            return 0

        col_str = match.group(1)
        col_num = 0
        for char in col_str:
            col_num = col_num * 26 + (ord(char) - ord("A") + 1)
        return col_num
