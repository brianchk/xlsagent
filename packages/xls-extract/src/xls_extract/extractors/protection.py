"""Protection settings extractor."""

from __future__ import annotations

from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from ..models import WorkbookProtectionInfo, SheetProtectionInfo
from .base import BaseExtractor


class ProtectionExtractor(BaseExtractor):
    """Extracts protection settings from workbook and sheets."""

    name = "protection"

    def extract(self) -> dict[str, Any]:
        """Extract protection settings.

        Returns:
            Dict with 'workbook' (WorkbookProtectionInfo) and 'sheets' (list of SheetProtectionInfo)
        """
        workbook_info = self._extract_workbook_protection()
        sheet_infos = self._extract_sheet_protection()

        return {
            "workbook": workbook_info,
            "sheets": sheet_infos,
        }

    def _extract_workbook_protection(self) -> WorkbookProtectionInfo | None:
        """Extract workbook-level protection settings."""
        try:
            security = self.workbook.security

            if security:
                return WorkbookProtectionInfo(
                    is_protected=True,
                    protect_structure=getattr(security, "lockStructure", False) or False,
                    protect_windows=getattr(security, "lockWindows", False) or False,
                )
        except Exception:
            pass

        return None

    def _extract_sheet_protection(self) -> list[SheetProtectionInfo]:
        """Extract sheet-level protection settings."""
        results = []

        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            if not isinstance(sheet, Worksheet):
                continue

            try:
                protection = sheet.protection

                if protection and protection.sheet:
                    # In openpyxl, False means allowed, True means blocked
                    info = SheetProtectionInfo(
                        sheet=sheet_name,
                        is_protected=True,
                        allow_select_locked=not getattr(protection, "selectLockedCells", False),
                        allow_select_unlocked=not getattr(protection, "selectUnlockedCells", False),
                        allow_format_cells=not getattr(protection, "formatCells", True),
                        allow_format_columns=not getattr(protection, "formatColumns", True),
                        allow_format_rows=not getattr(protection, "formatRows", True),
                        allow_insert_columns=not getattr(protection, "insertColumns", True),
                        allow_insert_rows=not getattr(protection, "insertRows", True),
                        allow_insert_hyperlinks=not getattr(protection, "insertHyperlinks", True),
                        allow_delete_columns=not getattr(protection, "deleteColumns", True),
                        allow_delete_rows=not getattr(protection, "deleteRows", True),
                        allow_sort=not getattr(protection, "sort", True),
                        allow_filter=not getattr(protection, "autoFilter", True),
                        allow_pivot_tables=not getattr(protection, "pivotTables", True),
                    )
                    results.append(info)

            except Exception:
                pass

        return results
