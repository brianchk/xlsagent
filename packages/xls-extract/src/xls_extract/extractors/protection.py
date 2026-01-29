"""Protection settings extractor."""

from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet

from ..models import ProtectionInfo
from .base import BaseExtractor


class ProtectionExtractor(BaseExtractor):
    """Extracts protection settings from workbook and sheets."""

    name = "protection"

    def extract(self) -> ProtectionInfo:
        """Extract protection settings.

        Returns:
            ProtectionInfo object
        """
        info = ProtectionInfo()

        # Check workbook-level protection
        self._extract_workbook_protection(info)

        # Check sheet-level protection
        self._extract_sheet_protection(info)

        return info

    def _extract_workbook_protection(self, info: ProtectionInfo) -> None:
        """Extract workbook-level protection settings."""
        try:
            security = self.workbook.security

            if security:
                info.workbook_protected = True

                # Check specific protection types
                if hasattr(security, "lockStructure"):
                    info.workbook_structure = security.lockStructure or False

                if hasattr(security, "lockWindows"):
                    info.workbook_windows = security.lockWindows or False

        except Exception:
            pass

    def _extract_sheet_protection(self, info: ProtectionInfo) -> None:
        """Extract sheet-level protection settings."""
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            if not isinstance(sheet, Worksheet):
                continue

            try:
                protection = sheet.protection

                if protection and protection.sheet:
                    sheet_details = {
                        "protected": True,
                        "password_protected": protection.password is not None,
                    }

                    # Extract specific protection flags
                    protection_flags = [
                        "selectLockedCells", "selectUnlockedCells",
                        "formatCells", "formatColumns", "formatRows",
                        "insertColumns", "insertRows", "insertHyperlinks",
                        "deleteColumns", "deleteRows",
                        "sort", "autoFilter", "pivotTables",
                        "objects", "scenarios",
                    ]

                    for flag in protection_flags:
                        if hasattr(protection, flag):
                            value = getattr(protection, flag)
                            # In openpyxl, False means allowed, True means protected
                            sheet_details[flag] = not value if value is not None else None

                    info.sheets[sheet_name] = sheet_details
                else:
                    info.sheets[sheet_name] = {"protected": False}

            except Exception:
                info.sheets[sheet_name] = {"protected": False}
