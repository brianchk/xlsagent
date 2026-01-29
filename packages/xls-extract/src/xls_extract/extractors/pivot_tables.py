"""Pivot table extractor."""

from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet

from ..models import PivotTableInfo
from .base import BaseExtractor


class PivotTableExtractor(BaseExtractor):
    """Extracts pivot table definitions from all sheets."""

    name = "pivot_tables"

    def extract(self) -> list[PivotTableInfo]:
        """Extract all pivot table definitions.

        Returns:
            List of PivotTableInfo objects
        """
        pivots = []

        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            if not isinstance(sheet, Worksheet):
                continue

            sheet_pivots = self._extract_sheet_pivots(sheet, sheet_name)
            pivots.extend(sheet_pivots)

        return pivots

    def _extract_sheet_pivots(self, sheet: Worksheet, sheet_name: str) -> list[PivotTableInfo]:
        """Extract pivot tables from a sheet."""
        pivots = []

        try:
            for pivot in sheet._pivots:
                info = self._create_pivot_info(pivot, sheet_name)
                if info:
                    pivots.append(info)
        except Exception:
            pass

        return pivots

    def _create_pivot_info(self, pivot, sheet_name: str) -> PivotTableInfo | None:
        """Create PivotTableInfo from a pivot table object."""
        try:
            name = getattr(pivot, "name", None) or "PivotTable"
            location = ""

            # Try to get location
            if hasattr(pivot, "location") and pivot.location:
                location = str(pivot.location.ref) if hasattr(pivot.location, "ref") else str(pivot.location)

            # Try to get source range from cache
            source_range = None
            cache_id = None
            if hasattr(pivot, "cacheId"):
                cache_id = pivot.cacheId

            # Extract field information (limited in openpyxl read-only mode)
            row_fields = []
            col_fields = []
            data_fields = []
            filter_fields = []

            try:
                if hasattr(pivot, "rowFields") and pivot.rowFields:
                    for field in pivot.rowFields.field:
                        if hasattr(field, "x") and field.x is not None:
                            row_fields.append(f"Field {field.x}")

                if hasattr(pivot, "colFields") and pivot.colFields:
                    for field in pivot.colFields.field:
                        if hasattr(field, "x") and field.x is not None:
                            col_fields.append(f"Field {field.x}")

                if hasattr(pivot, "dataFields") and pivot.dataFields:
                    for field in pivot.dataFields.dataField:
                        name = getattr(field, "name", None) or f"Field {getattr(field, 'fld', '?')}"
                        data_fields.append(name)

                if hasattr(pivot, "pageFields") and pivot.pageFields:
                    for field in pivot.pageFields.pageField:
                        if hasattr(field, "fld") and field.fld is not None:
                            filter_fields.append(f"Field {field.fld}")
            except Exception:
                pass

            return PivotTableInfo(
                name=name,
                sheet=sheet_name,
                location=location,
                source_range=source_range,
                row_fields=row_fields,
                column_fields=col_fields,
                data_fields=data_fields,
                filter_fields=filter_fields,
                cache_id=cache_id,
            )
        except Exception:
            return None
