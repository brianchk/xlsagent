"""AutoFilter extractor."""

from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet

from ..models import AutoFilterInfo
from .base import BaseExtractor


class FilterExtractor(BaseExtractor):
    """Extracts AutoFilter settings from all sheets."""

    name = "filters"

    def extract(self) -> list[AutoFilterInfo]:
        """Extract all AutoFilter settings.

        Returns:
            List of AutoFilterInfo objects
        """
        filters = []

        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            if not isinstance(sheet, Worksheet):
                continue

            filter_info = self._extract_sheet_filter(sheet, sheet_name)
            if filter_info:
                filters.append(filter_info)

        return filters

    def _extract_sheet_filter(self, sheet: Worksheet, sheet_name: str) -> AutoFilterInfo | None:
        """Extract AutoFilter from a sheet."""
        try:
            if not sheet.auto_filter or not sheet.auto_filter.ref:
                return None

            filter_ref = sheet.auto_filter.ref

            # Extract column filter details
            column_filters = {}
            try:
                if hasattr(sheet.auto_filter, "filterColumn") and sheet.auto_filter.filterColumn:
                    for fc in sheet.auto_filter.filterColumn:
                        col_id = fc.colId if hasattr(fc, "colId") else 0
                        filter_details = self._extract_filter_details(fc)
                        if filter_details:
                            column_filters[col_id] = filter_details
            except Exception:
                pass

            return AutoFilterInfo(
                sheet=sheet_name,
                range=filter_ref,
                column_filters=column_filters,
            )
        except Exception:
            return None

    def _extract_filter_details(self, filter_column) -> dict | None:
        """Extract details from a filter column."""
        details = {}

        try:
            # Check for filters (specific values)
            if hasattr(filter_column, "filters") and filter_column.filters:
                filters = filter_column.filters
                if hasattr(filters, "filter"):
                    details["type"] = "values"
                    details["values"] = [f.val for f in filters.filter if hasattr(f, "val")]
                if hasattr(filters, "blank") and filters.blank:
                    details["include_blank"] = True

            # Check for custom filters
            if hasattr(filter_column, "customFilters") and filter_column.customFilters:
                custom = filter_column.customFilters
                details["type"] = "custom"
                details["and_or"] = "and" if getattr(custom, "and_", False) else "or"
                custom_filters = []
                if hasattr(custom, "customFilter"):
                    for cf in custom.customFilter:
                        custom_filters.append({
                            "operator": getattr(cf, "operator", None),
                            "val": getattr(cf, "val", None),
                        })
                details["filters"] = custom_filters

            # Check for top10 filter
            if hasattr(filter_column, "top10") and filter_column.top10:
                t10 = filter_column.top10
                details["type"] = "top10"
                details["top"] = getattr(t10, "top", True)
                details["percent"] = getattr(t10, "percent", False)
                details["val"] = getattr(t10, "val", 10)

            # Check for dynamic filter
            if hasattr(filter_column, "dynamicFilter") and filter_column.dynamicFilter:
                dyn = filter_column.dynamicFilter
                details["type"] = "dynamic"
                details["dynamic_type"] = getattr(dyn, "type", None)

            # Check for color filter
            if hasattr(filter_column, "colorFilter") and filter_column.colorFilter:
                color = filter_column.colorFilter
                details["type"] = "color"
                details["cell_color"] = getattr(color, "cellColor", None)
                details["dxf_id"] = getattr(color, "dxfId", None)

        except Exception:
            pass

        return details if details else None
