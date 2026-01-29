"""Named range and LAMBDA function extractor."""

from __future__ import annotations

import re

from ..models import NamedRangeInfo
from .base import BaseExtractor


class NamedRangeExtractor(BaseExtractor):
    """Extracts named ranges and LAMBDA function definitions."""

    name = "named_ranges"

    def extract(self) -> list[NamedRangeInfo]:
        """Extract all named ranges and LAMBDA definitions.

        Returns:
            List of NamedRangeInfo objects
        """
        named_ranges = []

        try:
            # Try iterating over defined_names directly (openpyxl 3.1+)
            for defined_name in self.workbook.defined_names:
                info = self._create_named_range_info(defined_name)
                if info:
                    named_ranges.append(info)
        except TypeError:
            # Fallback for older API
            try:
                for name in self.workbook.defined_names.definedName:
                    info = self._create_named_range_info(name)
                    if info:
                        named_ranges.append(info)
            except Exception:
                pass
        except Exception:
            pass

        return named_ranges

    def _create_named_range_info(self, defined_name) -> NamedRangeInfo | None:
        """Create NamedRangeInfo from a defined name object."""
        try:
            name = defined_name.name
            value = defined_name.value or ""

            # Determine scope
            scope = None
            if defined_name.localSheetId is not None:
                try:
                    scope = self.workbook.sheetnames[defined_name.localSheetId]
                except (IndexError, TypeError):
                    pass

            # Check if this is a LAMBDA definition
            is_lambda = self._is_lambda_definition(value)

            # Clean up the value (translate prefixes)
            value_clean = self._clean_value(value)

            # Check if hidden
            hidden = getattr(defined_name, "hidden", False)

            # Get comment if available
            comment = getattr(defined_name, "comment", None)

            return NamedRangeInfo(
                name=name,
                value=value_clean,
                scope=scope,
                is_lambda=is_lambda,
                comment=comment,
                hidden=hidden,
            )
        except Exception:
            return None

    def _is_lambda_definition(self, value: str) -> bool:
        """Check if a named range value is a LAMBDA function."""
        value_upper = value.upper()
        # Check for LAMBDA keyword (with or without _xlfn prefix)
        return "LAMBDA(" in value_upper or "_XLPM.LAMBDA(" in value_upper or "_XLFN.LAMBDA(" in value_upper

    def _clean_value(self, value: str) -> str:
        """Clean up named range value by translating prefixes."""
        cleaned = value

        # Remove _xlfn. and _xlpm. prefixes
        cleaned = re.sub(r"_xlfn\.", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"_xlpm\.", "", cleaned, flags=re.IGNORECASE)

        return cleaned
