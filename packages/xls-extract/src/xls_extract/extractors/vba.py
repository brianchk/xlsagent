"""VBA macro extractor using oletools."""

from __future__ import annotations

import re
from pathlib import Path

from ..models import VBAModuleInfo
from .base import BaseExtractor


class VBAExtractor(BaseExtractor):
    """Extracts VBA code from macro-enabled workbooks using oletools."""

    name = "vba"

    def extract(self) -> list[VBAModuleInfo]:
        """Extract VBA modules from the workbook.

        Returns:
            List of VBAModuleInfo objects
        """
        # Only process macro-enabled files
        if not self._is_macro_enabled():
            return []

        try:
            from oletools.olevba import VBA_Parser
        except ImportError:
            # oletools not available
            return []

        modules = []

        try:
            vba_parser = VBA_Parser(str(self.file_path))

            if vba_parser.detect_vba_macros():
                for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                    if vba_code:
                        info = self._create_module_info(vba_filename, vba_code, stream_path)
                        if info:
                            modules.append(info)

            vba_parser.close()
        except Exception:
            pass

        return modules

    def _is_macro_enabled(self) -> bool:
        """Check if the file is macro-enabled."""
        suffix = self.file_path.suffix.lower()
        return suffix in (".xlsm", ".xlsb", ".xltm", ".xla", ".xlam")

    def _create_module_info(
        self, module_name: str, code: str, stream_path: str
    ) -> VBAModuleInfo | None:
        """Create VBAModuleInfo from extracted code."""
        try:
            # Determine module type from stream path and code
            module_type = self._determine_module_type(module_name, code, stream_path)

            # Extract procedure names
            procedures = self._extract_procedures(code)

            # Count lines (excluding empty lines)
            line_count = len([line for line in code.split("\n") if line.strip()])

            return VBAModuleInfo(
                name=module_name,
                module_type=module_type,
                code=code,
                line_count=line_count,
                procedures=procedures,
            )
        except Exception:
            return None

    def _determine_module_type(self, name: str, code: str, stream_path: str) -> str:
        """Determine the type of VBA module."""
        name_lower = name.lower()

        if "thisworkbook" in name_lower:
            return "ThisWorkbook"

        if name_lower.startswith("sheet") or "worksheet" in stream_path.lower():
            return "Sheet"

        # Check for class module indicators in code
        if code.strip().upper().startswith("VERSION 1.0 CLASS"):
            return "Class"

        # Check for common class module patterns
        class_patterns = [
            r"^\s*Private\s+Type\s+",
            r"^\s*Implements\s+",
            r"Property\s+(Get|Let|Set)\s+",
        ]
        for pattern in class_patterns:
            if re.search(pattern, code, re.MULTILINE | re.IGNORECASE):
                return "Class"

        return "Standard"

    def _extract_procedures(self, code: str) -> list[str]:
        """Extract procedure names from VBA code."""
        procedures = []

        # Pattern for Sub, Function, Property procedures
        patterns = [
            r"^\s*(Public|Private|Friend)?\s*(Sub|Function)\s+(\w+)",
            r"^\s*(Public|Private|Friend)?\s*Property\s+(Get|Let|Set)\s+(\w+)",
        ]

        for pattern in patterns:
            matches = re.findall(pattern, code, re.MULTILINE | re.IGNORECASE)
            for match in matches:
                # Extract procedure name (last captured group)
                proc_name = match[-1] if match else None
                if proc_name and proc_name not in procedures:
                    procedures.append(proc_name)

        return procedures

    def get_vba_project_name(self) -> str | None:
        """Get the VBA project name if available."""
        if not self._is_macro_enabled():
            return None

        try:
            from oletools.olevba import VBA_Parser

            vba_parser = VBA_Parser(str(self.file_path))
            # VBA project name is typically in the vba_project attribute
            name = None
            if hasattr(vba_parser, "vba_project_name"):
                name = vba_parser.vba_project_name
            vba_parser.close()
            return name
        except Exception:
            return None
