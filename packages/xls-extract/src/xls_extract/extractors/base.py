"""Base extractor protocol and utilities."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, TypeVar
from zipfile import ZipFile

from openpyxl import Workbook

T = TypeVar("T")


class BaseExtractor(ABC):
    """Base class for all extractors."""

    name: str = "base"

    def __init__(self, workbook: Workbook, file_path: Path):
        """Initialize extractor.

        Args:
            workbook: The openpyxl Workbook object
            file_path: Path to the xlsx file (for direct XML access)
        """
        self.workbook = workbook
        self.file_path = file_path

    @abstractmethod
    def extract(self) -> Any:
        """Extract data from the workbook.

        Returns:
            Extracted data (type depends on extractor)
        """
        pass

    def read_xml_from_xlsx(self, internal_path: str) -> bytes | None:
        """Read an XML file from inside the xlsx archive.

        Args:
            internal_path: Path inside the xlsx (e.g., 'xl/workbook.xml')

        Returns:
            XML content as bytes, or None if not found
        """
        try:
            with ZipFile(self.file_path, "r") as zf:
                if internal_path in zf.namelist():
                    return zf.read(internal_path)
        except Exception:
            pass
        return None

    def list_xlsx_contents(self) -> list[str]:
        """List all files inside the xlsx archive.

        Returns:
            List of internal file paths
        """
        try:
            with ZipFile(self.file_path, "r") as zf:
                return zf.namelist()
        except Exception:
            return []

    def read_file_from_xlsx(self, internal_path: str) -> bytes | None:
        """Read any file from inside the xlsx archive.

        Args:
            internal_path: Path inside the xlsx

        Returns:
            File content as bytes, or None if not found
        """
        return self.read_xml_from_xlsx(internal_path)
