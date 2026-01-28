"""Power Query M code extractor."""

from __future__ import annotations

import base64
import io
import re
from zipfile import ZipFile

from lxml import etree

from ..models import PowerQueryInfo
from .base import BaseExtractor


class PowerQueryExtractor(BaseExtractor):
    """Extracts Power Query M code from DataMashup embedded in xlsx."""

    name = "power_query"

    # Namespaces used in Power Query XML
    NAMESPACES = {
        "pq": "http://schemas.microsoft.com/DataMashup",
        "pkg": "http://schemas.microsoft.com/office/2006/xmlPackage",
    }

    def extract(self) -> list[PowerQueryInfo]:
        """Extract Power Query definitions.

        Returns:
            List of PowerQueryInfo objects
        """
        queries = []

        # Try to find and parse the DataMashup
        datamashup = self._find_datamashup()
        if datamashup:
            queries = self._parse_datamashup(datamashup)

        return queries

    def _find_datamashup(self) -> bytes | None:
        """Find the DataMashup content in customXml."""
        try:
            contents = self.list_xlsx_contents()

            # Look for customXml items
            for item in contents:
                if item.startswith("customXml/item") and item.endswith(".xml"):
                    xml_content = self.read_xml_from_xlsx(item)
                    if xml_content:
                        # Check if this is the DataMashup
                        if b"DataMashup" in xml_content or b"http://schemas.microsoft.com/DataMashup" in xml_content:
                            return xml_content

        except Exception:
            pass

        return None

    def _parse_datamashup(self, xml_content: bytes) -> list[PowerQueryInfo]:
        """Parse the DataMashup XML to extract queries."""
        queries = []

        try:
            root = etree.fromstring(xml_content)

            # The DataMashup contains a base64-encoded ZIP file with the actual M code
            # Find the Mashup content
            mashup_content = None

            # Try different paths based on format version
            for xpath in [
                ".//pq:Mashup",
                ".//*[local-name()='Mashup']",
                ".//pkg:part[@pkg:name='/package/formulas/Section1.m']",
            ]:
                try:
                    elements = root.xpath(xpath, namespaces=self.NAMESPACES)
                    if elements:
                        mashup_content = elements[0].text
                        break
                except Exception:
                    continue

            # If we have encoded content, decode and extract
            if mashup_content:
                queries = self._extract_from_mashup(mashup_content)
            else:
                # Try direct extraction from binary content
                queries = self._extract_from_binary(xml_content)

        except Exception:
            pass

        return queries

    def _extract_from_mashup(self, mashup_content: str) -> list[PowerQueryInfo]:
        """Extract queries from base64-encoded mashup content."""
        queries = []

        try:
            # Decode base64
            decoded = base64.b64decode(mashup_content)

            # The decoded content is a ZIP file
            with ZipFile(io.BytesIO(decoded), "r") as zf:
                # Look for M code files
                for name in zf.namelist():
                    if name.endswith(".m"):
                        m_code = zf.read(name).decode("utf-8", errors="replace")
                        # Parse the M code to extract individual queries
                        extracted = self._parse_m_code(m_code)
                        queries.extend(extracted)

                # Also look for metadata
                try:
                    if "[Content_Types].xml" in zf.namelist():
                        pass  # Could extract additional metadata
                except Exception:
                    pass

        except Exception:
            pass

        return queries

    def _extract_from_binary(self, content: bytes) -> list[PowerQueryInfo]:
        """Try to extract M code from binary/XML content directly."""
        queries = []

        try:
            # Convert to string and search for M code patterns
            text = content.decode("utf-8", errors="replace")

            # Look for shared query sections
            # Pattern: shared QueryName = let ... in ...;
            pattern = r"shared\s+(\w+)\s*=\s*(let\s+.*?in\s+\w+)\s*;"
            matches = re.findall(pattern, text, re.DOTALL | re.IGNORECASE)

            for name, formula in matches:
                queries.append(PowerQueryInfo(
                    name=name,
                    formula=formula.strip(),
                ))

            # If no matches, try simpler pattern for single query
            if not queries:
                let_pattern = r"(let\s+.*?in\s+\w+)"
                let_matches = re.findall(let_pattern, text, re.DOTALL | re.IGNORECASE)
                for idx, formula in enumerate(let_matches):
                    queries.append(PowerQueryInfo(
                        name=f"Query{idx + 1}",
                        formula=formula.strip(),
                    ))

        except Exception:
            pass

        return queries

    def _parse_m_code(self, m_code: str) -> list[PowerQueryInfo]:
        """Parse M code to extract individual query definitions."""
        queries = []

        try:
            # M code format: section Section1;
            #               shared QueryName = let ... in ...;
            #               shared QueryName2 = let ... in ...;

            # Remove section header
            m_code = re.sub(r"section\s+\w+\s*;", "", m_code, flags=re.IGNORECASE)

            # Split by 'shared' keyword (keeping the keyword)
            parts = re.split(r"(?=shared\s+)", m_code)

            for part in parts:
                part = part.strip()
                if not part or not part.lower().startswith("shared"):
                    continue

                # Parse: shared QueryName = formula;
                match = re.match(
                    r"shared\s+#?\"?(\w+)\"?\s*=\s*(.*?)\s*;?\s*$",
                    part,
                    re.DOTALL | re.IGNORECASE
                )
                if match:
                    name = match.group(1)
                    formula = match.group(2).strip()

                    # Clean up formula
                    formula = self._clean_formula(formula)

                    queries.append(PowerQueryInfo(
                        name=name,
                        formula=formula,
                    ))

        except Exception:
            pass

        return queries

    def _clean_formula(self, formula: str) -> str:
        """Clean up M code formula for display."""
        # Remove trailing semicolons
        formula = formula.rstrip(";").strip()

        # Fix common encoding issues
        formula = formula.replace("\r\n", "\n")
        formula = formula.replace("\r", "\n")

        return formula
