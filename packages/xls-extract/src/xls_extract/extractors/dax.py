"""DAX/Power Pivot detector."""

from __future__ import annotations

from lxml import etree

from .base import BaseExtractor


class DAXDetector(BaseExtractor):
    """Detects presence of DAX/Power Pivot in the workbook.

    Note: Full DAX extraction is not possible as it runs in-process in Excel.
    This detector identifies the presence of a Data Model and Power Pivot.
    """

    name = "dax"

    def extract(self) -> tuple[bool, str | None]:
        """Detect if workbook contains DAX/Power Pivot.

        Returns:
            Tuple of (has_dax: bool, detection_note: str | None)
        """
        has_dax = False
        notes = []

        # Check for Data Model presence
        if self._has_data_model():
            has_dax = True
            notes.append("Data Model detected")

        # Check for Power Pivot connections
        if self._has_power_pivot_connection():
            has_dax = True
            notes.append("Power Pivot connection detected")

        # Check for CUBE functions in formulas
        if self._has_cube_functions():
            has_dax = True
            notes.append("CUBE functions detected (likely using Data Model)")

        # Check for measures in pivot tables
        if self._has_measures():
            has_dax = True
            notes.append("Measures detected in pivot tables")

        note = "; ".join(notes) if notes else None

        if has_dax and not note:
            note = "Power Pivot/DAX presence detected but details cannot be fully extracted"

        return has_dax, note

    def _has_data_model(self) -> bool:
        """Check if workbook contains a Data Model."""
        contents = self.list_xlsx_contents()

        # Data Model is stored in xl/model/
        for item in contents:
            if item.startswith("xl/model/"):
                return True

        return False

    def _has_power_pivot_connection(self) -> bool:
        """Check for Power Pivot-specific connections."""
        content = self.read_xml_from_xlsx("xl/connections.xml")
        if not content:
            return False

        try:
            root = etree.fromstring(content)

            for conn in root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}connection"):
                # Check connection type and properties
                name = conn.get("name", "").lower()
                conn_type = conn.get("type", "")

                # Power Pivot connections often have specific names/types
                if "powerpivot" in name or "thisworkbookdatamodel" in name:
                    return True

                # Type 5 with specific model flag
                if conn_type == "5":
                    model_attr = conn.get("model")
                    if model_attr:
                        return True

        except Exception:
            pass

        return False

    def _has_cube_functions(self) -> bool:
        """Check if workbook uses CUBE functions (indicate Data Model usage)."""
        cube_functions = [
            "CUBEVALUE", "CUBEMEMBER", "CUBESET", "CUBERANKEDMEMBER",
            "CUBESETCOUNT", "CUBEMEMBERPROPERTY", "CUBEKPIMEMBER",
        ]

        for sheet_name in self.workbook.sheetnames:
            try:
                sheet = self.workbook[sheet_name]
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            value_upper = cell.value.upper()
                            for func in cube_functions:
                                if func in value_upper:
                                    return True
            except Exception:
                continue

        return False

    def _has_measures(self) -> bool:
        """Check for DAX measures in pivot cache definitions."""
        contents = self.list_xlsx_contents()

        for item in contents:
            if item.startswith("xl/pivotCache/pivotCacheDefinition") and item.endswith(".xml"):
                content = self.read_xml_from_xlsx(item)
                if content:
                    try:
                        # Check for measure-related elements
                        if b"<measure" in content.lower() or b"calculatedMember" in content:
                            return True
                    except Exception:
                        pass

        return False
