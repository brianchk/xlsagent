"""Data connections extractor."""

from __future__ import annotations

import re

from lxml import etree

from ..models import CellReference, DataConnectionInfo, ExternalRefInfo
from .base import BaseExtractor


class ConnectionExtractor(BaseExtractor):
    """Extracts data connections and external references."""

    name = "connections"

    NAMESPACES = {
        "": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }

    def extract(self) -> tuple[list[DataConnectionInfo], list[ExternalRefInfo]]:
        """Extract data connections and external references.

        Returns:
            Tuple of (connections, external_refs)
        """
        connections = self._extract_connections()

        # Build map of connection ID -> pivot cache IDs that use it
        conn_to_pivots = self._map_connections_to_pivots()

        # Attach pivot cache info to connections
        for conn in connections:
            if conn.name in conn_to_pivots:
                conn.used_by_pivot_caches = conn_to_pivots[conn.name]

        external_refs = self._extract_external_refs()

        return connections, external_refs

    def _extract_connections(self) -> list[DataConnectionInfo]:
        """Extract data connections from xl/connections.xml."""
        connections = []

        content = self.read_xml_from_xlsx("xl/connections.xml")
        if not content:
            return connections

        try:
            root = etree.fromstring(content)

            # Find all connection elements
            for conn in root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}connection"):
                info = self._parse_connection(conn)
                if info:
                    connections.append(info)

        except Exception:
            pass

        return connections

    def _parse_connection(self, conn) -> DataConnectionInfo | None:
        """Parse a connection element."""
        try:
            name = conn.get("name", "Unknown")
            conn_id = conn.get("id")
            conn_type = self._determine_connection_type(conn)

            connection_string = None
            command_text = None
            command_type = None
            description = conn.get("description")

            # Extract ODBC/OLEDB properties
            dbPr = conn.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dbPr")
            if dbPr is not None:
                connection_string = dbPr.get("connection")
                command_text = dbPr.get("command")
                command_type_attr = dbPr.get("commandType")
                if command_type_attr:
                    command_type = {
                        "1": "SQL",
                        "2": "Table",
                        "3": "Default",
                        "4": "DAX",
                        "5": "Cube",
                    }.get(command_type_attr, command_type_attr)

            # Extract web query properties
            webPr = conn.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}webPr")
            if webPr is not None:
                conn_type = "Web Query"
                connection_string = webPr.get("url")

            # Extract text file properties
            textPr = conn.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}textPr")
            if textPr is not None:
                conn_type = "Text File"
                connection_string = textPr.get("sourceFile")

            # Extract OLAP properties (Power Pivot / Analysis Services)
            olapPr = conn.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}olapPr")
            if olapPr is not None:
                conn_type = "OLAP/Power Pivot"

            # Detect DAX query in command text
            is_dax = False
            dax_query = None
            if command_text:
                # Clean up XML escape sequences
                cleaned_text = command_text.replace("_x000d_", "").replace("_x000a_", "\n")

                # Check if it looks like DAX
                dax_keywords = ["EVALUATE", "SUMMARIZE", "CALCULATE", "FILTER", "ALL", "VALUES", "RELATED"]
                command_upper = cleaned_text.upper()
                if any(kw in command_upper for kw in dax_keywords):
                    is_dax = True
                    dax_query = cleaned_text
                    if command_type is None:
                        command_type = "DAX"
                else:
                    # Still clean up the command text for display
                    command_text = cleaned_text

            return DataConnectionInfo(
                name=name,
                connection_type=conn_type,
                connection_string=connection_string,
                command_text=command_text,
                command_type=command_type,
                description=description,
                is_dax=is_dax,
                dax_query=dax_query,
                connection_id=conn_id,
            )
        except Exception:
            return None

    def _determine_connection_type(self, conn) -> str:
        """Determine the type of data connection."""
        conn_type_attr = conn.get("type")

        if conn_type_attr:
            type_map = {
                "1": "ODBC",
                "2": "DAO",
                "3": "File",
                "4": "Web Query",
                "5": "OLEDB",
                "6": "Text",
                "7": "ADO",
                "8": "DSP",
            }
            return type_map.get(conn_type_attr, f"Type {conn_type_attr}")

        return "Unknown"

    def _map_connections_to_pivots(self) -> dict[str, list[str]]:
        """Map connection names to pivot cache definitions that use them."""
        conn_to_pivots = {}

        contents = self.list_xlsx_contents()

        for item in contents:
            if item.startswith("xl/pivotCache/pivotCacheDefinition") and item.endswith(".xml"):
                content = self.read_xml_from_xlsx(item)
                if content:
                    try:
                        root = etree.fromstring(content)

                        # Get cache ID from filename
                        cache_id = item.split("pivotCacheDefinition")[1].replace(".xml", "")

                        # Check for connection reference
                        cache_source = root.find(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cacheSource")
                        if cache_source is not None:
                            conn_id = cache_source.get("connectionId")
                            if conn_id:
                                # Try to find connection name by ID
                                conn_name = self._get_connection_name_by_id(conn_id)
                                if conn_name:
                                    if conn_name not in conn_to_pivots:
                                        conn_to_pivots[conn_name] = []
                                    conn_to_pivots[conn_name].append(f"PivotCache{cache_id}")
                    except Exception:
                        pass

        return conn_to_pivots

    def _get_connection_name_by_id(self, conn_id: str) -> str | None:
        """Get connection name by its ID."""
        content = self.read_xml_from_xlsx("xl/connections.xml")
        if not content:
            return None

        try:
            root = etree.fromstring(content)
            for conn in root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}connection"):
                if conn.get("id") == conn_id:
                    return conn.get("name")
        except Exception:
            pass

        return None

    def _extract_external_refs(self) -> list[ExternalRefInfo]:
        """Extract external workbook references from formulas."""
        external_refs = []
        seen_refs = set()

        # Scan all sheets for external references in formulas
        for sheet_name in self.workbook.sheetnames:
            try:
                sheet = self.workbook[sheet_name]

                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                            refs = self._find_external_refs_in_formula(cell.value, sheet_name, cell.coordinate)
                            for ref in refs:
                                ref_key = (ref.target_workbook, ref.target_sheet)
                                if ref_key not in seen_refs:
                                    seen_refs.add(ref_key)
                                    external_refs.append(ref)
            except Exception:
                continue

        # Also check xl/externalLinks for linked workbooks
        external_refs.extend(self._extract_from_external_links())

        return external_refs

    def _find_external_refs_in_formula(
        self, formula: str, sheet_name: str, cell_coord: str
    ) -> list[ExternalRefInfo]:
        """Find external workbook references in a formula."""
        refs = []

        # Pattern: [WorkbookName.xlsx]SheetName!Range
        pattern = r"\[([^\]]+\.xlsx?)\](?:'?([^'!]+)'?)?!?([A-Z]+\d+(?::[A-Z]+\d+)?)?"

        for match in re.finditer(pattern, formula, re.IGNORECASE):
            workbook_name = match.group(1)
            target_sheet = match.group(2)
            target_range = match.group(3)

            refs.append(ExternalRefInfo(
                source_cell=CellReference(
                    sheet=sheet_name,
                    cell=cell_coord,
                    row=0,  # Would need to parse
                    col=0,
                ),
                target_workbook=workbook_name,
                target_sheet=target_sheet,
                target_range=target_range,
                is_broken=False,  # Would need to verify
            ))

        return refs

    def _extract_from_external_links(self) -> list[ExternalRefInfo]:
        """Extract external references from xl/externalLinks/."""
        refs = []

        contents = self.list_xlsx_contents()

        for item in contents:
            if item.startswith("xl/externalLinks/externalLink") and item.endswith(".xml"):
                content = self.read_xml_from_xlsx(item)
                if content:
                    try:
                        root = etree.fromstring(content)

                        # Find external book reference
                        ext_book = root.find(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}externalBook")
                        if ext_book is not None:
                            # The actual file path is in the relationships
                            rid = ext_book.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")

                            # Get filename from rels file
                            rels_path = item.replace("xl/externalLinks/", "xl/externalLinks/_rels/") + ".rels"
                            rels_content = self.read_xml_from_xlsx(rels_path)

                            if rels_content:
                                rels_root = etree.fromstring(rels_content)
                                for rel in rels_root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                                    if rel.get("Id") == rid:
                                        target = rel.get("Target", "")
                                        # Extract filename from path
                                        filename = target.split("/")[-1] if "/" in target else target

                                        refs.append(ExternalRefInfo(
                                            source_cell=CellReference(
                                                sheet="",
                                                cell="",
                                                row=0,
                                                col=0,
                                            ),
                                            target_workbook=filename,
                                            is_broken="file:///" not in target.lower(),
                                        ))

                    except Exception:
                        pass

        return refs
