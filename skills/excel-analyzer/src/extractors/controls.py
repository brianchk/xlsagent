"""Form controls and shapes extractor."""

from __future__ import annotations

import re
from lxml import etree

from ..models import ControlInfo
from .base import BaseExtractor


class ControlExtractor(BaseExtractor):
    """Extracts form controls, buttons, and shapes from the workbook."""

    name = "controls"

    # Namespaces used in drawings/controls
    NAMESPACES = {
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "v": "urn:schemas-microsoft-com:vml",
        "o": "urn:schemas-microsoft-com:office:office",
        "x": "urn:schemas-microsoft-com:office:excel",
    }

    def extract(self) -> list[ControlInfo]:
        """Extract all form controls and shapes.

        Returns:
            List of ControlInfo objects
        """
        controls = []

        # Get all files in the xlsx
        contents = self.list_xlsx_contents()

        # Process drawings (shapes, charts, images)
        for item in contents:
            if item.startswith("xl/drawings/drawing") and item.endswith(".xml"):
                sheet_name = self._get_sheet_for_drawing(item, contents)
                drawing_controls = self._parse_drawing(item, sheet_name)
                controls.extend(drawing_controls)

        # Process VML drawings (form controls, comments)
        for item in contents:
            if "vmlDrawing" in item and item.endswith(".vml"):
                sheet_name = self._get_sheet_for_vml(item, contents)
                vml_controls = self._parse_vml_drawing(item, sheet_name)
                controls.extend(vml_controls)

        # Process control properties
        for item in contents:
            if item.startswith("xl/ctrlProps/") and item.endswith(".xml"):
                control_info = self._parse_control_props(item)
                if control_info:
                    # Try to match with existing controls or add new
                    self._merge_control_props(controls, control_info)

        return controls

    def _get_sheet_for_drawing(self, drawing_path: str, contents: list[str]) -> str:
        """Determine which sheet a drawing belongs to."""
        # Parse drawing number from path
        match = re.search(r"drawing(\d+)\.xml", drawing_path)
        if not match:
            return "Unknown"

        drawing_num = match.group(1)

        # Look for relationship file that references this drawing
        for item in contents:
            if item.startswith("xl/worksheets/_rels/sheet") and item.endswith(".xml.rels"):
                rels_content = self.read_xml_from_xlsx(item)
                if rels_content:
                    try:
                        if f"drawing{drawing_num}.xml" in rels_content.decode("utf-8"):
                            # Extract sheet number
                            sheet_match = re.search(r"sheet(\d+)\.xml\.rels", item)
                            if sheet_match:
                                sheet_idx = int(sheet_match.group(1)) - 1
                                if 0 <= sheet_idx < len(self.workbook.sheetnames):
                                    return self.workbook.sheetnames[sheet_idx]
                    except Exception:
                        pass

        return "Unknown"

    def _get_sheet_for_vml(self, vml_path: str, contents: list[str]) -> str:
        """Determine which sheet a VML drawing belongs to."""
        # Similar logic to _get_sheet_for_drawing
        match = re.search(r"vmlDrawing(\d+)\.vml", vml_path)
        if not match:
            return "Unknown"

        try:
            # VML drawings are often numbered to match sheets
            vml_num = int(match.group(1))
            if 0 < vml_num <= len(self.workbook.sheetnames):
                return self.workbook.sheetnames[vml_num - 1]
        except Exception:
            pass

        return "Unknown"

    def _parse_drawing(self, drawing_path: str, sheet_name: str) -> list[ControlInfo]:
        """Parse a drawing XML file to extract shapes."""
        controls = []

        content = self.read_xml_from_xlsx(drawing_path)
        if not content:
            return controls

        try:
            root = etree.fromstring(content)

            # Find all anchored shapes
            for anchor in root.findall(".//xdr:twoCellAnchor", self.NAMESPACES):
                control = self._parse_shape_anchor(anchor, sheet_name)
                if control:
                    controls.append(control)

            for anchor in root.findall(".//xdr:oneCellAnchor", self.NAMESPACES):
                control = self._parse_shape_anchor(anchor, sheet_name)
                if control:
                    controls.append(control)

        except Exception:
            pass

        return controls

    def _parse_shape_anchor(self, anchor, sheet_name: str) -> ControlInfo | None:
        """Parse a shape anchor element."""
        try:
            # Try to find shape element
            sp = anchor.find(".//xdr:sp", self.NAMESPACES)
            if sp is not None:
                return self._parse_sp_element(sp, sheet_name, anchor)

            # Try to find picture
            pic = anchor.find(".//xdr:pic", self.NAMESPACES)
            if pic is not None:
                return self._parse_pic_element(pic, sheet_name, anchor)

            # Try to find chart
            graphicFrame = anchor.find(".//xdr:graphicFrame", self.NAMESPACES)
            if graphicFrame is not None:
                return None  # Charts are handled separately

        except Exception:
            pass

        return None

    def _parse_sp_element(self, sp, sheet_name: str, anchor) -> ControlInfo | None:
        """Parse a shape (sp) element."""
        try:
            # Get shape properties
            nvSpPr = sp.find(".//xdr:nvSpPr", self.NAMESPACES)
            name = "Shape"
            if nvSpPr is not None:
                cNvPr = nvSpPr.find(".//xdr:cNvPr", self.NAMESPACES)
                if cNvPr is not None:
                    name = cNvPr.get("name", "Shape")

            # Get position
            position = self._get_anchor_position(anchor)

            # Get text content if any
            text = None
            txBody = sp.find(".//xdr:txBody", self.NAMESPACES)
            if txBody is not None:
                texts = []
                for t in txBody.findall(".//a:t", self.NAMESPACES):
                    if t.text:
                        texts.append(t.text)
                text = " ".join(texts) if texts else None

            return ControlInfo(
                name=name,
                sheet=sheet_name,
                control_type="Shape",
                position=position,
                text=text,
            )
        except Exception:
            return None

    def _parse_pic_element(self, pic, sheet_name: str, anchor) -> ControlInfo | None:
        """Parse a picture element."""
        try:
            nvPicPr = pic.find(".//xdr:nvPicPr", self.NAMESPACES)
            name = "Picture"
            if nvPicPr is not None:
                cNvPr = nvPicPr.find(".//xdr:cNvPr", self.NAMESPACES)
                if cNvPr is not None:
                    name = cNvPr.get("name", "Picture")

            position = self._get_anchor_position(anchor)

            return ControlInfo(
                name=name,
                sheet=sheet_name,
                control_type="Picture",
                position=position,
            )
        except Exception:
            return None

    def _parse_vml_drawing(self, vml_path: str, sheet_name: str) -> list[ControlInfo]:
        """Parse VML drawing to extract form controls."""
        controls = []

        content = self.read_xml_from_xlsx(vml_path)
        if not content:
            return controls

        try:
            # VML uses different XML structure
            # Need to handle namespace declarations in content
            content_str = content.decode("utf-8", errors="replace")

            # Find form controls using regex (VML parsing can be tricky)
            # Look for x:ClientData elements which define form controls
            control_pattern = r"<v:shape[^>]*>.*?<x:ClientData[^>]*ObjectType=\"([^\"]+)\"[^>]*>.*?</x:ClientData>.*?</v:shape>"
            matches = re.findall(control_pattern, content_str, re.DOTALL | re.IGNORECASE)

            for idx, object_type in enumerate(matches):
                # Map ObjectType to control type
                control_type = self._map_object_type(object_type)

                controls.append(ControlInfo(
                    name=f"{control_type} {idx + 1}",
                    sheet=sheet_name,
                    control_type=control_type,
                ))

            # Look for buttons with macros
            button_pattern = r"<v:shape[^>]*>.*?<x:ClientData[^>]*>.*?<x:FmlaMacro>([^<]+)</x:FmlaMacro>.*?</x:ClientData>.*?</v:shape>"
            button_matches = re.findall(button_pattern, content_str, re.DOTALL | re.IGNORECASE)

            for idx, macro in enumerate(button_matches):
                # Check if we already have this button
                macro_found = False
                for ctrl in controls:
                    if ctrl.macro is None and ctrl.control_type == "Button":
                        ctrl.macro = macro
                        macro_found = True
                        break

                if not macro_found:
                    controls.append(ControlInfo(
                        name=f"Button {len(controls) + 1}",
                        sheet=sheet_name,
                        control_type="Button",
                        macro=macro,
                    ))

        except Exception:
            pass

        return controls

    def _map_object_type(self, object_type: str) -> str:
        """Map VML ObjectType to friendly control type name."""
        type_map = {
            "Button": "Button",
            "Checkbox": "CheckBox",
            "CheckBox": "CheckBox",
            "Drop": "ComboBox",
            "Edit": "EditBox",
            "GBox": "GroupBox",
            "Label": "Label",
            "List": "ListBox",
            "Radio": "OptionButton",
            "Scroll": "ScrollBar",
            "Spin": "SpinButton",
            "Note": "Comment",
        }
        return type_map.get(object_type, object_type)

    def _parse_control_props(self, props_path: str) -> dict | None:
        """Parse control properties XML."""
        content = self.read_xml_from_xlsx(props_path)
        if not content:
            return None

        try:
            root = etree.fromstring(content)

            props = {}

            # Extract linked cell
            linked_cell = root.get("fmlaLink")
            if linked_cell:
                props["linked_cell"] = linked_cell

            # Extract macro
            macro = root.get("fmlaMacro")
            if macro:
                props["macro"] = macro

            # Extract other properties
            for attr in ["noThreeD", "checked", "dropStyle", "sel", "val", "min", "max", "inc", "page"]:
                value = root.get(attr)
                if value is not None:
                    props[attr] = value

            return props if props else None

        except Exception:
            return None

    def _merge_control_props(self, controls: list[ControlInfo], props: dict) -> None:
        """Merge control properties into existing controls."""
        # This is a simplified merge - in practice would need to match by ID
        if props.get("linked_cell"):
            for ctrl in controls:
                if ctrl.linked_cell is None:
                    ctrl.linked_cell = props["linked_cell"]
                    break

        if props.get("macro"):
            for ctrl in controls:
                if ctrl.macro is None:
                    ctrl.macro = props["macro"]
                    break

    def _get_anchor_position(self, anchor) -> str | None:
        """Extract position from anchor element."""
        try:
            from_elem = anchor.find(".//xdr:from", self.NAMESPACES)
            if from_elem is not None:
                col = from_elem.find("xdr:col", self.NAMESPACES)
                row = from_elem.find("xdr:row", self.NAMESPACES)
                if col is not None and row is not None:
                    return f"Col {col.text}, Row {row.text}"
        except Exception:
            pass
        return None
