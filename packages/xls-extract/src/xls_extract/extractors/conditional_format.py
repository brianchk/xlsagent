"""Conditional formatting extractor."""

from __future__ import annotations

from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    IconSetRule,
    Rule,
)
from openpyxl.worksheet.worksheet import Worksheet

from ..models import CFRuleType, ConditionalFormatInfo
from .base import BaseExtractor


class ConditionalFormatExtractor(BaseExtractor):
    """Extracts conditional formatting rules from all sheets."""

    name = "conditional_formatting"

    def extract(self) -> list[ConditionalFormatInfo]:
        """Extract all conditional formatting rules.

        Returns:
            List of ConditionalFormatInfo objects
        """
        rules = []

        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            if not isinstance(sheet, Worksheet):
                continue

            sheet_rules = self._extract_sheet_rules(sheet, sheet_name)
            rules.extend(sheet_rules)

        return rules

    def _extract_sheet_rules(self, sheet: Worksheet, sheet_name: str) -> list[ConditionalFormatInfo]:
        """Extract conditional formatting rules from a sheet."""
        rules = []

        try:
            # Iterate over conditional formatting ranges
            for cf_range in sheet.conditional_formatting:
                range_string = str(cf_range)
                # Get the rules for this range
                try:
                    cf_rules = sheet.conditional_formatting[cf_range]
                    if hasattr(cf_rules, 'cfRule'):
                        for rule in cf_rules.cfRule:
                            info = self._create_rule_info(range_string, rule)
                            if info:
                                rules.append(info)
                    elif hasattr(cf_rules, '__iter__'):
                        for rule in cf_rules:
                            info = self._create_rule_info(range_string, rule)
                            if info:
                                rules.append(info)
                except (TypeError, KeyError):
                    pass
        except Exception:
            pass

        # Fallback: try internal API
        if not rules:
            try:
                for range_string, cf_rules in sheet.conditional_formatting._cf_rules.items():
                    for rule in cf_rules:
                        info = self._create_rule_info(str(range_string), rule)
                        if info:
                            rules.append(info)
            except Exception:
                pass

        return rules

    def _create_rule_info(self, range_string: str, rule: Rule) -> ConditionalFormatInfo | None:
        """Create ConditionalFormatInfo from a rule object."""
        try:
            rule_type = self._determine_rule_type(rule)
            description = self._describe_rule(rule, rule_type)

            info = ConditionalFormatInfo(
                range=range_string,
                rule_type=rule_type,
                priority=getattr(rule, "priority", 0) or 0,
                formula=self._get_formula(rule),
                operator=getattr(rule, "operator", None),
                stop_if_true=getattr(rule, "stopIfTrue", False) or False,
                description=description,
            )

            # Extract values for certain rule types
            info.values = self._extract_values(rule)

            return info
        except Exception:
            return None

    def _determine_rule_type(self, rule: Rule) -> CFRuleType:
        """Determine the type of conditional formatting rule."""
        rule_type_attr = getattr(rule, "type", None)

        if isinstance(rule, ColorScaleRule) or rule_type_attr == "colorScale":
            return CFRuleType.COLOR_SCALE
        elif isinstance(rule, DataBarRule) or rule_type_attr == "dataBar":
            return CFRuleType.DATA_BAR
        elif isinstance(rule, IconSetRule) or rule_type_attr == "iconSet":
            return CFRuleType.ICON_SET
        elif isinstance(rule, CellIsRule) or rule_type_attr == "cellIs":
            return CFRuleType.CELL_IS
        elif isinstance(rule, FormulaRule) or rule_type_attr == "expression":
            return CFRuleType.FORMULA
        elif rule_type_attr == "top10":
            return CFRuleType.TOP_BOTTOM
        elif rule_type_attr == "aboveAverage":
            return CFRuleType.ABOVE_AVERAGE
        elif rule_type_attr == "duplicateValues":
            return CFRuleType.DUPLICATE
        elif rule_type_attr == "uniqueValues":
            return CFRuleType.UNIQUE
        elif rule_type_attr == "containsText":
            return CFRuleType.TEXT_CONTAINS
        elif rule_type_attr == "timePeriod":
            return CFRuleType.DATE_OCCURRING
        elif rule_type_attr == "containsBlanks":
            return CFRuleType.BLANK
        elif rule_type_attr == "notContainsBlanks":
            return CFRuleType.NOT_BLANK
        elif rule_type_attr == "containsErrors":
            return CFRuleType.ERROR
        elif rule_type_attr == "notContainsErrors":
            return CFRuleType.NOT_ERROR

        return CFRuleType.FORMULA

    def _get_formula(self, rule: Rule) -> str | None:
        """Extract formula from a rule."""
        try:
            if hasattr(rule, "formula") and rule.formula:
                if isinstance(rule.formula, (list, tuple)):
                    return str(rule.formula[0]) if rule.formula else None
                return str(rule.formula)
        except Exception:
            pass
        return None

    def _extract_values(self, rule: Rule) -> list:
        """Extract values/thresholds from a rule."""
        values = []
        try:
            # For cell is rules with values
            if hasattr(rule, "formula") and rule.formula:
                if isinstance(rule.formula, (list, tuple)):
                    values.extend(str(v) for v in rule.formula)

            # For color scale, data bar, icon set - extract CFVO values
            if hasattr(rule, "colorScale") and rule.colorScale:
                for cfvo in getattr(rule.colorScale, "cfvo", []):
                    values.append({"type": cfvo.type, "val": cfvo.val})

            if hasattr(rule, "dataBar") and rule.dataBar:
                for cfvo in getattr(rule.dataBar, "cfvo", []):
                    values.append({"type": cfvo.type, "val": cfvo.val})

            if hasattr(rule, "iconSet") and rule.iconSet:
                for cfvo in getattr(rule.iconSet, "cfvo", []):
                    values.append({"type": cfvo.type, "val": cfvo.val})

        except Exception:
            pass
        return values

    def _describe_rule(self, rule: Rule, rule_type: CFRuleType) -> str:
        """Create a human-readable description of the rule."""
        try:
            if rule_type == CFRuleType.COLOR_SCALE:
                return "Color scale (gradient coloring based on values)"

            if rule_type == CFRuleType.DATA_BAR:
                return "Data bar (in-cell bar chart)"

            if rule_type == CFRuleType.ICON_SET:
                icon_style = "default"
                if hasattr(rule, "iconSet") and rule.iconSet:
                    icon_style = getattr(rule.iconSet, "iconSet", "default")
                return f"Icon set ({icon_style})"

            if rule_type == CFRuleType.CELL_IS:
                operator = getattr(rule, "operator", "")
                formula = self._get_formula(rule)
                return f"Cell is {operator} {formula or ''}"

            if rule_type == CFRuleType.FORMULA:
                formula = self._get_formula(rule)
                return f"Formula: {formula or 'custom'}"

            if rule_type == CFRuleType.TOP_BOTTOM:
                rank = getattr(rule, "rank", 10)
                bottom = getattr(rule, "bottom", False)
                percent = getattr(rule, "percent", False)
                direction = "Bottom" if bottom else "Top"
                unit = "%" if percent else ""
                return f"{direction} {rank}{unit}"

            if rule_type == CFRuleType.ABOVE_AVERAGE:
                above = not getattr(rule, "aboveAverage", True) == False
                std_dev = getattr(rule, "stdDev", None)
                desc = "Above" if above else "Below"
                if std_dev:
                    desc += f" {std_dev} std dev from"
                return f"{desc} average"

            if rule_type == CFRuleType.DUPLICATE:
                return "Duplicate values"

            if rule_type == CFRuleType.UNIQUE:
                return "Unique values"

            if rule_type == CFRuleType.TEXT_CONTAINS:
                text = getattr(rule, "text", "")
                return f"Text contains: {text}"

            if rule_type == CFRuleType.DATE_OCCURRING:
                period = getattr(rule, "timePeriod", "")
                return f"Date occurring: {period}"

            if rule_type == CFRuleType.BLANK:
                return "Cell is blank"

            if rule_type == CFRuleType.NOT_BLANK:
                return "Cell is not blank"

            if rule_type == CFRuleType.ERROR:
                return "Cell contains error"

            if rule_type == CFRuleType.NOT_ERROR:
                return "Cell does not contain error"

        except Exception:
            pass

        return str(rule_type.value)
