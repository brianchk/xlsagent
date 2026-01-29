"""Chart extractor."""

from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet

from ..models import ChartInfo
from .base import BaseExtractor


class ChartExtractor(BaseExtractor):
    """Extracts chart definitions from all sheets."""

    name = "charts"

    # Map openpyxl chart types to readable names
    CHART_TYPE_NAMES = {
        "AreaChart": "Area Chart",
        "AreaChart3D": "3D Area Chart",
        "BarChart": "Bar Chart",
        "BarChart3D": "3D Bar Chart",
        "BubbleChart": "Bubble Chart",
        "DoughnutChart": "Doughnut Chart",
        "LineChart": "Line Chart",
        "LineChart3D": "3D Line Chart",
        "PieChart": "Pie Chart",
        "PieChart3D": "3D Pie Chart",
        "ProjectedPieChart": "Projected Pie Chart",
        "RadarChart": "Radar Chart",
        "ScatterChart": "Scatter Chart",
        "StockChart": "Stock Chart",
        "SurfaceChart": "Surface Chart",
        "SurfaceChart3D": "3D Surface Chart",
    }

    def extract(self) -> list[ChartInfo]:
        """Extract all chart definitions.

        Returns:
            List of ChartInfo objects
        """
        charts = []

        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            if not isinstance(sheet, Worksheet):
                continue

            sheet_charts = self._extract_sheet_charts(sheet, sheet_name)
            charts.extend(sheet_charts)

        return charts

    def _extract_sheet_charts(self, sheet: Worksheet, sheet_name: str) -> list[ChartInfo]:
        """Extract charts from a sheet."""
        charts = []

        try:
            for idx, chart in enumerate(sheet._charts):
                info = self._create_chart_info(chart, sheet_name, idx)
                if info:
                    charts.append(info)
        except Exception:
            pass

        return charts

    def _create_chart_info(self, chart, sheet_name: str, index: int) -> ChartInfo | None:
        """Create ChartInfo from a chart object."""
        try:
            # Get chart type
            chart_class_name = chart.__class__.__name__
            chart_type = self.CHART_TYPE_NAMES.get(chart_class_name, chart_class_name)

            # Get chart title - handle openpyxl's complex title structure
            title = self._extract_title(chart)

            # Get data ranges from all series
            data_range = self._extract_data_ranges(chart)

            # Get position
            position = None
            try:
                if hasattr(chart, "anchor"):
                    anchor = chart.anchor
                    if hasattr(anchor, "_from"):
                        pos = anchor._from
                        position = f"Col {pos.col}, Row {pos.row}"
            except Exception:
                pass

            return ChartInfo(
                name=title or f"Chart {index + 1}",
                sheet=sheet_name,
                chart_type=chart_type,
                title=title,
                data_range=data_range,
                position=position,
            )
        except Exception:
            return None

    def _extract_title(self, chart) -> str | None:
        """Extract the actual text from a chart title."""
        try:
            if not chart.title:
                return None

            title_obj = chart.title

            # Direct string
            if isinstance(title_obj, str):
                return title_obj

            # Try to get text from RichText structure
            if hasattr(title_obj, 'tx') and title_obj.tx:
                tx = title_obj.tx
                if hasattr(tx, 'rich') and tx.rich:
                    # Extract text from rich text paragraphs
                    texts = []
                    for p in tx.rich.p or []:
                        for r in p.r or []:
                            if hasattr(r, 't') and r.t:
                                texts.append(r.t)
                    if texts:
                        return ''.join(texts)
                if hasattr(tx, 'strRef') and tx.strRef:
                    # Title references a cell
                    if hasattr(tx.strRef, 'f') and tx.strRef.f:
                        return f"={tx.strRef.f}"

            # Try simple text attribute
            if hasattr(title_obj, 'text') and title_obj.text:
                text = title_obj.text
                if isinstance(text, str):
                    return text

            return None
        except Exception:
            return None

    def _extract_data_ranges(self, chart) -> str | None:
        """Extract data ranges from chart series."""
        try:
            ranges = []
            if chart.series:
                for series in chart.series[:3]:  # Limit to first 3 series
                    # Get values reference
                    if hasattr(series, 'val') and series.val:
                        val = series.val
                        if hasattr(val, 'numRef') and val.numRef and val.numRef.f:
                            ranges.append(val.numRef.f)

                    # Get categories reference
                    if hasattr(series, 'cat') and series.cat:
                        cat = series.cat
                        if hasattr(cat, 'numRef') and cat.numRef and cat.numRef.f:
                            if cat.numRef.f not in ranges:
                                ranges.append(f"(cat) {cat.numRef.f}")
                        elif hasattr(cat, 'strRef') and cat.strRef and cat.strRef.f:
                            if cat.strRef.f not in ranges:
                                ranges.append(f"(cat) {cat.strRef.f}")

            if ranges:
                result = "; ".join(ranges)
                if len(chart.series) > 3:
                    result += f" ... (+{len(chart.series) - 3} more series)"
                return result
            return None
        except Exception:
            return None
