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

            # Get chart title
            title = None
            try:
                if chart.title:
                    if hasattr(chart.title, "text"):
                        title = chart.title.text
                    elif isinstance(chart.title, str):
                        title = chart.title
            except Exception:
                pass

            # Get data range (from first series if available)
            data_range = None
            try:
                if chart.series and len(chart.series) > 0:
                    first_series = chart.series[0]
                    if hasattr(first_series, "val") and first_series.val:
                        if hasattr(first_series.val, "numRef") and first_series.val.numRef:
                            data_range = first_series.val.numRef.f
            except Exception:
                pass

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
