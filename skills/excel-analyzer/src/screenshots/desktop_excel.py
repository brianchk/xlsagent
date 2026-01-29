"""Desktop Excel screenshot capture via xlwings (Windows only)."""

from __future__ import annotations

import platform
import time
from datetime import datetime
from pathlib import Path

from ..models import ScreenshotInfo, SheetInfo, SheetVisibility


class DesktopExcelScreenshotter:
    """Captures screenshots of Excel sheets via desktop Excel application (Windows only)."""

    # Window dimensions for screenshots
    WINDOW_WIDTH = 1920
    WINDOW_HEIGHT = 1200

    # Zoom levels
    ZOOM_NORMAL = 100
    ZOOM_MIN = 25  # Minimum zoom for bird's eye view

    def __init__(self, output_dir: Path):
        """Initialize screenshotter.

        Args:
            output_dir: Directory to save screenshots
        """
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def capture_all_sheets(
        self,
        file_path: Path,
        sheets: list[SheetInfo],
    ) -> list[ScreenshotInfo]:
        """Capture screenshots of all visible sheets.

        Args:
            file_path: Path to the Excel file
            sheets: List of sheet info from extraction

        Returns:
            List of ScreenshotInfo objects
        """
        # Only supported on Windows
        if platform.system() != "Windows":
            print("  Screenshots only supported on Windows", flush=True)
            return []

        try:
            import xlwings as xw
        except ImportError:
            print("  xlwings not installed, skipping screenshots", flush=True)
            return []

        screenshots = []

        # Open Excel (visible but minimized to reduce disruption)
        print("  Opening Excel...", flush=True)
        try:
            app = xw.App(visible=True, add_book=False)
            app.display_alerts = False
            app.screen_updating = True  # Keep enabled so sheets render properly
            # Windows-specific: suppress dialogs
            try:
                app.api.AskToUpdateLinks = False
                app.api.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
        except Exception as e:
            print(f"  Could not start Excel: {e}", flush=True)
            return []

        try:
            # Open workbook read-only
            print(f"  Opening workbook: {file_path.name}", flush=True)
            wb = app.books.open(
                str(file_path),
                read_only=True,
                update_links=False,
                ignore_read_only_recommended=True,
            )

            # Set window size and position
            self._set_window_size(app)

            # Give Excel time to render
            time.sleep(0.5)

            # Capture each visible sheet
            for sheet_info in sheets:
                if sheet_info.visibility != SheetVisibility.VISIBLE:
                    continue

                sheet_screenshots = self._capture_sheet(wb, sheet_info)
                screenshots.extend(sheet_screenshots)

            # Capture individual charts
            chart_screenshots = self._capture_charts(wb)
            screenshots.extend(chart_screenshots)

            # Close workbook without saving
            wb.close()

        except Exception as e:
            print(f"  Error during screenshot capture: {e}", flush=True)
        finally:
            try:
                app.quit()
            except Exception:
                pass

        return screenshots

    def _set_window_size(self, app) -> None:
        """Set Excel window to standard size for consistent screenshots."""
        try:
            app.api.ActiveWindow.WindowState = -4143  # xlNormal
            app.api.ActiveWindow.Top = 0
            app.api.ActiveWindow.Left = 0
            app.api.ActiveWindow.Width = self.WINDOW_WIDTH
            app.api.ActiveWindow.Height = self.WINDOW_HEIGHT
        except Exception as e:
            print(f"  Could not set window size: {e}", flush=True)

    def _capture_sheet(self, wb, sheet_info: SheetInfo) -> list[ScreenshotInfo]:
        """Capture screenshots of a single sheet (normal + bird's eye view)."""
        screenshots = []
        try:
            print(f"  Capturing: {sheet_info.name}", flush=True)

            # Activate the sheet
            sheet = wb.sheets[sheet_info.name]
            sheet.activate()

            # Scroll to top-left
            sheet.range("A1").select()
            time.sleep(0.3)

            # Screenshot 1: Normal view (100% zoom)
            self._set_zoom(sheet, self.ZOOM_NORMAL)
            time.sleep(0.2)

            normal_path = self.output_dir / f"{self._sanitize_filename(sheet_info.name)}_{self.ZOOM_NORMAL}.png"
            if self._take_screenshot(normal_path, sheet):
                screenshots.append(ScreenshotInfo(
                    sheet=sheet_info.name,
                    path=normal_path,
                    captured_at=datetime.now().isoformat(),
                ))

            # Screenshot 2: Bird's eye view - zoom out until content fits
            birdseye_zoom = self._calculate_fit_zoom(sheet, sheet_info)
            # Always capture bird's eye if different from normal (even if same due to small sheet)
            if birdseye_zoom < self.ZOOM_NORMAL or sheet_info.row_count > 35 or sheet_info.col_count > 12:
                self._set_zoom(sheet, birdseye_zoom)
                time.sleep(0.3)

                birdseye_path = self.output_dir / f"{self._sanitize_filename(sheet_info.name)}_{birdseye_zoom}.png"
                if self._take_screenshot(birdseye_path, sheet):
                    screenshots.append(ScreenshotInfo(
                        sheet=sheet_info.name,
                        path=birdseye_path,
                        captured_at=datetime.now().isoformat(),
                    ))

            # Reset zoom
            self._set_zoom(sheet, self.ZOOM_NORMAL)

        except Exception as e:
            print(f"  Error capturing sheet '{sheet_info.name}': {e}", flush=True)

        return screenshots

    def _calculate_fit_zoom(self, sheet, sheet_info: SheetInfo) -> int:
        """Calculate zoom level to fit all content in the window.

        Returns zoom level between ZOOM_MIN and ZOOM_NORMAL.
        """
        try:
            # Get used range directly from Excel for accuracy
            try:
                used_range = sheet.api.UsedRange
                rows = used_range.Rows.Count
                cols = used_range.Columns.Count
                print(f"      Excel UsedRange: {rows} rows x {cols} cols", flush=True)
            except Exception:
                # Fallback to extracted data
                rows = sheet_info.row_count or 50
                cols = sheet_info.col_count or 20
                print(f"      Using extracted data: {rows} rows x {cols} cols", flush=True)

            # Approximate visible rows/cols at 100% zoom in our window
            # Be conservative - assume less visible to ensure content fits
            # At 100% zoom in 1920x1200: roughly 30 rows, 10 columns visible
            visible_rows_100 = 30
            visible_cols_100 = 10

            # Calculate zoom needed to fit all rows and all columns
            zoom_for_rows = int((visible_rows_100 / rows) * 100) if rows > 0 else 100
            zoom_for_cols = int((visible_cols_100 / cols) * 100) if cols > 0 else 100

            print(f"      Zoom needed: rows={zoom_for_rows}%, cols={zoom_for_cols}%", flush=True)

            # Use the smaller zoom (more zoomed out) to fit both dimensions
            zoom = min(zoom_for_rows, zoom_for_cols, self.ZOOM_NORMAL)

            # Clamp to minimum
            zoom = max(zoom, self.ZOOM_MIN)

            # Round to nearest 5%
            zoom = (zoom // 5) * 5

            print(f"      -> Using bird's eye zoom: {zoom}%", flush=True)
            return zoom

        except Exception:
            return 50  # Default fallback

    def _set_zoom(self, sheet, zoom_level: int) -> None:
        """Set the zoom level for the active sheet."""
        try:
            sheet.book.app.api.ActiveWindow.Zoom = zoom_level
        except Exception as e:
            print(f"  Could not set zoom to {zoom_level}%: {e}", flush=True)

    def _take_screenshot(self, output_path: Path, sheet=None) -> bool:
        """Take a screenshot of the sheet area (not the whole window).

        Uses Excel's CopyPicture to capture just the visible cells.
        """
        try:
            from PIL import ImageGrab
            import pythoncom
            import win32clipboard

            if sheet is None:
                print("  No sheet provided for screenshot", flush=True)
                return False

            # Get the visible range in the current window
            try:
                excel_window = sheet.book.app.api.ActiveWindow
                visible_range = excel_window.VisibleRange

                # Copy the visible range as a picture
                visible_range.CopyPicture(1, 2)  # xlScreen=1, xlPicture=2 (or xlBitmap=-4147)
                time.sleep(0.2)

                # Grab from clipboard
                img = ImageGrab.grabclipboard()
                if img:
                    img.save(str(output_path), "PNG")
                    return output_path.exists()
                else:
                    print("  Could not grab image from clipboard", flush=True)
                    return False

            except Exception as e:
                print(f"  CopyPicture failed: {e}", flush=True)
                return False

        except ImportError as e:
            print(f"  Required modules not available: {e}", flush=True)
            return False
        except Exception as e:
            print(f"  Screenshot failed: {e}", flush=True)
            return False

    def _capture_charts(self, wb) -> list[ScreenshotInfo]:
        """Capture individual chart images."""
        screenshots = []
        chart_dir = self.output_dir / "charts"
        chart_dir.mkdir(exist_ok=True)

        print("  Capturing charts...", flush=True)
        try:
            for sheet in wb.sheets:
                try:
                    # Access charts through Excel COM API directly
                    excel_sheet = sheet.api
                    chart_objects = excel_sheet.ChartObjects()

                    if chart_objects.Count == 0:
                        continue

                    for i in range(1, chart_objects.Count + 1):  # Excel is 1-indexed
                        try:
                            chart_obj = chart_objects.Item(i)
                            chart_name = chart_obj.Name or f"Chart_{i}"

                            safe_name = self._sanitize_filename(f"{sheet.name}_{chart_name}")
                            # Use absolute path with forward slashes for Windows COM
                            output_path = chart_dir / f"{safe_name}.png"
                            output_path_str = str(output_path.resolve())

                            # Try CopyPicture + paste to new chart sheet + export
                            # This is more reliable than direct Export
                            try:
                                chart_obj.CopyPicture(1, 2)  # xlScreen=1, xlPicture=2
                                time.sleep(0.1)

                                # Save via PIL from clipboard
                                from PIL import ImageGrab
                                img = ImageGrab.grabclipboard()
                                if img:
                                    img.save(output_path_str, "PNG")
                            except Exception:
                                # Fallback: try direct export
                                chart = chart_obj.Chart
                                chart.Export(output_path_str, "PNG")

                            if output_path.exists():
                                print(f"    Chart: {chart_name} ({sheet.name})", flush=True)
                                screenshots.append(ScreenshotInfo(
                                    sheet=sheet.name,
                                    path=output_path,
                                    captured_at=datetime.now().isoformat(),
                                ))
                        except Exception as e:
                            print(f"    Could not export chart {i} on {sheet.name}: {e}", flush=True)
                except Exception:
                    # Sheet might not have charts or API access failed
                    continue

        except Exception as e:
            print(f"  Error capturing charts: {e}", flush=True)

        return screenshots

    def _sanitize_filename(self, name: str) -> str:
        """Sanitize a string for use as a filename."""
        invalid_chars = '<>:"/\\|?*'
        result = name
        for char in invalid_chars:
            result = result.replace(char, "_")
        if len(result) > 100:
            result = result[:100]
        return result


def capture_desktop_screenshots(
    file_path: Path,
    sheets: list[SheetInfo],
    output_dir: Path,
) -> list[ScreenshotInfo]:
    """Convenience function to capture desktop Excel screenshots (Windows only).

    Args:
        file_path: Path to the Excel file
        sheets: List of sheet info
        output_dir: Where to save screenshots

    Returns:
        List of ScreenshotInfo objects
    """
    screenshotter = DesktopExcelScreenshotter(output_dir)
    return screenshotter.capture_all_sheets(file_path, sheets)
