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
    ZOOM_MIN = 40  # Minimum zoom for bird's eye view (25% was too extreme)

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

    # Target size for detail view in points (roughly 1600x900 pixels at 96 DPI)
    # 1 point = 1/72 inch, 96 DPI means 1 point ≈ 1.33 pixels
    DETAIL_TARGET_WIDTH_PT = 1200  # ~1600 pixels
    DETAIL_TARGET_HEIGHT_PT = 675  # ~900 pixels

    # Fallback if we can't calculate sizes
    DETAIL_FALLBACK_SIZE = (40, 20)

    def _calculate_detail_size(self, sheet) -> tuple[int, int]:
        """Calculate rows and cols that fit within target pixel dimensions."""
        try:
            ws = sheet.api

            # Calculate columns that fit within target width
            total_width = 0
            num_cols = 0
            for col in range(1, 100):  # Max 100 columns
                try:
                    col_width_pt = ws.Columns(col).Width  # Width in points
                    if total_width + col_width_pt > self.DETAIL_TARGET_WIDTH_PT:
                        break
                    total_width += col_width_pt
                    num_cols = col
                except Exception:
                    break

            # Calculate rows that fit within target height
            total_height = 0
            num_rows = 0
            for row in range(1, 200):  # Max 200 rows
                try:
                    row_height_pt = ws.Rows(row).Height  # Height in points
                    if total_height + row_height_pt > self.DETAIL_TARGET_HEIGHT_PT:
                        break
                    total_height += row_height_pt
                    num_rows = row
                except Exception:
                    break

            # Ensure minimum size
            num_rows = max(num_rows, 10)
            num_cols = max(num_cols, 5)

            print(f"    Detail size: {num_rows} rows x {num_cols} cols (based on cell dimensions)", flush=True)
            return (num_rows, num_cols)

        except Exception as e:
            print(f"    Could not calculate detail size: {e}, using fallback", flush=True)
            return self.DETAIL_FALLBACK_SIZE

    def _capture_sheet(self, wb, sheet_info: SheetInfo) -> list[ScreenshotInfo]:
        """Capture screenshots of a single sheet using Range.CopyPicture."""
        screenshots = []
        try:
            print(f"  Capturing: {sheet_info.name}", flush=True)

            # Activate the sheet
            sheet = wb.sheets[sheet_info.name]
            sheet.activate()
            time.sleep(0.2)

            # Capture 1: Full sheet (all actual data, not just UsedRange)
            full_path = self.output_dir / f"{self._sanitize_filename(sheet_info.name)}_full.png"
            if self._capture_range_as_image(sheet, full_path, fixed_size=None):
                screenshots.append(ScreenshotInfo(
                    sheet=sheet_info.name,
                    path=full_path,
                    captured_at=datetime.now().isoformat(),
                ))

            # Capture 2: Detail view (size based on actual row/col dimensions)
            detail_size = self._calculate_detail_size(sheet)
            detail_path = self.output_dir / f"{self._sanitize_filename(sheet_info.name)}_detail.png"
            if self._capture_range_as_image(sheet, detail_path, fixed_size=detail_size):
                screenshots.append(ScreenshotInfo(
                    sheet=sheet_info.name,
                    path=detail_path,
                    captured_at=datetime.now().isoformat(),
                ))

        except Exception as e:
            print(f"  Error capturing sheet '{sheet_info.name}': {e}", flush=True)

        return screenshots

    def _get_actual_data_range(self, sheet) -> tuple[int, int]:
        """Find the actual last row and column with data (not just formatting)."""
        try:
            ws = sheet.api

            # Find last row with data
            last_row = 1
            try:
                # Search from bottom up for any value
                found = ws.Cells.Find(
                    What="*",
                    SearchOrder=-4134,  # xlByRows
                    SearchDirection=2,  # xlPrevious
                )
                if found:
                    last_row = found.Row
            except Exception:
                last_row = ws.UsedRange.Rows.Count

            # Find last column with data
            last_col = 1
            try:
                # Search from right to left for any value
                found = ws.Cells.Find(
                    What="*",
                    SearchOrder=-4152,  # xlByColumns
                    SearchDirection=2,  # xlPrevious
                )
                if found:
                    last_col = found.Column
            except Exception:
                last_col = ws.UsedRange.Columns.Count

            return last_row, last_col
        except Exception:
            return 100, 20  # Fallback

    def _capture_range_as_image(self, sheet, output_path: Path, fixed_size: tuple[int, int] | None = None) -> bool:
        """Capture a range as an image using CopyPicture (no window chrome).

        Args:
            sheet: The sheet to capture
            output_path: Where to save the image
            fixed_size: If provided, capture fixed (rows, cols) from A1. Otherwise capture all data.
        """
        try:
            from PIL import ImageGrab

            ws = sheet.api

            if fixed_size:
                # Fixed size capture from A1 (detail view)
                rows, cols = fixed_size
                capture_range = ws.Range(ws.Cells(1, 1), ws.Cells(rows, cols))
                label = f"detail ({rows} rows x {cols} cols)"
            else:
                # Find actual data extent (not just UsedRange which includes formatting)
                last_row, last_col = self._get_actual_data_range(sheet)
                capture_range = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col))
                label = f"full ({last_row} rows x {last_col} cols)"

            print(f"    Capturing {label}...", flush=True)

            # Copy range as picture to clipboard
            # xlScreen = 1, xlBitmap = 2
            capture_range.CopyPicture(Appearance=1, Format=2)
            time.sleep(0.3)

            # Grab from clipboard
            img = ImageGrab.grabclipboard()
            if img:
                img.save(str(output_path), "PNG")
                return output_path.exists()
            else:
                print(f"    Could not grab image from clipboard", flush=True)
                return False

        except Exception as e:
            print(f"    Range capture failed: {e}", flush=True)
            return False

    def _hide_excel_ui(self, app) -> dict:
        """Hide Excel UI elements to maximize sheet area. Returns previous state."""
        state = {}
        try:
            api = app.api

            # Save current state
            state['formula_bar'] = api.DisplayFormulaBar
            state['status_bar'] = api.DisplayStatusBar

            # Hide formula bar and status bar
            api.DisplayFormulaBar = False
            api.DisplayStatusBar = False

            # Minimize ribbon (may cause brief focus steal but maximizes sheet area)
            try:
                api.ExecuteExcel4Macro('SHOW.TOOLBAR("Ribbon",False)')
            except Exception:
                try:
                    api.CommandBars.ExecuteMso("MinimizeRibbon")
                except Exception:
                    pass

            time.sleep(0.2)  # Let UI update

        except Exception as e:
            print(f"    Warning: Could not hide UI elements: {e}", flush=True)

        return state

    def _restore_excel_ui(self, app, state: dict) -> None:
        """Restore Excel UI elements to previous state."""
        try:
            api = app.api

            # Restore formula bar and status bar
            if 'formula_bar' in state:
                api.DisplayFormulaBar = state['formula_bar']
            if 'status_bar' in state:
                api.DisplayStatusBar = state['status_bar']

            # Restore ribbon
            try:
                api.ExecuteExcel4Macro('SHOW.TOOLBAR("Ribbon",True)')
            except Exception:
                try:
                    api.CommandBars.ExecuteMso("MinimizeRibbon")  # Toggle back
                except Exception:
                    pass

        except Exception as e:
            print(f"    Warning: Could not restore UI elements: {e}", flush=True)

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
        """Take a screenshot of the Excel window and crop to sheet area."""
        try:
            import win32gui
            import win32ui
            import win32con
            from PIL import Image
            import ctypes

            # Find Excel window
            def find_excel_window(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if "Excel" in title:
                        windows.append(hwnd)
                return True

            windows = []
            win32gui.EnumWindows(find_excel_window, windows)

            if not windows:
                print("  Could not find Excel window", flush=True)
                return False

            hwnd = windows[0]

            # Get window dimensions
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top

            # Capture the window
            hwnd_dc = win32gui.GetWindowDC(hwnd)
            mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc = mfc_dc.CreateCompatibleDC()

            bitmap = win32ui.CreateBitmap()
            bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
            save_dc.SelectObject(bitmap)

            # Use PrintWindow for capture
            result = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)

            if result == 0:
                save_dc.BitBlt((0, 0), (width, height), mfc_dc, (0, 0), win32con.SRCCOPY)

            # Convert to PIL Image and save
            bmpinfo = bitmap.GetInfo()
            bmpstr = bitmap.GetBitmapBits(True)
            img = Image.frombuffer(
                'RGB',
                (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                bmpstr, 'raw', 'BGRX', 0, 1
            )

            # No cropping - UI elements are hidden programmatically
            img.save(str(output_path))

            # Cleanup
            win32gui.DeleteObject(bitmap.GetHandle())
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwnd_dc)

            return output_path.exists()

        except ImportError as e:
            print(f"  Required Windows modules not available: {e}", flush=True)
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
                                    is_chart=True,
                                    chart_name=chart_name,
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
