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
    ZOOM_BIRDSEYE = 65

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
            if self._take_screenshot(normal_path):
                screenshots.append(ScreenshotInfo(
                    sheet=sheet_info.name,
                    path=normal_path,
                    captured_at=datetime.now().isoformat(),
                ))

            # Screenshot 2: Bird's eye view
            self._set_zoom(sheet, self.ZOOM_BIRDSEYE)
            time.sleep(0.3)

            birdseye_path = self.output_dir / f"{self._sanitize_filename(sheet_info.name)}_{self.ZOOM_BIRDSEYE}.png"
            if self._take_screenshot(birdseye_path):
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

    def _set_zoom(self, sheet, zoom_level: int) -> None:
        """Set the zoom level for the active sheet."""
        try:
            sheet.book.app.api.ActiveWindow.Zoom = zoom_level
        except Exception as e:
            print(f"  Could not set zoom to {zoom_level}%: {e}", flush=True)

    def _take_screenshot(self, output_path: Path) -> bool:
        """Take a screenshot of the Excel window."""
        try:
            import win32gui
            import win32ui
            import win32con
            from PIL import Image

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

            # Note: We don't call SetForegroundWindow to avoid stealing focus
            # PrintWindow with flag 2 can capture background windows

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

            # Use PrintWindow for better capture
            import ctypes
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
