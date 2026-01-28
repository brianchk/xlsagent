"""Desktop Excel screenshot capture via xlwings (macOS/Windows)."""

from __future__ import annotations

import platform
import subprocess
import time
from datetime import datetime
from pathlib import Path

from ..models import ScreenshotInfo, SheetInfo, SheetVisibility


class DesktopExcelScreenshotter:
    """Captures screenshots of Excel sheets via desktop Excel application."""

    # Window dimensions for screenshots
    WINDOW_WIDTH = 1920
    WINDOW_HEIGHT = 1200

    # Zoom levels
    ZOOM_NORMAL = 100
    ZOOM_BIRDSEYE = 25

    def __init__(self, output_dir: Path):
        """Initialize screenshotter.

        Args:
            output_dir: Directory to save screenshots
        """
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.system = platform.system()

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
        try:
            import xlwings as xw
        except ImportError:
            print("  xlwings not installed, skipping desktop screenshots", flush=True)
            return []

        screenshots = []

        # Open Excel
        print("  Opening Excel...", flush=True)
        try:
            # Use visible=True so we can screenshot
            app = xw.App(visible=True, add_book=False)
            # Suppress all prompts and alerts
            app.display_alerts = False
            app.screen_updating = True
            # Disable automatic link updates
            app.api.AskToUpdateLinks = False
        except Exception as e:
            print(f"  Could not start Excel: {e}", flush=True)
            return []

        try:
            # Open the workbook read-only, suppressing prompts
            print(f"  Opening workbook: {file_path.name}", flush=True)
            wb = app.books.open(
                str(file_path),
                read_only=True,
                update_links=False,  # Don't update external links
                ignore_read_only_recommended=True,  # Skip read-only prompt
            )

            # Disable macros execution (already suppressed by display_alerts=False)
            # Set window size
            self._set_window_size(app)

            # Give Excel time to render
            time.sleep(1)

            # Capture each visible sheet
            for sheet_info in sheets:
                if sheet_info.visibility != SheetVisibility.VISIBLE:
                    print(f"  Skipping hidden sheet: {sheet_info.name}", flush=True)
                    continue

                sheet_screenshots = self._capture_sheet(wb, sheet_info)
                screenshots.extend(sheet_screenshots)

            # Close workbook without saving
            wb.close()

        except Exception as e:
            print(f"  Error during screenshot capture: {e}", flush=True)
        finally:
            # Quit Excel
            try:
                app.quit()
            except Exception:
                pass

        return screenshots

    def _set_window_size(self, app) -> None:
        """Set Excel window to standard size for consistent screenshots."""
        try:
            if self.system == "Darwin":
                # macOS: Use AppleScript to resize window
                script = f'''
                tell application "Microsoft Excel"
                    activate
                    set bounds of window 1 to {{0, 0, {self.WINDOW_WIDTH}, {self.WINDOW_HEIGHT}}}
                end tell
                '''
                subprocess.run(
                    ["osascript", "-e", script],
                    capture_output=True,
                    timeout=5
                )
            elif self.system == "Windows":
                # Windows: Use xlwings API
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

            # Give time for rendering
            time.sleep(0.3)

            # Screenshot 1: Normal view (100% zoom)
            self._set_zoom(sheet, self.ZOOM_NORMAL)
            time.sleep(0.2)

            normal_path = self.output_dir / f"{self._sanitize_filename(sheet_info.name)}_100.png"
            if self._take_screenshot(normal_path):
                screenshots.append(ScreenshotInfo(
                    sheet=sheet_info.name,
                    path=normal_path,
                    captured_at=datetime.now().isoformat(),
                ))

            # Screenshot 2: Bird's eye view (25% zoom)
            self._set_zoom(sheet, self.ZOOM_BIRDSEYE)
            time.sleep(0.3)

            birdseye_path = self.output_dir / f"{self._sanitize_filename(sheet_info.name)}_25.png"
            if self._take_screenshot(birdseye_path):
                screenshots.append(ScreenshotInfo(
                    sheet=sheet_info.name,
                    path=birdseye_path,
                    captured_at=datetime.now().isoformat(),
                ))

            # Reset zoom to normal
            self._set_zoom(sheet, self.ZOOM_NORMAL)

        except Exception as e:
            print(f"  Error capturing sheet '{sheet_info.name}': {e}", flush=True)

        return screenshots

    def _set_zoom(self, sheet, zoom_level: int) -> None:
        """Set the zoom level for the active sheet."""
        try:
            if self.system == "Darwin":
                # macOS: Use AppleScript to set zoom
                script = f'''
                tell application "Microsoft Excel"
                    set view of active window to normal view
                    set zoom of active window to {zoom_level}
                end tell
                '''
                subprocess.run(
                    ["osascript", "-e", script],
                    capture_output=True,
                    timeout=5
                )
            else:
                # Windows: Use xlwings API
                sheet.book.app.api.ActiveWindow.Zoom = zoom_level
        except Exception as e:
            print(f"  Could not set zoom to {zoom_level}%: {e}", flush=True)

    def _take_screenshot(self, output_path: Path) -> bool:
        """Take a screenshot of the Excel window only."""
        if self.system == "Darwin":
            return self._screenshot_macos(output_path)
        elif self.system == "Windows":
            return self._screenshot_windows(output_path)
        else:
            print(f"  Unsupported platform: {self.system}", flush=True)
            return False

    def _screenshot_macos(self, output_path: Path) -> bool:
        """Take screenshot of Excel window on macOS using screencapture."""
        try:
            # Bring Excel to front
            subprocess.run(
                ["osascript", "-e", 'tell application "Microsoft Excel" to activate'],
                capture_output=True,
                timeout=5
            )
            time.sleep(0.2)

            # Get the window position and size via System Events
            script = '''
            tell application "System Events"
                set excelProcess to first application process whose name is "Microsoft Excel"
                set frontWindow to window 1 of excelProcess
                set winPos to position of frontWindow
                set winSize to size of frontWindow
                return {item 1 of winPos, item 2 of winPos, item 1 of winSize, item 2 of winSize}
            end tell
            '''
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=10
            )

            if result.returncode != 0 or not result.stdout.strip():
                print(f"  Could not get Excel window bounds: {result.stderr}", flush=True)
                return False

            # Parse the bounds: x, y, width, height
            bounds = result.stdout.strip().split(", ")
            if len(bounds) != 4:
                print(f"  Invalid window bounds: {result.stdout}", flush=True)
                return False

            x, y, w, h = [int(b) for b in bounds]

            # Capture the window region
            region = f"{x},{y},{w},{h}"
            capture_result = subprocess.run(
                ["screencapture", f"-R{region}", "-x", "-o", str(output_path)],
                capture_output=True,
                timeout=10
            )

            if not output_path.exists():
                print(f"  Screenshot failed: {capture_result.stderr.decode() if capture_result.stderr else 'unknown error'}", flush=True)
                return False

            return True

        except Exception as e:
            print(f"  macOS screenshot failed: {e}", flush=True)
            return False

    def _screenshot_windows(self, output_path: Path) -> bool:
        """Take screenshot of Excel window on Windows using win32gui + PIL."""
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

            # Bring window to front
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.2)

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

            # Use PrintWindow for better capture (works with DWM)
            import ctypes
            result = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)

            if result == 0:
                # Fallback to BitBlt
                save_dc.BitBlt((0, 0), (width, height), mfc_dc, (0, 0), win32con.SRCCOPY)

            # Convert to PIL Image
            bmpinfo = bitmap.GetInfo()
            bmpstr = bitmap.GetBitmapBits(True)
            img = Image.frombuffer(
                'RGB',
                (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                bmpstr, 'raw', 'BGRX', 0, 1
            )

            # Save
            img.save(str(output_path))

            # Cleanup
            win32gui.DeleteObject(bitmap.GetHandle())
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwnd_dc)

            return output_path.exists()

        except ImportError as e:
            # Fallback to pyautogui if win32 modules not available
            print(f"  win32 modules not available ({e}), trying pyautogui...", flush=True)
            try:
                import pyautogui
                time.sleep(0.2)
                screenshot = pyautogui.screenshot()
                screenshot.save(str(output_path))
                return True
            except ImportError:
                print("  Neither win32gui nor pyautogui available", flush=True)
                return False
        except Exception as e:
            print(f"  Windows screenshot failed: {e}", flush=True)
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
    """Convenience function to capture desktop Excel screenshots.

    Args:
        file_path: Path to the Excel file
        sheets: List of sheet info
        output_dir: Where to save screenshots

    Returns:
        List of ScreenshotInfo objects
    """
    screenshotter = DesktopExcelScreenshotter(output_dir)
    return screenshotter.capture_all_sheets(file_path, sheets)
