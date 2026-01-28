"""Desktop Excel screenshot capture via xlwings (macOS/Windows)."""

from __future__ import annotations

import asyncio
import platform
import subprocess
import time
from datetime import datetime
from pathlib import Path

from ..models import ScreenshotInfo, SheetInfo, SheetVisibility


class DesktopExcelScreenshotter:
    """Captures screenshots of Excel sheets via desktop Excel application."""

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
            app.display_alerts = False
            app.screen_updating = True
        except Exception as e:
            print(f"  Could not start Excel: {e}", flush=True)
            return []

        try:
            # Open the workbook
            print(f"  Opening workbook: {file_path.name}", flush=True)
            wb = app.books.open(str(file_path), read_only=True)

            # Give Excel time to render
            time.sleep(2)

            # Capture each visible sheet
            for sheet_info in sheets:
                if sheet_info.visibility != SheetVisibility.VISIBLE:
                    print(f"  Skipping hidden sheet: {sheet_info.name}", flush=True)
                    continue

                screenshot_info = self._capture_sheet(wb, sheet_info)
                if screenshot_info:
                    screenshots.append(screenshot_info)

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

    def _capture_sheet(self, wb, sheet_info: SheetInfo) -> ScreenshotInfo | None:
        """Capture a screenshot of a single sheet."""
        try:
            print(f"  Capturing: {sheet_info.name}", flush=True)

            # Activate the sheet
            sheet = wb.sheets[sheet_info.name]
            sheet.activate()

            # Scroll to top-left
            sheet.range("A1").select()

            # Give time for rendering
            time.sleep(0.5)

            # Take screenshot
            screenshot_path = self.output_dir / f"{self._sanitize_filename(sheet_info.name)}.png"

            if self.system == "Darwin":
                success = self._screenshot_macos(screenshot_path)
            elif self.system == "Windows":
                success = self._screenshot_windows(screenshot_path)
            else:
                print(f"  Unsupported platform: {self.system}", flush=True)
                return None

            if success and screenshot_path.exists():
                return ScreenshotInfo(
                    sheet=sheet_info.name,
                    path=screenshot_path,
                    captured_at=datetime.now().isoformat(),
                )

        except Exception as e:
            print(f"  Error capturing sheet '{sheet_info.name}': {e}", flush=True)

        return None

    def _screenshot_macos(self, output_path: Path) -> bool:
        """Take screenshot on macOS using screencapture."""
        try:
            # Get Excel window ID using AppleScript
            script = '''
            tell application "Microsoft Excel"
                set windowId to id of window 1
            end tell
            return windowId
            '''
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=10
            )

            if result.returncode == 0:
                window_id = result.stdout.strip()
                # Capture specific window
                subprocess.run(
                    ["screencapture", "-l", window_id, "-o", str(output_path)],
                    timeout=10
                )
                return output_path.exists()
            else:
                # Fallback: capture frontmost window
                # First bring Excel to front
                subprocess.run(
                    ["osascript", "-e", 'tell application "Microsoft Excel" to activate'],
                    timeout=5
                )
                time.sleep(0.3)
                # Capture the frontmost window
                subprocess.run(
                    ["screencapture", "-w", "-o", str(output_path)],
                    timeout=10
                )
                return output_path.exists()

        except Exception as e:
            print(f"  macOS screenshot failed: {e}", flush=True)
            return False

    def _screenshot_windows(self, output_path: Path) -> bool:
        """Take screenshot on Windows using win32gui + PIL."""
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
