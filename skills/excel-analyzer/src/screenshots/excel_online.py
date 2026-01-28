"""Excel Online screenshot capture via Playwright."""

from __future__ import annotations

import asyncio
from datetime import datetime
from pathlib import Path

from playwright.async_api import BrowserContext, Page

from ..models import ScreenshotInfo, SheetInfo, SheetVisibility


class ExcelOnlineScreenshotter:
    """Captures screenshots of Excel sheets via Excel Online."""

    # Selectors for Excel Online UI elements
    SELECTORS = {
        # Sheet tab bar
        "sheet_tab": '[data-automation-id="SheetTab-{name}"]',
        "sheet_tabs_container": '[data-automation-id="SheetTabBar"]',
        "active_sheet_tab": '[data-automation-id="SheetTab"][aria-selected="true"]',

        # Sheet content area
        "spreadsheet_canvas": '[data-automation-id="spreadsheet"]',
        "grid_container": '.ewr-sheet-container',

        # Hidden sheet menu
        "sheet_tab_context_menu": '[data-automation-id="SheetTabContextMenu"]',
        "unhide_sheets_item": 'button:has-text("Unhide")',
        "unhide_dialog": '[data-automation-id="UnhideSheetDialog"]',

        # Loading indicators
        "loading": '[data-automation-id="loading"]',
        "calculating": '.calculating-indicator',

        # User menu (indicates authenticated)
        "user_menu": '[data-automation-id="mectrl_main_trigger"]',
    }

    def __init__(self, output_dir: Path):
        """Initialize screenshotter.

        Args:
            output_dir: Directory to save screenshots
        """
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)

    async def capture_all_sheets(
        self,
        context: BrowserContext,
        sharepoint_url: str,
        sheets: list[SheetInfo],
    ) -> list[ScreenshotInfo]:
        """Capture screenshots of all sheets.

        Args:
            context: Authenticated browser context
            sharepoint_url: URL to the Excel file in SharePoint
            sheets: List of sheet info from extraction

        Returns:
            List of ScreenshotInfo objects
        """
        screenshots = []
        page = await context.new_page()

        try:
            # Navigate to Excel Online
            await page.goto(sharepoint_url)
            await self._wait_for_excel_online_load(page)

            # Capture each sheet
            for sheet in sheets:
                screenshot_info = await self._capture_sheet(page, sheet)
                if screenshot_info:
                    screenshots.append(screenshot_info)

        except Exception as e:
            print(f"Error capturing screenshots: {e}")
        finally:
            await page.close()

        return screenshots

    async def _wait_for_excel_online_load(self, page: Page, timeout: int = 30000) -> None:
        """Wait for Excel Online to fully load."""
        try:
            # Wait for loading to complete
            await page.wait_for_load_state("networkidle", timeout=timeout)

            # Wait for the spreadsheet canvas to appear
            await page.wait_for_selector(
                self.SELECTORS["spreadsheet_canvas"],
                timeout=timeout
            )

            # Wait a bit for rendering
            await asyncio.sleep(2)

            # Wait for any calculating indicator to disappear
            try:
                calculating = await page.query_selector(self.SELECTORS["calculating"])
                if calculating:
                    await page.wait_for_selector(
                        self.SELECTORS["calculating"],
                        state="hidden",
                        timeout=30000
                    )
            except Exception:
                pass

        except Exception as e:
            print(f"Warning: Excel Online may not have fully loaded: {e}")

    async def _capture_sheet(self, page: Page, sheet: SheetInfo) -> ScreenshotInfo | None:
        """Capture a screenshot of a single sheet."""
        try:
            # Handle hidden sheets
            was_hidden = False
            if sheet.visibility == SheetVisibility.HIDDEN:
                success = await self._unhide_sheet(page, sheet.name)
                if success:
                    was_hidden = True
                else:
                    print(f"Could not unhide sheet: {sheet.name}")
                    return None
            elif sheet.visibility == SheetVisibility.VERY_HIDDEN:
                # Very hidden sheets cannot be unhidden via UI
                print(f"Sheet '{sheet.name}' is very hidden (requires VBA to view)")
                return None

            # Navigate to the sheet
            await self._navigate_to_sheet(page, sheet.name)

            # Wait for sheet to render
            await asyncio.sleep(1)
            await page.wait_for_load_state("networkidle")

            # Take screenshot
            screenshot_path = self.output_dir / f"{self._sanitize_filename(sheet.name)}.png"

            # Try to capture just the spreadsheet area
            try:
                spreadsheet = await page.query_selector(self.SELECTORS["spreadsheet_canvas"])
                if spreadsheet:
                    await spreadsheet.screenshot(path=str(screenshot_path))
                else:
                    # Fallback to full page
                    await page.screenshot(path=str(screenshot_path), full_page=True)
            except Exception:
                # Fallback to full page
                await page.screenshot(path=str(screenshot_path), full_page=True)

            # Re-hide sheet if it was hidden
            if was_hidden:
                await self._hide_sheet(page, sheet.name)

            return ScreenshotInfo(
                sheet=sheet.name,
                path=screenshot_path,
                captured_at=datetime.now().isoformat(),
            )

        except Exception as e:
            print(f"Error capturing sheet '{sheet.name}': {e}")
            return None

    async def _navigate_to_sheet(self, page: Page, sheet_name: str) -> None:
        """Navigate to a specific sheet tab."""
        try:
            # Try to find and click the sheet tab
            # First try by data attribute
            selector = self.SELECTORS["sheet_tab"].format(name=sheet_name)
            tab = await page.query_selector(selector)

            if not tab:
                # Try by text content
                tab = await page.query_selector(f'[role="tab"]:has-text("{sheet_name}")')

            if not tab:
                # Try in sheet tab bar
                tabs_container = await page.query_selector(self.SELECTORS["sheet_tabs_container"])
                if tabs_container:
                    tab = await tabs_container.query_selector(f'text="{sheet_name}"')

            if tab:
                await tab.click()
                await asyncio.sleep(0.5)
            else:
                print(f"Could not find tab for sheet: {sheet_name}")

        except Exception as e:
            print(f"Error navigating to sheet '{sheet_name}': {e}")

    async def _unhide_sheet(self, page: Page, sheet_name: str) -> bool:
        """Attempt to unhide a hidden sheet via Excel Online UI."""
        try:
            # Right-click on any visible sheet tab to get context menu
            active_tab = await page.query_selector(self.SELECTORS["active_sheet_tab"])
            if active_tab:
                await active_tab.click(button="right")
                await asyncio.sleep(0.5)

                # Look for Unhide option
                unhide_button = await page.query_selector(self.SELECTORS["unhide_sheets_item"])
                if unhide_button:
                    await unhide_button.click()
                    await asyncio.sleep(0.5)

                    # Find and select the sheet in the unhide dialog
                    dialog = await page.query_selector(self.SELECTORS["unhide_dialog"])
                    if dialog:
                        sheet_option = await dialog.query_selector(f'text="{sheet_name}"')
                        if sheet_option:
                            await sheet_option.click()

                            # Click OK/Unhide button
                            ok_button = await dialog.query_selector('button:has-text("OK"), button:has-text("Unhide")')
                            if ok_button:
                                await ok_button.click()
                                await asyncio.sleep(1)
                                return True

                # Close any open menu
                await page.keyboard.press("Escape")

        except Exception as e:
            print(f"Error unhiding sheet: {e}")

        return False

    async def _hide_sheet(self, page: Page, sheet_name: str) -> bool:
        """Re-hide a sheet after capturing."""
        try:
            # Navigate to the sheet first
            await self._navigate_to_sheet(page, sheet_name)
            await asyncio.sleep(0.3)

            # Right-click on the sheet tab
            selector = self.SELECTORS["sheet_tab"].format(name=sheet_name)
            tab = await page.query_selector(selector)

            if not tab:
                tab = await page.query_selector(f'[role="tab"]:has-text("{sheet_name}")')

            if tab:
                await tab.click(button="right")
                await asyncio.sleep(0.5)

                # Look for Hide option
                hide_button = await page.query_selector('button:has-text("Hide")')
                if hide_button:
                    await hide_button.click()
                    await asyncio.sleep(0.5)
                    return True

                # Close menu
                await page.keyboard.press("Escape")

        except Exception as e:
            print(f"Error hiding sheet: {e}")

        return False

    def _sanitize_filename(self, name: str) -> str:
        """Sanitize a string for use as a filename."""
        # Replace problematic characters
        invalid_chars = '<>:"/\\|?*'
        result = name
        for char in invalid_chars:
            result = result.replace(char, "_")

        # Trim length
        if len(result) > 100:
            result = result[:100]

        return result


async def capture_screenshots(
    context: BrowserContext,
    sharepoint_url: str,
    sheets: list[SheetInfo],
    output_dir: Path,
) -> list[ScreenshotInfo]:
    """Convenience function to capture screenshots.

    Args:
        context: Authenticated browser context
        sharepoint_url: SharePoint URL to Excel file
        sheets: List of sheet info
        output_dir: Where to save screenshots

    Returns:
        List of ScreenshotInfo objects
    """
    screenshotter = ExcelOnlineScreenshotter(output_dir)
    return await screenshotter.capture_all_sheets(context, sharepoint_url, sheets)
