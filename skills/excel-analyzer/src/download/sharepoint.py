"""SharePoint file downloader using authenticated browser session."""

from __future__ import annotations

import asyncio
import re
import tempfile
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from playwright.async_api import BrowserContext, Page, async_playwright

from ..auth import SessionStore, SSOHandler


class SharePointDownloader:
    """Downloads Excel files from SharePoint using browser authentication."""

    # Patterns to extract file info from SharePoint URLs
    URL_PATTERNS = [
        # Shared Documents pattern
        r"/Shared%20Documents/([^?]+\.xlsm?)(?:\?|$)",
        r"/Shared Documents/([^?]+\.xlsm?)(?:\?|$)",
        # Standard sharing link: /:x:/r/sites/SiteName/Shared Documents/file.xlsx
        r"/sites/([^/]+)/([^?]+\.xlsx?)(?:\?|$)",
        # Personal OneDrive: /personal/user_company_com/Documents/file.xlsx
        r"/personal/([^/]+)/([^?]+\.xlsx?)(?:\?|$)",
        # Direct link with file parameter
        r"[?&]file=([^&]+\.xlsx?)",
    ]

    def __init__(self, session_store: SessionStore | None = None):
        """Initialize the downloader."""
        self.session_store = session_store or SessionStore()
        self.sso_handler = SSOHandler(self.session_store)

    def _extract_filename(self, url: str) -> str:
        """Extract filename from SharePoint URL."""
        # Try URL patterns
        for pattern in self.URL_PATTERNS:
            match = re.search(pattern, url, re.IGNORECASE)
            if match:
                path_or_file = match.group(match.lastindex)
                filename = Path(unquote(path_or_file)).name
                if filename:
                    return filename

        # Try query parameters
        parsed = urlparse(url)
        params = parse_qs(parsed.query)
        if "file" in params:
            return Path(unquote(params["file"][0])).name

        # Default fallback
        return "workbook.xlsx"

    def _is_sharepoint_url(self, url: str) -> bool:
        """Check if URL is a SharePoint/OneDrive URL."""
        sharepoint_domains = [
            "sharepoint.com",
            "sharepoint.us",
            "onedrive.com",
            "office.com",
        ]
        parsed = urlparse(url)
        return any(domain in parsed.netloc.lower() for domain in sharepoint_domains)

    async def download(
        self,
        url: str,
        output_dir: Path | None = None,
        filename: str | None = None,
    ) -> Path:
        """Download an Excel file from SharePoint.

        Args:
            url: SharePoint URL to the Excel file
            output_dir: Directory to save the file. Defaults to temp directory.
            filename: Override filename. Defaults to extracting from URL.

        Returns:
            Path to downloaded file
        """
        if not self._is_sharepoint_url(url):
            raise ValueError(f"URL does not appear to be a SharePoint/OneDrive URL: {url}")

        if output_dir is None:
            output_dir = Path(tempfile.mkdtemp(prefix="excel_analyzer_"))
        output_dir.mkdir(parents=True, exist_ok=True)

        if filename is None:
            filename = self._extract_filename(url)

        # Ensure proper extension
        if not filename.lower().endswith(('.xlsx', '.xlsm', '.xlsb')):
            filename = filename + '.xlsm'

        output_path = output_dir / filename

        print(f"Downloading: {filename}", flush=True)
        print(f"From: {url[:80]}...", flush=True)

        async with async_playwright() as p:
            # Launch browser - visible for auth if needed
            browser = await p.chromium.launch(headless=False)

            try:
                # Get authenticated context
                context = await self.sso_handler.get_authenticated_context(browser, url)

                # Download the file
                await self._download_file(context, url, output_path)

                print(f"✓ Downloaded: {output_path}", flush=True)
                return output_path

            finally:
                await browser.close()

    async def _download_file(
        self,
        context: BrowserContext,
        url: str,
        output_path: Path,
    ) -> None:
        """Download file using the authenticated context."""
        page = await context.new_page()

        try:
            # Navigate to the file
            print("Opening file in Excel Online...", flush=True)
            await page.goto(url, timeout=60000)
            # Wait for initial load, but don't wait for networkidle as Excel Online
            # continuously makes requests. Instead wait for domcontentloaded then
            # give it a moment to initialize.
            await page.wait_for_load_state("domcontentloaded", timeout=60000)

            # Wait for Excel Online to fully initialize
            await asyncio.sleep(5)

            # Debug: print current URL and save screenshot
            print(f"Current URL: {page.url}", flush=True)
            debug_screenshot = output_path.parent / "debug_page.png"
            await page.screenshot(path=str(debug_screenshot))
            print(f"Debug screenshot saved: {debug_screenshot}", flush=True)

            # Try to click download button
            downloaded = await self._try_download_button(page, output_path)

            if not downloaded:
                # Try File menu download via locators
                downloaded = await self._try_file_menu_download(page, output_path)

            if not downloaded:
                # Try keyboard shortcut (Alt+F or File menu via keyboard)
                downloaded = await self._try_keyboard_download(page, output_path)

            if not downloaded:
                # Try direct download URL
                await self._try_direct_download(page, url, output_path)

        finally:
            await page.close()

    async def _try_download_button(self, page: Page, output_path: Path) -> bool:
        """Try to download using the download button."""
        download_selectors = [
            # SharePoint document library download button
            'button[data-automationid="downloadCommand"]',
            '[data-automation-id="downloadCommand"]',
            # Excel Online toolbar
            'button[aria-label*="Download"]',
            'button[aria-label*="download"]',
            '[aria-label="Download"]',
            '[aria-label="Download a Copy"]',
            # SharePoint command bar
            'button[name="Download"]',
            'button[name="download"]',
            # Generic download icons
            'button:has([data-icon-name="Download"])',
            '[data-icon-name="Download"]',
        ]

        print("Trying download button selectors...", flush=True)
        for selector in download_selectors:
            try:
                btn = await page.wait_for_selector(selector, timeout=2000)
                if btn and await btn.is_visible():
                    print(f"  Found download button with selector: {selector}", flush=True)

                    async with page.expect_download(timeout=60000) as download_info:
                        await btn.click()

                    download = await download_info.value
                    await download.save_as(output_path)
                    return True
            except Exception:
                pass

        print("  No download button found", flush=True)
        return False

    async def _try_file_menu_download(self, page: Page, output_path: Path) -> bool:
        """Try to download via File menu."""
        try:
            # Try clicking File tab/menu in ribbon
            file_menu_selectors = [
                # Excel Online ribbon File tab
                '#FileTabButton',
                'button[id*="FileTabButton"]',
                '[data-automation-id="FileMenu"]',
                'button[aria-label="File"]',
                'a[aria-label="File"]',
                '#jewel-menu-button',
                # Text-based fallback
                'button:has-text("File")',
                'a:has-text("File"):first',
            ]

            for selector in file_menu_selectors:
                try:
                    print(f"  Trying File menu selector: {selector}", flush=True)
                    menu = await page.wait_for_selector(selector, timeout=3000)
                    if menu and await menu.is_visible():
                        print("  Found File menu, clicking...", flush=True)
                        await menu.click()
                        await asyncio.sleep(1)

                        # Look for download option in the backstage/file menu
                        download_selectors = [
                            '[data-automationid="SaveAsBtn"]',
                            'button[aria-label*="Download"]',
                            'button[aria-label*="download"]',
                            'a[aria-label*="Download"]',
                            'button:has-text("Download")',
                            'a:has-text("Download")',
                            'button:has-text("Save As")',
                            '[data-automationid*="Download"]',
                        ]

                        for dl_selector in download_selectors:
                            try:
                                dl_btn = await page.wait_for_selector(dl_selector, timeout=2000)
                                if dl_btn and await dl_btn.is_visible():
                                    print(f"  Found download option, clicking...", flush=True)
                                    async with page.expect_download(timeout=60000) as download_info:
                                        await dl_btn.click()
                                    download = await download_info.value
                                    await download.save_as(output_path)
                                    return True
                            except Exception:
                                continue

                        # If we got here, close the menu
                        await page.keyboard.press("Escape")
                except Exception:
                    continue

            # Close any open menu
            await page.keyboard.press("Escape")

        except Exception as e:
            print(f"  File menu download failed: {e}", flush=True)

        return False

    async def _try_keyboard_download(self, page: Page, output_path: Path) -> bool:
        """Try to download using keyboard shortcuts."""
        print("Trying keyboard shortcuts for download...", flush=True)
        try:
            # Try Alt+F to open File menu (works in some Office apps)
            await page.keyboard.press("Alt+F")
            await asyncio.sleep(1)

            # Try to find and click download option
            try:
                async with page.expect_download(timeout=30000) as download_info:
                    # Look for download text and click
                    download_link = page.locator("text=Download").first
                    if await download_link.is_visible():
                        await download_link.click()
                    else:
                        # Try Save As
                        save_link = page.locator("text=Save As").first
                        if await save_link.is_visible():
                            await save_link.click()
                        else:
                            await page.keyboard.press("Escape")
                            return False

                download = await download_info.value
                await download.save_as(output_path)
                return True
            except Exception:
                await page.keyboard.press("Escape")

            # Try clicking on File text in the ribbon using coordinates
            # Based on screenshot, File is at approx (23, 44) from top-left
            print("  Trying click on File ribbon tab by text...", flush=True)
            try:
                file_locator = page.get_by_role("tab", name="File")
                if await file_locator.is_visible(timeout=2000):
                    await file_locator.click()
                    await asyncio.sleep(1)

                    async with page.expect_download(timeout=30000) as download_info:
                        download_locator = page.get_by_text("Download", exact=False).first
                        await download_locator.click()

                    download = await download_info.value
                    await download.save_as(output_path)
                    return True
            except Exception:
                pass

        except Exception as e:
            print(f"  Keyboard download failed: {e}", flush=True)

        await page.keyboard.press("Escape")
        return False

    async def _try_direct_download(
        self,
        page: Page,
        url: str,
        output_path: Path,
    ) -> None:
        """Try direct download via modified URL."""
        print("Trying direct download...", flush=True)

        parsed = urlparse(url)
        domain = f"{parsed.scheme}://{parsed.netloc}"

        # Try to extract file path from the ORIGINAL sharing URL
        # Pattern: /:x:/r/Shared Documents/path/file.xlsx or /sites/SiteName/Shared Documents/...
        original_patterns = [
            r"/:x:/r(/Shared%20Documents/[^?]+)",
            r"/:x:/r(/sites/[^/]+/Shared%20Documents/[^?]+)",
            r"/r(/Shared%20Documents/[^?]+)",
            r"/r(/sites/[^/]+/Shared%20Documents/[^?]+)",
        ]

        file_path = None
        for pattern in original_patterns:
            match = re.search(pattern, url, re.IGNORECASE)
            if match:
                file_path = unquote(match.group(1))
                print(f"  Extracted file path from URL: {file_path}", flush=True)
                break

        # Also try from current URL
        if not file_path:
            current_url = page.url
            patterns = [
                r"(/Shared%20Documents/[^?]+)",
                r"(/sites/[^/]+/Shared%20Documents/[^?]+)",
                r"(/personal/[^/]+/Documents/[^?]+)",
            ]
            for pattern in patterns:
                match = re.search(pattern, current_url, re.IGNORECASE)
                if match:
                    file_path = unquote(match.group(1))
                    print(f"  Extracted file path from current URL: {file_path}", flush=True)
                    break

        if file_path:
            # Try REST API download
            download_url = f"{domain}/_api/web/GetFileByServerRelativeUrl('{file_path}')/$value"
            print(f"  Trying REST API: {download_url[:80]}...", flush=True)

            try:
                async with page.expect_download(timeout=60000) as download_info:
                    await page.goto(download_url)

                download = await download_info.value
                await download.save_as(output_path)
                return
            except Exception as e:
                print(f"  REST API download failed: {e}", flush=True)

            # Try direct file URL
            direct_url = f"{domain}{file_path}"
            print(f"  Trying direct URL: {direct_url[:80]}...", flush=True)

            try:
                async with page.expect_download(timeout=60000) as download_info:
                    await page.goto(direct_url)

                download = await download_info.value
                await download.save_as(output_path)
                return
            except Exception as e:
                print(f"  Direct URL download failed: {e}", flush=True)

        # Last resort: try adding download parameter
        if "?" in url:
            download_url = url + "&download=1"
        else:
            download_url = url + "?download=1"

        print(f"  Trying download=1 parameter...", flush=True)
        try:
            async with page.expect_download(timeout=60000) as download_info:
                await page.goto(download_url)

            download = await download_info.value
            await download.save_as(output_path)
        except Exception as e:
            raise RuntimeError(
                f"Could not download file. Please download manually and provide local path.\n"
                f"Error: {e}"
            )


def download_from_sharepoint(url: str, output_dir: Path | None = None) -> Path:
    """Synchronous wrapper for downloading from SharePoint."""
    downloader = SharePointDownloader()
    return asyncio.run(downloader.download(url, output_dir))
