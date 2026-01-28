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
        # Standard sharing link: /:x:/r/sites/SiteName/Shared Documents/file.xlsx
        r"/sites/([^/]+)/([^?]+\.xlsx?)(?:\?|$)",
        # Personal OneDrive: /personal/user_company_com/Documents/file.xlsx
        r"/personal/([^/]+)/([^?]+\.xlsx?)(?:\?|$)",
        # Direct link with file parameter
        r"[?&]file=([^&]+\.xlsx?)",
    ]

    def __init__(self, session_store: SessionStore | None = None):
        """Initialize the downloader.

        Args:
            session_store: Session store for auth persistence
        """
        self.session_store = session_store or SessionStore()
        self.sso_handler = SSOHandler(self.session_store)

    def _extract_filename(self, url: str) -> str:
        """Extract filename from SharePoint URL.

        Args:
            url: SharePoint URL

        Returns:
            Extracted filename or default name
        """
        # Try URL patterns
        for pattern in self.URL_PATTERNS:
            match = re.search(pattern, url, re.IGNORECASE)
            if match:
                # Get the last group which should be the path or filename
                path_or_file = match.group(match.lastindex)
                # Extract just the filename
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
        """Check if URL is a SharePoint/OneDrive URL.

        Args:
            url: URL to check

        Returns:
            True if URL appears to be SharePoint/OneDrive
        """
        sharepoint_domains = [
            "sharepoint.com",
            "sharepoint.us",
            "onedrive.com",
            "office.com",
        ]
        parsed = urlparse(url)
        return any(domain in parsed.netloc.lower() for domain in sharepoint_domains)

    async def _get_download_url(self, page: Page, original_url: str) -> str:
        """Get the direct download URL for a SharePoint file.

        SharePoint URLs need to be converted to download URLs.

        Args:
            page: Authenticated browser page
            original_url: Original SharePoint URL

        Returns:
            Direct download URL
        """
        # Navigate to the file
        await page.goto(original_url)
        await page.wait_for_load_state("networkidle")

        # Try to find and click download button
        download_selectors = [
            'button[aria-label*="Download"]',
            'button[data-automationid="downloadCommand"]',
            '[data-automation-id="downloadCommand"]',
            'button:has-text("Download")',
        ]

        for selector in download_selectors:
            try:
                download_btn = await page.wait_for_selector(selector, timeout=5000)
                if download_btn:
                    # Instead of clicking, get the download URL
                    # SharePoint often uses a REST API for downloads
                    break
            except Exception:
                continue

        # Construct download URL from the file URL
        # SharePoint REST API pattern: /_api/web/GetFileByServerRelativeUrl('path')/\$value
        current_url = page.url
        parsed = urlparse(current_url)

        # Try to extract file path for API download
        if "/sites/" in current_url or "/personal/" in current_url:
            # Extract the relative path
            match = re.search(r"(/sites/[^?]+\.xlsx?|/personal/[^?]+\.xlsx?)", current_url, re.IGNORECASE)
            if match:
                file_path = unquote(match.group(1))
                download_url = f"{parsed.scheme}://{parsed.netloc}/_api/web/GetFileByServerRelativeUrl('{file_path}')/$value"
                return download_url

        # Fallback: try to modify URL to trigger download
        if "?" in original_url:
            return original_url + "&download=1"
        else:
            return original_url + "?download=1"

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

        output_path = output_dir / filename

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)

            try:
                # Get authenticated context
                context = await self.sso_handler.get_authenticated_context(url, browser)
                page = await context.new_page()

                # Navigate to file
                print(f"Navigating to SharePoint file...")
                await page.goto(url)
                await page.wait_for_load_state("networkidle")

                # Set up download handler
                async with page.expect_download(timeout=60000) as download_info:
                    # Try clicking download button
                    download_clicked = False
                    download_selectors = [
                        'button[aria-label*="Download"]',
                        'button[data-automationid="downloadCommand"]',
                        '[data-automation-id="downloadCommand"]',
                        'button:has-text("Download")',
                        '[aria-label="Download"]',
                    ]

                    for selector in download_selectors:
                        try:
                            btn = await page.wait_for_selector(selector, timeout=3000)
                            if btn:
                                await btn.click()
                                download_clicked = True
                                break
                        except Exception:
                            continue

                    if not download_clicked:
                        # Try keyboard shortcut or menu
                        try:
                            # Open File menu
                            await page.keyboard.press("Alt+F")
                            await asyncio.sleep(0.5)
                            # Look for Save As or Download option
                            await page.click('text="Download"', timeout=3000)
                        except Exception:
                            # Try direct download URL
                            download_url = await self._get_download_url(page, url)
                            await page.goto(download_url)

                download = await download_info.value
                await download.save_as(output_path)

                print(f"Downloaded: {output_path}")
                return output_path

            finally:
                await browser.close()

    async def download_with_context(
        self,
        url: str,
        context: BrowserContext,
        output_dir: Path | None = None,
        filename: str | None = None,
    ) -> Path:
        """Download using an existing authenticated context.

        Args:
            url: SharePoint URL
            context: Pre-authenticated browser context
            output_dir: Output directory
            filename: Override filename

        Returns:
            Path to downloaded file
        """
        if output_dir is None:
            output_dir = Path(tempfile.mkdtemp(prefix="excel_analyzer_"))
        output_dir.mkdir(parents=True, exist_ok=True)

        if filename is None:
            filename = self._extract_filename(url)

        output_path = output_dir / filename
        page = await context.new_page()

        try:
            await page.goto(url)
            await page.wait_for_load_state("networkidle")

            async with page.expect_download(timeout=60000) as download_info:
                # Click download button
                download_selectors = [
                    'button[aria-label*="Download"]',
                    'button[data-automationid="downloadCommand"]',
                    '[data-automation-id="downloadCommand"]',
                    'button:has-text("Download")',
                ]

                for selector in download_selectors:
                    try:
                        btn = await page.wait_for_selector(selector, timeout=3000)
                        if btn:
                            await btn.click()
                            break
                    except Exception:
                        continue

            download = await download_info.value
            await download.save_as(output_path)

            return output_path

        finally:
            await page.close()


def download_from_sharepoint(url: str, output_dir: Path | None = None) -> Path:
    """Synchronous wrapper for downloading from SharePoint.

    Args:
        url: SharePoint URL
        output_dir: Output directory

    Returns:
        Path to downloaded file
    """
    downloader = SharePointDownloader()
    return asyncio.run(downloader.download(url, output_dir))
