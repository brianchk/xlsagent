"""M365/SharePoint SSO authentication handler using Playwright."""

from __future__ import annotations

import asyncio
import re
from pathlib import Path
from urllib.parse import urlparse

from playwright.async_api import Browser, BrowserContext, Page, async_playwright

from .session_store import SessionStore


class SSOHandler:
    """Handles M365 SSO authentication via interactive browser login."""

    # URLs that indicate successful SharePoint authentication
    SHAREPOINT_SUCCESS_PATTERNS = [
        r".*\.sharepoint\.com/.*",
        r".*\.sharepoint\.com/_layouts/.*",
    ]

    # URLs that indicate login pages
    LOGIN_URL_PATTERNS = [
        r"login\.microsoftonline\.com",
        r"login\.live\.com",
        r"adfs\.",
        r"federation",
    ]

    def __init__(self, session_store: SessionStore | None = None):
        """Initialize SSO handler.

        Args:
            session_store: Session store for persisting auth state
        """
        self.session_store = session_store or SessionStore()

    def _extract_domain(self, url: str) -> str:
        """Extract the SharePoint domain from a URL."""
        parsed = urlparse(url)
        return parsed.netloc

    def _is_login_page(self, url: str) -> bool:
        """Check if URL is a login page."""
        return any(re.search(pattern, url) for pattern in self.LOGIN_URL_PATTERNS)

    def _is_sharepoint_authenticated(self, url: str) -> bool:
        """Check if URL indicates successful SharePoint authentication."""
        return any(re.search(pattern, url) for pattern in self.SHAREPOINT_SUCCESS_PATTERNS)

    async def authenticate(
        self,
        sharepoint_url: str,
        headless: bool = False,
        timeout_ms: int = 300000,  # 5 minutes for user to complete login
    ) -> BrowserContext:
        """Authenticate to SharePoint via interactive browser login.

        Opens a browser window for the user to complete M365 SSO authentication.
        Once authenticated, saves the session for reuse.

        Args:
            sharepoint_url: The SharePoint URL to authenticate for
            headless: If True, runs headless (for testing). Usually False for interactive login.
            timeout_ms: Timeout for user to complete login

        Returns:
            Authenticated browser context
        """
        domain = self._extract_domain(sharepoint_url)
        state_path = self.session_store.get_state_path(domain)

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=headless)

            # Check for existing session
            if self.session_store.has_valid_session(domain) and state_path.exists():
                try:
                    context = await browser.new_context(storage_state=str(state_path))
                    page = await context.new_page()

                    # Verify session still works
                    await page.goto(sharepoint_url, timeout=30000)
                    await page.wait_for_load_state("networkidle")

                    if not self._is_login_page(page.url):
                        print(f"Using cached session for {domain}")
                        return context

                    # Session expired, close and re-auth
                    await context.close()
                    print(f"Cached session expired for {domain}, re-authenticating...")
                except Exception:
                    print(f"Cached session invalid for {domain}, re-authenticating...")

            # Need interactive login
            context = await browser.new_context()
            page = await context.new_page()

            print(f"\nOpening browser for M365 authentication...")
            print(f"Please complete the login process in the browser window.")
            print(f"URL: {sharepoint_url}\n")

            await page.goto(sharepoint_url)

            # Wait for successful authentication
            try:
                await self._wait_for_auth(page, domain, timeout_ms)
            except Exception as e:
                await context.close()
                await browser.close()
                raise RuntimeError(f"Authentication failed: {e}") from e

            # Save session state
            await context.storage_state(path=str(state_path))
            self.session_store.save_session(domain, {"url": sharepoint_url})

            print(f"\nAuthentication successful! Session saved for {domain}")

            return context

    async def _wait_for_auth(self, page: Page, domain: str, timeout_ms: int) -> None:
        """Wait for user to complete authentication.

        Args:
            page: The browser page
            domain: Expected SharePoint domain
            timeout_ms: Maximum time to wait
        """
        start_time = asyncio.get_event_loop().time()

        while True:
            elapsed = (asyncio.get_event_loop().time() - start_time) * 1000
            if elapsed > timeout_ms:
                raise TimeoutError("Authentication timed out. Please try again.")

            current_url = page.url

            # Check if we've reached SharePoint successfully
            if domain in current_url and not self._is_login_page(current_url):
                # Wait a bit more for page to fully load
                try:
                    await page.wait_for_load_state("networkidle", timeout=10000)
                except Exception:
                    pass  # May already be idle

                # Verify we're actually authenticated (look for user indicator)
                try:
                    # SharePoint typically shows user info in the header
                    await page.wait_for_selector(
                        "[data-automation-id='mectrl_main_trigger'], #O365_MainLink_Me, .ms-Persona",
                        timeout=5000
                    )
                    return
                except Exception:
                    # May be authenticated without visible user element, check URL again
                    if domain in page.url and not self._is_login_page(page.url):
                        return

            await asyncio.sleep(0.5)

    async def get_authenticated_context(
        self,
        sharepoint_url: str,
        browser: Browser | None = None,
    ) -> BrowserContext:
        """Get an authenticated browser context for SharePoint.

        Uses cached session if available, otherwise prompts for interactive login.

        Args:
            sharepoint_url: The SharePoint URL
            browser: Optional existing browser instance to use

        Returns:
            Authenticated browser context
        """
        domain = self._extract_domain(sharepoint_url)
        state_path = self.session_store.get_state_path(domain)

        # If we have a valid session, use it
        if self.session_store.has_valid_session(domain) and state_path.exists():
            if browser is None:
                async with async_playwright() as p:
                    browser = await p.chromium.launch(headless=True)
                    context = await browser.new_context(storage_state=str(state_path))

                    # Quick validation
                    page = await context.new_page()
                    await page.goto(sharepoint_url, timeout=30000)

                    if not self._is_login_page(page.url):
                        return context

                    await context.close()
            else:
                context = await browser.new_context(storage_state=str(state_path))
                page = await context.new_page()
                await page.goto(sharepoint_url, timeout=30000)

                if not self._is_login_page(page.url):
                    return context

                await context.close()

        # Need to re-authenticate
        return await self.authenticate(sharepoint_url, headless=False)

    def clear_session(self, sharepoint_url: str) -> None:
        """Clear cached session for a SharePoint domain.

        Args:
            sharepoint_url: URL of the SharePoint site
        """
        domain = self._extract_domain(sharepoint_url)
        self.session_store.clear_session(domain)
        print(f"Cleared session for {domain}")
