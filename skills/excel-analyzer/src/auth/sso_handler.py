"""M365/SharePoint SSO authentication handler using Playwright."""

from __future__ import annotations

import asyncio
import re
from urllib.parse import urlparse

from playwright.async_api import Browser, BrowserContext, Page

from .session_store import SessionStore


class SSOHandler:
    """Handles M365 SSO authentication via interactive browser login."""

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

    async def get_authenticated_context(
        self,
        browser: Browser,
        sharepoint_url: str,
        timeout_ms: int = 300000,
    ) -> BrowserContext:
        """Get an authenticated browser context for SharePoint.

        Uses cached session if available, otherwise prompts for interactive login.
        The browser must be provided and managed by the caller.

        Args:
            browser: Browser instance (caller manages lifecycle)
            sharepoint_url: The SharePoint URL
            timeout_ms: Timeout for interactive login if needed

        Returns:
            Authenticated browser context
        """
        domain = self._extract_domain(sharepoint_url)
        state_path = self.session_store.get_state_path(domain)

        # Try cached session first
        if self.session_store.has_valid_session(domain) and state_path.exists():
            try:
                context = await browser.new_context(storage_state=str(state_path))
                page = await context.new_page()

                # Validate session still works
                await page.goto(sharepoint_url, timeout=60000)
                await page.wait_for_load_state("domcontentloaded", timeout=30000)

                if not self._is_login_page(page.url):
                    print(f"Using cached session for {domain}", flush=True)
                    await page.close()
                    return context

                # Session expired
                await context.close()
                print(f"Cached session expired for {domain}, re-authenticating...", flush=True)
            except Exception as e:
                print(f"Cached session invalid for {domain}: {e}", flush=True)
                try:
                    await context.close()
                except Exception:
                    pass

        # Need interactive login
        return await self._interactive_login(browser, sharepoint_url, domain, timeout_ms)

    async def _interactive_login(
        self,
        browser: Browser,
        sharepoint_url: str,
        domain: str,
        timeout_ms: int,
    ) -> BrowserContext:
        """Perform interactive login."""
        state_path = self.session_store.get_state_path(domain)

        context = await browser.new_context()
        page = await context.new_page()

        print(f"\n{'='*60}", flush=True)
        print("🔐 SharePoint Authentication Required", flush=True)
        print(f"{'='*60}", flush=True)
        print(f"A browser window has opened. Please log in to SharePoint.", flush=True)
        print(f"Domain: {domain}", flush=True)
        print(f"You have 5 minutes to complete login.", flush=True)
        print(f"{'='*60}\n", flush=True)

        await page.goto(sharepoint_url)

        # Wait for authentication
        await self._wait_for_auth(page, domain, timeout_ms)

        # Save session
        await context.storage_state(path=str(state_path))
        self.session_store.save_session(domain, {"url": sharepoint_url})

        print(f"\n✓ Authentication successful! Session saved for {domain}", flush=True)
        await page.close()

        return context

    async def _wait_for_auth(self, page: Page, domain: str, timeout_ms: int) -> None:
        """Wait for user to complete authentication."""
        start_time = asyncio.get_event_loop().time()

        while True:
            elapsed = (asyncio.get_event_loop().time() - start_time) * 1000
            if elapsed > timeout_ms:
                raise TimeoutError("Authentication timed out. Please try again.")

            current_url = page.url

            # Check if we've reached SharePoint successfully
            if domain in current_url and not self._is_login_page(current_url):
                try:
                    await page.wait_for_load_state("networkidle", timeout=10000)
                except Exception:
                    pass

                # Look for indicators we're authenticated
                try:
                    await page.wait_for_selector(
                        "[data-automation-id='mectrl_main_trigger'], #O365_MainLink_Me, .ms-Persona, #appLauncherTop",
                        timeout=5000
                    )
                    return
                except Exception:
                    # May be authenticated without visible user element
                    if domain in page.url and not self._is_login_page(page.url):
                        return

            await asyncio.sleep(0.5)

    def clear_session(self, sharepoint_url: str) -> None:
        """Clear cached session for a SharePoint domain."""
        domain = self._extract_domain(sharepoint_url)
        self.session_store.clear_session(domain)
        print(f"Cleared session for {domain}", flush=True)
