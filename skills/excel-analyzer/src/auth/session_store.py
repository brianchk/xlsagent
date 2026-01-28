"""Session storage for M365/SharePoint authentication."""

from __future__ import annotations

import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any


class SessionStore:
    """Manages persistent browser session storage for M365 authentication."""

    def __init__(self, storage_dir: Path | None = None):
        """Initialize session store.

        Args:
            storage_dir: Directory to store session data. Defaults to ~/.claude/excel-analyzer/sessions/
        """
        if storage_dir is None:
            storage_dir = Path.home() / ".claude" / "excel-analyzer" / "sessions"
        self.storage_dir = storage_dir
        self.storage_dir.mkdir(parents=True, exist_ok=True)

    def _get_session_path(self, domain: str) -> Path:
        """Get path for domain-specific session file."""
        # Sanitize domain for filename
        safe_domain = domain.replace(".", "_").replace("/", "_")
        return self.storage_dir / f"{safe_domain}_session.json"

    def _get_state_path(self, domain: str) -> Path:
        """Get path for Playwright state file."""
        safe_domain = domain.replace(".", "_").replace("/", "_")
        return self.storage_dir / f"{safe_domain}_state.json"

    def get_state_path(self, domain: str) -> Path:
        """Get the Playwright browser state path for a domain."""
        return self._get_state_path(domain)

    def has_valid_session(self, domain: str, max_age_hours: int = 8) -> bool:
        """Check if a valid session exists for the given domain.

        Args:
            domain: The SharePoint domain (e.g., 'company.sharepoint.com')
            max_age_hours: Maximum age of session in hours before considered expired

        Returns:
            True if a valid, non-expired session exists
        """
        session_path = self._get_session_path(domain)
        state_path = self._get_state_path(domain)

        if not session_path.exists() or not state_path.exists():
            return False

        try:
            with open(session_path) as f:
                session_data = json.load(f)

            created_at = datetime.fromisoformat(session_data.get("created_at", ""))
            if datetime.now() - created_at > timedelta(hours=max_age_hours):
                return False

            return True
        except (json.JSONDecodeError, ValueError, KeyError):
            return False

    def save_session(self, domain: str, metadata: dict[str, Any] | None = None) -> None:
        """Save session metadata after successful authentication.

        Note: The actual browser state is saved by Playwright's storage_state method.
        This saves additional metadata about the session.

        Args:
            domain: The SharePoint domain
            metadata: Additional metadata to store with the session
        """
        session_path = self._get_session_path(domain)

        session_data = {
            "domain": domain,
            "created_at": datetime.now().isoformat(),
            "metadata": metadata or {},
        }

        with open(session_path, "w") as f:
            json.dump(session_data, f, indent=2)

    def clear_session(self, domain: str) -> None:
        """Clear session for a domain.

        Args:
            domain: The SharePoint domain to clear
        """
        session_path = self._get_session_path(domain)
        state_path = self._get_state_path(domain)

        if session_path.exists():
            session_path.unlink()
        if state_path.exists():
            state_path.unlink()

    def clear_all_sessions(self) -> None:
        """Clear all stored sessions."""
        for path in self.storage_dir.glob("*_session.json"):
            path.unlink()
        for path in self.storage_dir.glob("*_state.json"):
            path.unlink()

    def list_sessions(self) -> list[dict[str, Any]]:
        """List all stored sessions with their metadata.

        Returns:
            List of session info dictionaries
        """
        sessions = []
        for path in self.storage_dir.glob("*_session.json"):
            try:
                with open(path) as f:
                    data = json.load(f)
                    data["valid"] = self.has_valid_session(data.get("domain", ""))
                    sessions.append(data)
            except (json.JSONDecodeError, KeyError):
                continue
        return sessions
