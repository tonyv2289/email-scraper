"""
Microsoft Graph API Authentication Module

Handles OAuth2 authentication with Microsoft Graph API using MSAL.
Supports interactive login for personal/work accounts.
"""

import json
import os
from pathlib import Path
from typing import Optional

import msal
from rich.console import Console

console = Console()

# Microsoft Graph API scopes needed for email access
SCOPES = [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Mail.ReadBasic",
    "https://graph.microsoft.com/User.Read",
]

# Token cache file location
TOKEN_CACHE_FILE = Path.home() / ".outlook_scraper_token_cache.json"


class GraphAuthenticator:
    """Handles authentication with Microsoft Graph API."""

    def __init__(self, client_id: str, tenant_id: str = "common"):
        """
        Initialize the authenticator.

        Args:
            client_id: Azure AD application (client) ID
            tenant_id: Azure AD tenant ID (use "common" for multi-tenant)
        """
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self._token_cache = msal.SerializableTokenCache()
        self._load_token_cache()

        self.app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            token_cache=self._token_cache,
        )

    def _load_token_cache(self) -> None:
        """Load token cache from file if it exists."""
        if TOKEN_CACHE_FILE.exists():
            try:
                self._token_cache.deserialize(TOKEN_CACHE_FILE.read_text())
            except (json.JSONDecodeError, Exception) as e:
                console.print(f"[yellow]Warning: Could not load token cache: {e}[/yellow]")

    def _save_token_cache(self) -> None:
        """Save token cache to file."""
        if self._token_cache.has_state_changed:
            TOKEN_CACHE_FILE.write_text(self._token_cache.serialize())

    def get_access_token(self) -> Optional[str]:
        """
        Get an access token for Microsoft Graph API.

        First attempts to get a cached token, then falls back to interactive login.

        Returns:
            Access token string or None if authentication fails
        """
        accounts = self.app.get_accounts()

        if accounts:
            # Try to get token silently for the first account
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._save_token_cache()
                return result["access_token"]

        # No cached token, need interactive login
        console.print("\n[bold blue]Microsoft Graph API Authentication[/bold blue]")
        console.print("You'll need to sign in with your Microsoft account.")
        console.print("A browser window will open for authentication.\n")

        try:
            result = self.app.acquire_token_interactive(
                scopes=SCOPES,
                prompt="select_account",
            )
        except Exception as e:
            # Fallback to device code flow if interactive fails
            console.print(f"[yellow]Interactive login failed: {e}[/yellow]")
            console.print("[yellow]Falling back to device code authentication...[/yellow]")
            result = self._device_code_flow()

        if result and "access_token" in result:
            self._save_token_cache()
            console.print("[green]Authentication successful![/green]\n")
            return result["access_token"]

        if result and "error" in result:
            console.print(f"[red]Authentication error: {result['error']}[/red]")
            console.print(f"[red]{result.get('error_description', '')}[/red]")

        return None

    def _device_code_flow(self) -> Optional[dict]:
        """
        Authenticate using device code flow (for environments without browser).

        Returns:
            Token result dictionary or None
        """
        flow = self.app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flow:
            console.print(f"[red]Could not create device flow: {flow.get('error')}[/red]")
            return None

        console.print(f"\n[bold]To sign in:[/bold]")
        console.print(f"1. Go to: [cyan]{flow['verification_uri']}[/cyan]")
        console.print(f"2. Enter code: [bold yellow]{flow['user_code']}[/bold yellow]\n")

        result = self.app.acquire_token_by_device_flow(flow)
        return result

    def logout(self) -> None:
        """Clear the token cache and log out."""
        if TOKEN_CACHE_FILE.exists():
            TOKEN_CACHE_FILE.unlink()
        console.print("[green]Logged out successfully.[/green]")


def get_authenticator() -> GraphAuthenticator:
    """
    Create an authenticator from environment variables.

    Returns:
        GraphAuthenticator instance

    Raises:
        ValueError: If CLIENT_ID is not set
    """
    client_id = os.getenv("CLIENT_ID")
    tenant_id = os.getenv("TENANT_ID", "common")

    if not client_id:
        raise ValueError(
            "CLIENT_ID environment variable is required.\n"
            "Please set up an Azure AD app registration and configure .env file.\n"
            "See .env.example for details."
        )

    return GraphAuthenticator(client_id=client_id, tenant_id=tenant_id)
