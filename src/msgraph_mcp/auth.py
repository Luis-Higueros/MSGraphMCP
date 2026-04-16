"""Authentication helpers for Microsoft Graph API using MSAL."""

from __future__ import annotations

import os
from typing import Optional

import msal


GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]

# Delegated (user) permission scopes for interactive login
DELEGATED_SCOPES = [
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Team.ReadBasic.All",
    "https://graph.microsoft.com/Channel.ReadBasic.All",
    "https://graph.microsoft.com/ChannelMessage.Read.All",
    "https://graph.microsoft.com/ChannelMessage.Send",
    "https://graph.microsoft.com/Chat.Read",
    "https://graph.microsoft.com/Chat.ReadWrite",
    "https://graph.microsoft.com/Calendars.ReadWrite",
    "https://graph.microsoft.com/Mail.ReadWrite",
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/Files.Read.All",
    "https://graph.microsoft.com/Presence.Read.All",
]


class GraphAuthError(Exception):
    """Raised when authentication to Microsoft Graph fails."""


def _get_env(key: str, required: bool = True) -> Optional[str]:
    """Read a configuration value from the environment."""
    value = os.environ.get(key)
    if required and not value:
        raise GraphAuthError(
            f"Required environment variable '{key}' is not set. "
            "Please configure it in your .env file or environment."
        )
    return value


class GraphAuthProvider:
    """Provides access tokens for Microsoft Graph API calls.

    Supports two authentication flows:

    * **Client Credentials** (app-only): Uses CLIENT_ID + CLIENT_SECRET with
      application permissions.  Set ``AUTH_FLOW=client_credentials`` (default).

    * **Device Code** (delegated): Guides the user through a browser-based
      login to acquire delegated permissions.  Set ``AUTH_FLOW=device_code``.

    Required environment variables:
        TENANT_ID   – Azure AD tenant ID
        CLIENT_ID   – App registration client ID
        CLIENT_SECRET – App secret (only for client_credentials flow)
    """

    def __init__(self) -> None:
        self._tenant_id = _get_env("TENANT_ID")
        self._client_id = _get_env("CLIENT_ID")
        self._auth_flow = os.environ.get("AUTH_FLOW", "client_credentials").lower()
        self._authority = f"https://login.microsoftonline.com/{self._tenant_id}"
        self._app: Optional[msal.ClientApplication] = None
        self._token_cache = msal.SerializableTokenCache()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_access_token(self) -> str:
        """Return a valid access token, refreshing if necessary."""
        if self._auth_flow == "client_credentials":
            return self._client_credentials_token()
        elif self._auth_flow == "device_code":
            return self._device_code_token()
        else:
            raise GraphAuthError(
                f"Unknown AUTH_FLOW value '{self._auth_flow}'. "
                "Use 'client_credentials' or 'device_code'."
            )

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _client_credentials_token(self) -> str:
        client_secret = _get_env("CLIENT_SECRET")
        if self._app is None or not isinstance(
            self._app, msal.ConfidentialClientApplication
        ):
            self._app = msal.ConfidentialClientApplication(
                self._client_id,
                authority=self._authority,
                client_credential=client_secret,
                token_cache=self._token_cache,
            )

        result = self._app.acquire_token_silent(GRAPH_SCOPES, account=None)
        if not result:
            result = self._app.acquire_token_for_client(scopes=GRAPH_SCOPES)

        return self._extract_token(result)

    def _device_code_token(self) -> str:
        if self._app is None or not isinstance(
            self._app, msal.PublicClientApplication
        ):
            self._app = msal.PublicClientApplication(
                self._client_id,
                authority=self._authority,
                token_cache=self._token_cache,
            )

        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(
                DELEGATED_SCOPES, account=accounts[0]
            )
            if result:
                return self._extract_token(result)

        flow = self._app.initiate_device_flow(scopes=DELEGATED_SCOPES)
        if "user_code" not in flow:
            raise GraphAuthError("Failed to initiate device flow.")

        print(flow["message"])  # prints the "go to … and enter code …" message
        result = self._app.acquire_token_by_device_flow(flow)
        return self._extract_token(result)

    @staticmethod
    def _extract_token(result: Optional[dict]) -> str:
        if not result or "access_token" not in result:
            error = (result or {}).get("error_description", "Unknown error")
            raise GraphAuthError(f"Failed to acquire access token: {error}")
        return result["access_token"]
