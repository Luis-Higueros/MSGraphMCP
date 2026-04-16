"""Thin async HTTP client wrapper around the Microsoft Graph REST API."""

from __future__ import annotations

from typing import Any, Optional

import httpx

from .auth import GraphAuthProvider


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


class GraphClientError(Exception):
    """Raised when a Graph API call returns an error response."""

    def __init__(self, status_code: int, message: str) -> None:
        self.status_code = status_code
        super().__init__(f"Graph API error {status_code}: {message}")


class GraphClient:
    """Async Microsoft Graph API client.

    Usage::

        async with GraphClient(auth_provider) as client:
            data = await client.get("/me")
    """

    def __init__(self, auth_provider: GraphAuthProvider) -> None:
        self._auth = auth_provider
        self._http: Optional[httpx.AsyncClient] = None

    # ------------------------------------------------------------------
    # Context-manager / lifecycle
    # ------------------------------------------------------------------

    async def __aenter__(self) -> "GraphClient":
        self._http = httpx.AsyncClient(base_url=GRAPH_BASE_URL, timeout=30)
        return self

    async def __aexit__(self, *_: Any) -> None:
        if self._http:
            await self._http.aclose()
            self._http = None

    # ------------------------------------------------------------------
    # HTTP helpers
    # ------------------------------------------------------------------

    def _headers(self) -> dict[str, str]:
        token = self._auth.get_access_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    def _client(self) -> httpx.AsyncClient:
        if self._http is None:
            # Lazy initialisation: create the underlying client on first call.
            self._http = httpx.AsyncClient(base_url=GRAPH_BASE_URL, timeout=30)
        return self._http

    async def get(
        self,
        path: str,
        params: Optional[dict[str, Any]] = None,
    ) -> Any:
        response = await self._client().get(
            path, headers=self._headers(), params=params
        )
        return self._handle(response)

    async def post(self, path: str, json: Any = None) -> Any:
        response = await self._client().post(
            path, headers=self._headers(), json=json
        )
        return self._handle(response)

    async def patch(self, path: str, json: Any = None) -> Any:
        response = await self._client().patch(
            path, headers=self._headers(), json=json
        )
        return self._handle(response)

    async def delete(self, path: str) -> None:
        response = await self._client().delete(path, headers=self._headers())
        if response.status_code not in (200, 204):
            self._handle(response)

    # ------------------------------------------------------------------
    # Response handling
    # ------------------------------------------------------------------

    @staticmethod
    def _handle(response: httpx.Response) -> Any:
        if response.status_code in (200, 201, 202):
            if response.content:
                return response.json()
            return {}
        try:
            body = response.json()
            message = body.get("error", {}).get("message", response.text)
        except Exception:
            message = response.text
        raise GraphClientError(response.status_code, message)
