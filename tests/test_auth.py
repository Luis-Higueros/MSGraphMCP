"""Tests for the authentication module."""

import os
import pytest
from unittest.mock import MagicMock, patch

from msgraph_mcp.auth import GraphAuthProvider, GraphAuthError


class TestGraphAuthProvider:
    def _make_provider(self, env: dict | None = None) -> GraphAuthProvider:
        defaults = {
            "TENANT_ID": "test-tenant",
            "CLIENT_ID": "test-client",
            "CLIENT_SECRET": "test-secret",
            "AUTH_FLOW": "client_credentials",
        }
        combined = {**defaults, **(env or {})}
        with patch.dict(os.environ, combined, clear=False):
            return GraphAuthProvider()

    def test_missing_tenant_id_raises(self):
        with patch.dict(os.environ, {}, clear=True):
            with pytest.raises(GraphAuthError, match="TENANT_ID"):
                GraphAuthProvider()

    def test_unknown_auth_flow_raises(self):
        provider = self._make_provider({"AUTH_FLOW": "magic"})
        with pytest.raises(GraphAuthError, match="AUTH_FLOW"):
            provider.get_access_token()

    def test_client_credentials_acquires_token(self):
        provider = self._make_provider()

        mock_app = MagicMock()
        mock_app.acquire_token_silent.return_value = None
        mock_app.acquire_token_for_client.return_value = {
            "access_token": "test-token"
        }

        with patch.dict(os.environ, {"CLIENT_SECRET": "test-secret"}):
            with patch("msal.ConfidentialClientApplication", return_value=mock_app):
                token = provider.get_access_token()

        assert token == "test-token"
        mock_app.acquire_token_for_client.assert_called_once()

    def test_client_credentials_uses_cached_token(self):
        provider = self._make_provider()

        mock_app = MagicMock()
        mock_app.acquire_token_silent.return_value = {"access_token": "cached-token"}

        with patch.dict(os.environ, {"CLIENT_SECRET": "test-secret"}):
            with patch("msal.ConfidentialClientApplication", return_value=mock_app):
                token = provider.get_access_token()

        assert token == "cached-token"
        mock_app.acquire_token_for_client.assert_not_called()

    def test_failed_token_acquisition_raises(self):
        provider = self._make_provider()

        mock_app = MagicMock()
        mock_app.acquire_token_silent.return_value = None
        mock_app.acquire_token_for_client.return_value = {
            "error": "invalid_client",
            "error_description": "Bad credentials",
        }

        with patch.dict(os.environ, {"CLIENT_SECRET": "test-secret"}):
            with patch("msal.ConfidentialClientApplication", return_value=mock_app):
                with pytest.raises(GraphAuthError, match="Bad credentials"):
                    provider.get_access_token()
