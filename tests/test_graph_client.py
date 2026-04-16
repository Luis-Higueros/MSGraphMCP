"""Tests for the GraphClient HTTP wrapper."""

import pytest
import httpx
import respx

from msgraph_mcp.graph_client import GraphClient, GraphClientError, GRAPH_BASE_URL


@pytest.fixture
def client(mock_auth):
    return GraphClient(mock_auth)


class TestGraphClient:
    @respx.mock
    @pytest.mark.asyncio
    async def test_get_returns_json(self, client, mock_auth):
        respx.get(f"{GRAPH_BASE_URL}/me").mock(
            return_value=httpx.Response(200, json={"displayName": "Test User"})
        )
        result = await client.get("/me")
        assert result == {"displayName": "Test User"}

    @respx.mock
    @pytest.mark.asyncio
    async def test_get_includes_auth_header(self, client, mock_auth):
        route = respx.get(f"{GRAPH_BASE_URL}/me").mock(
            return_value=httpx.Response(200, json={})
        )
        await client.get("/me")
        assert route.called
        request = route.calls[0].request
        assert request.headers["Authorization"] == "Bearer dummy-token"

    @respx.mock
    @pytest.mark.asyncio
    async def test_post_returns_json(self, client):
        respx.post(f"{GRAPH_BASE_URL}/me/sendMail").mock(
            return_value=httpx.Response(202, json={})
        )
        result = await client.post("/me/sendMail", json={"message": {}})
        assert result == {}

    @respx.mock
    @pytest.mark.asyncio
    async def test_error_response_raises(self, client):
        respx.get(f"{GRAPH_BASE_URL}/me").mock(
            return_value=httpx.Response(
                403,
                json={"error": {"message": "Access denied"}},
            )
        )
        with pytest.raises(GraphClientError) as exc_info:
            await client.get("/me")
        assert exc_info.value.status_code == 403
        assert "Access denied" in str(exc_info.value)

    @respx.mock
    @pytest.mark.asyncio
    async def test_get_with_params(self, client):
        route = respx.get(f"{GRAPH_BASE_URL}/me/events").mock(
            return_value=httpx.Response(200, json={"value": []})
        )
        result = await client.get("/me/events", params={"$top": 5})
        assert result == {"value": []}
        assert route.called
