"""Tests for Users MCP tools."""

import pytest
import httpx
import respx
from mcp.server.fastmcp import FastMCP

from msgraph_mcp.graph_client import GRAPH_BASE_URL
from msgraph_mcp.tools.users import register_users_tools
from tests.conftest import parse_tool_result


@pytest.fixture
def mcp_and_client(mock_client):
    mcp = FastMCP("test")
    register_users_tools(mcp, mock_client)
    return mcp, mock_client


class TestUsersTools:
    @respx.mock
    @pytest.mark.asyncio
    async def test_get_current_user(self, mcp_and_client):
        mcp, _ = mcp_and_client
        respx.get(f"{GRAPH_BASE_URL}/me").mock(
            return_value=httpx.Response(
                200,
                json={
                    "id": "user-1",
                    "displayName": "John Doe",
                    "mail": "john@example.com",
                    "jobTitle": "Engineer",
                    "department": "Engineering",
                    "officeLocation": "Building A",
                    "mobilePhone": "+1-555-0100",
                    "businessPhones": ["+1-555-0200"],
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("get_current_user", {}))
        assert result["displayName"] == "John Doe"
        assert result["email"] == "john@example.com"

    @respx.mock
    @pytest.mark.asyncio
    async def test_search_users(self, mcp_and_client):
        mcp, _ = mcp_and_client
        respx.get(f"{GRAPH_BASE_URL}/users").mock(
            return_value=httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "user-2",
                            "displayName": "Jane Smith",
                            "mail": "jane@example.com",
                            "jobTitle": "Designer",
                            "department": "Product",
                        }
                    ]
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("search_users", {"query": "Jane"}))
        assert len(result) == 1
        assert result[0]["displayName"] == "Jane Smith"

    @respx.mock
    @pytest.mark.asyncio
    async def test_get_user_presence(self, mcp_and_client):
        mcp, _ = mcp_and_client
        user_id = "user-1"
        respx.get(f"{GRAPH_BASE_URL}/users/{user_id}/presence").mock(
            return_value=httpx.Response(
                200,
                json={
                    "id": user_id,
                    "availability": "Available",
                    "activity": "Available",
                    "statusMessage": None,
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("get_user_presence", {"user_id": user_id}))
        assert result["availability"] == "Available"
        assert result["userId"] == user_id

    @respx.mock
    @pytest.mark.asyncio
    async def test_get_my_presence(self, mcp_and_client):
        mcp, _ = mcp_and_client
        respx.get(f"{GRAPH_BASE_URL}/me/presence").mock(
            return_value=httpx.Response(
                200,
                json={
                    "availability": "Busy",
                    "activity": "InAMeeting",
                    "statusMessage": {
                        "message": {"content": "In standup"}
                    },
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("get_my_presence", {}))
        assert result["availability"] == "Busy"
        assert result["statusMessage"] == "In standup"
