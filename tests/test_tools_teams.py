"""Tests for Teams MCP tools."""

import pytest
import httpx
import respx
from mcp.server.fastmcp import FastMCP

from msgraph_mcp.graph_client import GraphClient, GRAPH_BASE_URL
from msgraph_mcp.tools.teams import register_teams_tools
from tests.conftest import parse_tool_result


@pytest.fixture
def mcp_and_client(mock_client):
    mcp = FastMCP("test")
    register_teams_tools(mcp, mock_client)
    return mcp, mock_client


class TestTeamsTools:
    @respx.mock
    @pytest.mark.asyncio
    async def test_list_teams(self, mcp_and_client):
        mcp, client = mcp_and_client
        respx.get(f"{GRAPH_BASE_URL}/me/joinedTeams").mock(
            return_value=httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "team-1",
                            "displayName": "Engineering",
                            "description": "Eng team",
                        }
                    ]
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("list_teams", {}))
        assert len(result) == 1
        assert result[0]["displayName"] == "Engineering"
        assert result[0]["id"] == "team-1"

    @respx.mock
    @pytest.mark.asyncio
    async def test_list_channels(self, mcp_and_client):
        mcp, client = mcp_and_client
        team_id = "team-abc"
        respx.get(f"{GRAPH_BASE_URL}/teams/{team_id}/channels").mock(
            return_value=httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "ch-1",
                            "displayName": "General",
                            "description": "",
                            "membershipType": "standard",
                        }
                    ]
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("list_channels", {"team_id": team_id}))
        assert len(result) == 1
        assert result[0]["displayName"] == "General"

    @respx.mock
    @pytest.mark.asyncio
    async def test_get_channel_messages(self, mcp_and_client):
        mcp, client = mcp_and_client
        team_id, channel_id = "team-1", "ch-1"
        respx.get(
            f"{GRAPH_BASE_URL}/teams/{team_id}/channels/{channel_id}/messages"
        ).mock(
            return_value=httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "msg-1",
                            "createdDateTime": "2024-01-01T10:00:00Z",
                            "from": {"user": {"displayName": "Alice"}},
                            "body": {"content": "Hello!", "contentType": "text"},
                            "importance": "normal",
                            "subject": None,
                        }
                    ]
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool(
            "get_channel_messages",
            {"team_id": team_id, "channel_id": channel_id},
        ))
        assert len(result) == 1
        assert result[0]["body"] == "Hello!"
        assert result[0]["sender"] == "Alice"

    @respx.mock
    @pytest.mark.asyncio
    async def test_send_channel_message(self, mcp_and_client):
        mcp, client = mcp_and_client
        team_id, channel_id = "team-1", "ch-1"
        respx.post(
            f"{GRAPH_BASE_URL}/teams/{team_id}/channels/{channel_id}/messages"
        ).mock(
            return_value=httpx.Response(
                201,
                json={
                    "id": "new-msg",
                    "createdDateTime": "2024-01-01T11:00:00Z",
                    "webUrl": "https://teams.microsoft.com/...",
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool(
            "send_channel_message",
            {
                "team_id": team_id,
                "channel_id": channel_id,
                "message": "Test message",
            },
        ))
        assert result["id"] == "new-msg"

    @respx.mock
    @pytest.mark.asyncio
    async def test_list_team_members(self, mcp_and_client):
        mcp, client = mcp_and_client
        team_id = "team-1"
        respx.get(f"{GRAPH_BASE_URL}/teams/{team_id}/members").mock(
            return_value=httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "member-1",
                            "displayName": "Bob",
                            "email": "bob@example.com",
                            "roles": ["owner"],
                        }
                    ]
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("list_team_members", {"team_id": team_id}))
        assert len(result) == 1
        assert result[0]["displayName"] == "Bob"
        assert "owner" in result[0]["roles"]

    @respx.mock
    @pytest.mark.asyncio
    async def test_top_param_capped_at_50(self, mcp_and_client):
        mcp, client = mcp_and_client
        team_id, channel_id = "team-1", "ch-1"
        route = respx.get(
            f"{GRAPH_BASE_URL}/teams/{team_id}/channels/{channel_id}/messages"
        ).mock(return_value=httpx.Response(200, json={"value": []}))
        await mcp.call_tool(
            "get_channel_messages",
            {"team_id": team_id, "channel_id": channel_id, "top": 200},
        )
        assert route.called
        request = route.calls[0].request
        # $top should be capped at 50
        assert "%24top=50" in str(request.url) or "$top=50" in str(request.url)
