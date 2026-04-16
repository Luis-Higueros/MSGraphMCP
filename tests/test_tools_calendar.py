"""Tests for Calendar MCP tools."""

import pytest
import httpx
import respx
from mcp.server.fastmcp import FastMCP

from msgraph_mcp.graph_client import GRAPH_BASE_URL
from msgraph_mcp.tools.calendar import register_calendar_tools
from tests.conftest import parse_tool_result


@pytest.fixture
def mcp_and_client(mock_client):
    mcp = FastMCP("test")
    register_calendar_tools(mcp, mock_client)
    return mcp, mock_client


class TestCalendarTools:
    @respx.mock
    @pytest.mark.asyncio
    async def test_list_calendar_events(self, mcp_and_client):
        mcp, _ = mcp_and_client
        respx.get(f"{GRAPH_BASE_URL}/me/events").mock(
            return_value=httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "evt-1",
                            "subject": "Team Standup",
                            "start": {
                                "dateTime": "2024-06-01T09:00:00",
                                "timeZone": "UTC",
                            },
                            "end": {
                                "dateTime": "2024-06-01T09:30:00",
                                "timeZone": "UTC",
                            },
                            "location": {"displayName": "Teams"},
                            "organizer": {
                                "emailAddress": {"name": "Alice"}
                            },
                            "webLink": "https://outlook.com/...",
                            "isOnlineMeeting": True,
                            "onlineMeetingUrl": "https://teams.microsoft.com/...",
                            "bodyPreview": "Daily standup",
                        }
                    ]
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("list_calendar_events", {}))
        assert len(result) == 1
        assert result[0]["subject"] == "Team Standup"
        assert result[0]["isOnlineMeeting"] is True

    @respx.mock
    @pytest.mark.asyncio
    async def test_create_calendar_event(self, mcp_and_client):
        mcp, _ = mcp_and_client
        respx.post(f"{GRAPH_BASE_URL}/me/events").mock(
            return_value=httpx.Response(
                201,
                json={
                    "id": "new-evt",
                    "subject": "New Meeting",
                    "start": {"dateTime": "2024-07-01T14:00:00"},
                    "end": {"dateTime": "2024-07-01T15:00:00"},
                    "webLink": "https://outlook.com/...",
                    "onlineMeetingUrl": None,
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool(
            "create_calendar_event",
            {
                "subject": "New Meeting",
                "start_datetime": "2024-07-01T14:00:00",
                "end_datetime": "2024-07-01T15:00:00",
            },
        ))
        assert result["id"] == "new-evt"
        assert result["subject"] == "New Meeting"

    @respx.mock
    @pytest.mark.asyncio
    async def test_get_calendar_event(self, mcp_and_client):
        mcp, _ = mcp_and_client
        event_id = "evt-123"
        respx.get(f"{GRAPH_BASE_URL}/me/events/{event_id}").mock(
            return_value=httpx.Response(
                200,
                json={
                    "id": event_id,
                    "subject": "All Hands",
                    "bodyPreview": "Company meeting",
                    "body": {"content": "<p>Agenda</p>"},
                    "start": {"dateTime": "2024-08-01T10:00:00"},
                    "end": {"dateTime": "2024-08-01T11:00:00"},
                    "location": {"displayName": "Main Hall"},
                    "organizer": {"emailAddress": {"name": "CEO"}},
                    "attendees": [
                        {
                            "emailAddress": {
                                "name": "Bob",
                                "address": "bob@example.com",
                            },
                            "status": {"response": "accepted"},
                        }
                    ],
                    "isOnlineMeeting": False,
                    "onlineMeetingUrl": None,
                    "webLink": "https://outlook.com/...",
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("get_calendar_event", {"event_id": event_id}))
        assert result["subject"] == "All Hands"
        assert len(result["attendees"]) == 1
        assert result["attendees"][0]["name"] == "Bob"
