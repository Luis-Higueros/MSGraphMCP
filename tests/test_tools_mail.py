"""Tests for Mail MCP tools."""

import pytest
import httpx
import respx
from mcp.server.fastmcp import FastMCP

from msgraph_mcp.graph_client import GRAPH_BASE_URL
from msgraph_mcp.tools.mail import register_mail_tools
from tests.conftest import parse_tool_result


@pytest.fixture
def mcp_and_client(mock_client):
    mcp = FastMCP("test")
    register_mail_tools(mcp, mock_client)
    return mcp, mock_client


class TestMailTools:
    @respx.mock
    @pytest.mark.asyncio
    async def test_list_emails(self, mcp_and_client):
        mcp, _ = mcp_and_client
        respx.get(f"{GRAPH_BASE_URL}/me/mailFolders/inbox/messages").mock(
            return_value=httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "msg-1",
                            "subject": "Hello",
                            "from": {
                                "emailAddress": {
                                    "name": "Alice",
                                    "address": "alice@example.com",
                                }
                            },
                            "receivedDateTime": "2024-01-01T10:00:00Z",
                            "bodyPreview": "Hi there",
                            "isRead": False,
                            "hasAttachments": False,
                            "importance": "normal",
                        }
                    ]
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("list_emails", {}))
        assert len(result) == 1
        assert result[0]["subject"] == "Hello"
        assert result[0]["isRead"] is False

    @respx.mock
    @pytest.mark.asyncio
    async def test_send_email(self, mcp_and_client):
        mcp, _ = mcp_and_client
        respx.post(f"{GRAPH_BASE_URL}/me/sendMail").mock(
            return_value=httpx.Response(202, json={})
        )
        result = parse_tool_result(await mcp.call_tool(
            "send_email",
            {
                "to_emails": ["bob@example.com"],
                "subject": "Test",
                "body": "<p>Hello Bob</p>",
            },
        ))
        assert result["status"] == "sent"
        assert "bob@example.com" in result["to"]

    @respx.mock
    @pytest.mark.asyncio
    async def test_get_email(self, mcp_and_client):
        mcp, _ = mcp_and_client
        msg_id = "msg-abc"
        respx.get(f"{GRAPH_BASE_URL}/me/messages/{msg_id}").mock(
            return_value=httpx.Response(
                200,
                json={
                    "id": msg_id,
                    "subject": "Weekly Report",
                    "from": {
                        "emailAddress": {
                            "name": "Manager",
                            "address": "mgr@example.com",
                        }
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": "me@example.com"}}
                    ],
                    "ccRecipients": [],
                    "body": {
                        "contentType": "html",
                        "content": "<p>Weekly update</p>",
                    },
                    "receivedDateTime": "2024-01-05T09:00:00Z",
                    "isRead": True,
                    "hasAttachments": False,
                    "importance": "normal",
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("get_email", {"message_id": msg_id}))
        assert result["subject"] == "Weekly Report"
        assert "me@example.com" in result["to"]

    @respx.mock
    @pytest.mark.asyncio
    async def test_reply_to_email(self, mcp_and_client):
        mcp, _ = mcp_and_client
        msg_id = "msg-1"
        respx.post(f"{GRAPH_BASE_URL}/me/messages/{msg_id}/reply").mock(
            return_value=httpx.Response(202, json={})
        )
        result = parse_tool_result(await mcp.call_tool(
            "reply_to_email",
            {"message_id": msg_id, "body": "Got it, thanks!"},
        ))
        assert result["status"] == "sent"
        assert result["replyAll"] is False
