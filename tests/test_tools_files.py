"""Tests for Files/OneDrive MCP tools."""

import pytest
import httpx
import respx
from mcp.server.fastmcp import FastMCP

from msgraph_mcp.graph_client import GRAPH_BASE_URL
from msgraph_mcp.tools.files import register_files_tools
from tests.conftest import parse_tool_result


@pytest.fixture
def mcp_and_client(mock_client):
    mcp = FastMCP("test")
    register_files_tools(mcp, mock_client)
    return mcp, mock_client


class TestFilesTools:
    @respx.mock
    @pytest.mark.asyncio
    async def test_list_drive_items_root(self, mcp_and_client):
        mcp, _ = mcp_and_client
        respx.get(f"{GRAPH_BASE_URL}/me/drive/root/children").mock(
            return_value=httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "file-1",
                            "name": "report.docx",
                            "file": {"mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
                            "size": 10240,
                            "lastModifiedDateTime": "2024-01-10T12:00:00Z",
                            "createdDateTime": "2024-01-01T10:00:00Z",
                            "webUrl": "https://example.sharepoint.com/...",
                        }
                    ]
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("list_drive_items", {}))
        assert len(result) == 1
        assert result[0]["name"] == "report.docx"
        assert result[0]["type"] == "file"

    @respx.mock
    @pytest.mark.asyncio
    async def test_list_drive_items_subfolder(self, mcp_and_client):
        mcp, _ = mcp_and_client
        respx.get(
            f"{GRAPH_BASE_URL}/me/drive/root:/Documents/Reports:/children"
        ).mock(
            return_value=httpx.Response(200, json={"value": []})
        )
        result = parse_tool_result(await mcp.call_tool(
            "list_drive_items", {"folder_path": "/Documents/Reports"}
        ))
        assert result == []

    @respx.mock
    @pytest.mark.asyncio
    async def test_get_drive_item(self, mcp_and_client):
        mcp, _ = mcp_and_client
        item_id = "item-abc"
        respx.get(f"{GRAPH_BASE_URL}/me/drive/items/{item_id}").mock(
            return_value=httpx.Response(
                200,
                json={
                    "id": item_id,
                    "name": "budget.xlsx",
                    "file": {"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                    "size": 5120,
                    "lastModifiedDateTime": "2024-02-01T08:00:00Z",
                    "createdDateTime": "2024-01-15T09:00:00Z",
                    "webUrl": "https://example.sharepoint.com/...",
                    "@microsoft.graph.downloadUrl": "https://download.example.com/...",
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("get_drive_item", {"item_id": item_id}))
        assert result["name"] == "budget.xlsx"
        assert "downloadUrl" in result

    @respx.mock
    @pytest.mark.asyncio
    async def test_search_drive_items(self, mcp_and_client):
        mcp, _ = mcp_and_client
        respx.get(
            f"{GRAPH_BASE_URL}/me/drive/root/search(q='budget')"
        ).mock(
            return_value=httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "file-2",
                            "name": "budget_2024.xlsx",
                            "file": {"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                            "size": 3072,
                            "lastModifiedDateTime": "2024-03-01T10:00:00Z",
                            "webUrl": "https://example.sharepoint.com/...",
                        }
                    ]
                },
            )
        )
        result = parse_tool_result(await mcp.call_tool("search_drive_items", {"query": "budget"}))
        assert len(result) == 1
        assert result[0]["name"] == "budget_2024.xlsx"
