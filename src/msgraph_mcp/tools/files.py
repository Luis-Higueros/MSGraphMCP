"""OneDrive / SharePoint Files tools for the MS Graph MCP server."""

from __future__ import annotations

from typing import Optional

from mcp.server.fastmcp import FastMCP

from ..graph_client import GraphClient


def register_files_tools(mcp: FastMCP, client: GraphClient) -> None:
    """Register all Files (OneDrive/SharePoint) MCP tools onto *mcp*."""

    @mcp.tool()
    async def list_drive_items(
        folder_path: str = "/",
        top: int = 20,
    ) -> list[dict]:
        """List files and folders in OneDrive.

        Args:
            folder_path: Path to the folder to list (default '/' for root).
                Use forward-slash separated paths, e.g. '/Documents/Reports'.
            top: Maximum number of items to return (default 20, max 100).

        Returns a list of items with id, name, type, size, lastModifiedDateTime,
        and webUrl.
        """
        top = min(top, 100)
        if folder_path in ("/", ""):
            endpoint = "/me/drive/root/children"
        else:
            # Encode path for Graph API
            clean_path = folder_path.strip("/")
            endpoint = f"/me/drive/root:/{clean_path}:/children"

        data = await client.get(
            endpoint,
            params={
                "$top": top,
                "$select": (
                    "id,name,file,folder,size,lastModifiedDateTime,"
                    "createdDateTime,webUrl,parentReference"
                ),
                "$orderby": "name",
            },
        )
        items = data.get("value", [])
        return [
            {
                "id": item.get("id"),
                "name": item.get("name"),
                "type": "folder" if "folder" in item else "file",
                "mimeType": item.get("file", {}).get("mimeType") if "file" in item else None,
                "size": item.get("size"),
                "lastModifiedDateTime": item.get("lastModifiedDateTime"),
                "createdDateTime": item.get("createdDateTime"),
                "webUrl": item.get("webUrl"),
            }
            for item in items
        ]

    @mcp.tool()
    async def get_drive_item(item_id: str) -> dict:
        """Get metadata for a specific OneDrive file or folder.

        Args:
            item_id: The unique identifier of the drive item.

        Returns item metadata including name, type, size, dates, webUrl,
        and download URL (for files).
        """
        item = await client.get(f"/me/drive/items/{item_id}")
        result = {
            "id": item.get("id"),
            "name": item.get("name"),
            "type": "folder" if "folder" in item else "file",
            "size": item.get("size"),
            "lastModifiedDateTime": item.get("lastModifiedDateTime"),
            "createdDateTime": item.get("createdDateTime"),
            "webUrl": item.get("webUrl"),
        }
        if "file" in item:
            result["mimeType"] = item["file"].get("mimeType")
        if "@microsoft.graph.downloadUrl" in item:
            result["downloadUrl"] = item["@microsoft.graph.downloadUrl"]
        return result

    @mcp.tool()
    async def search_drive_items(
        query: str,
        top: int = 10,
    ) -> list[dict]:
        """Search for files and folders in OneDrive by name or content.

        Args:
            query: Search query string.
            top: Maximum number of results to return (default 10, max 50).

        Returns a list of matching items with id, name, type, size,
        lastModifiedDateTime, and webUrl.
        """
        top = min(top, 50)
        data = await client.get(
            f"/me/drive/root/search(q='{query}')",
            params={
                "$top": top,
                "$select": (
                    "id,name,file,folder,size,lastModifiedDateTime,webUrl"
                ),
            },
        )
        items = data.get("value", [])
        return [
            {
                "id": item.get("id"),
                "name": item.get("name"),
                "type": "folder" if "folder" in item else "file",
                "mimeType": item.get("file", {}).get("mimeType") if "file" in item else None,
                "size": item.get("size"),
                "lastModifiedDateTime": item.get("lastModifiedDateTime"),
                "webUrl": item.get("webUrl"),
            }
            for item in items
        ]

    @mcp.tool()
    async def list_sharepoint_sites(
        top: int = 10,
    ) -> list[dict]:
        """List SharePoint sites accessible to the current user.

        Args:
            top: Maximum number of sites to return (default 10, max 50).

        Returns a list of sites with id, displayName, webUrl, and description.
        """
        top = min(top, 50)
        data = await client.get(
            "/sites",
            params={
                "$top": top,
                "search": "*",
                "$select": "id,displayName,webUrl,description",
            },
        )
        sites = data.get("value", [])
        return [
            {
                "id": s.get("id"),
                "displayName": s.get("displayName"),
                "webUrl": s.get("webUrl"),
                "description": s.get("description", ""),
            }
            for s in sites
        ]

    @mcp.tool()
    async def list_sharepoint_site_files(
        site_id: str,
        folder_path: str = "/",
        top: int = 20,
    ) -> list[dict]:
        """List files in a SharePoint site's document library.

        Args:
            site_id: The unique identifier of the SharePoint site.
            folder_path: Path within the site's default drive (default '/').
            top: Maximum number of items to return (default 20, max 100).

        Returns a list of files and folders with id, name, type, size,
        lastModifiedDateTime, and webUrl.
        """
        top = min(top, 100)
        if folder_path in ("/", ""):
            endpoint = f"/sites/{site_id}/drive/root/children"
        else:
            clean_path = folder_path.strip("/")
            endpoint = f"/sites/{site_id}/drive/root:/{clean_path}:/children"

        data = await client.get(
            endpoint,
            params={
                "$top": top,
                "$select": (
                    "id,name,file,folder,size,lastModifiedDateTime,webUrl"
                ),
                "$orderby": "name",
            },
        )
        items = data.get("value", [])
        return [
            {
                "id": item.get("id"),
                "name": item.get("name"),
                "type": "folder" if "folder" in item else "file",
                "mimeType": item.get("file", {}).get("mimeType") if "file" in item else None,
                "size": item.get("size"),
                "lastModifiedDateTime": item.get("lastModifiedDateTime"),
                "webUrl": item.get("webUrl"),
            }
            for item in items
        ]
