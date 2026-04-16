"""Microsoft Teams tools for the MS Graph MCP server."""

from __future__ import annotations

from typing import Optional

from mcp.server.fastmcp import FastMCP

from ..graph_client import GraphClient


def register_teams_tools(mcp: FastMCP, client: GraphClient) -> None:
    """Register all Teams-related MCP tools onto *mcp*."""

    @mcp.tool()
    async def list_teams() -> list[dict]:
        """List all Microsoft Teams that the current user is a member of.

        Returns a list of teams with id, displayName, and description.
        """
        data = await client.get("/me/joinedTeams")
        teams = data.get("value", [])
        return [
            {
                "id": t.get("id"),
                "displayName": t.get("displayName"),
                "description": t.get("description", ""),
            }
            for t in teams
        ]

    @mcp.tool()
    async def get_team(team_id: str) -> dict:
        """Get details of a specific Microsoft Team.

        Args:
            team_id: The unique identifier of the team.

        Returns team details including id, displayName, description,
        visibility, and webUrl.
        """
        team = await client.get(f"/teams/{team_id}")
        return {
            "id": team.get("id"),
            "displayName": team.get("displayName"),
            "description": team.get("description", ""),
            "visibility": team.get("visibility"),
            "webUrl": team.get("webUrl"),
        }

    @mcp.tool()
    async def list_channels(team_id: str) -> list[dict]:
        """List all channels in a Microsoft Team.

        Args:
            team_id: The unique identifier of the team.

        Returns a list of channels with id, displayName, description,
        and membershipType.
        """
        data = await client.get(f"/teams/{team_id}/channels")
        channels = data.get("value", [])
        return [
            {
                "id": c.get("id"),
                "displayName": c.get("displayName"),
                "description": c.get("description", ""),
                "membershipType": c.get("membershipType"),
            }
            for c in channels
        ]

    @mcp.tool()
    async def get_channel_messages(
        team_id: str,
        channel_id: str,
        top: int = 20,
    ) -> list[dict]:
        """Retrieve recent messages from a Teams channel.

        Args:
            team_id: The unique identifier of the team.
            channel_id: The unique identifier of the channel.
            top: Maximum number of messages to return (default 20, max 50).

        Returns a list of messages with id, createdDateTime, sender,
        body content, and importance.
        """
        top = min(top, 50)
        data = await client.get(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            params={"$top": top},
        )
        messages = data.get("value", [])
        return [
            {
                "id": m.get("id"),
                "createdDateTime": m.get("createdDateTime"),
                "sender": (
                    m.get("from", {}).get("user", {}).get("displayName")
                    if m.get("from")
                    else None
                ),
                "body": m.get("body", {}).get("content", ""),
                "contentType": m.get("body", {}).get("contentType", "text"),
                "importance": m.get("importance", "normal"),
                "subject": m.get("subject"),
            }
            for m in messages
        ]

    @mcp.tool()
    async def send_channel_message(
        team_id: str,
        channel_id: str,
        message: str,
        content_type: str = "text",
    ) -> dict:
        """Send a message to a Microsoft Teams channel.

        Args:
            team_id: The unique identifier of the team.
            channel_id: The unique identifier of the channel.
            message: The text content of the message to send.
            content_type: Either 'text' (default) or 'html'.

        Returns the created message object with id and createdDateTime.
        """
        payload = {
            "body": {
                "contentType": content_type,
                "content": message,
            }
        }
        result = await client.post(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            json=payload,
        )
        return {
            "id": result.get("id"),
            "createdDateTime": result.get("createdDateTime"),
            "webUrl": result.get("webUrl"),
        }

    @mcp.tool()
    async def create_channel(
        team_id: str,
        display_name: str,
        description: str = "",
        membership_type: str = "standard",
    ) -> dict:
        """Create a new channel in a Microsoft Team.

        Args:
            team_id: The unique identifier of the team.
            display_name: The display name for the new channel.
            description: An optional description for the channel.
            membership_type: Channel type - 'standard' (default) or 'private'.

        Returns the created channel with id and displayName.
        """
        payload = {
            "displayName": display_name,
            "description": description,
            "membershipType": membership_type,
        }
        result = await client.post(f"/teams/{team_id}/channels", json=payload)
        return {
            "id": result.get("id"),
            "displayName": result.get("displayName"),
            "description": result.get("description", ""),
            "membershipType": result.get("membershipType"),
        }

    @mcp.tool()
    async def list_team_members(team_id: str) -> list[dict]:
        """List all members of a Microsoft Team.

        Args:
            team_id: The unique identifier of the team.

        Returns a list of members with id, displayName, email, and roles.
        """
        data = await client.get(f"/teams/{team_id}/members")
        members = data.get("value", [])
        return [
            {
                "id": m.get("id"),
                "displayName": m.get("displayName"),
                "email": m.get("email"),
                "roles": m.get("roles", []),
            }
            for m in members
        ]

    @mcp.tool()
    async def list_chats() -> list[dict]:
        """List all chats (1:1 and group) for the current user.

        Returns a list of chats with id, chatType, topic, and lastUpdatedDateTime.
        """
        data = await client.get("/me/chats", params={"$expand": "members"})
        chats = data.get("value", [])
        return [
            {
                "id": c.get("id"),
                "chatType": c.get("chatType"),
                "topic": c.get("topic"),
                "lastUpdatedDateTime": c.get("lastUpdatedDateTime"),
                "memberCount": len(c.get("members", [])),
            }
            for c in chats
        ]

    @mcp.tool()
    async def get_chat_messages(
        chat_id: str,
        top: int = 20,
    ) -> list[dict]:
        """Retrieve recent messages from a Teams chat (1:1 or group).

        Args:
            chat_id: The unique identifier of the chat.
            top: Maximum number of messages to return (default 20, max 50).

        Returns a list of messages with id, createdDateTime, sender,
        and body content.
        """
        top = min(top, 50)
        data = await client.get(
            f"/me/chats/{chat_id}/messages",
            params={"$top": top},
        )
        messages = data.get("value", [])
        return [
            {
                "id": m.get("id"),
                "createdDateTime": m.get("createdDateTime"),
                "sender": (
                    m.get("from", {}).get("user", {}).get("displayName")
                    if m.get("from")
                    else None
                ),
                "body": m.get("body", {}).get("content", ""),
                "contentType": m.get("body", {}).get("contentType", "text"),
            }
            for m in messages
        ]

    @mcp.tool()
    async def send_chat_message(
        chat_id: str,
        message: str,
        content_type: str = "text",
    ) -> dict:
        """Send a message to a Teams chat (1:1 or group).

        Args:
            chat_id: The unique identifier of the chat.
            message: The text content of the message to send.
            content_type: Either 'text' (default) or 'html'.

        Returns the created message object with id and createdDateTime.
        """
        payload = {
            "body": {
                "contentType": content_type,
                "content": message,
            }
        }
        result = await client.post(
            f"/me/chats/{chat_id}/messages",
            json=payload,
        )
        return {
            "id": result.get("id"),
            "createdDateTime": result.get("createdDateTime"),
        }
