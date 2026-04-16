"""Microsoft 365 Users / People tools for the MS Graph MCP server."""

from __future__ import annotations

from typing import Optional

from mcp.server.fastmcp import FastMCP

from ..graph_client import GraphClient


def register_users_tools(mcp: FastMCP, client: GraphClient) -> None:
    """Register all Users-related MCP tools onto *mcp*."""

    @mcp.tool()
    async def get_current_user() -> dict:
        """Get the profile of the currently authenticated user.

        Returns displayName, email, jobTitle, department, officeLocation,
        mobilePhone, and businessPhones.
        """
        user = await client.get(
            "/me",
            params={
                "$select": (
                    "id,displayName,mail,userPrincipalName,jobTitle,"
                    "department,officeLocation,mobilePhone,businessPhones"
                )
            },
        )
        return {
            "id": user.get("id"),
            "displayName": user.get("displayName"),
            "email": user.get("mail") or user.get("userPrincipalName"),
            "jobTitle": user.get("jobTitle"),
            "department": user.get("department"),
            "officeLocation": user.get("officeLocation"),
            "mobilePhone": user.get("mobilePhone"),
            "businessPhones": user.get("businessPhones", []),
        }

    @mcp.tool()
    async def get_user(user_id_or_email: str) -> dict:
        """Get the profile of a specific user by their ID or email address.

        Args:
            user_id_or_email: The user's Azure AD object ID or email address.

        Returns displayName, email, jobTitle, department, officeLocation,
        and phone numbers.
        """
        user = await client.get(
            f"/users/{user_id_or_email}",
            params={
                "$select": (
                    "id,displayName,mail,userPrincipalName,jobTitle,"
                    "department,officeLocation,mobilePhone,businessPhones"
                )
            },
        )
        return {
            "id": user.get("id"),
            "displayName": user.get("displayName"),
            "email": user.get("mail") or user.get("userPrincipalName"),
            "jobTitle": user.get("jobTitle"),
            "department": user.get("department"),
            "officeLocation": user.get("officeLocation"),
            "mobilePhone": user.get("mobilePhone"),
            "businessPhones": user.get("businessPhones", []),
        }

    @mcp.tool()
    async def search_users(
        query: str,
        top: int = 10,
    ) -> list[dict]:
        """Search for users in the organization by name or email.

        Args:
            query: Search string matched against displayName, email, and
                userPrincipalName.
            top: Maximum number of results to return (default 10, max 25).

        Returns a list of matching users with id, displayName, email,
        jobTitle, and department.
        """
        top = min(top, 25)
        data = await client.get(
            "/users",
            params={
                "$search": f'"displayName:{query}" OR "mail:{query}"',
                "$top": top,
                "$select": (
                    "id,displayName,mail,userPrincipalName,jobTitle,department"
                ),
                "$orderby": "displayName",
                "ConsistencyLevel": "eventual",
            },
        )
        users = data.get("value", [])
        return [
            {
                "id": u.get("id"),
                "displayName": u.get("displayName"),
                "email": u.get("mail") or u.get("userPrincipalName"),
                "jobTitle": u.get("jobTitle"),
                "department": u.get("department"),
            }
            for u in users
        ]

    @mcp.tool()
    async def get_user_presence(user_id: str) -> dict:
        """Get the presence status of a user (available, busy, away, etc.).

        Args:
            user_id: The Azure AD object ID of the user.

        Returns availability, activity, and status message.
        """
        presence = await client.get(f"/users/{user_id}/presence")
        return {
            "userId": presence.get("id"),
            "availability": presence.get("availability"),
            "activity": presence.get("activity"),
            "statusMessage": (
                presence.get("statusMessage", {}).get("message", {}).get("content")
                if presence.get("statusMessage")
                else None
            ),
        }

    @mcp.tool()
    async def get_my_presence() -> dict:
        """Get the presence status of the currently authenticated user.

        Returns availability, activity, and status message.
        """
        presence = await client.get("/me/presence")
        return {
            "availability": presence.get("availability"),
            "activity": presence.get("activity"),
            "statusMessage": (
                presence.get("statusMessage", {}).get("message", {}).get("content")
                if presence.get("statusMessage")
                else None
            ),
        }

    @mcp.tool()
    async def list_my_manager_and_reports() -> dict:
        """Get the current user's manager and direct reports.

        Returns manager info and a list of direct reports with displayName,
        email, and jobTitle.
        """
        try:
            manager_data = await client.get(
                "/me/manager",
                params={"$select": "id,displayName,mail,jobTitle"},
            )
            manager = {
                "id": manager_data.get("id"),
                "displayName": manager_data.get("displayName"),
                "email": manager_data.get("mail"),
                "jobTitle": manager_data.get("jobTitle"),
            }
        except Exception:
            manager = None

        reports_data = await client.get(
            "/me/directReports",
            params={"$select": "id,displayName,mail,jobTitle"},
        )
        reports = [
            {
                "id": r.get("id"),
                "displayName": r.get("displayName"),
                "email": r.get("mail"),
                "jobTitle": r.get("jobTitle"),
            }
            for r in reports_data.get("value", [])
        ]

        return {"manager": manager, "directReports": reports}
