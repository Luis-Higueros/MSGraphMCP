"""Microsoft 365 Calendar tools for the MS Graph MCP server."""

from __future__ import annotations

from typing import Optional

from mcp.server.fastmcp import FastMCP

from ..graph_client import GraphClient


def register_calendar_tools(mcp: FastMCP, client: GraphClient) -> None:
    """Register all Calendar-related MCP tools onto *mcp*."""

    @mcp.tool()
    async def list_calendar_events(
        top: int = 10,
        start_datetime: Optional[str] = None,
        end_datetime: Optional[str] = None,
    ) -> list[dict]:
        """List upcoming calendar events for the current user.

        Args:
            top: Maximum number of events to return (default 10, max 50).
            start_datetime: ISO 8601 start date/time filter
                (e.g. '2024-01-01T00:00:00'). Defaults to now.
            end_datetime: ISO 8601 end date/time filter
                (e.g. '2024-12-31T23:59:59').

        Returns a list of events with id, subject, start, end, location,
        organizer, and webLink.
        """
        top = min(top, 50)
        params: dict = {
            "$top": top,
            "$orderby": "start/dateTime",
            "$select": (
                "id,subject,start,end,location,organizer,"
                "webLink,isOnlineMeeting,onlineMeetingUrl,bodyPreview"
            ),
        }
        if start_datetime or end_datetime:
            filters = []
            if start_datetime:
                filters.append(f"start/dateTime ge '{start_datetime}'")
            if end_datetime:
                filters.append(f"end/dateTime le '{end_datetime}'")
            params["$filter"] = " and ".join(filters)

        data = await client.get("/me/events", params=params)
        events = data.get("value", [])
        return [
            {
                "id": e.get("id"),
                "subject": e.get("subject"),
                "start": e.get("start", {}).get("dateTime"),
                "startTimeZone": e.get("start", {}).get("timeZone"),
                "end": e.get("end", {}).get("dateTime"),
                "endTimeZone": e.get("end", {}).get("timeZone"),
                "location": e.get("location", {}).get("displayName"),
                "organizer": (
                    e.get("organizer", {})
                    .get("emailAddress", {})
                    .get("name")
                ),
                "webLink": e.get("webLink"),
                "isOnlineMeeting": e.get("isOnlineMeeting", False),
                "onlineMeetingUrl": e.get("onlineMeetingUrl"),
                "bodyPreview": e.get("bodyPreview", ""),
            }
            for e in events
        ]

    @mcp.tool()
    async def get_calendar_event(event_id: str) -> dict:
        """Get full details of a specific calendar event.

        Args:
            event_id: The unique identifier of the calendar event.

        Returns full event details including attendees, body, recurrence,
        and online meeting info.
        """
        event = await client.get(f"/me/events/{event_id}")
        attendees = [
            {
                "name": a.get("emailAddress", {}).get("name"),
                "email": a.get("emailAddress", {}).get("address"),
                "status": a.get("status", {}).get("response"),
            }
            for a in event.get("attendees", [])
        ]
        return {
            "id": event.get("id"),
            "subject": event.get("subject"),
            "bodyPreview": event.get("bodyPreview"),
            "body": event.get("body", {}).get("content"),
            "start": event.get("start", {}).get("dateTime"),
            "end": event.get("end", {}).get("dateTime"),
            "location": event.get("location", {}).get("displayName"),
            "organizer": (
                event.get("organizer", {})
                .get("emailAddress", {})
                .get("name")
            ),
            "attendees": attendees,
            "isOnlineMeeting": event.get("isOnlineMeeting", False),
            "onlineMeetingUrl": event.get("onlineMeetingUrl"),
            "webLink": event.get("webLink"),
        }

    @mcp.tool()
    async def create_calendar_event(
        subject: str,
        start_datetime: str,
        end_datetime: str,
        time_zone: str = "UTC",
        body: str = "",
        location: str = "",
        attendee_emails: Optional[list[str]] = None,
        is_online_meeting: bool = False,
    ) -> dict:
        """Create a new calendar event.

        Args:
            subject: The title/subject of the event.
            start_datetime: ISO 8601 start date/time (e.g. '2024-06-15T14:00:00').
            end_datetime: ISO 8601 end date/time (e.g. '2024-06-15T15:00:00').
            time_zone: IANA or Windows time zone name (default 'UTC').
            body: Optional HTML or plain-text body for the event invitation.
            location: Optional location name or address.
            attendee_emails: Optional list of attendee email addresses.
            is_online_meeting: Whether to create an online Teams meeting link.

        Returns the created event with id, subject, start, end, and webLink.
        """
        payload: dict = {
            "subject": subject,
            "start": {"dateTime": start_datetime, "timeZone": time_zone},
            "end": {"dateTime": end_datetime, "timeZone": time_zone},
            "isOnlineMeeting": is_online_meeting,
        }
        if body:
            payload["body"] = {"contentType": "html", "content": body}
        if location:
            payload["location"] = {"displayName": location}
        if attendee_emails:
            payload["attendees"] = [
                {
                    "emailAddress": {"address": email},
                    "type": "required",
                }
                for email in attendee_emails
            ]

        result = await client.post("/me/events", json=payload)
        return {
            "id": result.get("id"),
            "subject": result.get("subject"),
            "start": result.get("start", {}).get("dateTime"),
            "end": result.get("end", {}).get("dateTime"),
            "webLink": result.get("webLink"),
            "onlineMeetingUrl": result.get("onlineMeetingUrl"),
        }

    @mcp.tool()
    async def update_calendar_event(
        event_id: str,
        subject: Optional[str] = None,
        start_datetime: Optional[str] = None,
        end_datetime: Optional[str] = None,
        time_zone: Optional[str] = None,
        body: Optional[str] = None,
        location: Optional[str] = None,
    ) -> dict:
        """Update an existing calendar event.

        Args:
            event_id: The unique identifier of the calendar event.
            subject: New title/subject (optional).
            start_datetime: New start date/time in ISO 8601 format (optional).
            end_datetime: New end date/time in ISO 8601 format (optional).
            time_zone: Time zone for start/end (required if updating times).
            body: New body/description HTML content (optional).
            location: New location (optional).

        Returns the updated event with id, subject, start, and end.
        """
        payload: dict = {}
        if subject:
            payload["subject"] = subject
        if start_datetime:
            payload["start"] = {
                "dateTime": start_datetime,
                "timeZone": time_zone or "UTC",
            }
        if end_datetime:
            payload["end"] = {
                "dateTime": end_datetime,
                "timeZone": time_zone or "UTC",
            }
        if body is not None:
            payload["body"] = {"contentType": "html", "content": body}
        if location is not None:
            payload["location"] = {"displayName": location}

        result = await client.patch(f"/me/events/{event_id}", json=payload)
        return {
            "id": result.get("id"),
            "subject": result.get("subject"),
            "start": result.get("start", {}).get("dateTime"),
            "end": result.get("end", {}).get("dateTime"),
        }
