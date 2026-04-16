"""Microsoft 365 Mail tools for the MS Graph MCP server."""

from __future__ import annotations

from typing import Optional

from mcp.server.fastmcp import FastMCP

from ..graph_client import GraphClient


def register_mail_tools(mcp: FastMCP, client: GraphClient) -> None:
    """Register all Mail-related MCP tools onto *mcp*."""

    @mcp.tool()
    async def list_emails(
        folder: str = "inbox",
        top: int = 10,
        search: Optional[str] = None,
        unread_only: bool = False,
    ) -> list[dict]:
        """List emails in a mail folder.

        Args:
            folder: Mail folder to list (e.g. 'inbox', 'sentitems',
                'drafts', 'deleteditems'). Default is 'inbox'.
            top: Maximum number of emails to return (default 10, max 50).
            search: Optional keyword search query.
            unread_only: If True, return only unread emails.

        Returns a list of emails with id, subject, from, receivedDateTime,
        bodyPreview, and isRead.
        """
        top = min(top, 50)
        params: dict = {
            "$top": top,
            "$orderby": "receivedDateTime desc",
            "$select": (
                "id,subject,from,receivedDateTime,"
                "bodyPreview,isRead,hasAttachments,importance"
            ),
        }
        if search:
            params["$search"] = f'"{search}"'
        elif unread_only:
            params["$filter"] = "isRead eq false"

        data = await client.get(
            f"/me/mailFolders/{folder}/messages", params=params
        )
        messages = data.get("value", [])
        return [
            {
                "id": m.get("id"),
                "subject": m.get("subject"),
                "from": (
                    m.get("from", {}).get("emailAddress", {}).get("name")
                    or m.get("from", {}).get("emailAddress", {}).get("address")
                ),
                "fromEmail": (
                    m.get("from", {}).get("emailAddress", {}).get("address")
                ),
                "receivedDateTime": m.get("receivedDateTime"),
                "bodyPreview": m.get("bodyPreview", ""),
                "isRead": m.get("isRead", True),
                "hasAttachments": m.get("hasAttachments", False),
                "importance": m.get("importance", "normal"),
            }
            for m in messages
        ]

    @mcp.tool()
    async def get_email(message_id: str) -> dict:
        """Get the full content of a specific email message.

        Args:
            message_id: The unique identifier of the email message.

        Returns the full email with subject, from, to, cc, body,
        receivedDateTime, and attachment info.
        """
        msg = await client.get(
            f"/me/messages/{message_id}",
            params={"$select": "id,subject,from,toRecipients,ccRecipients,body,receivedDateTime,isRead,hasAttachments,importance"},
        )
        to_recipients = [
            r.get("emailAddress", {}).get("address")
            for r in msg.get("toRecipients", [])
        ]
        cc_recipients = [
            r.get("emailAddress", {}).get("address")
            for r in msg.get("ccRecipients", [])
        ]
        return {
            "id": msg.get("id"),
            "subject": msg.get("subject"),
            "from": msg.get("from", {}).get("emailAddress", {}).get("name"),
            "fromEmail": msg.get("from", {}).get("emailAddress", {}).get("address"),
            "to": to_recipients,
            "cc": cc_recipients,
            "receivedDateTime": msg.get("receivedDateTime"),
            "body": msg.get("body", {}).get("content", ""),
            "contentType": msg.get("body", {}).get("contentType", "text"),
            "isRead": msg.get("isRead", True),
            "hasAttachments": msg.get("hasAttachments", False),
            "importance": msg.get("importance", "normal"),
        }

    @mcp.tool()
    async def send_email(
        to_emails: list[str],
        subject: str,
        body: str,
        content_type: str = "html",
        cc_emails: Optional[list[str]] = None,
        bcc_emails: Optional[list[str]] = None,
        importance: str = "normal",
    ) -> dict:
        """Send an email message.

        Args:
            to_emails: List of recipient email addresses.
            subject: The subject line of the email.
            body: The body content of the email.
            content_type: Either 'html' (default) or 'text'.
            cc_emails: Optional list of CC recipient email addresses.
            bcc_emails: Optional list of BCC recipient email addresses.
            importance: Message importance: 'low', 'normal' (default), or 'high'.

        Returns a confirmation dict with status 'sent'.
        """

        def _recipients(emails: list[str]) -> list[dict]:
            return [{"emailAddress": {"address": e}} for e in emails]

        message: dict = {
            "subject": subject,
            "importance": importance,
            "body": {"contentType": content_type, "content": body},
            "toRecipients": _recipients(to_emails),
        }
        if cc_emails:
            message["ccRecipients"] = _recipients(cc_emails)
        if bcc_emails:
            message["bccRecipients"] = _recipients(bcc_emails)

        await client.post("/me/sendMail", json={"message": message})
        return {"status": "sent", "to": to_emails, "subject": subject}

    @mcp.tool()
    async def reply_to_email(
        message_id: str,
        body: str,
        content_type: str = "html",
        reply_all: bool = False,
    ) -> dict:
        """Reply to an existing email message.

        Args:
            message_id: The unique identifier of the email to reply to.
            body: The body content of the reply.
            content_type: Either 'html' (default) or 'text'.
            reply_all: If True, reply to all recipients; otherwise reply only to sender.

        Returns a confirmation dict with status 'sent'.
        """
        endpoint = "replyAll" if reply_all else "reply"
        payload = {
            "comment": body,
        }
        await client.post(
            f"/me/messages/{message_id}/{endpoint}", json=payload
        )
        return {"status": "sent", "replyAll": reply_all}

    @mcp.tool()
    async def move_email(
        message_id: str,
        destination_folder: str,
    ) -> dict:
        """Move an email to a different folder.

        Args:
            message_id: The unique identifier of the email message.
            destination_folder: The destination folder name or well-known folder
                name (e.g. 'inbox', 'deleteditems', 'archive').

        Returns the moved message id and the destination folder.
        """
        payload = {"destinationId": destination_folder}
        result = await client.post(
            f"/me/messages/{message_id}/move", json=payload
        )
        return {
            "id": result.get("id"),
            "destinationFolder": destination_folder,
        }
