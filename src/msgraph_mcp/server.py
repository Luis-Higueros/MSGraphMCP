"""Main entry point for the Microsoft Graph MCP server.

Run with::

    python -m msgraph_mcp.server          # stdio transport (default for Claude Desktop)
    python -m msgraph_mcp.server --sse    # SSE transport (for web clients)

Or install the package and run::

    msgraph-mcp
"""

from __future__ import annotations

import argparse
import os
import sys

from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

from .auth import GraphAuthProvider
from .graph_client import GraphClient
from .tools.calendar import register_calendar_tools
from .tools.files import register_files_tools
from .tools.mail import register_mail_tools
from .tools.teams import register_teams_tools
from .tools.users import register_users_tools

# Load .env file if present (does nothing if file is missing)
load_dotenv()


def create_server() -> tuple[FastMCP, GraphClient]:
    """Instantiate and configure the MCP server with all Graph tools."""
    auth_provider = GraphAuthProvider()
    graph_client = GraphClient(auth_provider)

    mcp = FastMCP(
        "MS Graph MCP",
        instructions=(
            "This MCP server exposes Microsoft Graph API capabilities for "
            "Microsoft Teams, Outlook Mail, Calendar, OneDrive/SharePoint, "
            "and user directory lookups.  All tools require valid Azure AD "
            "credentials configured via environment variables (see README)."
        ),
    )

    # Register tool groups – the graph_client is used as an async context
    # manager inside each tool.  We pass a lazy wrapper so tools open the
    # HTTP session on first use.
    register_teams_tools(mcp, graph_client)
    register_calendar_tools(mcp, graph_client)
    register_mail_tools(mcp, graph_client)
    register_users_tools(mcp, graph_client)
    register_files_tools(mcp, graph_client)

    return mcp, graph_client


def main() -> None:
    """CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Microsoft Graph MCP Server"
    )
    parser.add_argument(
        "--sse",
        action="store_true",
        default=False,
        help="Run with SSE transport instead of stdio (default: stdio).",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port for SSE transport (default: 8000).",
    )
    args = parser.parse_args()

    mcp, _client = create_server()

    if args.sse:
        import uvicorn  # type: ignore[import]

        uvicorn.run(mcp.sse_app(), host="0.0.0.0", port=args.port)
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
