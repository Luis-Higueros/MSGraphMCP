"""Shared fixtures for the test suite."""

import json
import pytest
from unittest.mock import MagicMock

from msgraph_mcp.graph_client import GraphClient
from msgraph_mcp.auth import GraphAuthProvider


@pytest.fixture
def mock_auth():
    """Return a mock GraphAuthProvider that provides a dummy token."""
    auth = MagicMock(spec=GraphAuthProvider)
    auth.get_access_token.return_value = "dummy-token"
    return auth


@pytest.fixture
def mock_client(mock_auth):
    """Return a GraphClient backed by the mock auth provider."""
    return GraphClient(mock_auth)


def parse_tool_result(result):
    """Parse the payload returned by FastMCP.call_tool().

    FastMCP.call_tool() has two return shapes:
    * For tools returning a single dict: a list of TextContent objects.
    * For tools returning a list:        a tuple of (list[TextContent], {'result': [...]}).

    This helper normalises both cases and returns the deserialized Python value.
    """
    assert result, "Expected non-empty tool result"
    # Tuple form: (list[TextContent], {'result': [...]})
    if isinstance(result, tuple):
        return result[1]["result"]
    # List form: [TextContent(...)]
    return json.loads(result[0].text)
