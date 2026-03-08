"""Server initialization tests."""

import importlib

from fastmcp import FastMCP


def test_server_module_imports_with_supported_fastmcp_api() -> None:
    """Importing the server should succeed with the installed FastMCP version."""
    module = importlib.import_module("excel_mcp_server.server")

    assert isinstance(module.mcp, FastMCP)
