"""Security-focused regression tests."""

import importlib

import pytest

from excel_mcp_server.utils.security import FormulaValidator, PathValidator


def test_server_module_imports_with_fastmcp_3():
    """FastMCP 3.x should accept the server configuration."""
    module = importlib.import_module("excel_mcp_server.server")
    assert module.mcp is not None


@pytest.mark.parametrize(
    ("formula", "expected_error"),
    [
        ('=WEBSERVICE("http://evil.test")', "blocked function"),
        ('=@WEBSERVICE("http://evil.test")', "blocked function"),
        ('=SUM(1, _xlfn.WEBSERVICE("http://evil.test"))', "blocked function"),
        ("=cmd|'/C calc'!A0", "blocked syntax"),
    ],
)
def test_formula_validator_blocks_unsafe_formula_shapes(formula, expected_error):
    """Unsafe formulas should be rejected even with prefixes or DDE syntax."""
    is_valid, error = FormulaValidator.validate_formula(formula)
    assert not is_valid
    assert expected_error in error


def test_formula_validator_ignores_blocked_names_inside_strings():
    """String literals mentioning blocked functions should remain valid."""
    is_valid, error = FormulaValidator.validate_formula('="WEBSERVICE("')
    assert is_valid
    assert error == ""


@pytest.mark.parametrize("path", ["../escape.xlsx", "..\\escape.xlsx"])
def test_path_validator_rejects_traversal_segments(path):
    """Relative parent traversal must be rejected."""
    is_valid, error = PathValidator.validate_path(path, check_allowed=False)
    assert not is_valid
    assert "traversal" in error.lower()
