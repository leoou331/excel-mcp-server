"""Security-focused regression tests."""

import tempfile
from pathlib import Path

import pytest

from excel_mcp_server import add_allowed_directory
from excel_mcp_server.config import settings
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
        ("='cmd|/C calc'!A0", "blocked syntax"),
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


def test_formula_validator_allows_pipe_inside_double_quoted_string_literal():
    """Quoted string content should not trigger the DDE pipe detector."""
    is_valid, error = FormulaValidator.validate_formula('="cmd|/C calc"')
    assert is_valid
    assert error == ""


@pytest.mark.parametrize("path", ["../escape.xlsx", "..\\escape.xlsx"])
def test_path_validator_rejects_traversal_segments(path):
    """Relative parent traversal must be rejected."""
    is_valid, error = PathValidator.validate_path(path, check_allowed=False)
    assert not is_valid
    assert "traversal" in error.lower()


def test_path_validator_rejects_allowed_directory_prefix_bypass():
    """Sibling directories with a shared prefix are not inside the allowlist."""
    original_allowed_directories = list(settings.security.allowed_directories)
    settings.security.allowed_directories.clear()

    try:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            allowed_dir = root / "allowed"
            bypass_dir = root / "allowed_evil"
            allowed_dir.mkdir()
            bypass_dir.mkdir()
            add_allowed_directory(allowed_dir)

            is_valid, error = PathValidator.validate_path(
                bypass_dir / "escape.xlsx",
                check_allowed=True,
            )

            assert not is_valid
            assert "allowed directories" in error.lower()
    finally:
        settings.security.allowed_directories[:] = original_allowed_directories
