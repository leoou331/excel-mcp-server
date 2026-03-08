"""Configuration for Excel MCP Server"""

from pathlib import Path
from typing import Optional

from pydantic import BaseModel, Field


class SecurityConfig(BaseModel):
    """Security configuration"""
    allowed_directories: list[Path] = Field(default_factory=list)
    max_file_size_bytes: int = 100 * 1024 * 1024  # 100MB
    max_cells_per_operation: int = 100_000
    max_rows: int = 1_048_576
    max_columns: int = 16_384  # XFD
    max_formula_length: int = 8192
    max_sheet_name_length: int = 31
    forbidden_sheet_chars: set[str] = Field(
        default_factory=lambda: set('[]:/\\?*')
    )
    blocked_formula_patterns: list[str] = Field(default_factory=lambda: [
        r'=.*cmd\|',
        r'=.*WEBSERVICE\(',
        r'=.*HYPERLINK\(',
        r'=.*EXECUTE\(',
        r'=.*CALL\(',
        r'=.*REGISTER\(',
    ])
    lock_timeout_seconds: int = 30
    lock_stale_seconds: int = 300


class Settings(BaseModel):
    """Application settings"""
    security: SecurityConfig = Field(default_factory=SecurityConfig)
    log_level: str = "INFO"
    default_timeout: int = 300

    class Config:
        env_prefix = "EXCEL_MCP_"


settings = Settings()


def add_allowed_directory(path: str | Path) -> None:
    """Add a directory to the allowed list"""
    p = Path(path).resolve()
    if p not in settings.security.allowed_directories:
        settings.security.allowed_directories.append(p)


def is_path_allowed(path: str | Path) -> bool:
    """Check if a path is within allowed directories"""
    if not settings.security.allowed_directories:
        return True
    resolved = Path(path).resolve()
    return any(
        resolved.is_relative_to(allowed)
        for allowed in settings.security.allowed_directories
    )
