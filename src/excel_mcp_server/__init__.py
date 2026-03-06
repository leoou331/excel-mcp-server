"""Excel MCP Server - A robust MCP server for Excel file manipulation"""

__version__ = "1.0.0"

from .config import settings, add_allowed_directory, is_path_allowed
from .models import (
    OperationResult,
    WorkbookResult,
    SheetResult,
    CellResult,
    RangeResult,
)
from .operations import workbook, sheet, cell, formatting
