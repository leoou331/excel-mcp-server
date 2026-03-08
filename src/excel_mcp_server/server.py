"""MCP Server for Excel operations"""

import logging
from typing import Any

from fastmcp import FastMCP

from . import operations
from .config import settings, add_allowed_directory
from .models import (
    CreateWorkbookRequest,
    WorkbookRequest,
    CellReadRequest,
    CellWriteRequest,
    RangeReadRequest,
    RangeWriteRequest,
    FormulaWriteRequest,
    SheetCreateRequest,
    SheetRenameRequest,
    SheetDeleteRequest,
    FontFormatRequest,
    FillFormatRequest,
    NumberFormatRequest,
)

# Configure logging
logging.basicConfig(
    level=getattr(logging, settings.log_level),
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Create FastMCP server
mcp = FastMCP(
    "Excel MCP Server",
    instructions="A robust MCP server for Excel file manipulation"
)


# ============ Workbook Operations ============

@mcp.tool()
def create_workbook(
    file_path: str,
    sheet_name: str | None = None
) -> dict[str, Any]:
    """Create a new Excel workbook.
    
    Args:
        file_path: Path for the new workbook
        sheet_name: Optional name for the first sheet
    """
    request = CreateWorkbookRequest(file_path=file_path, sheet_name=sheet_name)
    result = operations.workbook.create_workbook(request)
    return result.model_dump()


@mcp.tool()
def get_workbook_info(file_path: str) -> dict[str, Any]:
    """Get information about an Excel workbook.
    
    Args:
        file_path: Path to the workbook
    """
    request = WorkbookRequest(file_path=file_path)
    result = operations.workbook.get_workbook_info(request)
    return result.model_dump()


@mcp.tool()
def list_sheets(file_path: str) -> dict[str, Any]:
    """List all sheets in an Excel workbook.
    
    Args:
        file_path: Path to the workbook
    """
    request = WorkbookRequest(file_path=file_path)
    result = operations.workbook.list_sheets(request)
    return result.model_dump()


# ============ Sheet Operations ============

@mcp.tool()
def create_sheet(
    file_path: str,
    sheet_name: str,
    index: int | None = None
) -> dict[str, Any]:
    """Create a new sheet in a workbook.
    
    Args:
        file_path: Path to the workbook
        sheet_name: Name for the new sheet
        index: Optional position to insert (0-based)
    """
    request = SheetCreateRequest(file_path=file_path, sheet_name=sheet_name, index=index)
    result = operations.sheet.create_sheet(request)
    return result.model_dump()


@mcp.tool()
def delete_sheet(file_path: str, sheet_name: str) -> dict[str, Any]:
    """Delete a sheet from a workbook.
    
    Args:
        file_path: Path to the workbook
        sheet_name: Name of the sheet to delete
    """
    request = SheetDeleteRequest(file_path=file_path, sheet_name=sheet_name)
    result = operations.sheet.delete_sheet(request)
    return result.model_dump()


@mcp.tool()
def rename_sheet(
    file_path: str,
    old_name: str,
    new_name: str
) -> dict[str, Any]:
    """Rename a sheet in a workbook.
    
    Args:
        file_path: Path to the workbook
        old_name: Current name of the sheet
        new_name: New name for the sheet
    """
    request = SheetRenameRequest(file_path=file_path, old_name=old_name, new_name=new_name)
    result = operations.sheet.rename_sheet(request)
    return result.model_dump()


# ============ Cell Operations ============

@mcp.tool()
def read_cell(
    file_path: str,
    sheet_name: str,
    cell: str
) -> dict[str, Any]:
    """Read a value from a cell.
    
    Args:
        file_path: Path to the workbook
        sheet_name: Name of the sheet
        cell: Cell reference (e.g., 'A1')
    """
    request = CellReadRequest(file_path=file_path, sheet_name=sheet_name, cell=cell)
    result = operations.cell.read_cell(request)
    return result.model_dump()


@mcp.tool()
def write_cell(
    file_path: str,
    sheet_name: str,
    cell: str,
    value: Any
) -> dict[str, Any]:
    """Write a value to a cell.
    
    Args:
        file_path: Path to the workbook
        sheet_name: Name of the sheet
        cell: Cell reference (e.g., 'A1')
        value: Value to write
    """
    request = CellWriteRequest(file_path=file_path, sheet_name=sheet_name, cell=cell, value=value)
    result = operations.cell.write_cell(request)
    return result.model_dump()


@mcp.tool()
def read_range(
    file_path: str,
    sheet_name: str,
    range_ref: str
) -> dict[str, Any]:
    """Read data from a range of cells.
    
    Args:
        file_path: Path to the workbook
        sheet_name: Name of the sheet
        range_ref: Range reference (e.g., 'A1:D10')
    """
    request = RangeReadRequest(file_path=file_path, sheet_name=sheet_name, range_ref=range_ref)
    result = operations.cell.read_range(request)
    return result.model_dump()


@mcp.tool()
def write_range(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    data: list[list[Any]]
) -> dict[str, Any]:
    """Write data to a range of cells.
    
    Args:
        file_path: Path to the workbook
        sheet_name: Name of the sheet
        start_cell: Top-left cell of the range (e.g., 'A1')
        data: 2D array of values
    """
    request = RangeWriteRequest(file_path=file_path, sheet_name=sheet_name, start_cell=start_cell, data=data)
    result = operations.cell.write_range(request)
    return result.model_dump()


@mcp.tool()
def write_formula(
    file_path: str,
    sheet_name: str,
    cell: str,
    formula: str
) -> dict[str, Any]:
    """Write a formula to a cell.
    
    Args:
        file_path: Path to the workbook
        sheet_name: Name of the sheet
        cell: Cell reference (e.g., 'A1')
        formula: Excel formula (e.g., '=SUM(B1:B10)')
    """
    request = FormulaWriteRequest(file_path=file_path, sheet_name=sheet_name, cell=cell, formula=formula)
    result = operations.cell.write_formula(request)
    return result.model_dump()


# ============ Formatting Operations ============

@mcp.tool()
def format_font(
    file_path: str,
    sheet_name: str,
    range_ref: str,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: int | None = None,
    color: str | None = None
) -> dict[str, Any]:
    """Apply font formatting to a range.
    
    Args:
        file_path: Path to the workbook
        sheet_name: Name of the sheet
        range_ref: Range reference (e.g., 'A1:D10')
        bold: Make text bold
        italic: Make text italic
        font_size: Font size (8-72)
        color: Hex color code (e.g., 'FF0000')
    """
    request = FontFormatRequest(
        file_path=file_path,
        sheet_name=sheet_name,
        range_ref=range_ref,
        bold=bold,
        italic=italic,
        font_size=font_size,
        color=color
    )
    result = operations.formatting.format_font(request)
    return result.model_dump()


@mcp.tool()
def format_fill(
    file_path: str,
    sheet_name: str,
    range_ref: str,
    color: str
) -> dict[str, Any]:
    """Apply fill (background color) formatting to a range.
    
    Args:
        file_path: Path to the workbook
        sheet_name: Name of the sheet
        range_ref: Range reference (e.g., 'A1:D10')
        color: Hex color code (e.g., 'FFFF00')
    """
    request = FillFormatRequest(
        file_path=file_path,
        sheet_name=sheet_name,
        range_ref=range_ref,
        color=color
    )
    result = operations.formatting.format_fill(request)
    return result.model_dump()


@mcp.tool()
def format_number(
    file_path: str,
    sheet_name: str,
    range_ref: str,
    format_string: str
) -> dict[str, Any]:
    """Apply number formatting to a range.
    
    Args:
        file_path: Path to the workbook
        sheet_name: Name of the sheet
        range_ref: Range reference (e.g., 'A1:D10')
        format_string: Excel number format (e.g., '#,##0.00', '0%', 'mm/dd/yyyy')
    """
    request = NumberFormatRequest(
        file_path=file_path,
        sheet_name=sheet_name,
        range_ref=range_ref,
        format_string=format_string
    )
    result = operations.formatting.format_number(request)
    return result.model_dump()


# ============ Configuration ============

@mcp.tool()
def add_allowed_directory_tool(directory: str) -> dict[str, Any]:
    """Add a directory to the allowed directories whitelist.
    
    Args:
        directory: Directory path to allow
    """
    try:
        add_allowed_directory(directory)
        return {"success": True, "message": f"Added directory: {directory}"}
    except Exception as e:
        return {"success": False, "message": str(e)}


def main():
    """Run the MCP server."""
    logger.info("Starting Excel MCP Server")
    mcp.run()


if __name__ == "__main__":
    main()
