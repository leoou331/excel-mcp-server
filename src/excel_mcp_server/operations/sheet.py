"""Sheet operations"""

import logging
from pathlib import Path

from ..models import SheetCreateRequest, SheetDeleteRequest, SheetRenameRequest, SheetResult
from ..utils.file_manager import file_manager

logger = logging.getLogger(__name__)


def create_sheet(request: SheetCreateRequest) -> SheetResult:
    """Create a new sheet in a workbook."""
    try:
        path = Path(request.file_path).resolve()
        
        with file_manager.safe_write(path) as wb:
            if request.sheet_name in wb.sheetnames:
                return SheetResult(
                    success=False,
                    message=f"Sheet '{request.sheet_name}' already exists",
                    error_code="SHEET_EXISTS"
                )
            
            if request.index is not None:
                wb.create_sheet(title=request.sheet_name, index=request.index)
            else:
                wb.create_sheet(title=request.sheet_name)
        
        return SheetResult(
            success=True,
            message=f"Sheet '{request.sheet_name}' created successfully",
            sheet_name=request.sheet_name
        )
        
    except Exception as e:
        logger.exception(f"Failed to create sheet: {request.sheet_name}")
        return SheetResult(
            success=False,
            message=f"Failed to create sheet: {str(e)}",
            error_code="CREATE_ERROR"
        )


def delete_sheet(request: SheetDeleteRequest) -> SheetResult:
    """Delete a sheet from a workbook."""
    try:
        path = Path(request.file_path).resolve()
        
        with file_manager.safe_write(path) as wb:
            if request.sheet_name not in wb.sheetnames:
                return SheetResult(
                    success=False,
                    message=f"Sheet '{request.sheet_name}' not found. Available: {wb.sheetnames}",
                    error_code="SHEET_NOT_FOUND"
                )
            
            if len(wb.sheetnames) == 1:
                return SheetResult(
                    success=False,
                    message="Cannot delete the last sheet in a workbook",
                    error_code="LAST_SHEET"
                )
            
            del wb[request.sheet_name]
        
        return SheetResult(
            success=True,
            message=f"Sheet '{request.sheet_name}' deleted successfully",
            sheet_name=request.sheet_name
        )
        
    except Exception as e:
        logger.exception(f"Failed to delete sheet: {request.sheet_name}")
        return SheetResult(
            success=False,
            message=f"Failed to delete sheet: {str(e)}",
            error_code="DELETE_ERROR"
        )


def rename_sheet(request: SheetRenameRequest) -> SheetResult:
    """Rename a sheet in a workbook."""
    try:
        path = Path(request.file_path).resolve()
        
        with file_manager.safe_write(path) as wb:
            if request.old_name not in wb.sheetnames:
                return SheetResult(
                    success=False,
                    message=f"Sheet '{request.old_name}' not found. Available: {wb.sheetnames}",
                    error_code="SHEET_NOT_FOUND"
                )
            
            if request.new_name in wb.sheetnames:
                return SheetResult(
                    success=False,
                    message=f"Sheet '{request.new_name}' already exists",
                    error_code="SHEET_EXISTS"
                )
            
            wb[request.old_name].title = request.new_name
        
        return SheetResult(
            success=True,
            message=f"Sheet renamed from '{request.old_name}' to '{request.new_name}'",
            sheet_name=request.new_name
        )
        
    except Exception as e:
        logger.exception(f"Failed to rename sheet: {request.old_name}")
        return SheetResult(
            success=False,
            message=f"Failed to rename sheet: {str(e)}",
            error_code="RENAME_ERROR"
        )
