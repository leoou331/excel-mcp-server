"""Workbook operations"""

import logging
from pathlib import Path

from ..models import CreateWorkbookRequest, WorkbookRequest, WorkbookResult
from ..utils.file_manager import file_manager

logger = logging.getLogger(__name__)


def create_workbook(request: CreateWorkbookRequest) -> WorkbookResult:
    """Create a new Excel workbook."""
    try:
        path = Path(request.file_path).resolve()
        
        if path.exists():
            return WorkbookResult(
                success=False,
                message=f"File already exists: {path}",
                error_code="FILE_EXISTS"
            )
        
        path.parent.mkdir(parents=True, exist_ok=True)
        
        with file_manager.safe_write(path) as wb:
            if request.sheet_name and wb.active:
                wb.active.title = request.sheet_name
        
        file_size = path.stat().st_size if path.exists() else 0
        
        return WorkbookResult(
            success=True,
            message="Workbook created successfully",
            file_path=str(path),
            sheet_count=1,
            file_size=file_size
        )
        
    except Exception as e:
        logger.exception(f"Failed to create workbook: {request.file_path}")
        return WorkbookResult(
            success=False,
            message=f"Failed to create workbook: {str(e)}",
            error_code="CREATE_ERROR"
        )


def get_workbook_info(request: WorkbookRequest) -> WorkbookResult:
    """Get information about a workbook."""
    try:
        path = Path(request.file_path).resolve()
        
        with file_manager.safe_open(path, read_only=True) as wb:
            sheets = wb.sheetnames
            sheet_count = len(sheets)
        
        file_size = path.stat().st_size
        
        return WorkbookResult(
            success=True,
            message="Workbook info retrieved",
            file_path=str(path),
            sheets=sheets,
            sheet_count=sheet_count,
            file_size=file_size
        )
        
    except Exception as e:
        logger.exception(f"Failed to get workbook info: {request.file_path}")
        return WorkbookResult(
            success=False,
            message=f"Failed to get workbook info: {str(e)}",
            error_code="READ_ERROR"
        )


def list_sheets(request: WorkbookRequest) -> WorkbookResult:
    """List all sheets in a workbook."""
    try:
        path = Path(request.file_path).resolve()
        
        with file_manager.safe_open(path, read_only=True) as wb:
            sheets = wb.sheetnames
        
        return WorkbookResult(
            success=True,
            message="Sheets listed successfully",
            file_path=str(path),
            sheets=sheets,
            sheet_count=len(sheets)
        )
        
    except Exception as e:
        logger.exception(f"Failed to list sheets: {request.file_path}")
        return WorkbookResult(
            success=False,
            message=f"Failed to list sheets: {str(e)}",
            error_code="READ_ERROR"
        )
