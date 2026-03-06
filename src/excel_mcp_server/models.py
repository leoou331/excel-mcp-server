"""Pydantic models for Excel MCP Server"""

import logging
from typing import Any, Optional

from pydantic import BaseModel, Field, field_validator

from .utils.security import CellValidator, FormulaValidator, PathValidator, DataValidator

logger = logging.getLogger(__name__)


# ============ Base Models ============

class OperationResult(BaseModel):
    """Base result for all operations"""
    success: bool
    message: str = ""
    error_code: Optional[str] = None


class WorkbookResult(OperationResult):
    """Result for workbook operations"""
    file_path: Optional[str] = None
    sheets: Optional[list[str]] = None
    sheet_count: Optional[int] = None
    file_size: Optional[int] = None


class SheetResult(OperationResult):
    """Result for sheet operations"""
    sheet_name: Optional[str] = None


class CellResult(OperationResult):
    """Result for cell operations"""
    cell: Optional[str] = None
    value: Optional[Any] = None


class RangeResult(OperationResult):
    """Result for range operations"""
    range_ref: Optional[str] = None
    rows: Optional[int] = None
    cols: Optional[int] = None
    data: Optional[list[list[Any]]] = None


# ============ Request Models ============

class CreateWorkbookRequest(BaseModel):
    """Request to create a new workbook"""
    file_path: str
    sheet_name: Optional[str] = None
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=False)
        if not is_valid:
            raise ValueError(error)
        return v
    
    @field_validator('sheet_name')
    @classmethod
    def validate_sheet_name(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return v
        is_valid, error = PathValidator.validate_sheet_name(v)
        if not is_valid:
            raise ValueError(error)
        return v


class WorkbookRequest(BaseModel):
    """Request for workbook operations"""
    file_path: str
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v


class CellReadRequest(BaseModel):
    """Request to read a cell"""
    file_path: str
    sheet_name: str
    cell: str
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v
    
    @field_validator('cell')
    @classmethod
    def validate_cell(cls, v: str) -> str:
        is_valid, error, _ = CellValidator.validate_cell(v)
        if not is_valid:
            raise ValueError(error)
        return v.upper()


class CellWriteRequest(BaseModel):
    """Request to write to a cell"""
    file_path: str
    sheet_name: str
    cell: str
    value: Any
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v
    
    @field_validator('cell')
    @classmethod
    def validate_cell(cls, v: str) -> str:
        is_valid, error, _ = CellValidator.validate_cell(v)
        if not is_valid:
            raise ValueError(error)
        return v.upper()
    
    @field_validator('value')
    @classmethod
    def validate_value(cls, v: Any) -> Any:
        is_valid, error = DataValidator.validate_value(v)
        if not is_valid:
            raise ValueError(error)
        return v


class RangeReadRequest(BaseModel):
    """Request to read a range"""
    file_path: str
    sheet_name: str
    range_ref: str
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v
    
    @field_validator('range_ref')
    @classmethod
    def validate_range(cls, v: str) -> str:
        is_valid, error, _ = CellValidator.validate_range(v)
        if not is_valid:
            raise ValueError(error)
        return v.upper()


class RangeWriteRequest(BaseModel):
    """Request to write to a range"""
    file_path: str
    sheet_name: str
    start_cell: str
    data: list[list[Any]]
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v
    
    @field_validator('start_cell')
    @classmethod
    def validate_start_cell(cls, v: str) -> str:
        is_valid, error, _ = CellValidator.validate_cell(v)
        if not is_valid:
            raise ValueError(error)
        return v.upper()
    
    @field_validator('data')
    @classmethod
    def validate_data(cls, v: list[list[Any]]) -> list[list[Any]]:
        is_valid, error = DataValidator.validate_data_size(v)
        if not is_valid:
            raise ValueError(error)
        # Validate each value
        for row in v:
            for val in row:
                is_valid, error = DataValidator.validate_value(val)
                if not is_valid:
                    raise ValueError(f"Invalid value in data: {error}")
        return v


class FormulaWriteRequest(BaseModel):
    """Request to write a formula"""
    file_path: str
    sheet_name: str
    cell: str
    formula: str
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v
    
    @field_validator('cell')
    @classmethod
    def validate_cell(cls, v: str) -> str:
        is_valid, error, _ = CellValidator.validate_cell(v)
        if not is_valid:
            raise ValueError(error)
        return v.upper()
    
    @field_validator('formula')
    @classmethod
    def validate_formula(cls, v: str) -> str:
        is_valid, error = FormulaValidator.validate_formula(v)
        if not is_valid:
            raise ValueError(error)
        return v


class SheetCreateRequest(BaseModel):
    """Request to create a new sheet"""
    file_path: str
    sheet_name: str
    index: Optional[int] = None
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v
    
    @field_validator('sheet_name')
    @classmethod
    def validate_sheet_name(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_sheet_name(v)
        if not is_valid:
            raise ValueError(error)
        return v


class SheetRenameRequest(BaseModel):
    """Request to rename a sheet"""
    file_path: str
    old_name: str
    new_name: str
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v
    
    @field_validator('new_name')
    @classmethod
    def validate_new_name(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_sheet_name(v)
        if not is_valid:
            raise ValueError(error)
        return v


class SheetDeleteRequest(BaseModel):
    """Request to delete a sheet"""
    file_path: str
    sheet_name: str
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v


# ============ Formatting Models ============

class FontFormatRequest(BaseModel):
    """Request to format font"""
    file_path: str
    sheet_name: str
    range_ref: str
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_size: Optional[int] = Field(None, ge=8, le=72)
    color: Optional[str] = None
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v
    
    @field_validator('range_ref')
    @classmethod
    def validate_range(cls, v: str) -> str:
        is_valid, error, _ = CellValidator.validate_range(v)
        if not is_valid:
            raise ValueError(error)
        return v.upper()
    
    @field_validator('color')
    @classmethod
    def validate_color(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return v
        v = v.lstrip('#')
        if len(v) != 6 or not all(c in '0123456789ABCDEFabcdef' for c in v):
            raise ValueError("Invalid hex color")
        return v.upper()


class FillFormatRequest(BaseModel):
    """Request to format cell fill"""
    file_path: str
    sheet_name: str
    range_ref: str
    color: str
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v
    
    @field_validator('range_ref')
    @classmethod
    def validate_range(cls, v: str) -> str:
        is_valid, error, _ = CellValidator.validate_range(v)
        if not is_valid:
            raise ValueError(error)
        return v.upper()
    
    @field_validator('color')
    @classmethod
    def validate_color(cls, v: str) -> str:
        v = v.lstrip('#')
        if len(v) != 6 or not all(c in '0123456789ABCDEFabcdef' for c in v):
            raise ValueError("Invalid hex color")
        return v.upper()


class NumberFormatRequest(BaseModel):
    """Request to format numbers"""
    file_path: str
    sheet_name: str
    range_ref: str
    format_string: str
    
    @field_validator('file_path')
    @classmethod
    def validate_path(cls, v: str) -> str:
        is_valid, error = PathValidator.validate_path(v, must_exist=True)
        if not is_valid:
            raise ValueError(error)
        return v
    
    @field_validator('range_ref')
    @classmethod
    def validate_range(cls, v: str) -> str:
        is_valid, error, _ = CellValidator.validate_range(v)
        if not is_valid:
            raise ValueError(error)
        return v.upper()
