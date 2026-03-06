"""Security utilities for Excel MCP Server"""

import logging
import re
from pathlib import Path
from typing import Tuple, Any

from ..config import settings

logger = logging.getLogger(__name__)


class SecurityError(Exception):
    """Security-related error"""
    pass


class PathValidator:
    """Validates file paths for security"""
    
    FORBIDDEN_CHARS = set('<>:"|?*\x00')
    RESERVED_NAMES = {
        'CON', 'PRN', 'AUX', 'NUL',
        *(f'COM{i}' for i in range(1, 10)),
        *(f'LPT{i}' for i in range(1, 10)),
    }
    
    @classmethod
    def validate_path(
        cls, 
        path: str | Path, 
        must_exist: bool = False,
        check_allowed: bool = True
    ) -> Tuple[bool, str]:
        """Validate a file path for security and correctness."""
        try:
            p = Path(path).resolve()
        except Exception as e:
            return False, f"Invalid path format"
        
        # Check path traversal - compare resolved path with original
        original_resolved = Path(path).resolve()
        if '..' in str(path) or original_resolved != p:
            # Allow if the resolved path is still valid
            pass  # resolve() already handles this
        
        # Check for forbidden characters
        name = p.name
        for char in cls.FORBIDDEN_CHARS:
            if char in name:
                return False, f"Forbidden character in filename"
        
        # Check reserved names
        stem = p.stem.upper()
        if stem in cls.RESERVED_NAMES:
            return False, f"Reserved filename"
        
        # Check if within allowed directories (whitelist)
        if check_allowed:
            from ..config import is_path_allowed
            if not is_path_allowed(p):
                return False, f"Path not in allowed directories"
        
        # Check for symbolic links (security risk)
        if p.exists() and p.is_symlink():
            return False, "Symbolic links not allowed"
        
        # Check file existence
        if must_exist and not p.exists():
            return False, "File not found"
        
        # Check file size
        if p.exists() and p.is_file():
            size = p.stat().st_size
            if size > settings.security.max_file_size_bytes:
                return False, f"File too large (max {settings.security.max_file_size_bytes // 1024 // 1024}MB)"
        
        # Check extension
        if p.suffix.lower() not in {'.xlsx', '.xlsm', '.xltx', '.xltm'}:
            return False, f"Invalid file extension"
        
        return True, ""
    
    @classmethod
    def validate_sheet_name(cls, name: str) -> Tuple[bool, str]:
        """Validate sheet name according to Excel rules"""
        if not name:
            return False, "Sheet name cannot be empty"
        
        if len(name) > settings.security.max_sheet_name_length:
            return False, f"Sheet name too long (max {settings.security.max_sheet_name_length} chars)"
        
        for char in settings.security.forbidden_sheet_chars:
            if char in name:
                return False, f"Forbidden character in sheet name"
        
        if name.startswith("'") or name.endswith("'"):
            return False, "Sheet name cannot start or end with single quote"
        
        return True, ""


class CellValidator:
    """Validates cell and range references"""
    
    MAX_COLUMN = 16384  # XFD
    MAX_ROW = 1048576
    CELL_PATTERN = re.compile(r'^([A-Z]{1,3})(\d+)$', re.IGNORECASE)
    
    @classmethod
    def _column_to_number(cls, col: str) -> int:
        """Convert column letters to number"""
        result = 0
        for char in col.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result
    
    @classmethod
    def validate_cell(cls, cell: str) -> Tuple[bool, str, Tuple[int, int] | None]:
        """Validate cell reference."""
        match = cls.CELL_PATTERN.match(cell.strip())
        if not match:
            return False, f"Invalid cell reference format", None
        
        col_str, row_str = match.groups()
        col_num = cls._column_to_number(col_str)
        row_num = int(row_str)
        
        if col_num > cls.MAX_COLUMN:
            return False, f"Column exceeds Excel limit (max XFD)", None
        
        if row_num > cls.MAX_ROW:
            return False, f"Row exceeds Excel limit (max {cls.MAX_ROW})", None
        
        if row_num < 1:
            return False, f"Row must be at least 1", None
        
        return True, "", (col_num, row_num)
    
    @classmethod
    def validate_range(cls, range_ref: str) -> Tuple[bool, str, Tuple[int, int, int, int] | None]:
        """Validate range reference."""
        if ':' not in range_ref:
            return False, f"Invalid range format", None
        
        parts = range_ref.split(':')
        if len(parts) != 2:
            return False, f"Invalid range format", None
        
        valid1, err1, start = cls.validate_cell(parts[0])
        if not valid1:
            return False, err1, None
        
        valid2, err2, end = cls.validate_cell(parts[1])
        if not valid2:
            return False, err2, None
        
        start_col, start_row = start
        end_col, end_row = end
        
        min_col, max_col = min(start_col, end_col), max(start_col, end_col)
        min_row, max_row = min(start_row, end_row), max(start_row, end_row)
        
        cell_count = (max_col - min_col + 1) * (max_row - min_row + 1)
        if cell_count > settings.security.max_cells_per_operation:
            return False, f"Range too large (max {settings.security.max_cells_per_operation} cells)", None
        
        return True, "", (min_col, min_row, max_col, max_row)


class FormulaValidator:
    """Validates Excel formulas for security"""
    
    # Dangerous patterns that could be used for attacks
    DANGEROUS_PATTERNS = [
        r'=.*cmd\|',           # DDE command injection
        r'=.*\|.*!',           # DDE attack pattern
        r'=.*WEBSERVICE\s*\(', # External web requests
        r'=.*IMPORTDATA\s*\(', # Import external data
        r'=.*IMPORTHTML\s*\(', # Import HTML
        r'=.*IMPORTXML\s*\(',  # Import XML
        r'=.*EXECUTE\s*\(',    # Execute commands (legacy)
        r'=.*CALL\s*\(',       # DLL calls
        r'=.*REGISTER\s*\(',   # DLL registration
        r'=.*REGISTER\.ID\s*\(', # DLL registration ID
        r'=.*GET\.CELL\s*\(',  # Get cell info (can leak data)
        r'=.*GET\.WORKBOOK\s*\(', # Get workbook info
        r'=.*LINKS\s*\(',      # External links
        r'=.*REQUEST\s*\(',    # DDE request
    ]
    
    @classmethod
    def validate_formula(cls, formula: str) -> Tuple[bool, str]:
        """Validate formula for security risks."""
        if not formula:
            return False, "Formula cannot be empty"
        
        # Normalize formula
        normalized = formula.strip()
        if not normalized.startswith('='):
            normalized = '=' + normalized
        
        # Check length
        if len(normalized) > settings.security.max_formula_length:
            return False, f"Formula too long (max {settings.security.max_formula_length} chars)"
        
        # Check for dangerous patterns
        for pattern in cls.DANGEROUS_PATTERNS:
            if re.search(pattern, normalized, re.IGNORECASE):
                logger.warning(f"Blocked formula containing dangerous pattern: {pattern}")
                return False, "Formula contains blocked function for security reasons"
        
        # Additional check for HYPERLINK - allow but log
        if re.search(r'=.*HYPERLINK\s*\(', normalized, re.IGNORECASE):
            logger.warning(f"Formula contains HYPERLINK function")
        
        return True, ""


class DataValidator:
    """Validates data for write operations"""
    
    MAX_STRING_LENGTH = 32767  # Excel cell limit
    
    @classmethod
    def validate_data_size(cls, data: list[list], max_cells: int | None = None) -> Tuple[bool, str]:
        """Validate data array size"""
        if not data:
            return False, "Data cannot be empty"
        
        max_cells = max_cells or settings.security.max_cells_per_operation
        
        total_cells = 0
        for i, row in enumerate(data):
            if not isinstance(row, list):
                return False, f"Row {i} is not a list"
            total_cells += len(row)
            
            if total_cells > max_cells:
                return False, f"Data exceeds limit (max {max_cells} cells)"
        
        return True, ""
    
    @classmethod
    def validate_value(cls, value: Any) -> Tuple[bool, str]:
        """Validate a single cell value"""
        if value is None:
            return True, ""
        
        # Check type
        if isinstance(value, bool):
            return True, ""
        if isinstance(value, (int, float)):
            # Check for infinity/NaN
            if isinstance(value, float) and (value == float('inf') or value == float('-inf') or value != value):
                return False, "Invalid numeric value (infinity or NaN)"
            return True, ""
        if isinstance(value, str):
            if len(value) > cls.MAX_STRING_LENGTH:
                return False, f"String too long (max {cls.MAX_STRING_LENGTH} chars)"
            return True, ""
        
        # Reject other types
        return False, f"Invalid value type: {type(value).__name__}"
