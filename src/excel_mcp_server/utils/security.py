"""Security utilities for Excel MCP Server"""

import logging
from pathlib import Path
from typing import Tuple, Any

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError

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
        raw_parts = [
            part for part in str(path).replace('\\', '/').split('/')
            if part and part != '.'
        ]
        if '..' in raw_parts:
            return False, "Path traversal is not allowed"

        raw_path = Path(path)

        try:
            p = raw_path.resolve()
        except Exception:
            return False, f"Invalid path format"
        
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
        normalized = cell.strip().upper()
        split_at = 0
        while split_at < len(normalized) and normalized[split_at].isalpha():
            split_at += 1

        col_str = normalized[:split_at]
        row_str = normalized[split_at:]

        if not col_str or not row_str:
            return False, f"Invalid cell reference format", None

        if len(col_str) > 3 or not col_str.isalpha() or not row_str.isdigit():
            return False, f"Invalid cell reference format", None

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
    
    BLOCKED_FUNCTIONS = {
        "WEBSERVICE",
        "IMPORTDATA",
        "IMPORTHTML",
        "IMPORTXML",
        "EXECUTE",
        "CALL",
        "REGISTER",
        "REGISTER.ID",
        "GET.CELL",
        "GET.WORKBOOK",
        "LINKS",
        "REQUEST",
    }
    WARN_ONLY_FUNCTIONS = {"HYPERLINK"}
    FUNCTION_PREFIXES = ("_XLFN.", "_XLWS.")

    @classmethod
    def _normalize_function_name(cls, token_value: str) -> str:
        """Normalize a function token for security checks."""
        normalized = token_value.rstrip('(').strip().upper()

        while normalized.startswith('@'):
            normalized = normalized[1:]

        changed = True
        while changed:
            changed = False
            for prefix in cls.FUNCTION_PREFIXES:
                if normalized.startswith(prefix):
                    normalized = normalized[len(prefix):]
                    changed = True

        return normalized

    @classmethod
    def _contains_unquoted_pipe(cls, formula: str) -> bool:
        """Detect DDE-style pipe syntax outside quoted strings."""
        quote_char: str | None = None
        index = 0

        while index < len(formula):
            char = formula[index]
            if quote_char:
                if char == quote_char:
                    if index + 1 < len(formula) and formula[index + 1] == quote_char:
                        index += 2
                        continue
                    quote_char = None
                index += 1
                continue

            if char in {'"', "'"}:
                quote_char = char
                index += 1
                continue

            if char == '|':
                return True

            index += 1

        return False
    
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

        if cls._contains_unquoted_pipe(normalized):
            logger.warning("Blocked formula containing DDE-style pipe syntax")
            return False, "Formula contains blocked syntax for security reasons"

        try:
            tokens = Tokenizer(normalized).items
        except TokenizerError:
            logger.warning("Blocked formula that failed tokenization")
            return False, "Formula contains unsupported or unsafe syntax"

        for token in tokens:
            if token.type != "FUNC" or token.subtype != "OPEN":
                continue

            function_name = cls._normalize_function_name(token.value)
            if function_name in cls.BLOCKED_FUNCTIONS:
                logger.warning(
                    "Blocked formula containing dangerous function: %s",
                    function_name,
                )
                return False, "Formula contains blocked function for security reasons"

            if function_name in cls.WARN_ONLY_FUNCTIONS:
                logger.warning("Formula contains HYPERLINK function")
        
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
