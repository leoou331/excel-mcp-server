# Excel MCP Server

A robust Model Context Protocol (MCP) server for Excel file manipulation.

## Features

- **Safe File Operations**: File locking prevents concurrent access issues
- **Security Hardened**: Path validation, formula injection prevention, input sanitization
- **Resource Management**: Proper context managers ensure no file handle leaks
- **Type Safety**: Full Pydantic validation on all inputs

## Installation

```bash
pip install excel-mcp-server
```

## Usage

### As MCP Server

```bash
excel-mcp-server
```

### Configure allowed directories

```python
from excel_mcp_server import add_allowed_directory

add_allowed_directory("/path/to/excel/files")
```

## Tools

### Workbook Operations
- `create_workbook` - Create a new Excel workbook
- `get_workbook_info` - Get metadata about a workbook
- `list_sheets` - List all sheets in a workbook

### Sheet Operations
- `create_sheet` - Create a new sheet
- `delete_sheet` - Delete a sheet
- `rename_sheet` - Rename a sheet

### Cell Operations
- `read_cell` - Read a cell value
- `write_cell` - Write a cell value
- `read_range` - Read a range of cells
- `write_range` - Write data to a range
- `write_formula` - Write a formula

### Formatting Operations
- `format_font` - Apply font formatting
- `format_fill` - Apply background color
- `format_number` - Apply number format

## Security Features

1. **Path Validation**: Prevents path traversal attacks
2. **Directory Whitelist**: Restrict file access to allowed directories
3. **Formula Injection Prevention**: Blocks dangerous Excel formulas
4. **File Size Limits**: Prevents resource exhaustion
5. **File Locking**: Prevents concurrent write corruption

## License

MIT
