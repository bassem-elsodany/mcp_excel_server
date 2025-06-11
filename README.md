# Excel MCP Server

This project provides a set of tools for interacting with Excel workbooks and worksheets using the MCP (Message Control Protocol) server framework.

## Overview

The Excel MCP Server allows you to perform various operations on Excel files, such as creating workbooks, managing worksheets, reading and writing data, and more. It is designed to be used as a part of a larger system that requires Excel file manipulation.

## Features

- **Workbook Management**: Create, list, read, and write data to Excel workbooks.
- **Worksheet Management**: Create, delete, rename, copy, and move worksheets within a workbook.
- **Data Operations**: Read and write data to specific ranges in worksheets.

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd mcp_excel_server
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Workbook Tools

- **Create Workbook**: Creates a new Excel workbook.
  ```python
  from mcp_excel_server.api.tools.workbook import create_workbook
  result = create_workbook("my_workbook.xlsx")
  ```

- **List Workbooks**: Lists all available workbooks.
  ```python
  from mcp_excel_server.api.tools.workbook import list_workbooks
  result = list_workbooks()
  ```

- **Read Workbook Data**: Reads data from a specific range in a workbook.
  ```python
  from mcp_excel_server.api.tools.workbook import read_workbook_data
  result = read_workbook_data("my_workbook.xlsx", "Sheet1", "A1", "B2")
  ```

- **Write Workbook Data**: Writes data to a specific range in a workbook.
  ```python
  from mcp_excel_server.api.tools.workbook import write_workbook_data
  data = [["New1", "New2"], ["New3", "New4"]]
  result = write_workbook_data("my_workbook.xlsx", "Sheet1", data, "C1")
  ```

### Worksheet Tools

- **Create Worksheet**: Creates a new worksheet in an existing workbook.
  ```python
  from mcp_excel_server.api.tools.worksheet import create_worksheet
  result = create_worksheet("my_workbook.xlsx", "NewSheet")
  ```

- **Delete Worksheet**: Deletes a worksheet from a workbook.
  ```python
  from mcp_excel_server.api.tools.worksheet import delete_worksheet
  result = delete_worksheet("my_workbook.xlsx", "Sheet1")
  ```

- **Rename Worksheet**: Renames an existing worksheet.
  ```python
  from mcp_excel_server.api.tools.worksheet import rename_worksheet
  result = rename_worksheet("my_workbook.xlsx", "Sheet1", "NewName")
  ```

- **Copy Worksheet**: Copies an existing worksheet to a new one.
  ```python
  from mcp_excel_server.api.tools.worksheet import copy_worksheet
  result = copy_worksheet("my_workbook.xlsx", "Sheet1", "CopySheet")
  ```

- **Move Worksheet**: Moves a worksheet to a new position.
  ```python
  from mcp_excel_server.api.tools.worksheet import move_worksheet
  result = move_worksheet("my_workbook.xlsx", "Sheet1", 1)
  ```

## Configuration

The server uses a configuration file to set various parameters, such as the directory where Excel files are stored. The default configuration is located in `src/mcp_excel_server/config/settings.py`.

## Testing

To run the tests, use the following command:
```bash
python -m pytest
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details. 