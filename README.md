# Excel MCP Server

This project provides a set of tools for interacting with Excel workbooks and worksheets using the MCP (Message Control Protocol) server framework.

## Table of Contents

1. [What is MCP?](#what-is-mcp)
    - [MCP Server](#mcp-server)
    - [MCP Client](#mcp-client)
    - [MCP Tools](#mcp-tools)
2. [About MCP Excel Server](#about-mcp-excel-server)
3. [Available Tools](#available-tools)
    - [Workbook Tools](#workbook-tools)
    - [Worksheet Tools](#worksheet-tools)
    - [Range Tools](#range-tools)
4. [How to Use the Tools](#how-to-use-the-tools)
    - [Installation](#installation)
    - [Integration with AI Agents](#integration-with-ai-agents)
    - [Configuration for AI Agents](#configuration-for-ai-agents)
5. [MCP Inspector](#mcp-inspector)
6. [Contributing](#contributing)
7. [License](#license)

---

## What is MCP?

**MCP** stands for **Model Context Protocol**. It is an open protocol that standardizes how applications provide context to Large Language Models (LLMs). Think of MCP like a USB-C port for AI applications—it provides a standardized way to connect AI models to different data sources and tools.

MCP helps you build agents and complex workflows on top of LLMs by:
- Providing a growing list of pre-built integrations that your LLM can directly plug into
- Offering flexibility to switch between LLM providers and vendors
- Implementing best practices for securing your data within your infrastructure

### MCP Server
The **MCP Server** is a lightweight program that exposes specific capabilities through the standardized Model Context Protocol. It securely accesses local data sources (like your computer's files and databases) and remote services (available over the internet) to provide context to LLMs.

### MCP Client
The **MCP Client** is a protocol client that maintains a 1:1 connection with an MCP Server. It allows applications (like Claude Desktop, IDEs, or AI tools) to access data and tools provided by the server.

### MCP Tools
**MCP Tools** are specific capabilities exposed by MCP Servers. These tools enable LLMs to perform actions (e.g., creating a workbook, merging cells) and access data through a standardized interface.

---

## About MCP Excel Server

![MCP Excel Server](imgs/mcp_excel_server.jpg)

**MCP Excel Server** is an implementation of the MCP Server that provides tools for automating Excel file operations using Python and [openpyxl](https://openpyxl.readthedocs.io/). It allows you to:
- Create, open, and save Excel workbooks
- Add, rename, copy, move, and delete worksheets
- Read and write data to cells and ranges
- Merge, unmerge, copy, move, and delete cell ranges
- List available Excel files

---

## Available Tools

### Workbook Tools
- **[create_workbook](src/mcp_excel_server/api/tools/workbook.py)**: Create a new Excel workbook.
  - **Args:**
    - `filename`: Name of the Excel file to create.
    - `sheet_name`: Name of the initial worksheet. If not provided, defaults to "Sheet1".
  - **Returns:**
    - `success`: True if the workbook was created successfully.
    - `message`: A message describing the result.
    - `info`: Additional information about the created workbook.

- **[list_workbooks](src/mcp_excel_server/api/tools/workbook.py)**: List all Excel files in the configured directory.
  - **Args:**
    - None
  - **Returns:**
    - `success`: True if the operation succeeded.
    - `files`: List of Excel filenames.
    - `message`: Error message if success is False.

- **[read_workbook_data](src/mcp_excel_server/api/tools/workbook.py)**: Read data from an Excel worksheet.
  - **Args:**
    - `filepath`: Name of the Excel file.
    - `sheet_name`: Name of the worksheet to read from.
    - `start_cell`: Starting cell reference (e.g., "A1"). If not provided, defaults to "A1".
    - `end_cell`: Ending cell reference (e.g., "B10"). If None, reads to the end of data.
    - `preview_only`: If True, returns only a preview of the data.
  - **Returns:**
    - `success`: True if the data was read successfully.
    - `data`: Data from Excel worksheet as formatted string.
    - `message`: Additional message describing the result.

- **[write_workbook_data](src/mcp_excel_server/api/tools/workbook.py)**: Write data to an Excel worksheet.
  - **Args:**
    - `filepath`: Name of the Excel file.
    - `sheet_name`: Name of the worksheet to write to.
    - `data`: List of lists containing data to write (sublists are rows).
    - `start_cell`: Cell to start writing to. If None, appends after last row.
  - **Returns:**
    - `success`: True if the data was written successfully.
    - `message`: A message describing the result.

- **[get_workbook_info](src/mcp_excel_server/api/tools/workbook.py)**: Get information about an Excel workbook.
  - **Args:**
    - `filename`: Name of the Excel file.
  - **Returns:**
    - `success`: True if the workbook was found.
    - `message`: A message describing the result.
    - `info`: Additional information about the workbook.

### Worksheet Tools
- **[create_worksheet](src/mcp_excel_server/api/tools/worksheet.py)**: Creates a new worksheet in the specified Excel workbook.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `sheet_name`: The name for the new worksheet.
    - `index`: The position to insert the sheet (0-based). If not provided, appends to the end.
  - **Returns:**
    - `success`: True if the worksheet was created successfully.
    - `message`: A message describing the result.
    - `sheet_name`: The name of the created sheet.

- **[delete_worksheet](src/mcp_excel_server/api/tools/worksheet.py)**: Deletes a worksheet from the specified Excel workbook.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `sheet_name`: The name of the worksheet to delete.
  - **Returns:**
    - `success`: True if the worksheet was deleted successfully.
    - `message`: A message describing the result.

- **[rename_worksheet](src/mcp_excel_server/api/tools/worksheet.py)**: Renames a worksheet in the specified Excel workbook.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `old_name`: The current name of the worksheet.
    - `new_name`: The new name for the worksheet.
  - **Returns:**
    - `success`: True if the worksheet was renamed successfully.
    - `message`: A message describing the result.
    - `new_name`: The new name of the worksheet.

- **[copy_worksheet](src/mcp_excel_server/api/tools/worksheet.py)**: Creates a copy of a worksheet in the specified Excel workbook.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `sheet_name`: The name of the worksheet to copy.
    - `new_name`: The name for the new worksheet.
  - **Returns:**
    - `success`: True if the worksheet was copied successfully.
    - `message`: A message describing the result.
    - `new_name`: The name of the new worksheet.

- **[move_worksheet](src/mcp_excel_server/api/tools/worksheet.py)**: Moves a worksheet to a new position in the specified Excel workbook.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `sheet_name`: The name of the worksheet to move.
    - `index`: The new position for the worksheet (0-based).
  - **Returns:**
    - `success`: True if the worksheet was moved successfully.
    - `message`: A message describing the result.
    - `sheet_name`: The name of the moved worksheet.

- **[get_worksheet](src/mcp_excel_server/api/tools/worksheet.py)**: Retrieves information about a worksheet in the specified Excel workbook.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `sheet_name`: The name of the worksheet to retrieve information for.
  - **Returns:**
    - `success`: True if the worksheet was found.
    - `message`: A message describing the result.
    - `sheet`: Worksheet information (structure may vary).

- **[list_worksheets](src/mcp_excel_server/api/tools/worksheet.py)**: Lists all worksheets in the specified Excel workbook.
  - **Args:**
    - `filename`: The name of the Excel file.
  - **Returns:**
    - `success`: True if the operation succeeded.
    - `message`: A message describing the result.
    - `sheets`: List of worksheet names.

- **[merge_cells](src/mcp_excel_server/api/tools/worksheet.py)**: Merge a range of cells in a worksheet.
  - **Args:**
    - `filename`: Name of the Excel file.
    - `sheet_name`: Name of the worksheet.
    - `start_cell`: Top-left cell of range to merge.
    - `end_cell`: Bottom-right cell of range to merge.
  - **Returns:**
    - `success`: True if the cells were merged successfully.
    - `message`: A message describing the result.

- **[unmerge_cells](src/mcp_excel_server/api/tools/worksheet.py)**: Unmerge a range of cells in a worksheet.
  - **Args:**
    - `filename`: Name of the Excel file.
    - `sheet_name`: Name of the worksheet.
    - `start_cell`: Top-left cell of range to unmerge.
    - `end_cell`: Bottom-right cell of range to unmerge.
  - **Returns:**
    - `success`: True if the cells were unmerged successfully.
    - `message`: A message describing the result.

- **[filter_rows_by_column](src/mcp_excel_server/api/tools/worksheet.py)**: List all rows from a worksheet where a specified column matches a given value.
  - **Args:**
    - `filename`: Name of the Excel file.
    - `column_name`: The name of the column to filter on.
    - `filter_value`: The value to match in the column.
    - `sheet_name`: Name of the worksheet to read from. If not provided, defaults to the first sheet.
  - **Returns:**
    - `success`: True if the operation succeeded.
    - `data`: A formatted string of matching rows.
    - `message`: A message describing the result.

- **[filter_rows_by_columns](src/mcp_excel_server/api/tools/worksheet.py)**: List all rows from a worksheet where specified columns match given values.
  - **Args:**
    - `filename`: Name of the Excel file.
    - `column_names`: List of column names to filter on.
    - `filter_values`: List of values to match in the corresponding columns.
    - `sheet_name`: Name of the worksheet to read from. If not provided, defaults to the first sheet.
  - **Returns:**
    - `success`: True if the operation succeeded.
    - `data`: A formatted string of matching rows.
    - `message`: A message describing the result.

### Range Tools
- **[delete_range](src/mcp_excel_server/api/tools/range.py)**: Delete a range of cells.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `sheet_name`: The name of the worksheet.
    - `start_cell`: The top-left cell of the range (e.g., "A1").
    - `end_cell`: The bottom-right cell of the range (e.g., "B2").
    - `shift_direction`: The direction to shift cells after deletion ("up" or "left"). Defaults to "up".
  - **Returns:**
    - `success`: True if the range was deleted successfully.
    - `message`: A message describing the result.

- **[copy_range](src/mcp_excel_server/api/tools/range.py)**: Copy a range of cells.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `sheet_name`: The name of the worksheet.
    - `source_start`: The top-left cell of the source range (e.g., "A1").
    - `source_end`: The bottom-right cell of the source range (e.g., "B2").
    - `target_start`: The top-left cell of the target range (e.g., "C1").
  - **Returns:**
    - `success`: True if the range was copied successfully.
    - `message`: A message describing the result.

- **[move_range](src/mcp_excel_server/api/tools/range.py)**: Move a range of cells.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `sheet_name`: The name of the worksheet containing the range.
    - `source_range`: The cell range to move (e.g., 'A1:B10').
    - `target_range`: The cell range to move to (e.g., 'D1:E10').
  - **Returns:**
    - `success`: True if the range was moved successfully.
    - `message`: A message describing the result.

- **[merge_range](src/mcp_excel_server/api/tools/range.py)**: Merge a range of cells.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `sheet_name`: The name of the worksheet.
    - `start_cell`: The top-left cell of the range (e.g., "A1").
    - `end_cell`: The bottom-right cell of the range (e.g., "B2").
  - **Returns:**
    - `success`: True if the range was merged successfully.
    - `message`: A message describing the result.

- **[unmerge_range](src/mcp_excel_server/api/tools/range.py)**: Unmerge a range of cells.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `sheet_name`: The name of the worksheet.
    - `start_cell`: The top-left cell of the range (e.g., "A1").
    - `end_cell`: The bottom-right cell of the range (e.g., "B2").
  - **Returns:**
    - `success`: True if the range was unmerged successfully.
    - `message`: A message describing the result.

- **[validate_range](src/mcp_excel_server/api/tools/range.py)**: Validate a cell range.
  - **Args:**
    - `filename`: The name of the Excel file.
    - `sheet_name`: The name of the worksheet containing the range.
    - `range_str`: The cell range to validate (e.g., 'A1:B10').
  - **Returns:**
    - `success`: True if the range is valid.
    - `message`: A message describing the result.
    - `range_info`: Information about the range (start_cell, end_cell, num_rows, num_cols).

### Registering Tools

All MCP tools in this project are registered using the `@register_tool` decorator. This decorator is defined in [`src/mcp_excel_server/api/registry.py`](src/mcp_excel_server/api/registry.py) and serves the following purposes:

- **Automatic Registration:** When you annotate a function with `@register_tool`, it is automatically registered as an MCP tool with the server.
- **Name Prefixing:** The tool name is automatically prefixed with the value of `settings.tool_prefix` (by default, `"EXCEL_MCP_"`). For example, a function named `create_workbook` becomes a tool named `EXCEL_MCP_create_workbook`.
- **Integration with FastMCP:** The decorator uses the MCP server's `.tool()` method to make the function available for remote invocation by AI agents or other clients.

**Example Usage:**
```python
from mcp_excel_server.api.registry import register_tool

@register_tool
def create_worksheet(filename: str, sheet_name: str) -> dict:
    """Creates a new worksheet in the specified Excel workbook."""
    # ... implementation ...
```

**How it works:**
- The decorator constructs the tool name by concatenating the prefix and the function name.
- It registers the function with the MCP server so it can be called as a tool.
- The tool is then discoverable and invokable by any MCP-compatible client.

**Configuration:**
- You can change the prefix by setting the `EXCEL_MCP_SERVER_TOOL_PREFIX` environment variable or editing the `tool_prefix` in your settings.

---

## How to Use the Tools

### Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd mcp_excel_server
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Integration with AI Agents

MCP Excel Server is designed to be used with AI Agents like Claude, GitHub Copilot, and others. These agents can directly invoke the tools provided by the server without needing to send HTTP requests manually.

### Configuration for AI Agents

To integrate MCP Excel Server with AI Agents, follow these steps:

1. **Start the MCP Excel Server:**
   ```bash
   PYTHONPATH=src python src/mcp_excel_server/__main__.py sse
   ```

2. **Configure your AI Agent:**
   - For Claude Desktop, use the built-in MCP integration.
   - For GitHub Copilot, ensure your environment is set up to recognize the MCP Server's capabilities.

   **Configuration Examples:**

   **Stdio Transport Connection (for local integration):**
   ```json
   {
      "mcpServers": {
         "mcp-excel-stdio": {
            "command": "uvx",
            "args": ["mcp-excel-server", "stdio"]
         }
      }
   }
   ```

   **SSE Transport Connection:**
   ```json
   {
      "mcpServers": {
         "mcp-excel-server": {
            "url": "http://localhost:8800/sse"
         }
      }
   }
   ```

---

## MCP Inspector

**MCP Inspector** The MCP Inspector is an interactive developer tool for testing and debugging MCP servers. While the Debugging Guide covers the Inspector as part of the overall debugging toolkit, this document provides a detailed exploration of the Inspector's features and capabilities.

For more detailed information on using the MCP Inspector, refer to the [official documentation](https://modelcontextprotocol.io/docs/tools/inspector).

---

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details. 