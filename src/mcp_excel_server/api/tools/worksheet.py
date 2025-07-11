"""
Worksheet management tools for Excel MCP Server.
This module provides tools for managing worksheets in Excel workbooks.
"""

from typing import Optional, List, Dict, Any
from pathlib import Path

from mcp_excel_server.api.registry import register_tool
from mcp_excel_server.core.workbook import get_workbook, save_workbook
from mcp_excel_server.core.sheet import (
    create_sheet,
    copy_sheet,
    delete_sheet,
    rename_sheet,
    move_sheet,
    get_sheet,
    list_sheets,
    merge_range,
    unmerge_range
)
from mcp_excel_server.core.exceptions import SheetError, ValidationError
from mcp_excel_server.config.settings import settings

from mcp_excel_server.utils import get_logger
# Initialize logger
logger = get_logger(__name__)


@register_tool
def create_worksheet(
    filename: str,
    sheet_name: str,
    index: Optional[int] = None
) -> Dict[str, Any]:
    """
    Creates a new worksheet in the specified Excel workbook.

    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name for the new worksheet.
        index (int, optional): The position to insert the sheet (0-based). If not provided, appends to the end.

    Returns:
        dict: {
            "success": bool,  # True if the worksheet was created successfully
            "message": str,   # A message describing the result
            "sheet_name": str # The name of the created sheet
        }

    Example:
        create_worksheet(filename="report.xlsx", sheet_name="Data", index=1)
    """
    logger.debug(f"create_worksheet called with filename={filename}, sheet_name={sheet_name}, index={index}")
    try:
        wb = get_workbook(filename)
        if sheet_name in wb.sheetnames:
            response = {
                "success": False,
                "message": f"Sheet '{sheet_name}' already exists",
                "sheet_name": sheet_name
            }
            logger.debug(f"create_worksheet response (sheet exists): {response}")
            return response
            
        wb.create_sheet(sheet_name)
        save_workbook(wb, filename)
        response = {
            "success": True,
            "message": f"Created worksheet '{sheet_name}'",
            "sheet_name": sheet_name
        }
        logger.debug(f"create_worksheet response (success): {response}")
        return response
    except Exception as e:
        response = {
            "success": False,
            "message": str(e),
            "sheet_name": sheet_name
        }
        logger.debug(f"create_worksheet response (error): {response}")
        return response

@register_tool
def delete_worksheet(
    filename: str,
    sheet_name: str
) -> Dict[str, Any]:
    """
    Deletes a worksheet from the specified Excel workbook.

    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet to delete.

    Returns:
        dict: {
            "success": bool,  # True if the worksheet was deleted successfully
            "message": str    # A message describing the result
        }

    Example:
        delete_worksheet(filename="report.xlsx", sheet_name="Data")
    """
    logger.debug(f"delete_worksheet called with filename={filename}, sheet_name={sheet_name}")
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            response = {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found"
            }
            logger.debug(f"delete_worksheet response (sheet not found): {response}")
            return response
            
        if len(wb.sheetnames) == 1:
            response = {
                "success": False,
                "message": "Cannot delete the only sheet in workbook"
            }
            logger.debug(f"delete_worksheet response (last sheet): {response}")
            return response
            
        del wb[sheet_name]
        save_workbook(wb, filename)
        response = {
            "success": True,
            "message": f"Deleted worksheet '{sheet_name}'"
        }
        logger.debug(f"delete_worksheet response (success): {response}")
        return response
    except Exception as e:
        response = {
            "success": False,
            "message": str(e)
        }
        logger.debug(f"delete_worksheet response (error): {response}")
        return response

@register_tool
def rename_worksheet(
    filename: str,
    old_name: str,
    new_name: str
) -> Dict[str, Any]:
    """
    Renames a worksheet in the specified Excel workbook.

    Args:
        filename (str): The name of the Excel file.
        old_name (str): The current name of the worksheet.
        new_name (str): The new name for the worksheet.

    Returns:
        dict: {
            "success": bool,  # True if the worksheet was renamed successfully
            "message": str,   # A message describing the result
            "new_name": str   # The new name of the worksheet
        }

    Example:
        rename_worksheet(filename="report.xlsx", old_name="Sheet1", new_name="Data")
    """
    logger.debug(f"rename_worksheet called with filename={filename}, old_name={old_name}, new_name={new_name}")
    try:
        wb = get_workbook(filename)
        if old_name not in wb.sheetnames:
            response = {
                "success": False,
                "message": f"Sheet '{old_name}' not found",
                "new_name": new_name
            }
            logger.debug(f"rename_worksheet response (sheet not found): {response}")
            return response
            
        if new_name in wb.sheetnames:
            response = {
                "success": False,
                "message": f"Sheet '{new_name}' already exists",
                "new_name": new_name
            }
            logger.debug(f"rename_worksheet response (new name exists): {response}")
            return response
            
        rename_sheet(filename, old_name, new_name)
        save_workbook(wb, filename)
        response = {
            "success": True,
            "message": f"Renamed worksheet '{old_name}' to '{new_name}'",
            "new_name": new_name
        }
        logger.debug(f"rename_worksheet response (success): {response}")
        return response
    except Exception as e:
        response = {
            "success": False,
            "message": str(e),
            "new_name": new_name
        }
        logger.debug(f"rename_worksheet response (error): {response}")
        return response

@register_tool
def copy_worksheet(
    filename: str,
    sheet_name: str,
    new_name: str
) -> Dict[str, Any]:
    """
    Creates a copy of a worksheet in the specified Excel workbook.

    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet to copy.
        new_name (str): The name for the new worksheet.

    Returns:
        dict: {
            "success": bool,  # True if the worksheet was copied successfully
            "message": str,   # A message describing the result
            "new_name": str   # The name of the new worksheet
        }

    Example:
        copy_worksheet(filename="report.xlsx", sheet_name="Data", new_name="Data_Copy")
    """
    logger.debug(f"copy_worksheet called with filename={filename}, sheet_name={sheet_name}, new_name={new_name}")
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            response = {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found",
                "new_name": new_name
            }
            logger.debug(f"copy_worksheet response (sheet not found): {response}")
            return response
            
        if new_name in wb.sheetnames:
            response = {
                "success": False,
                "message": f"Sheet '{new_name}' already exists",
                "new_name": new_name
            }
            logger.debug(f"copy_worksheet response (new name exists): {response}")
            return response
            
        copy_sheet(filename, sheet_name, new_name)
        save_workbook(wb, filename)
        response = {
            "success": True,
            "message": f"Copied worksheet '{sheet_name}' to '{new_name}'",
            "new_name": new_name
        }
        logger.debug(f"copy_worksheet response (success): {response}")
        return response
    except Exception as e:
        response = {
            "success": False,
            "message": str(e),
            "new_name": new_name
        }
        logger.debug(f"copy_worksheet response (error): {response}")
        return response

@register_tool
def move_worksheet(
    filename: str,
    sheet_name: str,
    index: int
) -> Dict[str, Any]:
    """
    Moves a worksheet to a new position in the specified Excel workbook.

    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet to move.
        index (int): The new position for the worksheet (0-based).

    Returns:
        dict: {
            "success": bool,  # True if the worksheet was moved successfully
            "message": str,   # A message describing the result
            "sheet_name": str # The name of the moved worksheet
        }

    Example:
        move_worksheet(filename="report.xlsx", sheet_name="Data", index=0)
    """
    logger.debug(f"move_worksheet called with filename={filename}, sheet_name={sheet_name}, index={index}")
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            response = {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found",
                "sheet_name": sheet_name
            }
            logger.debug(f"move_worksheet response (sheet not found): {response}")
            return response
            
        if index < 0 or index >= len(wb.sheetnames):
            response = {
                "success": False,
                "message": f"Invalid index {index}. Must be between 0 and {len(wb.sheetnames)-1}",
                "sheet_name": sheet_name
            }
            logger.debug(f"move_worksheet response (invalid index): {response}")
            return response
            
        sheet = wb[sheet_name]
        wb.move_sheet(sheet_name, offset=index - wb.index(sheet))
        save_workbook(wb, filename)
        response = {
            "success": True,
            "message": f"Moved worksheet '{sheet_name}' to position {index}",
            "sheet_name": sheet_name
        }
        logger.debug(f"move_worksheet response (success): {response}")
        return response
    except Exception as e:
        response = {
            "success": False,
            "message": str(e),
            "sheet_name": sheet_name
        }
        logger.debug(f"move_worksheet response (error): {response}")
        return response

@register_tool
def get_worksheet(
    filename: str,
    sheet_name: str
) -> Dict[str, Any]:
    """
    Retrieves information about a worksheet in the specified Excel workbook.

    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet to retrieve information for.

    Returns:
        dict: {
            "success": bool,  # True if the worksheet was found
            "message": str,   # A message describing the result
            "sheet": dict     # Worksheet information (structure may vary)
        }

    Example:
        get_worksheet(filename="report.xlsx", sheet_name="Data")
    """
    logger.debug(f"get_worksheet called with filename={filename}, sheet_name={sheet_name}")
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            response = {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found",
                "sheet": None
            }
            logger.debug(f"get_worksheet response (sheet not found): {response}")
            return response
            
        sheet = get_sheet(filename, sheet_name)
        response = {
            "success": True,
            "message": f"Retrieved worksheet '{sheet_name}'",
            "sheet": {"title": sheet["title"]}
        }
        logger.debug(f"get_worksheet response (success): {response}")
        return response
    except Exception as e:
        response = {
            "success": False,
            "message": str(e),
            "sheet": None
        }
        logger.debug(f"get_worksheet response (error): {response}")
        return response

@register_tool
def list_worksheets(
    filename: str
) -> Dict[str, Any]:
    """
    Lists all worksheets in the specified Excel workbook.

    Args:
        filename (str): The name of the Excel file.

    Returns:
        dict: {
            "success": bool,  # True if the operation succeeded
            "message": str,   # A message describing the result
            "sheets": list[str] # List of worksheet names
        }

    Example:
        list_worksheets(filename="report.xlsx")
    """
    logger.debug(f"list_worksheets called with filename={filename}")
    try:
        wb = get_workbook(filename)
        sheets = list_sheets(filename)
        response = {
            "success": True,
            "message": f"Listed worksheets in '{filename}'",
            "sheets": sheets
        }
        logger.debug(f"list_worksheets response (success): {response}")
        return response
    except Exception as e:
        response = {
            "success": False,
            "message": str(e),
            "sheets": []
        }
        logger.debug(f"list_worksheets response (error): {response}")
        return response

@register_tool
def merge_cells(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str
) -> Dict[str, Any]:
    """Merge a range of cells in a worksheet.

    Args:
        filename: Name of the Excel file
        sheet_name: Name of the worksheet
        start_cell: Top-left cell of range to merge
        end_cell: Bottom-right cell of range to merge

    Returns:
        Dict containing:
            success (bool): True if the cells were merged successfully
            message (str): A message describing the result
    """
    logger.debug(f"merge_cells called with filename={filename}, sheet_name={sheet_name}, start_cell={start_cell}, end_cell={end_cell}")
    try:
        result = merge_range(filename, sheet_name, start_cell, end_cell)
        response = {
            "success": True,
            "message": result["message"]
        }
        logger.debug(f"merge_cells response (success): {response}")
        return response
    except (ValidationError, SheetError) as e:
        response = {
            "success": False,
            "message": f"Error: {str(e)}"
        }
        logger.debug(f"merge_cells response (validation error): {response}")
        return response
    except Exception as e:
        logger.error(f"Error merging cells: {e}")
        response = {
            "success": False,
            "message": str(e)
        }
        logger.debug(f"merge_cells response (error): {response}")
        return response

@register_tool
def unmerge_cells(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str
) -> Dict[str, Any]:
    """Unmerge a range of cells in a worksheet.

    Args:
        filename: Name of the Excel file
        sheet_name: Name of the worksheet
        start_cell: Top-left cell of range to unmerge
        end_cell: Bottom-right cell of range to unmerge

    Returns:
        Dict containing:
            success (bool): True if the cells were unmerged successfully
            message (str): A message describing the result
    """
    logger.debug(f"unmerge_cells called with filename={filename}, sheet_name={sheet_name}, start_cell={start_cell}, end_cell={end_cell}")
    try:
        result = unmerge_range(filename, sheet_name, start_cell, end_cell)
        response = {
            "success": True,
            "message": result["message"]
        }
        logger.debug(f"unmerge_cells response (success): {response}")
        return response
    except (ValidationError, SheetError) as e:
        response = {
            "success": False,
            "message": f"Error: {str(e)}"
        }
        logger.debug(f"unmerge_cells response (validation error): {response}")
        return response
    except Exception as e:
        logger.error(f"Error unmerging cells: {e}")
        response = {
            "success": False,
            "message": str(e)
        }
        logger.debug(f"unmerge_cells response (error): {response}")
        return response

@register_tool
def filter_rows_by_column(
    filename: str,
    column_name: str,
    filter_value: str,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """List all rows from a worksheet where a specified column matches a given value.

    Args:
        filename: Name of the Excel file.
        column_name: The name of the column to filter on.
        filter_value: The value to match in the column.
        sheet_name: Name of the worksheet to read from. If not provided, defaults to the first sheet.

    Returns:
        Dict containing:
            success (bool): True if the operation succeeded.
            data (str): A formatted string of matching rows.
            message (str): A message describing the result.
    """
    logger.debug(f"filter_rows_by_column called with filename={filename}, sheet_name={sheet_name}, column_name={column_name}, filter_value={filter_value}")
    try:
        wb = get_workbook(filename)
        if sheet_name is None:
            sheet_name = wb.sheetnames[0]  # Default to the first sheet
        ws = wb[sheet_name]
        # Find the column index for the given column name
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        try:
            col_idx = header_row.index(column_name) + 1  # 1-based index
        except ValueError:
            wb.close()
            return {
                "success": False,
                "data": "",
                "message": f"Column '{column_name}' not found in the worksheet."
            }
        # Collect rows where the specified column value matches
        matching_rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[col_idx - 1] == filter_value:
                matching_rows.append(row)
        wb.close()
        if not matching_rows:
            return {
                "success": True,
                "data": "No matching rows found.",
                "message": "No matching rows found."
            }
        # Convert matching rows to a formatted string
        data_str = "\n".join([str(row) for row in matching_rows])
        return {
            "success": True,
            "data": data_str,
            "message": f"Found {len(matching_rows)} matching rows."
        }
    except Exception as e:
        logger.error(f"Error filtering rows: {e}")
        return {
            "success": False,
            "data": "",
            "message": str(e)
        }

@register_tool
def filter_rows_by_columns(
    filename: str,
    column_names: List[str],
    filter_values: List[str],
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """List all rows from a worksheet where specified columns match given values.

    Args:
        filename: Name of the Excel file.
        column_names: List of column names to filter on.
        filter_values: List of values to match in the corresponding columns.
        sheet_name: Name of the worksheet to read from. If not provided, defaults to the first sheet.

    Returns:
        Dict containing:
            success (bool): True if the operation succeeded.
            data (str): A formatted string of matching rows.
            message (str): A message describing the result.
    """
    logger.debug(f"filter_rows_by_columns called with filename={filename}, sheet_name={sheet_name}, column_names={column_names}, filter_values={filter_values}")
    try:
        wb = get_workbook(filename)
        if sheet_name is None:
            sheet_name = wb.sheetnames[0]  # Default to the first sheet
        ws = wb[sheet_name]
        # Find the column indices for the given column names
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        col_indices = []
        for col_name in column_names:
            try:
                col_idx = header_row.index(col_name) + 1  # 1-based index
                col_indices.append(col_idx)
            except ValueError:
                wb.close()
                return {
                    "success": False,
                    "data": "",
                    "message": f"Column '{col_name}' not found in the worksheet."
                }
        # Collect rows where all specified column values match
        matching_rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if all(row[col_idx - 1] == filter_values[i] for i, col_idx in enumerate(col_indices)):
                matching_rows.append(row)
        wb.close()
        if not matching_rows:
            return {
                "success": True,
                "data": "No matching rows found.",
                "message": "No matching rows found."
            }
        # Convert matching rows to a formatted string
        data_str = "\n".join([str(row) for row in matching_rows])
        return {
            "success": True,
            "data": data_str,
            "message": f"Found {len(matching_rows)} matching rows."
        }
    except Exception as e:
        logger.error(f"Error filtering rows: {e}")
        return {
            "success": False,
            "data": "",
            "message": str(e)
        }


