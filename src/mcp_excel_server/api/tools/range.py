"""
Range management tools for Excel MCP Server.
This module provides tools for managing ranges in Excel workbooks.
"""

from typing import Optional, List, Dict, Any, Union

from mcp_excel_server.utils import get_logger
# Initialize logger
logger = get_logger(__name__)

from mcp_excel_server.api.registry import register_tool
from mcp_excel_server.core.workbook import get_workbook, save_workbook
from mcp_excel_server.core.sheet import (
    merge_range,
    unmerge_range,
    copy_range_operation,
    delete_range_operation
)
from mcp_excel_server.core.exceptions import RangeError

@register_tool
def delete_range_tool(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up"
) -> Dict[str, Any]:
    """
    Deletes a range of cells in an Excel worksheet.

    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet.
        start_cell (str): The top-left cell of the range (e.g., "A1").
        end_cell (str): The bottom-right cell of the range (e.g., "B2").
        shift_direction (str, optional): The direction to shift cells after deletion ("up" or "left"). Defaults to "up".

    Returns:
        dict: {
            "success": bool,  # True if the range was deleted successfully
            "message": str    # A message describing the result
        }

    Example:
        delete_range_tool(filename="report.xlsx", sheet_name="Sheet1", start_cell="A1", end_cell="B2")
    """
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found"
            }
            
        delete_range_operation(filename, sheet_name, start_cell, end_cell, shift_direction)
        save_workbook(wb, filename)
        return {
            "success": True,
            "message": f"Deleted range {start_cell}:{end_cell}"
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e)
        }

@register_tool
def copy_range_tool(
    filename: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str
) -> Dict[str, Any]:
    """
    Copies a range of cells from one location to another in an Excel worksheet.

    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet.
        source_start (str): The top-left cell of the source range (e.g., "A1").
        source_end (str): The bottom-right cell of the source range (e.g., "B2").
        target_start (str): The top-left cell of the target range (e.g., "C1").

    Returns:
        dict: {
            "success": bool,  # True if the range was copied successfully
            "message": str    # A message describing the result
        }

    Example:
        copy_range_tool(filename="report.xlsx", sheet_name="Sheet1", source_start="A1", source_end="B2", target_start="C1")
    """
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found"
            }
            
        copy_range_operation(filename, sheet_name, source_start, source_end, target_start)
        save_workbook(wb, filename)
        return {
            "success": True,
            "message": f"Copied range {source_start}:{source_end} to {target_start}"
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e)
        }

@register_tool
def move_range_tool(
    filename: str,
    sheet_name: str,
    source_range: str,
    target_range: str
) -> Dict[str, Any]:
    """
    Moves a range of cells to a new location in the specified Excel worksheet.

    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet containing the range.
        source_range (str): The cell range to move (e.g., 'A1:B10').
        target_range (str): The cell range to move to (e.g., 'D1:E10').

    Returns:
        dict: {
            "success": bool,  # True if the range was moved successfully
            "message": str    # A message describing the result
        }

    Example:
        move_range_tool(filename="report.xlsx", sheet_name="Data", source_range="A1:B10", target_range="D1:E10")
    """
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found"
            }
            
        # Parse source range
        source_start, source_end = source_range.split(':')
        target_start = target_range.split(':')[0]
        
        # Copy to new location
        copy_range_operation(filename, sheet_name, source_start, source_end, target_start)
        
        # Delete from old location
        delete_range_operation(filename, sheet_name, source_start, source_end)
        
        save_workbook(wb, filename)
        return {
            "success": True,
            "message": f"Moved range '{source_range}' to '{target_range}' in '{sheet_name}'"
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e)
        }

@register_tool
def merge_range_tool(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str
) -> Dict[str, Any]:
    """
    Merges a range of cells in an Excel worksheet.

    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet.
        start_cell (str): The top-left cell of the range (e.g., "A1").
        end_cell (str): The bottom-right cell of the range (e.g., "B2").

    Returns:
        dict: {
            "success": bool,  # True if the range was merged successfully
            "message": str    # A message describing the result
        }

    Example:
        merge_range_tool(filename="report.xlsx", sheet_name="Sheet1", start_cell="A1", end_cell="B2")
    """
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found"
            }
            
        merge_range(filename, sheet_name, start_cell, end_cell)
        save_workbook(wb, filename)
        return {
            "success": True,
            "message": f"Merged range {start_cell}:{end_cell}"
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e)
        }

@register_tool
def unmerge_range_tool(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str
) -> Dict[str, Any]:
    """
    Unmerges a range of cells in an Excel worksheet.

    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet.
        start_cell (str): The top-left cell of the range (e.g., "A1").
        end_cell (str): The bottom-right cell of the range (e.g., "B2").

    Returns:
        dict: {
            "success": bool,  # True if the range was unmerged successfully
            "message": str    # A message describing the result
        }

    Example:
        unmerge_range_tool(filename="report.xlsx", sheet_name="Sheet1", start_cell="A1", end_cell="B2")
    """
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found"
            }
            
        unmerge_range(filename, sheet_name, start_cell, end_cell)
        save_workbook(wb, filename)
        return {
            "success": True,
            "message": f"Unmerged range {start_cell}:{end_cell}"
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e)
        }
    
@register_tool
def validate_range_tool(
    filename: str,
    sheet_name: str,
    range_str: str
) -> Dict[str, Any]:
    """
    Validates a range in the specified Excel worksheet and returns details about the range.

    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet containing the range.
        range_str (str): The cell range to validate (e.g., 'A1:B10').

    Returns:
        dict: {
            "success": bool,  # True if the range is valid
            "message": str,   # A message describing the result
            "range_info": dict # Information about the range (start_cell, end_cell, num_rows, num_cols)
        }

    Example:
        validate_range_tool(filename="report.xlsx", sheet_name="Data", range_str="A1:B10")
    """
    try:
        range_info = validate_range(filename, sheet_name, range_str)
        return {
            "success": True,
            "message": f"Validated range '{range_str}' in '{sheet_name}'",
            "range_info": range_info
        }
    except RangeError as e:
        return {
            "success": False,
            "message": str(e)
        }