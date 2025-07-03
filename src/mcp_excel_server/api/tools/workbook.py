"""
Workbook management tools for Excel MCP Server.
This module provides tools for managing Excel workbooks, including creating, reading, writing,
and manipulating workbook data and properties.
"""

from typing import Optional, List, Dict, Any
import os
from pathlib import Path

from mcp_excel_server.api.registry import register_tool
from mcp_excel_server.core.workbook import (
    create_workbook as create_workbook_impl,
    save_workbook,
    get_workbook_info as get_workbook_info_impl,
    list_workbooks as list_workbooks_impl,
    get_workbook
)
from mcp_excel_server.core.exceptions import WorkbookError, ValidationError, DataError
from mcp_excel_server.config.settings import settings
from mcp_excel_server.core.data import read_excel_range, write_data, get_next_available_cell

from mcp_excel_server.utils import get_logger
# Initialize logger
logger = get_logger(__name__)


@register_tool
def create_workbook(filename: str, sheet_name: str = "Sheet1") -> Dict[str, Any]:
    """Create a new Excel workbook.

    Args:
        filename: Name of the Excel file to create
        sheet_name: Name of the initial worksheet

    Returns:
        Dict containing:
            success (bool): True if the workbook was created successfully
            message (str): A message describing the result
            info (dict): Additional information about the created workbook
    """
    logger.debug(f"create_workbook called with filename={filename}, sheet_name={sheet_name}")
    result = create_workbook_impl(filename, sheet_name)
    logger.debug(f"create_workbook result: {result}")
    return {
        "success": result["success"],
        "message": result["message"],
        "info": {
            "filename": result.get("filename"),
            "sheet_name": sheet_name
        } if result["success"] else {}
    }

    
@register_tool
def list_workbooks() -> Dict[str, Any]:
    """List all Excel files in the configured directory.

    The directory will be created if it doesn't exist. Only .xlsx and .xlsm files are included.

    Returns:
        Dict containing:
            success (bool): True if the operation succeeded
            files (list[str]): List of Excel filenames
            message (str): Error message if success is False
    """
    logger.debug("list_workbooks called")
    result = list_workbooks_impl()
    logger.debug(f"list_workbooks result: {result}")
    return result

@register_tool
def read_workbook_data(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    preview_only: bool = False
) -> Dict[str, Any]:
    """Read data from an Excel worksheet.

    Args:
        filepath: Name of the Excel file
        sheet_name: Name of the worksheet to read from
        start_cell: Starting cell reference (e.g., "A1")
        end_cell: Ending cell reference (e.g., "B10"). If None, reads to the end of data
        preview_only: If True, returns only a preview of the data

    Returns:
        Dict containing:
            success (bool): True if the data was read successfully
            data (str): Data from Excel worksheet as formatted string
            message (str): Additional message describing the result
    """
    try:
        logger.debug(f"read_workbook_data called with filepath={filepath}, sheet_name={sheet_name}, start_cell={start_cell}, end_cell={end_cell}, preview_only={preview_only}")
        # Don't use get_full_path here since read_excel_range already handles the path
        logger.debug(f"Using filepath: {filepath}")
        result = read_excel_range(filepath, sheet_name, start_cell, end_cell, preview_only)
        if not result:
            logger.debug("No data found in specified range")
            return {
                "success": True,
                "data": "No data found in specified range",
                "message": "No data found in specified range"
            }
        # Convert the list of dicts to a formatted string
        data_str = "\n".join([str(row) for row in result])
        logger.debug(f"Read data result: {data_str}")
        response = {
            "success": True,
            "data": data_str,
            "message": "Data read successfully"
        }
        logger.debug(f"read_workbook_data response: {response}")
        return response
    except Exception as e:
        logger.error(f"Error reading data: {e}")
        response = {
            "success": False,
            "data": "",
            "message": str(e)
        }
        logger.debug(f"read_workbook_data error response: {response}")
        return response

@register_tool
def write_workbook_data(
    filepath: str,
    sheet_name: str,
    data: List[List],
    start_cell: Optional[str] = None
) -> Dict[str, Any]:
    """Write data to an Excel worksheet.

    Args:
        filepath: Name of the Excel file
        sheet_name: Name of the worksheet to write to
        data: List of lists containing data to write (sublists are rows)
        start_cell: Cell to start writing to. If None, appends after last row

    Returns:
        Dict containing:
            success (bool): True if the data was written successfully
            message (str): A message describing the result
    """
    try:
        logger.debug(f"write_workbook_data called with filepath={filepath}, sheet_name={sheet_name}, data={data}, start_cell={start_cell}")
        # Don't use get_full_path here since write_data already handles the path
        if start_cell is None:
            # Find the next available cell in the sheet
            logger.debug("Finding next available cell")
            wb = get_workbook(filepath)
            ws = wb[sheet_name]
            start_cell = get_next_available_cell(ws)
            wb.close()
            logger.debug(f"Next available cell: {start_cell}")
        result = write_data(filepath, sheet_name, data, start_cell)
        response = {
            "success": True,
            "message": result["message"]
        }
        logger.debug(f"write_workbook_data response: {response}")
        return response
    except (ValidationError, DataError) as e:
        response = {
            "success": False,
            "message": f"Error: {str(e)}"
        }
        logger.debug(f"write_workbook_data validation error response: {response}")
        return response
    except Exception as e:
        logger.error(f"Error writing data: {e}")
        response = {
            "success": False,
            "message": str(e)
        }
        logger.debug(f"write_workbook_data error response: {response}")
        return response

@register_tool
def get_workbook_info(filename: str) -> Dict[str, Any]:
    """Get information about an Excel workbook.

    Args:
        filename: Name of the Excel file

    Returns:
        Dict containing:
            success (bool): True if the workbook was found
            message (str): A message describing the result
            info (dict): Additional information about the workbook
    """
    logger.debug(f"get_workbook_info called with filename={filename}")
    result = get_workbook_info_impl(filename)
    logger.debug(f"get_workbook_info result: {result}")
    return result