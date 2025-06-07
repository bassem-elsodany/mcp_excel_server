"""
Data and cell management tools for Excel MCP server.
This module provides functions for reading/writing data and managing cell operations
like merging/unmerging cells in Excel worksheets.
"""

import os
from typing import Dict, List, Optional, Tuple, Union, Any

from mcp_excel_server.api.server import register_tool
from mcp_excel_server.core.workbook import create_sheet as create_worksheet_impl, get_workbook
from mcp_excel_server.utils import get_logger
from mcp_excel_server.utils.helpers import get_excel_path
from mcp_excel_server.config.settings import settings

from mcp_excel_server.core.exceptions import (
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
)
from mcp_excel_server.core.data import read_excel_range, write_data, get_next_available_cell

from mcp_excel_server.core.validation import (
    validate_formula_in_cell_operation as validate_formula_impl,
    validate_range_in_sheet_operation as validate_range_impl
)
from mcp_excel_server.core.workbook import get_workbook_info
from mcp_excel_server.core.sheet import (
    copy_sheet,
    delete_sheet,
    rename_sheet,
    merge_range,
    unmerge_range,
    copy_range_operation,
    delete_range_operation
)


logger = get_logger(__name__)

@register_tool
def read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    preview_only: bool = False
) -> Dict[str, Any]:
    """
    Read data from Excel worksheet.
    
    Returns:  
    dict: {
        "success": bool,  # True if the data was read successfully
        "data": str,      # Data from Excel worksheet as json string, or error message if failed
        "message": str    # Additional message describing the result
    }
    """
    try:
        logger.debug(f"read_data_from_excel called with filepath={filepath}, sheet_name={sheet_name}, start_cell={start_cell}, end_cell={end_cell}, preview_only={preview_only}")
        full_path = get_excel_path(filepath)
        logger.debug(f"Resolved full_path: {full_path}")
        result = read_excel_range(full_path, sheet_name, start_cell, end_cell, preview_only)
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
        return {
            "success": True,
            "data": data_str,
            "message": "Data read successfully"
        }
    except Exception as e:
        logger.error(f"Error reading data: {e}")
        return {
            "success": False,
            "data": "",
            "message": str(e)
        }

@register_tool
def write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[List],
    start_cell: Optional[str] = None,  # Allow None to mean append
) -> Dict[str, Any]:
    """
    Write data to Excel worksheet.
    Excel formula will write to cell without any verification.

    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet to write to
        data: List of lists containing data to write to the worksheet, sublists are assumed to be rows
        start_cell: Cell to start writing to, default is None (append after last row)
    
    Returns:
        dict: {
            "success": bool,  # True if the data was written successfully
            "message": str    # A message describing the result
        }
    """
    try:
        full_path = get_excel_path(filepath)
        if start_cell is None:
            # Find the next available cell in the sheet
            wb = get_workbook(full_path)
            ws = wb[sheet_name]
            start_cell = get_next_available_cell(ws)
            wb.close()
        result = write_data(full_path, sheet_name, data, start_cell)
        return {
            "success": True,
            "message": result["message"]
        }
    except (ValidationError, DataError) as e:
        return {
            "success": False,
            "message": f"Error: {str(e)}"
        }
    except Exception as e:
        logger.error(f"Error writing data: {e}")
        return {
            "success": False,
            "message": str(e)
        }

@register_tool
def merge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """
    Merge a range of cells.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Top-left cell of range to merge
        end_cell: Bottom-right cell of range to merge
        
    Returns:
        dict: {
            "success": bool,  # True if the cells were merged successfully
            "message": str    # A message describing the result
        }
    """
    try:
        full_path = get_excel_path(filepath)
        result = merge_range(full_path, sheet_name, start_cell, end_cell)
        return {
            "success": True,
            "message": result["message"]
        }
    except (ValidationError, SheetError) as e:
        return {
            "success": False,
            "message": f"Error: {str(e)}"
        }
    except Exception as e:
        logger.error(f"Error merging cells: {e}")
        return {
            "success": False,
            "message": str(e)
        }

@register_tool
def unmerge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """
    Unmerge a range of cells.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Top-left cell of range to unmerge
        end_cell: Bottom-right cell of range to unmerge
        
    Returns:
        dict: {
            "success": bool,  # True if the cells were unmerged successfully
            "message": str    # A message describing the result
        }
    """
    try:
        full_path = get_excel_path(filepath)
        result = unmerge_range(full_path, sheet_name, start_cell, end_cell)
        return {
            "success": True,
            "message": result["message"]
        }
    except (ValidationError, SheetError) as e:
        return {
            "success": False,
            "message": f"Error: {str(e)}"
        }
    except Exception as e:
        logger.error(f"Error unmerging cells: {e}")
        return {
            "success": False,
            "message": str(e)
        }

@register_tool
def list_excel_files() -> Dict[str, Any]:
    """
    Returns a list of all Excel files available in the configured directory.

    Returns:
        dict: {
            "success": bool,  # True if the operation succeeded
            "files": list[str],  # List of Excel filenames (e.g., ["file1.xlsx", "file2.xlsm"])
            "message": str (optional)  # Error message if success is False
        }

    Example:
        list_excel_files()
        # -> { "success": True, "files": ["report.xlsx", "data.xlsm"] }
    """
    try:
        logger.debug(f"Listing Excel files in {settings.excel_mcp_folder}")
        # Ensure directory exists
        os.makedirs(settings.excel_mcp_folder, exist_ok=True)
        
        files = []
        for f in os.listdir(settings.excel_mcp_folder):
            if f.endswith('.xlsx') or f.endswith('.xlsm'):
                files.append(f)
        logger.debug(f"Found files: {files}")
        return {"success": True, "files": files}
    except Exception as e:
        logger.error(f"Error listing Excel files: {e}")
        return {"success": False, "message": str(e), "files": []}