"""
Worksheet management tools for Excel MCP Server.
This module provides tools for managing worksheets in Excel workbooks.
"""

from typing import Optional, List, Dict, Any

from mcp_excel_server.api.registry import register_tool
from mcp_excel_server.core.workbook import get_workbook, save_workbook
from mcp_excel_server.core.sheet import (
    create_sheet,
    copy_sheet,
    delete_sheet,
    rename_sheet,
    move_sheet,
    get_sheet,
    list_sheets
)
from mcp_excel_server.core.exceptions import SheetError

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
    try:
        wb = get_workbook(filename)
        if sheet_name in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{sheet_name}' already exists",
                "sheet_name": sheet_name
            }
            
        wb.create_sheet(sheet_name)
        save_workbook(wb, filename)
        return {
            "success": True,
            "message": f"Created worksheet '{sheet_name}'",
            "sheet_name": sheet_name
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e),
            "sheet_name": sheet_name
        }

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
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found"
            }
            
        if len(wb.sheetnames) == 1:
            return {
                "success": False,
                "message": "Cannot delete the only sheet in workbook"
            }
            
        del wb[sheet_name]
        save_workbook(wb, filename)
        return {
            "success": True,
            "message": f"Deleted worksheet '{sheet_name}'"
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e)
        }

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
    try:
        wb = get_workbook(filename)
        if old_name not in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{old_name}' not found",
                "new_name": new_name
            }
            
        if new_name in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{new_name}' already exists",
                "new_name": new_name
            }
            
        rename_sheet(filename, old_name, new_name)
        save_workbook(wb, filename)
        return {
            "success": True,
            "message": f"Renamed worksheet '{old_name}' to '{new_name}'",
            "new_name": new_name
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e),
            "new_name": new_name
        }

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
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found",
                "new_name": new_name
            }
            
        if new_name in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{new_name}' already exists",
                "new_name": new_name
            }
            
        copy_sheet(filename, sheet_name, new_name)
        save_workbook(wb, filename)
        return {
            "success": True,
            "message": f"Copied worksheet '{sheet_name}' to '{new_name}'",
            "new_name": new_name
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e),
            "new_name": new_name
        }

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
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found",
                "sheet_name": sheet_name
            }
            
        if index < 0 or index >= len(wb.sheetnames):
            return {
                "success": False,
                "message": f"Invalid index {index}. Must be between 0 and {len(wb.sheetnames)-1}",
                "sheet_name": sheet_name
            }
            
        sheet = wb[sheet_name]
        wb.move_sheet(sheet_name, offset=index - wb.index(sheet))
        save_workbook(wb, filename)
        return {
            "success": True,
            "message": f"Moved worksheet '{sheet_name}' to position {index}",
            "sheet_name": sheet_name
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e),
            "sheet_name": sheet_name
        }

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
    try:
        wb = get_workbook(filename)
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "message": f"Sheet '{sheet_name}' not found",
                "sheet": None
            }
            
        sheet = get_sheet(filename, sheet_name)
        return {
            "success": True,
            "message": f"Retrieved worksheet '{sheet_name}'",
            "sheet": {"title": sheet["title"]}
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e),
            "sheet": None
        }

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
    try:
        wb = get_workbook(filename)
        sheets = list_sheets(filename)
        return {
            "success": True,
            "message": f"Listed worksheets in '{filename}'",
            "sheets": sheets
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e),
            "sheets": []
        }


