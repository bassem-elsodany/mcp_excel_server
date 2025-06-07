"""
Workbook management tools for Excel MCP Server.
This module provides tools for managing Excel workbooks.
"""

from typing import Optional, List, Dict, Any
import os
from pathlib import Path

from mcp_excel_server.api.registry import register_tool
from mcp_excel_server.core.workbook import (
    create_workbook,
    get_workbook,
    save_workbook,
    get_workbook_info
)
from mcp_excel_server.core.exceptions import WorkbookError
from mcp_excel_server.config.settings import settings

from mcp_excel_server.utils import get_logger
# Initialize logger
logger = get_logger(__name__)


@register_tool
def create_workbook_tool(
    filename: str,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """
    Creates a new Excel workbook file.

    Args:
        filename (str): The name of the Excel file to create (e.g., "report.xlsx").
        sheet_name (str, optional): The name of the initial worksheet. If not provided, a default name is used.

    Returns:
        dict: {
            "success": bool,  # True if the workbook was created successfully
            "message": str,   # A message describing the result
            "filename": str   # The name of the created file
        }

    Example:
        create_workbook_tool(filename="report.xlsx", sheet_name="Summary")
    """
    try:
        path = Path(filename)
        if path.exists():
            return {
                "success": False,
                "message": f"Workbook already exists: {filename}"
            }
            
        wb = create_workbook(filename, sheet_name or "Sheet1")
        save_workbook(wb["workbook"], filename)
        return {
            "success": True,
            "message": f"Created workbook '{filename}'",
            "filename": filename
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e)
        }

@register_tool
def get_workbook_info_tool(
    filename: str,
    include_ranges: bool = False
) -> Dict[str, Any]:
    """
    Retrieves information about an Excel workbook, such as sheet names and properties.

    Args:
        filename (str): The name of the Excel file to inspect.
        include_ranges (bool, optional): If True, include named range information.

    Returns:
        dict: {
            "success": bool,  # True if the operation succeeded
            "message": str,   # A message describing the result
            "info": {
                "sheets": list[str],      # List of worksheet names
                "ranges": list[str],      # List of named ranges (if requested)
                "properties": dict        # Workbook properties (e.g., author, created date)
            }
        }

    Example:
        get_workbook_info_tool(filename="report.xlsx", include_ranges=True)
    """
    try:
        info = get_workbook_info(filename, include_ranges=include_ranges)
        return {
            "success": True,
            "message": f"Retrieved workbook info for '{filename}'",
            "info": info
        }
    except WorkbookError as e:
        return {
            "success": False,
            "message": str(e)
        }