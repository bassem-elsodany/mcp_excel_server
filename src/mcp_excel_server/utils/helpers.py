"""
Helper utilities for Excel MCP server.
This module provides various utility functions for the Excel MCP server,
including path handling, validation, and other helper methods.
"""

import os
from typing import Optional
from mcp_excel_server.config.settings import settings
from mcp_excel_server.utils import get_logger

logger = get_logger(__name__)

def get_excel_path(filename: str) -> str:
    """Get full path to Excel file.
    
    Args:
        filename: Name of Excel file
        
    Returns:
        Full path to Excel file
        
    Raises:
        ValueError: If filename is not an absolute path when not in SSE mode
    """
    logger.debug(f"get_excel_path called with filename: {filename}")
    # If filename is already an absolute path, return it
    if os.path.isabs(filename):
        logger.debug(f"Filename is absolute path: {filename}")
        return filename

    # Check if excel_files_path is set in settings
    logger.debug(f"settings.excel_mcp_folder: {settings.excel_mcp_folder}")
    if not settings.excel_mcp_folder:
        logger.error(f"excel_mcp_folder is not set in settings. Filename: {filename}")
        # Must use absolute path
        raise ValueError(f"Invalid filename: {filename}, must be an absolute path when excel_files_path is not set in settings")

    # If it's a relative path, resolve it based on settings.excel_files_path
    full_path = os.path.join(str(settings.excel_mcp_folder), filename)
    logger.debug(f"Resolved full path: {full_path}")
    return full_path 