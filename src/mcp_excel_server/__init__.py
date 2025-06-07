"""
Excel MCP Server - A Message Control Protocol tool for Excel operations.
"""

from mcp_excel_server.config.settings import settings

from mcp_excel_server.core.exceptions import ExcelMCPError
from mcp_excel_server.core.workbook import Workbook

__all__ = ["ExcelMCPError", "Workbook"] 