"""
Tool registration module for Excel MCP Server.
This module provides the tool registration functionality used by the server.
All registered tool names are automatically prefixed with 'BOI_'.
"""

from typing import Callable
from mcp.server.fastmcp import FastMCP
import functools

from mcp_excel_server.config.settings import settings

# Prefix for all tool names
TOOL_PREFIX = settings.tool_prefix

# Initialize FastMCP server
mcp = FastMCP(
    name="mcp-excel-server",
    version=settings.version,
    description="Excel MCP Server for manipulating Excel files",
    port=settings.server_port,
    host=settings.server_host,
    dependencies=[
        "openpyxl>=3.1.2",
        "pandas>=2.0.0",  # Required for data operations
        "numpy>=1.24.0"   # Required for calculations
    ]
)

def register_tool(func: Callable) -> Callable:
    """Register a function as an MCP tool with the 'BOI_' prefix.
    
    This decorator registers functions as tools with the MCP server, automatically
    prefixing the tool name with 'BOI_'.
    
    Args:
        func: The function to register as a tool
    
    Returns:
        The decorated function
    """
    tool_name = TOOL_PREFIX + func.__name__
    return mcp.tool(name=tool_name)(func)

# Export mcp and register_tool for use in other modules
__all__ = ['mcp', 'register_tool'] 