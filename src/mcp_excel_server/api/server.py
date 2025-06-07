"""
Excel MCP Server main entry point.

This script starts and manages the Excel MCP server, which allows users to control and automate Excel files through natural language commands or API requests. It supports two modes:

- SSE (Server-Sent Events) mode for real-time API communication
- stdio mode for command-line integration and testing

The server loads and registers all available Excel tools (workbook, worksheet, chart, pivot, formatting, etc.), sets up logging, and handles server startup and shutdown. Beginners should use this file to launch the server and understand how the API is exposed.
"""

# Standard library imports
import os
from typing import Callable

# Third-party imports
from mcp.server.fastmcp import FastMCP


from mcp_excel_server.config.settings import settings
from mcp_excel_server.api.registry import mcp, register_tool

# Import tools to ensure they are registered
import mcp_excel_server.api.tools

# Import exceptions
from mcp_excel_server.core.exceptions import (
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
)
from mcp_excel_server.utils import get_logger
# Initialize logger
logger = get_logger(__name__)

# Server configuration
SERVER_NAME = "excel-mcp-server"
SERVER_DESCRIPTION = "Excel MCP Server for manipulating Excel files"
SERVER_DEPENDENCIES = [
    "openpyxl>=3.1.2",
    "pandas>=2.0.0",  # Required for data operations
    "numpy>=1.24.0"   # Required for calculations
]

async def run_sse():
    """Run Excel MCP server in SSE (Server-Sent Events) mode.
    
    This function starts the server in SSE mode, which allows for real-time
    communication with clients. The server will listen for connections and
    handle tool requests asynchronously.
    
    The server can be stopped with a keyboard interrupt (Ctrl+C).
    """
    excel_files_path = settings.excel_mcp_folder
    os.makedirs(str(excel_files_path), exist_ok=True)
    
    try:
        # Log server startup
        logger.info("=" * 50)
        logger.info("MCP Excel Server Starting")
        logger.info("=" * 50)
        logger.info(f"Version: {settings.version}")
        logger.info(f"Files directory: {excel_files_path}")
        logger.info(f"Log level: {settings.log_level}")
        logger.info(f"Server running on http://{settings.server_host}:{settings.server_port}")
        logger.info("=" * 50)
        
        # Run the server
        await mcp.run_sse_async()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Shutting down server...")
        # No explicit shutdown needed as the server will clean up automatically
        logger.info("Server shutdown complete")

def run_stdio():
    """Run Excel MCP server in stdio mode.
    
    This function starts the server in stdio mode, which processes requests
    through standard input/output. This mode is useful for command-line
    integration and testing.
    
    The server can be stopped with a keyboard interrupt (Ctrl+C).
    """
    try:
        # Log server startup
        logger.info("=" * 50)
        logger.info("MCP Excel Server Starting (stdio mode)")
        logger.info("=" * 50)
        logger.info(f"Version: {settings.version}")
        logger.info(f"Log level: {settings.log_level}")
        logger.info("=" * 50)
        
        mcp.run(transport="stdio")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")

# Export mcp and register_tool for use in other modules
__all__ = ['mcp', 'register_tool'] 