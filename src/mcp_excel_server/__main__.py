"""
CLI entry point for Excel MCP Server.
This module provides the command-line interface for starting the server.
"""

import asyncio
import typer

from mcp_excel_server.api.server import run_sse, run_stdio
from mcp_excel_server.utils import get_logger

# Initialize logger
logger = get_logger(__name__)

app = typer.Typer(help="Excel MCP Server")

@app.command()
def sse():
    """Start Excel MCP Server in SSE mode"""
    logger.info("Excel MCP Server - SSE mode")
    logger.info("----------------------")
    logger.info("Press Ctrl+C to exit")
    try:
        asyncio.run(run_sse())
    except KeyboardInterrupt:
        logger.info("\nShutting down server...")
    except Exception as e:
        logger.error(f"\nError: {e}")
        import traceback
        traceback.print_exc()
    finally:
        logger.info("Service stopped.")

@app.command()
def stdio():
    """Start Excel MCP Server in stdio mode"""
    try:
        run_stdio()
    except KeyboardInterrupt:
        logger.info("\nShutting down server...")
    except Exception as e:
        logger.error(f"\nError: {e}")
        import traceback
        traceback.print_exc()
    finally:
        logger.info("Service stopped.")

if __name__ == "__main__":
    app() 