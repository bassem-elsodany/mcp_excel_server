"""
Logger utility for Excel MCP Server.

Provides a get_logger function that configures logging to write to daily log files
in the logs directory.
"""

import logging
import os
import sys
from typing import Optional
from pathlib import Path
from datetime import datetime
import json

from mcp_excel_server.config.settings import settings


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """
    Get a configured logger instance that writes to daily log files.
    Uses project settings if available, otherwise defaults.
    
    Args:
        name: Optional logger name. If not provided, uses "excel-mcp"
        
    Returns:
        Configured logger instance
    """
    logger = logging.getLogger(name or "excel-mcp")
    
    # Only configure if no handlers exist
    if not logger.handlers:
        try:
            # Get log level from settings
            log_level = getattr(logging, settings.log_level.upper())
            
            # Ensure logs directory exists
            logs_dir = Path(settings.excel_mcp_log_folder)
            logs_dir.mkdir(parents=True, exist_ok=True)
            
            # Get current date for log file
            current_date = datetime.now().strftime('%Y%m%d')
            log_file = logs_dir / f"excel-mcp-{current_date}.log"
            
            # Create a new file handler with explicit encoding
            file_handler = logging.FileHandler(
                str(log_file),  # Convert Path to string
                mode="a",
                encoding="utf-8"
            )
            
            # Set formatter with timestamp
            formatter = logging.Formatter(
                "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S"
            )
            file_handler.setFormatter(formatter)
            
            # Set level and add handler
            file_handler.setLevel(log_level)
            logger.addHandler(file_handler)
            logger.setLevel(log_level)
            
            # Disable propagation to root logger
            logger.propagate = False
            
            # Log initial message
            logger.info("Logger initialized")
            
        except Exception as e:
            # If there's an error configuring the logger, print to stderr
            print(f"Error configuring logger: {str(e)}", file=sys.stderr)
            # Configure a basic console logger as fallback
            basic_handler = logging.StreamHandler(sys.stderr)
            basic_handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
            logger.addHandler(basic_handler)
            logger.setLevel(logging.INFO)
            logger.error(f"Failed to configure logger: {str(e)}")
    
    return logger

def get_audit_logger():
    """Return a logger for audit events, writing to logs/excel-mcp-audit.log."""
    logger = logging.getLogger("excel-mcp-audit")
    if not logger.handlers:
        logs_dir = Path(settings.excel_mcp_log_folder)
        logs_dir.mkdir(parents=True, exist_ok=True)
        audit_file = logs_dir / "excel-mcp-audit.log"
        handler = logging.FileHandler(str(audit_file), encoding="utf-8")
        handler.setFormatter(logging.Formatter('%(message)s'))
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        logger.propagate = False
    return logger

def audit_event(action_type, details):
    """Log an audit event as a JSON line with timestamp, action, and details."""
    logger = get_audit_logger()
    event = {
        "timestamp": datetime.utcnow().isoformat(),
        "action": action_type,
        **details
    }
    logger.info(json.dumps(event)) 