"""
Settings management for Excel MCP Server.

This module provides functionality to load and manage environment-specific configuration
settings from .env.{env} files. It uses pydantic for settings validation and management.
"""

import os
from datetime import datetime
from pathlib import Path
from typing import Optional, List
from functools import lru_cache
from dotenv import load_dotenv
from pydantic_settings import BaseSettings
from pydantic import ConfigDict, field_validator, Field


# Load environment variables from .env (base) and then environment-specific .env file
load_dotenv(f'.env')

class AppSettings(BaseSettings):
    """Settings for Excel MCP Server.
    
    This class defines all configuration settings that can be loaded from environment
    files or environment variables. Settings are validated using pydantic.
    """
    
    # Excel files directory path
    excel_mcp_folder: str = os.getenv("EXCEL_MCP_FOLDER", "excel_files")
    excel_mcp_log_folder: str = os.getenv("EXCEL_MCP_LOG_FOLDER", "logs")
    
    # Logging configuration
    log_level: str = os.getenv("EXCEL_MCP_LOG_LEVEL", "INFO")
    log_file: str = os.getenv("EXCEL_MCP_LOG_FILE", f"excel-mcp-{datetime.now().strftime('%Y%m%d')}.log")
    file_log_level: str = os.getenv("EXCEL_MCP_FILE_LOG_LEVEL", "DEBUG")
    
    # Server configuration
    server_host: str = os.getenv("EXCEL_MCP_SERVER_HOST", "localhost")
    server_port: int = int(os.getenv("EXCEL_MCP_SERVER_PORT", "8000"))

    tool_prefix: str = os.getenv("EXCEL_MCP_SERVER_TOOL_PREFIX", "EXCEL_MCP_")
    
    # Excel operation settings
    max_rows_per_sheet: int = int(os.getenv("EXCEL_MCP_MAX_ROWS_PER_SHEET", "1048576"))
    max_columns_per_sheet: int = int(os.getenv("EXCEL_MCP_MAX_COLUMNS_PER_SHEET", "16384"))
    version: str = os.getenv("EXCEL_MCP_SERVER_VERSION", "0.1.0")
    
    @field_validator("log_level", "file_log_level")
    def validate_log_level(cls, v: str) -> str:
        """Validate logging level."""
        allowed_levels = {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}
        if v.upper() not in allowed_levels:
            raise ValueError(f"Log level must be one of: {', '.join(allowed_levels)}")
        return v.upper()
    
    @field_validator("server_port")
    def validate_port(cls, v: int) -> int:
        """Validate server port number."""
        if not 1 <= v <= 65535:
            raise ValueError("Port must be between 1 and 65535")
        return v
    
    @field_validator("excel_mcp_folder", "excel_mcp_log_folder")
    def validate_paths(cls, v: str) -> str:
        """Validate and create directory for paths if needed."""
        Path(v).mkdir(parents=True, exist_ok=True)
        return v

@lru_cache()
def get_settings() -> AppSettings:
    """Get cached settings instance"""
    return AppSettings()

# Create a global settings instance
settings = get_settings() 