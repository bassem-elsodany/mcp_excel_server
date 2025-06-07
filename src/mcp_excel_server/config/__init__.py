"""
Configuration package for Excel MCP Server.

This package handles loading and managing environment-specific configuration settings
from .env.{env} files.
"""

from .settings import AppSettings, get_settings

__all__ = ["AppSettings", "get_settings"] 