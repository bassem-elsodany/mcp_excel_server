"""
Test cases for Excel MCP Server workbook tools.

This module contains test cases for all workbook tools including:
- create_workbook_tool
- get_workbook_info_tool
"""

import os
import pytest
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
from openpyxl import Workbook

from mcp.server.fastmcp import FastMCP
from mcp_excel_server.core.exceptions import ValidationError, DataError, SheetError
from mcp_excel_server.config.settings import settings
from mcp_excel_server.api.tools.workbook import (
    create_workbook_tool,
    get_workbook_info_tool
)

# Test data
TEST_WORKBOOK = f"{settings.excel_mcp_folder}/test_workbook.xlsx"
TEST_NONE_EXISTENT_WORKBOOK = f"{settings.excel_mcp_folder}/nonexistent.xlsx"
TEST_SHEET = "Sheet1"

@pytest.fixture
def mock_mcp():
    """Create a mock FastMCP server instance."""
    return Mock(spec=FastMCP)

@pytest.fixture
def test_workbook():
    """Create a test workbook with sample data."""
    # Ensure test directory exists
    os.makedirs(settings.excel_mcp_folder, exist_ok=True)
    
    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet()
    ws.title = TEST_SHEET
    
    # Add some test data
    ws['A1'] = "Test1"
    ws['A2'] = "Test2"
    ws['B1'] = "Test3"
    ws['B2'] = "Test4"
    
    # Save workbook
    wb.save(TEST_WORKBOOK)
    
    yield TEST_WORKBOOK
    
    # Cleanup
    if os.path.exists(TEST_WORKBOOK):
        os.remove(TEST_WORKBOOK)
    if os.path.exists(TEST_NONE_EXISTENT_WORKBOOK):
        os.remove(TEST_NONE_EXISTENT_WORKBOOK)

class TestCreateWorkbookTool:
    """Test cases for create_workbook_tool."""
    
    def test_create_workbook_success(self, mock_mcp):
        """Test successful workbook creation."""
        result = create_workbook_tool(TEST_WORKBOOK)
        
        assert isinstance(result, dict)
        assert result["success"] is True
        assert "Created workbook" in result["message"]
        assert os.path.exists(TEST_WORKBOOK)
        
    def test_create_workbook_existing(self, mock_mcp, test_workbook):
        """Test creating workbook that already exists."""
        result = create_workbook_tool(TEST_WORKBOOK)
        
        assert isinstance(result, dict)
        assert result["success"] is False
        assert "already exists" in result["message"].lower()
        

class TestGetWorkbookInfoTool:
    """Test cases for get_workbook_info_tool."""
    
    def test_get_workbook_info_success(self, mock_mcp, test_workbook):
        """Test successful workbook info retrieval."""
        result = get_workbook_info_tool(TEST_WORKBOOK)
        
        assert isinstance(result, dict)
        assert result["success"] is True
        assert "info" in result
        assert "sheets" in result["info"]
        assert TEST_SHEET in result["info"]["sheets"]
        
    def test_get_workbook_info_nonexistent(self, mock_mcp):
        """Test getting info for nonexistent workbook."""
        result = get_workbook_info_tool(TEST_NONE_EXISTENT_WORKBOOK)
        
        assert isinstance(result, dict)
        assert result["success"] is False
        assert "not found" in result["message"].lower()
        
        
            