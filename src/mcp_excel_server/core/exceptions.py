class ExcelMCPError(Exception):
    """Base exception for Excel MCP errors."""
    pass

class WorkbookError(ExcelMCPError):
    """Raised when workbook operations fail."""
    pass

class SheetError(ExcelMCPError):
    """Raised when sheet operations fail."""
    pass

class DataError(ExcelMCPError):
    """Raised when data operations fail."""
    pass

class ValidationError(ExcelMCPError):
    """Raised when validation fails."""
    pass

class RangeError(ExcelMCPError):
    """Raised when range operations fail."""
    pass
