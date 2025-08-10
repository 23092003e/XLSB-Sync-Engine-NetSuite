# excel_processor/exceptions.py
"""Custom exception types for Excel processing operations"""

class ExcelProcessingError(Exception):
    """Base exception for Excel processing errors"""
    pass

class COMInitializationError(ExcelProcessingError):
    """Raised when COM interface initialization fails"""
    pass

class ExcelAppCreationError(ExcelProcessingError):
    """Raised when Excel application cannot be created"""
    pass

class WorkbookOpenError(ExcelProcessingError):
    """Raised when workbook cannot be opened"""
    pass

class SheetNotFoundError(ExcelProcessingError):
    """Raised when required sheet is not found"""
    pass

class HeaderRowNotFoundError(ExcelProcessingError):
    """Raised when header row cannot be located"""
    pass

class DataReadError(ExcelProcessingError):
    """Raised when data reading fails"""
    pass

class MemoryConstraintError(ExcelProcessingError):
    """Raised when memory constraints prevent processing"""
    pass

class SummaryDataError(ExcelProcessingError):
    """Raised when summary data operations fail"""
    pass

class RetryableError(ExcelProcessingError):
    """Base class for errors that should trigger retry logic"""
    pass

class COMTimeoutError(RetryableError):
    """Raised when COM operations timeout"""
    pass

class ExcelHangError(RetryableError):
    """Raised when Excel application appears to hang"""
    pass