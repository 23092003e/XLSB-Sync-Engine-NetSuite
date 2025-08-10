"""
Test utilities and fixtures for auto-fill logic testing

Provides common test fixtures, mock objects, and utility functions for testing
the XLSB processing system, particularly the auto-fill functionality.
"""

import unittest
from unittest.mock import Mock, MagicMock, patch, PropertyMock
import pandas as pd
from typing import List, Dict, Any, Optional
import sys
from pathlib import Path

# Add src to Python path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from excel_processor.models import ProcessingConfig, ProcessingResult
from excel_processor.processor import EnhancedExcelProcessor


class MockExcelSheet:
    """Mock Excel sheet object for testing"""
    
    def __init__(self, used_range_rows: int = 100, used_range_cols: int = 20):
        self.used_range_rows = used_range_rows
        self.used_range_cols = used_range_cols
        self.data = {}  # Dictionary to store cell data
        self.ranges = {}  # Dictionary to store range data
        
    @property
    def used_range(self):
        """Mock used range property"""
        mock_range = Mock()
        mock_range.last_cell.row = self.used_range_rows
        mock_range.last_cell.column = self.used_range_cols
        return mock_range
    
    def range(self, start, end=None):
        """Mock range method"""
        mock_range = Mock()
        
        # Handle single cell
        if end is None:
            row, col = start
            mock_range.value = self.data.get((row, col), '')
            return mock_range
            
        # Handle range
        start_row, start_col = start
        end_row, end_col = end
        
        if start_row == end_row:
            # Single row range
            values = []
            for col in range(start_col, end_col + 1):
                values.append(self.data.get((start_row, col), ''))
            mock_range.value = values
        else:
            # Multi-row range
            values = []
            for row in range(start_row, end_row + 1):
                row_values = []
                for col in range(start_col, end_col + 1):
                    row_values.append(self.data.get((row, col), ''))
                values.append(row_values)
            mock_range.value = values
            
        return mock_range
    
    def set_cell_value(self, row: int, col: int, value: Any):
        """Set a cell value for testing"""
        self.data[(row, col)] = value
    
    def set_range_values(self, start_row: int, start_col: int, values: List[List[Any]]):
        """Set range values for testing"""
        for r, row_values in enumerate(values):
            for c, value in enumerate(row_values):
                self.data[(start_row + r, start_col + c)] = value


class MockExcelWorkbook:
    """Mock Excel workbook object for testing"""
    
    def __init__(self, sheets: Dict[str, MockExcelSheet] = None):
        self.sheets_dict = sheets or {}
        self._saved = False
        self._closed = False
    
    @property
    def sheets(self):
        """Mock sheets property"""
        mock_sheets = Mock()
        
        # Mock indexing by name
        def getitem(name):
            if name in self.sheets_dict:
                return self.sheets_dict[name]
            raise KeyError(f"Sheet '{name}' not found")
        
        mock_sheets.__getitem__ = getitem
        return mock_sheets
    
    def save(self):
        """Mock save method"""
        self._saved = True
    
    def close(self):
        """Mock close method"""
        self._closed = True
    
    @property
    def is_saved(self):
        return self._saved
    
    @property
    def is_closed(self):
        return self._closed


class MockExcelApp:
    """Mock Excel application object for testing"""
    
    def __init__(self, workbooks: Dict[str, MockExcelWorkbook] = None):
        self.workbooks_dict = workbooks or {}
    
    @property
    def books(self):
        """Mock books property"""
        mock_books = Mock()
        
        def open_workbook(filepath):
            if filepath in self.workbooks_dict:
                return self.workbooks_dict[filepath]
            # Create a default workbook if not specified
            sheet = MockExcelSheet()
            workbook = MockExcelWorkbook({'1.Leasing income': sheet})
            return workbook
        
        mock_books.open = open_workbook
        return mock_books
    
    def quit(self):
        """Mock quit method"""
        pass


class AutoFillTestBase(unittest.TestCase):
    """Base test class for auto-fill functionality tests"""
    
    def setUp(self):
        """Set up test fixtures"""
        # Create test configuration
        self.config = ProcessingConfig(
            max_excel_instances=2,
            chunk_size=1000,
            memory_threshold_percent=70.0,
            timeout_seconds=300,
            column_mapping={
                'Unit name': 'Factory code',
                'Tenant ID': 'Tenant code',
                'Tenant': 'Tenant name'
            }
        )
        
        # Create processor instance
        self.processor = EnhancedExcelProcessor(self.config)
        
        # Create sample summary data
        self.sample_summary_data = pd.DataFrame([
            {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'Sub1'},
            {'Unit name': 'Factory01', 'Tenant ID': 'T002', 'Tenant': 'Tenant B', 'Subsidiary': 'Sub1'},
            {'Unit name': 'Factory02', 'Tenant ID': 'T003', 'Tenant': 'Tenant C', 'Subsidiary': 'Sub1'},
            {'Unit name': 'Factory02', 'Tenant ID': 'T004', 'Tenant': 'Tenant D', 'Subsidiary': 'Sub1'},
            {'Unit name': 'Factory03', 'Tenant ID': 'T005', 'Tenant': 'Tenant E', 'Subsidiary': 'Sub1'},
        ])
        
        # Create sample Excel data (existing green rows)
        self.sample_headers = ['Item2', 'Note', 'Factory code', 'Tenant code', 'Tenant name', 'Rent (USD)']
        self.sample_excel_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],  # Matched row
            ['Leasing period', 'Committed', 'Factory01', 'T002', 'Tenant B', '1500'],  # Matched row  
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty green row 1
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty green row 2
        ]
        
        # Set up processor with sample data
        self.processor.summary_data = self.sample_summary_data
        self.processor.summary_lookup = {}
        for idx, row in self.sample_summary_data.iterrows():
            k1 = f"{row['Unit name'].strip()}|{row['Tenant ID'].strip()}"
            k2 = f"{row['Unit name'].strip()}|{row['Tenant'].strip()}"
            self.processor.summary_lookup[k1] = (idx, row.to_dict())
            self.processor.summary_lookup[k2] = (idx, row.to_dict())
    
    def create_mock_sheet_with_data(self, headers: List[str], data: List[List[str]], 
                                   header_row: int = 1) -> MockExcelSheet:
        """Create a mock Excel sheet with specific data"""
        sheet = MockExcelSheet(used_range_rows=len(data) + header_row, 
                              used_range_cols=len(headers))
        
        # Set headers
        for col, header in enumerate(headers):
            sheet.set_cell_value(header_row, col + 1, header)
        
        # Set data
        for row_idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data):
                sheet.set_cell_value(header_row + 1 + row_idx, col_idx + 1, value)
        
        return sheet
    
    def create_dataframe_from_excel_data(self, headers: List[str], 
                                       data: List[List[str]]) -> pd.DataFrame:
        """Create a pandas DataFrame from Excel-like data"""
        return pd.DataFrame(data, columns=headers).astype(object).fillna('')
    
    def assert_auto_fill_results(self, rows_updated: int, rows_added: int, 
                               expected_updated: int, expected_added: int):
        """Assert auto-fill results match expectations"""
        self.assertEqual(rows_updated, expected_updated, 
                        f"Expected {expected_updated} rows updated, got {rows_updated}")
        self.assertEqual(rows_added, expected_added, 
                        f"Expected {expected_added} rows added, got {rows_added}")
    
    def get_empty_green_rows_count(self, df: pd.DataFrame) -> int:
        """Count empty green rows in a DataFrame"""
        empty_green_mask = (
            (df['Item2'].astype(str).str.strip() == 'Leasing period') &
            (df['Note'].astype(str).str.strip() == 'Committed') &
            ((df['Factory code'].astype(str).str.strip() == '') |
             (df['Tenant code'].astype(str).str.strip() == '') |
             (df['Tenant name'].astype(str).str.strip() == ''))
        )
        return len(df[empty_green_mask])
    
    def get_unmatched_summary_count(self, df: pd.DataFrame, summary_data: pd.DataFrame) -> int:
        """Calculate unmatched summary count"""
        # Simulate the matching logic from the processor
        matched_indices = set()
        
        for _, row in df.iterrows():
            if (row['Item2'] == 'Leasing period' and row['Note'] == 'Committed' and
                row['Factory code'] != '' and row['Tenant code'] != ''):
                
                k1 = f"{row['Factory code']}|{row['Tenant code']}"
                k2 = f"{row['Factory code']}|{row['Tenant name']}"
                
                # Check if this matches any summary row
                for idx, srow in summary_data.iterrows():
                    sk1 = f"{srow['Unit name']}|{srow['Tenant ID']}"
                    sk2 = f"{srow['Unit name']}|{srow['Tenant']}"
                    
                    if k1 == sk1 or k2 == sk2:
                        matched_indices.add(idx)
                        break
        
        return len(summary_data) - len(matched_indices)


def create_test_summary_data(count: int = 10) -> pd.DataFrame:
    """Create test summary data with specified number of rows"""
    data = []
    for i in range(count):
        data.append({
            'Unit name': f'Factory{i+1:02d}',
            'Tenant ID': f'T{i+1:03d}',
            'Tenant': f'Tenant {chr(65+i)}',
            'Subsidiary': 'TestSub'
        })
    return pd.DataFrame(data)


def create_large_test_data(rows: int = 1000) -> pd.DataFrame:
    """Create large test dataset for performance testing"""
    import random
    import string
    
    data = []
    for i in range(rows):
        data.append({
            'Unit name': f'Factory{random.randint(1,50):02d}',
            'Tenant ID': f'T{i+1:04d}',
            'Tenant': f'Tenant {"".join(random.choices(string.ascii_uppercase, k=3))}',
            'Subsidiary': f'Sub{random.randint(1,5)}'
        })
    return pd.DataFrame(data)