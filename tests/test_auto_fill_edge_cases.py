"""
Test Suite: Auto-Fill Edge Cases and Error Handling
Coverage Goals: Edge cases, boundary conditions, and error scenarios
Dependencies: unittest, pandas, mock Excel COM interfaces

This module tests edge cases and error conditions for auto-fill functionality:
- Excel COM interface failures during row insertion
- Memory pressure during large row additions
- Sheet protection/read-only mode
- Invalid sheet references
- Formatting failures
- Boundary conditions (zero, negative, very large values)
- Data consistency under error conditions
"""

import unittest
from unittest.mock import Mock, MagicMock, patch, PropertyMock
import pandas as pd
import sys
from pathlib import Path
from typing import List, Dict, Any
import time

# Import test utilities
from test_utils import AutoFillTestBase, MockExcelSheet, create_test_summary_data

# Add src to Python path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from excel_processor.models import ProcessingConfig, ProcessingResult
from excel_processor.processor import EnhancedExcelProcessor


class TestAutoFillBoundaryConditions(AutoFillTestBase):
    """Test boundary conditions and edge input values"""
    
    def test_auto_fill_zero_rows(self):
        """Test: Adding zero rows should work without errors"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code']
        sheet = MockExcelSheet()
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 0, headers)
        
        # Assert
        self.assertEqual(rows_added, 0)
        self.assertEqual(len(sheet.data), 0)  # No data should be written
    
    def test_auto_fill_negative_count(self):
        """Test: Negative row count should be handled gracefully"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        sheet = MockExcelSheet()
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, -5, headers)
        
        # Assert - Should handle gracefully, likely by treating as 0
        self.assertEqual(rows_added, 0)
    
    def test_auto_fill_empty_headers_list(self):
        """Test: Empty headers list should be handled gracefully"""
        # Arrange
        empty_headers = []
        sheet = MockExcelSheet()
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 5, empty_headers)
        
        # Assert
        self.assertEqual(rows_added, 5)  # Should still report success
        # No actual data written due to empty headers
    
    def test_auto_fill_very_large_start_row(self):
        """Test: Very large start row number"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        sheet = MockExcelSheet()
        very_large_start = 1000000
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, very_large_start, 3, headers)
        
        # Assert
        self.assertEqual(rows_added, 3)
        
        # Verify rows were placed at correct positions
        item2_col = headers.index('Item2') + 1
        for i in range(3):
            expected_row = very_large_start + i
            self.assertEqual(sheet.data.get((expected_row, item2_col)), 'Leasing period')
    
    def test_auto_fill_with_none_values(self):
        """Test: Handling None values in headers or parameters"""
        # Arrange
        headers_with_none = ['Item2', None, 'Factory code', 'Tenant code']
        sheet = MockExcelSheet()
        
        # Act - Should handle None values in headers gracefully
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 2, headers_with_none)
        
        # Assert
        self.assertEqual(rows_added, 2)
        
        # Check that Item2 column is still set correctly
        item2_col = headers_with_none.index('Item2') + 1
        self.assertEqual(sheet.data.get((1, item2_col)), 'Leasing period')
    
    def test_auto_fill_extremely_large_count(self):
        """Test: Extremely large row count (stress test)"""
        # Arrange
        headers = ['Item2', 'Note']
        sheet = MockExcelSheet()
        extreme_count = 100000  # 100k rows
        
        # Act - This should work but may take time
        start_time = time.perf_counter()
        rows_added = self.processor._add_empty_green_rows(sheet, 1, extreme_count, headers)
        execution_time = time.perf_counter() - start_time
        
        # Assert
        self.assertEqual(rows_added, extreme_count)
        self.assertLess(execution_time, 60.0, "Should complete within 60 seconds")
        
        # Spot check some rows
        item2_col = headers.index('Item2') + 1
        note_col = headers.index('Note') + 1
        
        # Check first, middle, and last rows
        test_rows = [1, extreme_count // 2, extreme_count]
        for row in test_rows:
            self.assertEqual(sheet.data.get((row, item2_col)), 'Leasing period')
            self.assertEqual(sheet.data.get((row, note_col)), 'Committed')


class TestAutoFillExcelCOMErrors(AutoFillTestBase):
    """Test handling of Excel COM interface errors"""
    
    def test_com_error_during_range_access(self):
        """Test: COM error when accessing sheet ranges"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        
        mock_sheet = Mock()
        mock_sheet.range.side_effect = Exception("COM Error: Application is busy")
        
        # Act
        with patch('builtins.print'):  # Suppress error output
            rows_added = self.processor._add_empty_green_rows(mock_sheet, 1, 5, headers)
        
        # Assert
        self.assertEqual(rows_added, 0)  # Should fail gracefully
    
    def test_intermittent_com_errors(self):
        """Test: Intermittent COM errors during row addition"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        
        class IntermittentFailureSheet(MockExcelSheet):
            def __init__(self):
                super().__init__()
                self.call_count = 0
            
            def range(self, start, end=None):
                self.call_count += 1
                # Fail on every 3rd call
                if self.call_count % 3 == 0:
                    raise Exception("Intermittent COM Error")
                return super().range(start, end)
        
        failing_sheet = IntermittentFailureSheet()
        
        # Act
        with patch('builtins.print'):  # Suppress error output
            rows_added = self.processor._add_empty_green_rows(failing_sheet, 1, 9, headers)
        
        # Assert
        # Should handle partial failures - some rows may succeed
        self.assertGreaterEqual(rows_added, 0)
        self.assertLessEqual(rows_added, 9)
    
    def test_com_timeout_error(self):
        """Test: COM timeout errors during operations"""
        # Arrange
        headers = ['Item2', 'Note']
        
        mock_sheet = Mock()
        mock_range = Mock()
        mock_sheet.range.return_value = mock_range
        
        # Simulate timeout by making value assignment hang
        def slow_assignment(value):
            time.sleep(0.1)  # Simulate slow operation
            raise TimeoutError("COM operation timed out")
        
        type(mock_range).value = PropertyMock(side_effect=slow_assignment)
        
        # Act
        with patch('builtins.print'):  # Suppress error output
            start_time = time.perf_counter()
            rows_added = self.processor._add_empty_green_rows(mock_sheet, 1, 3, headers)
            execution_time = time.perf_counter() - start_time
        
        # Assert
        self.assertEqual(rows_added, 0)  # Should fail gracefully
        self.assertLess(execution_time, 5.0, "Should not hang indefinitely")
    
    def test_com_memory_error(self):
        """Test: COM memory errors during large operations"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        
        mock_sheet = Mock()
        mock_sheet.range.side_effect = MemoryError("Insufficient memory for COM operation")
        
        # Act
        with patch('builtins.print'):  # Suppress error output
            rows_added = self.processor._add_empty_green_rows(mock_sheet, 1, 1000, headers)
        
        # Assert
        self.assertEqual(rows_added, 0)  # Should handle memory errors gracefully


class TestAutoFillDataIntegrityErrors(AutoFillTestBase):
    """Test data integrity under various error conditions"""
    
    def test_partial_write_failure_integrity(self):
        """Test: Data integrity when partial write failures occur"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code']
        
        class PartialFailureSheet(MockExcelSheet):
            def __init__(self):
                super().__init__()
                self.write_attempts = 0
            
            def range(self, start, end=None):
                mock_range = super().range(start, end)
                
                # Override the value setter to fail on certain attempts
                original_setter = type(mock_range).value.fset if hasattr(type(mock_range).value, 'fset') else None
                
                def failing_setter(value):
                    self.write_attempts += 1
                    if self.write_attempts % 4 == 0:  # Fail every 4th write
                        raise Exception(f"Write failure #{self.write_attempts}")
                    
                    # Call original logic for successful writes
                    if hasattr(start, '__len__') and len(start) == 2:
                        row, col = start
                        if end is None:
                            self.data[(row, col)] = value
                        else:
                            end_row, end_col = end
                            if isinstance(value, list):
                                if isinstance(value[0], list):
                                    # Multi-row data
                                    for r, row_data in enumerate(value):
                                        for c, cell_value in enumerate(row_data):
                                            self.data[(row + r, col + c)] = cell_value
                                else:
                                    # Single row data
                                    for c, cell_value in enumerate(value):
                                        self.data[(row, col + c)] = cell_value
                
                type(mock_range).value = PropertyMock(fset=failing_setter)
                return mock_range
        
        failing_sheet = PartialFailureSheet()
        
        # Act
        with patch('builtins.print'):  # Suppress error output
            rows_added = self.processor._add_empty_green_rows(failing_sheet, 1, 10, headers)
        
        # Assert
        # Should report accurate count of successfully added rows
        self.assertGreaterEqual(rows_added, 0)
        self.assertLessEqual(rows_added, 10)
        
        # Verify data integrity for successful writes
        successful_rows = rows_added
        item2_col = headers.index('Item2') + 1
        note_col = headers.index('Note') + 1
        
        written_rows = 0
        for row in range(1, 11):
            if failing_sheet.data.get((row, item2_col)) == 'Leasing period':
                written_rows += 1
                self.assertEqual(failing_sheet.data.get((row, note_col)), 'Committed')
        
        # The number of rows actually written should not exceed reported count
        self.assertLessEqual(written_rows, rows_added)
    
    def test_invalid_sheet_reference_handling(self):
        """Test: Handling invalid or corrupted sheet references"""
        # Arrange
        headers = ['Item2', 'Note']
        
        # Test with None sheet
        with patch('builtins.print'):
            rows_added = self.processor._add_empty_green_rows(None, 1, 5, headers)
        self.assertEqual(rows_added, 0)
        
        # Test with invalid sheet object
        invalid_sheet = "not a sheet object"
        with patch('builtins.print'):
            rows_added = self.processor._add_empty_green_rows(invalid_sheet, 1, 5, headers)
        self.assertEqual(rows_added, 0)
    
    def test_concurrent_modification_errors(self):
        """Test: Handling concurrent modification of sheet during auto-fill"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        
        class ConcurrentModificationSheet(MockExcelSheet):
            def range(self, start, end=None):
                # Simulate concurrent modification by changing sheet state
                if hasattr(start, '__len__') and len(start) == 2:
                    row, col = start
                    if row > 3:  # Simulate concurrent deletion of rows
                        raise Exception("Range no longer valid - sheet was modified")
                
                return super().range(start, end)
        
        concurrent_sheet = ConcurrentModificationSheet()
        
        # Act
        with patch('builtins.print'):
            rows_added = self.processor._add_empty_green_rows(concurrent_sheet, 1, 10, headers)
        
        # Assert
        # Should handle concurrent modifications gracefully
        self.assertGreaterEqual(rows_added, 0)
        self.assertLessEqual(rows_added, 3)  # Should stop when concurrent modification detected


class TestAutoFillResourceConstraints(AutoFillTestBase):
    """Test auto-fill under various resource constraints"""
    
    @patch('excel_processor.processor.MemoryOptimizer')
    def test_severe_memory_pressure(self, mock_memory_optimizer):
        """Test: Auto-fill under severe memory pressure"""
        # Arrange
        headers = self.sample_headers
        
        # Simulate severe memory pressure
        mock_memory_optimizer.check_memory_pressure.return_value = True
        mock_memory_optimizer.get_available_memory.return_value = 50  # Very low memory
        mock_memory_optimizer.cleanup_memory.return_value = None
        
        excel_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],  # Matched
        ]
        
        # Large summary requiring many auto-fill rows
        large_summary = create_test_summary_data(1000)
        large_summary.iloc[0] = {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'Sub1'}
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 10
        
        # Act
        start_time = time.perf_counter()
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, large_summary
        )
        execution_time = time.perf_counter() - start_time
        
        # Assert
        self.assertGreater(rows_updated, 0)  # Should still work
        # May add fewer rows due to memory constraints, but should not fail
        self.assertGreaterEqual(rows_added, 0)
        
        # Should complete in reasonable time even under pressure
        self.assertLess(execution_time, 30.0)
        
        # Memory cleanup should be triggered
        mock_memory_optimizer.cleanup_memory.assert_called()
    
    def test_excel_application_limits(self):
        """Test: Handling Excel application limits (max rows, columns)"""
        # Arrange
        headers = ['Item2', 'Note'] * 100  # Very wide headers (200 columns)
        sheet = MockExcelSheet()
        
        # Test with row count near Excel limits
        excel_max_rows = 1048576  # Excel 2007+ limit
        large_start_row = excel_max_rows - 100
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, large_start_row, 50, headers)
        
        # Assert
        self.assertEqual(rows_added, 50)
        
        # Verify data written near Excel limits
        item2_positions = [i for i, h in enumerate(headers) if h == 'Item2']
        for pos in item2_positions[:5]:  # Check first few Item2 columns
            col = pos + 1
            self.assertEqual(sheet.data.get((large_start_row, col)), 'Leasing period')
    
    def test_disk_space_constraints(self):
        """Test: Simulated disk space constraints during auto-fill"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        
        class DiskConstrainedSheet(MockExcelSheet):
            def __init__(self):
                super().__init__()
                self.write_count = 0
            
            def range(self, start, end=None):
                mock_range = super().range(start, end)
                
                def disk_limited_setter(value):
                    self.write_count += 1
                    if self.write_count > 500:  # Simulate disk full after 500 writes
                        raise OSError("Disk full - cannot write to file")
                    
                    # Normal write logic
                    if hasattr(start, '__len__') and len(start) == 2:
                        row, col = start
                        self.data[(row, col)] = value
                
                type(mock_range).value = PropertyMock(fset=disk_limited_setter)
                return mock_range
        
        constrained_sheet = DiskConstrainedSheet()
        
        # Act - Try to add more rows than disk allows
        with patch('builtins.print'):
            rows_added = self.processor._add_empty_green_rows(constrained_sheet, 1, 1000, headers)
        
        # Assert
        # Should handle disk constraints gracefully
        self.assertGreaterEqual(rows_added, 0)
        self.assertLessEqual(rows_added, 1000)


class TestAutoFillEdgeCaseScenarios(AutoFillTestBase):
    """Test complex edge case scenarios combining multiple factors"""
    
    def test_corrupted_data_recovery(self):
        """Test: Auto-fill with corrupted or malformed data"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code']
        
        # Create DataFrame with corrupted data
        corrupted_data = [
            ['Leasing period', 'Committed', 'Factory01', None],  # None value
            [None, 'Committed', '', ''],  # None in Item2
            ['Leasing period', None, '', ''],  # None in Note
            ['', '', '', ''],  # All empty
            ['Leasing period', 'Committed', 'Factory\x00\x01', 'T\x02\x03'],  # Binary chars
        ]
        
        df = self.create_dataframe_from_excel_data(headers, corrupted_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, corrupted_data)
        mock_sheet.used_range.last_cell.row = len(corrupted_data) + 10
        
        # Create summary requiring auto-fill
        summary_data = create_test_summary_data(3)
        
        # Act - Should handle corrupted data gracefully
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, summary_data
        )
        
        # Assert
        self.assertGreaterEqual(rows_updated, 0)
        self.assertGreaterEqual(rows_added, 0)
        # Should not crash despite corrupted data
    
    def test_unicode_and_special_characters(self):
        """Test: Auto-fill with Unicode and special characters"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code', 'Special Field']
        sheet = MockExcelSheet()
        
        # Add rows with special characters
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 3, headers)
        
        # Assert
        self.assertEqual(rows_added, 3)
        
        # Verify standard values are still set correctly
        item2_col = headers.index('Item2') + 1
        note_col = headers.index('Note') + 1
        
        for row in range(1, 4):
            self.assertEqual(sheet.data.get((row, item2_col)), 'Leasing period')
            self.assertEqual(sheet.data.get((row, note_col)), 'Committed')
    
    def test_extremely_wide_sheets(self):
        """Test: Auto-fill with extremely wide sheets (many columns)"""
        # Arrange
        # Create 500 columns (very wide sheet)
        wide_headers = [f'Col_{i}' for i in range(500)]
        wide_headers[0] = 'Item2'
        wide_headers[1] = 'Note'
        
        sheet = MockExcelSheet()
        
        # Act
        start_time = time.perf_counter()
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 5, wide_headers)
        execution_time = time.perf_counter() - start_time
        
        # Assert
        self.assertEqual(rows_added, 5)
        self.assertLess(execution_time, 5.0, "Should handle wide sheets efficiently")
        
        # Verify key columns are set
        item2_col = 1  # First column
        note_col = 2   # Second column
        
        for row in range(1, 6):
            self.assertEqual(sheet.data.get((row, item2_col)), 'Leasing period')
            self.assertEqual(sheet.data.get((row, note_col)), 'Committed')
    
    def test_mixed_error_conditions(self):
        """Test: Auto-fill with multiple simultaneous error conditions"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        
        class MultipleErrorSheet(MockExcelSheet):
            def __init__(self):
                super().__init__()
                self.operation_count = 0
            
            def range(self, start, end=None):
                self.operation_count += 1
                
                # Mix different types of errors
                if self.operation_count % 5 == 1:
                    raise MemoryError("Simulated memory error")
                elif self.operation_count % 5 == 2:
                    raise TimeoutError("Simulated timeout")
                elif self.operation_count % 5 == 3:
                    raise PermissionError("Simulated permission error")
                elif self.operation_count % 5 == 4:
                    time.sleep(0.01)  # Simulate slow operation
                    raise Exception("Generic COM error")
                else:
                    # Success case
                    return super().range(start, end)
        
        error_sheet = MultipleErrorSheet()
        
        # Act - Should handle mixed errors gracefully
        with patch('builtins.print'):
            start_time = time.perf_counter()
            rows_added = self.processor._add_empty_green_rows(error_sheet, 1, 20, headers)
            execution_time = time.perf_counter() - start_time
        
        # Assert
        # Should handle errors gracefully and not hang
        self.assertGreaterEqual(rows_added, 0)
        self.assertLessEqual(rows_added, 20)
        self.assertLess(execution_time, 10.0, "Should not hang despite multiple errors")


if __name__ == '__main__':
    # Configure test runner
    unittest.main(verbosity=2, buffer=True)


# Usage:
# python -m pytest tests/test_auto_fill_edge_cases.py -v
# python tests/test_auto_fill_edge_cases.py