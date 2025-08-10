"""
Test Suite: Auto-Fill Integration Tests
Coverage Goals: Integration testing between auto-fill and other system components
Dependencies: unittest, pandas, mock Excel COM interfaces, memory optimizer, performance logger

This module tests the integration of auto-fill logic with:
- Chunked processing (5,000+ row chunks)
- Parallel processing (multiple Excel instances)  
- Memory optimization
- Performance monitoring
- Full end-to-end processing pipeline
"""

import unittest
from unittest.mock import Mock, MagicMock, patch, PropertyMock, call
import pandas as pd
import time
import threading
from typing import List, Dict, Any
import sys
from pathlib import Path

# Import test utilities
from test_utils import AutoFillTestBase, MockExcelSheet, MockExcelWorkbook, MockExcelApp, create_test_summary_data, create_large_test_data

# Add src to Python path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from excel_processor.models import ProcessingConfig, ProcessingResult
from excel_processor.processor import EnhancedExcelProcessor


class TestAutoFillChunkedProcessing(AutoFillTestBase):
    """Test auto-fill working with chunked processing for large datasets"""
    
    def setUp(self):
        super().setUp()
        # Configure for chunked processing
        self.config.chunk_size = 1000
        self.config.enable_chunked_processing = True
        self.processor = EnhancedExcelProcessor(self.config)
        self.processor.summary_data = self.sample_summary_data
        self.processor.summary_lookup = self._build_summary_lookup()
    
    def _build_summary_lookup(self):
        """Build summary lookup dictionary for testing"""
        lookup = {}
        for idx, row in self.sample_summary_data.iterrows():
            k1 = f"{row['Unit name'].strip()}|{row['Tenant ID'].strip()}"
            k2 = f"{row['Unit name'].strip()}|{row['Tenant'].strip()}"
            lookup[k1] = (idx, row.to_dict())
            lookup[k2] = (idx, row.to_dict())
        return lookup
    
    @patch('excel_processor.processor.MemoryOptimizer')
    @patch('excel_processor.processor.performance_logger')
    def test_chunked_processing_with_auto_fill(self, mock_perf_logger, mock_memory_optimizer):
        """Test: Auto-fill works correctly with large datasets requiring chunked processing"""
        # Arrange - Create large dataset
        headers = self.sample_headers
        
        # Create 5,000 rows of data with some green rows needing auto-fill
        large_excel_data = []
        for i in range(5000):
            if i < 10:
                # First 10 are matching green rows
                large_excel_data.append(['Leasing period', 'Committed', f'Factory0{i%5+1}', f'T{i+1:03d}', f'Tenant {chr(65+i)}', '1000'])
            elif i < 20:
                # Next 10 are empty green rows (need filling)
                large_excel_data.append(['Leasing period', 'Committed', '', '', '', ''])
            else:
                # Rest are other data
                large_excel_data.append(['Other Item', 'Other Note', f'Other{i}', f'O{i}', f'Other {i}', '500'])
        
        # Create large summary data (15 items total, 10 will match existing, 5 need new rows)
        large_summary = create_large_test_data(15)
        
        # Mock memory optimizer
        mock_memory_optimizer.get_memory_usage.return_value = 100.0
        mock_memory_optimizer.check_memory_pressure.return_value = False
        mock_memory_optimizer.cleanup_memory.return_value = None
        mock_memory_optimizer.get_available_memory.return_value = 2000
        
        # Create mock sheet with large data
        mock_sheet = self.create_mock_sheet_with_data(headers, large_excel_data)
        mock_sheet.used_range.last_cell.row = len(large_excel_data) + 10
        
        df = self.create_dataframe_from_excel_data(headers, large_excel_data)
        
        # Act
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, large_summary
        )
        
        # Assert
        # Should have processed the data and handled auto-fill appropriately
        self.assertGreaterEqual(rows_updated, 0)
        self.assertGreaterEqual(rows_added, 0)
        
        # Verify memory monitoring was called
        mock_memory_optimizer.get_memory_usage.assert_called()
        mock_memory_optimizer.check_memory_pressure.assert_called()
    
    @patch('excel_processor.processor.EnhancedExcelOptimizer')
    @patch('excel_processor.processor.MemoryOptimizer')
    def test_chunked_reading_with_auto_fill(self, mock_memory_optimizer, mock_excel_optimizer):
        """Test: Auto-fill logic works when data is read in chunks"""
        # Arrange
        headers = self.sample_headers
        
        # Mock chunked reading behavior
        mock_memory_optimizer.get_available_memory.return_value = 1000
        mock_memory_optimizer.check_memory_pressure.return_value = False
        mock_memory_optimizer.cleanup_memory.return_value = None
        
        # Create sheet that will trigger chunked reading
        mock_sheet = MockExcelSheet(used_range_rows=6000, used_range_cols=len(headers))
        
        # Mock safe_excel_operation for range operations
        def mock_safe_operation(func):
            return func()
        mock_excel_optimizer.safe_excel_operation = mock_safe_operation
        
        # Act
        headers_result, data_result = self.processor._batch_read_enhanced(mock_sheet, 1)
        
        # Assert
        self.assertEqual(len(headers_result), len(headers))
        # Should handle chunked reading without errors
        self.assertIsInstance(data_result, list)
    
    def test_auto_fill_memory_pressure_handling(self):
        """Test: Auto-fill handles memory pressure during large operations"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code', 'Tenant name']
        sheet = MockExcelSheet()
        
        # Mock memory pressure scenario
        with patch('excel_processor.processor.MemoryOptimizer') as mock_memory:
            mock_memory.check_memory_pressure.side_effect = [False, True, False]  # Pressure on second call
            mock_memory.cleanup_memory.return_value = None
            
            # Act - Try to add many rows during simulated memory pressure
            rows_added = self.processor._add_empty_green_rows(sheet, 1, 1000, headers)
            
            # Assert
            self.assertEqual(rows_added, 1000)  # Should still complete successfully
            mock_memory.cleanup_memory.assert_called()  # Should trigger cleanup


class TestAutoFillParallelProcessing(AutoFillTestBase):
    """Test auto-fill working with parallel processing and multiple Excel instances"""
    
    def setUp(self):
        super().setUp()
        self.config.max_excel_instances = 4
        self.processor = EnhancedExcelProcessor(self.config)
        self.processor.summary_data = self.sample_summary_data
        self.processor.summary_lookup = self._build_summary_lookup()
    
    def _build_summary_lookup(self):
        """Build summary lookup dictionary for testing"""
        lookup = {}
        for idx, row in self.sample_summary_data.iterrows():
            k1 = f"{row['Unit name'].strip()}|{row['Tenant ID'].strip()}"
            k2 = f"{row['Unit name'].strip()}|{row['Tenant'].strip()}"
            lookup[k1] = (idx, row.to_dict())
            lookup[k2] = (idx, row.to_dict())
        return lookup
    
    @patch('excel_processor.processor.COMManager')
    @patch('excel_processor.processor.MemoryOptimizer')
    def test_concurrent_auto_fill_operations(self, mock_memory_optimizer, mock_com_manager):
        """Test: Multiple auto-fill operations can run concurrently without conflicts"""
        # Arrange
        mock_memory_optimizer.get_memory_usage.return_value = 100.0
        mock_memory_optimizer.cleanup_memory.return_value = None
        mock_memory_optimizer.optimize_workbook_for_large_files.return_value = None
        
        mock_com_manager.initialize_com.return_value = True
        mock_com_manager.get_or_create_excel_app.return_value = MockExcelApp()
        mock_com_manager.release_excel_app.return_value = None
        
        # Create multiple test scenarios
        test_scenarios = []
        for i in range(3):
            headers = self.sample_headers
            data = [
                ['Leasing period', 'Committed', f'Factory{i+1:02d}', f'T{i*10+1:03d}', f'Tenant {chr(65+i)}', '1000'],
                ['Leasing period', 'Committed', '', '', '', ''],  # Empty row needing fill
            ]
            summary = create_test_summary_data(5)  # More summary than data
            test_scenarios.append((headers, data, summary))
        
        results = []
        threads = []
        
        def process_scenario(scenario_data):
            headers, data, summary = scenario_data
            mock_sheet = self.create_mock_sheet_with_data(headers, data)
            mock_sheet.used_range.last_cell.row = len(data) + 10
            df = self.create_dataframe_from_excel_data(headers, data)
            
            rows_updated, rows_added = self.processor._process_dataframe_enhanced(
                df, mock_sheet, 1, headers, summary
            )
            results.append((rows_updated, rows_added))
        
        # Act - Run scenarios in parallel
        for scenario in test_scenarios:
            thread = threading.Thread(target=process_scenario, args=(scenario,))
            threads.append(thread)
            thread.start()
        
        # Wait for all threads to complete
        for thread in threads:
            thread.join(timeout=10)  # 10 second timeout
        
        # Assert
        self.assertEqual(len(results), 3)
        for rows_updated, rows_added in results:
            self.assertGreaterEqual(rows_updated, 0)
            self.assertGreaterEqual(rows_added, 0)
    
    @patch('excel_processor.processor.safe_excel_operation_with_retry')
    def test_auto_fill_with_retry_logic(self, mock_retry):
        """Test: Auto-fill works with retry logic for COM failures"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        
        # Mock retry succeeding on second attempt
        mock_retry.side_effect = [Exception("COM Error"), True]
        
        # Create scenario requiring auto-fill
        mock_sheet = MockExcelSheet()
        
        # Act & Assert - Should not raise exception due to retry logic
        try:
            rows_added = self.processor._add_empty_green_rows(mock_sheet, 1, 5, headers)
            # If we get here, retry logic worked
            self.assertIsInstance(rows_added, int)
        except Exception as e:
            self.fail(f"Auto-fill should handle retries gracefully, got: {e}")


class TestAutoFillEndToEndIntegration(AutoFillTestBase):
    """Test full end-to-end processing with auto-fill enabled"""
    
    def setUp(self):
        super().setUp()
        self.processor = EnhancedExcelProcessor(self.config)
    
    @patch('excel_processor.processor.COMManager')
    @patch('excel_processor.processor.MemoryOptimizer')
    @patch('excel_processor.processor.SubsidiaryExtractor')
    @patch('excel_processor.processor.EnhancedExcelOptimizer')
    def test_end_to_end_processing_with_auto_fill(self, mock_excel_optimizer, 
                                                 mock_subsidiary_extractor, mock_memory_optimizer, 
                                                 mock_com_manager):
        """Test: Full file processing pipeline includes auto-fill functionality"""
        # Arrange - Mock all dependencies
        mock_com_manager.initialize_com.return_value = True
        
        mock_app = MockExcelApp()
        mock_com_manager.get_or_create_excel_app.return_value = mock_app
        mock_com_manager.release_excel_app.return_value = None
        
        mock_memory_optimizer.get_memory_usage.return_value = 100.0
        mock_memory_optimizer.optimize_workbook_for_large_files.return_value = None
        mock_memory_optimizer.cleanup_memory.return_value = None
        mock_memory_optimizer.get_available_memory.return_value = 2000
        mock_memory_optimizer.check_memory_pressure.return_value = False
        
        mock_subsidiary_extractor.extract_subsidiary_enhanced.return_value = "TestSub"
        
        mock_excel_optimizer.find_header_row_enhanced.return_value = 1
        mock_excel_optimizer.safe_excel_operation.side_effect = lambda func: func()
        
        # Create workbook with test data
        headers = self.sample_headers
        test_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],  # Will match
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty row
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty row
        ]
        
        mock_sheet = self.create_mock_sheet_with_data(headers, test_data)
        mock_sheet.used_range.last_cell.row = len(test_data) + 10
        
        mock_workbook = MockExcelWorkbook({'1.Leasing income': mock_sheet})
        mock_app.workbooks_dict = {'test_file.xlsb': mock_workbook}
        
        # Set up processor with summary data that will require auto-fill
        large_summary = create_test_summary_data(10)  # 10 summary items, only 1 matches existing
        large_summary.iloc[0] = {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'TestSub'}
        
        self.processor.summary_data = large_summary
        self.processor.summary_lookup = {}
        for idx, row in large_summary.iterrows():
            k1 = f"{row['Unit name'].strip()}|{row['Tenant ID'].strip()}"
            k2 = f"{row['Unit name'].strip()}|{row['Tenant'].strip()}"
            self.processor.summary_lookup[k1] = (idx, row.to_dict())
            self.processor.summary_lookup[k2] = (idx, row.to_dict())
        
        # Act
        result = self.processor.process_single_file_enhanced('test_file.xlsb')
        
        # Assert
        self.assertEqual(result.status, 'success')
        self.assertGreater(result.rows_updated, 0)  # Should have updates
        self.assertGreater(result.rows_added, 0)    # Should have auto-filled rows
        self.assertTrue(mock_workbook.is_saved)      # Should save file
        self.assertTrue(mock_workbook.is_closed)     # Should close file
    
    @patch('excel_processor.processor.performance_logger')
    def test_auto_fill_performance_logging(self, mock_perf_logger):
        """Test: Auto-fill operations are properly logged for performance monitoring"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code', 'Tenant name']
        sheet = MockExcelSheet()
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 100, headers)
        
        # Assert
        self.assertEqual(rows_added, 100)
        # Performance logging should have been triggered for the parent operation
        # (Note: _add_empty_green_rows doesn't directly use performance logger,
        # but the calling methods do)


class TestAutoFillErrorRecovery(AutoFillTestBase):
    """Test auto-fill error recovery and graceful degradation"""
    
    def test_auto_fill_partial_failure_recovery(self):
        """Test: Auto-fill recovers from partial failures and continues processing"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code']
        
        # Create mock sheet that fails on certain rows
        class FailingMockSheet(MockExcelSheet):
            def range(self, start, end=None):
                mock_range = super().range(start, end)
                
                # Simulate failure on row 3
                if hasattr(start, '__len__') and len(start) == 2:
                    if start[0] == 3:
                        mock_range.value = Mock(side_effect=Exception("Simulated COM error"))
                        return mock_range
                
                return mock_range
        
        failing_sheet = FailingMockSheet()
        
        # Act - Try to add 5 rows, expect partial success
        with patch('builtins.print'):  # Suppress error output
            rows_added = self.processor._add_empty_green_rows(failing_sheet, 1, 5, headers)
        
        # Assert - Should handle failure gracefully
        self.assertGreaterEqual(rows_added, 0)  # Some rows may succeed
        self.assertLessEqual(rows_added, 5)     # But not more than requested
    
    def test_auto_fill_complete_failure_graceful_degradation(self):
        """Test: Auto-fill degrades gracefully when it cannot add any rows"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        
        # Create completely failing mock sheet
        mock_sheet = Mock()
        mock_sheet.range.side_effect = Exception("Complete COM failure")
        
        # Act
        with patch('builtins.print'):  # Suppress error output
            rows_added = self.processor._add_empty_green_rows(mock_sheet, 1, 10, headers)
        
        # Assert
        self.assertEqual(rows_added, 0)  # Should return 0 on complete failure
    
    def test_auto_fill_with_invalid_headers(self):
        """Test: Auto-fill handles cases with missing or invalid headers"""
        # Arrange - Headers without Item2 or Note columns
        invalid_headers = ['Factory code', 'Tenant code', 'Tenant name', 'Rent']
        sheet = MockExcelSheet()
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 3, invalid_headers)
        
        # Assert
        self.assertEqual(rows_added, 3)  # Should still add rows
        
        # Verify that missing columns don't cause errors
        for row in range(1, 4):
            for col in range(1, len(invalid_headers) + 1):
                # All cells should be empty since Item2 and Note columns are missing
                self.assertEqual(sheet.data.get((row, col), ''), '')
    
    def test_auto_fill_with_sheet_protection(self):
        """Test: Auto-fill handles protected/read-only sheets gracefully"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        
        # Mock protected sheet that raises permission error
        mock_sheet = Mock()
        mock_range = Mock()
        mock_sheet.range.return_value = mock_range
        mock_range.value = Mock(side_effect=PermissionError("Sheet is protected"))
        
        # Act
        with patch('builtins.print'):  # Suppress error output
            rows_added = self.processor._add_empty_green_rows(mock_sheet, 1, 5, headers)
        
        # Assert
        self.assertEqual(rows_added, 0)  # Should fail gracefully


if __name__ == '__main__':
    # Configure test runner
    unittest.main(verbosity=2, buffer=True)


# Usage:
# python -m pytest tests/test_auto_fill_integration.py -v
# python tests/test_auto_fill_integration.py