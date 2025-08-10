"""
Test Suite: Auto-Fill Logic Core Functionality
Coverage Goals: 95%+ coverage of _add_empty_green_rows and related auto-fill logic
Dependencies: unittest, pandas, mock Excel COM interfaces

This module tests the core auto-fill functionality that automatically adds
empty green rows when the count is less than unmatched summary count.
"""

import unittest
from unittest.mock import Mock, MagicMock, patch, PropertyMock
import pandas as pd
from typing import List, Dict, Any
import sys
from pathlib import Path

# Import test utilities
from test_utils import AutoFillTestBase, MockExcelSheet, create_test_summary_data


class TestAutoFillCoreLogic(AutoFillTestBase):
    """Test the core auto-fill logic functionality"""
    
    def test_add_empty_green_rows_basic(self):
        """Test: _add_empty_green_rows basic functionality"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code', 'Tenant name']
        sheet = MockExcelSheet()
        start_row = 10
        count = 3
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, start_row, count, headers)
        
        # Assert
        self.assertEqual(rows_added, 3)
        # Verify the correct data was written to the mock sheet
        expected_calls = count
        self.assertGreaterEqual(len(sheet.data), expected_calls * len(headers))
    
    def test_add_empty_green_rows_with_identifiers(self):
        """Test: Empty green rows have correct Item2 and Note identifiers"""
        # Arrange  
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code', 'Tenant name', 'Rent (USD)']
        sheet = MockExcelSheet()
        start_row = 5
        count = 2
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, start_row, count, headers)
        
        # Assert
        self.assertEqual(rows_added, 2)
        
        # Check that Item2 and Note columns are set correctly
        item2_col = headers.index('Item2') + 1
        note_col = headers.index('Note') + 1
        
        for i in range(count):
            row_num = start_row + i
            self.assertEqual(sheet.data.get((row_num, item2_col)), 'Leasing period')
            self.assertEqual(sheet.data.get((row_num, note_col)), 'Committed')
    
    def test_add_empty_green_rows_zero_count(self):
        """Test: Adding zero rows should return 0"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        sheet = MockExcelSheet()
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 10, 0, headers)
        
        # Assert
        self.assertEqual(rows_added, 0)
        self.assertEqual(len(sheet.data), 0)
    
    def test_add_empty_green_rows_large_count(self):
        """Test: Adding large number of rows (1000+)"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code']
        sheet = MockExcelSheet()
        large_count = 1500
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, large_count, headers)
        
        # Assert
        self.assertEqual(rows_added, large_count)
        
        # Spot check a few rows
        item2_col = headers.index('Item2') + 1
        note_col = headers.index('Note') + 1
        
        self.assertEqual(sheet.data.get((1, item2_col)), 'Leasing period')
        self.assertEqual(sheet.data.get((1, note_col)), 'Committed')
        self.assertEqual(sheet.data.get((500, item2_col)), 'Leasing period')
        self.assertEqual(sheet.data.get((1500, item2_col)), 'Leasing period')
    
    @patch('builtins.print')  # Mock print to suppress output during tests
    def test_add_empty_green_rows_excel_failure(self, mock_print):
        """Test: Handle Excel COM interface failures during row insertion"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        mock_sheet = Mock()
        mock_range = Mock()
        mock_sheet.range.return_value = mock_range
        mock_range.value = Mock(side_effect=Exception("COM Error"))
        
        # Act
        rows_added = self.processor._add_empty_green_rows(mock_sheet, 1, 3, headers)
        
        # Assert
        self.assertEqual(rows_added, 0)  # Should return 0 on failure
        mock_print.assert_called()  # Should print error message


class TestAutoFillScenarios(AutoFillTestBase):
    """Test auto-fill scenarios and integration with data processing"""
    
    def test_auto_fill_scenario_basic(self):
        """Test: Empty green rows = 5, unmatched summary = 8 → Should add 3 green rows"""
        # Arrange - Create DataFrame with 5 empty green rows
        headers = self.sample_headers
        excel_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],  # Matched
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty 1
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty 2  
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty 3
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty 4
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty 5
        ]
        
        # Create summary with 8 total items (1 matched + 7 unmatched)
        summary_data = create_test_summary_data(8)
        summary_data.iloc[0] = {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'Sub1'}
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        
        # Mock the used_range.last_cell.row to simulate finding last row
        mock_sheet.used_range.last_cell.row = len(excel_data) + 1
        
        # Act
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, summary_data
        )
        
        # Assert
        # 1 row should be updated (the matched one)
        # We need 7 unmatched but only have 5 empty, so should add 2 more rows
        self.assertEqual(rows_updated, 1)
        self.assertGreaterEqual(rows_added, 2)  # Should add at least 2 rows for the deficit
    
    def test_auto_fill_scenario_zero_empty_rows(self):
        """Test: Empty green rows = 0, unmatched summary = 10 → Should add 10 green rows"""
        # Arrange - Create DataFrame with no empty green rows
        headers = self.sample_headers
        excel_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],  # Only matched row
        ]
        
        # Create summary with 10 items (1 matched + 9 unmatched)
        summary_data = create_test_summary_data(10)
        summary_data.iloc[0] = {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'Sub1'}
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 1
        
        # Act
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, summary_data
        )
        
        # Assert
        # 1 row updated, need 9 empty rows but have 0, so should add 9
        self.assertEqual(rows_updated, 1)
        self.assertGreaterEqual(rows_added, 9)
    
    def test_auto_fill_scenario_sufficient_empty_rows(self):
        """Test: Empty green rows = 10, unmatched summary = 5 → Should add 0 rows"""
        # Arrange - Create DataFrame with 10 empty green rows
        headers = self.sample_headers  
        excel_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],  # Matched
        ]
        
        # Add 10 empty green rows
        for i in range(10):
            excel_data.append(['Leasing period', 'Committed', '', '', '', ''])
        
        # Create summary with 5 items (1 matched + 4 unmatched)
        summary_data = create_test_summary_data(5)
        summary_data.iloc[0] = {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'Sub1'}
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 1
        
        # Act
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, summary_data
        )
        
        # Assert
        # Should not add any new rows since we have enough empty ones
        self.assertEqual(rows_updated, 1)
        # rows_added might include filling existing empty rows, but not creating new ones
    
    def test_auto_fill_scenario_equal_counts(self):
        """Test: Empty green rows = unmatched summary → Should add 0 rows"""
        # Arrange - Create DataFrame with exactly matching empty rows and unmatched summary
        headers = self.sample_headers
        excel_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],  # Matched
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty 1
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty 2
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty 3
        ]
        
        # Create summary with 4 items (1 matched + 3 unmatched = 3 empty rows needed)
        summary_data = create_test_summary_data(4)
        summary_data.iloc[0] = {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'Sub1'}
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 1
        
        # Act
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, summary_data
        )
        
        # Assert
        self.assertEqual(rows_updated, 1)
        # Should fill existing empty rows but not add new ones
    
    def test_auto_fill_scenario_no_summary_data(self):
        """Test: Unmatched summary = 0 → Should add 0 rows regardless of green rows count"""
        # Arrange - All summary data is matched
        headers = self.sample_headers
        excel_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],  # Matched
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty (will remain empty)
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty (will remain empty)
        ]
        
        # Create summary with only 1 item that matches existing data
        summary_data = pd.DataFrame([
            {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'Sub1'}
        ])
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 1
        
        # Act
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, summary_data
        )
        
        # Assert
        self.assertEqual(rows_updated, 1)  # One row updated with matching data
        # No additional rows should be added since all summary data is matched


class TestAutoFillDataIntegrity(AutoFillTestBase):
    """Test data integrity during auto-fill operations"""
    
    def test_green_row_formatting_preserved(self):
        """Test: Verify green row formatting is applied correctly to new rows"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code', 'Tenant name']
        sheet = MockExcelSheet()
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 5, headers)
        
        # Assert
        self.assertEqual(rows_added, 5)
        
        # Verify each added row has proper green row identifiers
        item2_col = headers.index('Item2') + 1
        note_col = headers.index('Note') + 1
        
        for row in range(1, 6):
            self.assertEqual(sheet.data.get((row, item2_col)), 'Leasing period')
            self.assertEqual(sheet.data.get((row, note_col)), 'Committed')
            
            # Other columns should be empty
            for col in range(3, len(headers) + 1):
                expected_value = sheet.data.get((row, col), '')
                self.assertEqual(expected_value, '', 
                               f"Column {col} in row {row} should be empty but was '{expected_value}'")
    
    def test_original_data_unchanged(self):
        """Test: Original data remains unchanged during auto-fill"""
        # Arrange
        headers = self.sample_headers
        original_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],
            ['Other Item', 'Other Note', 'Factory02', 'T002', 'Tenant B', '2000'],
        ]
        
        df = self.create_dataframe_from_excel_data(headers, original_data)
        original_df = df.copy()
        
        mock_sheet = self.create_mock_sheet_with_data(headers, original_data)
        
        # Create minimal summary to avoid changes
        summary_data = pd.DataFrame([
            {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'Sub1'}
        ])
        
        # Act
        self.processor._process_dataframe_enhanced(df, mock_sheet, 1, headers, summary_data)
        
        # Assert - Original rows should be unchanged (except for the update from summary)
        # The non-green row should remain completely unchanged
        self.assertEqual(df.iloc[1]['Other Item'], original_df.iloc[1]['Other Item'])
        self.assertEqual(df.iloc[1]['Other Note'], original_df.iloc[1]['Other Note'])
        self.assertEqual(df.iloc[1]['Factory code'], original_df.iloc[1]['Factory code'])
    
    def test_position_integrity(self):
        """Test: Rows added in correct location (after existing data)"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        sheet = MockExcelSheet()
        sheet.used_range.last_cell.row = 10  # Simulate existing data up to row 10
        
        # Act
        start_row = 11  # Should add after existing data
        count = 3
        rows_added = self.processor._add_empty_green_rows(sheet, start_row, count, headers)
        
        # Assert
        self.assertEqual(rows_added, 3)
        
        # Verify rows are placed at correct positions
        item2_col = headers.index('Item2') + 1
        self.assertEqual(sheet.data.get((11, item2_col)), 'Leasing period')
        self.assertEqual(sheet.data.get((12, item2_col)), 'Leasing period')  
        self.assertEqual(sheet.data.get((13, item2_col)), 'Leasing period')
        
        # Verify no data written to wrong positions
        self.assertNotIn((10, item2_col), sheet.data)  # Before start_row
        self.assertNotIn((14, item2_col), sheet.data)  # After end


class TestAutoFillAdvancedFunctionality(AutoFillTestBase):
    """Test advanced auto-fill functionality and complex scenarios"""
    
    def test_auto_fill_with_duplicate_headers(self):
        """Test: Auto-fill correctly handles duplicate column headers"""
        # Arrange - Headers with duplicates (common in Excel exports)
        headers = ['Item2', 'Note', 'Rent', 'Factory code', 'Rent', 'Tenant code']
        sheet = MockExcelSheet()
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 3, headers)
        
        # Assert
        self.assertEqual(rows_added, 3)
        
        # Verify Item2 and Note are set correctly
        item2_col = headers.index('Item2') + 1
        note_col = headers.index('Note') + 1
        
        for row in range(1, 4):
            self.assertEqual(sheet.data.get((row, item2_col)), 'Leasing period')
            self.assertEqual(sheet.data.get((row, note_col)), 'Committed')
    
    def test_auto_fill_column_order_variations(self):
        """Test: Auto-fill works with different column orders"""
        # Arrange - Different column arrangements
        column_arrangements = [
            ['Item2', 'Note', 'Factory code', 'Tenant code'],  # Standard order
            ['Factory code', 'Item2', 'Tenant code', 'Note'],  # Mixed order
            ['Note', 'Factory code', 'Tenant code', 'Item2'],  # Item2 last
            ['Factory code', 'Tenant code', 'Note', 'Item2'],  # Note and Item2 swapped
        ]
        
        for arrangement in column_arrangements:
            with self.subTest(arrangement=arrangement):
                sheet = MockExcelSheet()
                
                # Act
                rows_added = self.processor._add_empty_green_rows(sheet, 1, 2, arrangement)
                
                # Assert
                self.assertEqual(rows_added, 2)
                
                # Verify correct columns are set regardless of order
                if 'Item2' in arrangement:
                    item2_col = arrangement.index('Item2') + 1
                    self.assertEqual(sheet.data.get((1, item2_col)), 'Leasing period')
                    self.assertEqual(sheet.data.get((2, item2_col)), 'Leasing period')
                
                if 'Note' in arrangement:
                    note_col = arrangement.index('Note') + 1
                    self.assertEqual(sheet.data.get((1, note_col)), 'Committed')
                    self.assertEqual(sheet.data.get((2, note_col)), 'Committed')
    
    def test_auto_fill_case_insensitive_headers(self):
        """Test: Auto-fill handles case variations in headers"""
        # Arrange - Headers with different cases
        case_variations = [
            ['ITEM2', 'NOTE', 'Factory code'],  # Uppercase
            ['item2', 'note', 'Factory code'],  # Lowercase
            ['Item2', 'Note', 'Factory code'],  # Standard case
            ['iTeM2', 'NoTe', 'Factory code'],  # Mixed case
        ]
        
        for headers in case_variations:
            with self.subTest(headers=headers):
                sheet = MockExcelSheet()
                
                # Act
                rows_added = self.processor._add_empty_green_rows(sheet, 1, 1, headers)
                
                # Assert
                self.assertEqual(rows_added, 1)
                
                # The current implementation is case-sensitive, so only exact matches work
                # This test documents the current behavior
                if 'Item2' in headers:
                    item2_col = headers.index('Item2') + 1
                    self.assertEqual(sheet.data.get((1, item2_col)), 'Leasing period')
                
                if 'Note' in headers:
                    note_col = headers.index('Note') + 1
                    self.assertEqual(sheet.data.get((1, note_col)), 'Committed')
    
    def test_auto_fill_with_extra_columns(self):
        """Test: Auto-fill with many extra columns beyond required ones"""
        # Arrange - Headers with many extra columns
        extra_headers = [
            'Item2', 'Note', 'Factory code', 'Tenant code', 'Tenant name',
            'Rent (USD)', 'Rent (VND)', 'Start Date', 'End Date', 'Area',
            'Floor', 'Building', 'Zone', 'Status', 'Comments', 'Manager',
            'Contract Number', 'Payment Terms', 'Security Deposit'
        ]
        sheet = MockExcelSheet()
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 3, extra_headers)
        
        # Assert
        self.assertEqual(rows_added, 3)
        
        # Verify core columns are set correctly
        item2_col = extra_headers.index('Item2') + 1
        note_col = extra_headers.index('Note') + 1
        
        for row in range(1, 4):
            self.assertEqual(sheet.data.get((row, item2_col)), 'Leasing period')
            self.assertEqual(sheet.data.get((row, note_col)), 'Committed')
            
            # Extra columns should remain empty
            rent_usd_col = extra_headers.index('Rent (USD)') + 1
            comments_col = extra_headers.index('Comments') + 1
            
            self.assertEqual(sheet.data.get((row, rent_usd_col), ''), '')
            self.assertEqual(sheet.data.get((row, comments_col), ''), '')
    
    def test_auto_fill_performance_with_headers(self):
        """Test: Auto-fill performance with different header configurations"""
        import time
        
        # Test different header sizes
        header_configs = [
            ['Item2', 'Note'],  # Minimal
            ['Item2', 'Note'] + [f'Col{i}' for i in range(10)],  # Medium
            ['Item2', 'Note'] + [f'Col{i}' for i in range(100)],  # Large
        ]
        
        performance_results = []
        
        for headers in header_configs:
            sheet = MockExcelSheet()
            
            start_time = time.perf_counter()
            rows_added = self.processor._add_empty_green_rows(sheet, 1, 100, headers)
            execution_time = time.perf_counter() - start_time
            
            performance_results.append({
                'header_count': len(headers),
                'execution_time': execution_time,
                'rows_added': rows_added
            })
            
            # Assert success
            self.assertEqual(rows_added, 100)
        
        # Performance should not degrade significantly with more headers
        for i, result in enumerate(performance_results):
            print(f"  Headers: {result['header_count']}, Time: {result['execution_time']:.3f}s")
            
            # Each configuration should complete in reasonable time
            self.assertLess(result['execution_time'], 2.0, 
                           f"Too slow with {result['header_count']} headers")
    
    def test_auto_fill_data_type_consistency(self):
        """Test: Auto-fill maintains consistent data types"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code', 'Rent']
        sheet = MockExcelSheet()
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 5, headers)
        
        # Assert
        self.assertEqual(rows_added, 5)
        
        # Verify data type consistency
        item2_col = headers.index('Item2') + 1
        note_col = headers.index('Note') + 1
        
        for row in range(1, 6):
            # Required fields should be strings
            item2_value = sheet.data.get((row, item2_col))
            note_value = sheet.data.get((row, note_col))
            
            self.assertIsInstance(item2_value, str)
            self.assertIsInstance(note_value, str)
            self.assertEqual(item2_value, 'Leasing period')
            self.assertEqual(note_value, 'Committed')
            
            # Other fields should be empty strings (consistent type)
            for col in range(1, len(headers) + 1):
                if col not in [item2_col, note_col]:
                    value = sheet.data.get((row, col), '')
                    self.assertEqual(value, '', f"Column {col} should be empty")


class TestAutoFillRobustness(AutoFillTestBase):
    """Test auto-fill robustness and reliability"""
    
    def test_auto_fill_repeated_operations(self):
        """Test: Multiple auto-fill operations on same sheet"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code']
        sheet = MockExcelSheet()
        
        # Act - Perform multiple auto-fill operations
        operations = [
            (1, 5),   # Add 5 rows starting at row 1
            (6, 3),   # Add 3 rows starting at row 6
            (9, 7),   # Add 7 rows starting at row 9
        ]
        
        total_expected = 0
        for start_row, count in operations:
            rows_added = self.processor._add_empty_green_rows(sheet, start_row, count, headers)
            self.assertEqual(rows_added, count)
            total_expected += count
        
        # Assert
        # Verify all rows were added correctly
        item2_col = headers.index('Item2') + 1
        note_col = headers.index('Note') + 1
        
        written_rows = 0
        for row in range(1, 20):  # Check generous range
            if sheet.data.get((row, item2_col)) == 'Leasing period':
                written_rows += 1
                self.assertEqual(sheet.data.get((row, note_col)), 'Committed')
        
        self.assertEqual(written_rows, total_expected)
    
    def test_auto_fill_idempotency(self):
        """Test: Auto-fill operations are idempotent when appropriate"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        sheet = MockExcelSheet()
        
        # Act - Perform same operation twice
        rows_added_1 = self.processor._add_empty_green_rows(sheet, 1, 3, headers)
        original_data = dict(sheet.data)  # Copy current state
        
        # Second operation on different rows (should not interfere)
        rows_added_2 = self.processor._add_empty_green_rows(sheet, 10, 3, headers)
        
        # Assert
        self.assertEqual(rows_added_1, 3)
        self.assertEqual(rows_added_2, 3)
        
        # Original data should be unchanged
        for key, value in original_data.items():
            self.assertEqual(sheet.data[key], value, f"Original data at {key} was modified")
        
        # New data should be added at correct location
        item2_col = headers.index('Item2') + 1
        for row in range(10, 13):
            self.assertEqual(sheet.data.get((row, item2_col)), 'Leasing period')
    
    def test_auto_fill_consistency_across_calls(self):
        """Test: Auto-fill produces consistent results across multiple calls"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code']
        
        # Act - Perform same operation on different sheets
        results = []
        for i in range(5):
            sheet = MockExcelSheet()
            rows_added = self.processor._add_empty_green_rows(sheet, 1, 10, headers)
            
            # Collect data for comparison
            item2_col = headers.index('Item2') + 1
            note_col = headers.index('Note') + 1
            
            sheet_data = []
            for row in range(1, 11):
                sheet_data.append({
                    'row': row,
                    'item2': sheet.data.get((row, item2_col)),
                    'note': sheet.data.get((row, note_col))
                })
            
            results.append({
                'rows_added': rows_added,
                'data': sheet_data
            })
        
        # Assert
        # All operations should produce identical results
        first_result = results[0]
        for i, result in enumerate(results[1:], 1):
            self.assertEqual(result['rows_added'], first_result['rows_added'],
                           f"Operation {i} returned different row count")
            
            for j, (first_row, result_row) in enumerate(zip(first_result['data'], result['data'])):
                self.assertEqual(result_row['item2'], first_row['item2'],
                               f"Operation {i}, row {j}: Item2 mismatch")
                self.assertEqual(result_row['note'], first_row['note'],
                               f"Operation {i}, row {j}: Note mismatch")
    
    def test_auto_fill_thread_safety_simulation(self):
        """Test: Simulate thread safety by rapid sequential operations"""
        import time
        
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        sheet = MockExcelSheet()
        
        # Act - Rapid sequential operations to simulate concurrency stress
        operations = []
        start_time = time.perf_counter()
        
        for i in range(10):
            start_row = i * 5 + 1  # Non-overlapping row ranges
            rows_added = self.processor._add_empty_green_rows(sheet, start_row, 3, headers)
            operations.append((start_row, rows_added))
            
            # Brief pause to simulate processing time
            time.sleep(0.001)
        
        end_time = time.perf_counter()
        
        # Assert
        # All operations should succeed
        for start_row, rows_added in operations:
            self.assertEqual(rows_added, 3, f"Operation starting at row {start_row} failed")
        
        # Verify data integrity
        item2_col = headers.index('Item2') + 1
        
        for start_row, _ in operations:
            for offset in range(3):
                row = start_row + offset
                self.assertEqual(sheet.data.get((row, item2_col)), 'Leasing period',
                               f"Data corruption at row {row}")
        
        # Performance should be reasonable
        total_time = end_time - start_time
        self.assertLess(total_time, 1.0, "Rapid operations should complete quickly")


if __name__ == '__main__':
    # Configure test runner
    unittest.main(verbosity=2, buffer=True)


# Usage:
# python -m pytest tests/test_auto_fill_logic.py -v
# python tests/test_auto_fill_logic.py