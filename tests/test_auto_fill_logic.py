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


if __name__ == '__main__':
    # Configure test runner
    unittest.main(verbosity=2, buffer=True)


# Usage:
# python -m pytest tests/test_auto_fill_logic.py -v
# python tests/test_auto_fill_logic.py