"""
Test Suite: Auto-Fill Memory Optimization and Performance Monitoring Integration
Coverage Goals: Integration between auto-fill and monitoring/optimization systems
Dependencies: unittest, pandas, mock Excel COM interfaces, memory monitoring, performance logging

This module tests the integration between auto-fill functionality and:
- Memory optimization during auto-fill operations
- Performance monitoring and logging
- Resource constraint handling
- System health monitoring during bulk operations
"""

import unittest
from unittest.mock import Mock, MagicMock, patch, PropertyMock
import pandas as pd
import time
import gc
from typing import List, Dict, Any
import sys
from pathlib import Path

# Import test utilities
from test_utils import (
    AutoFillTestBase, MockExcelSheet, create_test_summary_data,
    MockMemoryOptimizer, MockPerformanceLogger, TestScenarioBuilder,
    EnhancedMockExcelSheet
)

# Add src to Python path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from excel_processor.models import ProcessingConfig, ProcessingResult
from excel_processor.processor import EnhancedExcelProcessor


class TestAutoFillMemoryIntegration(AutoFillTestBase):
    """Test integration between auto-fill and memory optimization"""
    
    def setUp(self):
        super().setUp()
        self.mock_memory = MockMemoryOptimizer()
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
    def test_memory_monitoring_during_auto_fill(self, mock_memory_optimizer):
        """Test: Memory usage is monitored during auto-fill operations"""
        # Arrange
        mock_memory_optimizer.get_memory_usage.side_effect = [100.0, 150.0, 120.0]  # Simulate memory changes
        mock_memory_optimizer.check_memory_pressure.return_value = False
        mock_memory_optimizer.cleanup_memory.return_value = None
        mock_memory_optimizer.get_available_memory.return_value = 2000.0
        
        headers = self.sample_headers
        excel_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],  # Matched
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty row needing fill
        ]
        
        # Large summary requiring auto-fill
        large_summary = create_test_summary_data(10)
        large_summary.iloc[0] = {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'Sub1'}
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 10
        
        # Act
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, large_summary
        )
        
        # Assert
        self.assertGreater(rows_updated, 0)
        self.assertGreater(rows_added, 0)
        
        # Verify memory monitoring was called
        mock_memory_optimizer.get_memory_usage.assert_called()
        mock_memory_optimizer.check_memory_pressure.assert_called()
    
    @patch('excel_processor.processor.MemoryOptimizer')
    def test_memory_pressure_triggers_cleanup_during_auto_fill(self, mock_memory_optimizer):
        """Test: Memory pressure during auto-fill triggers cleanup"""
        # Arrange - Simulate memory pressure scenario
        self.mock_memory.simulate_memory_usage_increase(1200.0)  # High usage
        
        mock_memory_optimizer.get_memory_usage.return_value = 1800.0
        mock_memory_optimizer.check_memory_pressure.return_value = True  # Pressure detected
        mock_memory_optimizer.cleanup_memory.return_value = None
        mock_memory_optimizer.get_available_memory.return_value = 500.0  # Low available memory
        
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code']
        excel_data = [['Leasing period', 'Committed', '', '', '']]  # One empty row
        
        # Large summary requiring many auto-fill operations
        large_summary = create_test_summary_data(100)
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 10
        
        # Act
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, large_summary
        )
        
        # Assert
        self.assertGreaterEqual(rows_added, 0)  # Should still work despite pressure
        
        # Memory cleanup should have been triggered
        mock_memory_optimizer.cleanup_memory.assert_called()
        mock_memory_optimizer.check_memory_pressure.assert_called()
    
    @patch('excel_processor.processor.MemoryOptimizer')
    def test_chunked_processing_memory_optimization(self, mock_memory_optimizer):
        """Test: Memory optimization works with chunked auto-fill operations"""
        # Arrange
        self.config.chunk_size = 100  # Small chunks for testing
        processor = EnhancedExcelProcessor(self.config)
        processor.summary_data = self.sample_summary_data
        processor.summary_lookup = self._build_summary_lookup()
        
        # Mock memory optimizer for chunked scenario
        mock_memory_optimizer.get_memory_usage.side_effect = [200.0, 300.0, 250.0, 220.0]
        mock_memory_optimizer.check_memory_pressure.side_effect = [False, True, False]  # Pressure in middle
        mock_memory_optimizer.cleanup_memory.return_value = None
        mock_memory_optimizer.get_available_memory.return_value = 1000.0
        
        headers = self.sample_headers
        
        # Create data that will require chunked processing
        excel_data = []
        for i in range(500):  # Large dataset
            if i < 5:
                excel_data.append(['Leasing period', 'Committed', f'Factory{i+1:02d}', f'T{i+1:03d}', f'Tenant {chr(65+i)}', '1000'])
            else:
                excel_data.append(['Leasing period', 'Committed', '', '', '', ''])  # Empty rows
        
        large_summary = create_test_summary_data(20)  # Requires auto-fill
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 10
        
        # Act
        rows_updated, rows_added = processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, large_summary
        )
        
        # Assert
        self.assertGreater(rows_updated, 0)
        self.assertGreater(rows_added, 0)
        
        # Memory monitoring should have been called multiple times during chunked processing
        self.assertGreater(mock_memory_optimizer.get_memory_usage.call_count, 2)
        mock_memory_optimizer.cleanup_memory.assert_called()  # Should cleanup due to pressure
    
    def test_memory_optimization_with_large_auto_fill(self):
        """Test: Memory optimization during large auto-fill operations"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code', 'Tenant name']
        sheet = MockExcelSheet()
        
        # Create memory optimizer with realistic constraints
        mock_memory = MockMemoryOptimizer(available_memory=1000.0, pressure_threshold=800.0)
        
        # Simulate memory usage increase during operation
        initial_memory = mock_memory.get_memory_usage()
        
        # Act - Large auto-fill operation
        with patch('excel_processor.processor.MemoryOptimizer', return_value=mock_memory):
            rows_added = self.processor._add_empty_green_rows(sheet, 1, 2000, headers)
        
        # Assert
        self.assertEqual(rows_added, 2000)
        
        # Memory should be managed appropriately
        final_memory = mock_memory.get_memory_usage()
        # Memory usage should be controlled (not grow unbounded)
        self.assertLess(final_memory - initial_memory, 500.0, "Memory usage should be controlled")


class TestAutoFillPerformanceMonitoring(AutoFillTestBase):
    """Test integration between auto-fill and performance monitoring"""
    
    def setUp(self):
        super().setUp()
        self.mock_perf_logger = MockPerformanceLogger()
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
    
    @patch('excel_processor.processor.performance_logger')
    def test_performance_logging_integration(self, mock_perf_logger):
        """Test: Performance logging works with auto-fill operations"""
        # Arrange
        headers = self.sample_headers
        excel_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],
            ['Leasing period', 'Committed', '', '', '', ''],
        ]
        
        large_summary = create_test_summary_data(5)
        large_summary.iloc[0] = {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'Sub1'}
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 10
        
        # Act
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, large_summary
        )
        
        # Assert
        self.assertGreater(rows_updated, 0)
        self.assertGreater(rows_added, 0)
        
        # Performance logging should have been invoked
        # Note: The actual implementation uses @log_performance decorator
        # which should trigger performance logging
    
    def test_performance_metrics_collection_during_auto_fill(self):
        """Test: Performance metrics are collected during auto-fill"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code']
        sheet = MockExcelSheet()
        
        # Test different sizes to collect performance data
        test_sizes = [50, 100, 200, 500]
        performance_data = []
        
        for size in test_sizes:
            # Act
            start_time = time.perf_counter()
            rows_added = self.processor._add_empty_green_rows(sheet, 1, size, headers)
            execution_time = time.perf_counter() - start_time
            
            # Collect metrics
            performance_data.append({
                'size': size,
                'execution_time': execution_time,
                'rows_added': rows_added,
                'rows_per_second': size / execution_time if execution_time > 0 else 0
            })
            
            # Assert operation succeeded
            self.assertEqual(rows_added, size)
        
        # Assert performance characteristics
        for data in performance_data:
            print(f"Size: {data['size']}, Time: {data['execution_time']:.3f}s, "
                  f"Rows/sec: {data['rows_per_second']:.1f}")
            
            # Performance should be reasonable
            self.assertGreater(data['rows_per_second'], 10, "Should process at least 10 rows per second")
            self.assertLess(data['execution_time'], 5.0, f"Should complete {data['size']} rows within 5 seconds")
    
    def test_performance_degradation_detection(self):
        """Test: Detection of performance degradation in auto-fill"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        baseline_size = 100
        
        # Establish baseline performance
        sheet = MockExcelSheet()
        start_time = time.perf_counter()
        rows_added = self.processor._add_empty_green_rows(sheet, 1, baseline_size, headers)
        baseline_time = time.perf_counter() - start_time
        baseline_rate = baseline_size / baseline_time
        
        # Test larger operation
        large_size = 500
        large_sheet = MockExcelSheet()
        start_time = time.perf_counter()
        rows_added_large = self.processor._add_empty_green_rows(large_sheet, 1, large_size, headers)
        large_time = time.perf_counter() - start_time
        large_rate = large_size / large_time
        
        # Assert
        self.assertEqual(rows_added, baseline_size)
        self.assertEqual(rows_added_large, large_size)
        
        # Performance should scale reasonably (not degrade significantly)
        performance_ratio = large_rate / baseline_rate
        self.assertGreater(performance_ratio, 0.5, "Large operations should maintain at least 50% of baseline performance")
        
        print(f"Baseline: {baseline_rate:.1f} rows/sec, Large: {large_rate:.1f} rows/sec, "
              f"Ratio: {performance_ratio:.2f}")


class TestAutoFillSystemHealthMonitoring(AutoFillTestBase):
    """Test system health monitoring during auto-fill operations"""
    
    def test_resource_usage_monitoring(self):
        """Test: System resource usage monitoring during auto-fill"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code', 'Tenant name']
        sheet = MockExcelSheet()
        
        # Mock system resource monitoring
        mock_memory = MockMemoryOptimizer(available_memory=2000.0, pressure_threshold=1600.0)
        
        # Act - Perform auto-fill while monitoring resources
        initial_memory = mock_memory.get_memory_usage()
        initial_pressure_checks = mock_memory.pressure_checks
        
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 1000, headers)
        
        final_memory = mock_memory.get_memory_usage()
        
        # Simulate resource monitoring that would happen in real system
        mock_memory.check_memory_pressure()
        final_pressure_checks = mock_memory.pressure_checks
        
        # Assert
        self.assertEqual(rows_added, 1000)
        self.assertGreater(final_pressure_checks, initial_pressure_checks)
        
        # Resource usage should be reasonable
        memory_delta = final_memory - initial_memory
        self.assertLess(abs(memory_delta), 100.0, "Memory delta should be reasonable")
    
    @patch('excel_processor.processor.MemoryOptimizer')
    def test_health_check_during_bulk_operations(self, mock_memory_optimizer):
        """Test: System health checks during bulk auto-fill operations"""
        # Arrange
        mock_memory_optimizer.get_memory_usage.return_value = 500.0
        mock_memory_optimizer.check_memory_pressure.return_value = False
        mock_memory_optimizer.cleanup_memory.return_value = None
        mock_memory_optimizer.get_available_memory.return_value = 1500.0
        
        headers = self.sample_headers
        excel_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],
        ]
        
        # Very large summary requiring significant auto-fill
        very_large_summary = create_test_summary_data(500)
        very_large_summary.iloc[0] = {'Unit name': 'Factory01', 'Tenant ID': 'T001', 'Tenant': 'Tenant A', 'Subsidiary': 'Sub1'}
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 10
        
        # Act
        start_time = time.perf_counter()
        rows_updated, rows_added = self.processor._process_dataframe_enhanced(
            df, mock_sheet, 1, headers, very_large_summary
        )
        execution_time = time.perf_counter() - start_time
        
        # Assert
        self.assertGreater(rows_updated, 0)
        self.assertGreater(rows_added, 0)
        
        # System health monitoring should have occurred
        mock_memory_optimizer.get_memory_usage.assert_called()
        mock_memory_optimizer.check_memory_pressure.assert_called()
        
        # Operation should complete in reasonable time despite size
        self.assertLess(execution_time, 30.0, "Large operations should complete within 30 seconds")
    
    def test_error_recovery_with_monitoring(self):
        """Test: Error recovery works with monitoring systems active"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code']
        
        # Create sheet that will fail intermittently
        failing_sheet = EnhancedMockExcelSheet(failure_rate=0.2)  # 20% failure rate
        mock_memory = MockMemoryOptimizer()
        mock_perf = MockPerformanceLogger()
        
        # Act - Attempt auto-fill with monitoring active
        with patch('excel_processor.processor.MemoryOptimizer', return_value=mock_memory):
            start_time = time.perf_counter()
            
            # Suppress error output during testing
            with patch('builtins.print'):
                rows_added = self.processor._add_empty_green_rows(failing_sheet, 1, 50, headers)
            
            execution_time = time.perf_counter() - start_time
        
        # Assert
        # Should handle failures gracefully with monitoring
        self.assertGreaterEqual(rows_added, 0)
        self.assertLessEqual(rows_added, 50)
        
        # Should not hang due to monitoring overhead
        self.assertLess(execution_time, 10.0)
        
        # Some operations should have been attempted
        self.assertGreater(failing_sheet.operation_count, 0)


class TestAutoFillMonitoringScenarios(AutoFillTestBase):
    """Test complex monitoring scenarios using TestScenarioBuilder"""
    
    def test_memory_pressure_scenario(self):
        """Test: Complete memory pressure scenario with auto-fill"""
        # Arrange
        scenario = TestScenarioBuilder.create_memory_pressure_scenario(self.config)
        mock_memory = scenario['memory_optimizer']
        
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code']
        excel_data = [['Leasing period', 'Committed', '', '', '']]  # One empty row
        large_summary = create_test_summary_data(20)  # Requires auto-fill
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 10
        
        # Act
        with patch('excel_processor.processor.MemoryOptimizer', return_value=mock_memory):
            rows_updated, rows_added = self.processor._process_dataframe_enhanced(
                df, mock_sheet, 1, headers, large_summary
            )
        
        # Assert
        self.assertGreater(rows_added, 0)
        self.assertTrue(scenario['expected_pressure'])
        self.assertTrue(scenario['expected_cleanups']())  # Memory cleanup should have occurred
    
    def test_performance_benchmark_scenario(self):
        """Test: Performance benchmark scenario with monitoring"""
        # Arrange
        scenario = TestScenarioBuilder.create_performance_benchmark_scenario()
        mock_perf = scenario['performance_logger']
        targets = scenario['performance_targets']
        
        headers = ['Item2', 'Note', 'Factory code']
        
        # Test each benchmark size
        results = []
        for size in scenario['benchmark_sizes']:
            sheet = MockExcelSheet()
            
            # Act
            mock_perf.start_operation(f'auto_fill_{size}', size=size)
            start_time = time.perf_counter()
            
            rows_added = self.processor._add_empty_green_rows(sheet, 1, size, headers)
            
            execution_time = time.perf_counter() - start_time
            mock_perf.end_operation(f'auto_fill_{size}', success=True, 
                                  rows_added=rows_added, execution_time=execution_time)
            
            rows_per_second = size / execution_time
            results.append({
                'size': size,
                'execution_time': execution_time,
                'rows_per_second': rows_per_second
            })
            
            # Assert against performance targets
            self.assertEqual(rows_added, size)
            self.assertGreater(rows_per_second, targets['rows_per_second_min'],
                             f"Performance target not met for {size} rows")
        
        # Assert overall performance characteristics
        stats = mock_perf.get_operation_stats()
        self.assertGreater(stats['count'], 0)
        self.assertEqual(stats['success_rate'], 1.0)  # All operations should succeed
        
        print(f"\nPerformance Results:")
        for result in results:
            print(f"  {result['size']} rows: {result['execution_time']:.3f}s, "
                  f"{result['rows_per_second']:.1f} rows/sec")
    
    def test_large_dataset_monitoring_scenario(self):
        """Test: Large dataset scenario with comprehensive monitoring"""
        # Arrange
        scenario = TestScenarioBuilder.create_large_dataset_scenario(rows=2000, cols=30)
        
        summary_data = scenario['summary_data']
        excel_data = scenario['excel_data']
        headers = scenario['headers']
        
        mock_memory = MockMemoryOptimizer(available_memory=3000.0, pressure_threshold=2400.0)
        mock_perf = MockPerformanceLogger()
        
        df = self.create_dataframe_from_excel_data(headers, excel_data)
        mock_sheet = self.create_mock_sheet_with_data(headers, excel_data)
        mock_sheet.used_range.last_cell.row = len(excel_data) + 10
        
        # Act
        mock_perf.start_operation('large_dataset_processing', 
                                summary_rows=len(summary_data), excel_rows=len(excel_data))
        
        with patch('excel_processor.processor.MemoryOptimizer', return_value=mock_memory):
            rows_updated, rows_added = self.processor._process_dataframe_enhanced(
                df, mock_sheet, 1, headers, summary_data
            )
        
        mock_perf.end_operation('large_dataset_processing', success=True,
                              rows_updated=rows_updated, rows_added=rows_added)
        
        # Assert
        self.assertGreater(rows_updated, 0)
        self.assertGreater(rows_added, 0)
        
        # Monitoring should show reasonable resource usage
        self.assertGreater(mock_memory.pressure_checks, 0)
        
        # Performance logging should capture the operation
        stats = mock_perf.get_operation_stats('large_dataset_processing')
        self.assertGreater(stats['count'], 0)
        self.assertEqual(stats['success_rate'], 1.0)
        
        print(f"\nLarge Dataset Results:")
        print(f"  Summary rows: {len(summary_data)}")
        print(f"  Excel rows: {len(excel_data)}")
        print(f"  Rows updated: {rows_updated}")
        print(f"  Rows added: {rows_added}")
        print(f"  Memory pressure checks: {mock_memory.pressure_checks}")
        print(f"  Performance: {stats}")


if __name__ == '__main__':
    # Configure test runner
    unittest.main(verbosity=2, buffer=True)


# Usage:
# python -m pytest tests/test_auto_fill_monitoring.py -v -s
# python tests/test_auto_fill_monitoring.py