"""
Test Suite: Auto-Fill Performance Benchmarks
Coverage Goals: Performance testing and benchmarking of auto-fill operations
Dependencies: unittest, pandas, time, memory profiling, mock Excel COM interfaces

This module tests performance aspects of auto-fill functionality:
- Large scale operations (1000+ rows)
- Memory usage during bulk operations
- Speed benchmarks for row addition
- Concurrent auto-fill performance
- Memory pressure handling
"""

import unittest
from unittest.mock import Mock, MagicMock, patch, PropertyMock
import pandas as pd
import time
import threading
from typing import List, Dict, Any, Tuple
import sys
from pathlib import Path
import gc
import tracemalloc

# Import test utilities
from test_utils import AutoFillTestBase, MockExcelSheet, create_test_summary_data, create_large_test_data

# Add src to Python path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from excel_processor.models import ProcessingConfig, ProcessingResult
from excel_processor.processor import EnhancedExcelProcessor


class TestAutoFillPerformanceBenchmarks(AutoFillTestBase):
    """Performance benchmarks for auto-fill operations"""
    
    def setUp(self):
        super().setUp()
        # Configure for performance testing
        self.config.chunk_size = 5000
        self.config.enable_memory_monitoring = True
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
    
    def test_large_scale_auto_fill_performance(self):
        """Test: Performance of adding 1000+ green rows"""
        # Arrange
        headers = self.sample_headers
        sheet = MockExcelSheet()
        large_count = 1500
        
        # Start performance measurement
        start_time = time.perf_counter()
        start_memory = self._get_memory_usage()
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, large_count, headers)
        
        # End performance measurement
        end_time = time.perf_counter()
        end_memory = self._get_memory_usage()
        
        execution_time = end_time - start_time
        memory_delta = end_memory - start_memory
        
        # Assert
        self.assertEqual(rows_added, large_count)
        
        # Performance assertions
        self.assertLess(execution_time, 5.0, f"Adding {large_count} rows took {execution_time:.2f}s, should be < 5s")
        self.assertLess(memory_delta, 50.0, f"Memory usage increased by {memory_delta:.1f}MB, should be < 50MB")
        
        # Calculate performance metrics
        rows_per_second = large_count / execution_time
        memory_per_row_kb = (memory_delta * 1024) / large_count
        
        print(f"\nPerformance Metrics for {large_count} rows:")
        print(f"  Execution time: {execution_time:.3f}s")
        print(f"  Rows per second: {rows_per_second:.1f}")
        print(f"  Memory delta: {memory_delta:.1f}MB")
        print(f"  Memory per row: {memory_per_row_kb:.2f}KB")
        
        # Performance targets
        self.assertGreater(rows_per_second, 100, "Should process at least 100 rows per second")
        self.assertLess(memory_per_row_kb, 50, "Should use less than 50KB per row")
    
    def test_memory_usage_during_bulk_operations(self):
        """Test: Memory usage patterns during bulk row addition"""
        # Arrange
        headers = self.sample_headers
        sheet = MockExcelSheet()
        
        # Test different bulk sizes
        test_sizes = [100, 500, 1000, 2000]
        memory_measurements = []
        
        for size in test_sizes:
            # Force garbage collection before measurement
            gc.collect()
            
            # Start memory tracking
            tracemalloc.start()
            start_memory = self._get_memory_usage()
            
            # Act
            rows_added = self.processor._add_empty_green_rows(sheet, 1, size, headers)
            
            # End memory tracking
            end_memory = self._get_memory_usage()
            current, peak = tracemalloc.get_traced_memory()
            tracemalloc.stop()
            
            memory_delta = end_memory - start_memory
            peak_mb = peak / (1024 * 1024)
            
            memory_measurements.append({
                'size': size,
                'rows_added': rows_added,
                'memory_delta': memory_delta,
                'peak_memory': peak_mb
            })
            
            # Assert basic success
            self.assertEqual(rows_added, size)
        
        # Analyze memory scaling
        print(f"\nMemory Usage Analysis:")
        for measurement in memory_measurements:
            print(f"  {measurement['size']} rows: {measurement['memory_delta']:.1f}MB delta, {measurement['peak_memory']:.1f}MB peak")
        
        # Assert memory usage is reasonable and scales appropriately
        for measurement in memory_measurements:
            memory_per_row_kb = (measurement['memory_delta'] * 1024) / measurement['size']
            self.assertLess(memory_per_row_kb, 100, f"Memory per row should be < 100KB for {measurement['size']} rows")
            self.assertLess(measurement['peak_memory'], 100, f"Peak memory should be < 100MB for {measurement['size']} rows")
    
    def test_batch_vs_individual_row_performance(self):
        """Test: Performance comparison between batch and individual row operations"""
        headers = self.sample_headers
        row_count = 500
        
        # Test individual row approach (simulated)
        individual_sheet = MockExcelSheet()
        start_time = time.perf_counter()
        
        for i in range(row_count):
            # Simulate individual row operations
            self.processor._add_empty_green_rows(individual_sheet, i + 1, 1, headers)
        
        individual_time = time.perf_counter() - start_time
        
        # Test batch approach
        batch_sheet = MockExcelSheet()
        start_time = time.perf_counter()
        
        self.processor._add_empty_green_rows(batch_sheet, 1, row_count, headers)
        
        batch_time = time.perf_counter() - start_time
        
        # Assert batch is significantly faster
        speedup = individual_time / batch_time
        print(f"\nBatch vs Individual Performance:")
        print(f"  Individual: {individual_time:.3f}s")
        print(f"  Batch: {batch_time:.3f}s")
        print(f"  Speedup: {speedup:.1f}x")
        
        self.assertGreater(speedup, 2.0, "Batch processing should be at least 2x faster")
        self.assertLess(batch_time, individual_time, "Batch should be faster than individual operations")
    
    def test_concurrent_auto_fill_performance(self):
        """Test: Performance of concurrent auto-fill operations"""
        headers = self.sample_headers
        thread_count = 4
        rows_per_thread = 250
        
        # Prepare test data
        threads = []
        results = []
        
        def worker_function(thread_id: int):
            sheet = MockExcelSheet()
            start_time = time.perf_counter()
            
            rows_added = self.processor._add_empty_green_rows(
                sheet, 1, rows_per_thread, headers
            )
            
            execution_time = time.perf_counter() - start_time
            results.append({
                'thread_id': thread_id,
                'rows_added': rows_added,
                'execution_time': execution_time
            })
        
        # Start concurrent operations
        overall_start = time.perf_counter()
        
        for i in range(thread_count):
            thread = threading.Thread(target=worker_function, args=(i,))
            threads.append(thread)
            thread.start()
        
        # Wait for completion
        for thread in threads:
            thread.join(timeout=30)  # 30 second timeout
        
        overall_time = time.perf_counter() - overall_start
        
        # Assert all threads completed successfully
        self.assertEqual(len(results), thread_count)
        
        total_rows = sum(r['rows_added'] for r in results)
        max_thread_time = max(r['execution_time'] for r in results)
        avg_thread_time = sum(r['execution_time'] for r in results) / len(results)
        
        print(f"\nConcurrent Performance Results:")
        print(f"  Threads: {thread_count}")
        print(f"  Total rows: {total_rows}")
        print(f"  Overall time: {overall_time:.3f}s")
        print(f"  Max thread time: {max_thread_time:.3f}s")
        print(f"  Avg thread time: {avg_thread_time:.3f}s")
        print(f"  Effective rows/sec: {total_rows / overall_time:.1f}")
        
        # Performance assertions
        for result in results:
            self.assertEqual(result['rows_added'], rows_per_thread)
            self.assertLess(result['execution_time'], 10.0, f"Thread {result['thread_id']} took too long")
        
        # Concurrent execution should be faster than sequential
        sequential_estimate = avg_thread_time * thread_count
        concurrent_efficiency = (sequential_estimate / overall_time) * 100
        
        print(f"  Concurrent efficiency: {concurrent_efficiency:.1f}%")
        self.assertGreater(concurrent_efficiency, 150, "Concurrent execution should be at least 50% more efficient")
    
    def _get_memory_usage(self) -> float:
        """Get current memory usage in MB"""
        try:
            import psutil
            process = psutil.Process()
            return process.memory_info().rss / (1024 * 1024)  # Convert to MB
        except ImportError:
            # Fallback if psutil not available
            import sys
            return sys.getsizeof(gc.get_objects()) / (1024 * 1024)


class TestAutoFillMemoryPressureHandling(AutoFillTestBase):
    """Test auto-fill behavior under memory pressure conditions"""
    
    @patch('excel_processor.processor.MemoryOptimizer')
    def test_auto_fill_under_memory_pressure(self, mock_memory_optimizer):
        """Test: Auto-fill handles memory pressure gracefully"""
        # Arrange
        headers = self.sample_headers
        
        # Simulate progressive memory pressure
        memory_pressure_sequence = [False, False, True, True, False]
        mock_memory_optimizer.check_memory_pressure.side_effect = memory_pressure_sequence
        mock_memory_optimizer.cleanup_memory.return_value = None
        mock_memory_optimizer.get_memory_usage.return_value = 1000.0
        
        # Create large dataset that would require auto-fill
        excel_data = [
            ['Leasing period', 'Committed', 'Factory01', 'T001', 'Tenant A', '1000'],  # Matched
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty 1
            ['Leasing period', 'Committed', '', '', '', ''],  # Empty 2
        ]
        
        # Create summary requiring many more rows
        large_summary = create_test_summary_data(20)
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
        self.assertGreater(rows_updated, 0)
        self.assertGreater(rows_added, 0)
        
        # Should have triggered memory cleanup
        mock_memory_optimizer.cleanup_memory.assert_called()
        mock_memory_optimizer.check_memory_pressure.assert_called()
        
        # Should complete in reasonable time even with memory pressure
        self.assertLess(execution_time, 10.0, "Should handle memory pressure within 10 seconds")
    
    @patch('excel_processor.processor.MemoryOptimizer')
    def test_auto_fill_memory_optimization_integration(self, mock_memory_optimizer):
        """Test: Auto-fill integrates properly with memory optimization"""
        # Arrange
        headers = ['Item2', 'Note', 'Factory code', 'Tenant code', 'Tenant name']
        sheet = MockExcelSheet()
        
        # Configure memory optimizer mocks
        mock_memory_optimizer.get_available_memory.return_value = 500  # Limited memory
        mock_memory_optimizer.check_memory_pressure.return_value = True  # Pressure detected
        mock_memory_optimizer.cleanup_memory.return_value = None
        
        # Act
        rows_added = self.processor._add_empty_green_rows(sheet, 1, 1000, headers)
        
        # Assert
        self.assertEqual(rows_added, 1000)  # Should still succeed
        
        # Memory optimizer should not have been called directly by _add_empty_green_rows
        # but may be called by the parent processing methods
    
    def test_large_dataset_memory_efficiency(self):
        """Test: Memory efficiency with very large datasets"""
        # Arrange
        headers = self.sample_headers
        very_large_count = 5000
        
        # Monitor memory before operation
        gc.collect()  # Clean up before measurement
        initial_memory = self._get_memory_usage()
        
        # Create large mock sheet
        sheet = MockExcelSheet()
        
        # Act
        start_time = time.perf_counter()
        rows_added = self.processor._add_empty_green_rows(sheet, 1, very_large_count, headers)
        execution_time = time.perf_counter() - start_time
        
        # Monitor memory after operation
        final_memory = self._get_memory_usage()
        memory_delta = final_memory - initial_memory
        
        # Assert
        self.assertEqual(rows_added, very_large_count)
        
        # Memory efficiency targets
        memory_per_row_kb = (memory_delta * 1024) / very_large_count
        self.assertLess(memory_per_row_kb, 10, f"Should use < 10KB per row, used {memory_per_row_kb:.2f}KB")
        self.assertLess(execution_time, 30.0, f"Should complete within 30s, took {execution_time:.2f}s")
        
        print(f"\nLarge Dataset Memory Efficiency:")
        print(f"  Rows processed: {very_large_count}")
        print(f"  Execution time: {execution_time:.3f}s")
        print(f"  Memory delta: {memory_delta:.1f}MB")
        print(f"  Memory per row: {memory_per_row_kb:.3f}KB")
        print(f"  Rows per second: {very_large_count / execution_time:.1f}")
    
    def _get_memory_usage(self) -> float:
        """Get current memory usage in MB"""
        try:
            import psutil
            process = psutil.Process()
            return process.memory_info().rss / (1024 * 1024)
        except ImportError:
            # Fallback estimation
            return 0.0


class TestAutoFillPerformanceRegression(AutoFillTestBase):
    """Regression tests to ensure performance doesn't degrade over time"""
    
    def setUp(self):
        super().setUp()
        # Performance baseline targets (adjust based on your environment)
        self.performance_targets = {
            'rows_per_second_min': 50,  # Minimum rows per second
            'memory_per_row_max_kb': 20,  # Maximum KB per row
            'max_execution_time_1000_rows': 20.0,  # Maximum seconds for 1000 rows
            'concurrent_efficiency_min': 120  # Minimum concurrent efficiency %
        }
    
    def test_performance_regression_baseline(self):
        """Test: Ensure auto-fill performance meets baseline requirements"""
        # Arrange
        headers = self.sample_headers
        sheet = MockExcelSheet()
        baseline_row_count = 1000
        
        # Measure performance
        start_time = time.perf_counter()
        start_memory = self._get_memory_usage()
        
        rows_added = self.processor._add_empty_green_rows(sheet, 1, baseline_row_count, headers)
        
        end_time = time.perf_counter()
        end_memory = self._get_memory_usage()
        
        # Calculate metrics
        execution_time = end_time - start_time
        memory_delta = max(0, end_memory - start_memory)  # Protect against negative values
        rows_per_second = baseline_row_count / execution_time
        memory_per_row_kb = (memory_delta * 1024) / baseline_row_count if memory_delta > 0 else 0
        
        # Assert against performance targets
        self.assertEqual(rows_added, baseline_row_count)
        self.assertGreater(rows_per_second, self.performance_targets['rows_per_second_min'],
                          f"Performance regression: {rows_per_second:.1f} < {self.performance_targets['rows_per_second_min']} rows/sec")
        self.assertLess(execution_time, self.performance_targets['max_execution_time_1000_rows'],
                       f"Performance regression: {execution_time:.2f}s > {self.performance_targets['max_execution_time_1000_rows']}s")
        
        if memory_delta > 0:  # Only check if we have meaningful memory measurements
            self.assertLess(memory_per_row_kb, self.performance_targets['memory_per_row_max_kb'],
                           f"Memory regression: {memory_per_row_kb:.2f}KB > {self.performance_targets['memory_per_row_max_kb']}KB per row")
        
        print(f"\nPerformance Baseline Results:")
        print(f"  Rows per second: {rows_per_second:.1f} (target: >{self.performance_targets['rows_per_second_min']})")
        print(f"  Execution time: {execution_time:.3f}s (target: <{self.performance_targets['max_execution_time_1000_rows']}s)")
        print(f"  Memory per row: {memory_per_row_kb:.3f}KB (target: <{self.performance_targets['memory_per_row_max_kb']}KB)")
    
    def test_performance_with_different_row_sizes(self):
        """Test: Performance characteristics across different row sizes"""
        headers = self.sample_headers
        test_sizes = [10, 100, 500, 1000, 2000]
        results = []
        
        for size in test_sizes:
            sheet = MockExcelSheet()
            
            start_time = time.perf_counter()
            rows_added = self.processor._add_empty_green_rows(sheet, 1, size, headers)
            execution_time = time.perf_counter() - start_time
            
            rows_per_second = size / execution_time
            results.append({
                'size': size,
                'execution_time': execution_time,
                'rows_per_second': rows_per_second
            })
            
            # Each size should meet minimum performance
            self.assertEqual(rows_added, size)
            self.assertGreater(rows_per_second, self.performance_targets['rows_per_second_min'] * 0.5,
                              f"Size {size}: {rows_per_second:.1f} rows/sec too slow")
        
        print(f"\nPerformance Scaling Analysis:")
        for result in results:
            print(f"  {result['size']:4d} rows: {result['execution_time']:.3f}s, {result['rows_per_second']:6.1f} rows/sec")
        
        # Assert reasonable scaling (larger operations should maintain efficiency)
        large_ops = [r for r in results if r['size'] >= 500]
        for result in large_ops:
            self.assertGreater(result['rows_per_second'], self.performance_targets['rows_per_second_min'] * 0.8,
                              f"Large operations should maintain at least 80% of target performance")
    
    def _get_memory_usage(self) -> float:
        """Get current memory usage in MB"""
        try:
            import psutil
            process = psutil.Process()
            return process.memory_info().rss / (1024 * 1024)
        except ImportError:
            return 0.0  # Return 0 if we can't measure


if __name__ == '__main__':
    # Configure test runner for performance tests
    unittest.main(verbosity=2, buffer=True)


# Usage:
# python -m pytest tests/test_auto_fill_performance.py -v -s  # -s to see print statements
# python tests/test_auto_fill_performance.py