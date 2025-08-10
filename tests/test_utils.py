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


class MockCOMManager:
    """Mock COM Manager for testing Excel COM operations"""
    
    def __init__(self, success_rate: float = 1.0):
        self.success_rate = success_rate
        self.excel_apps = {}
        self.com_initialized = False
        self.operation_count = 0
    
    def initialize_com(self) -> bool:
        """Mock COM initialization"""
        self.operation_count += 1
        if self._should_succeed():
            self.com_initialized = True
            return True
        return False
    
    def get_or_create_excel_app(self, app_id: str) -> Optional[MockExcelApp]:
        """Mock Excel application creation/retrieval"""
        self.operation_count += 1
        if not self._should_succeed():
            return None
        
        if app_id not in self.excel_apps:
            self.excel_apps[app_id] = MockExcelApp()
        return self.excel_apps[app_id]
    
    def release_excel_app(self, app_id: str) -> bool:
        """Mock Excel application release"""
        self.operation_count += 1
        if app_id in self.excel_apps and self._should_succeed():
            self.excel_apps[app_id].quit()
            del self.excel_apps[app_id]
            return True
        return False
    
    def cleanup_all_apps(self):
        """Mock cleanup of all Excel applications"""
        self.operation_count += 1
        if self._should_succeed():
            for app in self.excel_apps.values():
                app.quit()
            self.excel_apps.clear()
    
    def _should_succeed(self) -> bool:
        """Determine if operation should succeed based on success rate"""
        import random
        return random.random() < self.success_rate


class MockMemoryOptimizer:
    """Mock Memory Optimizer for testing memory-related operations"""
    
    def __init__(self, available_memory: float = 2000.0, pressure_threshold: float = 1500.0):
        self.available_memory = available_memory
        self.pressure_threshold = pressure_threshold
        self.current_usage = 500.0
        self.cleanup_count = 0
        self.pressure_checks = 0
    
    def get_memory_usage(self) -> float:
        """Mock current memory usage"""
        return self.current_usage
    
    def get_available_memory(self) -> float:
        """Mock available memory"""
        return self.available_memory
    
    def check_memory_pressure(self) -> bool:
        """Mock memory pressure check"""
        self.pressure_checks += 1
        return self.current_usage > self.pressure_threshold
    
    def cleanup_memory(self) -> None:
        """Mock memory cleanup"""
        self.cleanup_count += 1
        # Simulate memory cleanup reducing usage
        self.current_usage = max(100.0, self.current_usage * 0.7)
    
    def optimize_workbook_for_large_files(self, workbook) -> None:
        """Mock workbook optimization"""
        pass
    
    def simulate_memory_usage_increase(self, amount: float):
        """Simulate increase in memory usage for testing"""
        self.current_usage += amount
        if self.current_usage > self.available_memory:
            self.current_usage = self.available_memory
    
    def reset_memory_state(self):
        """Reset to initial state for testing"""
        self.current_usage = 500.0
        self.cleanup_count = 0
        self.pressure_checks = 0


class MockPerformanceLogger:
    """Mock Performance Logger for testing performance monitoring"""
    
    def __init__(self):
        self.operations = []
        self.active_operations = {}
    
    def log_performance(self, operation_name: str):
        """Mock performance logging decorator"""
        def decorator(func):
            def wrapper(*args, **kwargs):
                import time
                start_time = time.perf_counter()
                
                try:
                    result = func(*args, **kwargs)
                    success = True
                    error = None
                except Exception as e:
                    result = None
                    success = False
                    error = str(e)
                    raise
                finally:
                    end_time = time.perf_counter()
                    execution_time = end_time - start_time
                    
                    self.operations.append({
                        'operation': operation_name,
                        'start_time': start_time,
                        'end_time': end_time,
                        'execution_time': execution_time,
                        'success': success,
                        'error': error
                    })
                
                return result
            return wrapper
        return decorator
    
    def start_operation(self, operation_name: str, **metadata):
        """Mock start of performance tracking"""
        import time
        self.active_operations[operation_name] = {
            'start_time': time.perf_counter(),
            'metadata': metadata
        }
    
    def end_operation(self, operation_name: str, success: bool = True, **result_data):
        """Mock end of performance tracking"""
        import time
        if operation_name in self.active_operations:
            start_info = self.active_operations.pop(operation_name)
            end_time = time.perf_counter()
            
            self.operations.append({
                'operation': operation_name,
                'start_time': start_info['start_time'],
                'end_time': end_time,
                'execution_time': end_time - start_info['start_time'],
                'success': success,
                'metadata': start_info['metadata'],
                'result_data': result_data
            })
    
    def get_operation_stats(self, operation_name: str = None) -> Dict[str, Any]:
        """Get statistics for logged operations"""
        if operation_name:
            ops = [op for op in self.operations if op['operation'] == operation_name]
        else:
            ops = self.operations
        
        if not ops:
            return {}
        
        execution_times = [op['execution_time'] for op in ops]
        return {
            'count': len(ops),
            'total_time': sum(execution_times),
            'avg_time': sum(execution_times) / len(execution_times),
            'min_time': min(execution_times),
            'max_time': max(execution_times),
            'success_rate': sum(1 for op in ops if op['success']) / len(ops)
        }


class MockRetryUtils:
    """Mock retry utilities for testing retry logic"""
    
    def __init__(self, max_attempts: int = 3, delay: float = 0.1):
        self.max_attempts = max_attempts
        self.delay = delay
        self.attempt_counts = {}
    
    def safe_excel_operation_with_retry(self, operation_func, operation_name: str, 
                                      max_attempts: Optional[int] = None):
        """Mock retry logic for Excel operations"""
        max_attempts = max_attempts or self.max_attempts
        operation_key = f"{operation_name}_{id(operation_func)}"
        
        if operation_key not in self.attempt_counts:
            self.attempt_counts[operation_key] = 0
        
        for attempt in range(max_attempts):
            self.attempt_counts[operation_key] += 1
            
            try:
                return operation_func()
            except Exception as e:
                if attempt == max_attempts - 1:
                    # Last attempt, re-raise the exception
                    raise
                
                # Simulate delay between retries
                import time
                time.sleep(self.delay)
        
        raise RuntimeError(f"All {max_attempts} attempts failed for {operation_name}")
    
    def get_attempt_count(self, operation_name: str, operation_func) -> int:
        """Get the number of attempts made for a specific operation"""
        operation_key = f"{operation_name}_{id(operation_func)}"
        return self.attempt_counts.get(operation_key, 0)
    
    def reset_counts(self):
        """Reset attempt counts for testing"""
        self.attempt_counts.clear()


class EnhancedMockExcelSheet(MockExcelSheet):
    """Enhanced mock Excel sheet with more realistic COM behaviors"""
    
    def __init__(self, used_range_rows: int = 100, used_range_cols: int = 20,
                 failure_rate: float = 0.0, slow_operations: bool = False):
        super().__init__(used_range_rows, used_range_cols)
        self.failure_rate = failure_rate
        self.slow_operations = slow_operations
        self.operation_count = 0
        self.last_error = None
    
    def range(self, start, end=None):
        """Enhanced range method with failure simulation"""
        self.operation_count += 1
        
        # Simulate operation failures
        if self._should_fail():
            error_types = [
                Exception("COM Error: The object invoked has disconnected from its clients"),
                MemoryError("Insufficient memory for operation"),
                TimeoutError("Operation timed out"),
                PermissionError("Access denied - sheet may be protected"),
                RuntimeError("Excel application is busy")
            ]
            
            import random
            error = random.choice(error_types)
            self.last_error = error
            raise error
        
        # Simulate slow operations
        if self.slow_operations:
            import time
            time.sleep(0.01)  # 10ms delay
        
        return super().range(start, end)
    
    def _should_fail(self) -> bool:
        """Determine if operation should fail based on failure rate"""
        import random
        return random.random() < self.failure_rate
    
    def set_failure_rate(self, rate: float):
        """Adjust failure rate for testing different scenarios"""
        self.failure_rate = max(0.0, min(1.0, rate))
    
    def enable_slow_operations(self, enabled: bool = True):
        """Enable/disable slow operation simulation"""
        self.slow_operations = enabled
    
    def get_operation_stats(self) -> Dict[str, Any]:
        """Get statistics about operations performed"""
        return {
            'total_operations': self.operation_count,
            'data_cells_written': len(self.data),
            'last_error': str(self.last_error) if self.last_error else None,
            'failure_rate': self.failure_rate
        }


class TestScenarioBuilder:
    """Helper class to build complex test scenarios"""
    
    @staticmethod
    def create_memory_pressure_scenario(processor_config: ProcessingConfig) -> Dict[str, Any]:
        """Create a scenario that simulates memory pressure"""
        mock_memory = MockMemoryOptimizer(available_memory=500.0, pressure_threshold=300.0)
        mock_memory.simulate_memory_usage_increase(400.0)  # Start near pressure
        
        return {
            'memory_optimizer': mock_memory,
            'config': processor_config,
            'expected_pressure': True,
            'expected_cleanups': lambda: mock_memory.cleanup_count > 0
        }
    
    @staticmethod
    def create_com_failure_scenario(failure_rate: float = 0.3) -> Dict[str, Any]:
        """Create a scenario with intermittent COM failures"""
        mock_com = MockCOMManager(success_rate=1.0 - failure_rate)
        mock_sheet = EnhancedMockExcelSheet(failure_rate=failure_rate)
        
        return {
            'com_manager': mock_com,
            'sheet': mock_sheet,
            'expected_failures': True,
            'failure_rate': failure_rate
        }
    
    @staticmethod
    def create_large_dataset_scenario(rows: int = 5000, cols: int = 50) -> Dict[str, Any]:
        """Create a scenario with large dataset processing"""
        large_summary = create_large_test_data(rows)
        headers = ['Item2', 'Note'] + [f'Col_{i}' for i in range(cols - 2)]
        
        # Create Excel data with mixed content
        excel_data = []
        for i in range(min(1000, rows // 5)):  # Subset for Excel sheet
            if i % 10 < 2:  # 20% are empty green rows
                row = ['Leasing period', 'Committed'] + [''] * (cols - 2)
            elif i % 10 < 4:  # 20% are filled green rows
                row = ['Leasing period', 'Committed', f'Factory{i:02d}', f'T{i:03d}'] + [''] * (cols - 4)
            else:  # 60% are other data
                row = ['Other Item', 'Other Note'] + [f'Data_{i}_{j}' for j in range(cols - 2)]
            excel_data.append(row)
        
        return {
            'summary_data': large_summary,
            'excel_data': excel_data,
            'headers': headers,
            'expected_auto_fill': True,
            'size_category': 'large' if rows > 1000 else 'medium'
        }
    
    @staticmethod
    def create_performance_benchmark_scenario() -> Dict[str, Any]:
        """Create a scenario for performance benchmarking"""
        mock_perf = MockPerformanceLogger()
        
        # Define performance targets
        targets = {
            'rows_per_second_min': 100,
            'memory_per_row_max_mb': 0.01,
            'max_execution_time_1000_rows': 10.0
        }
        
        return {
            'performance_logger': mock_perf,
            'performance_targets': targets,
            'benchmark_sizes': [100, 500, 1000, 2000],
            'expected_scaling': 'linear'
        }