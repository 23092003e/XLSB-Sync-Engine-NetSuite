# excel_processor/memory_optimizer.py
import gc
import psutil
import xlwings as xw
from typing import Optional

class MemoryOptimizer:
    @staticmethod
    def get_memory_usage() -> float:
        """Get current memory usage in MB"""
        process = psutil.Process()
        return process.memory_info().rss / 1024 / 1024
    
    @staticmethod
    def get_available_memory() -> float:
        """Get available system memory in MB"""
        return psutil.virtual_memory().available / 1024 / 1024
    
    @staticmethod
    def get_memory_percent() -> float:
        """Get current memory usage as percentage"""
        return psutil.virtual_memory().percent
    
    @staticmethod
    def check_memory_pressure(threshold_percent: float = 70.0) -> bool:
        """Check if memory usage exceeds threshold"""
        current_usage = MemoryOptimizer.get_memory_percent()
        return current_usage > threshold_percent
    
    @staticmethod
    def optimize_workbook_for_large_files(wb: xw.Book) -> None:
        """Apply optimizations for large XLSB files"""
        try:
            # Disable automatic recalculation
            wb.app.calculation = 'manual'
            
            # Turn off screen updating
            wb.app.screen_updating = False
            
            # Disable events
            wb.app.enable_events = False
            
            # Turn off alerts
            wb.app.display_alerts = False
            
            print("   ⚡ Applied large file optimizations")
        except Exception as e:
            print(f"   ⚠️ Memory optimization warning: {e}")
    
    @staticmethod
    def cleanup_memory():
        """Force garbage collection and memory cleanup"""
        import gc
        gc.collect()
        
    @staticmethod
    def monitor_memory_usage(operation_name: str):
        """Decorator to monitor memory usage of operations"""
        def decorator(func):
            def wrapper(*args, **kwargs):
                start_memory = MemoryOptimizer.get_memory_usage()
                start_percent = MemoryOptimizer.get_memory_percent()
                
                result = func(*args, **kwargs)
                
                end_memory = MemoryOptimizer.get_memory_usage()
                end_percent = MemoryOptimizer.get_memory_percent()
                memory_delta = end_memory - start_memory
                
                print(f"   📊 {operation_name}: {memory_delta:+.1f}MB (now {end_memory:.1f}MB, {end_percent:.1f}%)")
                
                # Warn if memory usage is high
                if end_percent > 75:
                    print(f"   ⚠️ High memory usage detected: {end_percent:.1f}%")
                    
                return result
            return wrapper
        return decorator
    
    @staticmethod
    def get_optimal_chunk_size(available_memory_mb: float, columns: int) -> int:
        """Calculate optimal chunk size based on available memory and columns"""
        # Estimate memory per row (bytes)
        estimated_bytes_per_cell = 50
        bytes_per_row = columns * estimated_bytes_per_cell
        
        # Use 20% of available memory for chunk processing
        chunk_memory_budget = available_memory_mb * 0.2 * 1024 * 1024
        
        # Calculate optimal chunk size
        optimal_chunk_size = int(chunk_memory_budget / bytes_per_row)
        
        # Ensure reasonable bounds
        return max(500, min(optimal_chunk_size, 10000))
        
    @staticmethod
    def log_memory_stats():
        """Log current memory statistics"""
        vm = psutil.virtual_memory()
        process_memory = MemoryOptimizer.get_memory_usage()
        
        print(f"   Memory Stats:")
        print(f"      System: {vm.percent:.1f}% used ({vm.used / 1024 / 1024 / 1024:.1f}GB / {vm.total / 1024 / 1024 / 1024:.1f}GB)")
        print(f"      Process: {process_memory:.1f}MB")
        print(f"      Available: {vm.available / 1024 / 1024:.1f}MB")