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
        gc.collect()
        
    @staticmethod
    def monitor_memory_usage(operation_name: str):
        """Decorator to monitor memory usage of operations"""
        def decorator(func):
            def wrapper(*args, **kwargs):
                start_memory = MemoryOptimizer.get_memory_usage()
                result = func(*args, **kwargs)
                end_memory = MemoryOptimizer.get_memory_usage()
                memory_delta = end_memory - start_memory
                print(f"   📊 {operation_name}: {memory_delta:+.1f}MB (now {end_memory:.1f}MB)")
                return result
            return wrapper
        return decorator