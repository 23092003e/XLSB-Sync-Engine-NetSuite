# excel_processor/performance_logger.py
import time
import json
import psutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional
from dataclasses import dataclass, asdict
from .memory_optimizer import MemoryOptimizer

@dataclass
class PerformanceMetrics:
    operation: str
    start_time: datetime
    end_time: Optional[datetime] = None
    duration_seconds: float = 0.0
    memory_start_mb: float = 0.0
    memory_end_mb: float = 0.0
    memory_delta_mb: float = 0.0
    memory_percent_start: float = 0.0
    memory_percent_end: float = 0.0
    cpu_percent: float = 0.0
    rows_processed: int = 0
    columns_processed: int = 0
    file_size_mb: float = 0.0
    success: bool = True
    error_message: str = ""
    additional_data: Dict[str, Any] = None

class PerformanceLogger:
    def __init__(self, log_file: str = "performance_log.json"):
        self.log_file = Path(log_file)
        self.active_operations: Dict[str, PerformanceMetrics] = {}
        
    def start_operation(self, operation_id: str, operation_name: str, 
                       file_path: str = None, **kwargs) -> str:
        """Start tracking performance for an operation"""
        metrics = PerformanceMetrics(
            operation=operation_name,
            start_time=datetime.now(),
            memory_start_mb=MemoryOptimizer.get_memory_usage(),
            memory_percent_start=MemoryOptimizer.get_memory_percent(),
            cpu_percent=psutil.cpu_percent(),
            additional_data=kwargs
        )
        
        # Get file size if path provided
        if file_path:
            try:
                metrics.file_size_mb = Path(file_path).stat().st_size / (1024 * 1024)
            except:
                pass
                
        self.active_operations[operation_id] = metrics
        return operation_id
        
    def end_operation(self, operation_id: str, success: bool = True, 
                     error_message: str = "", rows: int = 0, columns: int = 0, **kwargs):
        """End tracking and log performance metrics"""
        if operation_id not in self.active_operations:
            return
            
        metrics = self.active_operations[operation_id]
        metrics.end_time = datetime.now()
        metrics.duration_seconds = (metrics.end_time - metrics.start_time).total_seconds()
        metrics.memory_end_mb = MemoryOptimizer.get_memory_usage()
        metrics.memory_delta_mb = metrics.memory_end_mb - metrics.memory_start_mb
        metrics.memory_percent_end = MemoryOptimizer.get_memory_percent()
        metrics.success = success
        metrics.error_message = error_message
        metrics.rows_processed = rows
        metrics.columns_processed = columns
        
        # Add any additional data
        if kwargs:
            if metrics.additional_data is None:
                metrics.additional_data = {}
            metrics.additional_data.update(kwargs)
            
        # Log the metrics
        self._log_metrics(metrics)
        
        # Remove from active operations
        del self.active_operations[operation_id]
        
        return metrics
        
    def _log_metrics(self, metrics: PerformanceMetrics):
        """Write metrics to log file"""
        try:
            # Convert to dict for JSON serialization
            metrics_dict = asdict(metrics)
            metrics_dict['start_time'] = metrics.start_time.isoformat()
            if metrics.end_time:
                metrics_dict['end_time'] = metrics.end_time.isoformat()
                
            # Append to log file
            with open(self.log_file, 'a', encoding='utf-8') as f:
                json.dump(metrics_dict, f, ensure_ascii=False)
                f.write('\n')
                
            # Print performance summary
            self._print_performance_summary(metrics)
            
        except Exception as e:
            print(f"   ⚠️ Failed to log performance metrics: {e}")
            
    def _print_performance_summary(self, metrics: PerformanceMetrics):
        """Print a formatted performance summary"""
        status = "✅ SUCCESS" if metrics.success else "❌ FAILED"
        
        print(f"   📊 PERFORMANCE: {metrics.operation} - {status}")
        print(f"      Duration: {metrics.duration_seconds:.2f}s")
        print(f"      Memory: {metrics.memory_delta_mb:+.1f}MB ({metrics.memory_start_mb:.1f} → {metrics.memory_end_mb:.1f}MB)")
        print(f"      Memory %: {metrics.memory_percent_start:.1f}% → {metrics.memory_percent_end:.1f}%")
        
        if metrics.rows_processed > 0:
            rows_per_sec = metrics.rows_processed / max(metrics.duration_seconds, 0.001)
            print(f"      Data: {metrics.rows_processed:,} rows × {metrics.columns_processed} cols ({rows_per_sec:.1f} rows/sec)")
            
        if metrics.file_size_mb > 0:
            mb_per_sec = metrics.file_size_mb / max(metrics.duration_seconds, 0.001)
            print(f"      File: {metrics.file_size_mb:.1f}MB ({mb_per_sec:.1f}MB/sec)")
            
        if not metrics.success:
            print(f"      Error: {metrics.error_message}")
            
    def get_summary_stats(self, hours_back: int = 24) -> Dict[str, Any]:
        """Get summary statistics from recent performance logs"""
        try:
            cutoff_time = datetime.now().timestamp() - (hours_back * 3600)
            
            stats = {
                'total_operations': 0,
                'successful_operations': 0,
                'failed_operations': 0,
                'total_duration': 0.0,
                'total_rows_processed': 0,
                'average_duration': 0.0,
                'average_memory_usage': 0.0,
                'operations_by_type': {},
                'error_types': {}
            }
            
            if not self.log_file.exists():
                return stats
                
            with open(self.log_file, 'r', encoding='utf-8') as f:
                for line in f:
                    try:
                        metrics = json.loads(line.strip())
                        
                        # Check if within time window
                        start_time = datetime.fromisoformat(metrics['start_time']).timestamp()
                        if start_time < cutoff_time:
                            continue
                            
                        stats['total_operations'] += 1
                        
                        if metrics['success']:
                            stats['successful_operations'] += 1
                        else:
                            stats['failed_operations'] += 1
                            error_type = metrics.get('error_message', 'Unknown').split(':')[0]
                            stats['error_types'][error_type] = stats['error_types'].get(error_type, 0) + 1
                            
                        stats['total_duration'] += metrics.get('duration_seconds', 0)
                        stats['total_rows_processed'] += metrics.get('rows_processed', 0)
                        
                        op_type = metrics['operation']
                        stats['operations_by_type'][op_type] = stats['operations_by_type'].get(op_type, 0) + 1
                        
                    except Exception:
                        continue  # Skip malformed entries
                        
            # Calculate averages
            if stats['total_operations'] > 0:
                stats['average_duration'] = stats['total_duration'] / stats['total_operations']
                
            return stats
            
        except Exception as e:
            print(f"   ⚠️ Failed to get summary stats: {e}")
            return {}

# Global performance logger instance
performance_logger = PerformanceLogger("logs/performance.jsonl")

def log_performance(operation_name: str, file_path: str = None):
    """Decorator for automatic performance logging"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            operation_id = f"{operation_name}_{int(time.time() * 1000)}"
            
            performance_logger.start_operation(
                operation_id, 
                operation_name, 
                file_path or kwargs.get('filepath', 'unknown')
            )
            
            try:
                result = func(*args, **kwargs)
                
                # Extract metrics from result if it's a ProcessingResult
                rows = getattr(result, 'rows_updated', 0) + getattr(result, 'rows_added', 0)
                
                performance_logger.end_operation(
                    operation_id, 
                    success=True, 
                    rows=rows,
                    result_status=getattr(result, 'status', 'unknown')
                )
                
                return result
                
            except Exception as e:
                performance_logger.end_operation(
                    operation_id, 
                    success=False, 
                    error_message=str(e)
                )
                raise
                
        return wrapper
    return decorator