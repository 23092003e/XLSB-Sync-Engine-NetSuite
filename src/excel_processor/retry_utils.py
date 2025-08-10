# excel_processor/retry_utils.py
import time
import random
from functools import wraps
from typing import Callable, Type, Tuple, Any
from .exceptions import RetryableError, COMTimeoutError, ExcelHangError

class RetryConfig:
    def __init__(self, 
                 max_attempts: int = 3, 
                 base_delay: float = 1.0, 
                 max_delay: float = 60.0, 
                 backoff_factor: float = 2.0,
                 jitter: bool = True):
        self.max_attempts = max_attempts
        self.base_delay = base_delay
        self.max_delay = max_delay
        self.backoff_factor = backoff_factor
        self.jitter = jitter

    def get_delay(self, attempt: int) -> float:
        """Calculate delay for given attempt with exponential backoff"""
        delay = self.base_delay * (self.backoff_factor ** attempt)
        delay = min(delay, self.max_delay)
        
        if self.jitter:
            # Add random jitter (±25%)
            jitter_factor = 0.25
            jitter = delay * jitter_factor * (2 * random.random() - 1)
            delay += jitter
            
        return max(0, delay)

def retry_on_failure(config: RetryConfig = None, 
                    retryable_exceptions: Tuple[Type[Exception], ...] = (RetryableError,)):
    """
    Decorator that implements retry logic with exponential backoff
    """
    if config is None:
        config = RetryConfig()
    
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            last_exception = None
            
            for attempt in range(config.max_attempts):
                try:
                    return func(*args, **kwargs)
                    
                except retryable_exceptions as e:
                    last_exception = e
                    
                    if attempt == config.max_attempts - 1:
                        # Last attempt failed, raise the exception
                        print(f"   ❌ Final attempt {attempt + 1} failed: {e}")
                        raise e
                    
                    delay = config.get_delay(attempt)
                    print(f"   🔄 Attempt {attempt + 1} failed: {e}")
                    print(f"   ⏳ Retrying in {delay:.2f} seconds...")
                    
                    time.sleep(delay)
                    
                except Exception as e:
                    # Non-retryable exception, raise immediately
                    print(f"   ❌ Non-retryable error: {e}")
                    raise e
                    
            # This should never be reached, but just in case
            if last_exception:
                raise last_exception
            else:
                raise RuntimeError("Unexpected retry logic failure")
                
        return wrapper
    return decorator

def safe_excel_operation_with_retry(operation: Callable, 
                                   operation_name: str = "Excel operation",
                                   max_attempts: int = 3) -> Any:
    """
    Execute Excel operations with automatic retry and proper error handling
    """
    config = RetryConfig(max_attempts=max_attempts, base_delay=0.5, max_delay=5.0)
    
    @retry_on_failure(config, (COMTimeoutError, ExcelHangError))
    def execute_operation():
        try:
            return operation()
        except Exception as e:
            # Convert common Excel errors to retryable errors when appropriate
            error_msg = str(e).lower()
            
            if any(keyword in error_msg for keyword in ['timeout', 'busy', 'not responding']):
                raise COMTimeoutError(f"{operation_name} timed out: {e}")
            elif any(keyword in error_msg for keyword in ['rpc', 'com', 'remote procedure']):
                raise COMTimeoutError(f"{operation_name} COM error: {e}")
            else:
                # Non-retryable error, raise as-is
                raise e
    
    return execute_operation()

def cleanup_on_failure(cleanup_func: Callable):
    """
    Decorator that ensures cleanup is performed when an operation fails
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            try:
                return func(*args, **kwargs)
            except Exception as e:
                try:
                    cleanup_func()
                except Exception as cleanup_error:
                    print(f"   ⚠️ Cleanup failed: {cleanup_error}")
                raise e
        return wrapper
    return decorator