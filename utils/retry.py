import time
import logging
import threading
from functools import wraps
from typing import Callable, Any, Tuple, Type, Optional
import platform

logger = logging.getLogger(__name__)

class TimeoutError(Exception):
    """Custom timeout exception"""
    pass

def timeout_wrapper(func, timeout_duration, args, kwargs):
    """Wrapper function to run with timeout using threading"""
    result = [None]
    exception = [None]
    
    def target():
        try:
            result[0] = func(*args, **kwargs)
        except Exception as e:
            exception[0] = e
    
    thread = threading.Thread(target=target)
    thread.daemon = True
    thread.start()
    thread.join(timeout_duration)
    
    if thread.is_alive():
        # Thread is still running, timeout occurred
        raise TimeoutError(f"Function {func.__name__} timeout after {timeout_duration} seconds")
    
    if exception[0]:
        raise exception[0]
    
    return result[0]

def retry_with_backoff(
    max_retries: int = 2,
    base_delay: float = 15.0,
    max_delay: float = 120.0,
    backoff_factor: float = 2.0,
    exceptions: Tuple[Type[Exception], ...] = (Exception,),
    timeout: Optional[float] = None
):
    """
    Retry decorator with exponential backoff
    
    Args:
        max_retries: Maximum number of retry attempts
        base_delay: Initial delay between retries (seconds)
        max_delay: Maximum delay between retries (seconds)
        backoff_factor: Multiplier for delay after each retry
        exceptions: Tuple of exceptions to catch and retry
        timeout: Timeout for the function call (seconds)
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            last_exception = None
            delay = base_delay
            
            for attempt in range(max_retries + 1):
                try:
                    if timeout:
                        # Use threading-based timeout for cross-platform compatibility
                        result = timeout_wrapper(func, timeout, args, kwargs)
                    else:
                        result = func(*args, **kwargs)
                    
                    if attempt > 0:
                        logger.info(f"{func.__name__} succeeded on attempt {attempt + 1}")
                    
                    return result
                    
                except exceptions as e:
                    last_exception = e
                    
                    if attempt == max_retries:
                        logger.error(f"{func.__name__} failed after {max_retries + 1} attempts: {str(e)}")
                        raise e
                    
                    logger.warning(f"{func.__name__} failed on attempt {attempt + 1}: {str(e)}")
                    logger.info(f"Retrying in {delay:.1f} seconds...")
                    
                    time.sleep(delay)
                    delay = min(delay * backoff_factor, max_delay)
            
            raise last_exception
        
        return wrapper
    return decorator

# Alternative timeout decorator using signal for Unix systems only
def retry_with_signal_timeout(
    max_retries: int = 3,
    base_delay: float = 30.0,
    max_delay: float = 300.0,
    backoff_factor: float = 2.0,
    exceptions: Tuple[Type[Exception], ...] = (Exception,),
    timeout: Optional[float] = None
):
    """
    Retry decorator with signal-based timeout (Unix/Linux only)
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            # Only use signal on Unix systems
            if platform.system() != 'Windows' and timeout:
                import signal
                
                def timeout_handler(signum, frame):
                    raise TimeoutError(f"Function {func.__name__} timeout after {timeout} seconds")
                
                last_exception = None
                delay = base_delay
                
                for attempt in range(max_retries + 1):
                    try:
                        old_handler = signal.signal(signal.SIGALRM, timeout_handler)
                        signal.alarm(int(timeout))
                        
                        result = func(*args, **kwargs)
                        
                        signal.alarm(0)  # Cancel alarm
                        signal.signal(signal.SIGALRM, old_handler)
                        
                        if attempt > 0:
                            logger.info(f"{func.__name__} succeeded on attempt {attempt + 1}")
                        
                        return result
                        
                    except exceptions as e:
                        last_exception = e
                        
                        signal.alarm(0)  # Cancel alarm
                        signal.signal(signal.SIGALRM, old_handler)
                        
                        if attempt == max_retries:
                            logger.error(f"{func.__name__} failed after {max_retries + 1} attempts: {str(e)}")
                            raise e
                        
                        logger.warning(f"{func.__name__} failed on attempt {attempt + 1}: {str(e)}")
                        logger.info(f"Retrying in {delay:.1f} seconds...")
                        
                        time.sleep(delay)
                        delay = min(delay * backoff_factor, max_delay)
                
                raise last_exception
            else:
                # Fallback to threading-based timeout for Windows or no timeout
                return retry_with_backoff(max_retries, base_delay, max_delay, backoff_factor, exceptions, timeout)(func)(*args, **kwargs)
        
        return wrapper
    return decorator