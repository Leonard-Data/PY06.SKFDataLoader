import time
import logging
import threading
import random
import os
from functools import wraps
from typing import Callable, Any, Tuple, Type, Optional, List, Union
from dotenv import load_dotenv
from utils.exceptions import SKFError, is_retryable_error, wrap_exception, TimeoutError

# Load environment variables
load_dotenv()

logger = logging.getLogger(__name__)

class RetryConfig:
    """Configuration class for retry settings"""
    def __init__(self):
        # Load from environment variables with defaults
        self.max_retries = int(os.getenv('RETRY_MAX_RETRIES', '1'))
        self.base_delay = float(os.getenv('RETRY_BASE_DELAY', '10.0'))
        self.max_delay = float(os.getenv('RETRY_MAX_DELAY', '60.0'))
        self.backoff_factor = float(os.getenv('RETRY_BACKOFF_FACTOR', '1.5'))
        self.timeout = float(os.getenv('RETRY_TIMEOUT', '120.0')) if os.getenv('RETRY_TIMEOUT') else None
        self.jitter_enabled = os.getenv('RETRY_JITTER_ENABLED', 'true').lower() == 'true'
        self.jitter_max = float(os.getenv('RETRY_JITTER_MAX', '3.0'))
        
        # Log loaded configuration
        logger.info(f"Retry config loaded - Max retries: {self.max_retries}, "
                   f"Base delay: {self.base_delay}s, Max delay: {self.max_delay}s, "
                   f"Timeout: {self.timeout}s, Jitter: {self.jitter_enabled}")

# Global configuration instance
_retry_config = RetryConfig()

def get_retry_config() -> RetryConfig:
    """Get the global retry configuration"""
    return _retry_config

def reload_retry_config():
    """Reload retry configuration from environment"""
    global _retry_config
    _retry_config = RetryConfig()

def timeout_wrapper(func: Callable, timeout_duration: float, args: Tuple, kwargs: dict) -> Any:
    """Wrapper function to run with timeout using threading"""
    result: List[Optional[Any]] = [None]
    exception: List[Optional[Exception]] = [None]
    
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
    
    if exception[0] is not None:
        raise exception[0]
    
    return result[0]

def add_jitter(delay: float, jitter_max: Optional[float] = None) -> float:
    """Add random jitter to delay to avoid thundering herd"""
    if jitter_max is None:
        jitter_max = _retry_config.jitter_max
    
    if _retry_config.jitter_enabled:
        jitter = random.uniform(0, min(jitter_max, delay * 0.1))  # Max 10% of delay
        return delay + jitter
    return delay

def retry_with_backoff(
    max_retries: Optional[int] = None,
    base_delay: Optional[float] = None,
    max_delay: Optional[float] = None,
    backoff_factor: Optional[float] = None,
    exceptions: Tuple[Type[Exception], ...] = (Exception,),
    timeout: Optional[float] = None,
    use_config: bool = True,
    jitter: Optional[bool] = None,
    auto_wrap: bool = True  # Automatically wrap standard exceptions
):
    """
    Retry decorator with exponential backoff and intelligent exception handling
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            # Use config values if not provided
            config = _retry_config if use_config else None
            
            _max_retries = max_retries if max_retries is not None else (config.max_retries if config else 1)
            _base_delay = base_delay if base_delay is not None else (config.base_delay if config else 10.0)
            _max_delay = max_delay if max_delay is not None else (config.max_delay if config else 60.0)
            _backoff_factor = backoff_factor if backoff_factor is not None else (config.backoff_factor if config else 1.5)
            _timeout = timeout if timeout is not None else (config.timeout if config else None)
            _jitter = jitter if jitter is not None else (config.jitter_enabled if config else True)
            
            last_exception: Optional[Exception] = None
            delay = _base_delay
            
            for attempt in range(_max_retries + 1):
                try:
                    if _timeout:
                        result = timeout_wrapper(func, _timeout, args, kwargs)
                    else:
                        result = func(*args, **kwargs)
                    
                    if attempt > 0:
                        logger.info(f"{func.__name__} succeeded on attempt {attempt + 1}")
                    
                    return result
                    
                except exceptions as e:
                    # Wrap exception if auto_wrap is enabled
                    if auto_wrap and not isinstance(e, SKFError):
                        e = wrap_exception(e)
                    
                    last_exception = e
                    
                    # Check if this exception is retryable
                    should_retry = is_retryable_error(e) if isinstance(e, SKFError) else True
                    
                    if attempt == _max_retries or not should_retry:
                        error_code = getattr(e, 'code', e.__class__.__name__)
                        logger.error(f"{func.__name__} failed after {attempt + 1} attempts: [{error_code}] {str(e)}")
                        
                        if not should_retry:
                            logger.error(f"Exception {error_code} is not retryable - failing immediately")
                        
                        raise e
                    
                    error_code = getattr(e, 'code', e.__class__.__name__)
                    logger.warning(f"{func.__name__} failed on attempt {attempt + 1}: [{error_code}] {str(e)}")
                    
                    # Calculate next delay with jitter
                    actual_delay = add_jitter(delay) if _jitter else delay
                    logger.info(f"Retrying in {actual_delay:.1f} seconds...")
                    
                    time.sleep(actual_delay)
                    delay = min(delay * _backoff_factor, _max_delay)
            
            if last_exception is not None:
                raise last_exception
            else:
                raise SKFError(f"{func.__name__} failed but no exception was captured.")
        
        return wrapper
    return decorator

# Predefined configurations optimized for 15-minute schedule
def retry_fast(max_retries: int = 0, base_delay: float = 5.0, timeout: float = 30.0):
    """Quick retry for fast operations (almost no retry)"""
    return retry_with_backoff(
        max_retries=max_retries,
        base_delay=base_delay,
        max_delay=15.0,
        timeout=timeout,
        use_config=False
    )

def retry_medium(max_retries: int = 1, base_delay: float = 10.0, timeout: float = 60.0):
    """Medium retry for normal operations"""
    return retry_with_backoff(
        max_retries=max_retries,
        base_delay=base_delay,
        max_delay=30.0,
        timeout=timeout,
        use_config=False
    )

def retry_heavy(max_retries: int = 1, base_delay: float = 15.0, timeout: float = 120.0):
    """Heavy retry for file downloads/uploads"""
    return retry_with_backoff(
        max_retries=max_retries,
        base_delay=base_delay,
        max_delay=60.0,
        timeout=timeout,
        use_config=False
    )