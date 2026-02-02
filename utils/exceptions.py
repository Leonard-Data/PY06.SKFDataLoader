from typing import Optional

class SKFError(Exception):
    """Base class for SKF Data Loader exceptions."""
    code: Optional[str] = None
    retryable: bool = False
    
    def __init__(self, message: str = "", code: Optional[str] = None, *, retryable: Optional[bool] = None):
        super().__init__(message)
        if code:
            self.code = code
        if retryable is not None:
            self.retryable = retryable

# Authentication & Authorization errors
class AuthenticationError(SKFError):
    """Authentication/Token related errors."""
    code = "AUTH_ERROR"
    retryable = True  # Tokens can be refreshed

class TokenExpiredError(AuthenticationError):
    """Access token has expired."""
    code = "TOKEN_EXPIRED"
    retryable = True

class TokenNotAvailableError(AuthenticationError):
    """Access token not available."""
    code = "TOKEN_UNAVAILABLE"
    retryable = True

# Data/Business Logic errors
class DataError(SKFError):
    """Base class for data-related errors."""
    retryable = False

class MasterDataError(DataError):
    """Master data processing errors."""
    code = "MASTERDATA_ERROR"

class DataValidationError(DataError):
    """Data validation failed."""
    code = "DATA_VALIDATION_ERROR"

class NoDataFoundError(DataError):
    """No data found for processing."""
    code = "NO_DATA_FOUND"

class ConsolidationError(DataError):
    """Data consolidation failed."""
    code = "CONSOLIDATION_ERROR"

# File Operation errors
class FileError(SKFError):
    """Base class for file operation errors."""
    retryable = False

class FileNotFoundError(FileError):
    """File not found."""
    code = "FILE_NOT_FOUND"

class FileProcessingError(FileError):
    """File processing failed."""
    code = "FILE_PROCESSING_ERROR"

class ExcelCombinationError(FileError):
    """Excel file combination failed."""
    code = "EXCEL_COMBINATION_ERROR"

# SharePoint/Network related errors (retryable)
class SharePointError(SKFError):
    """Base class for SharePoint errors."""
    retryable = True

class SharePointConnectionError(SharePointError):
    """SharePoint connection failed."""
    code = "SHAREPOINT_CONNECTION_ERROR"
    retryable = True

class SharePointAccessDeniedError(SharePointError):
    """SharePoint access denied."""
    code = "SHAREPOINT_ACCESS_DENIED"
    retryable = False  # Permission issue, no point retrying

class SharePointDownloadError(SharePointError):
    """SharePoint file download failed."""
    code = "SHAREPOINT_DOWNLOAD_ERROR"
    retryable = True

class SharePointUploadError(SharePointError):
    """SharePoint file upload failed."""
    code = "SHAREPOINT_UPLOAD_ERROR"
    retryable = True

class SharePointTimeoutError(SharePointError):
    """SharePoint operation timed out."""
    code = "SHAREPOINT_TIMEOUT"
    retryable = True

# API/External Service errors (retryable)
class APIError(SKFError):
    """Base class for API errors."""
    retryable = True

class DataverseAPIError(APIError):
    """Dataverse API errors."""
    code = "DATAVERSE_API_ERROR"
    retryable = True

class GraphAPIError(APIError):
    """Microsoft Graph API errors."""
    code = "GRAPH_API_ERROR"
    retryable = True

class APITimeoutError(APIError):
    """API request timed out."""
    code = "API_TIMEOUT"
    retryable = True

class APIRateLimitError(APIError):
    """API rate limit exceeded."""
    code = "API_RATE_LIMIT"
    retryable = True

# Network/Infrastructure errors (retryable)
class NetworkError(SKFError):
    """Base class for network errors."""
    retryable = True

class ConnectionResetError(NetworkError):
    """Connection was reset by remote host."""
    code = "CONNECTION_RESET"
    retryable = True

class TimeoutError(NetworkError):
    """Operation timed out."""
    code = "TIMEOUT"
    retryable = True

class ConnectionError(NetworkError):
    """Connection failed."""
    code = "CONNECTION_ERROR"
    retryable = True

# Process/System errors
class ProcessError(SKFError):
    """Process management errors."""
    retryable = False

class ProcessKillError(ProcessError):
    """Failed to kill process."""
    code = "PROCESS_KILL_ERROR"

# Configuration errors
class ConfigurationError(SKFError):
    """Configuration related errors."""
    retryable = False

class EnvironmentVariableError(ConfigurationError):
    """Environment variable missing or invalid."""
    code = "ENV_VAR_ERROR"

# Critical errors (non-retryable)
class CriticalError(SKFError):
    """Critical system errors that should not be retried."""
    retryable = False

class FolderCreationError(CriticalError):
    """Failed to create working folder."""
    code = "FOLDER_CREATION_ERROR"

class MemoryError(CriticalError):
    """Memory allocation failed."""
    code = "MEMORY_ERROR"

class SystemError(CriticalError):
    """System-level error."""
    code = "SYSTEM_ERROR"

# Helper functions
def is_retryable_error(error: Exception) -> bool:
    """Check if an error is retryable."""
    if isinstance(error, SKFError):
        return error.retryable
    # For standard Python exceptions, decide case by case
    if isinstance(error, (ConnectionAbortedError, ConnectionResetError, OSError)):
        return True
    return False

def get_error_code(error: Exception) -> Optional[str]:
    """Get error code from exception."""
    if isinstance(error, SKFError):
        return error.code
    return error.__class__.__name__

def wrap_exception(error: Exception, message: str = "") -> SKFError:
    """Wrap standard exceptions into SKF exceptions."""
    import requests
    
    # Network/HTTP errors
    if isinstance(error, requests.exceptions.ConnectionError):
        return ConnectionError(message or str(error))
    elif isinstance(error, requests.exceptions.Timeout):
        return TimeoutError(message or str(error))
    elif isinstance(error, requests.exceptions.HTTPError):
        if error.response and error.response.status_code == 401:
            return AuthenticationError(message or str(error))
        elif error.response and error.response.status_code == 403:
            return SharePointAccessDeniedError(message or str(error))
        elif error.response and error.response.status_code == 429:
            return APIRateLimitError(message or str(error))
        else:
            return APIError(message or str(error))
    
    # File errors
    elif isinstance(error, FileNotFoundError):
        return FileNotFoundError(message or str(error))
    elif isinstance(error, PermissionError):
        return SharePointAccessDeniedError(message or str(error))
    
    # Memory/System errors
    elif isinstance(error, MemoryError):
        return MemoryError(message or str(error))
    elif isinstance(error, OSError):
        return SystemError(message or str(error))
    
    # Default: wrap as generic SKF error
    else:
        return SKFError(message or str(error), retryable=False)