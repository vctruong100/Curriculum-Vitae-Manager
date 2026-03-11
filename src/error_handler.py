"""
Custom error handling for common file permission errors.
"""

import os
from pathlib import Path
from typing import Optional, Tuple


class FilePermissionError(Exception):
    """Custom exception for file permission errors with user-friendly messages."""
    
    def __init__(self, file_path: Path, operation: str = "access"):
        self.file_path = file_path
        self.operation = operation
        super().__init__(self._get_message())
    
    def _get_message(self) -> str:
        """Generate user-friendly error message."""
        filename = self.file_path.name
        return (
            f"Cannot {self.operation} the file:\n\n"
            f"{self.file_path}\n\n"
            f"The file appears to be open in another program.\n"
            f"Please close '{filename}' and try again."
        )


def handle_file_operation(func):
    """
    Decorator to catch and convert PermissionError to FilePermissionError.
    
    Usage:
        @handle_file_operation
        def save_file(path):
            # file operations
    """
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except PermissionError as e:
            # Try to extract file path from args
            file_path = None
            for arg in args:
                if isinstance(arg, (str, Path)):
                    file_path = Path(arg)
                    break
            
            if file_path is None:
                # Check kwargs
                for key in ['path', 'file_path', 'output_path', 'cv_path', 'master_path']:
                    if key in kwargs:
                        file_path = Path(kwargs[key])
                        break
            
            if file_path:
                raise FilePermissionError(file_path, "save") from e
            else:
                # Fallback to generic message
                raise FilePermissionError(Path("the file"), "access") from e
    
    return wrapper


def check_file_writable(file_path: Path) -> Tuple[bool, Optional[str]]:
    """
    Check if a file is writable (not locked by another process).
    
    Returns: (is_writable, error_message)
    """
    if not file_path.exists():
        # File doesn't exist yet, check if directory is writable
        parent = file_path.parent
        if not parent.exists():
            return False, f"Directory does not exist: {parent}"
        return os.access(parent, os.W_OK), None
    
    # Try to open file in append mode to check if it's locked
    try:
        with open(file_path, 'a'):
            pass
        return True, None
    except PermissionError:
        return False, (
            f"Cannot access the file:\n\n"
            f"{file_path}\n\n"
            f"The file appears to be open in another program.\n"
            f"Please close '{file_path.name}' and try again."
        )
    except Exception as e:
        return False, str(e)
