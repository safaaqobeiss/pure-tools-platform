"""
Centralized logging module for Excel file generation.
Generates a log file for each new Excel file created.
"""

import os
import logging
import datetime
from typing import Optional, Dict, Any


class ExcelFileHandler(logging.Handler):
    """
    Custom logging handler that writes to Excel-specific log files.
    """
    
    def __init__(self, log_file_path: str):
        super().__init__()
        self.log_file_path = log_file_path
        self.file_handle = open(log_file_path, 'w', encoding='utf-8')
        self._closed = False
        
    def emit(self, record):
        """
        Emit a log record to the file.
        """
        try:
            # Check if file is still open before writing
            if self._closed or not self.file_handle or self.file_handle.closed:
                return  # Silently ignore if file is closed
            
            msg = self.format(record)
            self.file_handle.write(msg + '\n')
            self.file_handle.flush()  # Force immediate write
        except (ValueError, OSError):
            # File is closed or I/O error, silently ignore
            pass
        except Exception:
            self.handleError(record)
    
    def close(self):
        """
        Close the file handle.
        """
        if not self._closed and hasattr(self, 'file_handle') and self.file_handle and not self.file_handle.closed:
            try:
                self.file_handle.flush()
                self.file_handle.close()
            except Exception:
                pass
        self._closed = True
        super().close()


class ExcelFileLogger:
    """
    Specialized logger for Excel file creation.
    Creates a log file per generated Excel file.
    """
    
    def __init__(self, base_output_dir: str = None):
        """
        Initialize the Excel file logger.
        
        Args:
            base_output_dir: Base directory for logs. If None, uses Excel file directory.
        """
        self.base_output_dir = base_output_dir
        self.active_loggers: Dict[str, logging.Logger] = {}
        self.log_files: Dict[str, str] = {}
        
    def create_logger_for_excel_file(self, excel_file_path: str, log_prefix: str = "excel_generation") -> logging.Logger:
        """
        Creates a specific logger for an Excel file.
        
        Args:
            excel_file_path: Path to the Excel file
            log_prefix: Prefix for the log file name
            
        Returns:
            Configured logger for this Excel file
        """
        # Extract filename without extension
        excel_filename = os.path.splitext(os.path.basename(excel_file_path))[0]
        
        # Create unique name for the logger
        logger_name = f"{log_prefix}_{excel_filename}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        # Determine log directory
        if self.base_output_dir:
            log_dir = os.path.join(self.base_output_dir, "logs")
        else:
            log_dir = os.path.join(os.path.dirname(excel_file_path), "logs")
        
        # Create log directory if it doesn't exist
        os.makedirs(log_dir, exist_ok=True)
        
        # Create log filename
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        log_filename = f"{log_prefix}_{excel_filename}_{timestamp}.log"
        log_file_path = os.path.join(log_dir, log_filename)
        
        # Create the logger
        logger = logging.getLogger(logger_name)
        logger.setLevel(logging.DEBUG)
        
        # Avoid handler duplication
        if logger.handlers:
            logger.handlers.clear()
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Custom file handler that flushes immediately
        file_handler = ExcelFileHandler(log_file_path)
        file_handler.setLevel(logging.DEBUG)  # Capture all levels
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        
        # No console handler to avoid interference
        logger.propagate = False  # Prevent propagation to root logger
        
        # Store references
        self.active_loggers[excel_file_path] = logger
        self.log_files[excel_file_path] = log_file_path
        
        # Initial log with Excel file information
        logger.info("=" * 70)
        logger.info(f"START OF EXCEL FILE GENERATION")
        logger.info(f"Excel file: {excel_file_path}")
        logger.info(f"Log file: {log_file_path}")
        logger.info(f"Timestamp: {datetime.datetime.now().isoformat()}")
        logger.info("=" * 70)
        
        return logger
    
    def get_logger_for_excel_file(self, excel_file_path: str) -> Optional[logging.Logger]:
        """
        Retrieves existing logger for an Excel file.
        
        Args:
            excel_file_path: Path to the Excel file
            
        Returns:
            Existing logger or None
        """
        return self.active_loggers.get(excel_file_path)
    
    def log_excel_creation_start(self, excel_file_path: str, **metadata):
        """
        Log the start of Excel file creation.
        
        Args:
            excel_file_path: Path to the Excel file
            **metadata: Additional metadata to log
        """
        logger = self.get_logger_for_excel_file(excel_file_path)
        if logger:
            logger.info("📊 START OF EXCEL FILE CREATION")
            for key, value in metadata.items():
                logger.info(f"   {key}: {value}")
    
    def log_excel_creation_step(self, excel_file_path: str, step_name: str, details: str = ""):
        """
        Log a step in Excel file creation.
        
        Args:
            excel_file_path: Path to the Excel file
            step_name: Name of the step
            details: Additional details
        """
        logger = self.get_logger_for_excel_file(excel_file_path)
        if logger:
            logger.info(f"🔧 {step_name}")
            if details:
                logger.info(f"   {details}")
    
    def log_excel_creation_error(self, excel_file_path: str, error_msg: str, exception: Exception = None):
        """
        Log an error during Excel file creation.
        
        Args:
            excel_file_path: Path to the Excel file
            error_msg: Error message
            exception: Optional exception
        """
        logger = self.get_logger_for_excel_file(excel_file_path)
        if logger:
            logger.error(f"❌ ERROR: {error_msg}")
            if exception:
                logger.error(f"   Exception: {str(exception)}")
    
    def log_excel_creation_complete(self, excel_file_path: str, success: bool = True, **summary):
        """
        Log the completion of Excel file creation.
        
        Args:
            excel_file_path: Path to the Excel file
            success: Success indicator
            **summary: Summary of operations performed
        """
        logger = self.get_logger_for_excel_file(excel_file_path)
        if logger:
            status = "✅ SUCCESS" if success else "❌ FAILURE"
            logger.info("=" * 70)
            logger.info(f"END OF EXCEL FILE GENERATION - {status}")
            logger.info(f"Excel file: {excel_file_path}")
            logger.info(f"Log file: {self.log_files.get(excel_file_path, 'N/A')}")
            
            if summary:
                logger.info("SUMMARY:")
                for key, value in summary.items():
                    logger.info(f"   {key}: {value}")
            
            logger.info(f"End time: {datetime.datetime.now().isoformat()}")
            logger.info("=" * 70)
            
            # Force flush of all handlers to ensure everything is written
            self.flush_logger(excel_file_path)
    
    def flush_logger(self, excel_file_path: str):
        """
        Force flush all handlers for a specific Excel file logger.
        
        Args:
            excel_file_path: Path to the Excel file
        """
        logger = self.get_logger_for_excel_file(excel_file_path)
        if logger:
            for handler in logger.handlers:
                if isinstance(handler, ExcelFileHandler):
                    handler.file_handle.flush()
    
    def cleanup_logger(self, excel_file_path: str):
        """
        Cleans up the logger for a specific Excel file.
        
        Args:
            excel_file_path: Path to the Excel file
        """
        # Force final flush before cleanup
        self.flush_logger(excel_file_path)
        
        logger = self.active_loggers.get(excel_file_path)
        
        if logger:
            # Close all handlers and remove them
            for handler in logger.handlers[:]:
                handler.close()
                logger.removeHandler(handler)
        
        # Remove from our tracking dictionaries
        if excel_file_path in self.active_loggers:
            del self.active_loggers[excel_file_path]
        if excel_file_path in self.log_files:
            del self.log_files[excel_file_path]
            
        # Remove from dictionaries
        self.active_loggers.pop(excel_file_path, None)
        self.log_files.pop(excel_file_path, None)


# Global instance for ease of use
excel_file_logger = ExcelFileLogger()


def setup_excel_file_logging(excel_file_path: str, log_prefix: str = "excel_generation") -> logging.Logger:
    """
    Utility function to quickly configure logging for an Excel file.
    
    Args:
        excel_file_path: Path to the Excel file
        log_prefix: Prefix for the log file name
        
    Returns:
        Configured logger
    """
    logger = excel_file_logger.create_logger_for_excel_file(excel_file_path, log_prefix)
    # Don't setup global debug capture to avoid file corruption
    return logger


def log_excel_step(excel_file_path: str, step_name: str, details: str = ""):
    """
    Utility function to log a step.
    
    Args:
        excel_file_path: Path to the Excel file
        step_name: Name of the step
        details: Additional details
    """
    excel_file_logger.log_excel_creation_step(excel_file_path, step_name, details)
    # Force flush after each step to ensure logs are written immediately
    excel_file_logger.flush_logger(excel_file_path)


def log_excel_error(excel_file_path: str, error_msg: str, exception: Exception = None):
    """
    Utility function to log an error.
    
    Args:
        excel_file_path: Path to the Excel file
        error_msg: Error message
        exception: Optional exception
    """
    excel_file_logger.log_excel_creation_error(excel_file_path, error_msg, exception)
    # Force flush after errors to ensure they're captured
    excel_file_logger.flush_logger(excel_file_path)


def complete_excel_logging(excel_file_path: str, success: bool = True, **summary):
    """
    Utility function to finalize logging.
    
    Args:
        excel_file_path: Path to the Excel file
        success: Success indicator
        **summary: Summary of operations
    """
    excel_file_logger.log_excel_creation_complete(excel_file_path, success, **summary)
    excel_file_logger.cleanup_logger(excel_file_path)
