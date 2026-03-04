import logging
import os
import sys
from datetime import datetime
from pathlib import Path


class CSVProcessorLogger:
    """
    Custom logger that saves to file and optionally displays to console.
    """

    def __init__(
        self, name="CSVProcessor", log_to_file=True, log_to_console=True, log_dir="logs"
    ):
        self.logger = logging.getLogger(name)
        self.logger.setLevel(logging.DEBUG)
        self.logger.handlers.clear()

        # Create formatters
        file_formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s"
        )
        console_formatter = logging.Formatter("%(levelname)s: %(message)s")

        # File handler
        if log_to_file:
            # Create log directory if it doesn't exist
            os.makedirs(log_dir, exist_ok=True)

            # Create log filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_file = os.path.join(log_dir, f"csv_processor_{timestamp}.log")

            file_handler = logging.FileHandler(log_file, encoding="utf-8", mode="w")
            file_handler.setLevel(logging.DEBUG)
            file_handler.setFormatter(file_formatter)
            self.logger.addHandler(file_handler)

            # Store log file path for reference
            self.log_file = log_file
        else:
            self.log_file = None

        # Console handler (optional, controlled by verbose flag)
        if log_to_console:
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setLevel(logging.INFO)
            console_handler.setFormatter(console_formatter)
            self.logger.addHandler(console_handler)

    def get_logger(self):
        """Return the configured logger instance."""
        return self.logger

    def get_log_file(self):
        """Return the path to the current log file."""
        return self.log_file


_logger_instance = None


def get_logger(name="CSVProcessor", verbose=False):

    global _logger_instance
    if _logger_instance is None:
        _logger_instance = CSVProcessorLogger(
            name, log_to_file=True, log_to_console=verbose
        )
    return _logger_instance.get_logger()


def get_log_file_path():
    """Get the path to the current log file."""
    global _logger_instance
    if _logger_instance:
        return _logger_instance.get_log_file()
    return None


def set_verbosity(verbose):
    global _logger_instance
    if _logger_instance:
        logger = _logger_instance.get_logger()
        for handler in logger.handlers:
            if isinstance(handler, logging.StreamHandler) and not isinstance(
                handler, logging.FileHandler
            ):
                handler.setLevel(logging.INFO if verbose else logging.WARNING)
