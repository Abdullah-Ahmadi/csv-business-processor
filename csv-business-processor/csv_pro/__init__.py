"""Top-level package for CSV Business Processor.

This module exposes convenient imports and helpers for the package so the
project can be used as a module (e.g. `python -m csv_pro.cli`) or imported
in other Python code for programmatic use.
"""

__version__ = "0.1.0"

# CLI entrypoint (alias: `run`) — this is the function invoked when running
# the package as a module: `python -m csv_pro.cli` or programmatically.
from .cli import main as run

# Processors
from .processors import EcommerceProcessor, InventoryProcessor, FinanceProcessor

# Exporters
from .exporters import (
	ConsoleExporter,
	JSONExporter,
	ExcelExporter,
	PDFExporter,
)

# Utilities (formatters and file helpers)
from .utils import (
	format_currency,
	format_percentage,
	chk_input_file_format,
	chk_output_file_format,
	validate_csv,
	detect_csv_delimeter,
	generate_filename,
)

# Logger helpers
from .utils.logger import get_logger, get_log_file_path, set_verbosity

__all__ = [
	"__version__",
	"run",
	# processors
	"EcommerceProcessor",
	"InventoryProcessor",
	"FinanceProcessor",
	# exporters
	"ConsoleExporter",
	"JSONExporter",
	"ExcelExporter",
	"PDFExporter",
	# utils
	"format_currency",
	"format_percentage",
	"chk_input_file_format",
	"chk_output_file_format",
	"validate_csv",
	"detect_csv_delimeter",
	"generate_filename",
	# logger
	"get_logger",
	"get_log_file_path",
	"set_verbosity",
]
