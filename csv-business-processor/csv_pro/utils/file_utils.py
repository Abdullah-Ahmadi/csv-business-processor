import argparse
import os

from datetime import datetime


def chk_output_file_format(filename):
    if not filename.endswith((".xlsx", ".json", ".pdf", "console")):
        raise argparse.ArgumentTypeError(
            f"Invalid file format: {filename}. Must be .xlsx, .json, .pdf, or 'console'."
        )
    return filename


def chk_input_file_format(filename):
    if not filename.endswith(".csv"):
        raise argparse.ArgumentTypeError(
            f"Invalid file format: {filename}. Must be .csv ."
        )
    return filename


def validate_csv(filename):
    if not os.path.exists(filename):
        raise FileNotFoundError(f"File not found {filename}")


def detect_csv_delimeter(filepath):
    with open(filepath, "r") as file:
        first_line = file.readline()

        if ";" in first_line:
            return ";"
        elif "\t" in first_line:
            return "\t"
        else:
            return ","


def generate_filename(processor, file_extension):
    """Generate a filename"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    processor_name = processor.__class__.__name__.replace("Processor", "").lower()
    return f"{processor_name}_report_{timestamp}{file_extension}"
