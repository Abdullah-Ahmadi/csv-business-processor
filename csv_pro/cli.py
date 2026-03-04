import argparse
import os
import sys
import logging
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from utils.file_utils import chk_input_file_format, chk_output_file_format, validate_csv, detect_csv_delimeter, generate_filename
from utils.formatters import format_currency, format_percentage
from utils.logger import get_logger, get_log_file_path, set_verbosity
from processors import InventoryProcessor, EcommerceProcessor, FinanceProcessor
from exporters import ConsoleExporter, JSONExporter, ExcelExporter, PDFExporter


def main():
    parser = argparse.ArgumentParser(
        description="📊 CSV Processor - Transform business CSV files into actionable insights",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process e-commerce sales data and view in console
  python csvpro.py sales.csv --mode ecommerce --output console

  # Analyze inventory and generate Excel report
  python csvpro.py inventory.csv --mode inventory --output excel --outfile stock_report.xlsx

  # Process expenses and save as JSON
  python csvpro.py expenses.csv --mode finance --output json

  # Full pipeline with all options
  python csvpro.py data.csv --mode ecommerce --output pdf --outfile report.pdf --verbose
        """
    )

    # Required arguments
    parser.add_argument(
        "input",
        help="Input CSV file path"
    )

    parser.add_argument(
        "--mode", "-m",
        help="Processing mode",
        choices=["ecommerce", "inventory", "finance"],
        type=str.lower,
        required=True
    )

    parser.add_argument(
        "--output", "-o",
        help="Output format",
        default="console",
        type=str.lower,
        choices=["console", "json", "excel", "pdf"]
    )

    # Optional arguments
    parser.add_argument(
        "--outfile", "-f",
        help="Output file name (auto-generated if not specified)",
        type=str,
        default=None
    )

    parser.add_argument(
        "--verbose", "-v",
        help="Show detailed processing information",
        action="store_true"
    )

    parser.add_argument(
        "--delimiter", "-d",
        help="CSV delimiter (auto-detected if not specified)",
        type=str,
        default=None
    )

    args = parser.parse_args()

    # Set up logging based on verbose flag
    logger = get_logger("CSVProcessor", verbose=args.verbose)
    set_verbosity(args.verbose)

    # Log execution start
    logger.info("=" * 60)
    logger.info("CSV PROCESSOR v1.0 - Starting execution")
    logger.info("=" * 60)
    logger.info(f"Input file: {args.input}")
    logger.info(f"Mode: {args.mode}")
    logger.info(f"Output format: {args.output}")
    logger.info(f"Verbose mode: {args.verbose}")
    logger.info(f"Log file: {get_log_file_path()}")

    # Validate input file
    try:
        validate_csv(args.input)
        logger.info(f"Input file validated successfully: {args.input}")
    except (FileNotFoundError, ValueError) as e:
        logger.error(f"Input validation failed: {e}")
        print(f"\n❌ Error: {e}")
        print(f"📋 Check log for details: {get_log_file_path()}")
        sys.exit(1)

    # Display header
    print("\n" + "="*60)
    print("📊 CSV PROCESSOR v1.0")
    print("="*60)
    print(f"📁 Input: {args.input}")
    print(f"⚙️  Mode: {args.mode.upper()}")
    print(f"📤 Output: {args.output.upper()}")
    if args.verbose:
        print(f"🔍 Verbose mode: ON")
    print(f"📋 Log file: {get_log_file_path()}")
    print("-"*60)

    # Select processor based on mode
    processor = None
    if args.mode == "ecommerce":
        processor = EcommerceProcessor()
        logger.info("Selected EcommerceProcessor")
    elif args.mode == "inventory":
        processor = InventoryProcessor()
        logger.info("Selected InventoryProcessor")
    elif args.mode == "finance":
        processor = FinanceProcessor()
        logger.info("Selected FinanceProcessor")

    if processor is None:
        logger.error(f"Unknown mode: {args.mode}")
        print(f"❌ Error: Unknown mode '{args.mode}'")
        sys.exit(1)

    # Load and analyze data
    try:
        print("📂 Loading data...")
        logger.info(f"Loading data from {args.input}")
        processor.load_data(args.input)
        print(f"✅ Loaded {len(processor.data)} rows")
        logger.info(f"Loaded {len(processor.data)} rows, {len(processor.data.columns)} columns")

        print("🔍 Analyzing data...")
        logger.info("Starting analysis")
        processor.analyze()
        print(f"✅ Generated {len(processor.insights)} insights and {len(processor.alerts)} alerts")
        logger.info(f"Analysis complete: {len(processor.insights)} insights, {len(processor.alerts)} alerts")

        # Store input filename for reference
        processor.input_file = args.input

        # Log alerts for debugging
        for i, alert in enumerate(processor.alerts):
            logger.warning(f"Alert {i+1}: {alert}")

    except Exception as e:
        logger.error(f"Error processing data: {e}", exc_info=True)
        print(f"\n❌ Error processing data: {e}")
        print(f"📋 Check log for details: {get_log_file_path()}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)

    # Generate output
    try:
        print(f"📤 Exporting to {args.output.upper()}...")
        logger.info(f"Exporting to {args.output} format")

        exporter = None
        output_path = args.outfile
        output = args.output.lower()

        # Auto-generate filename if not provided and not console
        if output_path is None and output != "console":
            output_path = generate_filename(processor, f".{output}")
            logger.info(f"Auto-generated filename: {output_path}")

        # Select exporter
        if output == "console":
            exporter = ConsoleExporter()
            logger.info("Using ConsoleExporter")
            result = exporter.export(processor)
            logger.info("Console export complete")

        elif output == "json":
            exporter = JSONExporter()
            logger.info("Using JSONExporter")
            result = exporter.export(processor, output_path)
            logger.info(f"JSON export complete: {result}")
            print(f"✅ JSON saved: {result}")

        elif output == "excel":
            exporter = ExcelExporter()
            logger.info("Using ExcelExporter")
            result = exporter.export(processor, output_path)
            logger.info(f"Excel export complete: {result}")
            print(f"✅ Excel saved: {result}")

        elif output == "pdf":
            try:
                exporter = PDFExporter()
                logger.info("Using PDFExporter")
                result = exporter.export(processor, output_path)
                logger.info(f"PDF export complete: {result}")
                print(f"✅ PDF saved: {result}")
            except ImportError as e:
                logger.error(f"PDF export failed - missing fpdf2: {e}")
                print(f"❌ PDF export requires: pip install fpdf2")
                sys.exit(1)
            except Exception as e:
                logger.error(f"PDF export failed: {e}", exc_info=True)
                print(f"❌ PDF export failed: {e}")
                print(f"📋 Check log for details: {get_log_file_path()}")
                sys.exit(1)

        print("\n" + "="*60)
        print("✅ Processing complete!")
        print(f"📋 Log saved to: {get_log_file_path()}")
        print("="*60)
        logger.info("Processing completed successfully")

    except Exception as e:
        logger.error(f"Error exporting data: {e}", exc_info=True)
        print(f"\n❌ Error exporting data: {e}")
        print(f"📋 Check log for details: {get_log_file_path()}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
