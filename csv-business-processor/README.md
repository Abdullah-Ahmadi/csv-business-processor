CSV Business Processor

CSV Business Processor is a lightweight, extensible CLI-driven toolkit for transforming and exporting CSV datasets used by small business workflows (inventory, ecommerce, finance). It provides processors to interpret domain-specific CSV formats and multiple exporters (console, Excel, JSON, PDF) to produce human- and machine-friendly outputs.

## Key Features

- Process common business CSVs with domain-aware processors: `inventory`, `ecommerce`, and `finance`.
- Export results to multiple formats: console, Excel (`.xlsx`), JSON, and PDF.
- Run as a module using the included CLI (`csv_pro.cli`).
- Small, dependency-light codebase suitable for portfolio demonstration and extension.

## Quick Start

Prerequisites

- Python 3.8 or newer
- Install dependencies:

```bash
pip install -r requirements.txt
```

Run the inventory feature and print to console (example):

```bash
python -m csv_pro.cli sample_data/sample_inventory.csv --mode inventory --output console
```

Run and export to Excel:

```bash
python -m csv_pro.cli sample_data/sample_inventory.csv --mode inventory --output excel --out-file reports/inventory_report.xlsx
```

Available modes

- `inventory` — inventory count, stock replenishment hints, SKU lookups
- `ecommerce` — order and customer aggregates, revenue summaries
- `finance` — transaction categorization and basic P&L summaries

Available exporters

- `console` — human-readable text output (good for quick checks)
- `excel` — `.xlsx` workbook output for downstream analysis
- `json` — structured JSON suitable for APIs or further automation
- `pdf` — printable report output

## Examples and Screenshots

Console output example:

![Console Output](docs/console-output.png)

Excel output example (open in Excel or LibreOffice):

![Excel Output](docs/excel-output.png)

PDF report example:

![PDF Output](docs/pdf-output.png)

## Project Layout

- `csv_pro/cli.py` — module entrypoint and CLI interface
- `csv_pro/processors/` — domain processors (`inventory.py`, `ecommerce.py`, `finance.py`)
- `csv_pro/exporters/` — exporter implementations (console, Excel, JSON, PDF)
- `csv_pro/utils/` — helper utilities (file handling, formatting, logging)
- `sample_data/` — example CSV files to try out

## Extending the Project

To add a new processor or exporter, follow existing patterns in `csv_pro/processors` and `csv_pro/exporters`. Keep single-responsibility classes and add unit tests where appropriate.

## Contributing

Contributions, improvements, and bug reports are welcome. Open an issue or submit a pull request on the repository.

## License

This project includes a `LICENSE` file in the repository root.

## Author

Abdulla Ahmadi — https://github.com/Abdullah-Ahmadi

Project: CSV Business Processor
