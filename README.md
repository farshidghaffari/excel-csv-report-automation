# Excel / CSV Reporting Automation

[![Tests](https://github.com/farshidghaffari/excel-csv-report-automation/actions/workflows/tests.yml/badge.svg)](https://github.com/farshidghaffari/excel-csv-report-automation/actions/workflows/tests.yml)

**Supporting implementation · Data processing and reporting · Python / Pandas / OpenPyXL**

A reusable data-to-report workflow that validates structured CSV or Excel input, cleans business data, calculates revenue metrics, and produces a multi-sheet Excel report.

## Business Problem

Recurring business reports are often assembled through repeated spreadsheet work: importing files, renaming columns, correcting values, calculating totals, grouping data, and rebuilding the same summary tabs every week or month.

That workflow is slow, difficult to reproduce consistently, and vulnerable to manual errors. This implementation demonstrates how the repeatable portion can be turned into a clear processing pipeline with an explicit input schema and predictable output.

## Implemented Workflow

```mermaid
flowchart LR
    A["CSV / Excel input"] --> B["Schema validation"]
    B --> C["Cleaning and normalization"]
    C --> D["Revenue calculations"]
    D --> E["Multi-sheet Excel report"]
```

The current implementation:

- Reads `.csv` and `.xlsx` files; legacy `.xls` input requires a compatible Pandas engine installed separately
- Normalizes column names to a consistent lowercase format
- Verifies that all required columns are present
- Converts dates, quantities, and prices to usable data types
- Excludes rows with invalid dates, non-positive quantities, or negative prices
- Calculates revenue for each valid row
- Builds monthly, category, and product summaries
- Creates the output directory when it does not exist
- Exports a structured Excel workbook with five worksheets

## Required Input Schema

| Column | Purpose |
|---|---|
| `date` | Transaction or order date |
| `product` | Product name |
| `category` | Reporting category |
| `quantity` | Number of units |
| `unit_price` | Price per unit |

Example input:

```csv
date,product,category,quantity,unit_price
2026-01-05,USB Cable,Accessories,5,8.5
2026-01-07,Wireless Mouse,Accessories,2,24.9
```

If a required column is missing, the workflow stops with a clear `ValueError` instead of producing an incomplete report.

## Generated Workbook

| Worksheet | Contents |
|---|---|
| `Cleaned Data` | Valid rows with normalized values, calculated revenue, and reporting month |
| `Monthly Summary` | Quantity, revenue, average unit price, and order count by month |
| `Category Summary` | The same metrics grouped by category |
| `Product Summary` | The same metrics grouped by product |
| `Report Info` | Source file, processed row count, total revenue, and date range |

## Quick Start

### 1. Clone the repository

```bash
git clone https://github.com/farshidghaffari/excel-csv-report-automation.git
cd excel-csv-report-automation
```

### 2. Create and activate a virtual environment

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
```

macOS / Linux:

```bash
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Run the included demonstration

```bash
python examples/run_demo.py
```

The generated workbook will be saved to:

```text
output/sales_report.xlsx
```

## Testing

Run the automated tests with:

```bash
pytest -q
```

The current test suite verifies two critical behaviors:

- A valid source file produces an `.xlsx` report.
- A source missing required columns is rejected with a clear error.

## Reliability Decisions

- **Schema validation comes before reporting.** A report is not generated when required fields are absent.
- **Unsupported formats fail explicitly.** Only the documented CSV and Excel formats are accepted.
- **Source data is not modified in place.** The workflow writes a separate report file.
- **Report structure is deterministic.** The same valid schema produces the same set of worksheets and summary dimensions.
- **Empty datasets remain safe.** Report metadata handles an empty cleaned result without failing on the date range.

## Current Boundaries

This is a focused supporting implementation rather than a complete reporting platform.

- Summary dimensions and required columns are currently fixed in code.
- Invalid rows are excluded but are not yet exported to a separate rejection report.
- The workflow runs locally and does not yet include scheduling, notifications, or a web interface.
- The current automated tests cover report creation and required-column failure; broader edge-case coverage is planned as the implementation evolves.

These boundaries are kept explicit so the repository demonstrates implemented behavior without presenting planned features as completed work.

## Adaptation Paths

The same workflow pattern can be adapted for:

- Sales and revenue reporting
- Inventory and purchasing summaries
- Client order exports
- Finance and operational reports
- Data preparation before dashboard ingestion
- Scheduled weekly or monthly reporting

## Repository Structure

```text
excel-csv-report-automation/
├── examples/
│   └── run_demo.py
├── sample_data/
│   └── sales_sample.csv
├── src/report_automation/
│   ├── __init__.py
│   └── generate_report.py
├── tests/
│   └── test_generate_report.py
├── docs/
├── requirements.txt
└── README.md
```

## Related Links

- [Portfolio projects](https://farshidghaffari.net/projects/)
- [Reporting automation insight](https://farshidghaffari.net/insights/reliable-reporting-automation/)
- [Discuss a reporting workflow](https://farshidghaffari.net/contact/)

## Author

**Farshid Ghaffari**  
Business Automation & Integration Specialist  
AI Workflows · APIs · Backend Systems
