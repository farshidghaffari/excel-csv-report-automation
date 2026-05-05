# Excel / CSV Report Automation with Python

A practical Python automation project for cleaning Excel/CSV files and generating business-ready summary reports.

This project is designed as a portfolio-ready example for freelance clients who need:
- Excel and CSV report automation
- Data cleaning and validation
- Monthly sales summaries
- Category-level reporting
- Reusable Python scripts for repetitive reporting tasks

## What This Project Does

The script reads a CSV or Excel file, cleans the data, validates the required columns, and generates an Excel report with multiple sheets:

1. **Cleaned Data** — normalized and cleaned rows
2. **Monthly Summary** — revenue grouped by month
3. **Category Summary** — revenue grouped by category
4. **Product Summary** — revenue grouped by product
5. **Report Info** — basic report metadata

## Folder Structure

```text
excel-csv-report-automation/
├── README.md
├── requirements.txt
├── .gitignore
├── LICENSE
├── sample_data/
│   └── sales_sample.csv
├── src/
│   └── report_automation/
│       ├── __init__.py
│       └── generate_report.py
├── examples/
│   └── run_demo.py
├── tests/
│   └── test_generate_report.py
├── docs/
│   ├── project-overview.md
│   └── sample-output.md
├── output/
│   └── .gitkeep
└── site-snippets/
    └── projects-card.html
```

## Requirements

- Python 3.10+
- pandas
- openpyxl
- pytest

Install dependencies:

```bash
pip install -r requirements.txt
```

## Run Demo

```bash
python examples/run_demo.py
```

The generated Excel report will be saved in:

```text
output/sales_report.xlsx
```

## Input CSV Format

The input file should include these columns:

```text
date, product, category, quantity, unit_price
```

Example:

```csv
date,product,category,quantity,unit_price
2026-01-05,USB Cable,Accessories,5,8.5
2026-01-07,Wireless Mouse,Accessories,2,24.9
```

## Use Cases

This project can be adapted for:

- Sales reports
- Inventory reports
- Finance summaries
- Client order reports
- Marketing campaign reports
- CSV cleanup workflows
- Weekly/monthly business reporting

## Freelance Service Angle

This repository demonstrates a real service I can provide:

> I can automate repetitive Excel and CSV reports using Python, clean raw business data, and generate structured reports that save time and reduce manual errors.

## Related Portfolio Pages

- Project page: https://farshidghaffari.net/projects/
- Blog article: https://farshidghaffari.net/blog/automate-excel-csv-reports-python/
- Python services: https://farshidghaffari.net/services/

## Author

Farshid Ghaffari  
Python Developer — Automation, Backend APIs, Data Tools

Website: https://farshidghaffari.net  
GitHub: https://github.com/farshidghaffari
