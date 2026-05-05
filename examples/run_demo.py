from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from report_automation import generate_report


def main() -> None:
    input_file = PROJECT_ROOT / "sample_data" / "sales_sample.csv"
    output_file = PROJECT_ROOT / "output" / "sales_report.xlsx"

    report_path = generate_report(input_file=input_file, output_file=output_file)
    print(f"Report generated successfully: {report_path}")


if __name__ == "__main__":
    main()
