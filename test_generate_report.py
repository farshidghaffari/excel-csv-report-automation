from pathlib import Path
import sys

import pandas as pd
import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from report_automation.generate_report import generate_report


def test_generate_report_creates_excel_file(tmp_path: Path) -> None:
    input_file = tmp_path / "input.csv"
    output_file = tmp_path / "report.xlsx"

    input_file.write_text(
        "date,product,category,quantity,unit_price\n"
        "2026-01-01,USB Cable,Accessories,2,8.5\n"
        "2026-01-02,Keyboard,Hardware,1,45\n",
        encoding="utf-8",
    )

    result = generate_report(input_file=input_file, output_file=output_file)

    assert result.exists()
    assert result.suffix == ".xlsx"


def test_generate_report_rejects_missing_columns(tmp_path: Path) -> None:
    input_file = tmp_path / "bad_input.csv"
    output_file = tmp_path / "report.xlsx"

    input_file.write_text(
        "date,product,quantity\n"
        "2026-01-01,USB Cable,2\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Missing required columns"):
        generate_report(input_file=input_file, output_file=output_file)
