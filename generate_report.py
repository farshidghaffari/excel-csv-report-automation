from __future__ import annotations

from pathlib import Path
from typing import Literal

import pandas as pd


RequiredColumns = {"date", "product", "category", "quantity", "unit_price"}


def _read_input_file(input_path: Path) -> pd.DataFrame:
    """Read a CSV or Excel file into a DataFrame."""
    suffix = input_path.suffix.lower()

    if suffix == ".csv":
        return pd.read_csv(input_path)

    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(input_path)

    raise ValueError("Unsupported file type. Please provide a .csv, .xlsx, or .xls file.")


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names to lowercase snake_case style."""
    df = df.copy()
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_", regex=False)
        .str.replace("-", "_", regex=False)
    )
    return df


def _validate_columns(df: pd.DataFrame) -> None:
    """Validate that all required columns exist."""
    missing_columns = RequiredColumns - set(df.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(f"Missing required columns: {missing}")


def _clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Clean and enrich the sales data."""
    df = df.copy()

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["product"] = df["product"].astype(str).str.strip()
    df["category"] = df["category"].astype(str).str.strip()
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").fillna(0).astype(int)
    df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce").fillna(0)

    df = df.dropna(subset=["date"])
    df = df[df["quantity"] > 0]
    df = df[df["unit_price"] >= 0]

    df["total_revenue"] = df["quantity"] * df["unit_price"]
    df["month"] = df["date"].dt.to_period("M").astype(str)

    return df.sort_values("date").reset_index(drop=True)


def _build_summary(df: pd.DataFrame, group_by: Literal["month", "category", "product"]) -> pd.DataFrame:
    """Build a grouped revenue summary."""
    summary = (
        df.groupby(group_by, as_index=False)
        .agg(
            total_quantity=("quantity", "sum"),
            total_revenue=("total_revenue", "sum"),
            average_unit_price=("unit_price", "mean"),
            order_count=("product", "count"),
        )
        .sort_values("total_revenue", ascending=False)
    )

    summary["total_revenue"] = summary["total_revenue"].round(2)
    summary["average_unit_price"] = summary["average_unit_price"].round(2)

    return summary


def generate_report(input_file: str | Path, output_file: str | Path) -> Path:
    """Generate an Excel report from a CSV or Excel input file.

    Args:
        input_file: Path to the source CSV/XLSX file.
        output_file: Path where the generated XLSX report should be saved.

    Returns:
        Path to the generated report file.
    """
    input_path = Path(input_file)
    output_path = Path(output_file)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    raw_df = _read_input_file(input_path)
    normalized_df = _normalize_columns(raw_df)
    _validate_columns(normalized_df)

    clean_df = _clean_data(normalized_df)

    monthly_summary = _build_summary(clean_df, "month")
    category_summary = _build_summary(clean_df, "category")
    product_summary = _build_summary(clean_df, "product")

    report_info = pd.DataFrame(
        {
            "metric": [
                "source_file",
                "total_rows_after_cleaning",
                "total_revenue",
                "date_from",
                "date_to",
            ],
            "value": [
                str(input_path),
                len(clean_df),
                round(clean_df["total_revenue"].sum(), 2),
                clean_df["date"].min().strftime("%Y-%m-%d") if not clean_df.empty else "",
                clean_df["date"].max().strftime("%Y-%m-%d") if not clean_df.empty else "",
            ],
        }
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        clean_df.to_excel(writer, sheet_name="Cleaned Data", index=False)
        monthly_summary.to_excel(writer, sheet_name="Monthly Summary", index=False)
        category_summary.to_excel(writer, sheet_name="Category Summary", index=False)
        product_summary.to_excel(writer, sheet_name="Product Summary", index=False)
        report_info.to_excel(writer, sheet_name="Report Info", index=False)

    return output_path
