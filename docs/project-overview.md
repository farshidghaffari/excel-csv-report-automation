# Project Overview

A local Python workflow for repeatable CSV/Excel reporting: normalize headings, validate schema, split accepted and rejected rows, and produce a six-sheet workbook.

## Business problem and scope

Repeated spreadsheet preparation can hide invalid data and inconsistent calculations. This implementation makes its current row-level decisions visible; it does not correct invalid business records or provide a reporting service.

## Current pipeline

1. Read CSV, XLSX, or XLS using the existing Pandas readers.
2. Normalize headings and reject missing or colliding required schema.
3. Parse and validate the five required fields.
4. Normalize accepted records and retain rejected records with reasons.
5. Calculate accepted-row revenue and summaries by month, category, and product.
6. Export Cleaned Data, Rejected Rows, Monthly Summary, Category Summary, Product Summary, and Report Info.

See [README metric definitions](../README.md#metric-semantics-and-traceability) and the [synthetic example](sample-output.md). “Cleaning” here means normalization and separation, not automatic repair of invalid rows. `order_count` is an accepted-row count, not a unique-order count.

## Evidence and boundaries

The public implementation, existing tests/CI, and reproducible synthetic input demonstrate file processing and workbook generation. They do not establish production deployment, customer outcomes, API integration, scheduling, or accounting suitability. Source preservation requires distinct input/output paths; source-row numbers refer to parsed positions.
