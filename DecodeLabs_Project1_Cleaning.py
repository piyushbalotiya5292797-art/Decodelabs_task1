"""
DecodeLabs Project 1: Data Cleaning & Preparation
Uses the dataset provided for the task.
Only performs checks and date-format normalization required by the brief.
It does NOT fill missing values, rewrite text, alter prices, or delete rows unless duplicate OrderIDs exist.
"""

import pandas as pd
from pathlib import Path

INPUT_FILE = Path("Dataset_for_Data_Analytics.xlsx")
OUTPUT_XLSX = Path("DecodeLabs_Project1_Cleaned_Dataset.xlsx")
OUTPUT_CSV = Path("DecodeLabs_Project1_Cleaned_Dataset.csv")
SUMMARY_FILE = Path("DecodeLabs_Project1_Cleaning_Summary.txt")

df = pd.read_excel(INPUT_FILE)
original_rows = len(df)
original_columns = len(df.columns)

# 1. Identify missing/null values. Do not replace them without a reliable source.
missing_before = df.isna().sum()

# 2. Check duplicate OrderIDs. Remove later occurrences only if duplicates are present.
duplicate_id_rows = int(df["OrderID"].duplicated(keep="first").sum())
if duplicate_id_rows:
    df = df.drop_duplicates(subset=["OrderID"], keep="first").copy()

# 3. Validate and standardize date format. Invalid dates are not guessed.
parsed_dates = pd.to_datetime(df["Date"], errors="coerce")
invalid_dates = int(parsed_dates.isna().sum())
df["Date"] = parsed_dates

# 4. Check numeric columns are valid numeric data; do not alter valid values.
numeric_columns = ["Quantity", "UnitPrice", "ItemsInCart", "TotalPrice"]
numeric_invalid = {}
for column in numeric_columns:
    converted = pd.to_numeric(df[column], errors="coerce")
    numeric_invalid[column] = int(converted.isna().sum())
    if not pd.api.types.is_numeric_dtype(df[column]):
        df[column] = converted

# 5. Check text fields for leading/trailing spaces; do not rewrite text values.
text_columns = df.select_dtypes(include=["object", "string"]).columns
space_issues = {
    column: int((df[column].dropna().astype(str) != df[column].dropna().astype(str).str.strip()).sum())
    for column in text_columns
}

# 6. Final validation
duplicate_ids_remaining = int(df["OrderID"].duplicated().sum())
invalid_dates_remaining = int(pd.to_datetime(df["Date"], errors="coerce").isna().sum())
missing_after = df.isna().sum()

# Export Excel with date display format YYYY-MM-DD.
with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl", date_format="yyyy-mm-dd",
                    datetime_format="yyyy-mm-dd") as writer:
    df.to_excel(writer, index=False, sheet_name="CleanedData")
    ws = writer.sheets["CleanedData"]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    date_column = list(df.columns).index("Date") + 1
    for row in range(2, ws.max_row + 1):
        ws.cell(row=row, column=date_column).number_format = "yyyy-mm-dd"

# CSV has dates explicitly formatted as YYYY-MM-DD.
csv_df = df.copy()
csv_df["Date"] = csv_df["Date"].dt.strftime("%Y-%m-%d")
csv_df.to_csv(OUTPUT_CSV, index=False, na_rep="")

numeric_lines = "\n".join(
    f"- {column}: {df[column].dtype}; values failing numeric check: {numeric_invalid[column]}"
    for column in numeric_columns
)
space_lines = "\n".join(f"- {column}: {count} issue(s)" for column, count in space_issues.items())

summary = f"""DECODELABS PROJECT 1 — CLEANING SUMMARY
Source: Dataset_for_Data_Analytics.xlsx (dataset provided for the task)

Rows in source: {original_rows}
Columns in source: {original_columns}
Rows in output: {len(df)}
Duplicate OrderID rows removed: {duplicate_id_rows}
Duplicate OrderIDs remaining: {duplicate_ids_remaining}
Invalid/missing dates remaining: {invalid_dates_remaining}

Missing/null values in source:
{missing_before.to_string()}

Missing/null values in output:
{missing_after.to_string()}

Numeric columns checked:
{numeric_lines}

Text fields checked for leading/trailing spaces:
{space_lines}

Actions taken:
- Checked missing/null values and reported them; no values were invented or filled.
- Checked OrderID duplicates; later occurrences would be removed only if duplicates existed.
- Checked dates and standardized their output display/export format to YYYY-MM-DD.
- Checked numeric columns without changing valid numeric values.
- Checked text fields; did not rewrite text content.
- Preserved all columns and valid source values.

Important:
Missing CouponCode values remain missing because the source does not specify their correct values.
"""
SUMMARY_FILE.write_text(summary, encoding="utf-8")
print(summary)
