# DecodeLabs Internship — Task 1: Data Cleaning & Preparation

## About the Project

This project is part of my Data Analytics internship at DecodeLabs. The main goal of this task is to clean and prepare a dataset so that it can be used for further analysis.

I used Python and Pandas to check the dataset, verify data formats, and make sure the data was ready for use without making unnecessary changes to the original information.

## Objectives

- Identify missing or null values in the dataset.
- Check for duplicate records.
- Verify and standardize date formats.
- Check numeric and text columns for formatting issues.
- Validate the dataset after cleaning.

## Tools and Technologies

- **Python** — for data cleaning and processing.
- **Pandas** — for handling and analyzing tabular data.
- **Excel** — for working with the original and cleaned datasets.
- **GitHub** — for storing and managing the project files.

## Dataset Overview

The dataset contains 1,200 rows and 14 columns related to order information, including order IDs, dates, customer IDs, products, quantities, prices, payment methods, order statuses, and coupon codes.

## Data Cleaning Process

The following checks were performed:

1. **Missing Values:** Identified missing values in the dataset. Missing coupon codes were preserved rather than filled with guessed values.
2. **Duplicate Records:** Checked the `OrderID` column for duplicate IDs.
3. **Date Formatting:** Verified the dates and standardized their display to `YYYY-MM-DD`.
4. **Numeric Data:** Checked columns such as quantity, unit price, total price, and items in cart.
5. **Text Formatting:** Checked text fields for unnecessary leading or trailing spaces.
6. **Validation:** Checked the dataset after processing to confirm that the required conditions were met.

## Results

- Total records retained: 1,200
- Total columns retained: 14
- Duplicate OrderIDs: 0
- Invalid or missing dates: 0
- Missing coupon codes: 309, preserved as missing values
- Original dataset: Kept unchanged

## Project Files

- `Dataset_for_Data_Analytics.xlsx` — Original dataset.
- `DecodeLabs_Project1_Cleaned_Dataset.xlsx` — Cleaned dataset in Excel format.
- `DecodeLabs_Project1_Cleaned_Dataset.csv` — Cleaned dataset in CSV format.
- `DecodeLabs_Project1_Cleaning.py` — Python script used for the cleaning process.
- `DecodeLabs_Project1_Cleaning_Summary.txt` — Summary of the cleaning checks and results.
- `DecodeLabs_Project1_Report.docx` — Project report.

## How to Run the Project

1. Install Python on your computer.
2. Install the required libraries:

   ```bash
   py -m pip install pandas openpyxl
   ```

3. Keep the Python script and original Excel dataset in the same folder.
4. Open Command Prompt or a terminal in that folder.
5. Run the script:

   ```bash
   py DecodeLabs_Project1_Cleaning.py
   ```

6. Check the generated cleaned dataset and cleaning summary.

## Conclusion

This task helped me understand the importance of data cleaning and preparation before performing data analysis. It also gave me practical experience using Python and Pandas to validate a dataset while preserving the original information.

---

**Internship:** DecodeLabs — Data Analytics  
**Project:** Task 1 — Data Cleaning & Preparation
