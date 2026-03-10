# Excel Cleaner

A simple but robust Python script to automatically clean Excel files in
bulk.

This project was created to solve a very common problem in real-world
data work: **messy Excel files**.\
Many spreadsheets contain empty rows, empty columns, inconsistent column
names, invisible spaces, and duplicated records.

This script processes **every Excel file in the same directory**, cleans
the data across **all sheets**, and generates a cleaned version
automatically.

It works across **Windows, macOS, and Linux**.

------------------------------------------------------------------------

# What This Script Does

The script automatically:

-   Detects Excel files in the current folder
-   Processes **all sheets** in each Excel file
-   Cleans column names
-   Removes empty rows
-   Removes empty columns
-   Strips invisible spaces from text fields
-   Removes duplicated rows
-   Logs the cleaning process
-   Generates a clean Excel file as output

------------------------------------------------------------------------

# Supported Excel Formats

The script supports:

-   `.xlsx`
-   `.xls`
-   `.xlsm`

------------------------------------------------------------------------

# Example

### Before

    sales.xlsx
    customers.xlsx
    inventory.xlsx

### Run

    python clean_excel.py

### After

    clean_sales.xlsx
    clean_customers.xlsx
    clean_inventory.xlsx

Original files remain unchanged.

------------------------------------------------------------------------

# Project Structure

    excel-cleaner/
    │
    ├── clean_excel.py
    ├── sales.xlsx
    ├── customers.xlsx
    └── inventory.xlsx

The script scans the **current directory** and processes all Excel files
automatically.

------------------------------------------------------------------------

# Installation

## 1. Clone the repository

    git clone https://github.com/your-username/excel-cleaner.git
    cd excel-cleaner

## 2. Install dependencies

    pip install pandas openpyxl xlrd

------------------------------------------------------------------------

# Usage

1.  Place your Excel files in the same folder as the script.

Example:

    clean_excel.py
    sales.xlsx
    customers.xlsx
    report.xls

2.  Run the script:

```bash
    python clean_excel.py
```
    

3.  Cleaned files will be generated automatically:

```bash
    clean_sales.xlsx
    clean_customers.xlsx
    clean_report.xls
```


------------------------------------------------------------------------

# Data Cleaning Steps

For every sheet in every Excel file, the script performs the following:

### 1. Remove Empty Rows

Rows that contain no data are removed.

### 2. Remove Empty Columns

Columns that contain no values are removed.

### 3. Standardize Column Names

Column names are normalized:

Example:

    " Total Sales "
    "Customer ID"
    "Price-Unit"

becomes:

    total_sales
    customer_id
    price_unit

### 4. Trim Text Fields

Removes invisible leading and trailing spaces.

Example:

    "  John  "

becomes:

    "John"

### 5. Remove Duplicate Rows

Duplicate records are removed automatically.

------------------------------------------------------------------------

# Logging

The script prints logs during execution so you can see what is
happening.

Example output:

    Processing file: sales.xlsx
    Cleaning sheet: January
    Cleaning sheet: February
    File saved: clean_sales.xlsx | Rows removed: 42

This helps you understand what changes were made.

------------------------------------------------------------------------

# Why This Project Exists

Excel files are still one of the most common data sources in companies.

However, they often arrive:

-   poorly formatted
-   inconsistent
-   full of empty cells
-   duplicated
-   difficult to analyze

This script acts as a **lightweight data-cleaning step**, similar to a
small ETL stage, helping prepare spreadsheets before analysis or loading
into a data pipeline.

------------------------------------------------------------------------

# Requirements

-   Python 3.9+
-   pandas
-   openpyxl
-   xlrd

------------------------------------------------------------------------

# Future Improvements

Potential improvements:

-   CLI arguments (choose folder, output path)
-   CSV support
-   Data type inference
-   Data quality reports
-   Integration with data pipelines

------------------------------------------------------------------------

# License

MIT License

------------------------------------------------------------------------

# Author

Created as part of a series of practical automation scripts for
developers and data professionals.