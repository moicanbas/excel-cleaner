import pandas as pd
import re
from pathlib import Path
import logging

# ---------------------------------------
# Logging configuration
# ---------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(message)s"
)

# ---------------------------------------
# Get current directory
# ---------------------------------------

BASE_DIR = Path(__file__).resolve().parent


# ---------------------------------------
# Clean column names
# ---------------------------------------

def clean_column_names(columns):
    cleaned = []

    for col in columns:
        col = str(col).strip().lower()
        col = re.sub(r"[^\w\s]", "", col)
        col = re.sub(r"\s+", "_", col)

        cleaned.append(col)

    return cleaned


# ---------------------------------------
# Clean dataframe
# ---------------------------------------

def clean_dataframe(df):

    original_rows = len(df)

    # Remove completely empty rows
    df = df.dropna(how="all")

    # Remove completely empty columns
    df = df.dropna(axis=1, how="all")

    # Clean column names
    df.columns = clean_column_names(df.columns)

    # Trim string spaces
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Remove duplicated rows
    df = df.drop_duplicates()

    removed_rows = original_rows - len(df)

    return df, removed_rows


# ---------------------------------------
# Process Excel file
# ---------------------------------------

def process_excel(file_path):

    logging.info(f"\nProcessing file: {file_path.name}")

    try:

        # Read all sheets
        sheets = pd.read_excel(file_path, sheet_name=None)

        cleaned_sheets = {}

        total_removed = 0

        for sheet_name, df in sheets.items():

            logging.info(f"Cleaning sheet: {sheet_name}")

            cleaned_df, removed_rows = clean_dataframe(df)

            cleaned_sheets[sheet_name] = cleaned_df

            total_removed += removed_rows

        output_file = BASE_DIR / f"clean_{file_path.name}"

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

            for sheet_name, df in cleaned_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        logging.info(
            f"File saved: {output_file.name} | Rows removed: {total_removed}"
        )

    except Exception as e:

        logging.error(f"Error processing {file_path.name}: {e}")


# ---------------------------------------
# Main execution
# ---------------------------------------

def main():

    excel_files = list(BASE_DIR.glob("*.xlsx")) + \
                  list(BASE_DIR.glob("*.xls")) + \
                  list(BASE_DIR.glob("*.xlsm"))

    # Avoid reprocessing cleaned files
    excel_files = [f for f in excel_files if not f.name.startswith("clean_")]

    if not excel_files:
        logging.info("No Excel files found in this directory.")
        return

    for file in excel_files:
        process_excel(file)

    logging.info("\nDone.")


if __name__ == "__main__":
    main()