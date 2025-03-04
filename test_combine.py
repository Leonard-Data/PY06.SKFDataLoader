import pandas as pd
import os
from utils.functions import *
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

try:
    folder_path = f"C:\\Users\\Public\\Downloads\\ACP.RefreshData_test"
    combined_file_path = os.path.join(folder_path, "Generate.xlsx")

    # Check if directory exists
    if not os.path.exists(folder_path):
        raise FileNotFoundError(f"Directory not found: {folder_path}")

    # Check if directory is empty
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls')) and f != "Generate.xlsx"]
    if not excel_files:
        raise ValueError(f"No Excel files found in: {folder_path}")

    # Try to combine files
    try:
        result = combine_excel_files(folder_path, combined_file_path)
        if result is not None:
            logging.info(f"Successfully combined Excel files into: {combined_file_path}")
        else:
            raise ValueError("No data was combined")

    except ValueError as e:
        logging.error(f"Data processing error: {str(e)}")
        if os.path.exists(combined_file_path):
            os.remove(combined_file_path)
            logging.info("Removed incomplete output file")

except FileNotFoundError as e:
    logging.error(f"Directory error: {e}")
except ValueError as e:
    logging.error(f"Validation error: {e}")
except Exception as e:
    logging.error(f"Unexpected error: {str(e)}")