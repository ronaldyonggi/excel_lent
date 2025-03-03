# excel_lent/worker.py
import pandas as pd
from .config import SOFFICE_PATH
import subprocess
import os
from pathlib import Path
import time
import logging

def spawn_calc_worker(
    column_data: pd.Series, column_name: str, output_dir: str
) -> float:
    """
    Processes a single column of a DataFrame using headless LibreOffice Calc.

    Args:
        column_data (pd.Series): The data for the column.
        column_name (str): Name of the column.
        output_dir (str): The directory to save temporary files.

    Returns:
        float: The sum of the column data.
    """
    try:
        # Create temporary csv file for the single column
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        temp_csv_filepath = Path(output_dir) / f"temp_{column_name}_{timestamp}.csv"

        # Ensure the Series has a name before converting to CSV
        column_data.name = column_name
        column_data.to_csv(temp_csv_filepath, index=False, header=True)

        # Construct the Calc command
        output_filepath = Path(output_dir) / f"{column_name}_{timestamp}.xlsx"

        command = [
            SOFFICE_PATH,
            "--headless",
            "--convert-to",
            "xlsx",
            str(temp_csv_filepath),
            "--outdir",
            str(output_dir),  # Ensure output_dir is a string
            "--infilter=CSV:44,34,76,1,,true",  # Explicitly specify the first line is header
        ]

        # Start timing the process
        start_time = time.time()

        # Execute the command
        process = subprocess.Popen(
            command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
        )
        stdout, stderr = process.communicate()

        if process.returncode != 0:
            raise Exception(f"Calc process failed for column {column_name}: {stderr}")

        # Calculate the elapsed time
        elapsed_time = time.time() - start_time
        logging.info(f"Processing time for column {column_name}: {elapsed_time:.2f} seconds")

        # Clean up the temporary CSV file
        os.remove(temp_csv_filepath)

        # Calculate the sum directly from the column_data
        column_sum = column_data.sum()
        return column_sum

    except Exception as e:
        logging.error(f"Error processing column {column_name}: {e}")
        return 0.0  # Return 0 in case of error