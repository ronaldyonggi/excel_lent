# excel_lent/core.py
import pandas as pd
import time
import os
from typing import Dict
from pathlib import Path
from .config import SOFFICE_PATH  # Import from config.py
from .worker import spawn_calc_worker  # Import the worker function
import concurrent.futures
import psutil
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class ExcelLent:
    """
    A class to process Pandas DataFrames and interact with LibreOffice Calc.
    """

    def __init__(self, output_dir: str = '.'):
        """
        Initializes the ExcelLent class.

        Args:
            output_dir (str): The directory to save output files. Defaults to current directory.
        """
        self.output_dir = output_dir
        # Ensure the output directory exists
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)

    def process_dataframe(self, df: pd.DataFrame, filename: str) -> Dict[str, float]:
        """
        Processes a Pandas DataFrame, creates an Excel file, and calculates column sums.

        Args:
            df (pd.DataFrame): The input DataFrame.
            filename (str): The base filename for the Excel file (without extension).

        Returns:
            Dict[str, float]: A dictionary of column names and their sums.
        """
        column_sums: Dict[str, float] = {}

        try:
            # Create a ProcessPoolExecutor to manage worker processes
            with concurrent.futures.ProcessPoolExecutor() as executor:
                futures = {}
                # Iterate over each column and submit it to a worker
                for column_name in df.columns:
                    if pd.api.types.is_numeric_dtype(df[column_name]):
                        future = executor.submit(spawn_calc_worker, df[column_name], column_name, self.output_dir)
                        futures[future] = column_name

                # Collect results as they become available
                for future in concurrent.futures.as_completed(futures):
                    column_name = futures[future]
                    try:
                        column_sum = future.result()
                        column_sums[column_name] = column_sum
                    except Exception as e:
                        logging.error(f"Error processing column {column_name}: {e}")

            # Log resource usage
            self._log_resource_usage()

            # Create a combined Excel file
            self._save_combined_excel(df, filename, column_sums)

        except Exception as e:
            logging.error(f"An error occurred: {e}")

        return column_sums

    def _save_combined_excel(self, df: pd.DataFrame, filename: str, column_sums: Dict[str, float]):
        """
        Saves the combined Excel file with the original data and column sums.

        Args:
            df (pd.DataFrame): The input DataFrame.
            filename (str): The base filename for the Excel file (without extension).
            column_sums (Dict[str, float]): A dictionary of column names and their sums.
        """
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        output_filepath = Path(self.output_dir) / f"{filename}_{timestamp}.xlsx"

        with pd.ExcelWriter(output_filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Data', index=False)
            sums_df = pd.DataFrame(list(column_sums.items()), columns=['Column', 'Sum'])
            sums_df.to_excel(writer, sheet_name='Sums', index=False)

        logging.info(f"Successfully created combined Excel file at {output_filepath}")

    def _log_resource_usage(self):
        """
        Logs the resource usage of the current process.
        """
        process = psutil.Process(os.getpid())
        memory_info = process.memory_info()
        cpu_percent = process.cpu_percent(interval=1)
        logging.info(f"Memory Usage: {memory_info.rss} bytes")
        logging.info(f"CPU Usage: {cpu_percent}%")