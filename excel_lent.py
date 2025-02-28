import pandas as pd
import subprocess
import time
import os
from typing import Dict
from pathlib import Path


def create_sample_dataframe():
    """
    Generate a sample Pandas DataFrame for testing.

    Returns:
        pd.DataFrame: A DataFrame with various data types.
    """

    data = {
        'NumericColumn1': [1, 2, 3, 4, 5],
        'NumericColumn2': [10.5, 20.2, 30.8, 40.1, 50.6],
        'StringColumn': ['A', 'B', 'C', 'D', 'E'],
        'MixedColumn': [1, 'A', 2.5, 'B', 3]
    }

    df = pd.DataFrame(data)
    return df

def test_libreoffice_headless(soffice_path):
    """
    Test launching LibreOffice Calc in headless mode.

    Args:
        soffice_path (str): The path to the LibreOffice Calc executable.

    """
    try:
        # Start LibreOffice in headless mode
        process = subprocess.Popen([soffice_path, '--headless', '--calc'])

        # Wait for 5 secs
        time.sleep(2)

        # Check if process is still running
        if process.poll() is None:
            # If process is still running, terminate
            process.terminate()
            process.wait()
            print("LibreOffice headless test: Success! (process terminated)")
        else:
            print("LibreOffice headless test: Failed! (process exited immediately)")

    except FileNotFoundError:
        print(f"LibreOffice headless test: Failed! (soffice not found)")

    except Exception as e:
        print(f"LibreOffice headless test: Failed! (An unexpected error occurred: {e})")

class ExcelLent:
    """
    Processes Pandas DataFrames, creates Excel files using LibreOffice Calc, and calculates column sums.
    """

    def __init__(self, soffice_path: str):
        """
        Initializes the ExcelLent class.

        Args:
            soffice_path (str): The path to the LibreOffice Calc executable.
        """
        self.soffice_path = soffice_path

    def process_dataframe(self, df: pd.DataFrame, filename: str) -> Dict[str, float]:
        """
        Processes a Pandas DataFrame, creates an Excel file using LibreOffice Calc, and calculates column sums.

        Args:
            df (pd.DataFrame): The DataFrame to process.
            filename (str): The name of the Excel file to create.

        Returns:
            dict[str, float]: A dictionary mapping column names to their sums.
        """

        # Placeholder for column sums (will implement later)
        column_sums: Dict[str, float] = {}

        try:
            # Create temporary csv file
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            temp_csv_filepath = Path(f"temp_{timestamp}.csv")
            df.to_csv(temp_csv_filepath, index=False)

            # Construct the command
            output_filepath = Path(f"{filename}_{timestamp}.xlsx")

            command = [
                self.soffice_path,
                "--headless",
                "--convert-to",
                "xlsx",
                str(temp_csv_filepath),
                "--outdir",
                ".",
            ]

            # Execute the command via subprocess
            process = subprocess.Popen(
                command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
            )

            # Write csv data to stdin
            stdout, stderr = process.communicate()

            # Error handling
            if process.returncode != 0:
                raise Exception(f"Calc process failed with error: {stderr}")

            # Clean up temporary file
            os.remove(temp_csv_filepath)

            print(f"Sucessfully created {output_filepath}")

        except Exception as e:
            print(f"An error occured: {e}")

        return column_sums


if __name__ == "__main__":
    df = create_sample_dataframe()
    # print(df)

    soffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    # test_libreoffice_headless(soffice_path)

    # Create an instance of the ExcelLent class
    excel_lent = ExcelLent(soffice_path)

    # Filename for the output Excel file
    output_filename = "output"

    result = excel_lent.process_dataframe(df, output_filename)
