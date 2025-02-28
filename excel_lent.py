import pandas as pd
import subprocess
import time

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
if __name__ == '__main__':
    df = create_sample_dataframe()
    print(df)