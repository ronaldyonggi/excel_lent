# tests/test_core.py
import unittest
import pandas as pd
import os
import shutil
from pathlib import Path
from excel_lent.core import ExcelLent
from excel_lent.utils import create_sample_dataframe

class TestExcelLent(unittest.TestCase):

    def setUp(self):
        # Create a temporary output directory for tests
        self.output_dir = Path('tests/output')
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def tearDown(self):
        # Clean up the temporary output directory after each test
        if self.output_dir.exists():
            shutil.rmtree(self.output_dir)

    def test_process_dataframe(self):
        # Create a sample DataFrame
        df = create_sample_dataframe()
        excel_lent = ExcelLent(output_dir=self.output_dir)
        filename = "test_output"

        # Process the DataFrame
        result = excel_lent.process_dataframe(df, filename)

        # Expected column sums
        expected_sums = {
            'NumericColumn1': 15.0,
            'NumericColumn2': 151.2
        }

        # Check if the result matches the expected sums
        self.assertEqual(result, expected_sums)

        # Check if the combined Excel file was created
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        expected_file_path = self.output_dir / f"{filename}_{timestamp}.xlsx"
        self.assertTrue(expected_file_path.exists())

    def test_non_numeric_columns(self):
        # Create a DataFrame with only non-numeric columns
        df = pd.DataFrame({
            'StringColumn': ['A', 'B', 'C'],
            'MixedColumn': [1, 'A', 2.5]
        })
        excel_lent = ExcelLent(output_dir=self.output_dir)
        filename = "test_non_numeric"

        # Process the DataFrame
        result = excel_lent.process_dataframe(df, filename)

        # Expected column sums (should be empty since no numeric columns)
        expected_sums = {}

        # Check if the result matches the expected sums
        self.assertEqual(result, expected_sums)

        # Check if the combined Excel file was created
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        expected_file_path = self.output_dir / f"{filename}_{timestamp}.xlsx"
        self.assertTrue(expected_file_path.exists())

    def test_empty_dataframe(self):
        # Create an empty DataFrame
        df = pd.DataFrame()
        excel_lent = ExcelLent(output_dir=self.output_dir)
        filename = "test_empty"

        # Process the DataFrame
        result = excel_lent.process_dataframe(df, filename)

        # Expected column sums (should be empty since no columns)
        expected_sums = {}

        # Check if the result matches the expected sums
        self.assertEqual(result, expected_sums)

        # Check if the combined Excel file was created
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        expected_file_path = self.output_dir / f"{filename}_{timestamp}.xlsx"
        self.assertTrue(expected_file_path.exists())

if __name__ == '__main__':
    unittest.main()