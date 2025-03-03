# tests/test_utils.py
import unittest
import pandas as pd
from excel_lent.utils import create_sample_dataframe

class TestUtils(unittest.TestCase):

    def test_create_sample_dataframe(self):
        # Create a sample DataFrame
        df = create_sample_dataframe()

        # Check if the DataFrame has the correct columns
        expected_columns = ['NumericColumn1', 'NumericColumn2', 'StringColumn', 'MixedColumn']
        self.assertListEqual(list(df.columns), expected_columns)

        # Check if the DataFrame has the correct number of rows
        self.assertEqual(len(df), 5)

        # Check if the DataFrame has the correct data types
        self.assertTrue(pd.api.types.is_numeric_dtype(df['NumericColumn1']))
        self.assertTrue(pd.api.types.is_numeric_dtype(df['NumericColumn2']))
        self.assertFalse(pd.api.types.is_numeric_dtype(df['StringColumn']))
        self.assertFalse(pd.api.types.is_numeric_dtype(df['MixedColumn']))

if __name__ == '__main__':
    unittest.main()