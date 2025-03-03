import pandas as pd

def create_sample_dataframe():
    """
    Generates a sample Pandas DataFrame for testing.

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