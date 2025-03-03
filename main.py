# main.py
from excel_lent.core import ExcelLent
from excel_lent.utils import create_sample_dataframe

if __name__ == '__main__':
    df = create_sample_dataframe()
    excel_lent = ExcelLent(output_dir='output')  # Specify an output directory
    output_filename = "output"
    result = excel_lent.process_dataframe(df, output_filename)
    print(f"Column sums: {result}")