import pandas as pd
from pathlib import Path



def merge_csv_to_excel(csv_files, sheet_names, sep=';', output='merged.xlsx'):
    """ Merge `csv_files` into a single excel file.

        Parameters:
        -----------
        csv_files: List of Path
            Paths to the csv files to be merged.
        sheet_names: List of str
            Names of excel worksheets corresponding to the `csv_files`
        sep: str
            Delimeter of the csv file.
        output: Path or str
            Output of the merged excel file.
    """
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    for csv, sheet_name in zip(csv_files, sheet_names):
        df = pd.read_csv(csv, sep=sep)
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.save()


if __name__ == "__main__":
    PATH = Path()
    DATA_PATH = PATH/'test_data'
    CSVs = list(DATA_PATH.glob('*.csv'))
    names = [f.stem for f in CSVs]
    merge_csv_to_excel(CSVs, names, sep=';')
