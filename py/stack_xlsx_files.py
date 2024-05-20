"""
Merges a set of excel files by concatenating rows. A column is added to indicate
the source file for each row of data. Sheets are preserved as well.

Dependencies:
pip3 install pandas
pip3 install openpyxl

To run the script:
python3 merge_xlsx.py
"""

import pandas as pd
import glob
from pathlib import Path

# file pattern to find the excel files to merge
top = "/home/zeke/Downloads/initial_surveys/*.xls*"
output = "Combined.xlsx"

# find the files to combine
files = glob.glob(top)
files = [Path(file) for file in files if Path(file).name != output]
print(f"Merging excel data from:\n{files}")

combined = {}  # Dict[sheet_name: pd.DataFrame] to store the combined data
for file in files:
  xl = pd.ExcelFile(file)
  for sheet in xl.sheet_names:
    data = xl.parse(sheet)
    data.insert(0, "source_file", file.name)  # keep track of source file
    if sheet not in combined:  # first file creates the pd.DataFrame
        combined[sheet] = data
    else:  # subsequent files are concatenated to the existing pd.DataFrame
       combined[sheet] = pd.concat([combined[sheet], data])

# create the output .xlsx file and dump each combined sheet into it
writer = pd.ExcelWriter(file.parent / output, engine = 'openpyxl')
for (sheet_name, sheet) in combined.items():
   sheet.to_excel(writer, sheet_name=sheet_name)
writer.close()