# Methods to read and verify Excel OCS.

# OCS headers:
# - Description
# - SAP Element Number
# - Quantity
# - Unit of Measure
# - Estimated Duration
# - Currency
# - SAP Unit Price
# - Total Price

# Metadata:
# - Contract Title
# - Contract Number
# - SAP OA Number
# - OCS Number
# - Well Name
# - WBS Number

# Charging mechanism:
# - 'XX' unit/day 'daily/weekly/monthly' from 'start date/Phase' to 'end date/Phase' or 'XX occurrences'

import pandas as pd
import openpyxl
from settings import *

# Proposed workflow:
# - For each OCS file in the OCS folder, read each OCS file.
# - Use try-except to verify OCS validity and raise errors.
# - Detect revisions and ensure revision does not impact charged items before Today.
# -

# Proposed verification:
# - All required fields are present, e.g. OCS Number, WBS Number.
# - WBS Number and Well Name (if present) are correct.


def read_OCS(excel_file_path):
    try:
        sheet_name = 'OCS Input'
        table_name = 'OCSTable'
        wb = openpyxl.load_workbook(filename=excel_file_path, data_only=True)
        ws = wb[sheet_name]
        lookup_table = ws.tables[table_name]
        data = ws[lookup_table.ref]
        rows_list = [[col.value for col in row] for row in data]
        df = pd.DataFrame(data=rows_list[1:], index=None, columns=rows_list[0])
        df['Well Name'] = ws['B5'].value
        df['OCS Number'] = ws['B4'].value
        df['WBS Number'] = ws['B6'].value
        return df
    except Exception as e:
        print(f"Error on {excel_file_path} :", e)


# # Loop through all files in the folder and load workbooks into dataframes.
# df = pd.DataFrame()
# for f in OCS_DIR.iterdir():
#     data = read_OCS(f)
#     df = df.append(data)

# Load all OCS into a dataframe.
df_OCS = pd.concat([read_OCS(f) for f in sorted(OCS_DIR.iterdir(), key=lambda x: x.name)], ignore_index=True)

print(df_OCS)
