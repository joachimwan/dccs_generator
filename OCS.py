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
# - 'XX' unit/day 'daily/weekly/monthly' from 'start date/Phase' to 'end date/Phase'
# - 1 unit/day recur daily from start phase 50 to end phase 50 for minimum 7 days
# - 1 unit/day recur weekly on Monday Wednesday from start date 20238/01/01 to end date 2024/01/01 for maximum 10 days
# - 1 unit/day recur monthly on day 3 from start date 2023/01/01 to end date 2024/01/01
# - 1.2 unit/day on start date 2023/01/03
# - 0.5 unit/day on end phase 15

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


def create_instruction_dict(text):
    instruction_dict = {}
    text = text.split()
    instruction_dict['Number'] = text[0]
    if text[2] == 'recur':
        instruction_dict['Recurrence'] = text[2] + " " + text[3]
        if text[3] == 'weekly':
            days_of_week = {'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'}
            instruction_dict['Setting'] = set(days_of_week.intersection(text))
        elif text[3] == 'monthly':
            instruction_dict['Setting'] = text[6]
        else:  # Recur daily.
            instruction_dict['Setting'] = None
        from_index = text.index('from')
        instruction_dict['Start Type'] = text[from_index+1] + " " + text[from_index+2]
        instruction_dict['Start'] = text[from_index+3]
        instruction_dict['End Type'] = text[from_index+5] + " " + text[from_index+6]
        instruction_dict['End'] = text[from_index+7]
        try:
            instruction_dict['Min Max'] = text[from_index+9]
            instruction_dict['Occurrence'] = text[from_index+10]
        except Exception as e:
            print("Error:", e)
    else:  # No recurrence.
        instruction_dict['Start Type'] = text[3] + " " + text[4]
        instruction_dict['Start'] = text[5]
    return instruction_dict


# Load all OCS into a dataframe. Sorted by filename.
df_OCS = pd.concat([read_OCS(f) for f in sorted(OCS_DIR.iterdir(), key=lambda x: x.name)], ignore_index=True)

# Generate placeholder columns.
df_OCS['Total Cost (USD)'] = None
df_OCS['Total Units'] = None
