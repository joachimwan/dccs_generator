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
# - Well Name (can be 'Not assigned')
# - WBS Number (can be 'Not assigned')

# Charging mechanism:
# - 'XXX' unit/day from 'start date to end date' or for [list of 'Well-Phase'] for max XXX occurrences
# - 'XXX' unit/day for [list of 'Well-Phase'] for max XXX occurrences
# - 'XXX' unit/day on 'date'
# - 'XXX' unit/day on start/end of 'Well Name' phase 'Phase Code'
# - Rig rates: Charge 1 daily from 'start date to end date' or for [list of 'Well-Phase'] (may need Manual input)
# - Tariffs (e.g. ROE, Overhead): Charge 1 daily for [list of 'Well-Phase']
# - Tariffs (e.g. RTOC, LMP, Vessels): Charge 1 daily from 'start date' to 'end date'
# - Tariffs (e.g. Insurance): Lump sum charges
# - Aviation: Charge XXX on chopper days (or Manual input)
# - Consolidation for each OCS/Tariff: Manual input
# - Daily mud cost: Manual input
# - Lump sum charges: Charge XXX on 'start Well-Phase'
# - Installed equipment: Charge XXX on 'end Well-Phase'
# - Personnel and equipment rental: Charge XXX daily for [list of 'Well-Phase'] for max XXX occurrences
# - Some are specific to Well-Phase e.g. DD or TRS, some are continuous e.g. Mud logging or SCE rentals

import pandas as pd
import openpyxl
from settings import *

# Proposed workflow:
# - For each OCS file in the OCS folder, read each OCS file.
# - Use try-except to verify OCS validity and raise errors.
# - Detect revisions and ensure revision does not impact charged items before Today.
# - Parse information from charging mechanisms.
# - Generate OCS Number against WBS Number against Well Name.

# Proposed verification:
# - All required fields are present, e.g. OCS Number, WBS Number.
# - WBS Number and Well Name (if present) are correct.
# - All OCS has unique OCS Number.


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
        wb.close()
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
