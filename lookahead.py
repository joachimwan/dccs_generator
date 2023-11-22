# Methods to read and verify Excel Lookahead.

# Lookahead headers:
# - Start Time
# - Well Name
# - Phase Code
# - Phase
# - Description
# - AFE Time
# - DSV Time
# - Actual Time

# Metadata:
# - Start Time (the very first start time of the lookahead)

import pandas as pd
import openpyxl
from settings import *

# Proposed workflow:
# - Identify latest lookahead. Ensure lookahead is a named table.
# - Use try-except to verify lookahead validity and raise errors.
# - Identify number of unique wells in the lookahead.
# - For lines after Actual, fill with AFE Time. For lines without AFE Time, fill with DSV Time.
# - Recalculate Start Time after Actual.
# - For each well, generate Performance Tracker.

# Proposed verification:
# - Lookahead name, sheet name, and table name are as expected.
# - All expected Phase Codes and Phases for each well are present.


def read_lookahead():
    excel_file_path = LATEST_LOOKAHEAD_DIR
    sheet_name = 'Drilling Input'
    table_name = 'LookaheadTable'
    wb = openpyxl.load_workbook(filename=excel_file_path, data_only=True)
    sheet = wb[sheet_name]
    lookup_table = sheet.tables[table_name]
    data = sheet[lookup_table.ref]
    rows_list = [[col.value for col in row] for row in data]
    df = pd.DataFrame(data=rows_list[1:], index=None, columns=rows_list[0])
    return df
