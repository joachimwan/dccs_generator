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
# - Identify lookahead start time.
# - Compute Projection Time based on Actual Time, then AFE Time, then DSV Time.
# - Recalculate Projection Start Time based on Projection Time.
# - For each well, generate Performance Tracker.

# Proposed verification:
# - Lookahead name, sheet name, and table name are as expected.
# - All expected Phase Codes and Phases for each well are present.
# - Raise warning if there are gaps in Actual Time.


def read_lookahead():
    excel_file_path = LATEST_LOOKAHEAD_DIR
    sheet_name = 'Drilling Input'
    table_name = 'LookaheadTable'
    wb = openpyxl.load_workbook(filename=excel_file_path, data_only=True)
    ws = wb[sheet_name]
    lookup_table = ws.tables[table_name]
    data = ws[lookup_table.ref]
    rows_list = [[col.value for col in row] for row in data]
    df = pd.DataFrame(data=rows_list[1:], index=None, columns=rows_list[0])
    return df


# Load workbook into dataframe.
df_lookahead = read_lookahead()

# Rename headers and remove unnecessary columns.
df_lookahead = df_lookahead.rename({'GK WELL OPERATIONS LOOKAHEAD': 'Description',
                                    'DWOP': 'AFE Time',
                                    'DSV plan': 'DSV Time',
                                    'Actual \nTime': 'Actual Time'}, axis='columns')
df_lookahead = df_lookahead[['Start Time', 'Well Name', 'Phase Code', 'Phase', 'Description', 'AFE Time', 'DSV Time',
                             'Actual Time']]

# Identify number of unique wells in the lookahead.
for well in df_lookahead['Well Name'].unique():
    if well is not None:
        pass

# Identify lookahead start time.
lookahead_start_time = df_lookahead['Start Time'].iloc[0]

# Recalculate projection start time.
df_lookahead['Projection Time'] = df_lookahead['Actual Time'].fillna(df_lookahead['AFE Time']).fillna(df_lookahead['DSV Time'])
df_lookahead['Cumulative Projection Time'] = df_lookahead['Projection Time'].cumsum().shift(fill_value=0)
df_lookahead['Projection Start Time'] = df_lookahead['Cumulative Projection Time'].apply(lambda x: lookahead_start_time + pd.Timedelta(hours=x))
df_lookahead['Projection End Time'] = df_lookahead['Projection Start Time'] + pd.to_timedelta(df_lookahead['Projection Time'], unit='hours')
