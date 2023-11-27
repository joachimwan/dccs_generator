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
lookahead_wells = {well for well in df_lookahead['Well Name'].unique() if well is not None}

# Identify lookahead start time.
lookahead_start_time = df_lookahead['Start Time'].iloc[0]

# Recalculate projection start time.
df_lookahead['Projection Time'] = df_lookahead['Actual Time'].fillna(df_lookahead['AFE Time']).fillna(df_lookahead['DSV Time'])
df_lookahead['Cumulative Projection Time'] = df_lookahead['Projection Time'].cumsum().shift(fill_value=0)
df_lookahead['Projection Start Time'] = df_lookahead['Cumulative Projection Time'].apply(lambda x: lookahead_start_time + pd.Timedelta(hours=x))
df_lookahead['Projection End Time'] = df_lookahead['Projection Start Time'] + pd.to_timedelta(df_lookahead['Projection Time'], unit='hours')

# Identify projected lookahead end time.
lookahead_end_time = df_lookahead['Projection End Time'].iloc[-1]

# Generate performance tracker grouped by well.
grouped_df = df_lookahead.groupby(['Well Name', 'Phase Code', 'Phase']).agg(
    Projection_Start_Time=('Projection Start Time', 'first'),
    Projection_End_Time=('Projection End Time', 'last'),
    AFE_Time=('AFE Time', 'sum'),
    Actual_Time=('Actual Time', 'sum'))
grouped_df = grouped_df.reset_index()
grouped_df = grouped_df.rename({'Projection_Start_Time': 'Projection Start Time',
                                'Projection_End_Time': 'Projection End Time',
                                'AFE_Time': 'AFE Time',
                                'Actual_Time': 'Actual Time'}, axis='columns')


# Calculate intersection of two datetime ranges in number of days.
def calc_intersection(start1, end1, start2, end2):
    start = max(start1, start2)
    end = min(end1, end2)
    return (end-start).total_seconds()/(60*60*24) if start < end else 0


# Helper function.
def calc_days_in_date(date):
    grouped_df['Days in Date'] = grouped_df.apply(lambda row: calc_intersection(row['Projection Start Time'], row['Projection End Time'], date, date+pd.Timedelta(days=1)), axis=1)
    for well in lookahead_wells:
        df_date_well.loc[df_date_well['Date'] == date, well] = grouped_df[grouped_df['Well Name'] == well]['Days in Date'].sum()
    grouped_df.drop(['Days in Date'], axis=1, inplace=True)


# Generate operational days by well.
df_date_well = pd.DataFrame({'Date': pd.date_range(start=lookahead_start_time.date(), end=lookahead_end_time.date())})
df_date_well.update({well: None} for well in lookahead_wells)
df_date_well.apply(lambda row: calc_days_in_date(date=row['Date']), axis=1)
