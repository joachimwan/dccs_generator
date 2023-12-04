# Methods to read, verify, and write from Excel DCCS.

# DCCS headers:
# - UID
# - Cost Group
# - Description
# - Allocation (Tariffs or OCS number)
# - MYR Cost
# - USD Cost
# - Unit
# - Total Cost (USD)
# - Total Units
# - Dates << to be associated with Well Name and Phase from Lookahead

# Metadata:
# - Well name
# - First day of DCCS
# - USD/MYR conversion rate

# Association headers:
# - OCS number
# - PO number
# - WBS number

import re
import pandas as pd
import openpyxl
from lookahead import *
from OCS import *

# Proposed workflow:
# - Identify latest DCCS.
# - Use try-except to verify lookahead validity and raise errors.
# - Parse information from charging mechanisms.
# -

# Proposed verification:
# -

# Proposed data sources:
# - The latest Excel Campaign Lookahead
# - All OCS and revisions
# - All Tariffs and consolidation
# - The latest DCCS (to check for manual inputs)
# - OpenWells for rig rates and NPT information
# - How about aviation charges...?


# Parse information from charging mechanisms.
def create_instruction_dict(text, well):
    instruction_dict = {}
    text_split = text.split()
    instruction_dict['Number'] = text_split[0]
    instruction_dict['Recurrence'] = text_split[2]
    if text_split[2] == 'from':
        if re.findall(r'\b\d{4}/\d{2}/\d{2}\b', text_split[3]):
            instruction_dict['Start'] = pd.to_datetime(text_split[3], format='%Y/%m/%d %H:%M:%S').date()
        elif text_split[3] == 'start':
            instruction_dict['Start'] = grouped_df[grouped_df['Well Name'] == well][grouped_df['Phase Code'] == int(text_split[5])]['Projection Start Time'].iloc[0].date()
        elif text_split[3] == 'end':
            instruction_dict['Start'] = grouped_df[grouped_df['Well Name'] == well][grouped_df['Phase Code'] == int(text_split[5])]['Projection End Time'].iloc[0].date()
        to_index = text_split.index('to')
        if re.findall(r'\b\d{4}/\d{2}/\d{2}\b', text_split[to_index+1]):
            instruction_dict['End'] = pd.to_datetime(text_split[to_index+1], format='%Y/%m/%d %H:%M:%S').date()
        elif text_split[to_index+1] == 'start':
            instruction_dict['End'] = grouped_df[grouped_df['Well Name'] == well][grouped_df['Phase Code'] == int(text_split[to_index+3])]['Projection Start Time'].iloc[0].date()
        elif text_split[to_index+1] == 'end':
            instruction_dict['End'] = grouped_df[grouped_df['Well Name'] == well][grouped_df['Phase Code'] == int(text_split[to_index+3])]['Projection End Time'].iloc[0].date()
        instruction_dict['Dict'] = None
    elif text_split[2] == 'for':
        instruction_dict['Start'] = None
        instruction_dict['End'] = None
        instruction_dict['Dict'] = eval(text[text.find('{'):text.find('}')+1])
    else:  # No recurrence.
        if re.findall(r'\b\d{4}/\d{2}/\d{2}\b', text_split[3]):
            instruction_dict['Start'] = pd.to_datetime(text_split[3], format='%Y/%m/%d %H:%M:%S').date()
        elif text_split[3] == 'start':
            instruction_dict['Start'] = grouped_df[grouped_df['Well Name'] == well][grouped_df['Phase Code'] == int(text_split[5])]['Projection Start Time'].iloc[0].date()
        elif text_split[3] == 'end':
            instruction_dict['Start'] = grouped_df[grouped_df['Well Name'] == well][grouped_df['Phase Code'] == int(text_split[5])]['Projection End Time'].iloc[0].date()
        instruction_dict['End'] = None
        instruction_dict['Dict'] = None
    if text_split[-1] == 'occurrences':
        instruction_dict['Occurrence'] = text_split[-2]
    else:
        instruction_dict['Occurrence'] = 99999
    return instruction_dict


# Generate DCCS with placeholder columns.
df_DCCS = df_OCS.copy(deep=True)
df_DCCS['Total Cost (USD)'] = None
df_DCCS['Total Units'] = None
for date in well_date_range:
    df_DCCS[date.date()] = None
