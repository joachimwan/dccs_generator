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
# - USD/MYR conversion rate

# Association headers:
# - OCS number
# - PO number
# - WBS number

import re
from openpyxl.utils.cell import get_column_letter
from lookahead import *
from OCS import *

# Proposed workflow:
# - Identify latest DCCS.
# - Use try-except to verify lookahead validity and raise errors.
# - Parse information from charging mechanisms.
# - Auto charging using parsed information.
# TODO: - Create UID for each line item.
# TODO: - Handle manual inputs before Today.
# TODO: - Handle consolidation especially different well from Today.

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
    instruction_dict['Number'] = float(text_split[0])
    instruction_dict['Recurrence'] = text_split[2]
    if text_split[2] == 'from':
        if re.findall(r'\b\d{4}/\d{2}/\d{2}\b', text_split[3]):
            instruction_dict['Start'] = pd.to_datetime(text_split[3], format='%Y/%m/%d %H:%M:%S').date()
        elif text_split[3] == 'start':  # Start of phase, to be converted to a date, only to be checked against WBS.
            instruction_dict['Start'] = grouped_df[grouped_df['Well Name'] == well][grouped_df['Phase Code'] == int(text_split[5])]['Projection Start Time'].iloc[0].date()
        elif text_split[3] == 'end':  # End of phase, to be converted to a date, only to be checked against WBS.
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
        instruction_dict['Occurrence'] = float(text_split[-2])
    else:
        instruction_dict['Occurrence'] = 99999
    return instruction_dict


# Auto charging using parsed information.
def auto_charge(row):
    d = create_instruction_dict(row['Charging Mechanism'], row['Well Name'])
    grouped_df_by_WBS = grouped_df[grouped_df['Primary WBS'] == row['WBS Number']].groupby(['Primary WBS']).sum().reset_index()
    try:
        if d['Recurrence'] == 'from':
            # Filter by WBS, grouped by WBS, get value from date columns, multiply Number.
            for date in pd.date_range(start=d['Start'], end=d['End']):
                if date in well_date_range:
                    row[date.date()] = min(d['Number'] * grouped_df_by_WBS.iloc[0][date.date()], d['Occurrence'])
                    d['Occurrence'] -= min(d['Number'] * grouped_df_by_WBS.iloc[0][date.date()], d['Occurrence'])
                    if d['Occurrence'] == 0:
                        break
        elif d['Recurrence'] == 'for':
            # Filter by Well Name and Phase Code in Dict, grouped by WBS, get value from date columns, multiply Number.
            filtered_df = pd.DataFrame()
            for well, phase_list in d['Dict'].items():
                df_temp = grouped_df[(grouped_df['Well Name'] == well) & (grouped_df['Phase Code'].isin(phase_list))].groupby(['Well Name', 'Primary WBS']).sum().reset_index()
                filtered_df = pd.concat([filtered_df, df_temp], ignore_index=True)
            for date in well_date_range:
                row[date.date()] = min(d['Number'] * filtered_df.sum().to_frame().T.iloc[0][date.date()], d['Occurrence'])
                d['Occurrence'] -= min(d['Number'] * filtered_df.sum().to_frame().T.iloc[0][date.date()], d['Occurrence'])
                if d['Occurrence'] == 0:
                    break
        else:  # Lump sum.
            # If WBS Number exists on the Start date, charge Number.
            if grouped_df_by_WBS.iloc[0][d['Start']]:
                row[d['Start']] = d['Number']
                print(f"Warning: Lump sum charge applied on row {row.name}: {row['Well Name']} for {d['Number']} units!")
    except Exception as e:
        print(f"Charging error on row {row.name}: {row['Charging Mechanism']}! Error: {e}")
    return row


# Generate DCCS with placeholder columns.
df_DCCS = df_OCS.copy(deep=True)
df_DCCS['Daily Estimate (USD)'] = None
df_DCCS['Total Cost (USD)'] = None
df_DCCS['Total Units'] = None
df_DCCS = df_DCCS[['File Name', 'OCS Number', 'Well Name', 'WBS Number', 'AFE Number', 'Cost Group', 'Description',
                   'Daily Estimate (USD)', 'SAP Unit Price', 'Currency', 'Unit of Measure', 'Charging Mechanism',
                   'Total Cost (USD)', 'Total Units']]
for date in well_date_range:
    df_DCCS[date.date()] = None

# Charge DCCS as per charging mechanisms.
df_DCCS = df_DCCS.apply(lambda row: auto_charge(row), axis=1).replace({pd.np.nan: None, 0: None})

# Create EXCEL formulas.
sum_col_index = df_DCCS.columns.get_loc('Total Units')+1
price_col_index = df_DCCS.columns.get_loc('SAP Unit Price')+1
well_col_index = df_DCCS.columns.get_loc('Well Name')+1
start_row = 10
df_DCCS['Total Units'] = df_DCCS.apply(lambda row: f'=SUM({get_column_letter(sum_col_index+1)}{row.name + start_row + 2}:{get_column_letter(len(df_DCCS.columns))}{row.name + start_row + 2})', axis=1)
df_DCCS['Total Cost (USD)'] = df_DCCS.apply(lambda row: '={col1}{row1}*{col2}{row1}/IF({col3}{row1}="USD",1,$C$8)'.format(col1=get_column_letter(price_col_index), row1=row.name + start_row + 2, col2=get_column_letter(sum_col_index), col3=get_column_letter(price_col_index + 1)), axis=1)
df_DCCS['Daily Estimate (USD)'] = df_DCCS.apply(lambda row: '=({col1}{row1}=$C$6)*HLOOKUP($C$5,${col2}${row2}:${col3}{row1},ROW({col1}{row1})-{startrow},FALSE)*{col4}{row1}/IF({col5}{row1}="USD",1,$C$8)'.format(col1=get_column_letter(well_col_index), row1=row.name + start_row + 2, col2=get_column_letter(sum_col_index + 1), row2=start_row + 1, col3=get_column_letter(len(df_DCCS.columns)), col4=get_column_letter(price_col_index), col5=get_column_letter(price_col_index + 1), startrow=start_row), axis=1)

# Export DCCS to EXCEL and write to cells.
DCCS_filename = BASE_DIR.joinpath('DCCS', '{}_NTP_DCCS.xlsx'.format(TODAY.date().strftime("%Y%m%d")))
df_DCCS.to_excel(DCCS_filename, sheet_name='DCCS', index=False, startrow=start_row, freeze_panes=(start_row + 1, sum_col_index))
wb = openpyxl.load_workbook(DCCS_filename)
ws = wb['DCCS']
ws["B5"] = "Date"
ws["B6"] = "Well"
ws["B8"] = "USDMYR"
ws["B9"] = "Daily"
ws["C5"] = TODAY.date()
try:
    ws["C6"] = grouped_df[grouped_df[TODAY.date()-pd.Timedelta(days=1)]!=0]['Well Name'].unique()[0]
except Exception as e:
    print("Error:", e)
    ws["C6"] = grouped_df['Well Name'].unique()[0]
ws["C8"] = 4.5
daily_col_index = df_DCCS.columns.get_loc('Daily Estimate (USD)')+1
ws["C9"] = "=SUM({col1}{row1}:{col1}{row2})".format(col1=get_column_letter(daily_col_index), row1=start_row + 2, row2=start_row + 1 + len(df_DCCS.index))
wb.save(DCCS_filename)
wb.close()
