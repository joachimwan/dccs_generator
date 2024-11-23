# Methods to write Performance Tracker.

# Proposed workflow:
# - Bridge lookahead's day fraction by phase to DCCS via WBS.
# - Create adjusted day fraction by phase i) by WBS and ii) by Well.
# - Use day fraction by WBS if valid, else use day fraction by Well (e.g. Material, or DRO/COM cross-charging).

import openpyxl
from openpyxl.utils.cell import get_column_letter
import numpy as np
from datetime import datetime
from settings import *


def read_DCCS(excel_file_path=LATEST_DCCS_DIR):
    try:
        sheet_name = 'DCCS'
        wb = openpyxl.load_workbook(filename=excel_file_path, data_only=True)
        sheet = wb[sheet_name]
        sheet.delete_rows(1, 10)
        rows_list = [row for row in sheet.iter_rows(values_only=True)]
        df = pd.DataFrame(data=rows_list[1:], index=None, columns=rows_list[0])
        df.columns = [col.date() if isinstance(col, (datetime, pd.Timestamp)) else col for col in df.columns]
        wb.close()
        return df
    except Exception as e:
        print(f"Error:", e)


def read_day_fraction(excel_file_path=TODAY_DCCS_DIR):
    try:
        sheet_name = 'Day Fraction by Phase'
        wb = openpyxl.load_workbook(filename=excel_file_path, data_only=True)
        sheet = wb[sheet_name]
        rows_list = [row for row in sheet.iter_rows(values_only=True)]
        df = pd.DataFrame(data=rows_list[1:], index=None, columns=rows_list[0])
        df.columns = [col.date() if isinstance(col, (datetime, pd.Timestamp)) else col for col in df.columns]
        wb.close()
        return df
    except Exception as e:
        print(f"Error:", e)


# Load TODAY's DCCS.
df_DCCS = read_DCCS(TODAY_DCCS_DIR)
# Remove all empty date columns.
df_DCCS = df_DCCS.loc[:, ~((df_DCCS == 0) | (df_DCCS.isna()) | (df_DCCS == '')).all()]
# Unpivot date columns.
df_DCCS_melted = pd.melt(df_DCCS, id_vars=df_DCCS.columns[:df_DCCS.columns.get_loc('Total Units') + 1],
                         var_name='Date', value_name='Quantity')
# Remove unnecessary rows and columns.
df_DCCS_melted = df_DCCS_melted.replace({pd.np.nan: 0})
df_DCCS_melted = df_DCCS_melted[df_DCCS_melted['Quantity'] != 0]
df_DCCS_melted.drop(columns=['Daily Estimate (USD)', 'Charging Mechanism', 'Total Cost (USD)', 'Total Units'],
                    inplace=True)
df_DCCS_melted.reset_index(inplace=True, drop=True)
# Generate daily line cost.
df_DCCS_melted['Daily Line Cost (USD)'] = df_DCCS_melted.apply(
    lambda row: row['Quantity'] * row['SAP Unit Price'] / (1 if row['Currency'] == "USD" else USDMYR), axis=1)

# Load TODAY's day fraction generated from lookahead.
df_day_fraction = read_day_fraction()
# Unpivot date columns.
df_day_fraction_melted = pd.melt(df_day_fraction,
                                 id_vars=df_day_fraction.columns[:df_day_fraction.columns.get_loc('Planned Depth') + 1],
                                 var_name='Date', value_name='Day Fraction')
# Remove unnecessary rows and columns.
df_day_fraction_melted = df_day_fraction_melted[df_day_fraction_melted['Day Fraction'] != 0]
df_day_fraction_melted.reset_index(inplace=True, drop=True)
# Generate day fraction by WBS and by Well.
df_day_fraction_melted['Day Fraction by WBS'] = df_day_fraction_melted['Day Fraction'] / df_day_fraction_melted.groupby(['Primary WBS', 'Date'])['Day Fraction'].transform('sum')
df_day_fraction_melted['Day Fraction by Well'] = df_day_fraction_melted['Day Fraction'] / df_day_fraction_melted.groupby(['Well Name', 'Date'])['Day Fraction'].transform('sum')

# Split DCCS into non-material and material.
df_material = df_DCCS[df_DCCS.apply(lambda row: MATERIAL_WBS.get(row['Well Name']) == row['WBS Number'], axis=1)]
df_non_material = df_DCCS[~df_DCCS.apply(lambda row: MATERIAL_WBS.get(row['Well Name']) == row['WBS Number'], axis=1)]
# Generate Phase Code per Well for material rows.
df_material_expanded = pd.merge(df_material, df_AFE_WBS[['Well Name', 'Primary WBS', 'Phase Code', 'Phase']],
                                on=['Well Name'], how='inner')
# Generate Phase Code per WBS Number for non-material rows.
df_non_material_expanded = pd.merge(df_non_material, df_AFE_WBS[['Well Name', 'Primary WBS', 'Phase Code', 'Phase']],
                                    left_on=['Well Name', 'WBS Number'], right_on=['Well Name', 'Primary WBS'],
                                    how='left')
# Concatenate the results.
df_DCCS_expanded = pd.concat([df_non_material_expanded, df_material_expanded], ignore_index=True)
df_DCCS_expanded = df_DCCS_expanded[['File Name', 'OCS Number', 'Well Name', 'WBS Number', 'Cost Group', 'Item Number',
                                     'Description', 'SAP Unit Price', 'Currency', 'Unit of Measure', 'Phase Code',
                                     'Phase']]

# Generate day fraction per Phase Code for each date with non-zero day fraction.
df_DCCS_expanded = pd.merge(df_DCCS_expanded,
                            df_day_fraction_melted[['Well Name', 'Phase Code', 'Primary WBS', 'Date', 'Day Fraction',
                                                    'Day Fraction by WBS', 'Day Fraction by Well']],
                            left_on=['Well Name', 'Phase Code'],
                            right_on=['Well Name', 'Phase Code'], how='left')
# Generate daily line cost for each date.
df_DCCS_expanded = pd.merge(df_DCCS_expanded,
                            df_DCCS_melted[['OCS Number', 'Well Name', 'WBS Number', 'Item Number', 'Description',
                                            'Date', 'Quantity', 'Daily Line Cost (USD)']],
                            on=['OCS Number', 'Well Name', 'WBS Number', 'Item Number', 'Description', 'Date'],
                            how='left')
# Remove unnecessary rows.
df_DCCS_expanded = df_DCCS_expanded.replace({pd.np.nan: 0})
df_DCCS_expanded = df_DCCS_expanded[df_DCCS_expanded["Quantity"] != 0]
df_DCCS_expanded.reset_index(inplace=True, drop=True)
# Generate line cost for each DCCS row.
df_DCCS_expanded['Line Cost (USD)'] = np.where(df_DCCS_expanded['WBS Number'] == df_DCCS_expanded['Primary WBS'],
                                               df_DCCS_expanded['Daily Line Cost (USD)'] * df_DCCS_expanded[
                                                   'Day Fraction by WBS'],
                                               df_DCCS_expanded['Daily Line Cost (USD)'] * df_DCCS_expanded[
                                                   'Day Fraction by Well'])

# Export Performance Tracker to EXCEL.
with pd.ExcelWriter(TODAY_DCCS_DIR, mode='a', engine="openpyxl") as writer:
    df_DCCS_expanded.to_excel(writer, sheet_name='DCCS Expanded', index=False, freeze_panes=(1, 0))

# Open Performance Tracker tab in EXCEL.
wb = openpyxl.load_workbook(TODAY_DCCS_DIR)
ws = wb['DCCS Expanded']

# Configure formatting.
for i, header in enumerate(list(df_DCCS_expanded.columns)):
    if header in ['Description', 'Phase']:
        ws.column_dimensions[get_column_letter(i + 1)].width = 40
    if header == 'Date':
        ws.column_dimensions[get_column_letter(i + 1)].width = 13
ws.auto_filter.ref = f'A1:{get_column_letter(ws.max_column)}1'

# Save workbook.
wb.save(TODAY_DCCS_DIR)
wb.close()
