# Settings for directory paths for data sources.

from pathlib import Path
import pandas as pd

BASE_DIR = Path.cwd()

LATEST_LOOKAHEAD_DIR = BASE_DIR.joinpath('lookahead', '20240521_MLNGx2_NTP_Lookahead.xlsm')

OCS_DIR = BASE_DIR.joinpath('OCS')

LATEST_DCCS_DIR = BASE_DIR.joinpath('DCCS', '240101_NTP_DCCS')

df_AFE_WBS = pd.read_excel('Project_AFE.xlsx', index_col=0).reset_index()

MATERIAL_WBS = {'LATOK-1': 'C.MY.MLX.XG.23.010.1094M'}

# Correct to Malaysia's time zone if required using: + pd.Timedelta(hours=8).
TODAY = pd.to_datetime("today")

# FOREX.
USDMYR = 4.5
