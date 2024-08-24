# Settings for directory paths for data sources.

from pathlib import Path
import pandas as pd

BASE_DIR = Path.cwd()

LATEST_LOOKAHEAD_DIR = BASE_DIR.joinpath('lookahead', '20240824_MLNGx2_NTP_Lookahead JW.xlsm')

OCS_DIR = BASE_DIR.joinpath('OCS')

LATEST_DCCS_DIR = BASE_DIR.joinpath('DCCS', '20240101_NTP_DCCS.xlsx')

df_AFE_WBS = pd.read_excel('Project_AFE.xlsx', index_col=0).reset_index()

MATERIAL_WBS = {'F27-101': 'C.MY.27X.DD.22.003.1094M',
                'F22-101': 'C.MY.22X.DD.22.003.1094M',
                'SLS-101': 'C.MY.SLX.DD.22.003.1094M'}

# Correct to Malaysia's time zone if required using: + pd.Timedelta(hours=8).
TODAY = pd.to_datetime("today")

# FOREX.
USDMYR = 4.5
