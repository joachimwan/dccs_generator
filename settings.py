# Settings for directory paths for data sources.

from pathlib import Path
import pandas as pd

BASE_DIR = Path.cwd()

LATEST_LOOKAHEAD_DIR = BASE_DIR.joinpath('lookahead', '230101_Test_NTP_Lookahead.xlsm')

OCS_DIR = BASE_DIR.joinpath('OCS')

LATEST_DCCS_DIR = BASE_DIR.joinpath('DCCS', '230101_NTP_DCCS')

df_AFE_WBS = pd.read_excel('ProjectAFE.xlsx', index_col=0).reset_index()

# Corrected to Malaysia's time zone.
TODAY = pd.to_datetime("today") + pd.Timedelta(hours=8)

# FOREX.
USDMYR = 4.5
