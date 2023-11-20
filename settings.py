# Settings for directory paths for data sources.

from pathlib import Path

BASE_DIR = Path.cwd()

LATEST_LOOKAHEAD_DIR = BASE_DIR.joinpath('lookahead', '230101_NTP_Lookahead')

OCS_DIR = BASE_DIR.joinpath('OCS')

LATEST_DCCS_DIR = BASE_DIR.joinpath('DCCS', '230101_NTP_DCCS')
