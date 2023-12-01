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

import pandas as pd
import openpyxl
from lookahead import *
from OCS import *

# Proposed workflow:
# - Identify latest DCCS.
# - Use try-except to verify lookahead validity and raise errors.
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


def some_function():
    pass


# Generate DCCS with placeholder columns.
df_DCCS = df_OCS.copy(deep=True)
df_DCCS['Total Cost (USD)'] = None
df_DCCS['Total Units'] = None
for date in pd.date_range(start=lookahead_start_time.date(), end=lookahead_end_time.date()):
    df_DCCS[date.date()] = None
