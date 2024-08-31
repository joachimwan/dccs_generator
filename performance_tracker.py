# Methods to write Performance Tracker.

# Proposed workflow:
# - xxx

from DCCS import *

# Export Performance Tracker to EXCEL.
with pd.ExcelWriter(DCCS_filename, mode='a', engine="openpyxl") as writer:
    grouped_df.to_excel(writer, sheet_name='Performance Tracker', index=False)
