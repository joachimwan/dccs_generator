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

import openpyxl

# Proposed workflow:
# - For each DCCS file in the DCCS folder, read each DCCS file.
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
