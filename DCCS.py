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
# - First day of DCCS
# - USD/MYR conversion rate

# Association headers:
# - OCS number
# - PO number
# - WBS number

import openpyxl
