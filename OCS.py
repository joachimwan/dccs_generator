# Methods to read and verify Excel OCS.

# OCS headers:
# - Description
# - SAP Element Number
# - Quantity
# - Unit of Measure
# - Estimated Duration
# - Currency
# - SAP Unit Price
# - Total Price

# Metadata:
# - Contract Title
# - Contract Number
# - SAP OA Number
# - OCS Number
# - Well Name
# - WBS Number

# Charging mechanism:
# - 'XX' unit/day 'daily/weekly/monthly' from 'start date/Phase' to 'end date/Phase' or 'XX occurrences'

import pandas as pd
import openpyxl

# Proposed workflow:
# - For each OCS file in the OCS folder, read each OCS file.
# - Use try-except to verify OCS validity and raise errors.
# - Detect revisions and ensure revision does not impact charged items before Today.
# -

# Proposed verification:
# - All required fields are present, e.g. OCS Number, WBS Number.
# - WBS Number and Well Name (if present) are correct.
