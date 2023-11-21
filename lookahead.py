# Methods to read and verify Excel Lookahead.

# Lookahead headers:
# - Start Time
# - Well Name
# - Phase Code
# - Phase
# - Description
# - AFE Time
# - DSV Time
# - Actual Time

# Metadata:
# - Start Time (the very first start time of the lookahead)

import openpyxl

# Proposed workflow:
# - For each lookahead file in the lookahead folder, read each lookahead file.
# - Identify latest lookahead.
# - Use try-except to verify lookahead validity and raise errors.
# - Identify number of unique wells in the lookahead.
# - For lines after Actual, fill with AFE Time. For lines without AFE Time, fill with DSV Time.
# - Recalculate Start Time after Actual.
# - For each well, generate Performance Tracker.

# Proposed verification:
# - All expected Phase Codes and Phases for each well are present.
