import numpy as np
import pandas as pd
from settings import *

# C:\Users\Joachim.Wan\Desktop\OpsProject\dccs_generator
# Proposed workflow:
# - Adjust settings.py file to select data sources.
# - Read the latest Excel Campaign Lookahead. Extract and verify information.
# - Identify number of unique wells in the lookahead.
# - Read all OCS and OCS revisions. Extract and verify information.
# - Read all tariffs. Extract and verify information.
# - Read the latest DCCS. Extract and verify information. Update manual inputs.
# - Generate DCCS rows from validated OCS and tariffs.
# - Generate charges based on charging mechanisms and Lookahead (actual and projected AFE days).
# - Generate charges based on manual inputs.

if __name__ == '__main__':
    try:
        pass
    except Exception as e:
        print("Error:", e)
