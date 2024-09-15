from DCCS import *

# C:\Users\Joachim.Wan\Desktop\OpsProject\dccs_generator
# Proposed workflow:
# - Adjust settings.py file to select data sources.
# - Read the latest Excel Campaign Lookahead. Extract and verify information.
# - Read all OCS and OCS revisions. Extract and verify information.
# - Read all tariffs. Extract and verify information.
# - Read the latest DCCS. Extract and verify information. Update manual inputs.
# - Generate DCCS rows from validated OCS and tariffs.
# - Generate charging instructions based on charging mechanisms and Lookahead (projected days).
# - Generate charges based on manual inputs.

'''
numpy==1.22.3
openpyxl==3.0.9
pandas==1.4.2
'''

if __name__ == '__main__':
    try:
        pass
    except Exception as e:
        print("Error:", e)
