## read datatype

import os
import pandas as pd

# Determine the directory of the script
script_dir = os.path.dirname(os.path.abspath(__file__))

input_folder = os.path.join(script_dir, 'input')

# Iterate through all Excel files in the input folder
for excel_file in os.listdir(input_folder):
    if not excel_file.endswith('.xlsx'):
        continue

    # Construct the file path
    file_path = os.path.join(input_folder, excel_file)
    
    # Check if file exists
    if not os.path.isfile(file_path):
        print(f"File not found: {file_path}")
        continue
    
    try:
        # Read the financial data from the first worksheet of the Excel file
        df = pd.read_excel(file_path, sheet_name=0)
        print(df.dtypes)
    except PermissionError as e:
        print(f"Permission error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
