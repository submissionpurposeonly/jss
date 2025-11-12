#
# A simple Python script to merge multiple CSV files from a directory 
# and save the result as a single Excel (.xlsx) file.
# It assumes all CSV files have the same header row.
#

import os
import pandas as pd
import glob

# --- Configuration ---
# The path to the directory containing your CSV files.
# Use '.' if the script is in the same folder as the CSVs.
csv_directory = '.' 

# The name of the final merged Excel file you want to create.
output_file = 'ieee_merged_results.xlsx'
# -------------------

# Find all CSV files in the specified directory
try:
    all_files = glob.glob(os.path.join(csv_directory, "*.csv"))
    
    if not all_files:
        print(f"No CSV files found in the directory: {os.path.abspath(csv_directory)}")
        print("Please make sure your CSV files and the script are in the correct folder.")
    else:
        print(f"Found {len(all_files)} CSV files to merge.")
        
        # Create a list to hold the dataframes
        df_list = []
        for f in all_files:
            df = pd.read_csv(f)
            df_list.append(df)
            print(f"  - Reading {f}...")

        # Concatenate all dataframes in the list into a single dataframe
        print("\nMerging files...")
        merged_df = pd.concat(df_list, ignore_index=True)

        # Save the merged dataframe to a new Excel file
        # The engine 'openpyxl' is used for .xlsx files.
        # index=False prevents pandas from writing row indices to the file.
        print(f"Saving to Excel file: {output_file}...")
        merged_df.to_excel(output_file, index=False)

        print(f"\nMerge complete!")
        print(f"All data has been saved to: {output_file}")
        print(f"Total rows in merged file: {len(merged_df)}")

except ImportError:
    print("Error: The 'pandas' or 'openpyxl' library is not installed.")
    print("Please install them by running:")
    print("pip install pandas openpyxl")
except Exception as e:
    print(f"An error occurred: {e}")