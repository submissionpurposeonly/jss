#
# A Python script to merge multiple Excel (.xls) files from a directory
# and save the combined data into a single new Excel (.xlsx) file.
# This script requires the 'pandas', 'openpyxl', and 'xlrd' libraries.
#

import os
import glob
import pandas as pd

# --- Configuration ---
# The path to the directory containing your .xls files.
# Use '.' if the script is in the same folder as the Excel files.
xls_directory = '.' 

# The name of the final merged Excel file.
output_file = 'merged_results.xlsx'
# -------------------

try:
    # Find all files ending with .xls in the specified directory
    # Note: You can change "*.xls" to "*.xlsx" if your files are in the newer format.
    # Or use a more general pattern if you have mixed types.
    all_files = glob.glob(os.path.join(xls_directory, "*.xls"))
    
    if not all_files:
        print(f"No .xls files found in the directory: {os.path.abspath(xls_directory)}")
        print("Please make sure your .xls files and the script are in the correct folder.")
    else:
        print(f"Found {len(all_files)} .xls files to merge.")
        
        # A list to hold all the pandas DataFrames
        df_list = []
        
        for f in all_files:
            # Read each .xls file into a DataFrame.
            # The 'xlrd' engine is needed for older .xls files.
            try:
                df = pd.read_excel(f, engine='xlrd')
                df_list.append(df)
                print(f"  - Reading {f}...")
            except Exception as e:
                print(f"  - Could not read file {f}. Error: {e}")

        if not df_list:
             print("\nNo data was successfully read from the files. Exiting.")
        else:
            # Concatenate all DataFrames in the list into a single DataFrame
            print("\nMerging files...")
            merged_df = pd.concat(df_list, ignore_index=True)

            # Save the merged DataFrame to a new .xlsx file
            # The 'openpyxl' engine is used for writing .xlsx files.
            print(f"Saving to new Excel file: {output_file}...")
            merged_df.to_excel(output_file, index=False)
            
            print(f"\nMerge complete!")
            print(f"All data has been saved to: {output_file}")
            print(f"Total rows in merged file: {len(merged_df)}")

except ImportError:
    print("Error: Required libraries are not installed.")
    print("Please install them by running:")
    print("pip install pandas openpyxl xlrd")
except Exception as e:
    print(f"An unexpected error occurred: {e}")