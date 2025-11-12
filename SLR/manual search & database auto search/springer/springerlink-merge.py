#
# A Python script to find and merge all CSV files in a directory
# and save the combined data into a single Excel (.xlsx) file.
# This script requires the 'pandas' and 'openpyxl' libraries.
#

import os
import glob
import pandas as pd

# --- Configuration ---
# The path to the directory containing your CSV files.
# Use '.' if the script is in the same folder as the CSVs.
csv_directory = '.' 

# The name of the final merged Excel file you want to create.
output_file = 'merged_results.xlsx'
# -------------------

try:
    # Find all files ending with .csv in the specified directory
    all_files = glob.glob(os.path.join(csv_directory, "*.csv"))
    
    if not all_files:
        print(f"No CSV files found in the directory: {os.path.abspath(csv_directory)}")
        print("Please make sure your CSV files and the script are in the correct folder.")
    else:
        print(f"Found {len(all_files)} CSV files to merge.")
        
        # Create a list to hold the individual DataFrames
        df_list = []
        
        for f in all_files:
            try:
                # Read each CSV file into a DataFrame
                df = pd.read_csv(f)
                df_list.append(df)
                print(f"  - Reading {f}...")
            except Exception as e:
                print(f"  - Could not read file {f}. Error: {e}")

        if not df_list:
             print("\nNo data was successfully read from the files. Exiting.")
        else:
            # Concatenate all DataFrames in the list into a single DataFrame
            print("\nMerging all CSV files...")
            merged_df = pd.concat(df_list, ignore_index=True)

            # Save the merged DataFrame to a new Excel file
            print(f"Saving to new Excel file: {output_file}...")
            # 'index=False' prevents pandas from writing the DataFrame index as a column
            merged_df.to_excel(output_file, index=False)
            
            print(f"\nMerge complete!")
            print(f"All data has been successfully saved to: {output_file}")
            print(f"Total rows in the merged file: {len(merged_df)}")

except ImportError:
    print("Error: Required libraries 'pandas' or 'openpyxl' are not installed.")
    print("Please install them by running the following command in your terminal:")
    print("pip install pandas openpyxl")
except Exception as e:
    print(f"An unexpected error occurred: {e}")