#
# A Python script to merge multiple RIS files from a directory 
# and save the combined, structured data into a single Excel (.xlsx) file.
# This script requires the 'pandas', 'openpyxl', and 'rispy' libraries.
#

import os
import glob
import pandas as pd
import rispy

# --- Configuration ---
# The path to the directory containing your RIS files.
# Use '.' if the script is in the same folder as the RIS files.
ris_directory = '.' 

# The name of the final merged Excel file.
output_file = 'sciencedirect_merged_results.xlsx'
# -------------------

try:
    # Find all files ending with .ris in the specified directory
    all_files = glob.glob(os.path.join(ris_directory, "*.ris"))
    
    if not all_files:
        print(f"No .ris files found in the directory: {os.path.abspath(ris_directory)}")
        print("Please make sure your .ris files and the script are in the correct folder.")
    else:
        print(f"Found {len(all_files)} RIS files to merge.")
        
        # A list to hold all publication entries from all files
        all_entries = []
        
        for f in all_files:
            try:
                # Open and parse the RIS file using rispy
                # We try utf-8 encoding first, which is common.
                with open(f, 'r', encoding='utf-8') as ris_file:
                    entries = rispy.load(ris_file)
                    all_entries.extend(entries)
                    print(f"  - Reading {f}... Found {len(entries)} entries.")
            except UnicodeDecodeError:
                # If utf-8 fails, try a more lenient encoding like latin-1
                print(f"  - Warning: UTF-8 decoding failed for {f}. Trying with 'latin-1' encoding.")
                with open(f, 'r', encoding='latin-1') as ris_file:
                    entries = rispy.load(ris_file)
                    all_entries.extend(entries)
                    print(f"  - Successfully read {f} with 'latin-1'. Found {len(entries)} entries.")
            except Exception as e:
                print(f"  - Could not process file {f}. Error: {e}")

        if not all_entries:
            print("\nNo entries were successfully read from the files. Exiting.")
        else:
            # Convert the list of dictionaries to a pandas DataFrame
            # The keys from the dictionaries will automatically become column headers
            print("\nCreating DataFrame from all entries...")
            df = pd.DataFrame(all_entries)
            
            # Save the DataFrame to an Excel file
            print(f"Saving to Excel file: {output_file}...")
            df.to_excel(output_file, index=False)
            
            print(f"\nMerge complete!")
            print(f"All data has been saved to: {output_file}")
            print(f"Total records processed: {len(df)}")

except ImportError:
    print("Error: Required libraries are not installed.")
    print("Please install them by running:")
    print("pip install pandas openpyxl rispy")
except Exception as e:
    print(f"An unexpected error occurred: {e}")