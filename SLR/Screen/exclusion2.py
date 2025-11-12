# -*- coding: utf-8 -*-
import pandas as pd
from thefuzz import fuzz
import re

# --- Configuration Parameters ---
# Input Excel file name
INPUT_FILE = 'Exclusion345.xlsx'
# Output Excel file name for the deduplicated results
OUTPUT_FILE = 'Exclusion345_deduplicated.xlsx'
# The name of the column containing the titles (Please modify this to match your file)
TITLE_COLUMN = 'title'
# Similarity threshold (0-100). Titles with a similarity score above this value will be considered duplicates.
# 95 is a relatively strict and safe value, which you can adjust as needed.
SIMILARITY_THRESHOLD = 95

def normalize_text(text):
    """
    Function to normalize text for comparison:
    1. Convert to lowercase
    2. Remove all non-alphanumeric characters
    3. Remove extra whitespace
    """
    if not isinstance(text, str):
        return ""
    text = text.lower()
    text = re.sub(r'[^a-z0-9\s]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def deduplicate_titles(df, title_col):
    """
    Deduplicates titles in a DataFrame based on exact and fuzzy matching.
    """
    print("Starting the deduplication process...")
    initial_count = len(df)
    print(f"Original number of articles: {initial_count}")

    # --- Step 1: Exact deduplication based on normalized titles ---
    # Create a new column for the normalized title to use for comparison
    df['normalized_title'] = df[title_col].apply(normalize_text)
    
    # Drop rows where the normalized title is an exact duplicate, keeping the first occurrence
    df_exact_dedup = df.drop_duplicates(subset=['normalized_title'], keep='first')
    
    exact_dedup_count = len(df_exact_dedup)
    print(f"Remaining after exact deduplication: {exact_dedup_count} (Removed {initial_count - exact_dedup_count} exact duplicates)")

    # Reset the index to facilitate operations based on index
    df_processed = df_exact_dedup.reset_index(drop=True)

    # --- Step 2: Fuzzy matching for near-duplicate titles ---
    # Use a set to store the indices of rows to be dropped
    indices_to_drop = set()
    
    titles = df_processed['normalized_title'].tolist()
    num_titles = len(titles)

    print("\nStarting fuzzy matching deduplication (this may take a few minutes, please be patient)...")
    
    # This is an O(n^2) loop, which is feasible for a few thousand records
    for i in range(num_titles):
        # If the current row is already marked for deletion, skip to the next one
        if i in indices_to_drop:
            continue
            
        for j in range(i + 1, num_titles):
            if j in indices_to_drop:
                continue
            
            # Use token_sort_ratio to ignore word order, making the comparison more robust
            similarity_score = fuzz.token_sort_ratio(titles[i], titles[j])
            
            if similarity_score >= SIMILARITY_THRESHOLD:
                # If the similarity is high enough, mark the second article for deletion
                indices_to_drop.add(j)

    # Drop the marked rows from the DataFrame
    df_fuzzy_dedup = df_processed.drop(index=list(indices_to_drop))
    
    final_count = len(df_fuzzy_dedup)
    print(f"Remaining after fuzzy matching: {final_count} (Removed {len(indices_to_drop)} near duplicates)")

    # Remove the helper column to restore the original structure
    df_final = df_fuzzy_dedup.drop(columns=['normalized_title'])
    
    print(f"\nDeduplication complete! A total of {initial_count - final_count} articles were removed.")
    
    return df_final


# --- Main Program ---
if __name__ == "__main__":
    try:
        # Read the Excel file
        # The openpyxl engine needs to be installed: pip install openpyxl
        print(f"Reading file: {INPUT_FILE}")
        df = pd.read_excel(INPUT_FILE)
        
        # Execute the deduplication function
        deduplicated_df = deduplicate_titles(df, TITLE_COLUMN)
        
        # Save the results to a new Excel file
        print(f"Saving results to: {OUTPUT_FILE}")
        deduplicated_df.to_excel(OUTPUT_FILE, index=False)
        
        print("\nScript executed successfully!")

    except FileNotFoundError:
        print(f"Error: The file '{INPUT_FILE}' was not found. Please ensure the filename is correct and the script is in the same directory as the file.")
    except KeyError:
        print(f"Error: The column '{TITLE_COLUMN}' was not found in the file. Please check your Excel file and update the TITLE_COLUMN variable in the script.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

