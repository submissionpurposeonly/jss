import pandas as pd
import json
import time
import os
import sys

# --- Configuration ---
# IMPORTANT: Set your OpenAI API Key here.
# It's recommended to use an environment variable for security.
# For example: API_KEY = os.environ.get("OPENAI_API_KEY")
API_KEY = ""  # @param {type:"string"}

# You can choose which GPT model to use. gpt-4o is recommended for quality.
# gpt-3.5-turbo is faster and cheaper but might be less accurate.
MODEL_NAME = "gpt-4o" 

INPUT_FILE = 'Exclusion2_1588.xlsx'
OUTPUT_FILE = 'Exclusion_Screening_Results_OpenAI.xlsx'
TITLE_COLUMN = 'title'
ABSTRACT_COLUMN = 'abstract'

# --- OpenAI API Setup ---
# Make sure to install the OpenAI library: pip install openai
try:
    from openai import OpenAI
    client = OpenAI(api_key=API_KEY)
except ImportError:
    print("OpenAI Python library not found. Please install it using: pip install openai")
    sys.exit(1)
except Exception as e:
    print(f"Error configuring OpenAI API: {e}")
    sys.exit(1)


def get_screening_prompt(title, abstract):
    """
    Creates a detailed, structured prompt for the LLM to analyze a paper.
    """
    # This prompt is identical to the Gemini version as its logic is model-agnostic.
    prompt_content = f"""
    You are a meticulous senior researcher conducting a Systematic Literature Review (SLR) in Software Engineering. Your task is to analyze a research paper based on its title and abstract and decide if it should be excluded according to two specific criteria.

    **Research Paper Details:**
    - **Title:** "{title}"
    - **Abstract:** "{abstract}"

    **Exclusion Criteria:**
    1.  **EC7:** The article mentions the use of FM-based agents without describing the employed techniques or architecture. (Does the paper seem to focus on *how* the agent works internally, or does it just mention it as a tool?)
    2.  **EC8:** The study's primary contribution is the empirical evaluation of an agent’s performance on an SE task, rather than a novel contribution to the agent’s architectural design, patterns, or principles. (Is the main point "we created a new agent/method" or "we tested an existing agent/method"?)

    **Your Task:**
    Analyze the paper and provide a structured JSON output. For each criterion, you must provide:
    1.  A detailed "comment" explaining your reasoning, as if you were leaving a note for a colleague.
    2.  A final "decision" which must be either "Include" (the paper does NOT meet this exclusion criterion) or "Exclude" (the paper DOES meet this exclusion criterion).

    **JSON Output Format (MUST follow this structure exactly):**
    {{
      "EC7_Comment": "Your detailed analysis for EC7 here. Explain why you think the paper does or does not describe techniques/architecture based on the abstract.",
      "EC7_Decision": "Include or Exclude",
      "EC8_Comment": "Your detailed analysis for EC8 here. Explain if the primary contribution seems to be novel design or just performance evaluation.",
      "EC8_Decision": "Include or Exclude"
    }}
    """
    return prompt_content

def analyze_paper_with_openai(title, abstract):
    """
    Calls the OpenAI API to analyze a single paper and returns the structured JSON response.
    Includes retry logic for API calls.
    """
    if not title or not isinstance(title, str) or not abstract or not isinstance(abstract, str):
        print("Skipping row due to missing or invalid title/abstract.")
        return {
            "EC7_Comment": "Skipped: Missing title or abstract.", "EC7_Decision": "Error",
            "EC8_Comment": "Skipped: Missing title or abstract.", "EC8_Decision": "Error"
        }

    prompt = get_screening_prompt(title, abstract)
    retries = 3
    delay = 5  # seconds

    for attempt in range(retries):
        try:
            # For GPT-4 and newer, you must enable JSON mode for reliable JSON output
            response = client.chat.completions.create(
                model=MODEL_NAME,
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": "You are a helpful research assistant that always responds in JSON format."},
                    {"role": "user", "content": prompt}
                ]
            )
            # The response content is a JSON string
            return json.loads(response.choices[0].message.content)
        except Exception as e:
            print(f"API Error: {e}. Retrying in {delay} seconds... (Attempt {attempt + 1}/{retries})")
            time.sleep(delay)
    
    # If all retries fail
    print("API call failed after multiple retries.")
    return {
        "EC7_Comment": "API call failed after multiple retries.", "EC7_Decision": "Error",
        "EC8_Comment": "API call failed after multiple retries.", "EC8_Decision": "Error"
    }


def main():
    """
    Main function to read the Excel, process each row, and save the results.
    Includes logic to resume from a checkpoint.
    """
    try:
        if os.path.exists(OUTPUT_FILE):
            print(f"--- Resuming from previously saved file: {OUTPUT_FILE} ---")
            df = pd.read_excel(OUTPUT_FILE)
        else:
            print(f"--- Starting a new screening process ---")
            df = pd.read_excel(INPUT_FILE)
            df['EC7_Comment'] = ''
            df['EC7_Decision'] = ''
            df['EC8_Comment'] = ''
            df['EC8_Decision'] = ''
            df['Overall_Decision'] = ''
            
    except FileNotFoundError:
        print(f"Error: Input file '{INPUT_FILE}' not found.")
        return

    start_index = df[df['Overall_Decision'].isin(['', pd.NA, None])].index.min()
    
    if pd.isna(start_index):
        print("All articles have already been processed. Nothing to do.")
    else:
        print(f"Starting screening process for {len(df)} articles, resuming from article {start_index + 1}...")
        for index, row in df.iloc[start_index:].iterrows():
            print(f"Processing article {index + 1}/{len(df)}: {row[TITLE_COLUMN][:70]}...")
            
            title = row[TITLE_COLUMN]
            abstract = row[ABSTRACT_COLUMN]
            
            analysis_result = analyze_paper_with_openai(title, abstract)
            
            df.at[index, 'EC7_Comment'] = analysis_result.get('EC7_Comment', 'Error parsing response.')
            df.at[index, 'EC7_Decision'] = analysis_result.get('EC7_Decision', 'Error')
            df.at[index, 'EC8_Comment'] = analysis_result.get('EC8_Comment', 'Error parsing response.')
            df.at[index, 'EC8_Decision'] = analysis_result.get('EC8_Decision', 'Error')

            ec7_decision = df.at[index, 'EC7_Decision']
            ec8_decision = df.at[index, 'EC8_Decision']

            if ec7_decision == 'Exclude' or ec8_decision == 'Exclude':
                df.at[index, 'Overall_Decision'] = 'Exclude'
            elif ec7_decision == 'Include' and ec8_decision == 'Include':
                df.at[index, 'Overall_Decision'] = 'Include'
            else:
                df.at[index, 'Overall_Decision'] = 'Review Manually'

            if (index + 1) % 100 == 0:
                df.to_excel(OUTPUT_FILE, index=False)
                print(f"--- Progress saved at article {index + 1} ---")

    df.to_excel(OUTPUT_FILE, index=False)
    
    decision_counts = df['Overall_Decision'].value_counts()
    print("\n--- Screening Complete ---")
    print("Summary of decisions:")
    print(decision_counts)
    
    if 'Include' in decision_counts:
        included_count = decision_counts['Include']
        print(f"\nYou have approximately {included_count} articles to include in the next phase.")
    
    print(f"\nResults have been saved to '{OUTPUT_FILE}'.")


if __name__ == "__main__":
    if API_KEY == "YOUR_OPENAI_API_KEY" or not API_KEY:
        print("ERROR: Please replace 'YOUR_OPENAI_API_KEY' with your actual OpenAI API key in the script.")
    else:
        main()
