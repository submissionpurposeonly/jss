import pandas as pd
import time
import os
from openai import OpenAI, OpenAIError

# ========== CONFIG ==========
# IMPORTANT: Replace with your OpenAI API key below.
# For security, it's recommended to use environment variables for your API key.
api_key = ""
client = OpenAI(api_key=api_key)

INPUT_FILE = "final_merged_literature_data.xlsx"
OUTPUT_FILE = "phase2_screened_gpt_output.xlsx"
CHECKPOINT_INTERVAL = 100  # Save progress every 100 articles

MODEL = "gpt-3.5-turbo"  # or "gpt-4"
RATE_LIMIT_DELAY = 1  # Delay in seconds between each API call
# =============================

def safe_str(text):
    """Safely converts input to a clean string, handling potential NaN values."""
    return str(text).strip() if pd.notna(text) else ""

def build_prompt(title, abstract, keywords):
    """
    Builds the prompt for the GPT model based on the refined requirements.
    """
    return f"""
You are a research assistant conducting a systematic literature review (SLR).
Your task is to strictly evaluate the following paper based on its Title, Abstract, and Keywords.

Evaluate the paper against these three criteria:

1.  **FM/LLM Agent Focus**: Does the article explicitly claim that its core subject is the use of a Foundation Model (FM) or Large Language Model (LLM) based agent? A brief mention is not enough. The agent must be central to the paper's contribution.
2.  **Software Engineering Context**: Does the study involve a Software Engineering (SE) task or discuss software/system architecture? SE tasks include requirements, design, coding, testing, maintenance, etc.
3.  **Language**: Is the article written in English?

Your evaluation must be strict. If you are uncertain about any criterion based on the provided text, answer 'No'.

**Paper Details:**
- **Title**: {safe_str(title)}
- **Abstract**: {safe_str(abstract)}
- **Keywords**: {safe_str(keywords)}

**Your Response:**
Respond ONLY in the following format, without any explanations or introductory text:
FM/LLM: <Yes/No>
SE: <Yes/No>
English: <Yes/No>
""".strip()

def call_gpt(prompt, retries=3):
    """
    Calls the OpenAI API. Includes a retry mechanism with exponential backoff for robustness.
    """
    for attempt in range(retries):
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "user", "content": prompt}
                ],
                temperature=0  # Set to 0 for more deterministic and reproducible outputs
            )
            return response.choices[0].message.content
        except OpenAIError as e:
            wait_time = 2 ** attempt  # Exponential backoff
            print(f"⚠️ GPT API Error (Attempt {attempt + 1}/{retries}): {e}. Retrying in {wait_time}s...")
            time.sleep(wait_time)
    print("❌ Failed to get a response from GPT after multiple retries.")
    return "Error: API call failed"

def parse_criteria(reply):
    """
    Parses the raw text response from GPT into structured Yes/No values.
    """
    reply_lower = reply.lower() if isinstance(reply, str) else ""
    fm = "Yes" if "fm/llm: yes" in reply_lower else "No"
    se = "Yes" if "se: yes" in reply_lower else "No"
    en = "Yes" if "english: yes" in reply_lower else "No"
    # Determine inclusion based on whether all criteria are met with "Yes"
    include = fm == "Yes" and se == "Yes" and en == "Yes"
    return fm, se, en, include

def main():
    """
    Main function to run the literature screening process.
    """
    if os.path.exists(OUTPUT_FILE):
        print(f"📄 Found existing output file '{OUTPUT_FILE}'. Resuming screening from checkpoint.")
    else:
        print(f"🚀 Starting a new screening task from '{INPUT_FILE}'.")
        df = pd.read_excel(INPUT_FILE)
        # Prepare result columns for a new task
        df["fm_llm"] = ""
        df["se_related"] = ""
        df["english"] = ""
        df["included_by_gpt"] = False
        df["gpt_screening_result"] = ""

    total_articles = len(df)
    for idx in df.index:
        # Check if the current row has already been processed
        if pd.notna(df.at[idx, "gpt_screening_result"]) and df.at[idx, "gpt_screening_result"] != "":
            print(f"⏩ Skipping article {idx + 1}/{total_articles} (already processed).")
            continue

        row = df.iloc[idx]
        prompt = build_prompt(row.get("title"), row.get("abstract"), row.get("keywords"))
        
        print(f"🧠 Screening article {idx + 1}/{total_articles}...")
        result = call_gpt(prompt)
        fm, se, en, include = parse_criteria(result)

        # Update the DataFrame with the new results
        df.at[idx, "fm_llm"] = fm
        df.at[idx, "se_related"] = se
        df.at[idx, "english"] = en
        df.at[idx, "included_by_gpt"] = include
        df.at[idx, "gpt_screening_result"] = result

        # Checkpoint: Save progress at the specified interval
        if (idx + 1) % CHECKPOINT_INTERVAL == 0:
            print(f"💾 Checkpoint reached. Saving progress for the first {idx + 1} articles...")
            df.to_excel(OUTPUT_FILE, index=False)

        time.sleep(RATE_LIMIT_DELAY)

    # Final save to ensure the last batch of data is written to the file
    print("💾 Performing final save of all results...")
    df.to_excel(OUTPUT_FILE, index=False)
    
    included_count = df['included_by_gpt'].sum()
    print(f"✅ Screening complete! Total articles included: {included_count}/{total_articles}")

if __name__ == "__main__":
    main()
