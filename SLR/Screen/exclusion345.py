import pandas as pd
import os
import time
from openai import OpenAI
from tqdm import tqdm

# --- 配置 ---
# ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
# 只需要修改这一行！把你的密钥粘贴到下面的引号里
YOUR_OPENAI_API_KEY = ""
# ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

# ------------------- 下面的代码请不要修改 -------------------

# 检查密钥是否已填写
# 使用你填写的密钥初始化 OpenAI 客户端
try:
    client = OpenAI(api_key=YOUR_OPENAI_API_KEY)
except Exception as e:
    print(f"初始化 OpenAI 客户端时出错: {e}")
    exit()

# --- 函数定义 ---
def classify_with_gpt(prompt, max_retries=3):
    """
    使用 OpenAI GPT 模型进行分类，并包含重试机制。
    """
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful research assistant. Your task is to answer classification questions with only 'Yes' or 'No'."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0,
                max_tokens=5
            )
            text_response = response.choices[0].message.content.strip().capitalize()
            if 'Yes' in text_response:
                return 'Yes'
            elif 'No' in text_response:
                return 'No'
            else:
                return 'Uncertain'
        except Exception as e:
            print(f"API 调用出错: {e}。将在 {5 * (attempt + 1)} 秒后重试...")
            time.sleep(5 * (attempt + 1))
    return "API_Error"

def intelligent_screening(input_filename='phase2_screened_gpt_output.xlsx'):
    """
    使用 GPT API 对 SLR 数据进行智能筛选，并每100条保存一次进度。
    """
    try:
        df = pd.read_excel(input_filename)
        print(f"成功加载 '{input_filename}'。发现 {len(df)} 条记录。")
    except FileNotFoundError:
        print(f"错误: 文件 '{input_filename}' 未找到。")
        return

    # 定义输出文件名
    output_all_filename = 'slr_gpt_results_all.xlsx'

    # 为AI的判断结果创建新的列 (如果它们不存在)
    for col in ['AI_C3_PrimarySource', 'AI_C4_VenueType', 'AI_C5_GreyLiterature']:
        if col not in df.columns:
            df[col] = ''

    for index, row in tqdm(df.iterrows(), total=df.shape[0], desc="正在使用GPT筛选文章"):
        # 如果这一行已经被处理过了 (AI_C3列有值)，就跳过
        if pd.notna(row['AI_C3_PrimarySource']) and row['AI_C3_PrimarySource'] != '':
            continue

        data = {key: str(row.get(key, '')) for key in 
                ['ENTRYTYPE', 'title', 'isbn', 'publisher', 'source', 'booktitle', 'series', 'note', 'url']}

        prompt_c3 = f"""Is the following a primary research source (like a peer-reviewed journal article or conference paper)? Exclude books, theses (PhD/Master's), and editorials.
        - Entry Type: "{data['ENTRYTYPE']}"
        - Title: "{data['title']}"
        - ISBN: "{data['isbn']}"
        - Publisher: "{data['publisher']}"
        Answer with only 'Yes' or 'No'."""
        df.loc[index, 'AI_C3_PrimarySource'] = classify_with_gpt(prompt_c3)

        prompt_c4 = f"""Is the publication venue a main conference or journal? Exclude venues that are clearly a workshop, symposium, doctoral consortium, or companion proceeding.
        - Source/Journal: "{data['source']}"
        - Book Title: "{data['booktitle']}"
        - Series: "{data['series']}"
        Answer with only 'Yes' or 'No'."""
        df.loc[index, 'AI_C4_VenueType'] = classify_with_gpt(prompt_c4)
        
        prompt_c5 = f"""Is this a formal, peer-reviewed publication? Exclude non-refereed grey literature like technical reports or preprints from servers like arXiv.
        - Entry Type: "{data['ENTRYTYPE']}"
        - Note: "{data['note']}"
        - Publisher: "{data['publisher']}"
        Answer with only 'Yes' or 'No'."""
        df.loc[index, 'AI_C5_GreyLiterature'] = classify_with_gpt(prompt_c5)

        # --- 新增的自动保存逻辑 ---
        # 每处理100条记录就保存一次
        if (index + 1) % 100 == 0:
            try:
                df.to_excel(output_all_filename, index=False)
                # 使用 tqdm.write 打印信息，避免弄乱进度条
                tqdm.write(f"--- 进度已保存！已处理 {index + 1}/{len(df)} 篇文章 ---")
            except Exception as e:
                tqdm.write(f"--- 保存进度时出错: {e} ---")

    # --- 最终处理和保存 ---
    all_criteria_passed = (df['AI_C3_PrimarySource'] == 'Yes') & \
                          (df['AI_C4_VenueType'] == 'Yes') & \
                          (df['AI_C5_GreyLiterature'] == 'Yes')
    df['Included_AI_Final'] = pd.Series(all_criteria_passed).map({True: 'Yes', False: 'No'})
    
    # 最后再完整保存一次，确保所有数据都已写入
    df.to_excel(output_all_filename, index=False)
    print(f"\n筛选完成！所有AI辅助判断的结果已保存至 '{output_all_filename}'。")

    df_included = df[all_criteria_passed]
    output_included_filename = 'slr_gpt_results_included.xlsx'
    df_included.to_excel(output_included_filename, index=False)
    
    total_included = len(df_included)
    total_excluded = len(df) - total_included
    print(f"最终统计: {total_included} 篇文章被纳入, {total_excluded} 篇文章被排除。")
    print(f"已筛选出的文章数据保存至 '{output_included_filename}'。")

# --- 运行脚本 ---
if __name__ == '__main__':
    intelligent_screening()
