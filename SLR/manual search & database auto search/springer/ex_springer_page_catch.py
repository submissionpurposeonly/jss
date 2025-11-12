import pandas as pd
import requests
from bs4 import BeautifulSoup
import time

# 读取 Excel 文件
df = pd.read_excel("springerlink-merged_results.xlsx")

# 添加空列：page_count 和 status
if "page_count" not in df.columns:
    df["page_count"] = None
if "page_fetch_status" not in df.columns:
    df["page_fetch_status"] = None

# 抓取页码的函数
def fetch_springer_pages(doi_url):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        resp = requests.get(doi_url, headers=headers, timeout=10)
        if resp.status_code != 200:
            return None, f"HTTP {resp.status_code}"

        soup = BeautifulSoup(resp.text, "html.parser")

        # 从 meta 标签中提取页码
        start_meta = soup.find("meta", {"name": "citation_firstpage"})
        end_meta = soup.find("meta", {"name": "citation_lastpage"})

        start = int(start_meta["content"]) if start_meta else None
        end = int(end_meta["content"]) if end_meta else None

        if start and end and end >= start:
            return end - start + 1, "success"
        elif start:  # 有起始页但无结束页
            return 1, "single page"
        else:
            return None, "no page info"
    except Exception as e:
        return None, f"error: {str(e)}"

# 逐条抓取
for idx, row in df.iterrows():
    doi = row["Item DOI"]
    if pd.isna(doi) or not isinstance(doi, str):
        df.at[idx, "page_fetch_status"] = "no doi"
        continue

    doi_url = doi if doi.startswith("http") else f"https://doi.org/{doi}"
    print(f"Fetching for: {doi_url}")

    count, status = fetch_springer_pages(doi_url)
    df.at[idx, "page_count"] = count
    df.at[idx, "page_fetch_status"] = status
    print(f"→ page count: {count}, status: {status}")
    
    time.sleep(1)  # 限速，避免被封

# 保存结果
df.to_excel("springerlink_with_page_count.xlsx", index=False)
print("✅ 抓取完成，结果已保存为 springerlink_with_page_count.xlsx")