import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

# 读取Excel文件
df = pd.read_excel("sciencedirect_merged_results.xlsx")

def fetch_page_count_from_doi(doi_url):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        resp = requests.get(doi_url, headers=headers, timeout=10)
        if resp.status_code != 200:
            return None
        soup = BeautifulSoup(resp.text, "html.parser")

        # 尝试从 meta 标签中提取页码
        meta = soup.find("meta", {"name": "citation_firstpage"})
        start = int(meta["content"]) if meta else None

        meta = soup.find("meta", {"name": "citation_lastpage"})
        end = int(meta["content"]) if meta else None

        if start and end:
            return end - start + 1
    except Exception as e:
        return None

# 仅抓取没有 page_count 的行
for idx, row in df[df["page_count"].isna()].iterrows():
    doi_url = row["doi"]
    if pd.isna(doi_url) or not isinstance(doi_url, str) or "doi.org" not in doi_url:
        continue
    print(f"Fetching for: {doi_url}")
    page_count = fetch_page_count_from_doi(doi_url)
    if page_count:
        df.at[idx, "page_count"] = page_count
    time.sleep(1)  # 避免被封锁

# 保存结果
df.to_excel("sciencedirect_with_page_count.xlsx", index=False)