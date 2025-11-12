import pandas as pd
import time
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.service import Service

# 初始化浏览器（无头模式）
options = Options()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')

# 正确初始化浏览器的写法
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# 读取 Excel 文件
df = pd.read_excel("sciencedirect_merged_results.xlsx")

# 初始化新列
if "page_count" not in df.columns:
    df["page_count"] = None
if "page_fetch_status" not in df.columns:
    df["page_fetch_status"] = None

# 提取干净的 URL 函数
def extract_clean_url(raw_url):
    if not isinstance(raw_url, str):
        return None
    match = re.search(r'https://www\.sciencedirect\.com[^\']+', raw_url)
    return match.group(0) if match else None

# 抓取页码数
def fetch_page_count(url):
    try:
        driver.get(url)
        time.sleep(1.5)
        soup = BeautifulSoup(driver.page_source, "html.parser")

        # 查找包含页码的标签
        page_span = soup.find("span", string=re.compile(r"Pages", re.IGNORECASE))
        if page_span:
            text = page_span.text.strip()
            match = re.search(r'(\d+)\s*[-–]\s*(\d+)', text)
            if match:
                start = int(match.group(1))
                end = int(match.group(2))
                return end - start + 1, "success"
            else:
                return None, "found span, but no match"
        else:
            return None, "page not found"
    except Exception as e:
        return None, f"error: {str(e)}"

# 主循环
for idx, row in df.iterrows():
    raw_url = row.get("urls", None)
    url = extract_clean_url(raw_url)

    if not url:
        df.at[idx, "page_fetch_status"] = "no valid url"
        continue

    print(f"🔍 [{idx}] Fetching: {url}")
    count, status = fetch_page_count(url)
    df.at[idx, "page_count"] = count
    df.at[idx, "page_fetch_status"] = status
    print(f"→ page count: {count}, status: {status}")
    time.sleep(1.5)

# 保存结果
output_path = "sciencedirect_with_page_count.xlsx"
df.to_excel(output_path, index=False)
print(f"✅ 完成！结果保存为 {output_path}")

# 关闭浏览器
driver.quit()