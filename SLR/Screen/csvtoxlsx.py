import pandas as pd
import os

# 修改为你的 CSV 文件路径
csv_file = "merged_literature_data.csv"  # 例如 "merged_literature_data.csv"
xlsx_file = os.path.splitext(csv_file)[0] + ".xlsx"

# 读取 CSV 并写入为 Excel
df = pd.read_csv(csv_file, encoding='utf-8-sig')  # 如果乱码，尝试 encoding='gbk' 或 'utf-8'
df.to_excel(xlsx_file, index=False)

print(f"✔️ 成功将 {csv_file} 转换为 {xlsx_file}")