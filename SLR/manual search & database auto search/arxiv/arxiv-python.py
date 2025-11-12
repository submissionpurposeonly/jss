import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
import time
import re
import sys

def clean_text(text):
    """清理从 XML 中提取的文本，替换多个换行符和空格。"""
    if not text:
        return ""
    text = re.sub(r'\s*\n\s*', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def fetch_papers_for_query(search_query, start_date, end_date):
    """
    为单个查询获取所有符合条件的论文。
    这是一个核心函数，用于被循环调用。
    """
    base_url = 'http://export.arxiv.org/api/query?'
    papers_for_this_query = []
    
    # 1. 获取此子查询的总结果数
    try:
        initial_query_params = {'search_query': search_query, 'start': 0, 'max_results': 1}
        with urllib.request.urlopen(base_url + urllib.parse.urlencode(initial_query_params)) as response:
            initial_xml = response.read().decode('utf-8')
        root = ET.fromstring(initial_xml)
        namespace = {'atom': 'http://www.w3.org/2005/Atom', 'opensearch': 'http://a9.com/-/spec/opensearch/1.1/'}
        total_results = int(root.find('opensearch:totalResults', namespace).text)
        if total_results == 0:
            print(" -> 未找到结果，跳过。")
            return []
        print(f" -> 找到 {total_results} 篇相关论文，开始获取...")
    except Exception:
        print(" -> 获取总数时出错，跳过。")
        return []

    # 2. 分批获取数据
    start = 0
    results_per_page = 200
    while start < total_results:
        query_params = {
            'search_query': search_query,
            'start': start,
            'max_results': results_per_page,
            'sortBy': 'submittedDate',
            'sortOrder': 'descending'
        }
        url = base_url + urllib.parse.urlencode(query_params)
        
        try:
            with urllib.request.urlopen(url) as response:
                xml_data = response.read().decode('utf-8')
        except Exception:
            time.sleep(5) # 如果失败，等待后重试一次
            try:
                 with urllib.request.urlopen(url) as response:
                    xml_data = response.read().decode('utf-8')
            except Exception as e:
                print(f" -> 请求失败: {e}, 跳过此分页。")
                start += results_per_page
                continue

        root = ET.fromstring(xml_data)
        entries = root.findall('atom:entry', namespace)
        if not entries:
            break

        for entry in entries:
            published_str = entry.find('atom:published', namespace).text
            published_date = datetime.strptime(published_str, '%Y-%m-%dT%H:%M:%SZ')

            if start_date <= published_date <= end_date:
                title = clean_text(entry.find('atom:title', namespace).text)
                authors = ', '.join([author.find('atom:name', namespace).text for author in entry.findall('atom:author', namespace)])
                abstract = clean_text(entry.find('atom:summary', namespace).text)
                arxiv_id = entry.find('atom:id', namespace).text.split('/abs/')[-1]
                
                pdf_link_element = entry.find("atom:link[@title='pdf']", namespace)
                pdf_link = pdf_link_element.get('href') if pdf_link_element is not None else "N/A"
                
                primary_category_element = entry.find('atom:primary_category', namespace)
                primary_category = primary_category_element.get('term') if primary_category_element is not None else "N/A"

                papers_for_this_query.append({
                    'Title': title, 'Authors': authors, 'Abstract': abstract,
                    'Published Date': published_date.strftime('%Y-%m-%d'),
                    'arXiv ID': arxiv_id, 'PDF Link': pdf_link, 'Primary Category': primary_category
                })
        
        start += len(entries)
        time.sleep(3) 
        
    return papers_for_this_query


if __name__ == '__main__':
    # --- 定义查询的各个部分 ---
    part1_query_str = (
        '"FM-based agent" OR "Foundation Models" OR "Large Language Models" OR "LLMs" OR "Generative AI" OR "Conversational AI" OR "Transformer Models" OR "Autonomous Agents" OR "Agentic AI" OR "Multi-agent Systems" OR "MAS" OR "LLM-based Agents" OR "Generative Agents" OR "Autonomous Web Agent" OR "AWA"'
    )
    
    part2_terms = [
        "AI Agent Architecture", "Agent-based Systems", "Modular Architectures for AI Agent", "AI Agent Design", 
        "Autonomous Agent System Design", "Adaptive AI Architectures", "Adaptive AI Taxonomy", "Software Architecture", 
        "System Design", "Software Engineering Task", "Code Generation", "Automated Program Repair", "APR", 
        "Bug Fixing", "Fault Localization", "Software Testing", "Code Review", "Requirements Engineering", 
        "Software Maintenance", "Refactoring", "Code Snippet Adaptation", "Code Intent Extraction", 
        "Simulation Testing", "Security", "Prompt Engineering"
    ]

    # --- 设置日期和输出文件名 ---
    start_date_str = '2017-01-01'
    end_date_str = '2025-07-31'
    output_file = 'arXiv_Final_Results_v3.xlsx'

    start_date_obj = datetime.strptime(start_date_str, '%Y-%m-%d')
    end_date_obj = datetime.strptime(end_date_str, '%Y-%m-%d')
    
    master_paper_list = []
    processed_ids = set()

    # --- 循环执行每个子查询 ---
    for i, term in enumerate(part2_terms):
        print(f"\n--- 开始执行子查询 {i+1}/{len(part2_terms)}: (Term: '{term}') ---")
        
        # 采用新的、更细粒度的拆分方式
        abs_query = f'abs:(({part1_query_str}) AND "{term}")'
        final_query = f'(cat:cs.*) AND {abs_query}'
        
        papers_from_query = fetch_papers_for_query(final_query, start_date_obj, end_date_obj)
        
        new_papers_found = 0
        for paper in papers_from_query:
            if paper['arXiv ID'] not in processed_ids:
                master_paper_list.append(paper)
                processed_ids.add(paper['arXiv ID'])
                new_papers_found += 1
        
        print(f" -> 完成。本次查询新增了 {new_papers_found} 篇独一无二的论文。")
        print(f" -> 当前论文总数: {len(master_paper_list)}")

    # --- 所有查询完成，进行最后处理 ---
    print(f"\n\n所有子查询执行完毕！总共收集到 {len(master_paper_list)} 篇独一无二的论文。")
    
    if not master_paper_list:
        print("未找到任何符合条件的论文。")
    else:
        print(f"正在将所有数据写入唯一的 Excel 文件: {output_file}")
        df = pd.DataFrame(master_paper_list)
        # 按日期降序排序
        df_sorted = df.sort_values(by='Published Date', ascending=False)
        df_sorted.to_excel(output_file, index=False, engine='openpyxl')
        print(f"任务成功！最终数据已保存至 {output_file}")
