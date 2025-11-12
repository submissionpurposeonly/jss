import os
import bibtexparser
import pandas as pd
from bibtexparser.bparser import BibTexParser
from bibtexparser.customization import homogenize_latex_encoding

def parse_bib_files_to_excel(directory_path, output_filename="bib_summary.xlsx"):

    all_entries_data = []
    
    # 检查目录是否存在
    if not os.path.isdir(directory_path):
        print(f"error: directory '{directory_path}' not found")
        return

    print(f"start: '{directory_path}'...")

    # 
    for filename in os.listdir(directory_path):
        if filename.lower().endswith(".bib"):
            file_path = os.path.join(directory_path, filename)
            print(f"  -> processing: {filename}")
            
            try:
                with open(file_path, 'r', encoding='utf-8') as bib_file:
                    # 
                    parser = BibTexParser()
                    parser.customization = homogenize_latex_encoding
                    parser.ignore_errors = True
                    parser.common_strings = True
                    
                    bib_database = bibtexparser.load(bib_file, parser=parser)
                    
                    for entry in bib_database.entries:
                        #  'N/A'
                        title = entry.get('title', 'N/A')
                        # 
                        authors = entry.get('author', 'N/A').replace(' and ', ', ')
                        year = entry.get('year', 'N/A')
                        journal = entry.get('journal', entry.get('booktitle', 'N/A')) # 期刊或会议名
                        volume = entry.get('volume', 'N/A')
                        pages = entry.get('pages', 'N/A')
                        abstract = entry.get('abstract', 'N/A')
                        doi = entry.get('doi', 'N/A')
                        entry_type = entry.get('ENTRYTYPE', 'N/A')
                        citation_key = entry.get('ID', 'N/A')

                        all_entries_data.append({
                            'Citation Key': citation_key,
                            'Type': entry_type,
                            'Title': title,
                            'Authors': authors,
                            'Year': year,
                            'Journal/Conference': journal,
                            'Volume': volume,
                            'Pages': pages,
                            'DOI': doi,
                            'Abstract': abstract,
                            'Source File': filename # 记录来源文件
                        })
            except Exception as e:
                print(f"    -> process {filename} error: {e}")

    if not all_entries_data:
        print("no .bib ")
        return

    print(f"\ndone！all {len(all_entries_data)} refernces")
    
    # 创建 DataFrame 并保存到 Excel
    try:
        df = pd.DataFrame(all_entries_data)
        df.to_excel(output_filename, index=False, engine='openpyxl')
        print(f"saved to: {output_filename}")
    except Exception as e:
        print(f"write Excel error: {e}")


if __name__ == '__main__':

    bib_files_directory = "."

    # 
    output_excel_file = "my_literature_summary.xlsx"

    parse_bib_files_to_excel(bib_files_directory, output_excel_file)
