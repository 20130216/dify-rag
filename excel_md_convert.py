import sys
import os
import shutil
import datetime
import pandas as pd
from dify_rag.extractor.excel_extractor import ExcelExtractor
from dify_rag.extractor.html_extractor import HtmlExtractor

def excel_to_markdown(excel_path, md_path):
    """
    将单个 xlsx/xls 文件转为 markdown 文件，自动清洗 NaN。
    """
    import os
    import pandas as pd
    ext = os.path.splitext(excel_path)[1].lower()
    print("Pandas version:", pd.__version__)
    print("Pandas file:", pd.__file__)
    import xlrd
    print("xlrd version:", xlrd.__version__)
    if ext == ".xls":
        df = pd.read_excel(excel_path, engine="xlrd")
    else:
        df = pd.read_excel(excel_path)
    df = df.fillna("")
    html_content = df.to_html(index=False)
    
    
    df = pd.read_excel(excel_path)
    df = df.fillna("")  # 清洗NaN
    html_content = df.to_html(index=False)
    html_extractor = HtmlExtractor(
        file=html_content,
        title_convert_to_markdown=True,
        cut_table_to_line=False
    )
    docs = html_extractor.extract()
    with open(md_path, "w", encoding="utf-8") as f:
        for d in docs:
            clean_content = d.page_content.replace("NaN", "")
            f.write(clean_content)
            f.write("\n\n")

def process_single_file(excel_file):
    """
    处理单个文件，生成同目录下的 md 文件。
    """
    md_file = os.path.splitext(excel_file)[0] + ".md"
    excel_to_markdown(excel_file, md_file)
    print(f"已生成: {md_file}")

def process_directory(excel_dir):
    """
    递归处理目录下所有 xlsx/xls 文件，生成平行的 md 目录，结构完全对等。
    """
    # 生成平行目录名
    parent_dir = os.path.dirname(os.path.abspath(excel_dir.rstrip("/")))
    base_name = os.path.basename(os.path.abspath(excel_dir.rstrip("/")))
    now = datetime.datetime.now().strftime("%Y%m%d-%H%M")
    md_dir = os.path.join(parent_dir, f"{base_name}-md解析-{now}")

    # 递归遍历
    for root, dirs, files in os.walk(excel_dir):
        # 计算当前目录在原目录下的相对路径
        rel_path = os.path.relpath(root, excel_dir)
        # 在平行 md 目录下创建对应子目录
        md_subdir = os.path.join(md_dir, rel_path)
        os.makedirs(md_subdir, exist_ok=True)
        for file in files:
            if file.lower().endswith((".xlsx", ".xls")):
                excel_file = os.path.join(root, file)
                md_file = os.path.join(md_subdir, os.path.splitext(file)[0] + ".md")
                try:
                    excel_to_markdown(excel_file, md_file)
                    print(f"已生成: {md_file}")
                except Exception as e:
                    print(f"处理 {excel_file} 时出错: {e}")

    print(f"\n所有文件已处理完毕，md 文件保存在: {md_dir}")

def main():
    if len(sys.argv) != 2:
        print("用法: python tests/test_extractor/test_excel-md_extractor.py <excel文件或目录>")
        sys.exit(1)
    input_path = sys.argv[1]
    if os.path.isfile(input_path) and input_path.lower().endswith((".xlsx", ".xls")):
        process_single_file(input_path)
    elif os.path.isdir(input_path):
        process_directory(input_path)
    else:
        print("请输入有效的 xlsx/xls 文件或目录路径。")
        sys.exit(1)

if __name__ == "__main__":
    main()