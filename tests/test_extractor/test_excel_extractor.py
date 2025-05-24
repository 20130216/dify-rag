from dify_rag.extractor.excel_extractor import ExcelExtractor
from tests.log import logger

file_path = "tests/data/finance.xlsx"  # 请确保有这个测试文件

def test_excel_extractor():
    # 关键：让 HtmlExtractor 输出 markdown 格式
    extractor = ExcelExtractor(file_path)
    # 如果你想强制 markdown，可以在 ExcelExtractor 里加参数传递给 HtmlExtractor
    text_docs = extractor.extract()
    for d in text_docs:
        assert d.metadata
        assert d.page_content
        logger.info("----->")
        logger.info(f"Metadata: {d.metadata}")
        logger.info(f"{d.page_content} ({len(d.page_content)})")
        print(d.page_content)  # 直接打印，方便你肉眼检查是否为 markdown 格式

if __name__ == "__main__":
    test_excel_extractor()