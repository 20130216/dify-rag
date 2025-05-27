# File: tests/test_extractor/test_excel_extractor.py
# (Note: Renamed from i_test_excel_extractor.py for consistency if moving to final tests)

import os
from dify_rag.extractor.excel_extractor import ExcelExtractor
from tests.log import logger # Assuming tests/log.py exists and provides a logger

# Define the path to the Excel file with multiple sheets
# Please ensure this file exists at the specified path for the test to run correctly.
file_path = "tests/data/小桃管理团队岗位职责.xlsx"

# Define the output Markdown file path
# It will be in the same directory as the source Excel file, with a .md extension.
output_md_file_path = os.path.splitext(file_path)[0] + ".md"

def test_excel_extractor():
    """
    Test the ExcelExtractor to ensure it can parse multi-sheet Excel files
    and extract content into Document objects, then write to a Markdown file.
    """
    logger.info(f"Starting Excel extraction test for: {file_path}")

    try:
        # Initialize the ExcelExtractor with the file path
        extractor = ExcelExtractor(file_path=file_path)

        # Extract documents from the Excel file
        text_docs = extractor.extract()

        # Assert that documents were extracted
        assert text_docs, "No documents were extracted from the Excel file."
        logger.info(f"Successfully extracted {len(text_docs)} documents.")

        # Prepare content for writing to Markdown
        all_extracted_content = []
        for i, d in enumerate(text_docs):
            assert d.metadata is not None, f"Document {i} is missing metadata."
            assert d.page_content is not None, f"Document {i} is missing page_content."
            assert len(d.page_content) > 0, f"Document {i} has empty page_content."

            logger.info("-----> Document %d" % (i + 1))
            logger.info(f"Metadata: {d.metadata}")
            # Truncate content for logging to avoid excessive output
            display_content = d.page_content[:500] + "..." if len(d.page_content) > 500 else d.page_content
            logger.info(f"Content: {display_content} (Length: {len(d.page_content)})")

            # Collect content for Markdown file
            # Add a clear separator between contents from different "documents" (which might correspond to sheets or parts of sheets)
            if all_extracted_content: # Add separator if not the first document
                all_extracted_content.append("\n\n---\n\n") # Markdown horizontal rule as separator
            
            # Optionally add a title for each document if available in metadata
            # This depends on how HtmlExtractor formats titles. If it embeds them into page_content,
            # then simple page_content is enough.
            if "titles" in d.metadata and d.metadata["titles"]:
                # The titles in metadata are typically a list of hierarchical titles
                # We can choose to add them explicitly, or rely on page_content's internal formatting.
                # For this example, let's assume page_content already contains formatted content.
                # If you need to explicitly add titles from metadata, uncomment and adjust below:
                # all_extracted_content.append(f"# {d.metadata['titles'][0] if d.metadata['titles'] else 'Untitled'}\n\n")
                pass # Rely on content itself for structure

            all_extracted_content.append(d.page_content)

        # Join all collected content into a single string
        final_markdown_output = "".join(all_extracted_content)

        # Write the combined content to the Markdown file
        logger.info(f"Attempting to write extracted content to: {output_md_file_path}")
        with open(output_md_file_path, 'w', encoding='utf-8') as f:
            f.write(final_markdown_output)
        logger.info(f"Successfully wrote extracted content to: {output_md_file_path}")

    except FileNotFoundError:
        logger.error(f"Error: The file '{file_path}' was not found. Please ensure it exists.")
        raise
    except IOError as e:
        logger.error(f"Error writing to Markdown file {output_md_file_path}: {e}")
        raise
    except Exception as e:
        logger.error(f"An unexpected error occurred during Excel extraction or file writing: {e}")
        raise

if __name__ == "__main__":
    # This block allows running the test directly from the command line
    # For a more robust test setup, consider using pytest.
    logger.info("Running test_excel_extractor directly.")
    test_excel_extractor()
    logger.info("test_excel_extractor completed successfully.")
