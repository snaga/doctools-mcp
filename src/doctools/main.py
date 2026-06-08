from fastmcp import FastMCP
from dotenv import load_dotenv, find_dotenv
import sys
import os
import json
from doctools import __version__

# Load environment variables (explicitly from CWD to support global install)
load_dotenv(find_dotenv(usecwd=True))

from doctools.pdf_service import (
    extract_pages as pdf_extract_pages_svc,
    extract_text_as_markdown as pdf_extract_markdown_svc,
    get_page_count as pdf_get_page_count,
    export_pages_to_images as pdf_export_images_svc
)
from doctools.pptx_service import (
    count_slides as pptx_count_slides_svc,
    extract_text_as_markdown as pptx_extract_markdown_svc,
    extract_slides as pptx_extract_slides_svc,
    export_slides_to_images as pptx_export_images_svc,
    merge_pptx as pptx_merge_svc
)
from doctools.excel_service import (
    list_sheets as xlsx_list_sheets_svc,
    extract_csv as xlsx_extract_csv_svc,
    extract_sheet as xlsx_extract_sheet_svc,
    extract_markdown_to_file as xlsx_extract_markdown_svc,
    export_sheets_to_images as xlsx_export_images_svc
)
from doctools.csv_service import (
    csv_read_cells as csv_read_cells_svc,
    csv_search_values as csv_search_values_svc,
    csv_get_metadata as csv_get_metadata_svc,
    csv_extract_to_file as csv_extract_to_file_svc
)
from doctools.html_service import extract_text_as_markdown as html_extract_markdown_svc
from doctools.image_service import (
    get_image_metadata as image_get_metadata_svc,
    crop_image as image_crop_svc,
    save_clipboard_image as image_save_clipboard_svc
)
from doctools.pageindex_service import (
    get_tree as pageindex_get_tree_svc,
    get_node_content as pageindex_get_node_content_svc
)
from doctools.text_service import (
    convert_file_encoding as convert_encoding_svc, 
    read_head as read_head_svc, 
    read_tail as read_tail_svc, 
    grep_file as grep_file_svc,
    get_metadata as get_metadata_svc,
    set_clipboard_text as set_clipboard_text_svc
)
from doctools.util_service import zip_files as zip_files_svc, unzip_file as unzip_file_svc, format_error_response, get_summary_tree as get_summary_tree_svc
from doctools.search_service import search_index, list_indexes as list_indexes_svc, BASE_DIR_ENV_VAR

mcp = FastMCP("doctools-mcp")

# --- Helper to return JSON string ---
def to_json(data: dict) -> str:
    return json.dumps(data, ensure_ascii=False)

# --- Search Tools ---

@mcp.tool
def list_indexes() -> str:
    """
    List available search indexes and their descriptions.
    
    This tool allows you to discover which indexes are available for searching.
    It returns a list of index names and their descriptions (if available).
    
    Returns:
        JSON string containing a list of indexes.
        
    Example:
        Output: {"status": "success", "indexes": [{"name": "project_a", "description": "Docs for Project A"}, ...]}
    """
    try:
        base_dir = os.getenv(BASE_DIR_ENV_VAR)
        if not base_dir:
            return to_json({
                "status": "error", 
                "detail": f"{BASE_DIR_ENV_VAR} environment variable is not set.",
                "type": "EnvironmentError"
            })
        return to_json(list_indexes_svc(base_dir))
    except Exception as e:
        return to_json(format_error_response(e))

@mcp.tool
def search_documents(query: str, directory: str = None, index_name: str = None) -> str:
    """Search for Office documents using a pre-built Whoosh index.
    
    You can use Boolean operators (AND, OR, NOT) for advanced filtering.
    
    Examples:
    - '"user authentication"' : Search for the exact phrase.
    - 'user AND authentication' : Search for documents containing both words.
    - 'error OR exception' : Search for documents containing either word.
    - 'test NOT unit' : Search for 'test' but exclude documents containing 'unit'.
    - '(user OR account) AND (login OR auth)' : Complex logic.
    
    Args:
        query: Search keyword(s) with optional Boolean operators.
        directory: Optional path to filter results. 
                   Only documents within this directory (and its subdirectories) will be returned.
                   This filter uses **Prefix Match** on the absolute path of the document.
                   It is recommended to use an absolute path or a path relative to the project root.
                   Examples:
                   - 'src/doctools' (matches 'src/doctools/main.py', 'src/doctools/utils/helper.py')
                   - 'SPECS' (matches 'SPECS/design.md')
                   - 'C:/Users/name/Documents' (absolute path)
        index_name: Optional name of the index to search. 
                    Searches within `{WHOOSH_INDEX_BASE_DIR}/{index_name}`.
                    If omitted, defaults to "default".
    """
    try:
        base_dir = os.getenv(BASE_DIR_ENV_VAR)
        
        return to_json(search_index(
            query_str=query, 
            filter_dir=directory,
            base_dir=base_dir,
            index_name=index_name
        ))
    except Exception as e:
        return to_json(format_error_response(e))

# --- PageIndex Tools ---

@mcp.tool
def pageindex_get_summary_tree(target_dir: str, depth: int = 2) -> str:
    """
    Load {target_dir}/pageindex.json and recursively filter out children deeper than depth.
    Root is depth 0.
    
    Args:
        target_dir: The directory containing pageindex.json.
        depth: The maximum depth to include in the tree (default 2).
    """
    return to_json(get_summary_tree_svc(target_dir, depth))

@mcp.tool
def pageindex_get_tree(input_path: str, node_id: str = None, depth: int = 2) -> str:
    """
    PageIndex JSON を読み込み、ドキュメントの目次（ツリー構造）を返却します。
    
    大規模なドキュメントの構造を段階的に探索するのに適しています。
    最初は node_id 未指定でルートから取得し、必要に応じて特定の node_id を指定して深掘りしてください。
    
    Args:
        input_path: 対象ドキュメントのパス（例: 'path/to/doc.pdf'）。
                    '.pageindex.json' が同じ場所に存在する必要があります。
        node_id: 取得を開始する特定のノードID（任意）。
        depth: 取得する階層の深さ（デフォルト: 2）。
        
    Example:
        1. まず全体像を把握: pageindex_get_tree(path)
        2. 第3章の詳細が見たい: pageindex_get_tree(path, node_id='ch_3', depth=1)
    """
    return to_json(pageindex_get_tree_svc(input_path, node_id, depth))

@mcp.tool
def pageindex_get_node_content(
    file_path: str,
    node_type: str,
    node_id: str = None,
    pages: list[int] = None,
    sheet_name: str = None
) -> str:
    """
    ドキュメントの指定された箇所（ノード、ページ範囲、またはシート）からフルテキストを抽出します。
    
    優先順位:
    1. node_id が指定されている場合、PageIndex から該当箇所の範囲（ページ等）を特定して抽出します。
    2. node_id が無い場合、pages (PDF/PPTX) または sheet_name (Excel) を直接使用します。
    
    Args:
        file_path: 対象ドキュメントのパス。
        node_type: ドキュメントの種類 ('pdf', 'pptx', 'xlsx')。
        node_id: PageIndex 内のノードID（任意）。指定すると pages/sheet_name は無視されます。
        pages: 抽出したいページ番号またはスライド番号のリスト (1-based, 任意)。
        sheet_name: 抽出したい Excel のシート名（任意）。
        
    Example:
        - ノードIDで抽出: pageindex_get_node_content(path, 'pdf', node_id='sec_2_1')
        - 直接ページ指定: pageindex_get_node_content(path, 'pdf', pages=[1, 2, 5])
    """
    return to_json(pageindex_get_node_content_svc(file_path, node_type, node_id, pages, sheet_name))

# --- PDF Tools ---

@mcp.tool
def pdf_extract_pages(input_path: str, output_path: str, start_page: int, end_page: int) -> str:
    """Extract specified pages from a PDF file and save as a new file.
    
    Args:
        input_path: Path to the source PDF file.
        output_path: Path to save the extracted pages.
        start_page: Starting page number (1-based).
        end_page: Ending page number (1-based).
    """
    return to_json(pdf_extract_pages_svc(input_path, output_path, start_page, end_page))

@mcp.tool
def pdf_extract_markdown(input_path: str, output_path: str = None, start_page: int = 1, end_page: int = None) -> str:
    """Extract text from a PDF file and save it as a Markdown file.
    
    Args:
        input_path: Path to the PDF file.
        output_path: Path to the output Markdown file (optional).
        start_page: Starting page number (1-based, default 1).
        end_page: Ending page number (1-based, default None=last page).
    """
    return to_json(pdf_extract_markdown_svc(input_path, output_path, start_page, end_page))

@mcp.tool
def pdf_count_pages(input_path: str) -> str:
    """Get the total number of pages in a PDF file.
    
    Args:
        input_path: Path to the PDF file.
    """
    return to_json(pdf_get_page_count(input_path))

@mcp.tool
def pdf_extract_images(input_path: str, output_dir: str = None, pages: list[int] = None, dpi: int = 150) -> str:
    """Export specified PDF pages as images (PNG).
    
    Args:
        input_path: Path to the source PDF file.
        output_dir: Directory to save the images (optional).
        pages: List of 1-based page numbers to export (optional, default all).
        dpi: Resolution for the output images (default 150).
    """
    return to_json(pdf_export_images_svc(input_path, output_dir, pages, dpi))

# --- PowerPoint Tools ---

@mcp.tool
def pptx_count_slides(input_path: str) -> str:
    """Get the total number of slides in a PowerPoint file.
    
    Args:
        input_path: Path to the PPTX file.
    """
    return to_json(pptx_count_slides_svc(input_path))

@mcp.tool
def pptx_extract_markdown(input_path: str, output_path: str = None, start_slide: int = 1, end_slide: int = None) -> str:
    """Extract text from a PowerPoint file as Markdown and save to a file.
    
    Args:
        input_path: Path to the PPTX file.
        output_path: Path to the output Markdown file (optional).
        start_slide: Starting slide number (1-based, default 1).
        end_slide: Ending slide number (1-based, default None=last slide).
    """
    return to_json(pptx_extract_markdown_svc(input_path, output_path, start_slide, end_slide))

@mcp.tool
def pptx_extract_slides(input_path: str, output_path: str, start_slide: int, end_slide: int) -> str:
    """Extract specified slides from a PowerPoint file and save as a new file.
    
    Args:
        input_path: Path to the source PPTX file.
        output_path: Path to save the extracted slides.
        start_slide: Starting slide number (1-based).
        end_slide: Ending slide number (1-based).
    """
    return to_json(pptx_extract_slides_svc(input_path, output_path, start_slide, end_slide))

@mcp.tool
def pptx_extract_images(input_path: str, output_dir: str = None, slides: list[int] = None, width: int = 1280, height: int = 720) -> str:
    """Export specified slides from a PowerPoint file as images (PNG).
    
    Args:
        input_path: Path to the PPTX file.
        output_dir: Directory to save the images (optional).
        slides: List of 1-based slide numbers to export (optional, default all).
        width: Width of the exported images (default 1280).
        height: Height of the exported images (default 720).
    """
    return to_json(pptx_export_images_svc(input_path, output_dir, slides, width, height))

@mcp.tool
def pptx_merge(input_paths: list[str], output_path: str) -> str:
    """Merge multiple PowerPoint files into a single file.
    
    This tool combines several PPTX files into one, appending them in the order provided.
    It uses Windows PowerPoint to ensure that formatting and slide masters are preserved.
    
    Args:
        input_paths: A list of absolute paths to the source PPTX files.
                     Example: ["C:/docs/part1.pptx", "C:/docs/part2.pptx"]
        output_path: The absolute path where the merged PPTX file will be saved.
    """
    if not isinstance(input_paths, list):
        return to_json(format_error_response(TypeError("input_paths must be a list.")))
    if not input_paths:
        return to_json(format_error_response(ValueError("input_paths list cannot be empty.")))
        
    return to_json(pptx_merge_svc(input_paths, output_path))

# --- Excel Tools ---

@mcp.tool
def xlsx_list_sheets(input_path: str) -> str:
    """Get list of sheets in an Excel file.
    
    Args:
        input_path: Path to the XLSX file.
    """
    return to_json(xlsx_list_sheets_svc(input_path))

@mcp.tool
def xlsx_extract_markdown(input_path: str, output_path: str = None, sheet_name: str = None) -> str:
    """Extract data from an Excel sheet as Markdown table and save to a file.
    
    Args:
        input_path: Path to the XLSX file.
        output_path: Path to the output Markdown file (optional).
        sheet_name: Name of the sheet to extract (optional, defaults to first sheet).
    """
    return to_json(xlsx_extract_markdown_svc(input_path, output_path, sheet_name))

@mcp.tool
def xlsx_extract_csv(input_path: str, output_dir: str = None, sheet_names: list[str] = None, encoding: str = "shift_jis") -> str:
    """Extract specified Excel sheets (or all) as individual CSV files.
    
    Args:
        input_path: Path to the XLSX file.
        output_dir: Directory to save the CSV files (optional). 
                    Defaults to input file's directory.
        sheet_names: List of sheet names to extract (optional). 
                     Defaults to all sheets.
        encoding: Encoding for the output CSV files (optional, defaults to 'shift_jis').
    """
    return to_json(xlsx_extract_csv_svc(input_path, output_dir, sheet_names, encoding))

@mcp.tool
def xlsx_extract_sheet(input_path: str, output_path: str, sheet_name: str) -> str:
    """Extract a specific sheet from an Excel file and save as a new XLSX file.
    
    Args:
        input_path: Path to the source XLSX file.
        output_path: Path to save the new XLSX file.
        sheet_name: Name of the sheet to extract.
    """
    return to_json(xlsx_extract_sheet_svc(input_path, output_path, sheet_name))

@mcp.tool
def xlsx_extract_images(input_path: str, output_dir: str = None, sheet_names: list[str] = None, dpi: int = 150) -> str:
    """Export specified Excel sheets as images (PNG).
    
    This tool uses Windows Excel to render sheets accurately, ensuring that 
    tables and charts are captured as they appear in the application.
    It automatically optimizes the layout to fit the width of a single page.
    
    Args:
        input_path: Path to the source Excel file.
        output_dir: Directory to save the images (optional). 
                    Defaults to a new directory named after the input file.
        sheet_names: List of sheet names to export (optional). 
                     Defaults to the active sheet only.
        dpi: Resolution for the output images (default 150).
    """
    return to_json(xlsx_export_images_svc(input_path, output_dir, sheet_names, dpi))

# --- CSV Tools ---

@mcp.tool
def csv_read_cells(path: str, start_row: int = 1, end_row: int = None, columns: list[int] = None, header_row: int = 0, encoding: str = "shift_jis") -> str:
    """Read specific rows and columns from a CSV file and return as a compact JSON list.
    
    Args:
        path: CSV file path.
        start_row: 1-based start row index (default 1).
        end_row: 1-based end row index (inclusive, default None=last row).
        columns: List of 0-based column indices (optional, default None=all columns).
        header_row: 1-based row index to use as header (default 0=no header).
        encoding: File encoding (default 'shift_jis').
    """
    return to_json(csv_read_cells_svc(path, start_row, end_row, columns, header_row, encoding))

@mcp.tool
def csv_extract_to_file(input_path: str, output_path: str = None, start_row: int = 1, end_row: int = None, columns: list[int] = None, encoding: str = "shift_jis") -> str:
    """Extract specific rows and columns from a CSV file and save to another CSV file.
    
    Args:
        input_path: Source CSV file path.
        output_path: Target CSV file path (optional, auto-generated if None).
        start_row: 1-based start row index (default 1).
        end_row: 1-based end row index (inclusive, default None=last row).
        columns: List of 0-based column indices (optional, default None=all columns).
        encoding: Encoding for both reading and writing (default 'shift_jis').
    """
    return to_json(csv_extract_to_file_svc(input_path, output_path, start_row, end_row, columns, encoding))

@mcp.tool
def csv_search_values(path: str, query: str, encoding: str = "shift_jis") -> str:
    """Search for a string in a CSV file and return cell positions.
    
    Args:
        path: CSV file path.
        query: String to search for.
        encoding: File encoding (default 'shift_jis').
    """
    return to_json(csv_search_values_svc(path, query, encoding))

@mcp.tool
def csv_get_metadata(path: str) -> str:
    """Get CSV metadata: encoding, total rows, and max columns.
    
    Args:
        path: CSV file path.
    """
    return to_json(csv_get_metadata_svc(path))

# --- HTML Tools ---

@mcp.tool
def html_extract_markdown(input_path: str, output_path: str = None) -> str:
    """Extract text from an HTML file and save it as a Markdown file.
    
    Args:
        input_path: Path to the source HTML file.
        output_path: Path to the output Markdown file (optional).
    """
    return to_json(html_extract_markdown_svc(input_path, output_path))

# --- Image Tools ---

@mcp.tool
def image_get_metadata(path: str) -> str:
    """Get metadata (width and height) of an image file.
    
    Args:
        path: Path to the image file.
    """
    return to_json(image_get_metadata_svc(path))

@mcp.tool
def image_crop(path: str, left: int, top: int, right: int, bottom: int, output_path: str = None) -> str:
    """Crop an image with the specified coordinates.
    
    Args:
        path: Path to the source image file.
        left: The left boundary of the crop box.
        top: The top boundary of the crop box.
        right: The right boundary of the crop box.
        bottom: The bottom boundary of the crop box.
        output_path: Path for the output cropped image (optional).
    """
    return to_json(image_crop_svc(path, left, top, right, bottom, output_path))

@mcp.tool
def image_save_clipboard(output_dir: str = None, filename: str = None) -> str:
    """
    Windows クリップボードから画像を取得し、PNG ファイルとして保存します。

    Args:
        output_dir: 保存先のディレクトリパス（任意）。未指定時はカレントディレクトリ。
        filename: 保存するファイル名（任意）。未指定時は 'clipboard_YYYYMMDD_HHMMSS.png' を自動生成。

    Returns:
        JSON string containing the status and the absolute path of the saved image.
        
    Example:
        Output: {"status": "success", "message": "Image saved...", "output_path": "C:/path/to/clipboard_20260428_120000.png"}
    """
    return to_json(image_save_clipboard_svc(output_dir, filename))

# --- Text Tools ---

@mcp.tool
def text_read_head(path: str, n_lines: int = 10, encoding: str = "utf-8") -> str:
    """
    Read the first n lines of a text file.
    
    This tool allows you to peek at the beginning of a file to understand its structure or content
    without reading the entire file into memory.
    
    Args:
        path: The absolute path to the text file.
        n_lines: The number of lines to read from the beginning. Defaults to 10.
        encoding: The encoding of the file. Defaults to 'utf-8'.
    
    Returns:
        JSON string containing the lines read.
        
    Example:
        Input: path="C:/data/log.txt", n_lines=5
        Output: {"status": "success", "lines": ["2023-01-01 Start", "2023-01-01 Init", ...]}
    """
    return to_json(read_head_svc(path, n_lines, encoding))

@mcp.tool
def text_read_tail(path: str, n_lines: int = 10, encoding: str = "utf-8") -> str:
    """
    Read the last n lines of a text file.
    
    Useful for checking the latest entries in log files or confirming file termination.
    
    Args:
        path: The absolute path to the text file.
        n_lines: The number of lines to read from the end. Defaults to 10.
        encoding: The encoding of the file. Defaults to 'utf-8'.
        
    Returns:
        JSON string containing the lines read.
        
    Example:
        Input: path="C:/data/app.log", n_lines=3
        Output: {"status": "success", "lines": ["Error: ...", "Shutdown complete", "Exited code 0"]}
    """
    return to_json(read_tail_svc(path, n_lines, encoding))

@mcp.tool
def text_grep(path: str, pattern: str, encoding: str = "utf-8") -> str:
    """
    Search for a regex pattern in a text file and return matching lines.
    
    Args:
        path: The absolute path to the text file.
        pattern: The regular expression pattern to search for.
        encoding: The encoding of the file. Defaults to 'utf-8'.
        
    Returns:
        JSON string containing a list of matches (line number and text).
        
    Example:
        Input: path="C:/src/code.py", pattern="TODO"
        Output: {"status": "success", "matches": [{"line": 10, "text": "# TODO: fix bug"}, ...]}
    """
    return to_json(grep_file_svc(path, pattern, encoding))

@mcp.tool
def text_get_metadata(path: str) -> str:
    """
    Get metadata for a text file: encoding (auto-detected), size, and line count.
    
    This helps in understanding the file properties before processing it further.
    
    Args:
        path: The absolute path to the text file.
        
    Returns:
        JSON string containing encoding, size (bytes), and total lines.
        
    Example:
        Input: path="C:/data/large.csv"
        Output: {"status": "success", "encoding": "utf-8", "size": 1048576, "lines": 50000}
    """
    return to_json(get_metadata_svc(path))

@mcp.tool
def text_convert_encoding(input_path: str, output_encoding: str, output_path: str = None, input_encoding: str = "utf-8") -> str:
    """
    Convert a text file's encoding to another encoding.
    
    Args:
        input_path: The absolute path to the source text file.
        output_encoding: The target encoding (e.g., 'utf-8', 'shift_jis').
        output_path: The path where the converted file will be saved. If omitted, a path is generated automatically.
        input_encoding: The source encoding. Defaults to 'utf-8'. Set to None to attempt auto-detection (slower).
        
    Returns:
        JSON string containing the output path and used encodings.
        
    Example:
        Input: input_path="data.csv", output_encoding="utf-8", input_encoding="shift_jis"
        Output: {"status": "success", "output_path": "data_utf8.csv", ...}
    """
    res = convert_encoding_svc(input_path, output_encoding, output_path, input_encoding)
    return to_json(res)

@mcp.tool
def text_copy_clipboard(text: str) -> str:
    """
    Copy the specified text to the Windows clipboard.
    
    This tool is useful for quickly copying generated content, paths, or extracted text 
    to the system clipboard for use in other applications.
    
    Args:
        text: The text string to be copied to the clipboard.
        
    Returns:
        JSON string containing the status and a success message.
        
    Example:
        Input: text="Hello from MCP!"
        Output: {"status": "success", "message": "Text copied to clipboard successfully."}
    """
    return to_json(set_clipboard_text_svc(text))

# --- Utility Tools ---

@mcp.tool
def zip_files(file_paths: list[str], output_path: str = None) -> str:
    """Compress multiple files into a single ZIP archive.
    
    Args:
        file_paths: A list of paths to the files you want to compress.
        output_path: The path where the ZIP file will be saved. If omitted, it will be generated automatically.
    """
    res = zip_files_svc(file_paths, output_path)
    return to_json(res)

@mcp.tool
def unzip_file(zip_path: str, output_dir: str = None) -> str:
    """Extract all files from a ZIP archive.
    
    Args:
        zip_path: The path to the ZIP file to extract.
        output_dir: The directory where the files will be extracted. If omitted, a directory with the ZIP file's name will be created.
    """
    res = unzip_file_svc(zip_path, output_dir)
    return to_json(res)


def main():
    if len(sys.argv) > 1 and sys.argv[1] in ["--version", "-v"]:
        print(f"doctools-mcp v{__version__}")
        return

    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
