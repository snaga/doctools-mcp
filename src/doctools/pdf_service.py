from pypdf import PdfReader, PdfWriter
import os
import fitz  # PyMuPDF
from doctools.util_service import format_error_response

def export_pages_to_images(
    input_path: str,
    output_dir: str = None,
    pages: list[int] = None,
    dpi: int = 150
) -> dict:
    """
    Export specified PDF pages as images (PNG) using PyMuPDF (fitz).

    Args:
        input_path (str): Path to the source PDF file.
        output_dir (str, optional): Directory to save the images. 
            If None, a directory based on the input filename is created.
        pages (list[int], optional): List of 1-based page numbers to export. 
            If None, all pages are exported.
        dpi (int, optional): Resolution for the output images. Defaults to 150.

    Returns:
        dict: Result of the operation including status and list of output paths.
            Example: {"status": "success", "output_paths": ["path/to/page_1.png", ...]}
            On failure: {"status": "error", "detail": "..."}
    """
    try:
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")

        # Determine output directory
        if output_dir is None:
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_dir = os.path.join(os.path.dirname(os.path.abspath(input_path)), f"{base_name}_images")
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        doc = fitz.open(input_path)
        total_pages = len(doc)
        
        # Determine pages to export
        if pages is None:
            target_pages = list(range(1, total_pages + 1))
        else:
            target_pages = pages

        output_paths = []
        # Matrix for DPI scaling
        zoom = dpi / 72  # 72 is the default PDF DPI
        matrix = fitz.Matrix(zoom, zoom)

        for p_num in target_pages:
            if p_num < 1 or p_num > total_pages:
                # Skip invalid page numbers or raise error? 
                # Let's raise error for consistency with other functions.
                raise ValueError(f"Page number {p_num} is out of range (1-{total_pages}).")
            
            # fitz uses 0-based indexing
            page = doc.load_page(p_num - 1)
            pix = page.get_pixmap(matrix=matrix)
            
            output_filename = f"page_{p_num:03d}.png"
            output_path = os.path.join(output_dir, output_filename)
            pix.save(output_path)
            output_paths.append(os.path.abspath(output_path))

        doc.close()

        return {
            "status": "success",
            "output_paths": output_paths
        }

    except Exception as e:
        return format_error_response(e)

def extract_pages(input_path: str, output_path: str, start_page: int, end_page: int) -> dict:
    """
    Extract specified pages from a PDF file and save as a new file.
    
    Args:
        input_path: Path to the source PDF file.
        output_path: Path to save the extracted pages.
        start_page: Starting page number (1-based).
        end_page: Ending page number (1-based).
        
    Returns:
        dict: Result of the operation including status and output path.
    """
    try:
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
            
        reader = PdfReader(input_path)
        total_pages = len(reader.pages)
        
        # Validation
        if start_page < 1:
            raise ValueError(f"Start page must be >= 1. Got: {start_page}")
        if end_page > total_pages:
            raise ValueError(f"End page {end_page} exceeds total pages {total_pages}.")
        if start_page > end_page:
            raise ValueError(f"Start page {start_page} cannot be greater than end page {end_page}.")

        writer = PdfWriter()
        
        # pypdf uses 0-based indexing
        for i in range(start_page - 1, end_page):
            writer.add_page(reader.pages[i])
            
        with open(output_path, "wb") as output_file:
            writer.write(output_file)
            
        return {
            "status": "success",
            "message": "PDF pages extracted successfully.",
            "output_path": output_path
        }
        
    except Exception as e:
        return format_error_response(e)

def extract_text_as_markdown(input_path: str, output_path: str = None, start_page: int = 1, end_page: int = None) -> dict:
    """
    Extract text from a PDF file and save it as a Markdown file.
    
    Args:
        input_path: Path to the PDF file.
        output_path: Path to the output Markdown file (optional).
        start_page: Starting page number (1-based, default 1).
        end_page: Ending page number (1-based, default None=last page).
        
    Returns:
        dict: Result containing the absolute path to the generated Markdown file.
    """
    try:
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
            
        reader = PdfReader(input_path)
        total_pages = len(reader.pages)
        
        if end_page is None:
            end_page = total_pages
            
        # Validation
        if start_page < 1:
            raise ValueError(f"Start page must be >= 1. Got: {start_page}")
        if end_page > total_pages:
            raise ValueError(f"End page {end_page} exceeds total pages {total_pages}.")
        if start_page > end_page:
            raise ValueError(f"Start page {start_page} cannot be greater than end page {end_page}.")
            
        extracted_text = []
        for i in range(start_page - 1, end_page):
            page = reader.pages[i]
            text = page.extract_text()
            if text:
                extracted_text.append(f"## Page {i+1}\n\n{text}")
            else:
                extracted_text.append(f"## Page {i+1}\n\n(No text extracted)")
        
        if output_path is None:
            output_path = os.path.splitext(input_path)[0] + ".md"
            
        content = "\n\n".join(extracted_text)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(content)
                
        return {
            "status": "success",
            "output_path": os.path.abspath(output_path)
        }

    except Exception as e:
        return format_error_response(e)

def get_page_count(input_path: str) -> dict:
    """
    Get the total number of pages in a PDF file.
    
    Args:
        input_path: Path to the PDF file.
        
    Returns:
        dict: Result containing page count.
    """
    try:
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
            
        reader = PdfReader(input_path)
        return {
            "status": "success",
            "page_count": len(reader.pages)
        }
        
    except Exception as e:
        return format_error_response(e)
