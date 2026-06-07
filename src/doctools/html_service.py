import os
from markitdown import MarkItDown
from doctools.util_service import format_error_response

def extract_text_as_markdown(input_path: str, output_path: str = None) -> dict:
    """
    Extract text from an HTML file and save it as a Markdown file.
    
    Args:
        input_path: Path to the source HTML file.
        output_path: Path to the output Markdown file (optional).
        
    Returns:
        dict: Result of the operation including status and output path.
    """
    try:
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
            
        mid = MarkItDown()
        result = mid.convert(input_path)
        
        if output_path is None:
            output_path = os.path.splitext(input_path)[0] + ".md"
            
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(result.text_content)
            
        return {
            "status": "success",
            "output_path": os.path.abspath(output_path)
        }
        
    except Exception as e:
        return format_error_response(e)
