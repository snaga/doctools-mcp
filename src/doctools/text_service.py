import os
import re
from collections import deque
from charset_normalizer import from_path
import win32clipboard
import win32con
from doctools.util_service import format_error_response

def set_clipboard_text(text: str) -> dict:
    """
    Copy the specified text to the Windows clipboard.
    
    Args:
        text: The text to copy.
    """
    try:
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32con.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()
            
        return {
            "status": "success",
            "message": "Text copied to clipboard successfully."
        }
    except Exception as e:
        return format_error_response(e)

def convert_file_encoding(input_path: str, output_encoding: str, output_path: str = None, input_encoding: str = "utf-8") -> dict:
    """
    Convert a file's encoding to another encoding.
    
    Args:
        input_path: Source file path.
        output_encoding: Target encoding (e.g., 'utf-8', 'shift_jis').
        output_path: Target file path (optional, auto-generated if None).
        input_encoding: Source encoding (optional, default 'utf-8').
    """
    if not os.path.exists(input_path):
        return format_error_response(FileNotFoundError(f"File not found: {input_path}"))
        
    try:
        # 1. Detect input encoding ONLY if input_encoding is explicitly None
        detected_encoding = None
        if input_encoding is None:
            results = from_path(input_path)
            best_match = results.best()
            if best_match:
                input_encoding = best_match.encoding
                detected_encoding = input_encoding
            else:
                input_encoding = "utf-8" # Fallback
        
        # 2. Generate output path if not specified
        if output_path is None:
            base, ext = os.path.splitext(input_path)
            # Use a clean version of the encoding name for the filename
            enc_suffix = output_encoding.replace("-", "").replace("_", "").lower()
            output_path = f"{base}_{enc_suffix}{ext}"
            
        abs_output_path = os.path.abspath(output_path)
        
        # 3. Read and Write (Convert)
        with open(input_path, "r", encoding=input_encoding, errors="replace") as f_in:
            content = f_in.read()
            
        with open(abs_output_path, "w", encoding=output_encoding, newline="") as f_out:
            f_out.write(content)
            
        return {
            "status": "success",
            "output_path": abs_output_path,
            "input_encoding": input_encoding,
            "detected_encoding": detected_encoding,
            "output_encoding": output_encoding
        }
    except Exception as e:
        return format_error_response(e)

def read_head(path: str, n_lines: int = 10, encoding: str = "utf-8") -> dict:
    """
    Read the first n lines of a text file.
    
    Args:
        path: Path to the file.
        n_lines: Number of lines to read.
        encoding: File encoding (default 'utf-8').
    """
    if not os.path.exists(path):
        return format_error_response(FileNotFoundError(f"File not found: {path}"))
        
    try:
        lines = []
        with open(path, "r", encoding=encoding, errors="replace") as f:
            for i, line in enumerate(f):
                if i >= n_lines:
                    break
                lines.append(line.rstrip("\r\n"))
        return {"status": "success", "lines": lines}
    except Exception as e:
        return format_error_response(e)

def read_tail(path: str, n_lines: int = 10, encoding: str = "utf-8") -> dict:
    """
    Read the last n lines of a text file.
    
    Args:
        path: Path to the file.
        n_lines: Number of lines to read.
        encoding: File encoding (default 'utf-8').
    """
    if not os.path.exists(path):
        return format_error_response(FileNotFoundError(f"File not found: {path}"))
        
    try:
        # Use deque with maxlen to efficiently get the last n lines
        with open(path, "r", encoding=encoding, errors="replace") as f:
            lines = deque(f, maxlen=n_lines)
        return {"status": "success", "lines": [line.rstrip("\r\n") for line in lines]}
    except Exception as e:
        return format_error_response(e)

def grep_file(path: str, pattern: str, encoding: str = "utf-8") -> dict:
    """
    Search for a pattern in a text file and return matching lines with line numbers.
    
    Args:
        path: Path to the file.
        pattern: Regex pattern to search for.
        encoding: File encoding (default 'utf-8').
    """
    if not os.path.exists(path):
        return format_error_response(FileNotFoundError(f"File not found: {path}"))
        
    try:
        regex = re.compile(pattern)
        matches = []
        with open(path, "r", encoding=encoding, errors="replace") as f:
            for i, line in enumerate(f, 1):
                if regex.search(line):
                    matches.append({"line": i, "text": line.rstrip("\r\n")})
        return {"status": "success", "matches": matches}
    except Exception as e:
        return format_error_response(e)

def get_metadata(path: str) -> dict:
    """
    Get metadata for a text file: encoding (auto-detected), size, and line count.
    
    Args:
        path: Path to the file.
    """
    if not os.path.exists(path):
        return format_error_response(FileNotFoundError(f"File not found: {path}"))
        
    try:
        # 1. Get file size
        size = os.path.getsize(path)
        
        # 2. Detect encoding
        results = from_path(path)
        best_match = results.best()
        encoding = best_match.encoding if best_match else "utf-8"
        
        # 3. Count lines
        line_count = 0
        with open(path, "r", encoding=encoding, errors="replace") as f:
            for _ in f:
                line_count += 1
                
        return {
            "status": "success",
            "encoding": encoding,
            "size": size,
            "lines": line_count
        }
    except Exception as e:
        return format_error_response(e)
