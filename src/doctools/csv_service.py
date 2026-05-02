import csv
import os
from charset_normalizer import from_path
from doctools.util_service import format_error_response

def get_csv_encoding(path: str) -> dict:
    """
    Detect the encoding of a CSV file.
    
    Returns:
        dict: Result with encoding name and confidence.
    """
    try:
        results = from_path(path)
        best_match = results.best()
        if best_match:
            return {
                "encoding": best_match.encoding,
                "confidence": best_match.confidence
            }
        return {"encoding": "utf-8", "confidence": 0.0} # Fallback
    except Exception:
        return {"encoding": "shift_jis", "confidence": 0.0} # Common fallback for JP

def csv_get_metadata(path: str) -> dict:
    """
    Get CSV metadata: encoding, total rows, and max columns.
    """
    if not os.path.exists(path):
        return format_error_response(FileNotFoundError(f"File not found: {path}"))
        
    try:
        enc_info = get_csv_encoding(path)
        encoding = enc_info["encoding"]
        
        row_count = 0
        max_cols = 0
        
        with open(path, "r", encoding=encoding, errors="replace") as f:
            reader = csv.reader(f)
            for row in reader:
                row_count += 1
                max_cols = max(max_cols, len(row))
                
        return {
            "status": "success",
            "encoding": encoding,
            "total_rows": row_count,
            "max_columns": max_cols
        }
    except Exception as e:
        return format_error_response(e)

def csv_read_cells(path: str, start_row: int = 1, end_row: int = None, columns: list[int] = None, header_row: int = 0, encoding: str = "shift_jis") -> dict:
    """
    Read specific rows and columns from a CSV file and return as a compact JSON list.
    
    Args:
        path: CSV file path.
        start_row: 1-based start row index.
        end_row: 1-based end row index (inclusive).
        columns: List of 0-based column indices.
        header_row: 1-based row index to use as header (0 for no header).
        encoding: File encoding.
    """
    if not os.path.exists(path):
        return format_error_response(FileNotFoundError(f"File not found: {path}"))
        
    try:
        result_data = []
        header_labels = []
        
        with open(path, "r", encoding=encoding, errors="replace") as f:
            reader = csv.reader(f)
            
            # 1. Handle header if specified
            # We might need to peek or read ahead if header_row > end_row, 
            # but usually header is at the top.
            current_file_row = 0
            
            # Since we iterate sequentially, let's keep track of rows.
            # To handle header_row potentially being anywhere, we need a way to find it.
            # Simplest approach: iterate and collect.
            
            rows_to_fetch = []
            max_needed_row = end_row if end_row else float('inf')
            if header_row > 0:
                max_needed_row = max(max_needed_row, header_row)
            
            target_rows = set()
            if end_row:
                target_rows.update(range(start_row, end_row + 1))
            else:
                # If end_row is None, we don't know the end, so we read from start_row onwards.
                pass

            for row in reader:
                current_file_row += 1
                
                # Capture header if this is the row
                if current_file_row == header_row:
                    header_labels = row
                
                # Capture data row if within range
                is_in_range = current_file_row >= start_row and (end_row is None or current_file_row <= end_row)
                if is_in_range:
                    rows_to_fetch.append((current_file_row, row))
                
                # Optimization: break if we passed both end_row and header_row
                if end_row and current_file_row >= max_needed_row:
                    break
            
            # 2. Process collected rows with column filtering
            # Prepare column labels
            final_headers = ["row_num"]
            col_indices = []
            
            # Determine which columns to pick
            if columns:
                col_indices = columns
            else:
                # Use all columns from the first data row or header
                sample_row = rows_to_fetch[0][1] if rows_to_fetch else (header_labels if header_labels else [])
                col_indices = list(range(len(sample_row)))
            
            for ci in col_indices:
                label = f"col_{ci}"
                if header_row > 0 and ci < len(header_labels):
                    label = header_labels[ci]
                final_headers.append(label)
            
            result_data.append(final_headers)
            
            # Format data rows
            for row_num, row_content in rows_to_fetch:
                # Skip the header row itself if it's within the data range
                if row_num == header_row:
                    continue
                    
                formatted_row = [row_num]
                for ci in col_indices:
                    val = row_content[ci] if ci < len(row_content) else ""
                    formatted_row.append(val)
                result_data.append(formatted_row)
                
        return {
            "status": "success",
            "data": result_data
        }
    except Exception as e:
        return format_error_response(e)

def csv_search_values(path: str, query: str, encoding: str = "shift_jis") -> dict:
    """
    Search for a string in a CSV file and return cell positions.
    """
    if not os.path.exists(path):
        return format_error_response(FileNotFoundError(f"File not found: {path}"))
        
    try:
        results = []
        with open(path, "r", encoding=encoding, errors="replace") as f:
            reader = csv.reader(f)
            current_row = 0
            for row in reader:
                current_row += 1
                for col_idx, value in enumerate(row):
                    if query.lower() in value.lower():
                        results.append({
                            "row": current_row,
                            "col": col_idx,
                            "value": value
                        })
        return {
            "status": "success",
            "results": results
        }
    except Exception as e:
        return format_error_response(e)

def csv_extract_to_file(input_path: str, output_path: str = None, start_row: int = 1, end_row: int = None, columns: list[int] = None, encoding: str = "shift_jis") -> dict:
    """
    Extract specific rows and columns from a CSV file and save to another CSV file.
    
    Args:
        input_path: Source CSV file path.
        output_path: Target CSV file path (optional).
        start_row: 1-based start row index.
        end_row: 1-based end row index (inclusive).
        columns: List of 0-based column indices.
        encoding: Encoding for both reading and writing.
    """
    if not os.path.exists(input_path):
        return format_error_response(FileNotFoundError(f"File not found: {input_path}"))
        
    try:
        if output_path is None:
            base, _ = os.path.splitext(input_path)
            output_path = f"{base}_extracted.csv"
            
        abs_output_path = os.path.abspath(output_path)
        
        with open(input_path, "r", encoding=encoding, errors="replace") as fin:
            reader = csv.reader(fin)
            
            with open(abs_output_path, "w", encoding=encoding, newline="") as fout:
                writer = csv.writer(fout)
                
                current_file_row = 0
                for row in reader:
                    current_file_row += 1
                    
                    # Optimization: break if we passed end_row
                    if end_row and current_file_row > end_row:
                        break
                        
                    # Skip if before start_row
                    if current_file_row < start_row:
                        continue
                        
                    # Extract specific columns if specified
                    if columns:
                        extracted_row = [row[ci] if ci < len(row) else "" for ci in columns]
                    else:
                        extracted_row = row
                        
                    writer.writerow(extracted_row)
                    
        return {
            "status": "success",
            "output_path": abs_output_path
        }
    except Exception as e:
        return format_error_response(e)
