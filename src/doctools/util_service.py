import os
import zipfile
from jsonschema import validate, ValidationError

PAGEINDEX_SCHEMA = {
    "type": "object",
    "properties": {
        "title": {"type": "string"},
        "node_id": {"type": "string"},
        "summary": {"type": "string"},
        "start_index": {"type": "integer"},
        "end_index": {"type": "integer"},
        "sheet_name": {"type": "string"},
        "nodes": {
            "type": "array",
            "items": {"$ref": "#"}
        }
    },
    "required": ["title", "node_id", "summary"]
}

def validate_pageindex_json(data: dict) -> bool:
    """
    Validates the PageIndex JSON structure.
    
    Args:
        data (dict): The PageIndex JSON data to validate.
        
    Returns:
        bool: True if valid, False otherwise.
    """
    try:
        validate(instance=data, schema=PAGEINDEX_SCHEMA)
        return True
    except ValidationError:
        return False

def format_error_response(e: Exception) -> dict:
    """Formats an exception into a standardized error response dictionary."""
    return {
        "status": "error",
        "detail": str(e),
        "type": type(e).__name__
    }

def zip_files(file_paths: list[str], output_path: str = None) -> dict:
    """
    Compress multiple files and directories into a single ZIP archive.
    
    Args:
        file_paths: List of paths to files or directories to include in the ZIP.
        output_path: Path for the output ZIP file (optional).
    """
    if not file_paths:
        return format_error_response(ValueError("No files specified for compression."))
        
    # Validate input paths
    missing_paths = [p for p in file_paths if not os.path.exists(p)]
    if missing_paths:
        return format_error_response(FileNotFoundError(f"Paths not found: {missing_paths}"))
        
    try:
        # Generate default output path if not provided
        if output_path is None:
            first_path = file_paths[0]
            base, _ = os.path.splitext(first_path)
            output_path = f"{base}.zip"
            
        abs_output_path = os.path.abspath(output_path)
        
        with zipfile.ZipFile(abs_output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for path in file_paths:
                if os.path.isdir(path):
                    # For directories, walk through all subdirectories and files
                    # Preserve structure starting from the directory name itself
                    parent_dir = os.path.dirname(os.path.normpath(path))
                    for root, dirs, files in os.walk(path):
                        for file in files:
                            file_full_path = os.path.join(root, file)
                            # Create relative path from the parent of the target directory
                            arcname = os.path.relpath(file_full_path, parent_dir)
                            zipf.write(file_full_path, arcname=arcname)
                else:
                    # For individual files, preserve the provided path structure
                    # If it's absolute, make it relative to CWD to avoid drive letters in ZIP
                    if os.path.isabs(path):
                        arcname = os.path.relpath(path)
                    else:
                        arcname = path
                    zipf.write(path, arcname=arcname)
                
        return {
            "status": "success",
            "output_path": abs_output_path
        }
    except Exception as e:
        return format_error_response(e)

def filter_pageindex_tree(data: dict, node_id: str = None, depth: int = None) -> dict:
    """
    Filter the PageIndex JSON tree by node_id and/or depth.
    
    Args:
        data (dict): The PageIndex tree to filter.
        node_id (str, optional): The ID of the node to start the tree from.
        depth (int, optional): The maximum depth to include in the tree.
        
    Returns:
        dict: The filtered PageIndex tree.
    """
    def find_node(node: dict, target_id: str) -> dict:
        if node.get("node_id") == target_id:
            return node
        for child in node.get("nodes", []):
            result = find_node(child, target_id)
            if result:
                return result
        return None

    def prune_tree(node: dict, current_depth: int, max_depth: int) -> dict:
        # Shallow copy to avoid modifying original data
        new_node = {k: v for k, v in node.items() if k != "nodes"}
        
        if max_depth is not None and current_depth >= max_depth:
            new_node["nodes"] = []
            return new_node
            
        new_nodes = []
        for child in node.get("nodes", []):
            new_nodes.append(prune_tree(child, current_depth + 1, max_depth))
        
        new_node["nodes"] = new_nodes
        return new_node

    try:
        root_node = data
        if node_id:
            root_node = find_node(data, node_id)
            if not root_node:
                raise ValueError(f"Node ID '{node_id}' not found in the tree.")
        
        # depth 0 means just the node itself, no children.
        # depth 1 means node + its direct children (but children's children pruned).
        return prune_tree(root_node, 0, depth)
    except Exception as e:
        return format_error_response(e)

def unzip_file(zip_path: str, output_dir: str = None) -> dict:
    """
    Extract all files from a ZIP archive.
    
    Args:
        zip_path: Path to the source ZIP file.
        output_dir: Directory to extract files to (optional).
    """
    if not os.path.exists(zip_path):
        return format_error_response(FileNotFoundError(f"ZIP file not found: {zip_path}"))
        
    try:
        # Generate default output directory if not provided
        if output_dir is None:
            base, _ = os.path.splitext(zip_path)
            output_dir = base # Use ZIP filename as folder name
            
        abs_output_dir = os.path.abspath(output_dir)
        os.makedirs(abs_output_dir, exist_ok=True)
        
        extracted_files = []
        with zipfile.ZipFile(zip_path, 'r') as zipf:
            zipf.extractall(abs_output_dir)
            # List extracted files with absolute paths
            for name in zipf.namelist():
                extracted_files.append(os.path.join(abs_output_dir, name))
                
        return {
            "status": "success",
            "extracted_files": extracted_files
        }
    except zipfile.BadZipFile:
        return format_error_response(zipfile.BadZipFile("Invalid or corrupted ZIP file."))
    except Exception as e:
        return format_error_response(e)