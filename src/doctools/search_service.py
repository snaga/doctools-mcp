import os
import re
import json
from whoosh.index import open_dir, exists_in
from whoosh.qparser import QueryParser
from whoosh.query import Prefix
from doctools.util_service import format_error_response

# Base directory environment variable for multi-index support
BASE_DIR_ENV_VAR = "WHOOSH_INDEX_BASE_DIR"

def validate_index_name(name: str) -> bool:
    """
    Validates the index name.
    Allowed characters: alphanumeric, hyphen (-), underscore (_).
    Prevents directory traversal.
    """
    if not name:
        return False
    # Only allow a-z, A-Z, 0-9, -, _
    return re.match(r'^[a-zA-Z0-9_-]+$', name) is not None

def resolve_index_path(base_dir: str, index_name: str) -> str:
    """
    Resolves the absolute path for an index given a base directory and index name.
    Raises ValueError if index_name is invalid.
    """
    if not validate_index_name(index_name):
        raise ValueError(f"Invalid index name: {index_name}. Only alphanumeric, hyphen, and underscore are allowed.")
    
    return os.path.join(os.path.abspath(base_dir), index_name)

def list_indexes(base_dir: str) -> dict:
    """
    Lists available indexes in the base directory with their descriptions.
    
    Args:
        base_dir: The base directory containing index subdirectories.
        
    Returns:
        dict: List of indexes with 'name' and 'description'.
    """
    if not base_dir or not os.path.exists(base_dir):
        return format_error_response(ValueError(f"Base directory not found or not set: {base_dir}"))
    
    indexes = []
    try:
        # Scan subdirectories
        with os.scandir(base_dir) as entries:
            for entry in entries:
                if entry.is_dir():
                    index_name = entry.name
                    # Skip hidden directories if any
                    if index_name.startswith('.'):
                        continue
                    
                    # Only include valid Whoosh index directories
                    if not exists_in(entry.path):
                        continue
                        
                    description = ""
                    extensions = []
                    meta_path = os.path.join(entry.path, "meta.json")
                    if os.path.exists(meta_path):
                        try:
                            with open(meta_path, "r", encoding="utf-8") as f:
                                meta = json.load(f)
                                description = meta.get("description", "")
                                extensions = meta.get("extensions", [])
                        except Exception:
                            # Ignore metadata read errors
                            pass
                            
                    indexes.append({
                        "name": index_name,
                        "description": description,
                        "extensions": extensions
                    })
        
        return {"status": "success", "indexes": indexes}
        
    except Exception as e:
        return format_error_response(e)

def search_index(query_str: str = "", filter_dir: str = None, 
                 base_dir: str = None, index_name: str = None) -> dict:
    """
    Search the Whoosh index for the given query string.
    Supports Boolean operators (AND, OR, NOT) for advanced filtering.
    
    Args:
        query_str: Keyword(s) to search for.
        filter_dir: Optional path to filter results by directory (prefix match on absolute path).
        base_dir: Base directory for multi-index support (Required).
        index_name: Name of the index to search within `base_dir`. Defaults to "default".
        
    Returns:
        dict: Result containing hits (path, filename, content summary, and score).
    """
    if not base_dir:
        return format_error_response(ValueError(f"Environment variable {BASE_DIR_ENV_VAR} is not set."))

    # Fallback to "default" if index_name is not provided
    target_index_name = index_name if index_name else "default"
    
    try:
        target_index_path = resolve_index_path(base_dir, target_index_name)
    except ValueError as e:
        return format_error_response(e)

    # Strict existence check
    if not os.path.exists(target_index_path) or not exists_in(target_index_path):
        return format_error_response(FileNotFoundError(
            f"Valid Whoosh index not found at: {target_index_path}. "
            f"Please ensure the index exists and is initialized."
        ))
        
    try:
        ix = open_dir(target_index_path)
        results_list = []
        
        with ix.searcher() as searcher:
            # content is the field name defined in get_schema()
            query = QueryParser("content", ix.schema).parse(query_str)
            
            # Create filter query if directory is specified
            filter_query = None
            if filter_dir:
                # Convert to absolute path to match index storage
                abs_filter_path = os.path.abspath(filter_dir)
                filter_query = Prefix("path", abs_filter_path)
            
            hits = searcher.search(query, limit=20, filter=filter_query)
            
            for hit in hits:
                # Get summary highlights
                raw_summary = hit.highlights("content", top=3)
                fancy_summary = None
                if raw_summary:
                    # Replace emphasis tags with brackets
                    fancy_summary = raw_summary.replace('<b class="inverted">', '【').replace('</b>', '】')
                    # Remove other HTML tags and normalize whitespace
                    fancy_summary = re.sub(r'<[^>]+>', '', fancy_summary)
                    fancy_summary = re.sub(r'\s+', ' ', fancy_summary).strip()
                
                path = hit["path"]
                results_list.append({
                    "path": path,
                    "filename": os.path.basename(path),
                    "content": fancy_summary,
                    "score": hit.score
                })
                
        return {
            "status": "success",
            "results": results_list
        }
        
    except Exception as e:
        return format_error_response(e)
