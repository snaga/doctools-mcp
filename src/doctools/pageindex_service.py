import os
import json
from typing import Optional, List, Any, Union
from doctools.util_service import format_error_response
from doctools.pdf_service import extract_node_content as pdf_extract_node_content
from doctools.pptx_service import extract_node_content as pptx_extract_node_content
from doctools.excel_service import extract_node_content as excel_extract_node_content

PAGEINDEX_EXT = ".pageindex.json"

def get_tree(input_path: str, node_id: Optional[str] = None, depth: int = 2) -> dict:
    """
    PageIndex JSON を読み込み、指定ノードまたは階層までのサブツリーを返却する。
    
    Args:
        input_path: 対象ドキュメントのパス。
        node_id: 取得を開始するノードのID。省略時はルート。
        depth: 取得する階層の深さ。デフォルトは 2。
        
    Returns:
        dict: 成功時 {"status": "success", "tree": {...}}
    """
    try:
        index_path = input_path + PAGEINDEX_EXT
        if not os.path.exists(index_path):
            return {
                "status": "error",
                "detail": f"PageIndex file not found: {index_path}",
                "type": "FileNotFoundError"
            }

        with open(index_path, "r", encoding="utf-8") as f:
            tree = json.load(f)

        if node_id:
            # Find target node
            target_node = _find_node(tree, node_id)
            if not target_node:
                return {
                    "status": "error",
                    "detail": f"Node ID '{node_id}' not found in {index_path}.",
                    "type": "ValueError"
                }
            tree = target_node

        # Filter by depth
        filtered_tree = _filter_tree(tree, depth)

        return {
            "status": "success",
            "tree": filtered_tree
        }
    except Exception as e:
        return format_error_response(e)

def _find_node(node: dict, node_id: str) -> Optional[dict]:
    """再帰的に node_id を持つノードを探索するよ。"""
    if node.get("node_id") == node_id:
        return node
    
    for child in node.get("nodes", []):
        found = _find_node(child, node_id)
        if found:
            return found
    return None

def _filter_tree(node: dict, depth: int) -> dict:
    """指定された深さまでノードをフィルタリングするよ。"""
    # 子ノード以外をコピー
    result = {k: v for k, v in node.items() if k != "nodes"}
    
    if depth > 0 and "nodes" in node:
        result["nodes"] = [_filter_tree(child, depth - 1) for child in node["nodes"]]
    # depth が 0 の場合や nodes がない場合は nodes キーを含めない
        
    return result

def get_node_content(
    file_path: str,
    node_type: str,
    node_id: Optional[str] = None,
    pages: Optional[List[int]] = None,
    sheet_name: Optional[str] = None
) -> dict:
    """
    指定箇所のフルテキストを抽出し、返却する。
    優先順位: 1. node_id (インデックス解決), 2. pages/sheet_name (直接指定)
    
    Args:
        file_path: 対象ファイルのパス。
        node_type: ノードのタイプ ('pdf', 'pptx', 'xlsx' など)。
        node_id: インデックス内のノードID。
        pages: ページ番号またはスライド番号のリスト (1-based)。
        sheet_name: Excelのシート名。
        
    Returns:
        dict: 抽出されたテキスト内容。
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        # 1. Resolve node_id if provided
        if node_id:
            index_path = file_path + PAGEINDEX_EXT
            if os.path.exists(index_path):
                with open(index_path, "r", encoding="utf-8") as f:
                    tree = json.load(f)
                
                node = _find_node(tree, node_id)
                if node:
                    # インデックス情報からパラメータを上書きまたは補充
                    # node_type に基づいて適切なフィールドを採用するよ
                    if node_type == "pdf":
                        start = node.get("start_index")
                        end = node.get("end_index")
                        if start is not None and end is not None:
                            pages = list(range(start, end + 1))
                    elif node_type == "pptx":
                        start = node.get("start_index")
                        end = node.get("end_index")
                        if start is not None and end is not None:
                            pages = list(range(start, end + 1))
                    elif node_type == "xlsx":
                        sheet_name = node.get("sheet_name") or node.get("node_id") # node_id をシート名として使っている場合への配慮
                else:
                    return {
                        "status": "error",
                        "detail": f"Node ID '{node_id}' not found in index.",
                        "type": "ValueError"
                    }
            else:
                return {
                    "status": "error",
                    "detail": f"PageIndex file not found for {file_path}. Please build it first.",
                    "type": "FileNotFoundError"
                }

        # 2. Extract content based on resolved or direct parameters
        if node_type == "pdf":
            return pdf_extract_node_content(file_path, pages)
        elif node_type == "pptx":
            return pptx_extract_node_content(file_path, slides=pages)
        elif node_type == "xlsx":
            return excel_extract_node_content(file_path, sheet_name=sheet_name)
        else:
            raise ValueError(f"Unsupported node type: {node_type}")

    except Exception as e:
        return format_error_response(e)
