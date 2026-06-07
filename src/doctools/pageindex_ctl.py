"""
PageIndex Controller CLI.
PageIndex の構築・保守を行うコマンドラインインターフェース。
"""

import argparse
import sys
import os
import json
import logging
import csv
from dotenv import load_dotenv
import google.generativeai as genai
import vertexai
from vertexai.generative_models import GenerativeModel as VertexGenerativeModel

# モジュールインポートの調整
try:
    from doctools import pdf_service, pptx_service, excel_service
    from doctools.util_service import validate_pageindex_json
except ImportError:
    # 実行環境によっては直接インポートが必要な場合がある
    import pdf_service
    import pptx_service
    import excel_service
    from util_service import validate_pageindex_json

# ロギング設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    stream=sys.stdout
)
logger = logging.getLogger(__name__)

# --- 定数 ---
PAGEINDEX_EXT = ".pageindex.json"
DEFAULT_EXTS = [".pdf", ".pptx", ".xlsx"]

# 黄金プロンプト (Builder 向け)
PROMPT_TEMPLATE = """あなたはドキュメント解析のエキスパートです。提供された「各ページのタイトルと冒頭テキスト」を元に、PageIndex形式のツリー構造JSONを作成してください。

【ドキュメント情報】
ファイル名: {filename}

【物理構造データ】
{structure_data}

【ルール】
1. 全体を意味のある章・節（10ノード程度の中間階層）にグルーピングしてください。
2. 各ノードの 'summary' には、検索性を高めるための具体的なキーワード（パラメータ名、手法、専門用語）を凝縮してください。
3. リーフノードは5ページ程度を目安に分割してください。
4. node_id の命名規則:
   - 章・節: ch_1, sec_1_1 等
   - PDF ページ: p_001 (ゼロパディング)
   - スライド: slide_01
   - シート: sheet_name
5. title の記載ルール: 物理的な位置情報を補足する（例: 「第1章: タイトル」, 「スライド12: タイトル」）。
6. 出力は以下のJSONスキーマに厳格に従ってください。
   スキーマ:
   {{
     "type": "object",
     "properties": {{
       "title": {{"type": "string"}},
       "node_id": {{"type": "string"}},
       "summary": {{"type": "string"}},
       "start_index": {{"type": "integer"}},
       "end_index": {{"type": "integer"}},
       "sheet_name": {{"type": "string"}},
       "nodes": {{ "type": "array", "items": {{ "$ref": "#" }} }}
     }},
     "required": ["title", "node_id", "summary"]
   }}

【出力】
JSONのみを出力してください。Markdownのバッククォートなどは不要です。
"""

def setup_llm_client():
    """LLM クライアントのセットアップ (Gemini API or Vertex AI)"""
    load_dotenv()
    
    use_vertexai = os.getenv("GOOGLE_GENAI_USE_VERTEXAI", "").lower() == "true"
    
    if use_vertexai:
        project_id = os.getenv("GOOGLE_CLOUD_PROJECT")
        location = os.getenv("GOOGLE_CLOUD_LOCATION")
        
        if not project_id or not location:
            raise ValueError("❌ Vertex AI を使用するには GOOGLE_CLOUD_PROJECT と GOOGLE_CLOUD_LOCATION が必要です。")
        
        vertexai.init(project=project_id, location=location)
        logger.info(f"✨ Vertex AI を初期化したよ (Project: {project_id}, Location: {location})")
        return VertexGenerativeModel("gemini-1.5-flash")
    else:
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("❌ GEMINI_API_KEY または GOOGLE_GENAI_USE_VERTEXAI が設定されていません。")
        
        genai.configure(api_key=api_key)
        logger.info("✨ Gemini API を初期化したよ")
        return genai.GenerativeModel('gemini-1.5-flash')

def generate_summary(model, filename, structure_data):
    """Gemini API または Vertex AI を用いた要約生成"""
    prompt = PROMPT_TEMPLATE.format(
        filename=filename, 
        structure_data=json.dumps(structure_data, ensure_ascii=False, indent=2)
    )
    
    # モデルの型でクライアントを判別
    client_type = "Vertex AI" if isinstance(model, VertexGenerativeModel) else "Gemini API"
    logger.info(f"📡 {client_type} に要約作成を依頼中... (データ量: {len(json.dumps(structure_data))} bytes)")
    try:
        response = model.generate_content(prompt)
        
        # トークン統計とクォータ情報のログ出力
        usage = getattr(response, "usage_metadata", None)
        if usage:
            logger.info(
                f"📊 Token Usage: Prompt={usage.prompt_token_count}, "
                f"Candidates={usage.candidates_token_count}, Total={usage.total_token_count}"
            )
            # クォータ情報は API 経由で直接取得できない場合が多いが、
            # メタデータに含まれる場合は出力する。
        
        content = response.text.strip()
        # Markdown の囲みを除去
        if content.startswith("```"):
            lines = content.splitlines()
            if lines[0].startswith("```"):
                content = "\n".join(lines[1:-1])
            else:
                content = content.strip("`")
        
        return json.loads(content.strip())
    except Exception as e:
        logger.error(f"❌ {client_type} 呼び出しまたは JSON 解析に失敗したよ: {e}")
        return None

def build_single_pageindex(model, file_path):
    """単一ファイルの PageIndex 構築"""
    ext = os.path.splitext(file_path)[1].lower()
    
    logger.info(f"🔍 物理構造を抽出中: {os.path.basename(file_path)}")
    
    if ext == ".pdf":
        res = pdf_service.extract_structure(file_path)
    elif ext == ".pptx":
        res = pptx_service.extract_structure(file_path)
    elif ext == ".xlsx":
        res = excel_service.extract_structure(file_path)
    else:
        logger.warning(f"⚠️ サポートされていない拡張子です: {ext}")
        return False
        
    if res.get("status") != "success":
        logger.error(f"❌ 構造抽出失敗: {res.get('detail')}")
        return False
        
    structure = res.get("structure")
    
    # 要約生成（バリデーション失敗時は最大2回リトライ）
    max_retries = 2
    for attempt in range(max_retries + 1):
        if attempt > 0:
            logger.info(f"🔄 バリデーション失敗のため再生成を試みるよ... ({attempt}/{max_retries})")
            
        pageindex_data = generate_summary(model, os.path.basename(file_path), structure)
        
        if pageindex_data:
            if validate_pageindex_json(pageindex_data):
                # 保存 (分散保存の原則: 元ファイルと同じディレクトリ)
                output_path = file_path + PAGEINDEX_EXT
                try:
                    with open(output_path, "w", encoding="utf-8") as f:
                        json.dump(pageindex_data, f, ensure_ascii=False, indent=2)
                    logger.info(f"✅ PageIndex を保存したよ: {output_path}")
                    return True
                except Exception as e:
                    logger.error(f"❌ ファイル保存失敗: {e}")
                    return False
            else:
                logger.warning(f"⚠️ 生成された JSON がスキーマに適合しないよ (Attempt {attempt+1})")
        else:
            logger.warning(f"⚠️ 要約の生成に失敗したよ (Attempt {attempt+1})")
            
    logger.error(f"❌ {file_path} の PageIndex 構築を諦めたよ...。")
    return False

def build_pageindex(args):
    """
    Build PageIndex for a document.
    物理構造抽出 ➔ LLM による要約生成 ➔ バリデーション ➔ 分散保存。
    """
    try:
        model = setup_llm_client()
    except Exception as e:
        logger.error(f"❌ LLM クライアントの初期化に失敗したため、構築を中止するよ: {e}")
        return
        
    target_path = os.path.abspath(args.path)
    logger.info(f"🚀 PageIndex の構築を開始するよ: {target_path}")
    
    # 対象拡張子の決定
    target_exts = DEFAULT_EXTS
    if args.ext:
        # ".pdf .pptx" 形式の文字列をリストに変換
        target_exts = [e.strip() if e.startswith(".") else f".{e.strip()}" for e in args.ext.replace(",", " ").split()]

    # 対象ファイルの収集
    files_to_process = []
    if os.path.isfile(target_path):
        files_to_process.append(target_path)
    elif os.path.isdir(target_path):
        if args.recursive:
            for root, _, files in os.walk(target_path):
                for f in files:
                    if any(f.lower().endswith(ext) for ext in target_exts):
                        files_to_process.append(os.path.join(root, f))
        else:
            for f in os.listdir(target_path):
                full_path = os.path.join(target_path, f)
                if os.path.isfile(full_path) and any(f.lower().endswith(ext) for ext in target_exts):
                    files_to_process.append(full_path)
    else:
        logger.error(f"❌ パスが見つからないよ: {target_path}")
        return

    if not files_to_process:
        logger.warning("対象となるファイルが見つからなかったよ。拡張子設定を確認してね。")
        return

    logger.info(f"📂 {len(files_to_process)} 件のファイルを処理するよ...")

    success_count = 0
    for f in files_to_process:
        if build_single_pageindex(model, f):
            success_count += 1
            
    logger.info(f"✨ 構築完了！成功: {success_count}/{len(files_to_process)}")
    
    # 構築完了後、自動的に export 処理を呼び出し
    logger.info("📤 サマリーを自動更新（export）するよ...")
    export_pageindex(args)


def list_pageindex(args):
    """
    List PageIndex files.
    指定ディレクトリ内の .pageindex.json ファイル一覧表示。
    """
    target_dir = os.path.abspath(args.path)
    logger.info(f"📂 PageIndex ファイルを検索中: {target_dir}")
    
    found_files = []
    if os.path.isdir(target_dir):
        if args.recursive:
            for root, _, files in os.walk(target_dir):
                for f in files:
                    if f.endswith(PAGEINDEX_EXT):
                        found_files.append(os.path.join(root, f))
        else:
            for f in os.listdir(target_dir):
                if f.endswith(PAGEINDEX_EXT):
                    found_files.append(os.path.join(target_dir, f))
    elif os.path.isfile(target_dir) and target_dir.endswith(PAGEINDEX_EXT):
        found_files.append(target_dir)
                
    for f in found_files:
        print(f)
    
    logger.info(f"✨ 合計 {len(found_files)} 個のインデックスが見つかったよ！")


def delete_pageindex(args):
    """
    Delete PageIndex for a document.
    指定ファイルまたはディレクトリ配下の .pageindex.json 削除。
    """
    target_path = os.path.abspath(args.path)
    logger.info(f"🗑️  PageIndex の削除準備: {target_path}")
    
    files_to_delete = []
    if os.path.isfile(target_path):
        if target_path.endswith(PAGEINDEX_EXT):
            files_to_delete.append(target_path)
        else:
            idx_path = target_path + PAGEINDEX_EXT
            if os.path.exists(idx_path):
                files_to_delete.append(idx_path)
    elif os.path.isdir(target_path):
        if args.recursive:
            for root, _, files in os.walk(target_path):
                for f in files:
                    if f.endswith(PAGEINDEX_EXT):
                        files_to_delete.append(os.path.join(root, f))
        else:
            for f in os.listdir(target_path):
                if f.endswith(PAGEINDEX_EXT):
                    files_to_delete.append(os.path.join(target_path, f))

    if not files_to_delete:
        logger.info("削除対象の PageIndex ファイルは見つからなかったよ。")
        return

    for f in files_to_delete:
        try:
            os.remove(f)
            logger.info(f"🔥 削除しました: {f}")
        except Exception as e:
            logger.error(f"❌ 削除失敗: {f} ({e})")
            
    logger.info(f"✨ 合計 {len(files_to_delete)} 個のファイルを削除したよ！")


def export_pageindex(args):
    """
    Export PageIndex to CSV.
    全インデックス情報の CSV エクスポート。
    """
    target_path = os.path.abspath(args.path)
    if os.path.isfile(target_path):
        base_search_dir = os.path.dirname(target_path)
    else:
        base_search_dir = target_path

    logger.info(f"📤 PageIndex のエクスポート準備: {base_search_dir}")
    
    recursive = getattr(args, "recursive", True)
    
    found_files = []
    if os.path.isdir(base_search_dir):
        if recursive:
            for root, _, files in os.walk(base_search_dir):
                for f in files:
                    if f.endswith(PAGEINDEX_EXT):
                        found_files.append(os.path.join(root, f))
        else:
            for f in os.listdir(base_search_dir):
                if f.endswith(PAGEINDEX_EXT):
                    found_files.append(os.path.join(base_search_dir, f))
    elif os.path.isfile(target_path) and target_path.endswith(PAGEINDEX_EXT):
        found_files.append(target_path)

    if not found_files:
        logger.warning("エクスポート対象の PageIndex ファイルが見つからなかったよ。")
        return

    output_csv = os.path.join(base_search_dir, "pageindex_summary.csv")
    
    try:
        with open(output_csv, "w", encoding="utf-8", newline="") as f:
            # 全フィールドをダブルクォートで括る設定
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerow(["filepath", "title", "summary"])
            
            for idx_file in found_files:
                try:
                    with open(idx_file, "r", encoding="utf-8") as jf:
                        data = json.load(jf)
                        writer.writerow([
                            idx_file,
                            data.get("title", ""),
                            data.get("summary", "").replace("\n", " ")
                        ])
                except Exception as e:
                    logger.error(f"❌ {idx_file} の読み込みに失敗したよ: {e}")

        logger.info(f"✅ {output_csv} にサマリー情報をエクスポートしたよ！✨")
        print(os.path.abspath(output_csv))
    except Exception as e:
        logger.error(f"❌ CSV 書き込みに失敗したよ: {e}")


def main():
    """
    Main entry point for PageIndex Controller.
    """
    # Windows等の環境でUTF-8出力を強制する
    if sys.stdout.encoding.lower() != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except AttributeError:
            pass

    parser = argparse.ArgumentParser(description="PageIndex Controller CLI - PageIndex の構築・保守を行うよ！✨")
    subparsers = parser.add_subparsers(dest="command", help="利用可能なコマンド")

    # build command
    build_parser = subparsers.add_parser("build", help="PageIndex を構築するよ")
    build_parser.add_argument("path", help="対象ファイルまたはディレクトリのパス")
    build_parser.add_argument("--output", help="出力先パス（省略時は対象と同じ場所に .pageindex.json を作るよ）")
    build_parser.add_argument("--recursive", action="store_true", help="ディレクトリを再帰的に処理するよ")
    build_parser.add_argument("--ext", help='対象とする拡張子（例: ".pdf .pptx"）')
    build_parser.set_defaults(func=build_pageindex)

    # list command
    list_parser = subparsers.add_parser("list", help=".pageindex.json ファイルを一覧表示するよ")
    list_parser.add_argument("path", help="検索対象のディレクトリパス")
    list_parser.add_argument("--recursive", action="store_true", help="再帰的に検索するよ")
    list_parser.set_defaults(func=list_pageindex)

    # delete command
    delete_parser = subparsers.add_parser("delete", help=".pageindex.json を削除するよ")
    delete_parser.add_argument("path", help="対象ファイルまたはディレクトリのパス")
    delete_parser.add_argument("--recursive", action="store_true", help="ディレクトリ配下を再帰的に削除するよ")
    delete_parser.set_defaults(func=delete_pageindex)

    # export command
    export_parser = subparsers.add_parser("export", help="インデックス情報を CSV にエクスポートするよ")
    export_parser.add_argument("path", help="走査対象のディレクトリパス")
    export_parser.add_argument("--recursive", action="store_true", help="再帰的に検索するよ")
    export_parser.set_defaults(func=export_pageindex)

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    if hasattr(args, "func"):
        args.func(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
