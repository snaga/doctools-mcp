import os
import re
import json
import argparse
import time
import sys
from datetime import datetime
import multiprocessing
from queue import Empty
from markitdown import MarkItDown
from whoosh.index import open_dir, create_in, exists_in
from whoosh.fields import Schema, ID, TEXT, NUMERIC
from whoosh.analysis import NgramAnalyzer
from whoosh.query import Prefix
from doctools.search_service import resolve_index_path, BASE_DIR_ENV_VAR, list_indexes

# デフォルトの対象拡張子
DEFAULT_EXTENSIONS = ('.pptx', '.xlsx', '.docx', '.pdf', '.txt', '.md')

# インデックス作成時と同じスキーマを使用
def get_schema():
    jp_analyzer = NgramAnalyzer(minsize=2, maxsize=4)
    return Schema(
        path=ID(stored=True, unique=True),
        mtime=NUMERIC(stored=True),
        content=TEXT(stored=True, analyzer=jp_analyzer)
    )

class ErrorLogger:
    def __init__(self, index_dir):
        self.index_dir = index_dir
        self.errors = []
    
    def log(self, file_path, error_type, message):
        self.errors.append({
            "path": file_path,
            "type": error_type,
            "message": str(message),
            "timestamp": datetime.now().isoformat()
        })
    
    def save(self):
        if not self.errors:
            return
        path = os.path.join(self.index_dir, "errors.json")
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.errors, f, indent=2, ensure_ascii=False)
            print(f"⚠️ エラーログを保存したよ: {path}")
        except Exception as e:
            print(f"❌ エラーログの保存に失敗: {e}")

def _worker_convert(full_path, queue):
    """子プロセスで実行される変換ワーカー"""
    # テスト用の特別な挙動: パスに 'infinite_loop' が含まれる場合は無限ループ
    if "infinite_loop" in full_path:
        while True:
            time.sleep(0.1)

    try:
        # MarkItDownはプロセスごとに初期化（Pickle問題を回避）
        mid = MarkItDown()
        result = mid.convert(full_path)
        # 成功時はテキスト内容を返す
        queue.put({"status": "success", "text": result.text_content})
    except Exception as e:
        # 失敗時はエラー情報を返す
        queue.put({"status": "error", "type": type(e).__name__, "message": str(e)})

def convert_with_timeout(full_path, timeout):
    """MarkItDown変換を別プロセスで実行し、タイムアウト制御する"""
    queue = multiprocessing.Queue()
    process = multiprocessing.Process(target=_worker_convert, args=(full_path, queue))
    
    process.start()
    
    try:
        # 💡 join の前に get(timeout=...) を使うことでデッドロックを回避するよ
        result = queue.get(timeout=timeout)
        
        # 正常終了したプロセスの後始末
        process.join(0.5)
        if process.is_alive():
            process.terminate()
            
        if result["status"] == "success":
            return result["text"]
        else:
            error_msg = f"{result['type']}: {result['message']}"
            raise Exception(error_msg)
            
    except Empty:
        # タイムアウト発生時
        if process.is_alive():
            process.terminate()
            process.join(0.5)
            if process.is_alive():
                try:
                    process.kill()
                except AttributeError:
                    pass
        raise TimeoutError(f"Processing timed out after {timeout} seconds")
    except Exception as e:
        if process.is_alive():
            process.terminate()
        raise e
    finally:
        # 確実にゾンビプロセスを残さないようにするよ
        if process.is_alive():
            process.join(0.1)

def save_metadata(index_path, description=None, extensions=None):
    """Saves index metadata to meta.json."""
    if not description and not extensions:
        return
        
    meta_path = os.path.join(index_path, "meta.json")
    meta = {}
    
    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)
        except Exception:
            pass
            
    if description:
        meta["description"] = description
        print(f"📝 説明を保存したよ: {description}")
        
    if extensions:
        meta["extensions"] = list(extensions)
        print(f"📁 対象拡張子を保存したよ: {', '.join(extensions)}")

    meta["updated_at"] = datetime.now().isoformat()
    if "created_at" not in meta:
        meta["created_at"] = meta["updated_at"]
    
    try:
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"⚠️ メタデータの保存に失敗: {e}")

def get_base_dir(args):
    """Resolves base directory from args or env var."""
    base_dir = args.base_dir or os.getenv(BASE_DIR_ENV_VAR)
    if not base_dir:
        raise ValueError(f"Base directory not specified. Use --base-dir or set {BASE_DIR_ENV_VAR}.")
    return base_dir

# --- 同期機能 (削除ファイルの検出と削除) --- 
def sync_index(ix, target_dir, dry_run=False):
    """ファイルシステムに存在しないドキュメントをインデックスから削除する"""
    print(f"\n🔄 同期チェック中 (削除対象の確認)...")
    abs_target_dir = os.path.abspath(target_dir)
    to_delete = []

    with ix.searcher() as searcher:
        for fields in searcher.all_stored_fields():
            path = fields.get('path')
            if not path:
                continue
            if not path.startswith(abs_target_dir):
                continue
            if not os.path.exists(path):
                to_delete.append(path)

    if not to_delete:
        print("✅ 削除対象のファイルはありません。")
        return

    print(f"⚠️ {len(to_delete)} 件のファイルが見つかりません。インデックスから削除します。")
    if dry_run:
        for p in to_delete:
            print(f"  [Dry-run] Delete: {p}")
        return

    writer = ix.writer()
    for p in to_delete:
        writer.delete_by_term('path', p)
        print(f"🗑️  削除: {os.path.basename(p)}")
    writer.commit()
    print("✨ 同期完了: インデックスをクリーンアップしたよ！")


# --- ドキュメント処理 (追加・更新) --- 
def process_documents(ix, docs_path, timeout, dry_run=False, error_logger=None, valid_exts=None):
    writer = ix.writer()
    abs_docs = os.path.abspath(docs_path)
    if valid_exts is None:
        valid_exts = DEFAULT_EXTENSIONS
    
    print(f"🔄 スキャン開始: {abs_docs}")
    print(f"📋 対象拡張子: {', '.join(valid_exts)}")
    
    count_processed = 0
    count_skipped = 0
    count_error = 0
    
    indexed_mtimes = {}
    with ix.searcher() as searcher:
        for fields in searcher.all_stored_fields():
            path = fields.get('path')
            mtime = fields.get('mtime')
            if path and mtime is not None:
                indexed_mtimes[path] = mtime

    for root, dirs, files in os.walk(abs_docs):
        for file in files:
            if file.lower().endswith(valid_exts):
                full_path = os.path.abspath(os.path.join(root, file))
                current_mtime = os.path.getmtime(full_path)
                
                if full_path in indexed_mtimes:
                    if current_mtime <= indexed_mtimes[full_path]:
                        count_skipped += 1
                        continue
                
                if dry_run:
                    action = "Update" if full_path in indexed_mtimes else "Add"
                    print(f"  [Dry-run] {action}: {file}")
                    count_processed += 1
                    continue

                print(f"📖 処理中: {file} ...", end="", flush=True)
                start_time = time.time()
                try:
                    text_content = convert_with_timeout(full_path, timeout)
                    clean_text = re.sub(r'[#*`|:>-]', ' ', text_content)
                    
                    writer.update_document(
                        path=full_path,
                        mtime=current_mtime,
                        content=clean_text
                    )
                    
                    elapsed = time.time() - start_time
                    print(f" 完了！ ({elapsed:.2f}s)")
                    count_processed += 1
                    
                except Exception as e:
                    elapsed = time.time() - start_time
                    print(f" ❌ 失敗 ({elapsed:.2f}s): {e}")
                    count_error += 1
                    if error_logger:
                        error_logger.log(full_path, type(e).__name__, str(e))

    if not dry_run:
        writer.commit()
    else:
        writer.cancel()
    
    print("-" * 40)
    print(f"📊 処理結果:")
    print(f"   ✅ 処理・更新: {count_processed} 件")
    print(f"   ⏭️  スキップ (変更なし): {count_skipped} 件")
    print(f"   ❌ エラー: {count_error} 件")
    print("-" * 40)


# --- 1. ビルド機能 (統合コマンド) --- 
def build_index(index_path, docs_path, description=None, extensions=None, timeout=30, dry_run=False, no_sync=False):
    start_total = time.time()
    error_logger = ErrorLogger(index_path)

    if not os.path.exists(index_path):
        if dry_run:
            print(f"🚀 [Dry-run] 新規インデックス作成予定: {index_path}")
            pass 
        else:
            os.makedirs(index_path)
    
    ix = None
    try:
        # メタデータの読み込み（既存の拡張子設定を確認するため）
        meta_path = os.path.join(index_path, "meta.json")
        existing_exts = None
        if os.path.exists(meta_path):
            try:
                with open(meta_path, "r", encoding="utf-8") as f:
                    meta = json.load(f)
                    existing_exts = meta.get("extensions")
            except Exception:
                pass

        # 拡張子の決定優先順位: 1. 引数, 2. メタデータ, 3. デフォルト
        valid_exts = extensions or existing_exts or DEFAULT_EXTENSIONS
        valid_exts = tuple(valid_exts) # os.walk用

        if exists_in(index_path):
            ix = open_dir(index_path)
            print(f"📂 既存インデックスを開きました: {index_path}")
        else:
            if dry_run:
                print("⚠️ インデックスが存在しないため、新規作成としてシミュレーションします。")
                from whoosh.ramindex import RamStorage
                ix = RamStorage().create_index(get_schema())
            else:
                ix = create_in(index_path, get_schema())
                print(f"🚀 インデックスを新規作成しました: {index_path}")

        # メタデータの保存（説明または拡張子が指定された場合）
        if not dry_run:
            save_metadata(index_path, description=description, extensions=valid_exts)
        
        process_documents(ix, docs_path, timeout, dry_run, error_logger, valid_exts=valid_exts)
        
        if not no_sync:
            sync_index(ix, docs_path, dry_run)
        
        if not dry_run:
            error_logger.save()

    finally:
        if ix:
            ix.close()

    total_elapsed = time.time() - start_total
    print(f"\n✨ 全工程完了！ 所要時間: {total_elapsed:.2f}秒")


# --- 2. 一覧表示機能 --- 
def list_files(index_path):
    ix = open_dir(index_path)
    try:
        with ix.searcher() as searcher:
            print(f"\n📂 インデックスに登録済みのファイル一覧 ({index_path}):")
            print("-" * 60)
            count = 0
            for fields in searcher.all_stored_fields():
                path = fields['path']
                filename = os.path.basename(path)
                print(f"📄 {filename}")
                print(f"   🔗 {path}")
                count += 1
            print("-" * 60)
            print(f"✨ 合計 {count} 件あるよ！")
    finally:
        ix.close()

# --- 3. 削除機能（Prefix対応） --- 
def delete_path(index_path, target_path):
    ix = open_dir(index_path)
    try:
        writer = ix.writer()
        abs_target = os.path.abspath(target_path)
        writer.delete_by_query(Prefix("path", abs_target))
        writer.commit(optimize=True)
        print(f"🗑️  削除完了: {abs_target} 配下のデータを消したよ！")
    finally:
        ix.close()


# --- メイン処理 (CLI設定) --- 
def main():
    # Windows等の環境でUTF-8出力を強制する
    if sys.stdout.encoding.lower() != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except AttributeError:
            pass

    parser = argparse.ArgumentParser(description="検索インデックス管理ツール")
    subparsers = parser.add_subparsers(dest="command", help="実行する操作", required=True)

    def add_common_args(p, required=True):
        if required:
            p.add_argument("index_name", help="操作対象のインデックス名")
        else:
            p.add_argument("index_name", nargs='?', help="操作対象のインデックス名 (未指定時は一覧表示)")
        p.add_argument("--base-dir", help=f"インデックスのベースディレクトリ (未指定時は {BASE_DIR_ENV_VAR})")

    build_parser = subparsers.add_parser("build", help="インデックスの構築（作成・更新・同期）")
    build_parser.add_argument("target", help="スキャン対象のドキュメントディレクトリパス")
    build_parser.add_argument("--description", help="インデックスの説明")
    build_parser.add_argument("--ext", nargs='+', help="対象とする拡張子 (例: .py .sql). 未指定時は既存設定またはデフォルトを使用")
    build_parser.add_argument("--timeout", type=int, default=30, help="1ファイルあたりの処理タイムアウト(秒)")
    build_parser.add_argument("--dry-run", action="store_true", help="実際の変更を行わずに確認する")
    build_parser.add_argument("--no-sync", action="store_true", help="削除されたファイルの同期（削除）をスキップする")
    add_common_args(build_parser)

    list_parser = subparsers.add_parser("list", help="登録済みファイルの一覧表示、またはインデックス一覧表示")
    add_common_args(list_parser, required=False)

    del_parser = subparsers.add_parser("delete", help="指定パス(Prefix)の削除")
    del_parser.add_argument("target", help="削除したいディレクトリまたはファイルのパス")
    add_common_args(del_parser)

    args = parser.parse_args()

    try:
        base_dir = get_base_dir(args)
        # index_name がない場合は None になる (listコマンドのみ)
        index_path = resolve_index_path(base_dir, args.index_name) if args.index_name else None
    except ValueError as e:
        print(f"❌ エラー: {e}")
        return

    if args.command == "list" and not args.index_name:
        # インデックス一覧を表示
        result = list_indexes(base_dir)
        if result["status"] == "success":
            print(f"\n利用可能なインデックス一覧 (ベース: {base_dir}):")
            print("-" * 80)
            if not result["indexes"]:
                print(" (インデックスはまだありません)")
            for idx in result["indexes"]:
                name = idx["name"]
                desc = idx["description"] or "(説明なし)"
                exts = ", ".join(idx.get("extensions", [])) or "デフォルト"
                # 拡張子リストが長すぎる場合は切り詰める
                if len(exts) > 30:
                    exts = exts[:27] + "..."
                
                print(f"📁 {name:15} | 📝 {desc:25} | 📄 {exts}")
            print("-" * 80)
        else:
            print(f"❌ エラー: {result.get('detail', '一覧の取得に失敗したよ')}")
        return

    if args.command != "build" and not os.path.exists(index_path):
        print(f"❌ インデックスが見つからないよ！: {index_path}")
        return

    if args.command == "build":
        # 拡張子の正規化 (.py or py -> .py)
        normalized_exts = None
        if args.ext:
            normalized_exts = [e if e.startswith('.') else f".{e}" for e in args.ext]
            normalized_exts = [e.lower() for e in normalized_exts]

        # Windowsでmultiprocessingを使用する場合、freeze_supportが必要（メインプロセス内）
        # ただし、通常のスクリプト実行時は main() が guard されていれば良い。
        build_index(
            index_path, 
            args.target, 
            description=args.description, 
            extensions=normalized_exts,
            timeout=args.timeout, 
            dry_run=args.dry_run, 
            no_sync=args.no_sync
        )
    elif args.command == "list":
        list_files(index_path)
    elif args.command == "delete":
        delete_path(index_path, args.target)

if __name__ == "__main__":
    # Windowsでの子プロセス無限生成を防ぐ
    multiprocessing.freeze_support()
    main()