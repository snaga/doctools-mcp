# DocTools MCP

Gemini-CLI 等の AI エージェントが、PDF, PowerPoint, Excel, CSV, HTML などのローカルドキュメントを自由に閲覧・検索・加工できるようにするための MCP (Model Context Protocol) サーバーです。

## 概要

このプロジェクトは、ドキュメント操作に特化した MCP サーバーを提供します。
大容量データの扱いを考慮し、抽出結果をファイルとして保存したり、コンパクトなデータ構造で返却したりすることで、LLM のコンテキストウィンドウを効率的に利用できるように設計されています。

## 主要機能

### DocTools サーバー (`doctools-mcp`)

#### 全文検索ツール
- `search_documents`: 作成済みのインデックスを使用して、Office 文書を高速に検索します。
    - **マルチインデックス対応**: 複数のインデックス（プロジェクトごとなど）を切り替えて検索可能。
    - **Boolean Query**: `AND`, `OR`, `NOT`, `""` を用いた検索に対応。
    - **ディレクトリ絞り込み**: `directory` 引数により、特定のフォルダ配下のみを対象とした絞り込み（Faceting）が可能（前方一致）。
- `list_indexes`: 利用可能なインデックスの一覧と、それぞれの概要（Description）を取得します。

#### PDF ツール
- `pdf_extract_markdown`: PDF からテキストを抽出し、Markdown ファイルとして保存します。
- `pdf_extract_images`: PDF の特定ページを画像 (PNG/JPG) として保存します。マルチモーダル LLM による図解・数式の理解に最適です。
- `pdf_extract_pages`: PDF の特定ページを別ファイルとして切り出します。
- `pdf_count_pages`: PDF の総ページ数を取得します。

#### PowerPoint ツール
- `pptx_extract_markdown`: スライド内容（表を含む）を抽出し、Markdown ファイルとして保存します。
- `pptx_extract_images`: スライドを画像 (PNG/JPG) として保存します。システム構成図やグラフの理解に強力な力を発揮します。
- `pptx_count_slides`: 総スライド数を取得します。
- `pptx_extract_slides`: 指定範囲のスライドを別ファイルとして保存します。

#### Excel ツール
- `xlsx_list_sheets`: 全てのシート名を取得します。
- `xlsx_extract_markdown`: 指定シートの内容を Markdown テーブル形式で抽出し、ファイルとして保存します。
- `xlsx_extract_csv`: 指定シート（または全シート）を個別の CSV ファイルとして書き出します。
- `xlsx_extract_sheet`: 指定シートのみを抽出した新しい Excel ファイルを作成します。
- `xlsx_extract_images`: 指定シートを画像 (PNG/JPG) として保存します。レイアウトを横幅 1 ページに最適化（Fit-to-Page）して出力するため、Excel で描画されたフローチャートや設計図の読み取りに最適です。

#### CSV ツール
- `csv_read_cells`: 指定された行・列の範囲（矩形領域）を抽出します。
- `csv_search_values`: 文字列検索を行い、ヒットしたセルの位置と値を返却します。
- `csv_get_metadata`: 文字コード、総行数、列数などのメタ情報を取得します。
- `csv_extract_to_file`: 指定範囲のデータを別ファイルとして保存します。

#### テキスト処理ツール
- `text_read_head`: テキストファイルの先頭から指定行数を読み込みます。
- `text_read_tail`: テキストファイルの末尾から指定行数を読み込みます。
- `text_grep`: 正規表現を使用してテキストファイル内を検索します。
- `text_get_metadata`: テキストファイルのエンコーディング（自動判定）、サイズ、行数を取得します。
- `text_convert_encoding`: ファイルの文字コードを変換します (Shift-JIS -> UTF-8 等)。

#### HTML ツール
- `html_extract_markdown`: HTML からテキストを抽出し、Markdown ファイルとして保存します。

#### ユーティリティ
- `zip_files`: 指定したファイル群を ZIP アーカイブに圧縮します。
- `unzip_file`: ZIP アーカイブを解凍します。

## インストール

現在のディレクトリでパッケージとしてインストールすることで、CLI コマンドとして利用可能になります。

```bash
pip install .
```

これにより、以下のコマンドがシステムに登録されます：
- `doctools-mcp`: ドキュメント操作・検索 MCP サーバーの起動。
- `index_ctl`: 全文検索インデックスの管理・構築ツール。

## セットアップ

### 1. 環境設定
`.env` ファイルを作成し、インデックスの保存先を指定してください。

```env
# 複数のインデックスを格納するベースディレクトリ (必須)
WHOOSH_INDEX_BASE_DIR=C:\path\to\your\index_base
```

### 2. 検索インデックスの構築
CLI ツール `index_ctl` を使用してインデックスを作成・更新します。

```bash
# 基本: 存在すれば差分更新、なければ新規作成します
index_ctl build <docs_dir> <index_name>

# 説明（Description）付きで構築
index_ctl build <docs_dir> <index_name> --description "Project A Specifications"

# 対象とする拡張子をカスタマイズ（例: PythonとSQLを追加）
index_ctl build <docs_dir> <index_name> --ext .pdf .md .py .sql

# 処理の詳細（タイムアウト設定やDry-run）
index_ctl build <docs_dir> <index_name> --timeout 60 --dry-run
```

- **拡張子のカスタマイズ**: `--ext` オプションで対象の拡張子を自由に指定できます。
- **設定の永続化**: 一度 `--ext` で拡張子を指定して構築すると、その設定はインデックスのメタデータとして保存されます。次回の更新（`build`）時には、`--ext` を省略しても前回の設定が自動的に引き継がれます。
- **デフォルト拡張子**: オプションも過去の設定もない場合は、`.pdf, .pptx, .xlsx, .docx, .txt, .md` が対象になります。

## 使い方

MCP サーバーは標準入出力 (stdio) モードで起動します。

```bash
# MCP サーバーの起動
doctools-mcp
```

### インデックスのメンテナンス (`index_ctl`)

```bash
# 利用可能なインデックスの一覧表示
index_ctl list

# 特定のインデックスに登録済みファイルの一覧表示
index_ctl list <index_name>

# インデックスからのパス削除
index_ctl delete <path> <index_name>
```

各コマンドで操作対象の `index_name` を指定することで、マルチインデックス環境を管理できます。`list` コマンドを引数なしで実行すると、ベースディレクトリ配下のインデックス名とその説明の一覧が表示されます。

