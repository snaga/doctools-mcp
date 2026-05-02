# テクノロジースタック

## 言語
- Python

## コアライブラリ
- **fastmcp**
  - **選定理由**: デコレータを用いたPythonicな記述、型ヒントによるJSON Schemaの自動生成、開発効率の高さ。
- **pypdf**
  - **選定理由**: PDFの分割、マージ、テキスト抽出のための軽量で標準的なライブラリ。
- **PyMuPDF (fitz)**
  - **選定理由**: PDF および Excel (PDF経由) のページを画像 (PNG/JPG) に変換する際に、高速かつ高品質なレンダリングが可能。外部ソフトウェア (Poppler等) への依存がなく、ライブラリ単体で完結するため採用。
- **python-pptx**
  - **選定理由**: PowerPoint (.pptx) ファイルの構造解析とテキスト抽出に最適。
- **pywin32**
  - **選定理由**: Windows 環境において PowerPoint および Excel を直接操作し、レイアウト崩れのない高精度なスライド画像化、スライド結合、およびハックを用いた Excel シートの PDF エクスポートを実現するために採用。また、Windows クリップボード（画像・テキスト）の操作を低レイヤーで行うためにも利用する。COM (Component Object Model) 経由での自動化を利用する。
- **openpyxl**
  - **選定理由**: Excel (.xlsx) ファイルの読み書きにおける標準的かつ軽量なライブラリ。依存関係を最小限に抑えるため採用。
- **MarkItDown**
  - **選定理由**: Microsoft 製の Office 文書テキスト抽出ライブラリ。生成 AI での利用に適した形式で抽出可能。
- **Whoosh**
  - **選定理由**: 純 Python 製の全文検索エンジン。追加のサーバー構築が不要で、N-gram による日本語検索にも対応。
- **charset-normalizer**
  - **選定理由**: テキストファイルのエンコーディングを高精度に自動判定するために採用。CSVやテキスト操作機能の信頼性を向上させる。
- **Pillow (PIL)**
  - **選定理由**: Python での標準的な画像処理ライブラリ。PNG ファイルのメタデータ取得、切り抜き（Crop）などの操作を軽量かつ高速に行えるため採用。

## アーキテクチャパターン
- **Single Purpose Server**: ドキュメント操作と検索に特化した単一のサーバー (`doctools`) として構成。
- **Service Layering**: ビジネスロジックを `service.py` に分離し、`main.py` (FastMCP) はインターフェースに専念する。
- **Standalone Indexer**: 文書をスキャンしインデックスを構築する独立した CLI スクリプト (`index_ctl.py`)。
- **Handle-based Data Access**: 大容量データはファイルパス（ハンドル）として返し、AIが必要な場合にのみ中身を読みに行く設計。

## プラットフォーム固有の制約
- **Windows / Office 自動化**: PowerPoint のスライド画像化機能 (`pptx_extract_images`) および Excel のシート画像化機能 (`xlsx_extract_images`) は、Microsoft Office がインストールされた Windows 環境でのみ動作する。非 Windows 環境や未インストールの環境では、代替手段（例: スライド抽出のみ）に制限される。

## 実装上の重要なハック・最適化
- **Excel シート画像化の最適化 (Fit-to-Page Hack)**:
  Excel シートを PDF 経由で画像化する際、図形や表がページ跨ぎで分断されるのを防ぐため、PDF 出力直前に `PageSetup` を動的に変更し、横幅を強制的に 1 ページに収める設定 (`FitToPagesWide = 1`, `Zoom = False`) を適用する。これにより、LLM にとって読み取りやすい一貫した画像出力を実現する。
