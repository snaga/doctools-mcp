import os
from pathlib import Path
from typing import Any, Dict, Optional
from PIL import Image, ImageGrab
from datetime import datetime

def get_image_metadata(path: str) -> Dict[str, Any]:
    """
    指定された画像ファイルを開き、幅 (width) と高さ (height) を取得して辞書で返します。

    Args:
        path: 画像ファイルへのパス。

    Returns:
        'width', 'height', 'status' を含む辞書。
        エラーが発生した場合は、'status' と 'detail' を返します。
    """
    try:
        img_path = Path(path)
        if not img_path.exists():
            return {"status": "error", "detail": f"File not found: {path}"}

        with Image.open(img_path) as img:
            width, height = img.size
            return {
                "status": "success",
                "width": width,
                "height": height,
                "format": img.format,
                "mode": img.mode
            }
    except Exception as e:
        return {"status": "error", "detail": str(e)}

def crop_image(
    path: str, 
    left: int, 
    top: int, 
    right: int, 
    bottom: int, 
    output_path: Optional[str] = None
) -> Dict[str, Any]:
    """
    指定された矩形領域 (left, top, right, bottom) で画像を切り抜きます。

    Args:
        path: 元の画像ファイルへのパス。
        left: 切り抜き領域の左端の座標。
        top: 切り抜き領域の上端の座標。
        right: 切り抜き領域の右端の座標。
        bottom: 切り抜き領域の下端の座標。
        output_path: 切り抜き後の画像を保存するパス（任意）。指定されない場合は、元のファイル名に '_cropped' を付加したパスを生成します。

    Returns:
        'status', 'message', 'output_path' を含む辞書。
        エラーが発生した場合は、'status' と 'detail' を返します。
    """
    try:
        img_path = Path(path)
        if not img_path.exists():
            return {"status": "error", "detail": f"File not found: {path}"}

        # 座標の妥当性チェック
        if right <= left or bottom <= top:
            return {
                "status": "error", 
                "detail": f"Invalid coordinates: right ({right}) must be > left ({left}) and bottom ({bottom}) must be > top ({top})."
            }

        with Image.open(img_path) as img:
            width, height = img.size
            
            # 座標が画像サイズを超えていないかチェック
            if left < 0 or top < 0 or right > width or bottom > height:
                return {
                    "status": "error",
                    "detail": f"Coordinates out of bounds: Image size is {width}x{height}, but crop box is ({left}, {top}, {right}, {bottom})."
                }

            # 切り抜き実行
            cropped_img = img.crop((left, top, right, bottom))

            # 出力パスの決定
            if output_path is None:
                output_path = str(img_path.with_name(f"{img_path.stem}_cropped{img_path.suffix}"))
            
            abs_output_path = os.path.abspath(output_path)
            
            # 切り抜き後の画像を保存
            cropped_img.save(abs_output_path)

            return {
                "status": "success",
                "message": "Image cropped successfully.",
                "output_path": abs_output_path
            }
    except Exception as e:
        return {"status": "error", "detail": str(e)}

def save_clipboard_image(
    output_dir: Optional[str] = None, 
    filename: Optional[str] = None
) -> Dict[str, Any]:
    """
    Windows クリップボードから画像を取得し、PNG ファイルとして保存します。

    Args:
        output_dir: 保存先のディレクトリパス（任意）。未指定時はカレントディレクトリ。
        filename: 保存するファイル名（任意）。未指定時は 'clipboard_YYYYMMDD_HHMMSS.png' を自動生成。

    Returns:
        'status', 'message', 'output_path' を含む辞書。
        画像がない場合やエラー時は、'status' と 'detail' を返します。
    """
    try:
        # クリップボードから画像を取得
        # Note: ImageGrab.grabclipboard() は Windows/macOS で動作する
        img = ImageGrab.grabclipboard()

        if img is None:
            return {"status": "error", "detail": "No image found in clipboard."}

        # クリップボードの内容がファイルパス（リスト）の場合も Pillow がよしなに扱うことがあるが、
        # ここでは明示的な画像データのみを対象とする
        if not isinstance(img, Image.Image):
            return {"status": "error", "detail": "Clipboard content is not an image."}

        # 保存先ディレクトリの決定
        target_dir = Path(output_dir) if output_dir else Path.cwd()
        if not target_dir.exists():
            target_dir.mkdir(parents=True, exist_ok=True)

        # ファイル名の決定
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"clipboard_{timestamp}.png"
        elif not filename.lower().endswith(".png"):
            filename += ".png"

        output_path = target_dir / filename
        abs_output_path = str(output_path.absolute())

        # 保存実行
        img.save(abs_output_path, "PNG")

        return {
            "status": "success",
            "message": "Image saved from clipboard successfully.",
            "output_path": abs_output_path
        }
    except Exception as e:
        return {"status": "error", "detail": str(e)}
