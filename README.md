# AreaCap4Kind Fork

このプロジェクトは `areacap4kind1529.py` を元にしたカスタマイズ用のフォークです。

## ファイル構成
- `areacap4kind1529.py`: メインのPythonスクリプト
- `1529.ico`: ビルド用のアイコンファイル
- `areacap4kind1529.spec`: PyInstaller用のビルド設定ファイル

## 実行方法
```bash
python areacap4kind1529.py
```

## 補足
PDF作成に成功すると、中間生成物のJPEGファイル（`screenshots/page_*.jpg`）は自動削除され、`screenshots/output.pdf` のみが残ります。

## 依存ライブラリ
以下のライブラリが必要です：
- `pyautogui`
- `keyboard`
- `Pillow` (PIL)
