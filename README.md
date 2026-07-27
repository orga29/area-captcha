# AreaCap4Kind Fork

このプロジェクトは、Windowsで指定した画面領域を連続キャプチャし、1つのPDFへ結合するツールです。ローカルのPython環境で利用します。

## ファイル構成

- `areacap4kind1529.py`: メインのPythonスクリプト
- `start.bat`: Windowsでダブルクリックして起動するバッチファイル
- `1529.ico`: ビルド用のアイコンファイル
- `areacap4kind1529.spec`: PyInstaller用のビルド設定ファイル
- `scripts/Sign-AreaCap.ps1`: ビルド済みEXEを署名・検証するPowerShellスクリプト

## Windowsでの起動（推奨）

事前にPython 3をインストールし、インストール時に **Add Python to PATH** を有効にしてください。初回のみ、PowerShellまたはコマンドプロンプトで依存ライブラリを入れます。

```powershell
python -m pip install pyautogui keyboard Pillow
```

その後は `start.bat` をダブルクリックして起動できます。Python環境だけを確認するには、コマンドプロンプトで次を実行します。

```bat
start.bat --check
```

## Pythonからの起動

```powershell
python areacap4kind1529.py
```

## 補足

PDF作成に成功すると、中間生成物のJPEGファイル（`screenshots/page_*.jpg`）は自動削除され、`screenshots/output.pdf` のみが残ります。

## 依存ライブラリ

以下のライブラリが必要です。

- `pyautogui`
- `keyboard`
- `Pillow` (PIL)

## コード署名（任意）

PyInstallerでEXEを作成して他者へ配布する場合だけ、Smart App Control対策としてコード署名を検討してください。手順は [docs/code-signing.md](docs/code-signing.md) を参照してください。証明書や秘密鍵はリポジトリに保存しません。
