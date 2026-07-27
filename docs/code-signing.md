# Windows コード署名

Smart App Control で実行できるようにするには、PyInstaller が生成する `dist/areacap.exe` をリリースごとにコード署名します。Python ソース (`.py`) を署名対象にはしません。

## 事前準備

- Windows SDK の **Signing Tools for Desktop Apps** をインストールし、`signtool.exe` を利用可能にする。
- Microsoft Trusted Root Program に含まれる CA が発行した、**RSA** のコード署名証明書を用意する。
- PFX とそのパスワード、秘密鍵を Git に保存しない。
- 証明書発行元が指定するタイムスタンプ URL を確認する。既定値は DigiCert の URL。

Smart App Control は ECC 署名をサポートしないため、RSA 証明書を使います。

## ビルド

```powershell
pyinstaller .\areacap4kind1529.spec --noconfirm
```

この操作の後、`dist\areacap.exe` が存在することを確認します。

## PFX を使う場合

パスワードを画面に表示せず入力して署名します。スクリプトは一時的に CurrentUser の証明書ストアへ PFX を読み込み、署名後に削除します。

```powershell
$password = Read-Host 'PFX password' -AsSecureString
.\scripts\Sign-AreaCap.ps1 -PfxPath 'C:\secure\areacap-codesign.pfx' -PfxPassword $password
```

## Windows 証明書ストアを使う場合

既に CurrentUser の Personal ストアへ秘密鍵付き証明書を入れている場合は、SHA-1 拇印を指定します。

```powershell
.\scripts\Sign-AreaCap.ps1 -CertificateThumbprint '0123456789ABCDEF0123456789ABCDEF01234567'
```

CA 指定のタイムスタンプ URL を使うには `-TimestampUrl` を追加します。

```powershell
.\scripts\Sign-AreaCap.ps1 -CertificateThumbprint '<thumbprint>' -TimestampUrl 'https://example.invalid/timestamp'
```

最後に実行される `signtool verify /pa /v` が成功することを確認してから、署名済みの EXE を配布してください。
