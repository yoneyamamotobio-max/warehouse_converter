# 倉庫 inventory-data JSON ビューア

倉庫管理アプリから出力された `inventory-data-*.json` を読み込み、Excel風の表で閲覧、検索、並べ替え、列順変更、コピー、`.xlsx` 出力を行う別アプリです。既存の倉庫管理アプリ本体は変更しません。

## セットアップ

最初に依存ライブラリをインストールしてください。

```powershell
pip install -r requirements.txt
```

`requirements.txt` には以下を含めています。

- PySide6
- openpyxl
- PyInstaller

## 起動

```powershell
python inventory_json_viewer.py
```

または、Windowsでは以下のバッチファイルから起動できます。

```powershell
run_viewer.bat
```

ファイルを指定して起動することもできます。

```powershell
python inventory_json_viewer.py "C:\Users\yone\Downloads\inventory-data-2026-06-16.json"
```

## 主な機能

- `.json` ファイルのドラッグ&ドロップ読み込み
- 「ファイルを選択」ボタンからの読み込み
- 在庫明細、在庫集計、パレット一覧、出庫履歴、マップメモのタブ表示
- 検索対象列を選んだ部分一致検索、全列検索
- ヘッダークリックで並べ替え
- 列幅変更、列ヘッダーのドラッグによる列順変更
- 列幅と列順の保存
- セル編集不可
- Ctrl+C、右クリック「コピー」「ヘッダー付きコピー」
- 検索結果のみ、または全件を `.xlsx` 出力
- xlsx側でヘッダー固定、オートフィルター、列幅調整

## EXE 化

`build.bat` を実行すると、PyInstallerで単体EXEを作成します。

```powershell
build.bat
```

内部では以下を実行します。

```powershell
pip install -r requirements.txt
python -m PyInstaller --noconfirm --onefile --windowed --name InventoryJsonViewer --icon icon.ico --add-data "icon.ico;." --collect-all openpyxl inventory_json_viewer.py
```

生成物:

```text
dist\InventoryJsonViewer.exe
```

`openpyxl` はEXE内に同梱されるため、EXE利用者が別途 `pip install openpyxl` を実行する必要はありません。
