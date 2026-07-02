# 在庫データExcel変換ツール

倉庫管理アプリから出力された `inventory-data-*.json` を読み込み、Excel形式（`.xlsx`）に変換して出力するアプリです。既存の倉庫管理アプリ本体は変更しません。

## セットアップ

最初に依存ライブラリをインストールしてください。

```powershell
pip install --disable-pip-version-check -r requirements.txt
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
- 「JSONを選択」ボタンからの読み込み
- 在庫明細、在庫集計、パレット一覧、出庫履歴、マップメモをExcelシートとして出力
- 読み込んだ順番のままExcelへ出力
- 見出し行付き `.xlsx` 出力
- xlsx側でヘッダー固定、オートフィルター、列幅自動調整
- パレット番号は文字列、枚数や在庫日数は数値として出力
- 起動画面に位置確認用の真上ビューグリッドを表示
- グリッドは見取り図のみで、実際のパレットやアイテム情報は表示しません

## EXE 化

`build.bat` を実行すると、PyInstallerで単体EXEを作成します。

```powershell
build.bat
```

内部では以下を実行します。

```powershell
pip install -r requirements.txt
python -m PyInstaller --noconfirm --onefile --windowed --name InventoryExcelConverter --icon icon.ico --add-data "icon.ico;." --collect-all openpyxl inventory_json_viewer.py
```

生成物:

```text
dist\InventoryExcelConverter.exe
```

`openpyxl` はEXE内に同梱されるため、EXE利用者が別途 `pip install openpyxl` を実行する必要はありません。
