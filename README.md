# ExcelReadMCP

ExcelReadMCP は、Cursor や GitHub Copilot などの Model Context Protocol 対応クライアントから Excel ファイルを読み取るための専用サーバーです。読み込み・検索に特化しており、書き込みや整形系の操作は提供しません。

## 提供ツール

| ツール名 | 説明 |
| --- | --- |
| `excel_read_info` | ワークブックのメタ情報（シート数、シート名、ファイルサイズなど）を返します。 |
| `excel_read_range` | 指定シート（または先頭シート）の内容をレコード配列として返します。 |
| `excel_read_all_sheets` | 全シートを読み込み、シートごとのデータと処理状況を返します。 |
| `excel_quick_overview` | ファイル概要と各シートのサンプル行を返します。 |
| `excel_search` | ワークブック全体（または指定シート）から文字列を検索します。 |

> **重要:** すべてのツールで `file_path` には **絶対パス** を指定してください。相対パスを渡すとエラーになります。

## 共通セットアップ

```powershell
cd <path-to-ExcelReadMCP>
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

> **実行例:** フォルダを `C:\Projects\ExcelReadMCP` に展開した場合は、`cd C:\Projects\ExcelReadMCP` → `python -m venv .venv` → `.\.venv\Scripts\Activate.ps1` → `pip install -r requirements.txt` の順で PowerShell から実行します。

> **補足:** MCP ライブラリとして公式の `mcp` パッケージ（現在の安定版は 1.18.0）を利用しているため、`requirements.txt` ではそのバージョン以上を指定しています。

## Cursor でのセットアップ

1. `mcp_config.json` を開き、`<path-to-ExcelReadMCP>` を実際の絶対パスに置き換えます。
2. その内容を Cursor が参照する MCP 設定ファイル（例: `%USERPROFILE%\.cursor\mcp.json`）へ追記します。
3. Cursor を再起動し、Settings > Features > MCP に `excel-read-mcp` が表示されることを確認します。
4. Composer（`Ctrl` + `I`）で「`C:\path\to\workbook.xlsx` のシート一覧を取得して」などと指示し、ツールが利用できることを確認します。

## GitHub Copilot でのセットアップ

1. `mcp_config.json` を開き、`<path-to-ExcelReadMCP>` を実際の絶対パスに置き換えます。
2. Windows の場合は `%APPDATA%\GitHub Copilot\mcp.json` を編集し、`excel-read-tools` の設定を追記します（ファイルが無い場合は新規作成してください）。
3. VS Code を再起動し、Copilot Chat のツール一覧に `excel-read-tools` が表示されることを確認します。
4. Copilot Chat に「`C:\path\to\workbook.xlsx` の内容を確認して」などと指示し、ツールの応答をテストします。

## サーバーの起動

```powershell
cd <path-to-ExcelReadMCP>
start_mcp_server.bat
```

仮想環境が存在する場合は `.venv` の Python、存在しない場合はシステムの `python` が使用されます。
