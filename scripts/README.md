# Apps Script（ソース）

Google スプレッドシート向けのスクリプトを置いています。

| ファイル | 役割 |
|----------|------|
| [create-spreadsheet.gs](create-spreadsheet.gs) | シート生成、`createRequirementsSheet`、行追加パネル、Markdown 書き出し、ID 採番・同期、メニュー。**テンプレや出力ルールを変えるときはこのファイル**。 |

## 使い方

1. ブラウザで Google の **Apps Script** エディタを開き、[create-spreadsheet.gs](create-spreadsheet.gs) の内容を**すべてコピー**して貼り付けて保存する（手順の詳細はリポジトリ直下の [README.md](../README.md) を参照）。
2. [appsscript.json](../appsscript.json) は Advanced サービス（Google Sheets API）用マニフェストです。README の「`appsscript.json` の入れ方」に従って同じプロジェクトに含めます。
