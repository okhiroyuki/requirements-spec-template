# 要求仕様書テンプレート（Markdown / Google スプレッドシート）

顧客レビューと開発側の記述を両立させるための、**要求仕様書のひな形**と**スプレッドシート自動生成**のセットです。

## 同梱ファイル

| ファイル | 役割 |
|----------|------|
| [`requirements-spec-template.md`](requirements-spec-template.md) | Markdown 版の要求仕様書テンプレート。Git / Notion / 社内 Wiki などにそのまま貼りやすい構成。 |
| [`google-sheets-guide.md`](google-sheets-guide.md) | Google スプレッドシート版の**タブ一覧・列定義・運用ルール・共有の推奨**をまとめたガイド。 |
| [`create-spreadsheet.gs`](create-spreadsheet.gs) | 上記ガイドに沿ったシート（タブ・ヘッダー・ドロップダウン・条件付き書式など）を**一括で作成する** Google Apps Script。 |

## どれを使うか

- **ドキュメントをテキストで管理したい** → `requirements-spec-template.md` をコピーして編集。
- **顧客とスプレッドシート上でコメント・レビューしたい** → `google-sheets-guide.md` を読み、`create-spreadsheet.gs` でブックを生成してから運用。

Markdown 版とスプレッドシート版は章立て・要求の種類が対応しており、必要に応じて片方を正とし、もう一方へ同期する想定です。

## Google スプレッドシートを作る手順

1. 新しい [Google スプレッドシート](https://sheets.google.com/) を作成する。
2. **拡張機能** → **Apps Script** を開く。
3. エディタのデフォルトコードを削除し、`create-spreadsheet.gs` の内容を**すべて**貼り付けて保存する。
4. 関数 **`createRequirementsSheet`** を選び、**実行**する（初回は権限の承認が必要）。
5. スプレッドシートに戻ると、ガイドに記載されたタブ（📋 概要、👤 アクター、🎯 ビジネス要求 など）が生成されている。

詳しい列の意味や顧客共有の注意点は **`google-sheets-guide.md`** を参照してください。

## 記述の前提（抜粋）

機能要求は「**[条件] のとき、[主語] は [動作] する**」の形、非機能要求は**数値・測定可能な基準**で書く、といった原則は `requirements-spec-template.md` 冒頭に記載しています。
