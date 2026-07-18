# このリポジトリでの注意事項（Claude 向け）

## `.gs` ファイルのトップレベル宣言は `var` のままにする

Apps Script は、1つのプロジェクト内の複数 `.gs` ファイルを「同じグローバルオブジェクトを共有する
別々のトップレベルスクリプト」として実行する（ブラウザで複数の `<script>` タグを並べるのと同じモデル）。

- トップレベルの `var` / `function` 宣言はグローバルオブジェクトに載るため、**他ファイルから参照できる**。
- トップレベルの `let` / `const` 宣言はそのファイル内だけのスコープになり、**他ファイルからは参照できない**
  （ReferenceError になる）。

このリポジトリでは `template-setup.gs` のシート名定数（`UC_LIST_SHEET_NAME`、`TEMPLATE_SHEET_NAMES` など）や
`validation.gs` の `VALIDATION_ROW_HEADROOM` のように、複数ファイルから参照される定数がある。
**これらトップレベルの共有定数は `var` のまま**にすること。`let`/`const` に変換すると、他ファイルから
参照できなくなり本番の Apps Script 環境で壊れる。

一方、**関数内（ローカル）の `var` は `let`/`const` に変換して問題ない**。関数内変数は元々そのファイル内・
その関数内でしか使われないため、他ファイルとの共有スコープの制約を受けない。

### 検証方法

`test/support/gasSandbox.js` は Apps Script と同じ「複数ファイルを同じグローバルスコープに順次ロードする」
挙動を Node の `vm` モジュールで再現している。トップレベル宣言を誤って `let`/`const` にすると、他ファイルの
関数からその定数を参照するテストが `undefined`/`ReferenceError` で落ちるため、この方式で検出できる。
`.gs` ファイルを変更したときは、この観点でテストが機能しているかどうかも意識すること。

## 開発フロー

- テストは `pnpm test`（Vitest）。`create-spreadsheet.gs` は既に廃止され、`template-setup.gs` /
  `template-sheets.gs` / `validation.gs` / `ids.gs` / `menu.gs` / `markdown-export.gs` の6ファイルに
  分割されている。
- ブランチは `main` から作り直す運用（このリポジトリの PR は squash 的にマージされるため、過去の
  feature ブランチをそのまま使い回さない）。
