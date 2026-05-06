/**
 * 要求仕様書スプレッドシート — 自動生成スクリプト
 *
 * 使い方:
 *   1. Google スプレッドシートを新規作成
 *   2. 拡張機能 > Apps Script を開く
 *   3. このコードを貼り付けて保存
 *   4. 関数「createRequirementsSheet」を選択して実行
 */

function createRequirementsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  // 既存の空白シートを削除（デフォルトの「シート1」）
  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length === 1) {
    // 一時シートを作ってから削除（最低1枚必要なため）
    const tmp = ss.insertSheet('_tmp');
    ss.deleteSheet(defaultSheet);
  }

  setupOverview(ss);
  setupActors(ss);
  setupBusinessReqs(ss);
  setupUseCases(ss);
  setupFunctionalReqs(ss);
  setupNonFunctionalReqs(ss);
  setupConstraints(ss);
  setupExternalIF(ss);
  setupOpenIssues(ss);
  setupGlossary(ss);
  setupChangeLog(ss);

  // 一時シート削除
  const tmp = ss.getSheetByName('_tmp');
  if (tmp) ss.deleteSheet(tmp);

  // 最初のタブをアクティブに
  ss.setActiveSheet(ss.getSheetByName('📋 概要'));

  SpreadsheetApp.getUi().alert('✅ 要求仕様書テンプレートの作成が完了しました！');
}

// ─────────────────────────────────────────────
// ユーティリティ
// ─────────────────────────────────────────────

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** ヘッダー行のスタイル設定 */
function styleHeader(sheet, row, cols) {
  const range = sheet.getRange(row, 1, 1, cols);
  range.setBackground('#1a73e8')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setVerticalAlignment('middle');
  sheet.setFrozenRows(row);
}

/** 列幅を一括設定 */
function setColWidths(sheet, widths) {
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

/** ドロップダウン検証 */
function setDropdown(sheet, row, col, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, col).setDataValidation(rule);
}

/** ステータス列に色付けする条件付き書式 */
function addStatusFormatting(sheet, col, lastRow) {
  const range = sheet.getRange(2, col, lastRow - 1, 1);
  const rules = [
    { text: '合意済',     bg: '#e6f4ea', fg: '#137333' },
    { text: '解決済',     bg: '#e6f4ea', fg: '#137333' },
    { text: '未解決',     bg: '#fce8e6', fg: '#c5221f' },
    { text: 'レビュー中', bg: '#fef7e0', fg: '#f57c00' },
    { text: '差し戻し',   bg: '#fce8e6', fg: '#c5221f' },
    { text: '保留',       bg: '#f1f3f4', fg: '#5f6368' },
    { text: '草案',       bg: '#f8f9fa', fg: '#5f6368' },
  ];
  const cfRules = rules.map(r =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(r.text)
      .setBackground(r.bg)
      .setFontColor(r.fg)
      .setRanges([range])
      .build()
  );
  sheet.setConditionalFormatRules(cfRules);
}

// ─────────────────────────────────────────────
// タブ 1: 📋 概要
// ─────────────────────────────────────────────

function setupOverview(ss) {
  const sh = getOrCreateSheet(ss, '📋 概要');
  sh.clearContents();
  sh.clearFormats();

  // タイトル
  sh.getRange('A1').setValue('要求仕様書').setFontSize(16).setFontWeight('bold');
  sh.getRange('A1').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.getRange('A1:D1').merge();

  // ドキュメント管理
  const meta = [
    ['ドキュメントID', 'REQ-XXXX',      'バージョン',      '1.0.0'],
    ['ステータス',     '草案',           '作成日',          ''],
    ['最終更新日',     '',               '作成者',          ''],
    ['承認者（顧客）', '',               '承認者（自社）',   ''],
  ];
  sh.getRange(3, 1, meta.length, 4).setValues(meta);
  sh.getRange(3, 1, meta.length, 1).setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(3, 3, meta.length, 1).setFontWeight('bold').setBackground('#e8f0fe');

  // ステータスのドロップダウン
  setDropdown(sh, 4, 2, ['草案', 'レビュー中', '承認済']);

  // セクションヘッダー
  const sections = [
    [9,  'プロジェクト概要'],
    [13, 'スコープ（IN）'],
    [17, 'スコープ（OUT）'],
    [21, '成功指標'],
  ];
  sections.forEach(([row, title]) => {
    sh.getRange(row, 1).setValue(title).setFontWeight('bold').setBackground('#e8f0fe');
    sh.getRange(row, 1, 1, 4).merge();
  });

  sh.getRange(10, 1).setValue('目的');
  sh.getRange(10, 2, 1, 3).merge().setValue('（例）受注管理の手作業ミスを撲滅するため、受注データの自動入力・照合システムを構築する。');
  sh.getRange(11, 1).setValue('現状（As-Is）');
  sh.getRange(12, 1).setValue('課題');

  // 成功指標テーブル
  const kpiHeader = ['指標', '現状値', '目標値', '測定方法'];
  sh.getRange(22, 1, 1, 4).setValues([kpiHeader]);
  styleHeader(sh, 22, 4);

  setColWidths(sh, [160, 200, 160, 200]);
  sh.setRowHeight(1, 36);
}

// ─────────────────────────────────────────────
// タブ 2: 👤 アクター
// ─────────────────────────────────────────────

function setupActors(ss) {
  const sh = getOrCreateSheet(ss, '👤 アクター');
  sh.clearContents(); sh.clearFormats();

  const headers = ['アクターID', 'アクター名', '説明・ロール', '利用頻度', '備考'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  const data = [
    ['ACT-01', '一般ユーザー', 'システムを日常業務で利用する担当者', '毎日', ''],
    ['ACT-02', '管理者',       'ユーザー管理・マスタ管理を行う担当者', '週次', ''],
    ['ACT-03', '外部システム', '連携する外部 API / システム',          'リアルタイム', ''],
  ];
  sh.getRange(2, 1, data.length, headers.length).setValues(data);

  setColWidths(sh, [100, 140, 300, 120, 200]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
}

// ─────────────────────────────────────────────
// タブ 3: 🎯 ビジネス要求
// ─────────────────────────────────────────────

function setupBusinessReqs(ss) {
  const sh = getOrCreateSheet(ss, '🎯 ビジネス要求');
  sh.clearContents(); sh.clearFormats();

  const headers = ['要求ID', 'ビジネス要求（1文）', '背景・理由', '優先度', '成功指標', '顧客コメント ✏️', 'ステータス'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  const data = [
    ['BR-01', '受注ミスを月 XX 件以下にする', '現状は手入力による転記ミスが多発している', 'Must',   '', '', '草案'],
    ['BR-02', '受注処理時間を XX% 短縮する',  '担当者の残業時間増加が課題',               'Should', '', '', '草案'],
    ['BR-03', '顧客への納期回答を即日化する',  '顧客満足度向上のため',                     'Could',  '', '', '草案'],
  ];
  sh.getRange(2, 1, data.length, headers.length).setValues(data);

  // 優先度ドロップダウン
  for (let r = 2; r <= data.length + 1; r++) {
    setDropdown(sh, r, 4, ['Must', 'Should', 'Could']);
    setDropdown(sh, r, 7, ['草案', 'レビュー中', '合意済', '保留', '廃止']);
  }

  // 顧客コメント列を薄黄色に
  sh.getRange(2, 6, 50, 1).setBackground('#fffde7');

  addStatusFormatting(sh, 7, 30);
  setColWidths(sh, [80, 280, 240, 80, 180, 200, 100]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
}

// ─────────────────────────────────────────────
// タブ 4: 📖 ユースケース
// ─────────────────────────────────────────────

function setupUseCases(ss) {
  const sh = getOrCreateSheet(ss, '📖 ユースケース');
  sh.clearContents(); sh.clearFormats();

  // 一覧テーブル
  sh.getRange(1, 1).setValue('▼ ユースケース一覧').setFontWeight('bold');
  const listHeaders = ['UCID', 'アクター', 'ユースケース名', '関連BR', 'ステータス'];
  sh.getRange(2, 1, 1, listHeaders.length).setValues([listHeaders]);
  styleHeader(sh, 2, listHeaders.length);

  const listData = [
    ['UC-01', 'ACT-01', '受注データを登録する', 'BR-01', '草案'],
    ['UC-02', 'ACT-01', '受注一覧を照会する',   'BR-01', '草案'],
    ['UC-03', 'ACT-02', 'ユーザーを管理する',   '',      '草案'],
  ];
  sh.getRange(3, 1, listData.length, listHeaders.length).setValues(listData);

  // UC 詳細テンプレート（UC-01）
  const detailStart = 8;
  sh.getRange(detailStart, 1).setValue('▼ UC-01: 受注データを登録する').setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(detailStart, 1, 1, 5).merge();

  const ucDetail = [
    ['アクター',          'ACT-01（一般ユーザー）', '', '', ''],
    ['事前条件',          'ユーザーがシステムにログイン済であること', '', '', ''],
    ['事後条件（正常）',  '受注データが保存され、受注番号が発番される', '', '', ''],
    ['事後条件（異常）',  'エラーメッセージが表示され、データは保存されない', '', '', ''],
  ];
  sh.getRange(detailStart + 1, 1, ucDetail.length, 5).setValues(ucDetail);
  sh.getRange(detailStart + 1, 1, ucDetail.length, 1).setFontWeight('bold').setBackground('#f8f9fa');

  const flowStart = detailStart + ucDetail.length + 2;
  sh.getRange(flowStart, 1).setValue('基本フロー').setFontWeight('bold');
  const flows = [
    ['1', '一般ユーザーが受注登録画面を開く'],
    ['2', 'システムが受注登録フォームを表示する'],
    ['3', '一般ユーザーが受注情報（顧客・品目・数量・希望納期）を入力する'],
    ['4', '一般ユーザーが「登録」ボタンをクリックする'],
    ['5', 'システムが入力値を検証し、問題がなければデータを保存する'],
    ['6', 'システムが受注番号を発番し、登録完了画面を表示する'],
  ];
  sh.getRange(flowStart + 1, 1, flows.length, 2).setValues(flows);

  const altStart = flowStart + flows.length + 2;
  sh.getRange(altStart, 1).setValue('代替フロー').setFontWeight('bold');
  const alts = [
    ['3a', '必須項目が未入力の場合 → システムは対象項目をハイライトし、エラーメッセージを表示。3 に戻る'],
    ['5a', 'システムエラー発生時 → エラーをログに記録し、ユーザーに「登録に失敗しました。再度お試しください」を表示'],
  ];
  sh.getRange(altStart + 1, 1, alts.length, 2).setValues(alts);

  setColWidths(sh, [160, 320, 160, 100, 100]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
}

// ─────────────────────────────────────────────
// タブ 5: ⚙️ 機能要求
// ─────────────────────────────────────────────

function setupFunctionalReqs(ss) {
  const sh = getOrCreateSheet(ss, '⚙️ 機能要求');
  sh.clearContents(); sh.clearFormats();

  const headers = ['要求ID', '機能名', '関連UC', '要求記述（条件＋主語＋動作）', '受け入れ基準①', '受け入れ基準②', '受け入れ基準③', '優先度', '顧客コメント ✏️', '合意ステータス', '備考'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  const data = [
    ['FR-01', '受注データ登録', 'UC-01',
     'ユーザーが「登録」ボタンをクリックしたとき、システムは入力フォームの全項目を検証し、エラーがない場合はデータを保存して完了画面に遷移する',
     '正常入力時、受注番号が発番され完了画面が表示される',
     '必須項目未入力時、対象項目が赤枠でハイライトされエラーメッセージが表示される',
     'ネットワークエラー時、エラーがログに記録されユーザーに通知される',
     'Must', '', '草案', ''],
    ['FR-02', '受注一覧照会', 'UC-02',
     'ユーザーが受注一覧画面を開いたとき、システムは直近 90 日分の受注データを受注日降順で表示する',
     '受注一覧に直近 90 日分のデータが表示される',
     '受注日・顧客名・ステータスでフィルタリングできる',
     'CSV ダウンロードが実行できる',
     'Must', '', '草案', ''],
    ['FR-03', 'ユーザー管理', 'UC-03',
     '管理者がユーザー管理画面で「追加」ボタンをクリックしたとき、システムは新規ユーザー登録フォームを表示する',
     '新規ユーザーが登録でき、登録後にログイン可能になる',
     '重複メールアドレスでの登録時、エラーメッセージが表示される',
     '',
     'Must', '', '草案', ''],
  ];
  sh.getRange(2, 1, data.length, headers.length).setValues(data);

  // ドロップダウン
  for (let r = 2; r <= data.length + 1; r++) {
    setDropdown(sh, r, 8,  ['Must', 'Should', 'Could']);
    setDropdown(sh, r, 10, ['草案', 'レビュー中', '合意済', '差し戻し', '廃止']);
  }

  // 顧客コメント列
  sh.getRange(2, 9, 50, 1).setBackground('#fffde7');

  addStatusFormatting(sh, 10, 30);
  setColWidths(sh, [80, 140, 80, 340, 200, 200, 200, 80, 180, 110, 140]);
  sh.setRowHeights(1, sh.getLastRow(), 48);
  sh.getRange(2, 4, data.length, 1).setWrap(true);
}

// ─────────────────────────────────────────────
// タブ 6: 🔒 非機能要求
// ─────────────────────────────────────────────

function setupNonFunctionalReqs(ss) {
  const sh = getOrCreateSheet(ss, '🔒 非機能要求');
  sh.clearContents(); sh.clearFormats();

  const headers = ['要求ID', 'カテゴリ', '項目名', '要求値（数値必須）', '測定条件', '測定方法', '顧客コメント ✏️', '合意ステータス'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  const data = [
    ['NFR-P01', '性能',    '画面応答時間',         '3 秒以内',              '同時接続 100 ユーザー時', '負荷テスト（JMeter 等）', '', '草案'],
    ['NFR-P02', '性能',    'バッチ処理時間',        '10,000 件を 30 分以内', '業務時間外バッチ実行時',   '実測',                   '', '草案'],
    ['NFR-R01', '可用性',  '稼働率',               '99.9% 以上（月次）',     '計画メンテナンス除く',     'SLA レポート',           '', '草案'],
    ['NFR-R02', '可用性',  'RTO（目標復旧時間）',  '障害発生から 4 時間以内', '',                        '障害訓練',               '', '草案'],
    ['NFR-R03', '可用性',  'RPO（目標復旧時点）',  '最大 1 時間前の状態',    '',                        'バックアップ確認',        '', '草案'],
    ['NFR-S01', 'セキュリティ', '認証方式',         'ID＋パスワード ＋ MFA',  '',                        'セキュリティレビュー',    '', '草案'],
    ['NFR-S02', 'セキュリティ', '通信暗号化',       'TLS 1.2 以上',          '',                        'SSLラボスキャン',        '', '草案'],
    ['NFR-S03', 'セキュリティ', '監査ログ保持',     '操作ログを 12 ヶ月保持', '',                        'ログ確認',               '', '草案'],
    ['NFR-M01', '保守性',  '設定変更',             'コード変更なしで変更可',  'マスタテーブル管理対象項目', '管理画面操作確認',      '', '草案'],
    ['NFR-U01', 'UX',      '対応ブラウザ',         'Chrome / Safari / Edge 最新版', '',               'クロスブラウザテスト',    '', '草案'],
    ['NFR-U02', 'UX',      'レスポンシブ対応',     '1024px 以上の横幅で崩れない', '',                  '実機・エミュレータ確認', '', '草案'],
  ];
  sh.getRange(2, 1, data.length, headers.length).setValues(data);

  for (let r = 2; r <= data.length + 1; r++) {
    setDropdown(sh, r, 2, ['性能', '可用性', 'セキュリティ', '保守性', 'UX']);
    setDropdown(sh, r, 8, ['草案', 'レビュー中', '合意済', '差し戻し', '廃止']);
  }

  sh.getRange(2, 7, 50, 1).setBackground('#fffde7');
  addStatusFormatting(sh, 8, 30);
  setColWidths(sh, [90, 110, 180, 220, 220, 180, 180, 110]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
}

// ─────────────────────────────────────────────
// タブ 7: 🚧 制約条件
// ─────────────────────────────────────────────

function setupConstraints(ss) {
  const sh = getOrCreateSheet(ss, '🚧 制約条件');
  sh.clearContents(); sh.clearFormats();

  const headers = ['制約ID', 'カテゴリ', '制約内容', '理由', '顧客コメント ✏️', 'ステータス'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  const data = [
    ['CON-T01', '技術',         '既存の XX システム（Oracle DB）と連携すること', 'インフラ刷新は本プロジェクト外のため',     '', '草案'],
    ['CON-T02', '技術',         'クラウドは AWS を使用すること',                '社内インフラポリシーによる',               '', '草案'],
    ['CON-B01', 'ビジネス',     '本番稼働は YYYY-MM-DD 以降であること',          '顧客の会計年度切り替えに合わせるため',     '', '草案'],
    ['CON-B02', 'ビジネス',     '既存データの移行は対象外とする',               'データクレンジングは別プロジェクトで対応', '', '草案'],
    ['CON-L01', '法規制',       '個人情報は国内リージョンに保存すること',        '個人情報保護法・社内セキュリティポリシー', '', '草案'],
  ];
  sh.getRange(2, 1, data.length, headers.length).setValues(data);

  for (let r = 2; r <= data.length + 1; r++) {
    setDropdown(sh, r, 2, ['技術', 'ビジネス', '法規制', '運用']);
    setDropdown(sh, r, 6, ['草案', '合意済', '廃止']);
  }

  sh.getRange(2, 5, 50, 1).setBackground('#fffde7');
  setColWidths(sh, [90, 100, 300, 260, 180, 90]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
}

// ─────────────────────────────────────────────
// タブ 8: 🔗 外部IF
// ─────────────────────────────────────────────

function setupExternalIF(ss) {
  const sh = getOrCreateSheet(ss, '🔗 外部IF');
  sh.clearContents(); sh.clearFormats();

  const headers = ['IF-ID', '連携先システム', '方向', 'プロトコル／形式', '頻度', 'データ概要', '担当部署', '備考'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  const data = [
    ['IF-01', '既存受注管理システム',  'OUT（送信）', 'REST API / JSON',     'リアルタイム', '受注データ',       '顧客 IT 部門', ''],
    ['IF-02', '会計システム',          'OUT（送信）', 'CSV ファイル連携',    '日次（深夜）',  '請求データ',       '顧客 経理部門', ''],
    ['IF-03', 'メール通知（SendGrid）', 'OUT（送信）', 'REST API / JSON',    'イベント駆動',  '通知メール',       '自社',         'API キー管理要'],
  ];
  sh.getRange(2, 1, data.length, headers.length).setValues(data);

  for (let r = 2; r <= data.length + 1; r++) {
    setDropdown(sh, r, 3, ['IN（受信）', 'OUT（送信）', '双方向']);
  }

  setColWidths(sh, [70, 200, 110, 160, 120, 160, 130, 160]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
}

// ─────────────────────────────────────────────
// タブ 9: ❓ 未解決事項
// ─────────────────────────────────────────────

function setupOpenIssues(ss) {
  const sh = getOrCreateSheet(ss, '❓ 未解決事項');
  sh.clearContents(); sh.clearFormats();

  const headers = ['Issue-ID', '内容', '影響する要求ID', '担当者', '期限', '回答・決定内容', 'ステータス'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  const data = [
    ['OI-01', '受注データの保持期間について法的要件を確認する必要がある', 'NFR-S03', '顧客 法務担当', '2026-05-20', '', '未解決'],
    ['OI-02', '既存システムの API 仕様書の提供依頼中',                   'IF-01',   '顧客 IT 部門', '2026-05-15', '', '未解決'],
    ['OI-03', 'バッチ処理の実行時刻について業務部門と調整中',             'NFR-P02', '顧客 業務担当', '2026-05-10', '深夜 2:00〜4:00 を想定', '保留'],
  ];
  sh.getRange(2, 1, data.length, headers.length).setValues(data);

  for (let r = 2; r <= data.length + 1; r++) {
    setDropdown(sh, r, 7, ['未解決', '解決済', '保留', '取り下げ']);
  }

  addStatusFormatting(sh, 7, 30);
  setColWidths(sh, [80, 320, 160, 140, 110, 240, 100]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
}

// ─────────────────────────────────────────────
// タブ 10: 📚 用語集
// ─────────────────────────────────────────────

function setupGlossary(ss) {
  const sh = getOrCreateSheet(ss, '📚 用語集');
  sh.clearContents(); sh.clearFormats();

  const headers = ['用語', '定義', '類義語・注意', '初出箇所'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  const data = [
    ['受注',         '顧客からの発注を自社が受け付けること。発注書または電話での確認を以て成立とする。', '注文、オーダー',   'BR-01'],
    ['受注番号',     'システムが発番する受注を一意に識別する番号。形式: RCV-YYYYMMDD-NNNN',           'オーダー番号',    'FR-01'],
    ['管理者',       'ユーザー管理・マスタ管理の権限を持つシステム利用者。人事発令で任命される。',     '管理ユーザー',    'UC-03'],
    ['バッチ処理',   '業務時間外に自動実行される一括データ処理。通常、前日分のデータを翌日深夜に処理する。', 'バッチ',        'NFR-P02'],
  ];
  sh.getRange(2, 1, data.length, headers.length).setValues(data);

  setColWidths(sh, [140, 380, 180, 100]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
  sh.getRange(2, 2, data.length, 1).setWrap(true);
}

// ─────────────────────────────────────────────
// タブ 11: ✅ 変更履歴
// ─────────────────────────────────────────────

function setupChangeLog(ss) {
  const sh = getOrCreateSheet(ss, '✅ 変更履歴');
  sh.clearContents(); sh.clearFormats();

  const headers = ['バージョン', '日付', '変更者', '変更内容', '影響箇所'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const data = [
    ['1.0.0', today, '', '初版作成', '全体'],
  ];
  sh.getRange(2, 1, data.length, headers.length).setValues(data);

  setColWidths(sh, [110, 120, 140, 340, 180]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
}
