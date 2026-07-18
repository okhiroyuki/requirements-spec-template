/** 各タブのヘッダー・列幅・初期サンプル行の定義（setupXxx 系 + シードデータ）。 */

function setupOverview(ss) {
  const sh = getOrCreateSheet(ss, '📋 概要');
  resetSheetCellsForTemplate_(sh);

  sh.getRange('A1').setValue('要求仕様書').setFontSize(16).setFontWeight('bold');
  sh.getRange('A1').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.getRange('A1:D1').merge();

  const meta = [
    ['ドキュメントID', 'REQ-XXXX',      'バージョン',      '1.0'],
    ['ステータス',     '草案',           '作成日',          ''],
    ['最終更新日',     '',               '作成者',          ''],
    ['承認者（顧客）', '',               '承認者（自社）',   ''],
  ];
  sh.getRange(3, 4).setNumberFormat('@');
  sh.getRange(3, 1, meta.length, 4).setValues(meta);
  sh.getRange(3, 1, meta.length, 1).setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(3, 3, meta.length, 1).setFontWeight('bold').setBackground('#e8f0fe');

  setDropdown(sh, 4, 2, ['草案', 'レビュー中', '承認済']);

  const sections = [
    [9,  'プロジェクト概要'],
    [14, 'スコープ（IN）'],
    [18, 'スコープ（OUT）'],
    [22, '成功指標'],
  ];
  sections.forEach(([row, title]) => {
    sh.getRange(row, 1).setValue(title).setFontWeight('bold').setBackground('#e8f0fe');
    sh.getRange(row, 1, 1, 4).merge();
  });

  sh.getRange(10, 1).setValue('概要');
  sh.getRange(10, 2, 1, 3)
    .merge()
    .setValue(
      '（例）本プロジェクトは、受注の受理から在庫照会・納期回答・出荷連携までを対象とし、既存の受注・在庫基盤と連携しながら入力・照合・承認の体験を改善する。関係部門が同一の事実データを参照できる状態を目指す。'
    )
    .setWrap(true);
  sh.getRange(11, 1).setValue('目的');
  sh.getRange(11, 2, 1, 3)
    .merge()
    .setValue('（例）受注管理の手作業ミスを撲滅するため、受注データの自動入力・照合システムを構築する。')
    .setWrap(true);
  sh.getRange(12, 1).setValue('現状（As-Is）');
  sh.getRange(12, 2, 1, 3)
    .merge()
    .setValue(
      '（例）注文はメール・帳票・既存ツールで受けており、基幹への手入力やスプレッドシートでの追跡が中心である。'
    )
    .setWrap(true);
  sh.getRange(13, 1).setValue('課題');
  sh.getRange(13, 2, 1, 3)
    .merge()
    .setValue(
      '（例）入力遅延・二重入力・照合漏れにより、在庫引当や納期回答に時間がかかっている。'
    )
    .setWrap(true);

  const kpiHeader = ['指標', '現状値', '目標値', '測定方法'];
  sh.getRange(23, 1, 1, 4).setValues([kpiHeader]);
  styleHeader(sh, 23, 4, false);

  setColWidths(sh, [160, 200, 160, 200]);
  sh.setRowHeight(1, 36);
  sh.setRowHeights(10, 4, 52);
  sh.setFrozenRows(0);
}


function setupActors(ss) {
  const sh = getOrCreateSheet(ss, '👤 アクター');
  resetSheetCellsForTemplate_(sh);

  const headers = ['アクターID', 'アクター名', '説明・ロール', '利用頻度', '備考'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [100, 140, 300, 120, 200]);
  sh.setRowHeights(1, 1, 24);
}


function setupBusinessReqs(ss) {
  const sh = getOrCreateSheet(ss, '🎯 ビジネス要求');
  resetSheetCellsForTemplate_(sh);

  const headers = ['要求ID', 'ビジネス要求（1文）', '背景・理由', '優先度', '成功指標', '顧客コメント ✏️', 'ステータス'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [96, 280, 240, 80, 180, 200, 100]);
  sh.setRowHeights(1, 1, 24);
}


function setupBusinessUseCases(ss) {
  let sh = getOrCreateSheet(ss, BUC_SHEET_NAME);
  resetSheetCellsForTemplate_(sh);

  let headers = ['BUCID', '業務名', '業務の概要', '関連BR', '参考：ビジネス要求'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [88, 200, 320, 100, 280]);
  sh.setRowHeights(1, 1, 24);
}


function setupBucDetail(ss) {
  let sh = getOrCreateSheet(ss, BUC_DETAIL_SHEET_NAME);
  resetSheetCellsForTemplate_(sh, VALIDATION_ROW_HEADROOM, 10);
  writeBucDetailBlockAtRow_(sh, 1, 'BUC-001', '受注登録・検証業務', false, [
    ['1', '顧客が注文書を送付する', ''],
    ['2', '一般ユーザーが注文内容をシステムに入力する', 'UC-001'],
    ['3', 'システムがマスタと照合し、不備があれば警告を出す', 'UC-001'],
  ]);
  setColWidths(sh, [64, 520, 112]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
}

/** BUC詳細の ▼ 見出しと 3 列手順表を rowStart から書き込む。 */
function writeBucDetailBlockAtRow_(sh, rowStart, bucIdToken, bucName, skeletonOnly, stepRows) {
  skeletonOnly = !!skeletonOnly;
  stepRows = stepRows || [];

  let heading = '▼ ' + bucIdToken + ': ' + bucName;
  sh.getRange(rowStart, 1).setValue(heading).setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(rowStart, 1, 1, 3).merge();

  let hdrRow = rowStart + 1;
  let labels = [['手順', '行動内容', '関連UC']];
  sh.getRange(hdrRow, 1, 1, 3).setValues(labels);
  sh.getRange(hdrRow, 1, 1, 3).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');

  if (skeletonOnly) {
    // データ行は書かない（空の「1 |  | 」だけが Markdown に出るのを防ぐ。手順はユーザーが追加する）
    return;
  }
  if (stepRows.length > 0) {
    let dStart = hdrRow + 1;
    sh.getRange(dStart, 1, stepRows.length, 3).setValues(stepRows);
    sh.getRange(dStart, 2, stepRows.length, 1).setWrap(true);
  }
}


function setupUseCaseList(ss) {
  const sh = getOrCreateSheet(ss, UC_LIST_SHEET_NAME);
  resetSheetCellsForTemplate_(sh, VALIDATION_ROW_HEADROOM, 10);

  const listHeaders = ['UCID', 'アクター名', 'ユースケース名', '概要', 'ステータス'];
  sh.getRange(1, 1, 1, listHeaders.length).setValues([listHeaders]);
  styleHeader(sh, 1, listHeaders.length);

  setColWidths(sh, [160, 280, 240, 300, 120]);
  sh.setRowHeights(1, 1, 24);
}


function setupUseCaseDetail(ss) {
  const sh = getOrCreateSheet(ss, UC_DETAIL_SHEET_NAME);
  resetSheetCellsForTemplate_(sh, VALIDATION_ROW_HEADROOM, 10);

  writeUcDetailBlockAtRow_(sh, 1, 'UC-001', '受注データを登録する', '一般ユーザー');

  setColWidths(sh, [160, 320, 160, 100, 100]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
}

/** UC 詳細ブロック（▼ 見出し〜代替フロー）を rowStart から書き込む。 */
function writeUcDetailBlockAtRow_(sh, rowStart, ucIdToken, ucName, ucActorLabel, skeletonOnly) {
  skeletonOnly = !!skeletonOnly;

  let heading = '▼ ' + ucIdToken + ': ' + ucName;
  sh.getRange(rowStart, 1).setValue(heading).setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(rowStart, 1, 1, 5).merge();

  let ucDetail;
  let flows;
  let alts;

  if (skeletonOnly) {
    ucDetail = [
      ['アクター', ucActorLabel || '', '', '', ''],
      ['事前条件', '', '', '', ''],
      ['事後条件（正常）', '', '', '', ''],
      ['事後条件（異常）', '', '', '', ''],
    ];
  } else {
    ucDetail = [
      ['アクター', ucActorLabel, '', '', ''],
      ['事前条件', 'ユーザーがシステムにログイン済であること', '', '', ''],
      ['事後条件（正常）', '受注データが保存され、受注番号が発番される', '', '', ''],
      ['事後条件（異常）', 'エラーメッセージが表示され、データは保存されない', '', '', ''],
    ];
    flows = [
      ['1', '一般ユーザーが受注登録画面を開く'],
      ['2', 'システムが受注登録フォームを表示する'],
      ['3', '一般ユーザーが受注情報（顧客・品目・数量・希望納期）を入力する'],
      ['4', '一般ユーザーが「登録」ボタンをクリックする'],
      ['5', 'システムが入力値を検証し、問題がなければデータを保存する'],
      ['6', 'システムが受注番号を発番し、登録完了画面を表示する'],
    ];
    alts = [
      ['3a', '必須項目が未入力の場合 → システムは対象項目をハイライトし、エラーメッセージを表示。3 に戻る'],
      ['5a', 'システムエラー発生時 → エラーをログに記録し、ユーザーに「登録に失敗しました。再度お試しください」を表示'],
    ];
  }

  let metaStart = rowStart + 1;
  sh.getRange(metaStart, 1, ucDetail.length, 5).setValues(ucDetail);
  sh.getRange(metaStart, 1, ucDetail.length, 1).setFontWeight('bold').setBackground('#f8f9fa');

  let flowStart = rowStart + ucDetail.length + 2;
  sh.getRange(flowStart, 1).setValue('基本フロー').setFontWeight('bold');

  if (skeletonOnly) {
    let altStartSk = flowStart + 2;
    sh.getRange(altStartSk, 1).setValue('代替フロー').setFontWeight('bold');
  } else {
    let flowDataStart = flowStart + 1;
    sh.getRange(flowDataStart, 1, flows.length, 2).setValues(flows);

    let altStart = flowStart + flows.length + 2;
    sh.getRange(altStart, 1).setValue('代替フロー').setFontWeight('bold');
    let altDataStart = altStart + 1;
    sh.getRange(altDataStart, 1, alts.length, 2).setValues(alts);
  }
}


function setupFunctionalReqs(ss) {
  const sh = getOrCreateSheet(ss, '⚙️ 機能要求');
  resetSheetCellsForTemplate_(sh);

  const headers = ['要求ID', '機能名', '関連UC', '要求記述（条件＋主語＋動作）', '受け入れ基準①', '受け入れ基準②', '受け入れ基準③', '優先度', '顧客コメント ✏️', '合意ステータス', '備考'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [96, 140, 88, 340, 200, 200, 200, 80, 180, 110, 140]);
  sh.setRowHeights(1, 1, 48);
}


function setupNonFunctionalReqs(ss) {
  const sh = getOrCreateSheet(ss, '🔒 非機能要求');
  resetSheetCellsForTemplate_(sh);

  const headers = ['要求ID', 'カテゴリ', '項目名', '要求値（数値必須）', '測定条件', '測定方法', '顧客コメント ✏️', '合意ステータス'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [90, 110, 180, 220, 220, 180, 180, 110]);
  sh.setRowHeights(1, 1, 24);
}


function setupConstraints(ss) {
  const sh = getOrCreateSheet(ss, '🚧 制約条件');
  resetSheetCellsForTemplate_(sh);

  const headers = ['制約ID', 'カテゴリ', '制約内容', '理由', '顧客コメント ✏️', 'ステータス'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [90, 100, 300, 260, 180, 90]);
  sh.setRowHeights(1, 1, 24);
}


function setupExternalIF(ss) {
  const sh = getOrCreateSheet(ss, '🔗 外部IF');
  resetSheetCellsForTemplate_(sh);

  const headers = ['IF-ID', '連携先システム', '方向', 'プロトコル／形式', '頻度', 'データ概要', '担当部署', '備考'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [70, 200, 110, 160, 120, 160, 130, 160]);
  sh.setRowHeights(1, 1, 24);
}


function setupAssumptions(ss) {
  let sh = getOrCreateSheet(ss, '📌 前提条件');
  resetSheetCellsForTemplate_(sh);

  let headers = ['前提ID', '前提条件', 'リスク（崩れた場合の影響）'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [100, 380, 280]);
  sh.setRowHeights(1, 1, 24);
}


function setupOpenIssues(ss) {
  const sh = getOrCreateSheet(ss, '❓ 未解決事項');
  resetSheetCellsForTemplate_(sh);

  const headers = ['Issue-ID', '内容', '影響する要求ID', '担当者', '期限', '回答・決定内容', 'ステータス'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [96, 320, 160, 140, 110, 240, 100]);
  sh.setRowHeights(1, 1, 24);
}


function setupGlossary(ss) {
  const sh = getOrCreateSheet(ss, '📚 用語集');
  resetSheetCellsForTemplate_(sh);

  const headers = ['用語', '定義', '類義語・注意', '初出箇所'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [140, 380, 180, 100]);
  sh.setRowHeights(1, 1, 24);
}


function setupChangeLog(ss) {
  const sh = getOrCreateSheet(ss, '✅ 変更履歴');
  resetSheetCellsForTemplate_(sh);

  const headers = ['バージョン', '日付', '変更者', '変更内容', '影響箇所'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  sh.getRange(2, 1, 50, 1).setNumberFormat('@');
  setColWidths(sh, [110, 120, 140, 340, 180]);
  sh.setRowHeights(1, 1, 24);
}


function seedTemplateSampleRows_(ss) {
  let sh;
  let data;
  let n;

  sh = ss.getSheetByName('👤 アクター');
  if (sh) {
    data = [
      ['ACT-001', '一般ユーザー', 'システムを日常業務で利用する担当者', '毎日', ''],
      ['ACT-002', '管理者', 'ユーザー管理・マスタ管理を行う担当者', '週次', ''],
      ['ACT-003', '既存受注管理システム', '本プロジェクトから連携するレガシー受注・在庫基盤など', 'リアルタイム', ''],
      ['ACT-004', '会計システム', '請求データ連携などの経理・会計向け連携システム', '日次（深夜）', ''],
      ['ACT-005', 'APIサーバー', 'SendGrid 等の外部 HTTP API／通知エンドポイントを束ねて表すときの代表名', 'イベント駆動', ''],
      ['ACT-006', '顧客', '業務コンテキスト上の発注・問い合わせ側の主体', '', ''],
    ];
    sh.getRange(2, 1, data.length, 5).setValues(data);
    sh.setRowHeights(1, sh.getLastRow(), 24);
  }

  sh = ss.getSheetByName('🎯 ビジネス要求');
  if (sh) {
    data = [
      ['BR-001', '受注入力時にマスタデータと整合性を自動照合し、入力不備をその場で検知・修正できる状態にすること', '現状は手入力による転記ミスが多発している', 'Must', '', '', '草案'],
      ['BR-002', '現在手動で行っている承認フローを自動化し、担当者の介在なしに後続工程へデータを連携可能にすること', '担当者の残業時間増加が課題', 'Should', '', '', '草案'],
      ['BR-003', '在庫情報と配送スケジュールをリアルタイムに参照し、問い合わせに対して即座に正確な納期を算出・回答できること', '顧客満足度向上のため', 'Could', '', '', '草案'],
    ];
    sh.getRange(2, 1, data.length, 7).setValues(data);
    sh.getRange(2, 6, 50, 1).setBackground('#fffde7');
    sh.setRowHeights(1, sh.getLastRow(), 24);
  }

  sh = ss.getSheetByName(BUC_SHEET_NAME);
  if (sh) {
    data = [
      ['BUC-001', '受注登録・検証業務', '顧客からの注文を受け、内容を精査して受理する', 'BR-001'],
      ['BUC-002', '受注承認・出荷連携業務', '受理した注文を承認し、出荷工程へデータを送る', 'BR-002'],
      ['BUC-003', '納期回答業務', '在庫と配送状況を確認し、顧客へ納期を伝える', 'BR-003'],
    ];
    sh.getRange(2, 1, data.length, 4).setValues(data);
    let bucFormulas = [];
    let br;
    for (br = 0; br < data.length; br++) {
      bucFormulas.push([bucBrMirrorFormula_(2 + br)]);
    }
    sh.getRange(2, 5, data.length, 1).setFormulas(bucFormulas);
    sh.setRowHeights(1, sh.getLastRow(), 24);
    sh.getRange(2, 3, data.length, 1).setWrap(true);
    sh.getRange(2, 5, data.length, 1).setWrap(true);
  }

  sh = ss.getSheetByName(UC_LIST_SHEET_NAME);
  if (sh) {
    data = [
      ['UC-001', '一般ユーザー', '受注データを登録する', '注文内容を入力し、マスタ照合・検証のうえ受注として保存する', '草案'],
      ['UC-002', '一般ユーザー', '受注一覧を照会する', '条件を指定して受注情報を検索し、一覧で確認する', '草案'],
      ['UC-003', '管理者', 'ユーザーを管理する', 'アカウントの追加・権限変更・無効化など利用者を維持管理する', '草案'],
    ];
    sh.getRange(2, 1, data.length, 5).setValues(data);
    sh.getRange(2, 4, data.length, 1).setWrap(true);
    sh.setRowHeights(1, sh.getLastRow(), 24);
  }

  sh = ss.getSheetByName('⚙️ 機能要求');
  if (sh) {
    data = [
      ['FR-001', '受注データ登録', 'UC-001',
        'ユーザーが「登録」ボタンをクリックしたとき、システムは入力フォームの全項目を検証し、エラーがない場合はデータを保存して完了画面に遷移する',
        '正常入力時、受注番号が発番され完了画面が表示される',
        '必須項目未入力時、対象項目が赤枠でハイライトされエラーメッセージが表示される',
        'ネットワークエラー時、エラーがログに記録されユーザーに通知される',
        'Must', '', '草案', ''],
      ['FR-002', '受注一覧照会', 'UC-002',
        'ユーザーが受注一覧画面を開いたとき、システムは直近 90 日分の受注データを受注日降順で表示する',
        '受注一覧に直近 90 日分のデータが表示される',
        '受注日・顧客名・ステータスでフィルタリングできる',
        'CSV ダウンロードが実行できる',
        'Must', '', '草案', ''],
      ['FR-003', 'ユーザー管理', 'UC-003',
        '管理者がユーザー管理画面で「追加」ボタンをクリックしたとき、システムは新規ユーザー登録フォームを表示する',
        '新規ユーザーが登録でき、登録後にログイン可能になる',
        '重複メールアドレスでの登録時、エラーメッセージが表示される',
        '',
        'Must', '', '草案', ''],
    ];
    n = data.length;
    sh.getRange(2, 1, n, 11).setValues(data);
    sh.getRange(2, 9, 50, 1).setBackground('#fffde7');
    sh.setRowHeights(1, sh.getLastRow(), 48);
    sh.getRange(2, 4, n, 1).setWrap(true);
  }

  sh = ss.getSheetByName('🔒 非機能要求');
  if (sh) {
    data = [
      ['NFR-001', '性能', '画面応答時間', '3 秒以内', '同時接続 100 ユーザー時', '負荷テスト（JMeter 等）', '', '草案'],
      ['NFR-002', '性能', 'バッチ処理時間', '10,000 件を 30 分以内', '業務時間外バッチ実行時', '実測', '', '草案'],
      ['NFR-003', '可用性', '稼働率', '99.9% 以上（月次）', '計画メンテナンス除く', 'SLA レポート', '', '草案'],
      ['NFR-004', '可用性', 'RTO（目標復旧時間）', '障害発生から 4 時間以内', '', '障害訓練', '', '草案'],
      ['NFR-005', '可用性', 'RPO（目標復旧時点）', '最大 1 時間前の状態', '', 'バックアップ確認', '', '草案'],
      ['NFR-006', 'セキュリティ', '認証方式', 'ID＋パスワード ＋ MFA', '', 'セキュリティレビュー', '', '草案'],
      ['NFR-007', 'セキュリティ', '通信暗号化', 'TLS 1.2 以上', '', 'SSLラボスキャン', '', '草案'],
      ['NFR-008', 'セキュリティ', '監査ログ保持', '操作ログを 12 ヶ月保持', '', 'ログ確認', '', '草案'],
      ['NFR-009', '保守性', '設定変更', 'コード変更なしで変更可', 'マスタテーブル管理対象項目', '管理画面操作確認', '', '草案'],
      ['NFR-010', 'UX', '対応ブラウザ', 'Chrome / Safari / Edge 最新版', '', 'クロスブラウザテスト', '', '草案'],
      ['NFR-011', 'UX', 'レスポンシブ対応', '1024px 以上の横幅で崩れない', '', '実機・エミュレータ確認', '', '草案'],
    ];
    sh.getRange(2, 1, data.length, 8).setValues(data);
    sh.getRange(2, 7, 50, 1).setBackground('#fffde7');
    sh.setRowHeights(1, sh.getLastRow(), 24);
  }

  sh = ss.getSheetByName('🚧 制約条件');
  if (sh) {
    data = [
      ['CON-001', '技術', '既存の XX システム（Oracle DB）と連携すること', 'インフラ刷新は本プロジェクト外のため', '', '草案'],
      ['CON-002', '技術', 'クラウドは AWS を使用すること', '社内インフラポリシーによる', '', '草案'],
      ['CON-003', 'ビジネス', '本番稼働は YYYY-MM-DD 以降であること', '顧客の会計年度切り替えに合わせるため', '', '草案'],
      ['CON-004', 'ビジネス', '既存データの移行は対象外とする', 'データクレンジングは別プロジェクトで対応', '', '草案'],
      ['CON-005', '法規制', '個人情報は国内リージョンに保存すること', '個人情報保護法・社内セキュリティポリシー', '', '草案'],
    ];
    sh.getRange(2, 1, data.length, 6).setValues(data);
    sh.getRange(2, 5, 50, 1).setBackground('#fffde7');
    sh.setRowHeights(1, sh.getLastRow(), 24);
  }

  sh = ss.getSheetByName('🔗 外部IF');
  if (sh) {
    data = [
      ['IF-001', '既存受注管理システム', 'OUT（送信）', 'REST API / JSON', 'リアルタイム', '受注データ', '顧客 IT 部門', ''],
      ['IF-002', '会計システム', 'OUT（送信）', 'CSV ファイル連携', '日次（深夜）', '請求データ', '顧客 経理部門', ''],
      ['IF-003', 'APIサーバー', 'OUT（送信）', 'REST API / JSON（例: SendGrid）', 'イベント駆動', '通知メール', '自社', 'API キー管理要'],
    ];
    sh.getRange(2, 1, data.length, 8).setValues(data);
    sh.setRowHeights(1, sh.getLastRow(), 24);
  }

  sh = ss.getSheetByName('📌 前提条件');
  if (sh) {
    data = [['ASM-001', '（例）既存受注 DB のスキーマ変更は本プロジェクトのスコープ外である', '連携 IF の仕様見直し・スケジュール遅延の可能性']];
    sh.getRange(2, 1, 1, 3).setValues(data);
    sh.getRange(2, 2, 1, 2).setWrap(true);
    sh.setRowHeights(1, sh.getLastRow(), 24);
  }

  sh = ss.getSheetByName('❓ 未解決事項');
  if (sh) {
    data = [
      ['OI-001', '受注データの保持期間について法的要件を確認する必要がある', 'NFR-008', '顧客 法務担当', '2026-05-20', '', '未解決'],
      ['OI-002', '既存システムの API 仕様書の提供依頼中', 'IF-001', '顧客 IT 部門', '2026-05-15', '', '未解決'],
      ['OI-003', 'バッチ処理の実行時刻について業務部門と調整中', 'NFR-002', '顧客 業務担当', '2026-05-10', '深夜 2:00〜4:00 を想定', '保留'],
    ];
    sh.getRange(2, 1, data.length, 7).setValues(data);
    sh.setRowHeights(1, sh.getLastRow(), 24);
  }

  sh = ss.getSheetByName('📚 用語集');
  if (sh) {
    data = [
      ['受注', '顧客からの発注を自社が受け付けること。発注書または電話での確認を以て成立とする。', '注文、オーダー', 'BR-001'],
      ['受注番号', 'システムが発番する受注を一意に識別する番号。形式: RCV-YYYYMMDD-NNNN', 'オーダー番号', 'FR-001'],
      ['管理者', 'ユーザー管理・マスタ管理の権限を持つシステム利用者。人事発令で任命される。', '管理ユーザー', 'UC-003'],
      ['バッチ処理', '業務時間外に自動実行される一括データ処理。通常、前日分のデータを翌日深夜に処理する。', 'バッチ', 'NFR-002'],
    ];
    n = data.length;
    sh.getRange(2, 1, n, 4).setValues(data);
    sh.getRange(2, 2, n, 1).setWrap(true);
    sh.setRowHeights(1, sh.getLastRow(), 24);
  }

  sh = ss.getSheetByName('✅ 変更履歴');
  if (sh) {
    let tz = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';
    let createdDate = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    sh.getRange(2, 1, 1, 5).setValues([['1.0', createdDate, '', '初版作成', '全体']]);
    sh.setRowHeights(1, sh.getLastRow(), 24);
  }
}
