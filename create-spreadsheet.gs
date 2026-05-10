/**
 * 要求仕様書スプレッドシート — テンプレート自動生成
 *
 * 使い方:
 *   1. Google スプレッドシートを新規作成
 *   2. 拡張機能 > Apps Script を開く
 *   3. このコードを貼り付けて保存
 *   4. 関数「createRequirementsSheet」を選択して実行（実行のたびに全シートを初期サンプルで上書きする）
 *   作成完了ダイアログにメニュー利用の注意（再読み込み）が表示される。反映後はメニュー「要求仕様書」から各機能が使える。
 *
 *   一覧・ステータス等の入力規則は SpreadsheetApp（従来の矢印ドロップダウン）。旧版は Sheets API で「テーブル」化していたが、型付き列と入力規則が両立しないため未使用。
 *   appsscript.json の Sheets サービスは任意（コンテナバインドのままでも問題なし）。
 *
 *   シートの列・項目・Markdown 書き出しなどテンプレートまわりを変えたいときは、このファイルを編集する。
 */

var UC_LIST_SHEET_NAME = '📖 UC一覧';
var UC_DETAIL_SHEET_NAME = '📖 UC詳細';

var ID_SHEET_NAME = '🔢 ID管理';

/** 同期時に必ず行を用意するキー（テンプレに現れないキーは 0） */
var ID_COUNTER_KEYS = [
  'BR', 'FR', 'UC', 'IF', 'OI', 'ACT', 'ASM',
  'NFR',
  'CON'
];

/**
 * 全シートをクリアし、初期サンプルを再展開する。実行のたびに同じ（確認ダイアログなし）。
 * 流れ: 旧 tbl_* 削除（データ消しを避けるため seed より前）→ setup → seed → 入力規則。
 */
function createRequirementsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  deleteExistingReqSpecTables_(ss.getId());
  SpreadsheetApp.flush();

  // 既存の空白シートを削除（デフォルトの「シート1」）
  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length === 1) {
    const tmp = ss.insertSheet('_tmp');
    ss.deleteSheet(defaultSheet);
  }

  setupOverview(ss);
  setupActors(ss);
  setupBusinessReqs(ss);
  setupUseCaseList(ss);
  setupUseCaseDetail(ss);
  setupFunctionalReqs(ss);
  setupNonFunctionalReqs(ss);
  setupConstraints(ss);
  setupExternalIF(ss);
  setupAssumptions(ss);
  setupOpenIssues(ss);
  setupGlossary(ss);
  setupChangeLog(ss);
  setupIdSheetHeaderOnly_(ss);

  const tmp = ss.getSheetByName('_tmp');
  if (tmp) ss.deleteSheet(tmp);

  ss.setActiveSheet(ss.getSheetByName('📋 概要'));

  SpreadsheetApp.flush();

  seedTemplateSampleRows_(ss);
  SpreadsheetApp.flush();

  applyRequirementDropdowns_(ss);
  SpreadsheetApp.flush();

  syncIdCountersFromBookCore(ss);

  applyStatusFormattingAfterTables_(ss);
  SpreadsheetApp.flush();

  reorderReqSpecSheetTabs_(ss);
  ss.setActiveSheet(ss.getSheetByName('📋 概要'));

  try {
    SpreadsheetApp.getUi().alert(
      '✅ 要求仕様書テンプレートの作成が完了しました！\n\n' +
        'メニューバーに「要求仕様書」が出るまで、スプレッドシートを再読み込みするか、タブを一度閉じて開き直してください。'
    );
  } catch (ignore) {
    Logger.log('createRequirementsSheet: 完了ダイアログを表示できませんでした。');
  }
}

// ─────────────────────────────────────────────
// ユーティリティ
// ─────────────────────────────────────────────

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * テンプレのタブ順を固定する（ガイド・行追加パネルと同じ論理順）。
 * insertSheet の既定位置の差を吸収する。
 */
function reorderReqSpecSheetTabs_(ss) {
  var names = [
    '📋 概要',
    '👤 アクター',
    '🎯 ビジネス要求',
    UC_LIST_SHEET_NAME,
    UC_DETAIL_SHEET_NAME,
    '⚙️ 機能要求',
    '🔒 非機能要求',
    '🚧 制約条件',
    '🔗 外部IF',
    '📌 前提条件',
    '❓ 未解決事項',
    '📚 用語集',
    '✅ 変更履歴',
    ID_SHEET_NAME,
  ];
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(i + 1);
    }
  }
}

/**
 * getUi() が使えないコンテキスト（サイドバーからのサーバー呼び出し等）でも落ちない通知。
 * まずダイアログ、無理ならトースト、それも無理なら Logger。
 */
function notifyUser_(message, title) {
  title = title || '要求仕様書';
  try {
    SpreadsheetApp.getUi().alert(title ? title + '\n\n' + message : message);
    return;
  } catch (ignore) {}
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(message, title, 12);
  } catch (ignore2) {
    Logger.log('[' + title + '] ' + message);
  }
}

function showSidebarSafe_(htmlOutput) {
  try {
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    return true;
  } catch (e) {
    notifyUser_(
      'スプレッドシートを開いた状態でメニューから実行してください。',
      'サイドバーを開けません'
    );
    return false;
  }
}

function showModalDialogSafe_(htmlOutput, dialogTitle) {
  try {
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogTitle);
    return true;
  } catch (e) {
    notifyUser_(
      'ダイアログを表示できませんでした。ブックを開いた状態でメニューから実行してください。\n' + String(e.message || e),
      dialogTitle || 'エラー'
    );
    return false;
  }
}

/**
 * 軽い完了フィードバック（ID 再同期など）。行追加メニューでは呼ばない。
 */
function toastDone_(message, title) {
  title = title || '完了';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(message, title, 5);
  } catch (ignore) {
    Logger.log('[' + title + '] ' + message);
  }
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

/** ドロップダウン検証（一覧から選択する入力規則） */
function setDropdown(sheet, row, col, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, col).setDataValidation(rule);
}

/**
 * 👤 アクター の A 列→B 列（アクター名）マップ。
 * @return {!Object<string, string>}
 */
function readActorMap_(ss) {
  var sh = ss.getSheetByName('👤 アクター');
  if (!sh) return {};
  var lr = sh.getLastRow();
  if (lr < 2) return {};
  // getRange(r,c,numRows,numCols): 第3引数は終端行ではなく行数（2行目〜lr行目 → lr-1 行）
  var vals = sh.getRange(2, 1, lr - 1, 2).getValues();
  var map = {};
  for (var i = 0; i < vals.length; i++) {
    var id = String(vals[i][0]).trim();
    if (!id) continue;
    map[id] = String(vals[i][1] != null ? vals[i][1] : '').trim();
  }
  return map;
}

/**
 * セル値の先頭の ACT-nnn を取り出す。
 */
function extractActorIdFromCell_(text) {
  var m = String(text || '').trim().match(/^(ACT-\d+)/);
  return m ? m[1] : '';
}

/** アクター名（B 列・重複は先勝ち）→ アクターID。Markdown 出力の名前解決用。 */
function readActorNameToIdMap_(ss) {
  var sh = ss.getSheetByName('👤 アクター');
  if (!sh) return {};
  var lr = sh.getLastRow();
  if (lr < 2) return {};
  var vals = sh.getRange(2, 1, lr - 1, 2).getValues();
  var map = {};
  var i;
  for (i = 0; i < vals.length; i++) {
    var id = String(vals[i][0]).trim();
    var name = String(vals[i][1] != null ? vals[i][1] : '').trim();
    if (!id || !name) continue;
    if (!(name in map)) map[name] = id;
  }
  return map;
}

/**
 * Markdown 出力用にアクター欄を「ACT-xxx（表示名）」へ解決する。
 * ブックに入っているのが ID でもアクター名でもよい。同名が複数ある場合は先頭行の ID を使う。
 */
function resolveActorLabelForMarkdown_(cellValue, actorMap, actorNameToId) {
  actorMap = actorMap || {};
  actorNameToId = actorNameToId || {};
  var raw = String(cellValue != null ? cellValue : '').trim();
  if (!raw) return '';
  var id = extractActorIdFromCell_(raw);
  if (id) {
    var nm = actorMap[id];
    if (nm) return id + '（' + nm + '）';
    return raw;
  }
  id = actorNameToId[raw];
  if (id) return id + '（' + raw + '）';
  return raw;
}

/**
 * 👤 アクター で実データがある A 列の最終行まで（行の決定用）。B 列入力規則は同じ行範囲を使う。
 */
function getActorIdValidationRange_(ss) {
  var actorSh = ss.getSheetByName('👤 アクター');
  if (!actorSh) return null;
  var lr = actorSh.getLastRow();
  if (lr < 2) return null;
  var colA = actorSh.getRange(2, 1, lr - 1, 1).getValues();
  var lastData = 1;
  var i;
  for (i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim() !== '') lastData = i + 2;
  }
  if (lastData < 2) return null;
  var numRows = lastData - 2 + 1;
  return actorSh.getRange(2, 1, numRows, 1);
}

/** 👤 アクターの「アクター名」列（B）のうち、ID 行と同じ範囲を返す（UC 一覧プルダウン用）。 */
function getActorNameValidationRange_(ss) {
  var idR = getActorIdValidationRange_(ss);
  if (!idR) return null;
  var sh = idR.getSheet();
  var startRow = idR.getRow();
  var numRows = idR.getLastRow() - startRow + 1;
  return sh.getRange(startRow, 2, numRows, 1);
}

/**
 * シートの A 列で最終データ行までの ID セル範囲（requireValueInRange の一覧元）。
 */
function getFirstColumnIdRange_(sheet, maxScanRow) {
  if (!sheet) return null;
  maxScanRow = maxScanRow || 500;
  var lr = Math.min(sheet.getLastRow(), maxScanRow);
  if (lr < 2) return null;
  var numScan = lr - 1;
  var colA = sheet.getRange(2, 1, numScan, 1).getValues();
  var lastData = 1;
  var i;
  for (i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim() !== '') lastData = i + 2;
  }
  if (lastData < 2) return null;
  var numRows = lastData - 2 + 1;
  return sheet.getRange(2, 1, numRows, 1);
}

/** 🎯 ビジネス要求 の BR-ID（A 列）一覧範囲 */
function getBrIdListRange_(ss) {
  return getFirstColumnIdRange_(ss.getSheetByName('🎯 ビジネス要求'));
}

/** 📖 UC一覧 の UC-ID（A 列）一覧範囲 */
function getUcIdListRange_(ss) {
  return getFirstColumnIdRange_(ss.getSheetByName(UC_LIST_SHEET_NAME));
}

/**
 * 📖 UC一覧 の「関連BR」列（4）に、🎯 ビジネス要求の BR-ID を選ぶ入力規則を付与する。
 */
function applyUcListRelatedBrValidation_(ss) {
  var listSh = ss.getSheetByName(UC_LIST_SHEET_NAME);
  if (!listSh) return;
  var vr = getBrIdListRange_(ss);
  var rule = vr
    ? SpreadsheetApp.newDataValidation()
        .requireValueInRange(vr, true)
        .setAllowInvalid(true)
        .build()
    : null;

  var lr = Math.min(listSh.getLastRow(), 500);
  var r;
  try {
    for (r = 2; r <= lr; r++) {
      var ucCell = String(listSh.getRange(r, 1).getValue()).trim();
      if (!/^UC-\d+$/.test(ucCell)) continue;
      var cell = listSh.getRange(r, 4);
      if (rule) cell.setDataValidation(rule);
      else cell.clearDataValidation();
    }
  } catch (e) {
    var msg = String(e.message || e);
    if (msg.indexOf('型付き') !== -1 || /typed column/i.test(msg)) {
      Logger.log('applyUcListRelatedBrValidation_: skip typed table column — ' + msg);
      return;
    }
    throw e;
  }
}

/**
 * ⚙️ 機能要求 の「関連UC」列（3）に、📖 UC一覧の UC-ID を選ぶ入力規則を付与する。
 */
function applyFrRelatedUcValidation_(ss) {
  var frSh = ss.getSheetByName('⚙️ 機能要求');
  if (!frSh) return;
  var vr = getUcIdListRange_(ss);
  var rule = vr
    ? SpreadsheetApp.newDataValidation()
        .requireValueInRange(vr, true)
        .setAllowInvalid(true)
        .build()
    : null;

  var lr = Math.min(frSh.getLastRow(), 500);
  var r;
  try {
    for (r = 2; r <= lr; r++) {
      var frCell = String(frSh.getRange(r, 1).getValue()).trim();
      if (!/^FR-\d+$/.test(frCell)) continue;
      var cell = frSh.getRange(r, 3);
      if (rule) cell.setDataValidation(rule);
      else cell.clearDataValidation();
    }
  } catch (e) {
    var msg2 = String(e.message || e);
    if (msg2.indexOf('型付き') !== -1 || /typed column/i.test(msg2)) {
      Logger.log('applyFrRelatedUcValidation_: skip typed table column — ' + msg2);
      return;
    }
    throw e;
  }
}

/** アクター名・関連 BR／UC・外部 IF 連携先など、別シート参照の入力規則をまとめて付与 */
function applyAllReferenceValidations_(ss) {
  applyUcListActorValidation_(ss);
  applyUcListRelatedBrValidation_(ss);
  applyFrRelatedUcValidation_(ss);
  applyExternalIfPartnerValidation_(ss);
}

/**
 * 📖 UC一覧 の「アクター名」列に、👤 アクター の B 列から選ぶ入力規則を付与する。
 */
function applyUcListActorValidation_(ss) {
  var listSh = ss.getSheetByName(UC_LIST_SHEET_NAME);
  var vr = getActorNameValidationRange_(ss);
  if (!listSh || !vr) return;

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(vr, true)
    .setAllowInvalid(true)
    .build();

  var lr = listSh.getLastRow();
  var r;
  try {
    for (r = 2; r <= lr; r++) {
      var ucCell = String(listSh.getRange(r, 1).getValue()).trim();
      if (/^UC-\d+$/.test(ucCell)) {
        listSh.getRange(r, 2).setDataValidation(rule);
      }
    }
  } catch (e) {
    var msg = String(e.message || e);
    if (msg.indexOf('型付き') !== -1 || /typed column/i.test(msg)) {
      Logger.log('applyUcListActorValidation_: skip typed table column — ' + msg);
      return;
    }
    throw e;
  }
}

/**
 * 🔗 外部IF のデータ行について「連携先システム」列へ、👤 アクター B 列のアクター名を選べる入力規則を付与する。
 * （システム名は 👤 にアクターとして足しておけば UC・IF で同じ名前に揃えられる）
 */
function applyExternalIfPartnerValidation_(ss) {
  var ifSh = ss.getSheetByName('🔗 外部IF');
  var vr = getActorNameValidationRange_(ss);
  if (!ifSh || !vr) return;

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(vr, true)
    .setAllowInvalid(true)
    .build();

  var lr = Math.min(ifSh.getLastRow(), 500);
  var r;
  try {
    for (r = 2; r <= lr; r++) {
      var idCell = String(ifSh.getRange(r, 1).getValue()).trim();
      if (/^IF-\d+$/.test(idCell)) {
        ifSh.getRange(r, 2).setDataValidation(rule);
      }
    }
  } catch (e) {
    var msg = String(e.message || e);
    if (msg.indexOf('型付き') !== -1 || /typed column/i.test(msg)) {
      Logger.log('applyExternalIfPartnerValidation_: skip typed table column — ' + msg);
      return;
    }
    throw e;
  }
}

/** メニュー用：優先度・ステータスに加え、BR／UC／アクターなど一覧参照の入力規則をすべて再適用 */
function menuRefreshAllInputValidations() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    applyRequirementDropdowns_(ss);
    toastDone_(
      '🎯 BR・📖 UC・👤 アクターなどを参照する入力規則を含め、ブック全体のドロップダウンを再適用しました。',
      '入力規則'
    );
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

/** 後方互換：旧メニュー関数から同一処理を呼ぶ */
function menuRefreshUcListActorValidation() {
  menuRefreshAllInputValidations();
}

/** 全シートに SpreadsheetApp の入力規則を付与（tbl 削除は createRequirementsSheet の先頭で済ませる）。 */
function applyRequirementDropdowns_(ss) {
  applyLegacyDropdowns_(ss);
  applyAllReferenceValidations_(ss);
}

/**
 * 再実行時の同名テーブル衝突を避けるため、当スクリプトが付与したテーブル名を列挙して削除する。
 */
function deleteExistingReqSpecTables_(spreadsheetId) {
  var known = {
    tbl_actor: true,
    tbl_uc_list: true,
    tbl_br: true,
    tbl_fr: true,
    tbl_nfr: true,
    tbl_con: true,
    tbl_if: true,
    tbl_oi: true,
    tbl_asm: true,
    tbl_glossary: true,
    tbl_changelog: true,
    tbl_id_mgmt: true,
  };
  try {
    var res = Sheets.Spreadsheets.get(spreadsheetId);
    var dels = [];
    (res.sheets || []).forEach(function (sheet) {
      var tables = sheet.tables;
      if (!tables || !tables.length) return;
      tables.forEach(function (t) {
        if (t.tableId && t.name && known[t.name]) {
          dels.push({ deleteTable: { tableId: t.tableId } });
        }
      });
    });
    if (dels.length) {
      Sheets.Spreadsheets.batchUpdate({ requests: dels }, spreadsheetId);
    }
  } catch (e) {
    Logger.log('deleteExistingReqSpecTables_: ' + e);
  }
}

/**
 * 旧実装: Sheets API の addTable でチップ型ドロップダウンを付けていた。
 * 型付き列と SpreadsheetApp の入力規則（例: UC 一覧のアクター名列）が両立しないため未使用。
 */
function applyAllDataTables_(ss) {
  return false;
}

/** 各データシートのドロップダウン入力規則を一括付与 */
function applyLegacyDropdowns_(ss) {
  var lrCap = 500;

  var shBR = ss.getSheetByName('🎯 ビジネス要求');
  if (shBR) {
    var lrBR = Math.min(shBR.getLastRow(), lrCap);
    for (var rBR = 2; rBR <= lrBR; rBR++) {
      setDropdown(shBR, rBR, 4, ['Must', 'Should', 'Could']);
      setDropdown(shBR, rBR, 7, ['草案', 'レビュー中', '合意済', '保留', '廃止']);
    }
  }

  var shFR = ss.getSheetByName('⚙️ 機能要求');
  if (shFR) {
    var lrFR = Math.min(shFR.getLastRow(), lrCap);
    for (var rFR = 2; rFR <= lrFR; rFR++) {
      setDropdown(shFR, rFR, 8, ['Must', 'Should', 'Could']);
      setDropdown(shFR, rFR, 10, ['草案', 'レビュー中', '合意済', '差し戻し', '廃止']);
    }
  }

  var shNFR = ss.getSheetByName('🔒 非機能要求');
  if (shNFR) {
    var lrNFR = Math.min(shNFR.getLastRow(), lrCap);
    for (var rNFR = 2; rNFR <= lrNFR; rNFR++) {
      setDropdown(shNFR, rNFR, 2, ['性能', '可用性', 'セキュリティ', '保守性', 'UX']);
      setDropdown(shNFR, rNFR, 8, ['草案', 'レビュー中', '合意済', '差し戻し', '廃止']);
    }
  }

  var shCON = ss.getSheetByName('🚧 制約条件');
  if (shCON) {
    var lrCON = Math.min(shCON.getLastRow(), lrCap);
    for (var rCON = 2; rCON <= lrCON; rCON++) {
      setDropdown(shCON, rCON, 2, ['技術', 'ビジネス', '法規制', '運用']);
      setDropdown(shCON, rCON, 6, ['草案', '合意済', '廃止']);
    }
  }

  var shIF = ss.getSheetByName('🔗 外部IF');
  if (shIF) {
    var lrIF = Math.min(shIF.getLastRow(), lrCap);
    for (var rIF = 2; rIF <= lrIF; rIF++) {
      setDropdown(shIF, rIF, 3, ['IN（受信）', 'OUT（送信）', '双方向']);
    }
  }

  var shOI = ss.getSheetByName('❓ 未解決事項');
  if (shOI) {
    var lrOI = Math.min(shOI.getLastRow(), lrCap);
    for (var rOI = 2; rOI <= lrOI; rOI++) {
      setDropdown(shOI, rOI, 7, ['未解決', '解決済', '保留', '取り下げ']);
    }
  }

  applyUcListDropdownsLegacy_(ss);
}

/** 📖 UC一覧のデータ行にステータス列の入力規則 */
function applyUcListDropdownsLegacy_(ss) {
  var sh = ss.getSheetByName(UC_LIST_SHEET_NAME);
  if (!sh) return;
  var lr = sh.getLastRow();
  if (lr < 2) return;
  var opts = ['草案', 'レビュー中', '合意済', '保留', '廃止'];
  for (var r = 2; r <= lr; r++) {
    var text = String(sh.getRange(r, 1).getValue()).trim();
    if (/^UC-\d+$/.test(text)) {
      setDropdown(sh, r, 5, opts);
    }
  }
}

/** ステータス列の文字色のみを条件付き書式で付与する（セル背景は付けない） */
function addStatusFormatting(sheet, col, lastRow) {
  const range = sheet.getRange(2, col, lastRow - 1, 1);
  const rules = [
    { text: '合意済',     fg: '#137333' },
    { text: '解決済',     fg: '#137333' },
    { text: '未解決',     fg: '#c5221f' },
    { text: 'レビュー中', fg: '#f57c00' },
    { text: '差し戻し',   fg: '#c5221f' },
    { text: '保留',       fg: '#5f6368' },
    { text: '草案',       fg: '#5f6368' },
  ];
  const cfRules = rules.map(r =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(r.text)
      .setFontColor(r.fg)
      .setBold(true)
      .setRanges([range])
      .build()
  );
  sheet.setConditionalFormatRules(cfRules);
}

/** ドロップダウン適用後に条件付き書式（ステータス列の前景色）を付与 */
function applyStatusFormattingAfterTables_(ss) {
  var sh;
  sh = ss.getSheetByName('🎯 ビジネス要求');
  if (sh) addStatusFormatting(sh, 7, 30);
  sh = ss.getSheetByName('⚙️ 機能要求');
  if (sh) addStatusFormatting(sh, 10, 30);
  sh = ss.getSheetByName('🔒 非機能要求');
  if (sh) addStatusFormatting(sh, 8, 30);
  sh = ss.getSheetByName('❓ 未解決事項');
  if (sh) addStatusFormatting(sh, 7, 30);
  sh = ss.getSheetByName(UC_LIST_SHEET_NAME);
  if (sh) addStatusFormatting(sh, 5, 500);
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

  setColWidths(sh, [100, 140, 300, 120, 200]);
  sh.setRowHeights(1, 1, 24);
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

  setColWidths(sh, [96, 280, 240, 80, 180, 200, 100]);
  sh.setRowHeights(1, 1, 24);
}

// ─────────────────────────────────────────────
// タブ 4a: 📖 UC一覧
// ─────────────────────────────────────────────

function setupUseCaseList(ss) {
  const sh = getOrCreateSheet(ss, UC_LIST_SHEET_NAME);
  sh.clearContents(); sh.clearFormats();

  const listHeaders = ['UCID', 'アクター名', 'ユースケース名', '関連BR', 'ステータス'];
  sh.getRange(1, 1, 1, listHeaders.length).setValues([listHeaders]);
  styleHeader(sh, 1, listHeaders.length);

  setColWidths(sh, [160, 320, 160, 100, 100]);
  sh.setRowHeights(1, 1, 24);
}

// ─────────────────────────────────────────────
// タブ 4b: 📖 UC詳細（一覧シートとは別）
// ─────────────────────────────────────────────

function setupUseCaseDetail(ss) {
  const sh = getOrCreateSheet(ss, UC_DETAIL_SHEET_NAME);
  sh.clearContents(); sh.clearFormats();

  writeUcDetailBlockAtRow_(sh, 1, 'UC-001', '受注データを登録する', '一般ユーザー');

  setColWidths(sh, [160, 320, 160, 100, 100]);
  sh.setRowHeights(1, sh.getLastRow(), 24);
}

/**
 * UC 詳細ブロックを rowStart から書き込む（見出し〜代替フローまで）。
 * ucActorLabel は 📖 UC一覧 の「アクター名」列と同様に、👤 アクター の B 列と一致する名前を入れる。
 * skeletonOnly が true のときは項目見出しのみ（メニューからの追加）。本文・フロー番号は入れない。
 */
function writeUcDetailBlockAtRow_(sh, rowStart, ucIdToken, ucName, ucActorLabel, skeletonOnly) {
  skeletonOnly = !!skeletonOnly;

  var heading = '▼ ' + ucIdToken + ': ' + ucName;
  sh.getRange(rowStart, 1).setValue(heading).setFontWeight('bold').setBackground('#e8f0fe');
  // getRange(r,c,行数,列数): 第3引数は「終端行」ではなく行数。ここを rowStart にすると rowStart 行ぶん縦マージになり見出しが下端に寄る
  sh.getRange(rowStart, 1, 1, 5).merge();

  var ucDetail;
  var flows;
  var alts;

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

  var metaStart = rowStart + 1;
  sh.getRange(metaStart, 1, ucDetail.length, 5).setValues(ucDetail);
  sh.getRange(metaStart, 1, ucDetail.length, 1).setFontWeight('bold').setBackground('#f8f9fa');

  var flowStart = rowStart + ucDetail.length + 2;
  sh.getRange(flowStart, 1).setValue('基本フロー').setFontWeight('bold');

  if (skeletonOnly) {
    // 見出しのみ（行番号・3a 等は入れない）。1 行空けて代替フロー見出し
    var altStartSk = flowStart + 2;
    sh.getRange(altStartSk, 1).setValue('代替フロー').setFontWeight('bold');
  } else {
    var flowDataStart = flowStart + 1;
    sh.getRange(flowDataStart, 1, flows.length, 2).setValues(flows);

    var altStart = flowStart + flows.length + 2;
    sh.getRange(altStart, 1).setValue('代替フロー').setFontWeight('bold');
    var altDataStart = altStart + 1;
    sh.getRange(altDataStart, 1, alts.length, 2).setValues(alts);
  }
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

  setColWidths(sh, [96, 140, 88, 340, 200, 200, 200, 80, 180, 110, 140]);
  sh.setRowHeights(1, 1, 48);
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

  setColWidths(sh, [90, 110, 180, 220, 220, 180, 180, 110]);
  sh.setRowHeights(1, 1, 24);
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

  setColWidths(sh, [90, 100, 300, 260, 180, 90]);
  sh.setRowHeights(1, 1, 24);
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

  setColWidths(sh, [70, 200, 110, 160, 120, 160, 130, 160]);
  sh.setRowHeights(1, 1, 24);
}

// ─────────────────────────────────────────────
// タブ 9: 📌 前提条件（ASM）
// ─────────────────────────────────────────────

function setupAssumptions(ss) {
  var sh = getOrCreateSheet(ss, '📌 前提条件');
  sh.clearContents();
  sh.clearFormats();

  var headers = ['前提ID', '前提条件', 'リスク（崩れた場合の影響）'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [100, 380, 280]);
  sh.setRowHeights(1, 1, 24);
}

// ─────────────────────────────────────────────
// タブ 10: ❓ 未解決事項
// ─────────────────────────────────────────────

function setupOpenIssues(ss) {
  const sh = getOrCreateSheet(ss, '❓ 未解決事項');
  sh.clearContents(); sh.clearFormats();

  const headers = ['Issue-ID', '内容', '影響する要求ID', '担当者', '期限', '回答・決定内容', 'ステータス'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [96, 320, 160, 140, 110, 240, 100]);
  sh.setRowHeights(1, 1, 24);
}

// ─────────────────────────────────────────────
// タブ 11: 📚 用語集
// ─────────────────────────────────────────────

function setupGlossary(ss) {
  const sh = getOrCreateSheet(ss, '📚 用語集');
  sh.clearContents(); sh.clearFormats();

  const headers = ['用語', '定義', '類義語・注意', '初出箇所'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [140, 380, 180, 100]);
  sh.setRowHeights(1, 1, 24);
}

// ─────────────────────────────────────────────
// タブ 12: ✅ 変更履歴
// ─────────────────────────────────────────────

function setupChangeLog(ss) {
  const sh = getOrCreateSheet(ss, '✅ 変更履歴');
  sh.clearContents(); sh.clearFormats();

  const headers = ['バージョン', '日付', '変更者', '変更内容', '影響箇所'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, 1, headers.length);

  setColWidths(sh, [110, 120, 140, 340, 180]);
  sh.setRowHeights(1, 1, 24);
}

// ─────────────────────────────────────────────
// setup・入力規則のあとにサンプル行を入れる
// ─────────────────────────────────────────────

function seedTemplateSampleRows_(ss) {
  var sh;
  var data;
  var n;

  sh = ss.getSheetByName('👤 アクター');
  if (sh) {
    data = [
      ['ACT-001', '一般ユーザー', 'システムを日常業務で利用する担当者', '毎日', ''],
      ['ACT-002', '管理者', 'ユーザー管理・マスタ管理を行う担当者', '週次', ''],
      ['ACT-003', '既存受注管理システム', '本プロジェクトから連携するレガシー受注・在庫基盤など', 'リアルタイム', '🔗 IF-001 想定'],
      ['ACT-004', '会計システム', '請求データ連携などの経理・会計向け連携システム', '日次（深夜）', '🔗 IF-002 想定'],
      ['ACT-005', 'APIサーバー', 'SendGrid 等の外部 HTTP API／通知エンドポイントを束ねて表すときの代表名', 'イベント駆動', '🔗 IF-003 想定'],
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

  sh = ss.getSheetByName(UC_LIST_SHEET_NAME);
  if (sh) {
    data = [
      ['UC-001', '一般ユーザー', '受注データを登録する', 'BR-001', '草案'],
      ['UC-002', '一般ユーザー', '受注一覧を照会する', 'BR-001', '草案'],
      ['UC-003', '管理者', 'ユーザーを管理する', '', '草案'],
    ];
    sh.getRange(2, 1, data.length, 5).setValues(data);
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
    var tz = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';
    var createdDate = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    sh.getRange(2, 1, 1, 5).setValues([['1.0.0', createdDate, '', '初版作成', '全体']]);
    sh.setRowHeights(1, sh.getLastRow(), 24);
  }
}

/** 🔢 ID管理：ヘッダのみ（seed 後 syncIdCountersFromBookCore で中身を埋める） */
function setupIdSheetHeaderOnly_(ss) {
  var sh = getOrCreateSheet(ss, ID_SHEET_NAME);
  sh.clearContents();
  sh.clearFormats();
  sh.getRange(1, 1, 1, 3).setValues([['キー', '最終発番（数値）', '説明']]);
  styleHeader(sh, 1, 3);
  setColWidths(sh, [100, 130, 320]);
  try {
    sh.hideSheet();
  } catch (e) {}
}

// ─────────────────────────────────────────────
// ID 採番（🔢 ID管理）・行追加パネル（menuAdd*）
// ─────────────────────────────────────────────

/** 従来どおりヘッダ＋カウンタ同期（単体実行用）。createRequirementsSheet は setupIdSheetHeaderOnly_＋seed 後に sync */
function setupIdManagement(ss) {
  setupIdSheetHeaderOnly_(ss);
  syncIdCountersFromBookCore(ss);
}

/**
 * メニュー「🔢 IDカウンタをブックから再同期」用。
 * 手編集後やメニュー実行前にブック内の ID を走査して 🔢 ID管理 を更新する。
 */
function syncIdCountersFromBook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  syncIdCountersFromBookCore(ss);
  if (ss.getSheetByName(ID_SHEET_NAME)) {
    toastDone_('🔢 ID管理をブック内の ID に合わせました', '再同期');
  }
}

function syncIdCountersFromBookCore(ss) {
  var sh = ss.getSheetByName(ID_SHEET_NAME);
  if (!sh) {
    notifyUser_('🔢 ID管理 シートがありません。先に createRequirementsSheet を実行してください。', 'ID 管理');
    return;
  }
  var maxMap = scanMaxIdsFromBook(ss);
  var rows = [['キー', '最終発番（数値）', '説明']];
  for (var i = 0; i < ID_COUNTER_KEYS.length; i++) {
    var k = ID_COUNTER_KEYS[i];
    var n = maxMap[k];
    if (n == null || n === '') n = 0;
    rows.push([k, n, '']);
  }
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  styleHeader(sh, 1, 3);
}

/**
 * ブック内の各シートから ID を走査し、キーごとの最大連番を返す。
 */
function scanMaxIdsFromBook(ss) {
  var maxMap = {};

  function bump(key, num) {
    var n = parseInt(num, 10);
    if (isNaN(n)) return;
    if (maxMap[key] == null || n > maxMap[key]) maxMap[key] = n;
  }

  function scanColumn(sheetName, col, visitor) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    var lr = sheet.getLastRow();
    if (lr < 2) return;
    var vals = sheet.getRange(2, col, lr - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      visitor(String(vals[i][0]).trim());
    }
  }

  scanColumn('🎯 ビジネス要求', 1, function (text) {
    var m = text.match(/^BR-(\d+)$/);
    if (m) bump('BR', m[1]);
  });

  scanColumn('⚙️ 機能要求', 1, function (text) {
    var m = text.match(/^FR-(\d+)$/);
    if (m) bump('FR', m[1]);
  });

  scanColumn('🔗 外部IF', 1, function (text) {
    var m = text.match(/^IF-(\d+)$/);
    if (m) bump('IF', m[1]);
  });

  scanColumn('❓ 未解決事項', 1, function (text) {
    var m = text.match(/^OI-(\d+)$/);
    if (m) bump('OI', m[1]);
  });

  scanColumn('👤 アクター', 1, function (text) {
    var m = text.match(/^ACT-(\d+)$/);
    if (m) bump('ACT', m[1]);
  });

  scanColumn('📌 前提条件', 1, function (text) {
    var m = text.match(/^ASM-(\d+)$/);
    if (m) bump('ASM', m[1]);
  });

  scanColumn('🔒 非機能要求', 1, function (text) {
    var m1 = text.match(/^NFR-(\d+)$/);
    if (m1) bump('NFR', m1[1]);
    // 旧形式 NFR-P01 等（移行済みブック用）
    var m2 = text.match(/^NFR-([A-Z])(\d+)$/);
    if (m2) bump('NFR', m2[2]);
  });

  scanColumn('🚧 制約条件', 1, function (text) {
    var m1 = text.match(/^CON-(\d+)$/);
    if (m1) bump('CON', m1[1]);
    // 旧形式 CON-T01 等（移行済みブック用）
    var m2 = text.match(/^CON-([A-Z])(\d+)$/);
    if (m2) bump('CON', m2[2]);
  });

  // UC 一覧（A 列）および UC 詳細（見出し ▼ UC-nnn）
  ;[UC_LIST_SHEET_NAME, UC_DETAIL_SHEET_NAME].forEach(function (name) {
    var ucSh = ss.getSheetByName(name);
    if (!ucSh) return;
    var lr2 = ucSh.getLastRow();
    for (var r = 1; r <= lr2; r++) {
      var text = String(ucSh.getRange(r, 1).getValue()).trim();
      var m1 = text.match(/^UC-(\d+)$/);
      if (m1) bump('UC', m1[1]);
      var m2 = text.match(/▼\s*UC-(\d+)/);
      if (m2) bump('UC', m2[1]);
    }
  });

  return maxMap;
}

/**
 * ロック付きで連番を +1 し、表示用 ID 文字列を返す。🔢 ID管理 を更新する。
 */
function issueNextId(ss, counterKey) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = ss.getSheetByName(ID_SHEET_NAME);
    if (!sh) throw new Error('ID管理シートがありません');

    var data = sh.getDataRange().getValues();
    var rowIndex = -1;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === counterKey) {
        rowIndex = r + 1;
        break;
      }
    }
    if (rowIndex < 0) throw new Error('未定義のカウンタキー: ' + counterKey);

    var last = Number(data[rowIndex - 1][1]);
    if (isNaN(last)) last = 0;
    var next = last + 1;
    sh.getRange(rowIndex, 2).setValue(next);

    return formatRequirementId(counterKey, next);
  } finally {
    lock.releaseLock();
  }
}

function formatRequirementId(counterKey, num) {
  var n = Number(num);
  if (isNaN(n) || n < 1) throw new Error('不正な連番: ' + num);
  var s = String(n);
  var pad = s.length < 3 ? ('000' + s).slice(-3) : s;
  var simple = { BR: 'BR-', FR: 'FR-', UC: 'UC-', IF: 'IF-', OI: 'OI-', ACT: 'ACT-', ASM: 'ASM-', CON: 'CON-', NFR: 'NFR-' };
  var p = simple[counterKey];
  if (!p) throw new Error('不正なキー: ' + counterKey);
  return p + pad;
}

// ───────── 行追加サイドパネル（メニューと同じ menuAdd* を呼ぶ） ─────────

function showAddRowPanel() {
  var html = HtmlService.createHtmlOutput(getAddRowPanelHtml_()).setTitle('行を追加');
  showSidebarSafe_(html);
}

function getAddRowPanelHtml_() {
  var esc = function (t) {
    return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  };
  var fail =
    'function fail(e){alert(e&&e.message?e.message:String(e));}';
  var btn = function (fn, label) {
    return (
      '<button type="button" onclick="google.script.run.withFailureHandler(fail).' +
      fn +
      '()">' +
      esc(label) +
      '</button>'
    );
  };
  return (
    '<!DOCTYPE html><html><head><base target="_top"><meta charset="UTF-8">' +
      '<style>' +
      'body{font-family:Roboto,Segoe UI,Arial,sans-serif;padding:14px;margin:0;background:#fff;color:#202124;}' +
      'h1{font-size:15px;font-weight:600;margin:0 0 8px;}' +
      '.desc{font-size:12px;color:#5f6368;line-height:1.45;margin:0 0 14px;}' +
      'button{display:block;width:100%;box-sizing:border-box;margin:0 0 8px;padding:11px 14px;' +
      'font-size:13px;text-align:left;border:1px solid #dadce0;border-radius:10px;background:#f8f9fa;' +
      'cursor:pointer;color:#174ea6;font-weight:500;}' +
      'button:hover{background:#e8f0fe;border-color:#1a73e8;}' +
      'button:active{background:#d2e3fc;}' +
      '</style></head><body>' +
      '<h1>行を追加</h1>' +
      '<p class="desc">表の末尾に 1 行追加し、ID を自動採番します（ボタンはシートタブと同じ並び）。パネルを開いたまま連続で押せます。</p>' +
      btn('menuAddACT', 'ACT · 👤 アクター') +
      btn('menuAddBR', 'BR · 🎯 ビジネス要求') +
      btn('menuAddUC', 'UC · 📖 UC一覧') +
      btn('menuAddFR', 'FR · ⚙️ 機能要求') +
      btn('menuAddNFR', 'NFR · 🔒 非機能要求') +
      btn('menuAddCON', 'CON · 🚧 制約条件') +
      btn('menuAddIF', 'IF · 🔗 外部 IF') +
      btn('menuAddASM', 'ASM · 📌 前提条件') +
      btn('menuAddOI', 'OI · ❓ 未解決事項') +
      '<script>' +
      fail +
      '</script>' +
      '</body></html>'
  );
}

// ───────── メニュー（onOpen） ─────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('要求仕様書')
    .addItem('🔢 IDカウンタをブックから再同期', 'syncIdCountersFromBook')
    .addItem('🔗 入力規則をすべて更新（BR／UC／アクター連携）', 'menuRefreshAllInputValidations')
    .addItem('🧩 行を追加パネル（サイドバー）', 'showAddRowPanel')
    .addItem('📖 選択行の UC 詳細を追加／更新', 'menuAppendUcDetailFromListRow')
    .addSeparator()
    .addItem('📝 Markdown を作成（表示・コピー）', 'exportRequirementsToMarkdown')
    .addToUi();
}

function menuAddBR() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var id = issueNextId(ss, 'BR');
    var sh = ss.getSheetByName('🎯 ビジネス要求');
    if (!sh) {
      notifyUser_('シート「🎯 ビジネス要求」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '', 'Must', '', '', '草案']);
    var row = sh.getLastRow();
    setDropdown(sh, row, 4, ['Must', 'Should', 'Could']);
    setDropdown(sh, row, 7, ['草案', 'レビュー中', '合意済', '保留', '廃止']);
    sh.getRange(row, 6).setBackground('#fffde7');
    applyAllReferenceValidations_(ss);
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddFR() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var id = issueNextId(ss, 'FR');
    var sh = ss.getSheetByName('⚙️ 機能要求');
    if (!sh) {
      notifyUser_('シート「⚙️ 機能要求」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '', '', '', '', '', 'Must', '', '草案', '']);
    var row = sh.getLastRow();
    setDropdown(sh, row, 8, ['Must', 'Should', 'Could']);
    setDropdown(sh, row, 10, ['草案', 'レビュー中', '合意済', '差し戻し', '廃止']);
    sh.getRange(row, 9).setBackground('#fffde7');
    sh.getRange(row, 4).setWrap(true);
    applyAllReferenceValidations_(ss);
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddUC() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var id = issueNextId(ss, 'UC');
    var sh = ss.getSheetByName(UC_LIST_SHEET_NAME);
    if (!sh) {
      notifyUser_('シート「' + UC_LIST_SHEET_NAME + '」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    // BR / FR などと同様に appendRow（テーブル API の矩形と insertRowAfter + setValues の組み合わせで
    // 「データは 1 行だが範囲は N 行」エラーになる場合があるため）
    sh.appendRow([id, '', '', '', '草案']);
    var row = sh.getLastRow();
    setDropdown(sh, row, 5, ['草案', 'レビュー中', '合意済', '保留', '廃止']);
    applyAllReferenceValidations_(ss);
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

/** 📖 UC詳細 の A 列見出し「▼ UC-xxx: …」の開始行。無ければ 0 */
function findUcDetailBlockStartRow_(detailSh, ucIdToken) {
  var lr = detailSh.getLastRow();
  var prefix = '▼ ' + ucIdToken;
  for (var r = 1; r <= lr; r++) {
    var t = String(detailSh.getRange(r, 1).getValue()).trim();
    if (t.indexOf(prefix) === 0) return r;
  }
  return 0;
}

/**
 * 追記用の先頭行。A 列に値がある最終行の直後に空行 1 行を挟む（書式だけ伸びた getLastRow に依存しない）。
 */
function getUcDetailAppendStartRow_(detailSh) {
  var lr = detailSh.getLastRow();
  if (lr < 1) return 1;
  var vals = detailSh.getRange(1, 1, lr, 1).getValues();
  var maxR = 0;
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() !== '') maxR = i + 1;
  }
  if (maxR === 0) return 1;
  return maxR + 2;
}

/**
 * アクティブセルが 📖 UC一覧 のデータ行のとき、その UC の詳細ブロックを 📖 UC詳細 に追加する。
 * 既にブロックがある場合はジャンプを確認する。
 */
function menuAppendUcDetailFromListRow() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var listSh = ss.getActiveSheet();
    if (listSh.getName() !== UC_LIST_SHEET_NAME) {
      notifyUser_('「' + UC_LIST_SHEET_NAME + '」タブを表示し、追加したい UC の行を選択してから実行してください。', 'UC 詳細');
      return;
    }
    var row = ss.getActiveRange().getRow();
    if (row < 2) {
      notifyUser_('データ行（2行目以降）を選択してください。', 'UC 詳細');
      return;
    }
    var ucId = String(listSh.getRange(row, 1).getValue()).trim();
    var actor = String(listSh.getRange(row, 2).getValue()).trim();
    var ucName = String(listSh.getRange(row, 3).getValue()).trim();
    if (!/^UC-\d+$/.test(ucId)) {
      notifyUser_('A列に UC-nnn 形式の UCID がある行を選択してください。', 'UC 詳細');
      return;
    }
    var detailSh = ss.getSheetByName(UC_DETAIL_SHEET_NAME);
    if (!detailSh) {
      notifyUser_('シート「' + UC_DETAIL_SHEET_NAME + '」がありません。createRequirementsSheet を実行してください。', 'UC 詳細');
      return;
    }
    var existing = findUcDetailBlockStartRow_(detailSh, ucId);
    if (existing > 0) {
      var ui = SpreadsheetApp.getUi();
      var resp = ui.alert(
        'UC 詳細',
        ucId + ' の詳細ブロックは既に「' + UC_DETAIL_SHEET_NAME + '」にあります。該当の見出しセルへ移動しますか？',
        ui.ButtonSet.YES_NO
      );
      if (resp === ui.Button.YES) {
        ss.setActiveSheet(detailSh);
        detailSh.getRange(existing, 1).activate();
      }
      return;
    }
    var startRow = getUcDetailAppendStartRow_(detailSh);
    writeUcDetailBlockAtRow_(detailSh, startRow, ucId, ucName || ucId, actor || '', true);
    ss.setActiveSheet(detailSh);
    detailSh.getRange(startRow, 1).activate();
    toastDone_('UC 詳細に項目見出しを追加しました（本文・フロー表は自分で記入）。', 'UC 詳細');
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddNFR() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var id = issueNextId(ss, 'NFR');
    var sh = ss.getSheetByName('🔒 非機能要求');
    if (!sh) {
      notifyUser_('シート「🔒 非機能要求」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '性能', '', '', '', '', '', '草案']);
    var row = sh.getLastRow();
    setDropdown(sh, row, 2, ['性能', '可用性', 'セキュリティ', '保守性', 'UX']);
    setDropdown(sh, row, 8, ['草案', 'レビュー中', '合意済', '差し戻し', '廃止']);
    sh.getRange(row, 7).setBackground('#fffde7');
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddCON() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var id = issueNextId(ss, 'CON');
    var sh = ss.getSheetByName('🚧 制約条件');
    if (!sh) {
      notifyUser_('シート「🚧 制約条件」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '技術', '', '', '', '草案']);
    var row = sh.getLastRow();
    setDropdown(sh, row, 2, ['技術', 'ビジネス', '法規制', '運用']);
    setDropdown(sh, row, 6, ['草案', '合意済', '廃止']);
    sh.getRange(row, 5).setBackground('#fffde7');
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddIF() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var id = issueNextId(ss, 'IF');
    var sh = ss.getSheetByName('🔗 外部IF');
    if (!sh) {
      notifyUser_('シート「🔗 外部IF」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', 'OUT（送信）', '', '', '', '', '']);
    var row = sh.getLastRow();
    setDropdown(sh, row, 3, ['IN（受信）', 'OUT（送信）', '双方向']);
    applyAllReferenceValidations_(ss);
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddOI() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var id = issueNextId(ss, 'OI');
    var sh = ss.getSheetByName('❓ 未解決事項');
    if (!sh) {
      notifyUser_('シート「❓ 未解決事項」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '', '', '', '', '未解決']);
    var row = sh.getLastRow();
    setDropdown(sh, row, 7, ['未解決', '解決済', '保留', '取り下げ']);
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddASM() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var id = issueNextId(ss, 'ASM');
    var sh = ss.getSheetByName('📌 前提条件');
    if (!sh) {
      notifyUser_('シート「📌 前提条件」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '']);
    var row = sh.getLastRow();
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddACT() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var id = issueNextId(ss, 'ACT');
    var sh = ss.getSheetByName('👤 アクター');
    if (!sh) {
      notifyUser_('シート「👤 アクター」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '', '', '']);
    var row = sh.getLastRow();
    applyAllReferenceValidations_(ss);
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

// ═════════════════════════════════════════════
// Markdown 出力（メニュー「📝 Markdown を作成」）
// ═════════════════════════════════════════════

function exportRequirementsToMarkdown() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let md = '# 要求仕様書\n\n';

    const overviewSheet = ss.getSheetByName('📋 概要');
    if (overviewSheet) {
      md += '## 📋 概要\n\n';

      md += '### ドキュメント管理\n';
      md += flattenOverviewDocManagementTable(overviewSheet) + '\n';

      md += '### プロジェクト概要\n';
      md += '- **目的:** ' + escapeMarkdown(overviewSheet.getRange('B10').getValue()) + '\n';
      md += '- **現状（As-Is）:** ' + escapeMarkdown(overviewSheet.getRange('B11').getValue()) + '\n';
      md += '- **課題:** ' + escapeMarkdown(overviewSheet.getRange('B12').getValue()) + '\n\n';

      md += '### スコープ（IN）\n\n';
      md += overviewScopeBulletBlock(overviewSheet, 14, 16);
      md += '\n### スコープ（OUT）\n\n';
      md += overviewScopeBulletBlock(overviewSheet, 18, 20);
      md += '\n### 成功指標\n';
      md += extractTableAsMarkdown(overviewSheet, 22, 1, 4) + '\n\n';
    }

    const ucList = ss.getSheetByName(UC_LIST_SHEET_NAME);
    const ucDetail = ss.getSheetByName(UC_DETAIL_SHEET_NAME);
    const legacyUcSheet = ss.getSheetByName('📖 ユースケース');
    const actorMap = readActorMap_(ss);
    const actorNameToId = readActorNameToIdMap_(ss);
    if (ucList || ucDetail) {
      md += parseUseCaseSheets_(ucList, ucDetail, actorMap, actorNameToId);
    } else if (legacyUcSheet) {
      md += parseLegacyCombinedUseCaseSheet_(legacyUcSheet, actorMap, actorNameToId);
    }

    const sheetsToProcess = [
      { name: '👤 アクター', startRow: 1, cols: 5 },
      { name: '🎯 ビジネス要求', startRow: 1, cols: 7 },
      { name: '⚙️ 機能要求', startRow: 1, cols: 11 },
      { name: '🔒 非機能要求', startRow: 1, cols: 8 },
      { name: '🚧 制約条件', startRow: 1, cols: 6 },
      { name: '🔗 外部IF', startRow: 1, cols: 8 },
      { name: '📌 前提条件', startRow: 1, cols: 3 },
      { name: '❓ 未解決事項', startRow: 1, cols: 7 },
      { name: '📚 用語集', startRow: 1, cols: 4 },
      { name: '✅ 変更履歴', startRow: 1, cols: 5 }
    ];

    sheetsToProcess.forEach(function (info) {
      const sheet = ss.getSheetByName(info.name);
      if (sheet) {
        md += '## ' + info.name + '\n\n';
        if (info.name === '🔗 外部IF') {
          md += extractExternalIfTableAsMarkdown_(sheet, actorMap, actorNameToId) + '\n\n';
        } else {
          md += extractTableAsMarkdown(sheet, info.startRow, 1, info.cols) + '\n\n';
        }
      }
    });

    showMarkdownDialog(md);
  } catch (e) {
    notifyUser_('Markdown の作成に失敗しました。\n' + String(e.message || e), 'Markdown');
  }
}

function flattenOverviewDocManagementTable(sheet) {
  const rows = [['項目', '内容']];
  for (let r = 3; r <= 6; r++) {
    rows.push([sheet.getRange(r, 1).getValue(), sheet.getRange(r, 2).getValue()]);
    rows.push([sheet.getRange(r, 3).getValue(), sheet.getRange(r, 4).getValue()]);
  }
  return arrayToMarkdownTable(rows);
}

function overviewScopeBulletBlock(sheet, startRow, endRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return '\n';

  const end = Math.min(endRow, lastRow);
  const numRows = end - startRow + 1;
  const data = sheet.getRange(startRow, 1, numRows, 4).getValues();

  const lines = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const label = String(row[0]).trim();
    const b = String(row[1]).trim();
    const c = String(row[2]).trim();
    const d = String(row[3]).trim();
    const rest = [b, c, d].filter(function (x) {
      return x !== '';
    }).join(' ');

    if (!label && !rest) continue;

    if (label && rest) {
      lines.push('- **' + escapeMarkdown(label) + ':** ' + escapeMarkdown(rest));
    } else if (rest) {
      lines.push('- ' + escapeMarkdown(rest));
    } else {
      lines.push('- ' + escapeMarkdown(label));
    }
  }

  return lines.length ? lines.join('\n') + '\n' : '\n';
}

/** 📖 UC一覧 + 📖 UC詳細 を Markdown にする */
function parseUseCaseSheets_(listSheet, detailSheet, actorMap, actorNameToId) {
  let md = '## 📖 ユースケース\n\n';
  if (listSheet && listSheet.getLastRow() >= 1) {
    md += '### ▼ ユースケース一覧\n\n';
    md += extractUcListTableAsMarkdown_(listSheet, actorMap, actorNameToId) + '\n\n';
  }
  if (detailSheet && detailSheet.getLastRow() > 0) {
    md += parseUseCaseDetailSheet_(detailSheet, actorMap, actorNameToId);
  }
  return md;
}

/** 📖 UC詳細 のみ（▼ UC-xxx ブロック・フロー） */
function parseUseCaseDetailSheet_(sheet, actorMap, actorNameToId) {
  let md = '';
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0) return md;

  const data = sheet.getRange(1, 1, lastRow, Math.max(lastCol, 5)).getValues();
  let i = 0;

  while (i < data.length) {
    const cellA = String(data[i][0]).trim();

    if (cellA.startsWith('▼ UC-')) {
      md += '### ' + cellA + '\n\n';
      i++;
      const metaTable = [['項目', '内容']];
      while (i < data.length) {
        const nextCellA = String(data[i][0]).trim();
        if (nextCellA === '基本フロー' || nextCellA === '代替フロー' || nextCellA.startsWith('▼') || nextCellA === '') {
          break;
        }
        let metaVal = data[i][1];
        if (String(data[i][0]).trim() === 'アクター') {
          metaVal = resolveActorLabelForMarkdown_(metaVal, actorMap, actorNameToId);
        }
        metaTable.push([data[i][0], metaVal]);
        i++;
      }
      if (metaTable.length > 1) {
        md += arrayToMarkdownTable(metaTable) + '\n\n';
      }
      continue;
    }

    if (cellA === '基本フロー' || cellA === '代替フロー') {
      md += '#### ' + cellA + '\n\n';
      i++;
      const flowTable = [['No.', 'アクション']];
      while (i < data.length) {
        const nextCellA = String(data[i][0]).trim();
        const nextCellB = String(data[i][1]).trim();

        if (nextCellA.startsWith('▼') || nextCellA === '基本フロー' || nextCellA === '代替フロー') break;
        if (nextCellA === '' && nextCellB === '') break;

        flowTable.push([data[i][0], data[i][1]]);
        i++;
      }
      if (flowTable.length > 1) {
        md += arrayToMarkdownTable(flowTable) + '\n\n';
      }

      while (i < data.length && String(data[i][0]).trim() === '' && String(data[i][1]).trim() === '') {
        i++;
      }
      continue;
    }

    i++;
  }

  return md;
}

/** 旧テンプレ「📖 ユースケース」1シートに一覧と詳細が同居していた場合の書き出し */
function parseLegacyCombinedUseCaseSheet_(sheet, actorMap, actorNameToId) {
  let md = '## 📖 ユースケース\n\n';
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0) return md;

  const data = sheet.getRange(1, 1, lastRow, Math.max(lastCol, 5)).getValues();
  let i = 0;

  while (i < data.length) {
    const cellA = String(data[i][0]).trim();

    if (cellA === '▼ ユースケース一覧') {
      md += '### ' + cellA + '\n\n';
      i++;
      const tableData = [];
      while (i < data.length && String(data[i][0]).trim() !== '' && !String(data[i][0]).startsWith('▼')) {
        var listRow = data[i].slice(0, 5);
        if (tableData.length >= 1) {
          listRow[1] = resolveActorLabelForMarkdown_(listRow[1], actorMap, actorNameToId);
        }
        tableData.push(listRow);
        i++;
      }
      md += arrayToMarkdownTable(tableData) + '\n\n';
      continue;
    }

    if (cellA.startsWith('▼ UC-')) {
      md += '### ' + cellA + '\n\n';
      i++;
      const metaTable = [['項目', '内容']];
      while (i < data.length) {
        const nextCellA = String(data[i][0]).trim();
        if (nextCellA === '基本フロー' || nextCellA === '代替フロー' || nextCellA.startsWith('▼') || nextCellA === '') {
          break;
        }
        var legMetaVal = data[i][1];
        if (String(data[i][0]).trim() === 'アクター') {
          legMetaVal = resolveActorLabelForMarkdown_(legMetaVal, actorMap, actorNameToId);
        }
        metaTable.push([data[i][0], legMetaVal]);
        i++;
      }
      if (metaTable.length > 1) {
        md += arrayToMarkdownTable(metaTable) + '\n\n';
      }
      continue;
    }

    if (cellA === '基本フロー' || cellA === '代替フロー') {
      md += '#### ' + cellA + '\n\n';
      i++;
      const flowTable = [['No.', 'アクション']];
      while (i < data.length) {
        const nextCellA = String(data[i][0]).trim();
        const nextCellB = String(data[i][1]).trim();

        if (nextCellA.startsWith('▼') || nextCellA === '基本フロー' || nextCellA === '代替フロー') break;
        if (nextCellA === '' && nextCellB === '') break;

        flowTable.push([data[i][0], data[i][1]]);
        i++;
      }
      if (flowTable.length > 1) {
        md += arrayToMarkdownTable(flowTable) + '\n\n';
      }

      while (i < data.length && String(data[i][0]).trim() === '' && String(data[i][1]).trim() === '') {
        i++;
      }
      continue;
    }

    i++;
  }

  return md;
}

function extractTableAsMarkdown(sheet, startRow, startCol, numCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return '';
  const numRows = lastRow - startRow + 1;
  const data = sheet.getRange(startRow, startCol, numRows, numCols).getValues();

  const filteredData = data.filter(function (row) {
    return row.join('').trim() !== '';
  });
  return arrayToMarkdownTable(filteredData);
}

/** 🔗 外部IF の Markdown。「連携先システム」は UC と同様に ACT-xxx（名前）へ解決。 */
function extractExternalIfTableAsMarkdown_(sheet, actorMap, actorNameToId) {
  var startRow = 1;
  var startCol = 1;
  var numCols = 8;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return '';
  var numRows = lastRow - startRow + 1;
  var data = sheet.getRange(startRow, startCol, numRows, numCols).getValues();
  var filteredData = data.filter(function (row) {
    return row.join('').trim() !== '';
  });
  var j;
  for (j = 1; j < filteredData.length; j++) {
    var rowCopy = filteredData[j].slice();
    if (/^IF-\d+$/.test(String(rowCopy[0]).trim())) {
      rowCopy[1] = resolveActorLabelForMarkdown_(rowCopy[1], actorMap, actorNameToId);
    }
    filteredData[j] = rowCopy;
  }
  return arrayToMarkdownTable(filteredData);
}

/** 📖 UC一覧 の Markdown（アクター名列は出力時に ACT-xxx（名前）へ解決）。 */
function extractUcListTableAsMarkdown_(sheet, actorMap, actorNameToId) {
  var startRow = 1;
  var startCol = 1;
  var numCols = 5;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return '';
  var numRows = lastRow - startRow + 1;
  var data = sheet.getRange(startRow, startCol, numRows, numCols).getValues();
  var filteredData = data.filter(function (row) {
    return row.join('').trim() !== '';
  });
  var j;
  for (j = 1; j < filteredData.length; j++) {
    var rowCopy = filteredData[j].slice();
    rowCopy[1] = resolveActorLabelForMarkdown_(rowCopy[1], actorMap, actorNameToId);
    filteredData[j] = rowCopy;
  }
  return arrayToMarkdownTable(filteredData);
}

function arrayToMarkdownTable(data) {
  if (!data || data.length === 0) return '';

  const headers = data[0];
  let md = '| ' + headers.map(function (h) { return escapeMarkdown(h); }).join(' | ') + ' |\n';
  md += '| ' + headers.map(function () { return '---'; }).join(' | ') + ' |\n';

  for (let i = 1; i < data.length; i++) {
    const rowStr = data[i].map(function (val) { return escapeMarkdown(val); }).join(' | ');
    md += '| ' + rowStr + ' |\n';
  }
  return md;
}

function escapeMarkdown(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val)
    .replace(/\r?\n/g, '<br>')
    .replace(/\|/g, '\\|');
}

function showMarkdownDialog(mdText) {
  const encodedMd = Utilities.base64Encode(Utilities.newBlob(mdText).getBytes());

  const htmlTemplate =
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<style>' +
    'body { font-family: sans-serif; margin: 10px; }' +
    'textarea { width: 100%; height: 350px; font-family: monospace; font-size: 14px; padding: 10px; box-sizing: border-box; }' +
    'button { padding: 10px 20px; font-size: 14px; cursor: pointer; background-color: #1a73e8; color: white; border: none; border-radius: 4px; margin-top: 10px; }' +
    'button:hover { background-color: #1557b0; }' +
    '</style>' +
    '</head>' +
    '<body>' +
    '<textarea id="mdOutput"></textarea>' +
    '<button onclick="copyToClipboard()">クリップボードにコピー</button>' +
    '<span id="msg" style="margin-left: 10px; color: green; display: none;">✅ コピーしました！</span>' +
    '<script>' +
    'var encodedStr = "' +
    encodedMd +
    '";' +
    'var decodedStr = decodeURIComponent(escape(atob(encodedStr)));' +
    'var textArea = document.getElementById("mdOutput");' +
    'textArea.value = decodedStr;' +
    'function copyToClipboard() {' +
    'textArea.select();' +
    'document.execCommand("copy");' +
    'var msg = document.getElementById("msg");' +
    'msg.style.display = "inline";' +
    'setTimeout(function() { msg.style.display = "none"; }, 2000);' +
    '}' +
    '</script>' +
    '</body>' +
    '</html>';

  const htmlOutput = HtmlService.createHtmlOutput(htmlTemplate)
    .setWidth(800)
    .setHeight(480)
    .setTitle('Markdown 出力結果');

  showModalDialogSafe_(htmlOutput, '📝 Markdown 出力');
}
