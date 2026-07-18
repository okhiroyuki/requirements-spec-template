/**
 * 要求仕様書スプレッドシート — テンプレート自動生成（メイン：定数・createRequirementsSheet・共通UIヘルパー）
 *
 * Apps Script プロジェクトには本ファイルのほか、validation.gs / template-sheets.gs /
 * ids.gs / menu.gs / markdown-export.gs を同じプロジェクトに追加する（同一プロジェクト内では
 * ファイルをまたいで関数・var を共有できるため import は不要）。セットアップ手順は README.md 参照。
 *
 * 関数「createRequirementsSheet」を実行すると、実行のたびに全シートが初期サンプルで上書きされる
 * （確認ダイアログなし）。作成完了ダイアログにメニュー利用の注意（再読み込み）が表示される。
 */

var UC_LIST_SHEET_NAME = '📖 UC一覧';
var UC_DETAIL_SHEET_NAME = '📖 UC詳細';
/** BUC：事業側の業務単位。BR に紐づく。 */
var BUC_SHEET_NAME = '📗 BUC';
/** 業務単位ごとの手順・行動内容・関連 UC（一覧シートとは別）。 */
var BUC_DETAIL_SHEET_NAME = '📙 BUC詳細';
/** 📗 BUC の「参考：ビジネス要求」（E 列）が参照するシート（要求ID・ビジネス要求（1文））。 */
var BUSINESS_REQ_SHEET_NAME = '🎯 ビジネス要求';

var ID_SHEET_NAME = '🔢 ID管理';

/** 同期時に必ず行を用意するキー（テンプレに現れないキーは 0） */
var ID_COUNTER_KEYS = [
  'BR',
  'BUC',
  'FR',
  'UC',
  'IF',
  'OI',
  'ACT',
  'ASM',
  'NFR',
  'CON'
];

/** 全シートをクリアし、初期サンプルを再展開する（確認ダイアログなし）。 */
function createRequirementsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length === 1) {
    const tmp = ss.insertSheet('_tmp');
    ss.deleteSheet(defaultSheet);
  }

  setupOverview(ss);
  setupActors(ss);
  setupBusinessReqs(ss);
  setupBusinessUseCases(ss);
  setupBucDetail(ss);
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


function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** 値・書式・入力規則をクリアしてテンプレ再展開の前提にする。 */
function resetSheetCellsForTemplate_(sh, maxRows, maxCols) {
  if (!sh) return;
  sh.clearContents();
  sh.clearFormats();
  var rows = Math.min(Math.max(parseInt(maxRows, 10) || 500, 1), sh.getMaxRows());
  var cols = Math.min(Math.max(parseInt(maxCols, 10) || 40, 1), sh.getMaxColumns());
  try {
    sh.getRange(1, 1, rows, cols).clearDataValidations();
  } catch (e) {
    Logger.log('resetSheetCellsForTemplate_(' + sh.getName() + '): ' + (e && e.message ? e.message : e));
  }
}

/** テンプレのシートタブ順を固定する。 */
function reorderReqSpecSheetTabs_(ss) {
  var names = [
    '📋 概要',
    '📌 前提条件',
    '👤 アクター',
    '🎯 ビジネス要求',
    BUC_SHEET_NAME,
    BUC_DETAIL_SHEET_NAME,
    UC_LIST_SHEET_NAME,
    UC_DETAIL_SHEET_NAME,
    '⚙️ 機能要求',
    '🔒 非機能要求',
    '🚧 制約条件',
    '🔗 外部IF',
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

/** Toast で完了を通知する。 */
function toastDone_(message, title) {
  title = title || '完了';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(message, title, 5);
  } catch (ignore) {
    Logger.log('[' + title + '] ' + message);
  }
}

/**
 * ヘッダー行のスタイル設定。
 * @param {boolean} [freezeHeaderRow=true] false のとき、行の固定（setFrozenRows）は行わない（📋 概要の成功指標表など）。
 */
function styleHeader(sheet, row, cols, freezeHeaderRow) {
  const range = sheet.getRange(row, 1, 1, cols);
  range.setBackground('#1a73e8')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setVerticalAlignment('middle');
  if (freezeHeaderRow !== false) {
    sheet.setFrozenRows(row);
  }
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
