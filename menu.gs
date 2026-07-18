/** カスタムメニュー（onOpen）、行追加パネル、BUC／UC 詳細ブロックの追加アクション。 */

function showAddRowPanel() {
  let html = HtmlService.createHtmlOutput(getAddRowPanelHtml_()).setTitle('行を追加');
  showSidebarSafe_(html);
}

function getAddRowPanelHtml_() {
  let esc = function (t) {
    return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  };
  let fail =
    'function fail(e){alert(e&&e.message?e.message:String(e));}';
  let btn = function (fn, label) {
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
      btn('menuAddASM', 'ASM · 📌 前提条件') +
      btn('menuAddACT', 'ACT · 👤 アクター') +
      btn('menuAddBR', 'BR · 🎯 ビジネス要求') +
      btn('menuAddBUC', 'BUC · 📗 BUC') +
      btn('menuAddUC', 'UC · 📖 UC一覧') +
      btn('menuAddFR', 'FR · ⚙️ 機能要求') +
      btn('menuAddNFR', 'NFR · 🔒 非機能要求') +
      btn('menuAddCON', 'CON · 🚧 制約条件') +
      btn('menuAddIF', 'IF · 🔗 外部 IF') +
      btn('menuAddOI', 'OI · ❓ 未解決事項') +
      '<script>' +
      fail +
      '</script>' +
      '</body></html>'
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('要求仕様書')
    .addItem('🔢 IDカウンタをブックから再同期', 'syncIdCountersFromBook')
    .addItem('🔗 入力規則をすべて更新（BR／UC／アクター連携）', 'menuRefreshAllInputValidations')
    .addItem('🧩 行を追加パネル（サイドバー）', 'showAddRowPanel')
    .addItem('📙 選択行の BUC 詳細を追加／更新', 'menuAppendBucDetailFromListRow')
    .addItem('📖 選択行の UC 詳細を追加／更新', 'menuAppendUcDetailFromListRow')
    .addSeparator()
    .addItem('📝 Markdown を作成（表示・コピー）', 'exportRequirementsToMarkdown')
    .addToUi();
}

function menuAddBUC() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let id = issueNextId(ss, 'BUC');
    let sh = ss.getSheetByName(BUC_SHEET_NAME);
    if (!sh) {
      notifyUser_('シート「' + BUC_SHEET_NAME + '」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '', '', '']);
    let newRow = sh.getLastRow();
    sh.getRange(newRow, 3).setWrap(true);
    sh.getRange(newRow, 5).setFormula(bucBrMirrorFormula_(newRow));
    sh.getRange(newRow, 5).setWrap(true);
    applyAllReferenceValidations_(ss);
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddBR() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let id = issueNextId(ss, 'BR');
    let sh = ss.getSheetByName('🎯 ビジネス要求');
    if (!sh) {
      notifyUser_('シート「🎯 ビジネス要求」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '', 'Must', '', '', '草案']);
    let row = sh.getLastRow();
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
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let id = issueNextId(ss, 'FR');
    let sh = ss.getSheetByName('⚙️ 機能要求');
    if (!sh) {
      notifyUser_('シート「⚙️ 機能要求」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '', '', '', '', '', 'Must', '', '草案', '']);
    let row = sh.getLastRow();
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
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let id = issueNextId(ss, 'UC');
    let sh = ss.getSheetByName(UC_LIST_SHEET_NAME);
    if (!sh) {
      notifyUser_('シート「' + UC_LIST_SHEET_NAME + '」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '', '', '草案']);
    let row = sh.getLastRow();
    setDropdown(sh, row, 5, ['草案', 'レビュー中', '合意済', '保留', '廃止']);
    applyAllReferenceValidations_(ss);
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

/** 📖 UC詳細 の A 列見出し「▼ UC-xxx: …」の開始行。無ければ 0 */
function findUcDetailBlockStartRow_(detailSh, ucIdToken) {
  let lr = detailSh.getLastRow();
  let prefix = '▼ ' + ucIdToken;
  for (let r = 1; r <= lr; r++) {
    let t = String(detailSh.getRange(r, 1).getValue()).trim();
    if (t.indexOf(prefix) === 0) return r;
  }
  return 0;
}

/**
 * 追記用の先頭行。A 列に値がある最終行の直後に空行 1 行を挟む（書式だけ伸びた getLastRow に依存しない）。
 */
function getUcDetailAppendStartRow_(detailSh) {
  let lr = detailSh.getLastRow();
  if (lr < 1) return 1;
  let vals = detailSh.getRange(1, 1, lr, 1).getValues();
  let maxR = 0;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() !== '') maxR = i + 1;
  }
  if (maxR === 0) return 1;
  return maxR + 2;
}

/** 📙 BUC詳細 の A 列見出し「▼ BUC-xxx: …」の開始行。無ければ 0 */
function findBucDetailBlockStartRow_(detailSh, bucIdToken) {
  let lr = detailSh.getLastRow();
  let prefix = '▼ ' + bucIdToken;
  for (let rb = 1; rb <= lr; rb++) {
    let tb = String(detailSh.getRange(rb, 1).getValue()).trim();
    if (tb.indexOf(prefix) === 0) return rb;
  }
  return 0;
}

/** 📙 BUC詳細 への追記開始行（末尾ブロックのあとに空行を挟む）。 */
function getBucDetailAppendStartRow_(detailSh) {
  let lrBd = detailSh.getLastRow();
  if (lrBd < 1) return 1;
  let valsBd = detailSh.getRange(1, 1, lrBd, 1).getValues();
  let maxRb = 0;
  let ib;
  for (ib = 0; ib < valsBd.length; ib++) {
    if (String(valsBd[ib][0]).trim() !== '') maxRb = ib + 1;
  }
  if (maxRb === 0) return 1;
  return maxRb + 2;
}

/** 📗 BUC の選択行から 📙 BUC詳細 に手順ブロックを追加する。 */
function menuAppendBucDetailFromListRow() {
  try {
    let ssBd = SpreadsheetApp.getActiveSpreadsheet();
    let listBd = ssBd.getActiveSheet();
    if (listBd.getName() !== BUC_SHEET_NAME) {
      notifyUser_('「' + BUC_SHEET_NAME + '」タブを表示し、追加したい業務の行を選択してから実行してください。', 'BUC 詳細');
      return;
    }
    let rowBd = ssBd.getActiveRange().getRow();
    if (rowBd < 2) {
      notifyUser_('データ行（2行目以降）を選択してください。', 'BUC 詳細');
      return;
    }
    let bucId = String(listBd.getRange(rowBd, 1).getValue()).trim();
    let bucName = String(listBd.getRange(rowBd, 2).getValue()).trim();
    if (!/^BUC-\d+$/.test(bucId)) {
      notifyUser_('A列に BUC-nnn 形式の BUCID がある行を選択してください。', 'BUC 詳細');
      return;
    }
    let detailBd = ssBd.getSheetByName(BUC_DETAIL_SHEET_NAME);
    if (!detailBd) {
      notifyUser_('シート「' + BUC_DETAIL_SHEET_NAME + '」がありません。createRequirementsSheet を実行してください。', 'BUC 詳細');
      return;
    }
    let existingBd = findBucDetailBlockStartRow_(detailBd, bucId);
    if (existingBd > 0) {
      let uiBd = SpreadsheetApp.getUi();
      let respBd = uiBd.alert(
        'BUC 詳細',
        bucId + ' の詳細ブロックは既に「' + BUC_DETAIL_SHEET_NAME + '」にあります。該当の見出しセルへ移動しますか？',
        uiBd.ButtonSet.YES_NO
      );
      if (respBd === uiBd.Button.YES) {
        ssBd.setActiveSheet(detailBd);
        detailBd.getRange(existingBd, 1).activate();
      }
      return;
    }
    let startRowBd = getBucDetailAppendStartRow_(detailBd);
    writeBucDetailBlockAtRow_(detailBd, startRowBd, bucId, bucName || bucId, true, []);
    applyAllReferenceValidations_(ssBd);
    ssBd.setActiveSheet(detailBd);
    detailBd.getRange(startRowBd, 1).activate();
    toastDone_(
      'BUC詳細に手順表を追加しました。1行目に手順番号と関連UC（📖 UC一覧）の入力規則が付きます。行動内容は自分で記入してください。',
      'BUC 詳細'
    );
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

/** 📖 UC一覧 の選択行から 📖 UC詳細 にブロックを追加する。 */
function menuAppendUcDetailFromListRow() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let listSh = ss.getActiveSheet();
    if (listSh.getName() !== UC_LIST_SHEET_NAME) {
      notifyUser_('「' + UC_LIST_SHEET_NAME + '」タブを表示し、追加したい UC の行を選択してから実行してください。', 'UC 詳細');
      return;
    }
    let row = ss.getActiveRange().getRow();
    if (row < 2) {
      notifyUser_('データ行（2行目以降）を選択してください。', 'UC 詳細');
      return;
    }
    let ucId = String(listSh.getRange(row, 1).getValue()).trim();
    let actor = String(listSh.getRange(row, 2).getValue()).trim();
    let ucName = String(listSh.getRange(row, 3).getValue()).trim();
    if (!/^UC-\d+$/.test(ucId)) {
      notifyUser_('A列に UC-nnn 形式の UCID がある行を選択してください。', 'UC 詳細');
      return;
    }
    let detailSh = ss.getSheetByName(UC_DETAIL_SHEET_NAME);
    if (!detailSh) {
      notifyUser_('シート「' + UC_DETAIL_SHEET_NAME + '」がありません。createRequirementsSheet を実行してください。', 'UC 詳細');
      return;
    }
    let existing = findUcDetailBlockStartRow_(detailSh, ucId);
    if (existing > 0) {
      let ui = SpreadsheetApp.getUi();
      let resp = ui.alert(
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
    let startRow = getUcDetailAppendStartRow_(detailSh);
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
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let id = issueNextId(ss, 'NFR');
    let sh = ss.getSheetByName('🔒 非機能要求');
    if (!sh) {
      notifyUser_('シート「🔒 非機能要求」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '性能', '', '', '', '', '', '草案']);
    let row = sh.getLastRow();
    setDropdown(sh, row, 2, ['性能', '可用性', 'セキュリティ', '保守性', 'UX']);
    setDropdown(sh, row, 8, ['草案', 'レビュー中', '合意済', '差し戻し', '廃止']);
    sh.getRange(row, 7).setBackground('#fffde7');
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddCON() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let id = issueNextId(ss, 'CON');
    let sh = ss.getSheetByName('🚧 制約条件');
    if (!sh) {
      notifyUser_('シート「🚧 制約条件」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '技術', '', '', '', '草案']);
    let row = sh.getLastRow();
    setDropdown(sh, row, 2, ['技術', 'ビジネス', '法規制', '運用']);
    setDropdown(sh, row, 6, ['草案', '合意済', '廃止']);
    sh.getRange(row, 5).setBackground('#fffde7');
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddIF() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let id = issueNextId(ss, 'IF');
    let sh = ss.getSheetByName('🔗 外部IF');
    if (!sh) {
      notifyUser_('シート「🔗 外部IF」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', 'OUT（送信）', '', '', '', '', '']);
    let row = sh.getLastRow();
    setDropdown(sh, row, 3, ['IN（受信）', 'OUT（送信）', '双方向']);
    applyAllReferenceValidations_(ss);
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddOI() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let id = issueNextId(ss, 'OI');
    let sh = ss.getSheetByName('❓ 未解決事項');
    if (!sh) {
      notifyUser_('シート「❓ 未解決事項」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '', '', '', '', '未解決']);
    let row = sh.getLastRow();
    setDropdown(sh, row, 7, ['未解決', '解決済', '保留', '取り下げ']);
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddASM() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let id = issueNextId(ss, 'ASM');
    let sh = ss.getSheetByName('📌 前提条件');
    if (!sh) {
      notifyUser_('シート「📌 前提条件」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '']);
    let row = sh.getLastRow();
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}

function menuAddACT() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let id = issueNextId(ss, 'ACT');
    let sh = ss.getSheetByName('👤 アクター');
    if (!sh) {
      notifyUser_('シート「👤 アクター」がありません。createRequirementsSheet を実行してください。', '行を追加');
      return;
    }
    sh.appendRow([id, '', '', '', '']);
    let row = sh.getLastRow();
    applyAllReferenceValidations_(ss);
  } catch (e) {
    notifyUser_(String(e.message || e), 'エラー');
  }
}
