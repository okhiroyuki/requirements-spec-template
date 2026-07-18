/** 🔢 ID管理 シートの読み書き・採番ロジック（BR-001 のようなゼロ埋めID文字列の生成を含む）。 */

/** 🔢 ID管理：ヘッダのみ（seed 後 syncIdCountersFromBookCore で中身を埋める） */
function setupIdSheetHeaderOnly_(ss) {
  var sh = getOrCreateSheet(ss, ID_SHEET_NAME);
  resetSheetCellsForTemplate_(sh);
  sh.getRange(1, 1, 1, 3).setValues([['キー', '最終発番（数値）', '説明']]);
  styleHeader(sh, 1, 3);
  setColWidths(sh, [100, 130, 320]);
  try {
    sh.hideSheet();
  } catch (e) {}
}


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

  scanColumn(BUC_SHEET_NAME, 1, function (text) {
    var m = text.match(/^BUC-(\d+)$/);
    if (m) bump('BUC', m[1]);
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
    var m2 = text.match(/^NFR-([A-Z])(\d+)$/);
    if (m2) bump('NFR', m2[2]);
  });

  scanColumn('🚧 制約条件', 1, function (text) {
    var m1 = text.match(/^CON-(\d+)$/);
    if (m1) bump('CON', m1[1]);
    var m2 = text.match(/^CON-([A-Z])(\d+)$/);
    if (m2) bump('CON', m2[2]);
  });

  ;[UC_LIST_SHEET_NAME, UC_DETAIL_SHEET_NAME, BUC_DETAIL_SHEET_NAME].forEach(function (name) {
    var ucSh = ss.getSheetByName(name);
    if (!ucSh) return;
    var lr2 = ucSh.getLastRow();
    for (var r = 1; r <= lr2; r++) {
      var text = String(ucSh.getRange(r, 1).getValue()).trim();
      var m1 = text.match(/^UC-(\d+)$/);
      if (m1) bump('UC', m1[1]);
      var m2 = text.match(/▼\s*UC-(\d+)/);
      if (m2) bump('UC', m2[1]);
      var m3 = text.match(/▼\s*BUC-(\d+)/);
      if (m3) bump('BUC', m3[1]);
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
  var simple = {
    BR: 'BR-',
    BUC: 'BUC-',
    FR: 'FR-',
    UC: 'UC-',
    IF: 'IF-',
    OI: 'OI-',
    ACT: 'ACT-',
    ASM: 'ASM-',
    CON: 'CON-',
    NFR: 'NFR-',
  };
  var p = simple[counterKey];
  if (!p) throw new Error('不正なキー: ' + counterKey);
  return p + pad;
}
