/** ドロップダウン・別シート参照（BR／UC／アクターなど）の入力規則、ステータス条件付き書式。 */

/**
 * 実データ行数を十分に超える見込みの余白行数。
 * UC-ID 一覧の参照範囲や条件付き書式など、「将来追加される行もあらかじめカバーしておきたい」
 * 用途で使う（都度スキャンする用途では使わない。sheet.getLastRow() は既に実データの最終行を
 * 返すので、そちらは追加の上限を設けない）。シートの実際の行数（getMaxRows）でクランプされる。
 */
var VALIDATION_ROW_HEADROOM = 2000;

/**
 * 👤 アクター の A 列→B 列（アクター名）マップ。
 * @return {!Object<string, string>}
 */
function readActorMap_(ss) {
  var sh = ss.getSheetByName('👤 アクター');
  if (!sh) return {};
  var lr = sh.getLastRow();
  if (lr < 2) return {};
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
function getFirstColumnIdRange_(sheet) {
  if (!sheet) return null;
  var lr = sheet.getLastRow();
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

/** 📖 UC一覧 A 列の UC-ID 範囲（入力規則の参照元。A2 起点で実データ+余白ぶんの行数を確保）。 */
function getUcIdListRange_(ss) {
  var sheet = ss.getSheetByName(UC_LIST_SHEET_NAME);
  if (!sheet) return null;
  var maxEnd = sheet.getMaxRows();
  if (maxEnd < 2) return null;
  var lastRow = Math.min(sheet.getLastRow(), maxEnd);
  if (lastRow < 2) return null;
  var endRow = Math.min(maxEnd, lastRow + VALIDATION_ROW_HEADROOM);
  var numRows = endRow - 2 + 1;
  return sheet.getRange(2, 1, numRows, 1);
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
        .setAllowInvalid(false)
        .build()
    : null;

  var lr = frSh.getLastRow();
  var r;
  try {
    for (r = 2; r <= lr; r++) {
      var frCell = String(frSh.getRange(r, 1).getValue()).trim();
      if (!/^FR-\d+$/.test(frCell)) continue;
      var cell = frSh.getRange(r, 3);
      if (rule) cell.setDataValidation(rule);
      else cell.clearDataValidations();
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

/**
 * 📗 BUC の E 列「参考：ビジネス要求」: D 列の関連BR に対応する 🎯 ビジネス要求 B 列を表示する数式。
 * @param {number} row 1 始まりの行番号（ヘッダーは 1 行目想定）
 */
function bucBrMirrorFormula_(row) {
  return (
    '=IF(D' +
    row +
    '="","",IFERROR(VLOOKUP(D' +
    row +
    ",'" +
    BUSINESS_REQ_SHEET_NAME +
    "'!$A:$B,2,FALSE),\"\"))"
  );
}

/**
 * 📗 BUC の「関連BR」列に、🎯 ビジネス要求の BR-ID を選ぶ入力規則を付与する。
 */
function applyBucRelatedBrValidation_(ss) {
  var bucSh = ss.getSheetByName(BUC_SHEET_NAME);
  if (!bucSh) return;
  var vr = getBrIdListRange_(ss);
  var rule = vr
    ? SpreadsheetApp.newDataValidation()
        .requireValueInRange(vr, true)
        .setAllowInvalid(false)
        .build()
    : null;

  var lr = bucSh.getLastRow();
  var r;
  try {
    for (r = 2; r <= lr; r++) {
      var idCell = String(bucSh.getRange(r, 1).getValue()).trim();
      if (!/^BUC-\d+$/.test(idCell)) continue;
      var cell = bucSh.getRange(r, 4);
      if (rule) cell.setDataValidation(rule);
      else cell.clearDataValidations();
    }
  } catch (e) {
    var msg = String(e.message || e);
    if (msg.indexOf('型付き') !== -1 || /typed column/i.test(msg)) {
      Logger.log('applyBucRelatedBrValidation_: skip typed table column — ' + msg);
      return;
    }
    throw e;
  }
}

/**
 * 📙 BUC詳細 手順表：B 列は自由入力、C 列は 📖 UC一覧 の UC-ID（該当なしは空欄）。
 * ブロックは次の「▼」行、または A〜C がすべて空の行まで。
 */
function applyBucDetailStepValidations_(ss) {
  var sh = ss.getSheetByName(BUC_DETAIL_SHEET_NAME);
  if (!sh) return;

  var vrUc = getUcIdListRange_(ss);
  var ruleUc = vrUc
    ? SpreadsheetApp.newDataValidation()
        .requireValueInRange(vrUc, true)
        .setAllowInvalid(false)
        .build()
    : null;

  var lrCap = sh.getLastRow();
  var r;

  try {
    r = 1;
    while (r <= lrCap) {
      var headingA = String(sh.getRange(r, 1).getValue()).trim();
      if (!/^▼\s*BUC-\d+/.test(headingA)) {
        r++;
        continue;
      }
      var hdrRow = r + 1;
      var sr;
      for (sr = hdrRow + 1; sr <= lrCap; sr++) {
        var qa = String(sh.getRange(sr, 1).getValue()).trim();
        var qb = String(sh.getRange(sr, 2).getValue()).trim();
        var qc = String(sh.getRange(sr, 3).getValue()).trim();

        if (qa.substring(0, 1) === '▼') break;

        var rowAllEmpty = qa === '' && qb === '' && qc === '';

        try {
          sh.getRange(sr, 2).clearDataValidations();
        } catch (ignoreClearB) {}

        try {
          if (ruleUc) sh.getRange(sr, 3).setDataValidation(ruleUc);
          else sh.getRange(sr, 3).clearDataValidations();
        } catch (eCell) {
          var msgCell = String(eCell.message || eCell);
          if (msgCell.indexOf('型付き') !== -1 || /typed column/i.test(msgCell)) {
            Logger.log('applyBucDetailStepValidations_: row ' + sr + ' col 3 — ' + msgCell);
          } else {
            throw eCell;
          }
        }

        if (rowAllEmpty) break;
      }
      r = sr;
    }
  } catch (e) {
    var msgDv = String(e.message || e);
    if (msgDv.indexOf('型付き') !== -1 || /typed column/i.test(msgDv)) {
      Logger.log('applyBucDetailStepValidations_: ' + msgDv);
      return;
    }
    throw e;
  }
}

/** アクター名・関連 BR／UC・外部 IF 連携先など、別シート参照の入力規則をまとめて付与 */
function applyAllReferenceValidations_(ss) {
  applyUcListActorValidation_(ss);
  applyBucRelatedBrValidation_(ss);
  applyBucDetailStepValidations_(ss);
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
    .setAllowInvalid(false)
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
    .setAllowInvalid(false)
    .build();

  var lr = ifSh.getLastRow();
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

/** 全シートに SpreadsheetApp の入力規則を付与 */
function applyRequirementDropdowns_(ss) {
  applyLegacyDropdowns_(ss);
  applyAllReferenceValidations_(ss);
}

/** 各データシートのドロップダウン入力規則を一括付与 */
function applyLegacyDropdowns_(ss) {
  var shBR = ss.getSheetByName('🎯 ビジネス要求');
  if (shBR) {
    var lrBR = shBR.getLastRow();
    for (var rBR = 2; rBR <= lrBR; rBR++) {
      setDropdown(shBR, rBR, 4, ['Must', 'Should', 'Could']);
      setDropdown(shBR, rBR, 7, ['草案', 'レビュー中', '合意済', '保留', '廃止']);
    }
  }

  var shFR = ss.getSheetByName('⚙️ 機能要求');
  if (shFR) {
    var lrFR = shFR.getLastRow();
    for (var rFR = 2; rFR <= lrFR; rFR++) {
      setDropdown(shFR, rFR, 8, ['Must', 'Should', 'Could']);
      setDropdown(shFR, rFR, 10, ['草案', 'レビュー中', '合意済', '差し戻し', '廃止']);
    }
  }

  var shNFR = ss.getSheetByName('🔒 非機能要求');
  if (shNFR) {
    var lrNFR = shNFR.getLastRow();
    for (var rNFR = 2; rNFR <= lrNFR; rNFR++) {
      setDropdown(shNFR, rNFR, 2, ['性能', '可用性', 'セキュリティ', '保守性', 'UX']);
      setDropdown(shNFR, rNFR, 8, ['草案', 'レビュー中', '合意済', '差し戻し', '廃止']);
    }
  }

  var shCON = ss.getSheetByName('🚧 制約条件');
  if (shCON) {
    var lrCON = shCON.getLastRow();
    for (var rCON = 2; rCON <= lrCON; rCON++) {
      setDropdown(shCON, rCON, 2, ['技術', 'ビジネス', '法規制', '運用']);
      setDropdown(shCON, rCON, 6, ['草案', '合意済', '廃止']);
    }
  }

  var shIF = ss.getSheetByName('🔗 外部IF');
  if (shIF) {
    var lrIF = shIF.getLastRow();
    for (var rIF = 2; rIF <= lrIF; rIF++) {
      setDropdown(shIF, rIF, 3, ['IN（受信）', 'OUT（送信）', '双方向']);
    }
  }

  var shOI = ss.getSheetByName('❓ 未解決事項');
  if (shOI) {
    var lrOI = shOI.getLastRow();
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

/**
 * ステータス列の文字色のみを条件付き書式で付与する（セル背景は付けない）。
 * desiredRows はシートの実際の行数（getMaxRows）まで自動的に切り詰める。
 */
function addStatusFormatting(sheet, col, desiredRows) {
  var rows = Math.min(desiredRows, sheet.getMaxRows() - 1);
  if (rows < 1) return;
  const range = sheet.getRange(2, col, rows, 1);
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

/** ステータス列に条件付き書式（文字色）を付与 */
function applyStatusFormattingAfterTables_(ss) {
  var sh;
  sh = ss.getSheetByName('🎯 ビジネス要求');
  if (sh) addStatusFormatting(sh, 7, VALIDATION_ROW_HEADROOM);
  sh = ss.getSheetByName('⚙️ 機能要求');
  if (sh) addStatusFormatting(sh, 10, VALIDATION_ROW_HEADROOM);
  sh = ss.getSheetByName('🔒 非機能要求');
  if (sh) addStatusFormatting(sh, 8, VALIDATION_ROW_HEADROOM);
  sh = ss.getSheetByName('❓ 未解決事項');
  if (sh) addStatusFormatting(sh, 7, VALIDATION_ROW_HEADROOM);
  sh = ss.getSheetByName(UC_LIST_SHEET_NAME);
  if (sh) addStatusFormatting(sh, 5, VALIDATION_ROW_HEADROOM);
}
