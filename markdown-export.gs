/** ブック全体を Markdown へ書き出す処理（タブ順の走査・各シートのテーブル整形・エスケープ）。 */

/** ブックのタブ順（getSheets）に従い、非表示・🔢 ID管理 はスキップする。 */
function isSheetSkippedInMarkdownExport_(sh) {
  if (!sh) return true;
  if (sh.getName() === ID_SHEET_NAME) return true;
  try {
    if (sh.isSheetHidden && sh.isSheetHidden()) return true;
  } catch (ignore) {}
  return false;
}

/** 📋 概要シートを Markdown の「## 📋 概要」ブロックへ。 */
function markdownOverviewSection_(overviewSheet) {
  if (!overviewSheet) return '';
  var out = '## 📋 概要\n\n';
  out += '### ドキュメント管理\n';
  out += flattenOverviewDocManagementTable(overviewSheet) + '\n';
  out += '### プロジェクト概要\n';
  out += '- **概要:** ' + escapeMarkdown(overviewSheet.getRange('B10').getValue()) + '\n';
  out += '- **目的:** ' + escapeMarkdown(overviewSheet.getRange('B11').getValue()) + '\n';
  out += '- **現状（As-Is）:** ' + escapeMarkdown(overviewSheet.getRange('B12').getValue()) + '\n';
  out += '- **課題:** ' + escapeMarkdown(overviewSheet.getRange('B13').getValue()) + '\n\n';
  out += '### スコープ（IN）\n\n';
  out += overviewScopeBulletBlock(overviewSheet, 15, 17);
  out += '\n### スコープ（OUT）\n\n';
  out += overviewScopeBulletBlock(overviewSheet, 19, 21);
  out += '\n### 成功指標\n';
  out += extractTableAsMarkdown(overviewSheet, 23, 1, 4) + '\n\n';
  return out;
}

function exportRequirementsToMarkdown() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var md = '# 要求仕様書\n\n';
    var actorMap = readActorMap_(ss);
    var actorNameToId = readActorNameToIdMap_(ss);
    var bucSectionOpened = false;
    var ucSectionOpened = false;

    function ensureBucSectionHeader_() {
      if (bucSectionOpened) return '';
      bucSectionOpened = true;
      return '## BUC\n\n';
    }
    function ensureUcSectionHeader_() {
      if (ucSectionOpened) return '';
      ucSectionOpened = true;
      return '## 📖 ユースケース\n\n';
    }

    var sheets = ss.getSheets();
    var si;
    for (si = 0; si < sheets.length; si++) {
      var sh = sheets[si];
      if (isSheetSkippedInMarkdownExport_(sh)) continue;
      var name = sh.getName();

      if (name === '📋 概要') {
        md += markdownOverviewSection_(sh);
        continue;
      }
      if (name === '📌 前提条件') {
        md += '## 📌 前提条件\n\n' + extractTableAsMarkdown(sh, 1, 1, 3) + '\n\n';
        continue;
      }
      if (name === '👤 アクター') {
        md += '## 👤 アクター\n\n' + extractTableAsMarkdown(sh, 1, 1, 5) + '\n\n';
        continue;
      }
      if (name === '🎯 ビジネス要求') {
        md += '## 🎯 ビジネス要求\n\n' + extractTableAsMarkdown(sh, 1, 1, 7) + '\n\n';
        continue;
      }
      if (name === BUC_SHEET_NAME) {
        md += ensureBucSectionHeader_();
        if (sh.getLastRow() >= 1) {
          md += '### 一覧\n\n' + extractTableAsMarkdown(sh, 1, 1, 5) + '\n\n';
        }
        continue;
      }
      if (name === BUC_DETAIL_SHEET_NAME) {
        md += ensureBucSectionHeader_();
        if (sh.getLastRow() > 0) {
          md += parseBucDetailSheet_(sh);
        }
        continue;
      }
      if (name === UC_LIST_SHEET_NAME) {
        md += ensureUcSectionHeader_();
        if (sh.getLastRow() >= 1) {
          md +=
            '### ▼ ユースケース一覧\n\n' +
            extractUcListTableAsMarkdown_(sh, actorMap, actorNameToId) +
            '\n\n';
        }
        continue;
      }
      if (name === UC_DETAIL_SHEET_NAME) {
        md += ensureUcSectionHeader_();
        if (sh.getLastRow() > 0) {
          md += parseUseCaseDetailSheet_(sh, actorMap, actorNameToId);
        }
        continue;
      }
      if (name === '⚙️ 機能要求') {
        md += '## ' + name + '\n\n' + extractTableAsMarkdown(sh, 1, 1, 11) + '\n\n';
        continue;
      }
      if (name === '🔒 非機能要求') {
        md += '## ' + name + '\n\n' + extractTableAsMarkdown(sh, 1, 1, 8) + '\n\n';
        continue;
      }
      if (name === '🚧 制約条件') {
        md += '## ' + name + '\n\n' + extractTableAsMarkdown(sh, 1, 1, 6) + '\n\n';
        continue;
      }
      if (name === '🔗 外部IF') {
        md += '## ' + name + '\n\n' + extractExternalIfTableAsMarkdown_(sh, actorMap, actorNameToId) + '\n\n';
        continue;
      }
      if (name === '❓ 未解決事項') {
        md += '## ' + name + '\n\n' + extractTableAsMarkdown(sh, 1, 1, 7) + '\n\n';
        continue;
      }
      if (name === '📚 用語集') {
        md += '## ' + name + '\n\n' + extractTableAsMarkdown(sh, 1, 1, 4) + '\n\n';
        continue;
      }
      if (name === '✅ 変更履歴') {
        md += '## ' + name + '\n\n' + extractTableAsMarkdown(sh, 1, 1, 5) + '\n\n';
        continue;
      }
    }

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

/** 📙 BUC詳細を Markdown 表にする（▼ BUC-nnn 単位・手順表は 3 列に正規化）。 */
function parseBucDetailSheet_(sheet) {
  var mdD = '';
  var lastRd = sheet.getLastRow();
  var lastCd = sheet.getLastColumn();
  if (lastRd === 0) return mdD;

  var datD = sheet.getRange(1, 1, lastRd, Math.max(lastCd, 4)).getValues();
  var ri = 0;
  while (ri < datD.length) {
    var cellAd = String(datD[ri][0]).trim();
    if (/^▼\s*BUC-/.test(cellAd)) {
      mdD += '### ' + cellAd + '\n\n';
      ri++;
      if (ri >= datD.length) break;
      var hdrA = String(datD[ri][0]).trim();
      var hdrB = String(datD[ri][1] != null ? datD[ri][1] : '').trim();
      var legacyTable = /^手順$/.test(hdrA) && hdrB === 'アクター';
      if (/^手順$/.test(hdrA)) {
        ri++;
      }
      var tableD = [['手順', '行動内容', '関連UC']];
      while (ri < datD.length) {
        var ra = String(datD[ri][0]).trim();
        var rb = String(datD[ri][1] != null ? datD[ri][1] : '').trim();
        var rc = String(datD[ri][2] != null ? datD[ri][2] : '').trim();
        if (ra.substring(0, 1) === '▼') break;

        var actionCell;
        var ucCell;
        if (legacyTable) {
          var subjL = rb;
          var bodyL = rc;
          if (subjL && bodyL) actionCell = subjL + 'が' + bodyL;
          else actionCell = subjL || bodyL;
          ucCell = datD[ri][3];
        } else {
          actionCell = datD[ri][1];
          ucCell = datD[ri][2];
        }

        if (
          ra === '' &&
          String(actionCell != null ? actionCell : '').trim() === '' &&
          String(ucCell != null ? ucCell : '').trim() === ''
        ) {
          break;
        }

        tableD.push([datD[ri][0], actionCell, ucCell]);
        ri++;
      }
      if (tableD.length > 1) {
        mdD += arrayToMarkdownTable(tableD) + '\n\n';
      }
      continue;
    }
    ri++;
  }
  return mdD;
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
        if (nextCellA === '基本フロー' || nextCellA === '代替フロー' || nextCellA.startsWith('▼')) {
          break;
        }
        if (nextCellA === '') {
          i++;
          continue;
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
      var lastFlowNo = '';
      while (i < data.length) {
        const nextCellA = String(data[i][0]).trim();
        const nextCellB = String(data[i][1]).trim();

        if (nextCellA.startsWith('▼') || nextCellA === '基本フロー' || nextCellA === '代替フロー') break;
        if (nextCellA === '' && nextCellB === '') break;

        var noOut = nextCellA;
        if (noOut === '' && nextCellB !== '') {
          noOut = lastFlowNo;
        } else if (noOut !== '') {
          lastFlowNo = noOut;
        }
        flowTable.push([noOut, data[i][1]]);
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
    'function showCopied() {' +
    'var msg = document.getElementById("msg");' +
    'msg.style.display = "inline";' +
    'setTimeout(function() { msg.style.display = "none"; }, 2000);' +
    '}' +
    'function copyWithExecCommand() {' +
    'textArea.select();' +
    'document.execCommand("copy");' +
    'showCopied();' +
    '}' +
    'function copyToClipboard() {' +
    // Apps Script dialogs run in a sandboxed iframe where the Clipboard API
    // may be unavailable or denied by permissions policy, so fall back to
    // the (deprecated but still broadly supported) execCommand approach.
    'if (navigator.clipboard && navigator.clipboard.writeText) {' +
    'navigator.clipboard.writeText(textArea.value).then(showCopied, copyWithExecCommand);' +
    '} else {' +
    'copyWithExecCommand();' +
    '}' +
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
