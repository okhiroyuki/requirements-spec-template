/**
 * Minimal stand-in for a Google Sheets Sheet object, backed by a plain 2D
 * array. Supports only the subset of the Sheet API the parsers under test
 * call: getLastRow / getLastColumn / getRange(...).getValues().
 */
export function makeMockSheet(rows) {
  const numCols = rows.reduce((max, row) => Math.max(max, row.length), 0);

  return {
    getLastRow: () => rows.length,
    getLastColumn: () => numCols,
    getRange(row, col, numRows, numColsArg) {
      const values = rows.slice(row - 1, row - 1 + numRows).map((r) => {
        const out = [];
        for (let c = 0; c < numColsArg; c++) {
          const v = r[col - 1 + c];
          out.push(v !== undefined ? v : '');
        }
        return out;
      });
      return { getValues: () => values };
    },
  };
}
