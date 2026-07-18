import { describe, expect, it } from 'vitest';
import { loadGasContext } from './support/gasSandbox.js';

/** Sheet stand-in whose only job is to record every getRange() call and answer getValues(). */
function makeIdColumnSheet(idCount) {
  const rows = [];
  for (let i = 1; i <= idCount; i++) rows.push([`BR-${String(i).padStart(3, '0')}`]);
  const calls = [];
  return {
    getLastRow: () => idCount + 1, // +1 for the header row
    getRange(row, col, numRows, numCols) {
      calls.push({ row, col, numRows, numCols });
      return { numRows, getValues: () => rows.slice(row - 2, row - 2 + numRows) };
    },
    _calls: calls,
  };
}

/** Sheet stand-in for getUcIdListRange_, which only needs getMaxRows/getLastRow/getRange. */
function makeRangeProbeSheet(maxRows, lastRow) {
  return {
    getMaxRows: () => maxRows,
    getLastRow: () => lastRow,
    getRange(row, col, numRows, numCols) {
      return { row, col, numRows, numCols };
    },
  };
}

describe('getFirstColumnIdRange_', () => {
  const gas = loadGasContext();

  it('covers every ID row even past the old 500-row cap', () => {
    const sheet = makeIdColumnSheet(600);
    const range = gas.getFirstColumnIdRange_(sheet);

    // The old Math.min(lastRow, 500) cap would have scanned only 499 data
    // rows and returned a 499-row range here instead of the full 600.
    expect(range.numRows).toBe(600);
    const finalCall = sheet._calls[sheet._calls.length - 1];
    expect(finalCall.numRows).toBe(600);
  });

  it('returns null when the sheet has no data rows', () => {
    const sheet = makeIdColumnSheet(0);
    expect(gas.getFirstColumnIdRange_(sheet)).toBeNull();
  });
});

describe('getUcIdListRange_', () => {
  const gas = loadGasContext();

  it('extends headroom past the old fixed 2000-row ceiling, bounded by the sheet\'s real max rows', () => {
    const ucSheet = makeRangeProbeSheet(5000, 2600);
    const ss = { getSheetByName: (name) => (name === gas.UC_LIST_SHEET_NAME ? ucSheet : null) };

    const range = gas.getUcIdListRange_(ss);

    // Old logic: maxEnd = min(getMaxRows(), 2000) = 2000, so anything beyond
    // row 2000 was silently excluded no matter how much real data existed.
    expect(range.numRows).toBe(2600 + gas.VALIDATION_ROW_HEADROOM - 1);
  });

  it('clamps the extended range to the sheet\'s actual max rows', () => {
    const ucSheet = makeRangeProbeSheet(5000, 4500);
    const ss = { getSheetByName: (name) => (name === gas.UC_LIST_SHEET_NAME ? ucSheet : null) };

    const range = gas.getUcIdListRange_(ss);

    expect(range.numRows).toBe(5000 - 2 + 1);
  });
});
