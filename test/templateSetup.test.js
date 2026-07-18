import { describe, expect, it } from 'vitest';
import { loadGasContext } from './support/gasSandbox.js';

function makeMockSpreadsheet(existingSheetNames) {
  return {
    getSheetByName: (name) => (existingSheetNames.includes(name) ? { name } : null),
  };
}

describe('bookHasExistingTemplateData_', () => {
  const gas = loadGasContext();

  it('returns false when none of the template tabs exist yet', () => {
    const ss = makeMockSpreadsheet([]);
    expect(gas.bookHasExistingTemplateData_(ss)).toBe(false);
  });

  it('returns true when any template tab already exists', () => {
    const ss = makeMockSpreadsheet(['📋 概要']);
    expect(gas.bookHasExistingTemplateData_(ss)).toBe(true);
  });

  it('checks every template tab name, not just the first', () => {
    const lastName = gas.TEMPLATE_SHEET_NAMES[gas.TEMPLATE_SHEET_NAMES.length - 1];
    const ss = makeMockSpreadsheet([lastName]);
    expect(gas.bookHasExistingTemplateData_(ss)).toBe(true);
  });
});
