import { describe, expect, it } from 'vitest';
import { loadGasContext } from './support/gasSandbox.js';

describe('escapeMarkdown', () => {
  const gas = loadGasContext();

  it('escapes pipes and turns newlines into <br>', () => {
    expect(gas.escapeMarkdown('a|b\nc')).toBe('a\\|b<br>c');
  });

  it('formats Date values as yyyy-MM-dd', () => {
    // Built with the sandbox's own Date constructor: `instanceof Date` inside
    // create-spreadsheet.gs refers to the vm context's Date, which is a
    // different realm from a Date built in this test file.
    const date = new gas.Date(gas.Date.UTC(2026, 4, 10));
    expect(gas.escapeMarkdown(date)).toBe('2026-05-10');
  });

  it('stringifies non-string values', () => {
    expect(gas.escapeMarkdown(42)).toBe('42');
  });
});

describe('sectionHeader_', () => {
  const gas = loadGasContext();

  it('formats an H2 heading followed by a blank line', () => {
    expect(gas.sectionHeader_('📗 BUC')).toBe('## 📗 BUC\n\n');
  });
});

describe('arrayToMarkdownTable', () => {
  const gas = loadGasContext();

  it('builds a header/divider/body table', () => {
    const md = gas.arrayToMarkdownTable([
      ['ID', 'Name'],
      ['BR-001', '受注'],
    ]);
    expect(md).toBe('| ID | Name |\n| --- | --- |\n| BR-001 | 受注 |\n');
  });

  it('returns an empty string for empty data', () => {
    expect(gas.arrayToMarkdownTable([])).toBe('');
  });
});
