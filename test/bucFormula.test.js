import { describe, expect, it } from 'vitest';
import { loadGasContext } from './support/gasSandbox.js';

describe('bucBrMirrorFormula_', () => {
  const gas = loadGasContext();

  it('builds a VLOOKUP formula against the business-requirements sheet', () => {
    expect(gas.bucBrMirrorFormula_(3)).toBe(
      '=IF(D3="","",IFERROR(VLOOKUP(D3,\'🎯 ビジネス要求\'!$A:$B,2,FALSE),""))'
    );
  });
});
