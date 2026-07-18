import { describe, expect, it } from 'vitest';
import { loadGasContext } from './support/gasSandbox.js';

describe('formatRequirementId', () => {
  const gas = loadGasContext();

  it('zero-pads to 3 digits', () => {
    expect(gas.formatRequirementId('BR', 1)).toBe('BR-001');
    expect(gas.formatRequirementId('UC', 42)).toBe('UC-042');
  });

  it('does not truncate numbers with 4+ digits', () => {
    expect(gas.formatRequirementId('FR', 1000)).toBe('FR-1000');
  });

  it('throws for an unknown counter key', () => {
    expect(() => gas.formatRequirementId('XX', 1)).toThrow();
  });

  it('throws for a non-positive sequence number', () => {
    expect(() => gas.formatRequirementId('BR', 0)).toThrow();
    expect(() => gas.formatRequirementId('BR', -1)).toThrow();
  });
});
