import { describe, expect, it } from 'vitest';
import { loadGasContext } from './support/gasSandbox.js';

describe('extractActorIdFromCell_', () => {
  const gas = loadGasContext();

  it('extracts a leading ACT-nnn token', () => {
    expect(gas.extractActorIdFromCell_('ACT-001（一般ユーザー）')).toBe('ACT-001');
  });

  it('returns an empty string when there is no leading ID', () => {
    expect(gas.extractActorIdFromCell_('一般ユーザー')).toBe('');
  });
});

describe('resolveActorLabelForMarkdown_', () => {
  const gas = loadGasContext();
  const actorMap = { 'ACT-001': '一般ユーザー' };
  const actorNameToId = { 一般ユーザー: 'ACT-001' };

  it('resolves a bare actor name to "ID（Name）"', () => {
    expect(gas.resolveActorLabelForMarkdown_('一般ユーザー', actorMap, actorNameToId)).toBe(
      'ACT-001（一般ユーザー）'
    );
  });

  it('keeps an existing ACT-nnn prefix and fills in the name', () => {
    expect(gas.resolveActorLabelForMarkdown_('ACT-001', actorMap, actorNameToId)).toBe(
      'ACT-001（一般ユーザー）'
    );
  });

  it('passes through names with no known mapping unchanged', () => {
    expect(gas.resolveActorLabelForMarkdown_('未知の担当者', actorMap, actorNameToId)).toBe(
      '未知の担当者'
    );
  });

  it('returns an empty string for empty input', () => {
    expect(gas.resolveActorLabelForMarkdown_('', actorMap, actorNameToId)).toBe('');
  });
});
