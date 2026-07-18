import { describe, expect, it } from 'vitest';
import { loadGasContext } from './support/gasSandbox.js';
import { makeMockSheet } from './support/mockSheet.js';

describe('parseBucDetailSheet_', () => {
  const gas = loadGasContext();

  it('renders a ▼ BUC block as a heading plus a 3-column table', () => {
    const sheet = makeMockSheet([
      ['▼ BUC-001: 受注登録・検証業務'],
      ['手順', '行動内容', '関連UC'],
      ['1', '顧客が注文書を送付する', ''],
      ['2', '一般ユーザーが注文内容をシステムに入力する', 'UC-001'],
    ]);

    const md = gas.parseBucDetailSheet_(sheet);

    expect(md).toContain('### ▼ BUC-001: 受注登録・検証業務');
    expect(md).toContain('| 手順 | 行動内容 | 関連UC |');
    expect(md).toContain('| 2 | 一般ユーザーが注文内容をシステムに入力する | UC-001 |');
  });

  it('ends a block at the next ▼ heading, skipping a blank separator row', () => {
    const sheet = makeMockSheet([
      ['▼ BUC-001: A業務'],
      ['手順', '行動内容', '関連UC'],
      ['1', 'ステップ1', ''],
      ['', '', ''],
      ['▼ BUC-002: B業務'],
      ['手順', '行動内容', '関連UC'],
      ['1', 'ステップ2', ''],
    ]);

    const md = gas.parseBucDetailSheet_(sheet);
    const blocks = md.split('### ▼ ').filter(Boolean);

    expect(blocks).toHaveLength(2);
    expect(md).toContain('BUC-001: A業務');
    expect(md).toContain('BUC-002: B業務');
  });
});

describe('parseUseCaseDetailSheet_', () => {
  const gas = loadGasContext();
  const actorMap = { 'ACT-001': '一般ユーザー' };
  const actorNameToId = { 一般ユーザー: 'ACT-001' };

  it('renders the meta table, basic flow, and alternate flow', () => {
    const sheet = makeMockSheet([
      ['▼ UC-001: 受注データを登録する'],
      ['アクター', '一般ユーザー'],
      ['事前条件', 'ログイン済であること'],
      ['基本フロー'],
      ['1', '画面を開く'],
      ['2', 'フォームを表示する'],
      ['代替フロー'],
      ['1a', '入力エラー時は差し戻す'],
    ]);

    const md = gas.parseUseCaseDetailSheet_(sheet, actorMap, actorNameToId);

    expect(md).toContain('### ▼ UC-001: 受注データを登録する');
    expect(md).toContain('| アクター | ACT-001（一般ユーザー） |');
    expect(md).toContain('#### 基本フロー');
    expect(md).toContain('| 2 | フォームを表示する |');
    expect(md).toContain('#### 代替フロー');
    expect(md).toContain('| 1a | 入力エラー時は差し戻す |');
  });
});
