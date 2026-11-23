import { describe, expect, it } from 'vitest';
import { render, screen } from '@testing-library/react';
import { ReportPreview } from '../src/components/StructuredPreview';
import { parseReport } from '../src/core/structured-output';

describe('report list preview', () => {
  it.each([
    [['1. 补齐 12 份资料。', '2. 保留 <待补充>&复核'], 'P'],
    [['补齐资料', '复核'], 'LI'],
    [['1. 补齐资料', '未编号的复核'], 'LI'],
  ])('matches Word list presentation and keeps authored wording', (bullets, tag) => {
    const report = parseReport(
      JSON.stringify({
        title: '验收报告',
        sections: [{ heading: '后续行动', bullets }],
      }),
    );
    render(<ReportPreview report={report} />);
    const items = screen.getAllByRole('listitem');
    expect(items.map((item) => item.tagName)).toEqual(bullets.map(() => tag));
    expect(items.map((item) => item.textContent)).toEqual(bullets);
    expect(screen.getAllByRole('list')).toHaveLength(1);
  });
});
