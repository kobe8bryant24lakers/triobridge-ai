import { act, render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { DataChart } from '../src/components/DataChart';
import { DeckGenerator } from '../src/components/DeckGenerator';
import type { PresentationRuntime } from '../src/components/presentation-runtime';
import { setUiLanguage, useUiLanguage } from '../src/i18n';
import { LAYOUT_LABELS } from '../src/presentation/content';
import { PATCH_LABELS } from '../src/presentation/editing';

const sources: PresentationRuntime['context'] = {
  includeImages: false,
  materials: [
    { id: 'm1', name: '季度资料.md', kind: 'text', text: '第一场有 18 人。', tables: [] },
  ],
};

function runtime(request: PresentationRuntime['request']): PresentationRuntime {
  return {
    context: sources,
    busy: false,
    demo: false,
    modelStamp: 'test',
    run: async (_label, work) => work(new AbortController().signal),
    request,
    notify: vi.fn(),
  };
}

afterEach(() => act(() => setUiLanguage('zh-CN')));

describe('PowerPoint workbench language', () => {
  it('switches chart controls and validation without changing pasted data or a custom title', async () => {
    const user = userEvent.setup();
    const props = runtime(vi.fn());
    function Chart() {
      useUiLanguage();
      return <DataChart {...props} />;
    }
    render(<Chart />);
    await user.type(screen.getByLabelText('含表头的数据'), '场次,人数\n第一场,18');
    act(() => setUiLanguage('en'));
    expect(screen.getByLabelText('Data with headers')).toHaveValue('场次,人数\n第一场,18');
    expect(screen.getByLabelText('Chart title')).toHaveValue('Business comparison');
    expect(screen.getByRole('option', { name: 'Column' })).toBeInTheDocument();
    expect(screen.getByRole('img', { name: /Column chart/ })).toBeInTheDocument();
    await user.clear(screen.getByLabelText('Chart title'));
    await user.type(screen.getByLabelText('Chart title'), '读者自己的标题');
    await user.clear(screen.getByLabelText('Maximum rows'));
    expect(screen.getByText(/Enter a starting row from 1 to 1,000/)).toBeInTheDocument();
    act(() => setUiLanguage('zh-CN'));
    expect(screen.getByLabelText('图表标题')).toHaveValue('读者自己的标题');
    expect(screen.getByText(/起始行请输入 1–1,000/)).toBeInTheDocument();
  });

  it('preserves user sources and generated content while translating prompts, layout and edit labels', async () => {
    setUiLanguage('en');
    const user = userEvent.setup();
    const result = JSON.stringify({
      title: '工作坊回顾',
      slides: [
        {
          title: '开始',
          layout: 'cover',
          subtitle: '活动回顾',
          notes: '原文备注',
          sources: ['m1'],
        },
        { title: '结束', layout: 'closing', subtitle: '下一次安排', notes: '', sources: ['m1'] },
      ],
    });
    const request = vi.fn<PresentationRuntime['request']>(async () => result);
    render(<DeckGenerator {...runtime(request)} />);
    const brief = '请用中文写两页工作坊回顾。';
    await user.type(screen.getByLabelText('Presentation brief', { selector: 'textarea' }), brief);
    const count = screen.getByRole('spinbutton', { name: 'Number of slides' });
    await user.clear(count);
    await user.type(count, '2');
    await user.click(screen.getByRole('button', { name: 'Generate deck' }));
    expect(await screen.findByLabelText('File title')).toHaveValue('工作坊回顾');
    expect(screen.getByRole('heading', { name: '开始' })).toBeInTheDocument();
    const [system, input] = request.mock.calls[0][0];
    expect(String(system.content)).toContain('Use the requested output language');
    expect(String(input.content)).toContain(brief);
    expect(String(input.content)).toContain('季度资料.md');
    expect(String(input.content)).toContain('第一场有 18 人。');
    expect(LAYOUT_LABELS.cover).toBe('Cover');
    expect(PATCH_LABELS.fontSize).toBe('Font size');
    act(() => setUiLanguage('zh-CN'));
    expect(LAYOUT_LABELS.cover).toBe('封面');
    expect(PATCH_LABELS.fontSize).toBe('字号');
  });
});
