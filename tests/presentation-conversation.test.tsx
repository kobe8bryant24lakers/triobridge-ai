import { act, fireEvent, render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { DeckGenerator } from '../src/components/DeckGenerator';
import { PresentationEdit } from '../src/components/PresentationEdit';
import { DataChart } from '../src/components/DataChart';
import type { PresentationRuntime } from '../src/components/presentation-runtime';
import type { ChatMessage } from '../src/core/types';
import { createDemoPresentation } from '../src/office/presentation-demo';
import { setUiLanguage } from '../src/i18n';

function fixture() {
  const access = createDemoPresentation(vi.fn());
  const apply = vi.spyOn(access, 'apply');
  const insert = vi.spyOn(access, 'insert');
  const errors = vi.fn();
  const controllers: AbortController[] = [];
  const request = vi.fn<PresentationRuntime['request']>();
  const props: PresentationRuntime = {
    access,
    context: {
      materials: [{ id: 'm1', name: '材料.txt', text: '汇报事实', kind: 'text', tables: [] }],
      includeImages: false,
    },
    busy: false,
    demo: false,
    modelStamp: 'model-1',
    run: async (_label, work) => {
      const controller = new AbortController();
      controllers.push(controller);
      try {
        await work(controller.signal);
      } catch (error) {
        errors(error);
      }
    },
    request,
    notify: vi.fn(),
  };
  return { props, access, apply, insert, errors, controllers, request };
}

function deck(title: string) {
  return JSON.stringify({
    title,
    slides: [
      { title: '开场', layout: 'cover', subtitle: '说明目标', notes: '', sources: [] },
      { title: '结语', layout: 'closing', subtitle: '下一步', notes: '', sources: [] },
    ],
  });
}

function content(messages: ChatMessage[], index: number) {
  return JSON.parse(messages[index].content as string);
}

async function followUp(value: string) {
  fireEvent.change(screen.getByLabelText('继续调整'), { target: { value } });
  await userEvent.click(screen.getByRole('button', { name: '发送修改要求' }));
}

async function startDeck() {
  fireEvent.change(screen.getByLabelText('演示要求'), { target: { value: '整理管理层季度汇报' } });
  fireEvent.change(screen.getByLabelText('生成页数'), { target: { value: '2' } });
  await userEvent.click(screen.getByRole('button', { name: '生成整套 PPT' }));
  await screen.findByRole('region', { name: '整稿预览' });
}

beforeEach(() => setUiLanguage('zh-CN'));

describe('PowerPoint multi-turn editing', () => {
  it('revises a deck through multiple turns with manual changes and starts fresh from the preview after clearing', async () => {
    const f = fixture();
    f.request
      .mockResolvedValueOnce(deck('初稿'))
      .mockResolvedValueOnce(deck('第二稿'))
      .mockResolvedValueOnce(deck('第三稿'))
      .mockResolvedValueOnce(deck('第四稿'));
    render(<DeckGenerator {...f.props} />);
    await startDeck();
    fireEvent.change(screen.getByLabelText('文件标题'), { target: { value: '手动命名的初稿' } });
    await followUp('强调结论');
    expect(screen.getByLabelText('文件标题')).toHaveValue('第二稿');
    const second = f.request.mock.calls[1][0];
    expect(second.map((m) => m.role)).toEqual(['system', 'user', 'assistant', 'user']);
    expect(content(second, 2).title).toBe('手动命名的初稿');
    await followUp('再简化结尾');
    const third = f.request.mock.calls[2][0];
    expect(third.map((m) => m.role)).toEqual([
      'system',
      'user',
      'assistant',
      'user',
      'assistant',
      'user',
    ]);
    expect(content(third, 2).title).toBe('手动命名的初稿');
    expect(third[3].content).toBe('强调结论');
    expect(content(third, 4).title).toBe('第二稿');
    expect(third[5].content).toBe('再简化结尾');
    expect(screen.getByText('查看 3 轮编辑记录')).toBeInTheDocument();
    await userEvent.click(screen.getByRole('button', { name: '清空编辑对话' }));
    expect(screen.getByLabelText('文件标题')).toHaveValue('第三稿');
    await followUp('换一个更清晰的标题');
    const fourth = f.request.mock.calls[3][0];
    expect(fourth).toHaveLength(4);
    expect(content(fourth, 2).title).toBe('第三稿');
    expect(JSON.stringify(fourth)).not.toContain('强调结论');
    expect(f.apply).not.toHaveBeenCalled();
    expect(f.insert).not.toHaveBeenCalled();
  });

  it('keeps the deck, history and follow-up on invalid output, failure or cancellation', async () => {
    const f = fixture();
    f.request
      .mockResolvedValueOnce(deck('原预览'))
      .mockResolvedValueOnce('{"title":"坏结果","slides":[]}')
      .mockRejectedValueOnce(new Error('network failed'));
    render(<DeckGenerator {...f.props} />);
    await startDeck();
    await followUp('完善结尾');
    expect(f.errors).toHaveBeenCalledTimes(1);
    expect(screen.getByLabelText('文件标题')).toHaveValue('原预览');
    expect(screen.getByLabelText('继续调整')).toHaveValue('完善结尾');
    await userEvent.click(screen.getByRole('button', { name: '发送修改要求' }));
    expect(f.errors).toHaveBeenCalledTimes(2);
    let finish!: (value: string) => void;
    f.request.mockImplementationOnce(
      () =>
        new Promise((resolve) => {
          finish = resolve;
        }),
    );
    await userEvent.click(screen.getByRole('button', { name: '发送修改要求' }));
    await act(async () => {
      f.controllers[f.controllers.length - 1].abort();
      finish(deck('不应替换'));
    });
    expect(screen.getByLabelText('文件标题')).toHaveValue('原预览');
    expect(screen.getByLabelText('继续调整')).toHaveValue('完善结尾');
    expect(screen.getByText('查看 1 轮编辑记录')).toBeInTheDocument();
    f.request.mockResolvedValueOnce(deck('重试成功'));
    await userEvent.click(screen.getByRole('button', { name: '发送修改要求' }));
    expect(screen.getByLabelText('文件标题')).toHaveValue('重试成功');
    expect(f.request.mock.calls[4][0]).toHaveLength(4);
    expect(f.insert).not.toHaveBeenCalled();
  });

  it('blocks old deck context when source tables change and resets it on regeneration', async () => {
    const f = fixture();
    f.request.mockResolvedValueOnce(deck('原预览')).mockResolvedValueOnce(deck('新资料预览'));
    const { rerender } = render(<DeckGenerator {...f.props} />);
    await startDeck();
    fireEvent.change(screen.getByLabelText('继续调整'), { target: { value: '继续改' } });
    const context = {
      ...f.props.context,
      materials: f.props.context.materials.map((m) => ({
        ...m,
        tables: [{ id: 'new-table', name: '数据', headers: ['年份', '营收'], rows: [[2026, 23]] }],
      })),
    };
    rerender(<DeckGenerator {...f.props} context={context} />);
    expect(screen.getByRole('button', { name: '发送修改要求' })).toBeDisabled();
    expect(f.request).toHaveBeenCalledTimes(1);
    await userEvent.click(screen.getByRole('button', { name: '重新生成整套 PPT' }));
    expect(screen.getByLabelText('文件标题')).toHaveValue('新资料预览');
    expect(f.request.mock.calls[1][0]).toHaveLength(2);
    expect(screen.getByLabelText('继续调整')).toHaveValue('');
  });

  it('keeps slide proposals cumulative against the original snapshot and respects deselected changes', async () => {
    const f = fixture();
    f.request
      .mockResolvedValueOnce(
        JSON.stringify({
          summary: '初稿方案',
          changes: [
            { slideId: 'demo-1', shapeId: 'title', text: '新标题' },
            { slideId: 'demo-1', shapeId: 'body-1', text: '不需要的正文' },
          ],
        }),
      )
      .mockResolvedValueOnce(
        JSON.stringify({
          summary: '第二轮方案',
          changes: [{ slideId: 'demo-1', shapeId: 'title', text: '新标题', bold: false }],
        }),
      )
      .mockResolvedValueOnce(
        JSON.stringify({
          summary: '第三轮方案',
          changes: [
            { slideId: 'demo-1', shapeId: 'title', text: '新标题', bold: false, fontSize: 36 },
          ],
        }),
      );
    render(<PresentationEdit {...f.props} />);
    await userEvent.click(screen.getByRole('button', { name: '读取页面与对象' }));
    fireEvent.change(screen.getByLabelText('编辑要求'), { target: { value: '修改标题与正文' } });
    await userEvent.click(screen.getByRole('button', { name: '生成局部修改方案' }));
    await screen.findByText('初稿方案');
    await userEvent.click(screen.getByRole('checkbox', { name: /第 1 页 · 要点 1/ }));
    await followUp('保留选中的修改，标题取消加粗');
    const second = f.request.mock.calls[1][0];
    expect(content(second, 2).changes).toEqual([
      { slideId: 'demo-1', shapeId: 'title', text: '新标题' },
    ]);
    expect(content(second, 1).request.presentation.slides[0].shapes[0].text).toBe('季度业务回顾');
    await followUp('标题再放大');
    const third = f.request.mock.calls[2][0];
    expect(content(third, 4).changes[0]).toMatchObject({ text: '新标题', bold: false });
    expect(content(third, 1).request.presentation.slides[0].shapes[0].text).toBe('季度业务回顾');
    expect(f.apply).not.toHaveBeenCalled();
    await userEvent.click(screen.getByRole('button', { name: '应用所选 1 项改动' }));
    await userEvent.click(screen.getByRole('button', { name: '确认应用改动' }));
    await waitFor(() => expect(f.apply).toHaveBeenCalledTimes(1));
    fireEvent.change(screen.getByLabelText('继续调整'), { target: { value: '继续修改' } });
    expect(screen.getByRole('button', { name: '发送修改要求' })).toBeDisabled();
    await userEvent.click(screen.getByRole('button', { name: '重新读取页面' }));
    expect(screen.queryByLabelText('继续调整')).not.toBeInTheDocument();
  });

  it('only sends explicitly read slides for a numbered request or mixed background instruction', async () => {
    const f = fixture();
    const read = vi.spyOn(f.access, 'read');
    f.request.mockResolvedValue(JSON.stringify({ summary: '修改选中页', changes: [] }));
    render(<PresentationEdit {...f.props} />);
    await userEvent.click(screen.getByRole('button', { name: '读取页面与对象' }));
    for (const brief of ['精简第 1 页的标题', '把第 2 页背景改为米白色，并改写整份演示稿']) {
      fireEvent.change(screen.getByLabelText('编辑要求'), { target: { value: brief } });
      await userEvent.click(screen.getByRole('button', { name: '生成局部修改方案' }));
    }
    expect(read.mock.calls).toEqual([['selected']]);
    expect(f.request).toHaveBeenCalledTimes(2);
    for (const [messages] of f.request.mock.calls) {
      const { presentation } = content(messages, 1).request;
      expect(presentation.slideIds).toEqual(['demo-1']);
      expect(presentation.slides).toHaveLength(1);
      expect(JSON.stringify(messages)).not.toContain('增长与挑战');
    }
  });

  it('keeps a cross-slide background preview local until the user explicitly reads the whole deck', async () => {
    const f = fixture();
    f.request.mockResolvedValue(JSON.stringify({ summary: '整稿调整', changes: [] }));
    render(<PresentationEdit {...f.props} />);
    await userEvent.click(screen.getByRole('button', { name: '读取页面与对象' }));
    fireEvent.change(screen.getByLabelText('编辑要求'), {
      target: { value: '把整份 PPT 背景改为米白色，但第 2 页改为深蓝色' },
    });
    await userEvent.click(screen.getByRole('button', { name: '生成局部修改方案' }));
    await followUp('把第 3 页背景改为深蓝色');
    expect(f.errors).not.toHaveBeenCalled();
    await followUp('精简每页的标题');
    expect(f.request).not.toHaveBeenCalled();
    expect(f.errors.mock.calls[0][0].message).toContain('重新读取页面');
    await userEvent.click(screen.getByRole('button', { name: '重新读取页面' }));
    expect(screen.getByLabelText('继续调整')).toHaveValue('精简每页的标题');
    await userEvent.click(screen.getByRole('button', { name: '发送修改要求' }));
    expect(f.request).toHaveBeenCalledTimes(1);
    expect(content(f.request.mock.calls[0][0], 1).request.presentation.slides).toHaveLength(3);
  });

  it('revises chart recommendations with manual settings, never alters data, and rejects invalid revisions atomically', async () => {
    const f = fixture();
    f.request
      .mockResolvedValueOnce(
        JSON.stringify({ type: 'column', title: '初稿标题', reason: '比较数据' }),
      )
      .mockResolvedValueOnce(
        JSON.stringify({ type: 'line', title: '趋势标题', reason: '展示变化' }),
      )
      .mockResolvedValueOnce(JSON.stringify({ type: 'bar', title: '分类标题', reason: '比较分类' }))
      .mockResolvedValueOnce(
        JSON.stringify({ type: 'pie', title: '错误标题', reason: '不能改变数据' }),
      );
    render(<DataChart {...f.props} />);
    fireEvent.change(screen.getByLabelText('含表头的数据'), {
      target: { value: '月份\t营收\n一月\t-1\n二月\t8' },
    });
    await userEvent.click(screen.getByRole('button', { name: '让模型推荐图表' }));
    fireEvent.change(screen.getByLabelText('图表标题'), { target: { value: '手动图表标题' } });
    fireEvent.change(screen.getByLabelText('单位'), { target: { value: '万元' } });
    await followUp('强调趋势');
    const second = f.request.mock.calls[1][0];
    expect(content(second, 2).title).toBe('手动图表标题');
    expect(content(second, 1).request.unit).toBe('万元');
    expect(content(second, 1).request.data.series[0].values).toEqual([-1, 8]);
    await followUp('改为横向比较');
    expect(f.request.mock.calls[2][0]).toHaveLength(6);
    expect(content(f.request.mock.calls[2][0], 4).type).toBe('line');
    await followUp('试试饼图');
    expect(f.errors).toHaveBeenCalledTimes(1);
    expect(screen.getByLabelText('图表类型')).toHaveValue('bar');
    expect(screen.getByLabelText('图表标题')).toHaveValue('分类标题');
    expect(screen.getByLabelText('继续调整')).toHaveValue('试试饼图');
    expect(screen.getByLabelText('含表头的数据')).toHaveValue('月份\t营收\n一月\t-1\n二月\t8');
    expect(f.insert).not.toHaveBeenCalled();
    fireEvent.change(screen.getByLabelText('含表头的数据'), {
      target: { value: '月份\t营收\n一月\t3\n二月\t8' },
    });
    expect(screen.getByRole('button', { name: '发送修改要求' })).toBeDisabled();
  });
});
