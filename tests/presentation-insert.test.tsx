import { beforeEach, describe, expect, it, vi } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { DeckGenerator } from '../src/components/DeckGenerator';
import { DataChart } from '../src/components/DataChart';
import type { PresentationRuntime } from '../src/components/presentation-runtime';
import type { PresentationAccess, PresentationInsertTarget } from '../src/presentation/types';
import { renderDeck } from '../src/presentation/render';

vi.mock('../src/presentation/render', async (original) => ({
  ...(await original<typeof import('../src/presentation/render')>()),
  renderDeck: vi.fn().mockResolvedValue('rendered-pptx'),
}));

function runtime() {
  const target: PresentationInsertTarget = {
    slideIds: ['original-1', 'original-2'],
    size: { width: 720, height: 540 },
  };
  const access: PresentationAccess = {
    capabilities: {
      read: true,
      selected: true,
      reorder: true,
      insert: true,
      create: true,
      thumbnail: false,
      backup: false,
    },
    read: vi.fn().mockRejectedValue(new Error('请检查当前选区。')),
    prepareInsert: vi.fn().mockResolvedValue(target),
    insert: vi.fn().mockResolvedValue(undefined),
    apply: vi.fn(),
    create: vi.fn(),
    thumbnail: vi.fn(),
    backup: vi.fn(),
  };
  const errors = vi.fn();
  const props: PresentationRuntime = {
    context: {
      materials: [
        { id: 'source', name: '季度材料.txt', kind: 'text', text: '季度经营汇报', tables: [] },
      ],
      includeImages: false,
    },
    access,
    busy: false,
    demo: true,
    modelStamp: 'test-model',
    run: async (_label, work) => {
      try {
        await work(new AbortController().signal);
      } catch (error) {
        errors(error);
      }
    },
    request: vi.fn(async (_messages, _signal, demoResult) => demoResult),
    notify: vi.fn(),
  };
  return { props, access, target, errors };
}

beforeEach(() => vi.clearAllMocks());

describe('append generated PowerPoint pages independently of selection', () => {
  it('generates from materials and appends a deck at the target size without reading selected objects', async () => {
    const user = userEvent.setup();
    const f = runtime();
    render(<DeckGenerator {...f.props} />);
    await user.type(screen.getByLabelText('演示要求'), '根据素材生成季度汇报');
    await user.click(screen.getByRole('button', { name: '生成整套 PPT' }));
    await screen.findByRole('region', { name: '整稿预览' });
    expect(f.access.prepareInsert).not.toHaveBeenCalled();
    await user.click(screen.getByRole('button', { name: '插入当前演示稿末尾' }));
    await user.click(screen.getByRole('button', { name: '确认插入 8 页' }));
    await waitFor(() => expect(f.access.insert).toHaveBeenCalledTimes(1));
    expect(f.access.read).not.toHaveBeenCalled();
    expect(f.access.prepareInsert).toHaveBeenCalledTimes(1);
    expect(renderDeck).toHaveBeenCalledWith(
      expect.objectContaining({ slides: expect.any(Array) }),
      f.props.context,
      expect.any(Object),
      f.target.size,
    );
    expect(f.access.insert).toHaveBeenCalledWith('rendered-pptx', 8, f.target);
    expect(vi.mocked(f.access.insert).mock.calls[0][2]).toBe(f.target);
    expect(f.errors).not.toHaveBeenCalled();
    expect(f.props.notify).toHaveBeenLastCalledWith(
      expect.stringContaining('已在当前演示稿末尾插入 8 页'),
    );
  });

  it('appends an editable chart without requiring a slide or text selection', async () => {
    const user = userEvent.setup();
    const f = runtime();
    render(<DataChart {...f.props} />);
    await user.type(screen.getByLabelText('含表头的数据'), '月份\t营收\n一月\t128\n二月\t156');
    await user.click(screen.getByRole('button', { name: '插入可编辑图表' }));
    await user.click(screen.getByRole('button', { name: '确认插入图表页' }));
    await waitFor(() => expect(f.access.insert).toHaveBeenCalledTimes(1));
    expect(f.access.read).not.toHaveBeenCalled();
    expect(f.access.prepareInsert).toHaveBeenCalledTimes(1);
    expect(vi.mocked(renderDeck).mock.calls[0][3]).toEqual(f.target.size);
    expect(f.access.insert).toHaveBeenCalledWith('rendered-pptx', 1, f.target);
    expect(vi.mocked(f.access.insert).mock.calls[0][2]).toBe(f.target);
    expect(f.errors).not.toHaveBeenCalled();
    expect(f.props.notify).toHaveBeenLastCalledWith(expect.stringContaining('已在末尾插入 1 页'));
  });

  it('does not render or insert when the target size cannot be read', async () => {
    const user = userEvent.setup();
    const f = runtime();
    f.target.size = null;
    render(<DeckGenerator {...f.props} />);
    await user.type(screen.getByLabelText('演示要求'), '季度汇报');
    await user.click(screen.getByRole('button', { name: '生成整套 PPT' }));
    await user.click(await screen.findByRole('button', { name: '插入当前演示稿末尾' }));
    await user.click(screen.getByRole('button', { name: '确认插入 8 页' }));
    await waitFor(() => expect(f.errors).toHaveBeenCalledTimes(1));
    expect(f.errors.mock.calls[0][0].message).toContain('无法确认当前页面尺寸');
    expect(renderDeck).not.toHaveBeenCalled();
    expect(f.access.insert).not.toHaveBeenCalled();
    expect(screen.getByRole('region', { name: '整稿预览' })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: '确认插入 8 页' })).toBeEnabled();
  });
});
