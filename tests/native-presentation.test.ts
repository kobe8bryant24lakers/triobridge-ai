import { describe, expect, it, vi } from 'vitest';
import { createNativeBridge } from '../src/offline/office';
import type { NativeWorkbench } from '../src/offline/transport';
import type { PresentationInsertTarget, PresentationSnapshot } from '../src/presentation/types';

describe('native PowerPoint insertion targets', () => {
  it('binds each prepared target to its own append token without reading the selection', async () => {
    const targets: PresentationInsertTarget[] = [
      { slideIds: [], size: { width: 960, height: 540 } },
      { slideIds: ['s1'], size: { width: 720, height: 540 } },
    ];
    let prepared = 0;
    const native: NativeWorkbench = {
      host: 'powerpoint',
      model: vi.fn(),
      cancel: vi.fn(),
      office: vi.fn(async (action) => {
        if (action === 'presentationTarget') {
          const index = prepared++;
          return { ok: true, value: { token: `insert-${index}`, target: targets[index] } };
        }
        if (action === 'append') return { ok: true };
        return { ok: false, message: '当前选区无法读取。' };
      }),
    };
    const access = createNativeBridge(native).presentation!;
    const first = await access.prepareInsert();
    const second = await access.prepareInsert();
    expect(first).toBe(targets[0]);
    expect(second).toBe(targets[1]);
    await access.insert('deck-one', 8, first);
    await access.insert('chart-two', 1, second);
    expect(vi.mocked(native.office).mock.calls).toEqual([
      ['presentationTarget', undefined],
      ['presentationTarget', undefined],
      ['append', { token: 'insert-0', base64: 'deck-one', count: 8 }],
      ['append', { token: 'insert-1', base64: 'chart-two', count: 1 }],
    ]);
  });

  it('keeps editing snapshot tokens separate from append targets', async () => {
    const snapshot: PresentationSnapshot = {
      scope: 'selected',
      slideIds: ['s1'],
      size: { width: 960, height: 540 },
      selectedShapeIds: [],
      skipped: 0,
      slides: [{ id: 's1', index: 0, title: '原页', shapes: [] }],
    };
    const target: PresentationInsertTarget = { slideIds: ['s1'], size: snapshot.size };
    const native: NativeWorkbench = {
      host: 'powerpoint',
      model: vi.fn(),
      cancel: vi.fn(),
      office: vi.fn(async (action) => ({
        ok: true,
        value:
          action === 'presentation'
            ? { token: 'edit-token', snapshot }
            : action === 'presentationTarget'
              ? { token: 'append-token', target }
              : undefined,
      })),
    };
    const access = createNativeBridge(native).presentation!;
    const editing = await access.read('selected');
    const insertion = await access.prepareInsert();
    await access.apply(editing, { summary: '保留现有内容', changes: [] });
    await access.insert('new-slides', 2, insertion);
    expect(native.office).toHaveBeenCalledWith('edit', {
      token: 'edit-token',
      proposal: { summary: '保留现有内容', changes: [] },
    });
    expect(native.office).toHaveBeenLastCalledWith('append', {
      token: 'append-token',
      base64: 'new-slides',
      count: 2,
    });
  });
});
