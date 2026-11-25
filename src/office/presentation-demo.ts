import type { PresentationAccess, PresentationSnapshot } from '../presentation/types';
import { assertSnapshotUnchanged, reorderedIds, validateEdit } from '../presentation/editing';

export function createDemoPresentation(onWrite: (text: string) => void): PresentationAccess {
  let snapshot: PresentationSnapshot = {
    scope: 'selected',
    slideIds: ['demo-1', 'demo-2', 'demo-3'],
    selectedShapeIds: ['body-1', 'body-2', 'body-3'],
    size: { width: 960, height: 540 },
    skipped: 0,
    slides: ['季度业务回顾', '增长与挑战', '下季度行动'].map((title, i) => ({
      id: `demo-${i + 1}`,
      index: i,
      title,
      shapes: [
        {
          id: 'title',
          name: '标题',
          type: 'TextBox',
          role: 'title',
          text: title,
          left: 40,
          top: 35,
          width: 880,
          height: 60,
          font: { name: 'Arial', size: 32, color: '#333333', bold: true },
          alignment: 'Left',
        },
        ...['客户响应时间从 24 小时缩短至 8 小时', '新版入门指南已上线', '优化跨团队协作流程'].map(
          (text, n) => ({
            id: `body-${n + 1}`,
            name: `要点 ${n + 1}`,
            type: 'TextBox',
            role: 'body' as const,
            text,
            left: 40 + n * 280,
            top: 150,
            width: 240,
            height: 180,
            font: {
              name: n === 1 ? 'Aptos' : 'Arial',
              size: n === 1 ? 18 : 20,
              color: '#444444',
              bold: false,
            },
            alignment: 'Left',
          }),
        ),
      ],
    })),
  };
  const read = async (scope: 'selected' | 'all') =>
    structuredClone({
      ...snapshot,
      scope,
      slides: scope === 'all' ? snapshot.slides : snapshot.slides.slice(0, 1),
    });
  return {
    capabilities: {
      read: true,
      selected: true,
      reorder: true,
      insert: true,
      create: false,
      thumbnail: false,
      backup: false,
    },
    read,
    prepareInsert: async () =>
      structuredClone({ slideIds: snapshot.slideIds, size: snapshot.size }),
    apply: async (before, input) => {
      assertSnapshotUnchanged(before, await read(before.scope));
      const plan = validateEdit(input, before);
      for (const p of plan.changes) {
        const shape = snapshot.slides
          .find((s) => s.id === p.slideId)!
          .shapes.find((s) => s.id === p.shapeId)!;
        for (const key of ['text', 'left', 'top', 'width', 'height', 'alignment'] as const)
          if (p[key] !== undefined) Object.assign(shape, { [key]: p[key] });
        if (p.fontName !== undefined) shape.font!.name = p.fontName;
        if (p.fontSize !== undefined) shape.font!.size = p.fontSize;
        if (p.fontColor !== undefined) shape.font!.color = p.fontColor;
        if (p.bold !== undefined) shape.font!.bold = p.bold;
      }
      if (plan.slideOrder) {
        snapshot.slideIds = reorderedIds(before, plan.slideOrder);
        snapshot.slides.sort(
          (a, b) => snapshot.slideIds.indexOf(a.id) - snapshot.slideIds.indexOf(b.id),
        );
        snapshot.slides.forEach((s, i) => (s.index = i));
      }
      onWrite(`演示编辑：${plan.summary}\n${JSON.stringify(plan, null, 2)}`);
    },
    insert: async (_base64, count) => {
      onWrite(`演示：已模拟在末尾插入 ${count} 页。下载 PPTX 可查看实际可编辑文件。`);
    },
    create: async () => {
      throw new Error('浏览器演示请下载 PPTX，在 PowerPoint 中打开。');
    },
    thumbnail: async () => '',
    backup: async () => {
      throw new Error('演示文档无需备份。');
    },
  };
}
