import { t } from '../i18n';
import type { PresentationAccess, PresentationSnapshot } from '../presentation/types';
import { assertSnapshotUnchanged, reorderedIds, validateEdit } from '../presentation/editing';

export function createDemoPresentation(onWrite: (text: string) => void): PresentationAccess {
  const initialSnapshot = (): PresentationSnapshot => ({
    scope: 'selected',
    slideIds: ['demo-1', 'demo-2', 'demo-3'],
    selectedShapeIds: ['body-1', 'body-2', 'body-3'],
    size: { width: 960, height: 540 },
    skipped: 0,
    slides: [
      t('季度业务回顾', 'Quarterly business review'),
      t('增长与挑战', 'Growth and challenges'),
      t('下季度行动', "Next quarter's actions"),
    ].map((title, i) => ({
      id: `demo-${i + 1}`,
      index: i,
      title,
      background: { kind: 'master' },
      shapes: [
        {
          id: 'title',
          name: t('标题', 'Title'),
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
        ...[
          t(
            '客户响应时间从 24 小时缩短至 8 小时',
            'Customer response time fell from 24 hours to 8 hours',
          ),
          t('新版入门指南已上线', 'The new onboarding guide is live'),
          t('优化跨团队协作流程', 'Improve cross-team handovers'),
        ].map((text, n) => ({
          id: `body-${n + 1}`,
          name: t(`要点 ${n + 1}`, `Bullet ${n + 1}`),
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
        })),
      ],
    })),
  });
  let stored: PresentationSnapshot | undefined;
  const current = () => (stored ??= initialSnapshot());
  const read = async (scope: 'selected' | 'all') =>
    structuredClone({
      ...current(),
      scope,
      slides: scope === 'all' ? current().slides : current().slides.slice(0, 1),
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
      backgrounds: true,
    },
    read,
    prepareInsert: async () =>
      structuredClone({ slideIds: current().slideIds, size: current().size }),
    apply: async (before, input) => {
      const snapshot = current();
      assertSnapshotUnchanged(before, await read(before.scope));
      const plan = validateEdit(input, before);
      for (const a of plan.actions || []) {
        if (a.kind !== 'background')
          throw Error(t('请打开 PPTX 工作副本。', 'Open a PPTX working copy.'));
        const source =
          a.fill.kind === 'reference'
            ? before.slides.find((s) => s.id === (a.fill as { sourceId: string }).sourceId)
                ?.background
            : a.fill;
        snapshot.slides.find((s) => s.id === a.slideId)!.background = structuredClone(source);
      }
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
      onWrite(
        t(
          `演示编辑：${plan.summary}\n${JSON.stringify(plan, null, 2)}`,
          `Demo edit: ${plan.summary}
${JSON.stringify(plan, null, 2)}`,
        ),
      );
    },
    insert: async (_base64, count) => {
      onWrite(
        t(
          `演示：已模拟在末尾插入 ${count} 页。下载 PPTX 可查看实际可编辑文件。`,
          `Demo: ${count} slides appended to the simulated presentation. Download the PPTX to see the editable file.`,
        ),
      );
    },
    create: async () => {
      throw new Error(
        t(
          '浏览器演示请下载 PPTX，在 PowerPoint 中打开。',
          'Download the PPTX from the browser demo and open it in PowerPoint.',
        ),
      );
    },
    thumbnail: async () => '',
    backup: async () => {
      throw new Error(t('演示文档无需备份。', 'Demo presentations do not need a backup.'));
    },
  };
}
