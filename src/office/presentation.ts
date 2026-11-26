import { uiMessage } from '../i18n-errors';
import { t } from '../i18n';
import type {
  EditProposal,
  PresentationAccess,
  PresentationInsertTarget,
  PresentationSnapshot,
  ShapeState,
  SlideBackground,
} from '../presentation/types';
import {
  assertSnapshotUnchanged,
  originalValue,
  reorderedIds,
  validateEdit,
  PATCH_LABELS,
} from '../presentation/editing';

const supports = (version: string) =>
  Office.context.requirements.isSetSupported('PowerPointApi', version);
const textTypes = new Set(['TextBox', 'GeometricShape', 'Placeholder']);
const round = (n: number) => Math.round(n * 100) / 100;

async function loadInsertTarget(ctx: PowerPoint.RequestContext): Promise<PresentationInsertTarget> {
  if (!supports('1.10'))
    throw new Error('当前 Office 无法读取准确页面尺寸，请下载 PPT 后手动导入。');
  const slides = ctx.presentation.slides,
    pageSize = ctx.presentation.pageSetup;
  slides.load('items/id');
  pageSize.load('slideWidth,slideHeight');
  await ctx.sync();
  if (slides.items.length > 500) throw new Error('演示稿超过 500 页，请拆分后再插入。');
  return {
    slideIds: slides.items.map((slide) => slide.id),
    size: { width: pageSize.slideWidth, height: pageSize.slideHeight },
  };
}

async function readPresentation(
  scope: 'selected' | 'all',
  fixedIds?: string[],
  prior?: PresentationSnapshot,
): Promise<PresentationSnapshot> {
  if (!supports('1.4'))
    throw new Error(
      '当前 PowerPoint 版本无法读取页面对象，请更新 Office。仍可根据素材生成并下载 PPT。',
    );
  if (scope === 'selected' && !supports('1.5'))
    throw new Error('当前版本无法读取选中页，请改为整稿范围。');
  return PowerPoint.run(async (ctx) => {
    const all = ctx.presentation.slides;
    all.load('items/id');
    const selected =
      !fixedIds && scope === 'selected' ? ctx.presentation.getSelectedSlides() : null;
    selected?.load('items/id');
    const selection =
      !fixedIds && scope === 'selected' ? ctx.presentation.getSelectedShapes() : null;
    selection?.load('items/id');
    const pageSize = supports('1.10') ? ctx.presentation.pageSetup : null;
    pageSize?.load('slideWidth,slideHeight');
    await ctx.sync();
    const slideIds = all.items.map((s) => s.id),
      ids = fixedIds || (selected ? selected.items.map((s) => s.id) : slideIds);
    if (!ids.length) throw new Error('请在左侧缩略图中选中至少一页。');
    if (ids.length > 60 || slideIds.length > 500)
      throw new Error('一次最多读取 60 页，演示稿最多 500 页。请缩小选区或拆分演示稿。');
    const slides = all.items.filter((s) => ids.includes(s.id));
    if (slides.length !== ids.length) throw new Error('目标页面已经删除，请重新读取。');
    const backgrounds = new Map(
      slides
        .filter((s) => supports('1.10') && s.background)
        .map((s) => {
          const bg = s.background;
          bg.load('isMasterBackgroundFollowed');
          bg.fill.load('type');
          const solid = bg.fill.getSolidFillOrNullObject();
          solid.load('color,transparency');
          const image = bg.fill.getPictureOrTextureFillOrNullObject();
          image.load('transparency');
          return [s.id, { bg, solid, image }] as const;
        }),
    );
    slides.forEach((s) =>
      s.shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height'),
    );
    await ctx.sync();
    if (slides.reduce((n, s) => n + s.shapes.items.length, 0) > 1500)
      throw new Error('所选页超过 1,500 个对象，请缩小范围。');
    const texts = slides.flatMap((s) => s.shapes.items.filter((sh) => textTypes.has(sh.type)));
    texts.forEach((s) => s.textFrame.load('hasText'));
    const placeholders = supports('1.8') ? texts.filter((s) => s.type === 'Placeholder') : [];
    placeholders.forEach((s) => s.placeholderFormat.load('type'));
    await ctx.sync();
    const filled = texts.filter((s) => s.textFrame.hasText);
    filled.forEach((s) => {
      s.textFrame.textRange.load('text');
      s.textFrame.textRange.font.load('name,size,color,bold');
      s.textFrame.textRange.paragraphFormat.load('horizontalAlignment');
    });
    await ctx.sync();
    let skipped = 0;
    const result: PresentationSnapshot = {
      scope,
      slideIds,
      size: pageSize ? { width: pageSize.slideWidth, height: pageSize.slideHeight } : null,
      selectedShapeIds: selection?.items.map((s) => s.id) || [],
      skipped: 0,
      slides: slides.map((slide) => {
        const shapes: ShapeState[] = slide.shapes.items.map((s) => {
          const shape: ShapeState = {
            id: s.id,
            name: s.name,
            type: s.type,
            left: round(s.left),
            top: round(s.top),
            width: round(s.width),
            height: round(s.height),
          };
          if (filled.includes(s)) {
            const t = s.textFrame.textRange,
              placeholder = placeholders.includes(s) ? s.placeholderFormat.type : '';
            shape.text = t.text;
            shape.font = {
              name: t.font.name,
              size: t.font.size,
              color: t.font.color?.toUpperCase() || null,
              bold: t.font.bold,
            };
            shape.alignment = t.paragraphFormat.horizontalAlignment;
            shape.role =
              /Footer|Header|SlideNumber|Date/.test(placeholder) ||
              s.name === 'TrioBridge Footer' ||
              (s.top >= (pageSize?.slideHeight || 540) * 0.88 && (t.font.size || 0) <= 14)
                ? 'footer'
                : /Title/.test(placeholder) ||
                    s.name === 'TrioBridge Title' ||
                    ((t.font.size || 0) >= 22 && s.top < (pageSize?.slideHeight || 540) * 0.24)
                  ? 'title'
                  : 'body';
          } else if (['Group', 'Table', 'Chart', 'SmartArt'].includes(s.type)) skipped++;
          return shape;
        });
        const fill = backgrounds.get(slide.id);
        let background: SlideBackground | undefined;
        if (fill) {
          if (fill.bg.isMasterBackgroundFollowed) background = { kind: 'master' };
          else if (!fill.solid.isNullObject && !fill.solid.transparency)
            background = { kind: 'solid', color: fill.solid.color.toUpperCase() };
          else if (!fill.image.isNullObject)
            background = {
              kind: 'image',
              imageId: 'office:' + slide.id,
              transparency: Math.round(fill.image.transparency * 100),
            };
          else background = { kind: 'other', label: fill.bg.fill.type };
        }
        return {
          ...(background ? { background } : {}),
          id: slide.id,
          index: slideIds.indexOf(slide.id),
          title:
            shapes.find((s) => s.role === 'title')?.text ||
            shapes.find((s) => s.text)?.text?.slice(0, 80) ||
            // Keep a synthetic title stable across interface-language changes during verification.
            prior?.slides.find((previous) => previous.id === slide.id)?.title ||
            t(`第 ${slideIds.indexOf(slide.id) + 1} 页`, `Slide ${slideIds.indexOf(slide.id) + 1}`),
          shapes,
        };
      }),
    };
    result.skipped = skipped;
    if (JSON.stringify(result).length > 150_000) throw new Error('所选页面内容过多，请缩小范围。');
    return result;
  });
}

async function applyEdits(
  snapshot: PresentationSnapshot,
  input: EditProposal,
  images: Record<string, string> = {},
) {
  const proposal = validateEdit(input, snapshot);
  const fills = (proposal.actions || []).map((a) => {
    if (a.kind !== 'background' || !supports('1.10'))
      throw Error(
        t('请在 PPTX 工作副本中应用这项编辑。', 'Open a PPTX working copy for this edit.'),
      );
    const fill =
      a.fill.kind === 'reference'
        ? snapshot.slides.find((s) => s.id === (a.fill as { sourceId: string }).sourceId)
            ?.background
        : a.fill;
    if (
      !fill ||
      !['master', 'solid', 'image'].includes(fill.kind) ||
      (fill.kind === 'image' && !images[fill.imageId])
    )
      throw Error(
        t(
          '请在 PPTX 工作副本中复制该背景或指定双色渐变。',
          'Open a PPTX working copy to copy this background or set both gradient colors.',
        ),
      );
    return { id: a.slideId, fill };
  });
  if (proposal.slideOrder && !supports('1.8'))
    throw new Error('当前 PowerPoint 不支持页面重排，请更新 Office。');
  // Re-read every scoped object before any writes, including style/geometry and page order.
  assertSnapshotUnchanged(snapshot, await readPresentation(snapshot.scope, undefined, snapshot));
  let started = false;
  try {
    await PowerPoint.run(async (ctx) => {
      for (const patch of proposal.changes) {
        const shape = ctx.presentation.slides.getItem(patch.slideId).shapes.getItem(patch.shapeId);
        if (patch.text !== undefined) {
          const before = snapshot.slides
            .find((s) => s.id === patch.slideId)!
            .shapes.find((s) => s.id === patch.shapeId)!.text!;
          // Preserve unaffected runs where possible instead of resetting the whole text frame.
          let start = 0,
            end = 0;
          while (
            start < Math.min(before.length, patch.text.length) &&
            before[start] === patch.text[start]
          )
            start++;
          while (
            end < Math.min(before.length, patch.text.length) - start &&
            before[before.length - 1 - end] === patch.text[patch.text.length - 1 - end]
          )
            end++;
          if (before.length - start - end > 0)
            shape.textFrame.textRange.getSubstring(start, before.length - start - end).text =
              patch.text.slice(start, patch.text.length - end);
          else shape.textFrame.textRange.text = patch.text;
        }
        if (patch.fontName !== undefined) shape.textFrame.textRange.font.name = patch.fontName;
        if (patch.fontSize !== undefined) shape.textFrame.textRange.font.size = patch.fontSize;
        if (patch.fontColor !== undefined) shape.textFrame.textRange.font.color = patch.fontColor;
        if (patch.bold !== undefined) shape.textFrame.textRange.font.bold = patch.bold;
        if (patch.alignment !== undefined)
          shape.textFrame.textRange.paragraphFormat.horizontalAlignment = patch.alignment;
        for (const key of ['left', 'top', 'width', 'height'] as const)
          if (patch[key] !== undefined) shape[key] = patch[key]!;
      }
      for (const { id, fill } of fills) {
        const bg = ctx.presentation.slides.getItem(id).background;
        bg.isMasterBackgroundFollowed = fill.kind === 'master';
        if (fill.kind === 'solid') bg.fill.setSolidFill({ color: fill.color, transparency: 0 });
        if (fill.kind === 'image')
          bg.fill.setPictureOrTextureFill({
            imageBase64: images[fill.imageId].split(',').pop(),
            transparency: fill.transparency / 100,
          });
      }
      if (proposal.slideOrder)
        reorderedIds(snapshot, proposal.slideOrder).forEach((id, index) =>
          ctx.presentation.slides.getItem(id).moveTo(index),
        );
      started = true;
      await ctx.sync();
    });
    const now = await readPresentation(
      snapshot.scope,
      snapshot.slides.map((s) => s.id),
    );
    if (
      JSON.stringify(now.slideIds) !== JSON.stringify(reorderedIds(snapshot, proposal.slideOrder))
    )
      throw new Error('页面顺序核对未通过。');
    for (const { id, fill } of fills) {
      const actual = now.slides.find((s) => s.id === id)?.background;
      if (
        actual?.kind !== fill.kind ||
        (fill.kind === 'solid' && actual.kind === 'solid' && actual.color !== fill.color) ||
        (fill.kind === 'image' &&
          actual.kind === 'image' &&
          actual.transparency !== fill.transparency)
      )
        throw Error(
          t('背景回读与预览不一致。', 'The background read back differs from the preview.'),
        );
    }
    for (const patch of proposal.changes) {
      const shape = now.slides
        .find((s) => s.id === patch.slideId)
        ?.shapes.find((s) => s.id === patch.shapeId);
      if (!shape) throw new Error('修改后的对象未找到。');
      for (const key of Object.keys(patch).filter((k) => !['slideId', 'shapeId'].includes(k))) {
        const actual = originalValue(shape, key),
          expected = (patch as unknown as Record<string, unknown>)[key];
        if (
          typeof expected === 'number' && typeof actual === 'number'
            ? Math.abs(expected - actual) > 0.15
            : actual !== expected
        )
          throw new Error(
            t(
              `部分改动的回读结果与方案不一致（${shape.name} · ${PATCH_LABELS[key] || key}）。`,
              `The value read back differs from the proposed change (${shape.name} · ${PATCH_LABELS[key] || key}).`,
            ),
          );
      }
    }
  } catch (error) {
    if (started)
      throw new Error(
        t(
          `Office 可能已应用部分改动，尚未确认全部成功。请检查页面，使用 Office 撤销或打开原稿备份；重新读取后再操作。${error instanceof Error && error.name !== 'RichApi.Error' ? ` ${error.message}` : ''}`,
          `Some changes may have been applied. Check the slides, use Undo or restore a backup, and read the presentation again before retrying.${error instanceof Error && error.name !== 'RichApi.Error' ? ` ${uiMessage(error.message)}` : ''}`,
        ),
      );
    throw error;
  }
}

async function backupPresentation(): Promise<Blob> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 4 * 1024 * 1024 },
      async (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error('当前 Office 无法导出原稿，请使用“另存为”备份。'));
          return;
        }
        const file = result.value;
        try {
          if (file.size > 150_000_000)
            throw new Error('原稿超过 150 MB，请使用 Office“另存为”备份。');
          const slices: Uint8Array<ArrayBuffer>[] = [];
          for (let i = 0; i < file.sliceCount; i++)
            slices.push(
              await new Promise((ok, fail) =>
                file.getSliceAsync(i, (r) =>
                  r.status === Office.AsyncResultStatus.Succeeded
                    ? ok(new Uint8Array(r.value.data))
                    : fail(new Error('原稿备份读取失败，请使用 Office“另存为”。')),
                ),
              ),
            );
          resolve(
            new Blob(slices, {
              type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            }),
          );
        } catch (error) {
          reject(error);
        } finally {
          file.closeAsync();
        }
      },
    );
  });
}
export function createPresentationAccess(): PresentationAccess {
  return {
    capabilities: {
      backgrounds: supports('1.10'),
      pictureBackground: supports('1.10'),
      read: supports('1.4'),
      selected: supports('1.5'),
      reorder: supports('1.8'),
      insert: supports('1.2'),
      create: supports('1.1'),
      thumbnail: supports('1.8'),
      backup: Office.context.requirements.isSetSupported('CompressedFile', '1.1'),
    },
    read: readPresentation,
    focus: async (id, shape) => {
      if (supports('1.5'))
        await PowerPoint.run(async (ctx) => {
          ctx.presentation.setSelectedSlides([id]);
          if (shape) ctx.presentation.slides.getItem(id).setSelectedShapes([shape]);
          await ctx.sync();
        });
    },
    prepareInsert: () => PowerPoint.run(loadInsertTarget),
    apply: applyEdits,
    create: async (base64) => {
      await PowerPoint.createPresentation(base64);
    },
    insert: async (base64, count, target, afterSlideId) => {
      if (!supports('1.2')) throw new Error('当前 Office 不支持插入整页，请下载后手动导入。');
      let started = false;
      try {
        await PowerPoint.run(async (ctx) => {
          const now = await loadInsertTarget(ctx);
          if (
            JSON.stringify(target.slideIds) !== JSON.stringify(now.slideIds) ||
            target.size?.width !== now.size?.width ||
            target.size?.height !== now.size?.height
          )
            throw new Error('演示稿页序或尺寸已变化，尚未插入。请检查当前演示稿并更新预览。');
          const slides = ctx.presentation.slides;
          const before = now.slideIds;
          const pos = afterSlideId === undefined ? before.length : before.indexOf(afterSlideId) + 1;
          if (afterSlideId !== undefined && !pos)
            throw Error(t('插入位置已变化。', 'The insertion point changed.'));
          ctx.presentation.insertSlidesFromBase64(base64, {
            formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
            ...(before.length ? { targetSlideId: before[pos - 1] } : {}),
          });
          started = true;
          await ctx.sync();
          slides.load('items/id');
          await ctx.sync();
          const after = slides.items.map((s) => s.id);
          if (
            after.length !== before.length + count ||
            before.some((id, i) => after[i < pos ? i : i + count] !== id)
          )
            throw new Error('插入后的页数或位置核对未通过。');
        });
      } catch (error) {
        if (started)
          throw new Error(
            '插入未能确认全部成功，请检查演示稿中的插入位置，必要时使用 Office 撤销。不要直接重复插入。',
          );
        throw error;
      }
    },
    thumbnail: async (id) =>
      PowerPoint.run(async (ctx) => {
        const image = ctx.presentation.slides.getItem(id).getImageAsBase64({ width: 800 });
        await ctx.sync();
        return `data:image/png;base64,${image.value}`;
      }),
    backup: backupPresentation,
  };
}
