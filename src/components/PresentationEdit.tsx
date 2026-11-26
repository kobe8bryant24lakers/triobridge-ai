import { englishCount } from './presentation-runtime';
import { t } from '../i18n';
import { useState } from 'react';
import { ArrowDown, ArrowUp, Download, RefreshCw, Sparkles } from 'lucide-react';
import type { EditProposal, PresentationSnapshot } from '../presentation/types';
import {
  alignProposal,
  editMessages,
  inspectSlides,
  originalValue,
  parseEdit,
  PATCH_LABELS,
  replaceProposal,
  styleProposal,
  validateEdit,
} from '../presentation/editing';
import { downloadBlob } from '../presentation/render';
import type { PresentationRuntime } from './presentation-runtime';
import {
  appendEditingTurn,
  buildEditingFollowUp,
  type EditingTurn,
} from '../core/editing-conversation';
import { EditingConversation } from './EditingConversation';

export function PresentationEdit({
  context,
  access,
  busy,
  run,
  request,
  notify,
  modelStamp,
  demo,
}: PresentationRuntime) {
  const [scope, setScope] = useState<'selected' | 'all'>(
      access?.capabilities.selected ? 'selected' : 'all',
    ),
    [snapshot, setSnapshot] = useState<PresentationSnapshot>();
  const [brief, setBrief] = useState(''),
    [proposal, setProposal] = useState<EditProposal>(),
    [chosen, setChosen] = useState<Set<number>>(new Set()),
    [includeOrder, setIncludeOrder] = useState(true);
  const [planStamp, setPlanStamp] = useState(''),
    [needsRead, setNeedsRead] = useState(false),
    [confirmed, setConfirmed] = useState(false),
    [written, setWritten] = useState(false),
    [showAll, setShowAll] = useState(false);
  const [find, setFind] = useState(''),
    [replacement, setReplacement] = useState(''),
    [font, setFont] = useState(''),
    [accent, setAccent] = useState('#28604A'),
    [thumbnail, setThumbnail] = useState('');
  const [turns, setTurns] = useState<EditingTurn[]>([]),
    [followUp, setFollowUp] = useState(''),
    [initialInstruction, setInitialInstruction] = useState('');
  const fingerprint = JSON.stringify([
    brief,
    context.materials.map((m) => [m.id, m.name, m.text, m.tables, m.image?.modelData]),
    context.includeImages,
    modelStamp,
  ]);
  const outdated =
    needsRead ||
    Boolean(snapshot && scope !== snapshot.scope) ||
    Boolean(planStamp && planStamp !== fingerprint);
  function preview(plan: EditProposal, stamp = '') {
    setProposal(plan);
    setChosen(new Set(plan.changes.map((_, i) => i)));
    setIncludeOrder(true);
    setConfirmed(false);
    setWritten(false);
    setPlanStamp(stamp);
    setShowAll(false);
    if (!plan.changes.length && !plan.slideOrder)
      notify(t(`没有需要应用的改动。${plan.summary}`, `No changes to apply. ${plan.summary}`));
  }
  async function read() {
    await run(
      t('读取现有页面与对象', 'Reading slides and objects'),
      async (signal) => {
        if (!access)
          throw new Error(
            t(
              '当前宿主未提供页面读取，请在 PowerPoint 插件中操作。',
              'Slide reading is unavailable here. Open the workbench in PowerPoint.',
            ),
          );
        const value = await access.read(scope);
        if (signal.aborted) return;
        const objects = value.slides.reduce((total, slide) => total + slide.shapes.length, 0);
        setSnapshot(value);
        setProposal(undefined);
        setNeedsRead(false);
        setWritten(false);
        setThumbnail('');
        setConfirmed(false);
        setTurns([]);
        setFollowUp('');
        setInitialInstruction('');
        notify(
          t(
            `已读取 ${value.slides.length} 页、${objects} 个对象。`,
            `Read ${englishCount(value.slides.length, 'slide')} and ${englishCount(objects, 'object')}.`,
          ),
        );
      },
      true,
    );
  }
  async function generate() {
    if (!snapshot) return;
    await run(
      t('生成已有页面的编辑方案', 'Generating a slide edit proposal'),
      async (signal) => {
        const first = snapshot.slides[0],
          title = first.shapes.find((s) => s.role === 'title') || first.shapes.find((s) => s.text);
        const demo = {
          summary: t(
            '固定演示：修改首个标题，体验改动预览与应用。真实模型会依据你的要求和素材生成方案。',
            'This demo changes the first title. A connected model can use your instructions and sources to propose edits.',
          ),
          changes: title
            ? [
                {
                  slideId: first.id,
                  shapeId: title.id,
                  text: t(`${title.text} · 精简版`, `${title.text} · Short version`),
                },
              ]
            : [],
        };
        const raw = await request(
          editMessages(context, snapshot, brief),
          signal,
          JSON.stringify(demo),
        );
        if (signal.aborted) return;
        const result = parseEdit(raw, snapshot);
        preview(result, fingerprint);
        setTurns([{ instruction: brief, result: JSON.stringify(result) }]);
        setInitialInstruction(brief);
        setFollowUp('');
      },
      true,
    );
  }
  async function quick(action: () => EditProposal) {
    await run(t('准备改动预览', 'Preparing change preview'), async () => {
      const result = action();
      preview(result, fingerprint);
      setInitialInstruction(result.summary);
      setTurns([{ instruction: result.summary, result: JSON.stringify(result) }]);
      setFollowUp('');
    });
  }
  async function refine() {
    if (!snapshot || !proposal || outdated || written || busy || !followUp.trim()) return;
    const instruction = followUp.trim();
    await run(
      t('根据对话调整修改方案', 'Refining the edit proposal'),
      async (signal) => {
        // Revisions remain a complete proposal against the original read snapshot.
        // Unchecked changes are absent from the current draft sent to the model.
        const current = JSON.stringify({
          summary: proposal.summary,
          changes: proposal.changes.filter((_, i) => chosen.has(i)),
          ...(includeOrder && proposal.slideOrder ? { slideOrder: proposal.slideOrder } : {}),
        });
        const messages = buildEditingFollowUp(
          editMessages(context, snapshot, initialInstruction),
          turns,
          current,
          instruction,
        );
        const raw = await request(messages, signal, current);
        if (signal.aborted) return;
        const result = parseEdit(raw, snapshot);
        preview(result, fingerprint);
        setTurns((previous) =>
          appendEditingTurn(previous, current, instruction, JSON.stringify(result)),
        );
        setFollowUp('');
        notify(
          demo
            ? t(
                '固定演示保留当前方案；连接真实模型后可按对话调整。',
                'The demo keeps the current proposal. Connect a model to refine it through conversation.',
              )
            : t(
                '已更新修改方案，请检查后确认应用。',
                'Edit proposal updated. Review it before confirming changes.',
              ),
        );
      },
      true,
    );
  }
  async function apply() {
    if (!snapshot || !proposal || !access || outdated || written) return;
    await run(t('应用并核对页面改动', 'Applying and verifying slide edits'), async () => {
      const selected = validateEdit(
        {
          summary: proposal.summary,
          changes: proposal.changes.filter((_, i) => chosen.has(i)),
          ...(includeOrder && proposal.slideOrder ? { slideOrder: proposal.slideOrder } : {}),
        },
        snapshot,
      );
      setNeedsRead(true);
      setConfirmed(false);
      await access.apply(snapshot, selected);
      setWritten(true);
      notify(
        t(
          `已应用并回读核对 ${selected.changes.length} 个对象${selected.slideOrder ? '及页面顺序' : ''}。请检查实际排版；继续编辑前重新读取页面。`,
          `Applied and verified ${englishCount(selected.changes.length, 'object')}${selected.slideOrder ? ' and the slide order' : ''}. Check the layout in PowerPoint. Read the slides again before continuing.`,
        ),
      );
    });
  }
  const objectCount =
    snapshot?.slides.reduce((total, slide) => total + slide.shapes.length, 0) ?? 0;
  const selectedCount = chosen.size + (includeOrder && proposal?.slideOrder ? 1 : 0),
    findings = snapshot ? inspectSlides(snapshot) : [];
  return (
    <div className="ppt-flow">
      <section className="ppt-compose" aria-label={t('已有页面编辑', 'Edit existing slides')}>
        <h2>{t('编辑现有 PPT', 'Edit an existing deck')}</h2>
        <p className="field-help">
          {t(
            '参考已选素材，改写局部内容、统一样式、调整对象位置或页面顺序。先读取，再预览具体改动。',
            'Read the slides to review proposed changes to text, formatting, object positions or slide order using your selected sources.',
          )}
        </p>
        <div className="ppt-field-grid">
          <label>
            {t('页面范围', 'Slide range')}
            <select
              value={scope}
              disabled={busy}
              onChange={(e) => {
                setScope(e.target.value as 'selected' | 'all');
                setConfirmed(false);
              }}
            >
              <option value="selected" disabled={!access?.capabilities.selected}>
                {t('当前选中页', 'Selected slides')}
              </option>
              <option value="all">{t('整份演示稿', 'Entire deck')}</option>
            </select>
          </label>
          <button
            className="secondary"
            disabled={busy || !access?.capabilities.read}
            onClick={() => void read()}
          >
            <RefreshCw size={16} />
            {snapshot
              ? t('重新读取页面', 'Read slides again')
              : t('读取页面与对象', 'Read slides and objects')}
          </button>
        </div>
        {!access?.capabilities.read && (
          <p className="notice warning">
            {t(
              '当前版本无法读取页面对象。请更新 PowerPoint；仍可使用素材生成并下载整套 PPT。',
              'This PowerPoint version cannot read slide objects. Update PowerPoint, or build a new deck from files and download it.',
            )}
          </p>
        )}
        {snapshot && (
          <>
            <p className="ppt-scope">
              {t(
                `已读取 ${snapshot.slides.length} 页 · ${objectCount} 个对象 · `,
                `Read ${englishCount(snapshot.slides.length, 'slide')} · ${englishCount(objectCount, 'object')} · `,
              )}
              {snapshot.size
                ? t(
                    `${snapshot.size.width} × ${snapshot.size.height} 磅`,
                    `${snapshot.size.width} × ${snapshot.size.height} pt`,
                  )
                : t('页面尺寸不可用', 'Slide size unavailable')}
            </p>
            <details className="ppt-page-list">
              <summary>{t('查看读取范围与检查结果', 'Review slides and findings')}</summary>
              <ol>
                {snapshot.slides.map((s) => (
                  <li key={s.id}>
                    <span>
                      {t('第 ', 'Slide ')}
                      {s.index + 1}
                      {t(' 页 · ', ' · ')}
                      {s.title}
                    </span>
                    {access?.capabilities.thumbnail && (
                      <button
                        className="text-button"
                        disabled={busy}
                        onClick={() =>
                          void run(t('读取页面缩略图', 'Reading slide thumbnail'), async () => {
                            setThumbnail(await access.thumbnail(s.id));
                          })
                        }
                      >
                        {t('预览原页', 'Preview original slide')}
                      </button>
                    )}
                  </li>
                ))}
              </ol>
              {findings.length ? (
                <ul>
                  {findings.map((f, i) => (
                    <li key={i}>{f}</li>
                  ))}
                </ul>
              ) : (
                <p className="field-help">
                  {t(
                    '未发现明显的越界或文字密度问题；这不代替视觉检查。',
                    'No obvious overflow or dense text was found. Check the slides visually.',
                  )}
                </p>
              )}
              {snapshot.skipped > 0 && (
                <p className="field-help">
                  {snapshot.skipped}
                  {t(
                    ' 个图表、表格、组合或 SmartArt 对象只读取位置，内部文字和数据未展开。',
                    ' charts, tables, groups or SmartArt objects were read by position only. Their internal text and data were skipped.',
                  )}
                </p>
              )}
              {thumbnail && (
                <img
                  className="original-slide-image"
                  src={thumbnail}
                  alt={t('当前读取的 PowerPoint 原页', 'Original PowerPoint slide')}
                />
              )}
            </details>
            <label className="field-label" htmlFor="ppt-edit-brief">
              {t('编辑要求', 'Editing instructions')}
            </label>
            <textarea
              id="ppt-edit-brief"
              rows={4}
              maxLength={6000}
              disabled={busy}
              value={brief}
              onChange={(e) => {
                setBrief(e.target.value);
                setConfirmed(false);
              }}
              placeholder={t(
                '例如：根据素材中的最新季度报告，更新选中页的结论；把正文压缩到 3 个要点，统一标题字号，并整理对象布局。',
                'Example: Update the conclusions using the latest report. Limit the body to three bullets, standardize title sizes and align the objects.',
              )}
            />
            <p className="field-help">
              {t('将参考 ', 'Using ')}
              {context.materials.length}{' '}
              {t(
                '份素材和已读取页面。仅修改明确列出的文字、样式、对象位置或页序。',
                'source(s) and the slides you read. Only the listed changes to text, styles, positions or slide order will be applied.',
              )}
            </p>
            <button
              className="primary"
              disabled={busy || !brief.trim() || needsRead || scope !== snapshot.scope}
              onClick={() => void generate()}
            >
              <Sparkles size={17} />
              {t('生成局部修改方案', 'Generate edit proposal')}
            </button>
            <details className="ppt-quick">
              <summary>{t('批量编辑与排版工具', 'Batch editing and layout')}</summary>
              <h3>{t('统一文字样式', 'Standardize text styles')}</h3>
              <div className="ppt-field-grid">
                <label>
                  {t('字体', 'Font')}
                  <select value={font} disabled={busy} onChange={(e) => setFont(e.target.value)}>
                    <option value="">{t('保留原字体', 'Keep original fonts')}</option>
                    <option>Aptos</option>
                    <option>Arial</option>
                    <option value="等线">{t('等线', 'DengXian')}</option>
                    <option>PingFang SC</option>
                  </select>
                </label>
                <label>
                  {t('标题颜色', 'Title color')}
                  <input
                    type="color"
                    value={accent}
                    disabled={busy}
                    onChange={(e) => setAccent(e.target.value)}
                  />
                </label>
              </div>
              <p className="field-help">
                {t(
                  '标题 30 磅，正文 20 磅；保留页脚与原字体。选择新字体时，混合字体会统一；中文或缺失字体可能由 Office 替换，回读无法确认时会提示。',
                  'Titles use 30 pt and body text 20 pt. Footers and original fonts are preserved unless you select a font. Mixed fonts will then be unified. Office may substitute missing fonts and report unverified changes.',
                )}
              </p>
              <button
                className="secondary"
                disabled={busy || needsRead || scope !== snapshot.scope}
                onClick={() => void quick(() => styleProposal(snapshot, font, accent))}
              >
                {t('预览统一样式', 'Preview styles')}
              </button>
              <h3>{t('批量查找替换', 'Find and replace')}</h3>
              <label>
                {t('查找文字', 'Find text')}
                <input
                  value={find}
                  maxLength={100}
                  disabled={busy}
                  onChange={(e) => setFind(e.target.value)}
                />
              </label>
              <label>
                {t('替换为', 'Replace with')}
                <input
                  value={replacement}
                  maxLength={200}
                  disabled={busy}
                  onChange={(e) => setReplacement(e.target.value)}
                />
              </label>
              <button
                className="secondary"
                disabled={busy || !find || needsRead || scope !== snapshot.scope}
                onClick={() => void quick(() => replaceProposal(snapshot, find, replacement))}
              >
                {t('预览批量替换', 'Preview replacements')}
              </button>
              <h3>{t('选中对象排版', 'Align selected objects')}</h3>
              <p className="field-help">
                {t(
                  '先在同一页多选对象，再读取“当前选中页”。等距排列需至少 3 个对象。',
                  'Select objects on one slide, then read Selected slides. Horizontal distribution requires at least three objects.',
                )}
              </p>
              <div className="button-row">
                {(['left', 'center', 'distribute'] as const).map((mode) => (
                  <button
                    key={mode}
                    className="secondary"
                    disabled={busy || needsRead || scope !== snapshot.scope || !snapshot.size}
                    onClick={() => void quick(() => alignProposal(snapshot, mode))}
                  >
                    {mode === 'left'
                      ? t('左对齐', 'Align left')
                      : mode === 'center'
                        ? t('居中对齐', 'Align centers')
                        : t('水平等距', 'Distribute horizontally')}
                  </button>
                ))}
              </div>
              <h3>{t('调整页面顺序', 'Reorder slides')}</h3>
              <p className="field-help">
                {t(
                  '只重排读取范围内的页面，范围外页面位置保持原样。选择整稿范围可调整全稿。',
                  'Only the slides you read will be reordered. Select Entire deck to change the order throughout the presentation.',
                )}
              </p>
              <ol className="reorder-list">
                {(proposal?.slideOrder || snapshot.slides.map((s) => s.id)).map((id, i, ids) => (
                  <li key={id}>
                    <span>
                      {i + 1}. {snapshot.slides.find((s) => s.id === id)?.title}
                    </span>
                    <div>
                      {[-1, 1].map((direction) => (
                        <button
                          key={direction}
                          className="icon-button"
                          aria-label={t(
                            `${i + 1} 页${direction < 0 ? '前移' : '后移'}`,
                            `Move slide ${i + 1} ${direction < 0 ? 'up' : 'down'}`,
                          )}
                          disabled={
                            busy ||
                            needsRead ||
                            scope !== snapshot.scope ||
                            !access?.capabilities.reorder ||
                            i + direction < 0 ||
                            i + direction >= ids.length
                          }
                          onClick={() =>
                            void quick(() => {
                              const order = [...ids];
                              [order[i], order[i + direction]] = [order[i + direction], order[i]];
                              return validateEdit(
                                {
                                  summary: t(
                                    '调整已读取页面的顺序。',
                                    'Reorder the slides you read.',
                                  ),
                                  changes: [],
                                  slideOrder: order,
                                },
                                snapshot,
                              );
                            })
                          }
                        >
                          {direction < 0 ? <ArrowUp size={15} /> : <ArrowDown size={15} />}
                        </button>
                      ))}
                    </div>
                  </li>
                ))}
              </ol>
            </details>
          </>
        )}
      </section>
      {proposal && snapshot && (
        <section className="ppt-result" aria-label={t('改动预览', 'Change preview')}>
          <div className="section-heading">
            <h2>{t('改动预览', 'Change preview')}</h2>
            <span>
              {t(
                `${proposal.changes.length} 个对象`,
                `${englishCount(proposal.changes.length, 'object')}`,
              )}
              {proposal.slideOrder ? t(' + 页序', ' + slide order') : ''}
            </span>
          </div>
          <p>{proposal.summary}</p>
          {outdated && !written && (
            <p className="notice warning" role="status">
              {t(
                '方案已过期或已尝试写入。请重新读取页面并生成方案，避免重复覆盖。',
                'This proposal is stale or a write was attempted. Read the slides again and generate a new proposal.',
              )}
            </p>
          )}
          {written && (
            <p className="notice success" role="status">
              {t(
                '改动已应用并核对。继续编辑前请重新读取。',
                'Changes applied and verified. Read the slides again to continue editing.',
              )}
            </p>
          )}
          {proposal.changes.length > 0 && (
            <label className="checkbox-label">
              <input
                type="checkbox"
                disabled={busy || written || outdated}
                checked={chosen.size === proposal.changes.length}
                onChange={(e) => {
                  setChosen(
                    e.target.checked ? new Set(proposal.changes.map((_, i) => i)) : new Set(),
                  );
                  setConfirmed(false);
                }}
              />
              {t('选择全部对象改动', 'Select all changes')}
            </label>
          )}
          <div className="change-list">
            {proposal.changes.slice(0, showAll ? undefined : 50).map((patch, i) => {
              const slide = snapshot.slides.find((s) => s.id === patch.slideId)!,
                shape = slide.shapes.find((s) => s.id === patch.shapeId)!;
              return (
                <article className="change-row" key={`${patch.slideId}/${patch.shapeId}`}>
                  <label className="checkbox-label">
                    <input
                      type="checkbox"
                      checked={chosen.has(i)}
                      disabled={busy || written || outdated}
                      onChange={(e) => {
                        const next = new Set(chosen);
                        if (e.target.checked) next.add(i);
                        else next.delete(i);
                        setChosen(next);
                        setConfirmed(false);
                      }}
                    />
                    <strong>
                      {t('第 ', 'Slide ')}
                      {slide.index + 1}
                      {t(' 页 · ', ' · ')}
                      {shape.name || shape.id}
                    </strong>
                  </label>
                  <dl>
                    {Object.keys(PATCH_LABELS)
                      .filter((k) => Object.prototype.hasOwnProperty.call(patch, k))
                      .map((key) => (
                        <div key={key}>
                          <dt>{PATCH_LABELS[key]}</dt>
                          <dd>
                            <span className="change-before">
                              {String(
                                originalValue(shape, key) ?? t('混合样式', 'Mixed formatting'),
                              )}
                            </span>
                            <span className="change-after">
                              {String((patch as unknown as Record<string, unknown>)[key]) ||
                                t('（清空）', '(clear)')}
                            </span>
                          </dd>
                        </div>
                      ))}
                  </dl>
                </article>
              );
            })}
          </div>
          {!showAll && proposal.changes.length > 50 && (
            <button className="text-button" onClick={() => setShowAll(true)}>
              {t('显示其余 ', 'Show ')}
              {proposal.changes.length - 50}
              {t(' 个改动', ' more changes')}
            </button>
          )}
          {proposal.slideOrder && (
            <label className="checkbox-label">
              <input
                type="checkbox"
                checked={includeOrder}
                disabled={busy || written || outdated}
                onChange={(e) => {
                  setIncludeOrder(e.target.checked);
                  setConfirmed(false);
                }}
              />
              <span>
                {t('调整页序：', 'New slide order: ')}
                {proposal.slideOrder
                  .map((id) =>
                    t(
                      `原第 ${snapshot.slides.find((s) => s.id === id)!.index + 1} 页`,
                      `Original slide ${snapshot.slides.find((s) => s.id === id)!.index + 1}`,
                    ),
                  )
                  .join(' → ')}
              </span>
            </label>
          )}
          <p className="field-help">
            {t(
              '改写文字可能影响框内局部格式；统一字体会替换混合样式。应用前建议备份原稿，应用后可使用 Office 撤销。',
              'Rewriting can change local formatting. Applying one font replaces mixed styles. Save a backup first. Use Office Undo to revert.',
            )}
          </p>
          <EditingConversation
            id="slides-conversation"
            turns={turns}
            value={followUp}
            onChange={setFollowUp}
            onSubmit={() => void refine()}
            onReset={() => {
              setTurns([]);
              setFollowUp('');
            }}
            disabled={busy}
            stale={outdated || written}
          />
          <div className="button-row">
            <button
              className="primary"
              disabled={busy || outdated || written || selectedCount === 0}
              onClick={() => setConfirmed(true)}
            >
              {t(
                `应用所选 ${selectedCount} 项改动`,
                `Apply ${englishCount(selectedCount, 'selected change')}`,
              )}
            </button>
            {access?.capabilities.backup && (
              <button
                className="secondary"
                disabled={busy}
                onClick={() =>
                  void run(t('导出原稿备份', 'Exporting a backup'), async () => {
                    downloadBlob(
                      await access.backup(),
                      t('TrioBridge-原稿备份.pptx', 'TrioBridge-Original.pptx'),
                    );
                    notify(
                      t(
                        '已请求下载原稿备份，请确认文件已保存。',
                        'The backup download was requested. Check that the file was saved.',
                      ),
                    );
                  })
                }
              >
                <Download size={16} />
                {t('下载原稿备份', 'Download original backup')}
              </button>
            )}
          </div>
          {confirmed && !outdated && (
            <div className="ppt-confirm">
              <p>
                {t(
                  `确认修改 ${snapshot.scope === 'all' ? '整稿' : '选中页'}范围内的 ${chosen.size} 个对象${includeOrder && proposal.slideOrder ? '及页序' : ''}？应用前会再次核对页面和内容。`,
                  `Apply changes to ${englishCount(chosen.size, 'object')} in ${snapshot.scope === 'all' ? 'the entire deck' : 'the selected slides'}${includeOrder && proposal.slideOrder ? ' and change the slide order' : ''}? The slides and content will be checked again.`,
                )}
              </p>
              <div className="button-row">
                <button className="primary" disabled={busy} onClick={() => void apply()}>
                  {t('确认应用改动', 'Confirm changes')}
                </button>
                <button className="secondary" disabled={busy} onClick={() => setConfirmed(false)}>
                  {t('返回检查', 'Back to review')}
                </button>
              </div>
            </div>
          )}
        </section>
      )}
    </div>
  );
}
