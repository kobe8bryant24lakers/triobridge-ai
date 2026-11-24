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

export function PresentationEdit({
  context,
  access,
  busy,
  run,
  request,
  notify,
  modelStamp,
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
  const fingerprint = JSON.stringify([
    brief,
    context.materials.map((m) => [m.id, m.name, m.text, m.image?.modelData]),
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
    if (!plan.changes.length && !plan.slideOrder) notify(`没有需要应用的改动。${plan.summary}`);
  }
  async function read() {
    await run('读取现有页面与对象', async (signal) => {
      if (!access) throw new Error('当前宿主未提供页面读取，请在 PowerPoint 插件中操作。');
      const value = await access.read(scope);
      if (signal.aborted) return;
      setSnapshot(value);
      setProposal(undefined);
      setNeedsRead(false);
      setWritten(false);
      setThumbnail('');
      setConfirmed(false);
      notify(
        `已读取 ${value.slides.length} 页、${value.slides.reduce((n, s) => n + s.shapes.length, 0)} 个对象。`,
      );
    });
  }
  async function generate() {
    if (!snapshot) return;
    await run('生成已有页面的编辑方案', async (signal) => {
      const first = snapshot.slides[0],
        title = first.shapes.find((s) => s.role === 'title') || first.shapes.find((s) => s.text);
      const demo = {
        summary:
          '固定演示：修改首个标题，体验改动预览与应用。真实模型会依据你的要求和素材生成方案。',
        changes: title
          ? [{ slideId: first.id, shapeId: title.id, text: `${title.text} · 精简版` }]
          : [],
      };
      const raw = await request(
        editMessages(context, snapshot, brief),
        signal,
        JSON.stringify(demo),
      );
      if (signal.aborted) return;
      preview(parseEdit(raw, snapshot), fingerprint);
    });
  }
  async function quick(action: () => EditProposal) {
    await run('准备改动预览', async () => {
      preview(action());
    });
  }
  async function apply() {
    if (!snapshot || !proposal || !access || outdated || written) return;
    await run('应用并核对页面改动', async () => {
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
        `已应用并回读核对 ${selected.changes.length} 个对象${selected.slideOrder ? '及页面顺序' : ''}。请检查实际排版；继续编辑前重新读取页面。`,
      );
    });
  }
  const selectedCount = chosen.size + (includeOrder && proposal?.slideOrder ? 1 : 0),
    findings = snapshot ? inspectSlides(snapshot) : [];
  return (
    <div className="ppt-flow">
      <section className="ppt-compose" aria-label="已有页面编辑">
        <h2>编辑现有 PPT</h2>
        <p className="field-help">
          参考已选素材，改写局部内容、统一样式、调整对象位置或页面顺序。先读取，再预览具体改动。
        </p>
        <div className="ppt-field-grid">
          <label>
            页面范围
            <select
              value={scope}
              disabled={busy}
              onChange={(e) => {
                setScope(e.target.value as 'selected' | 'all');
                setConfirmed(false);
              }}
            >
              <option value="selected" disabled={!access?.capabilities.selected}>
                当前选中页
              </option>
              <option value="all">整份演示稿</option>
            </select>
          </label>
          <button
            className="secondary"
            disabled={busy || !access?.capabilities.read}
            onClick={() => void read()}
          >
            <RefreshCw size={16} />
            {snapshot ? '重新读取页面' : '读取页面与对象'}
          </button>
        </div>
        {!access?.capabilities.read && (
          <p className="notice warning">
            当前版本无法读取页面对象。请更新 PowerPoint；仍可使用素材生成并下载整套 PPT。
          </p>
        )}
        {snapshot && (
          <>
            <p className="ppt-scope">
              已读取 {snapshot.slides.length} 页 ·{' '}
              {snapshot.slides.reduce((n, s) => n + s.shapes.length, 0)} 个对象 ·{' '}
              {snapshot.size
                ? `${snapshot.size.width} × ${snapshot.size.height} 磅`
                : '页面尺寸不可用'}
            </p>
            <details className="ppt-page-list">
              <summary>查看读取范围与检查结果</summary>
              <ol>
                {snapshot.slides.map((s) => (
                  <li key={s.id}>
                    <span>
                      第 {s.index + 1} 页 · {s.title}
                    </span>
                    {access?.capabilities.thumbnail && (
                      <button
                        className="text-button"
                        disabled={busy}
                        onClick={() =>
                          void run('读取页面缩略图', async () => {
                            setThumbnail(await access.thumbnail(s.id));
                          })
                        }
                      >
                        预览原页
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
                <p className="field-help">未发现明显的越界或文字密度问题；这不代替视觉检查。</p>
              )}
              {snapshot.skipped > 0 && (
                <p className="field-help">
                  {snapshot.skipped} 个图表、表格、组合或 SmartArt
                  对象只读取位置，内部文字和数据未展开。
                </p>
              )}
              {thumbnail && (
                <img
                  className="original-slide-image"
                  src={thumbnail}
                  alt="当前读取的 PowerPoint 原页"
                />
              )}
            </details>
            <label className="field-label" htmlFor="ppt-edit-brief">
              编辑要求
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
              placeholder="例如：根据素材中的最新季度报告，更新选中页的结论；把正文压缩到 3 个要点，统一标题字号，并整理对象布局。"
            />
            <p className="field-help">
              将参考 {context.materials.length}{' '}
              份素材和已读取页面。仅修改明确列出的文字、样式、对象位置或页序。
            </p>
            <button
              className="primary"
              disabled={busy || !brief.trim() || needsRead || scope !== snapshot.scope}
              onClick={() => void generate()}
            >
              <Sparkles size={17} />
              生成局部修改方案
            </button>
            <details className="ppt-quick">
              <summary>批量编辑与排版工具</summary>
              <h3>统一文字样式</h3>
              <div className="ppt-field-grid">
                <label>
                  字体
                  <select value={font} disabled={busy} onChange={(e) => setFont(e.target.value)}>
                    <option value="">保留原字体</option>
                    <option>Aptos</option>
                    <option>Arial</option>
                    <option>等线</option>
                    <option>PingFang SC</option>
                  </select>
                </label>
                <label>
                  标题颜色
                  <input
                    type="color"
                    value={accent}
                    disabled={busy}
                    onChange={(e) => setAccent(e.target.value)}
                  />
                </label>
              </div>
              <p className="field-help">
                标题 30 磅，正文 20
                磅；保留页脚与原字体。选择新字体时，混合字体会统一；中文或缺失字体可能由 Office
                替换，回读无法确认时会提示。
              </p>
              <button
                className="secondary"
                disabled={busy || needsRead || scope !== snapshot.scope}
                onClick={() => void quick(() => styleProposal(snapshot, font, accent))}
              >
                预览统一样式
              </button>
              <h3>批量查找替换</h3>
              <label>
                查找文字
                <input
                  value={find}
                  maxLength={100}
                  disabled={busy}
                  onChange={(e) => setFind(e.target.value)}
                />
              </label>
              <label>
                替换为
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
                预览批量替换
              </button>
              <h3>选中对象排版</h3>
              <p className="field-help">
                先在同一页多选对象，再读取“当前选中页”。等距排列需至少 3 个对象。
              </p>
              <div className="button-row">
                {(['left', 'center', 'distribute'] as const).map((mode) => (
                  <button
                    key={mode}
                    className="secondary"
                    disabled={busy || needsRead || scope !== snapshot.scope || !snapshot.size}
                    onClick={() => void quick(() => alignProposal(snapshot, mode))}
                  >
                    {mode === 'left' ? '左对齐' : mode === 'center' ? '居中对齐' : '水平等距'}
                  </button>
                ))}
              </div>
              <h3>调整页面顺序</h3>
              <p className="field-help">
                只重排读取范围内的页面，范围外页面位置保持原样。选择整稿范围可调整全稿。
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
                          aria-label={`${i + 1} 页${direction < 0 ? '前移' : '后移'}`}
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
                                  summary: '调整已读取页面的顺序。',
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
        <section className="ppt-result" aria-label="改动预览">
          <div className="section-heading">
            <h2>改动预览</h2>
            <span>
              {proposal.changes.length} 个对象{proposal.slideOrder ? ' + 页序' : ''}
            </span>
          </div>
          <p>{proposal.summary}</p>
          {outdated && !written && (
            <p className="notice warning" role="status">
              方案已过期或已尝试写入。请重新读取页面并生成方案，避免重复覆盖。
            </p>
          )}
          {written && (
            <p className="notice success" role="status">
              改动已应用并核对。继续编辑前请重新读取。
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
              选择全部对象改动
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
                      第 {slide.index + 1} 页 · {shape.name || shape.id}
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
                              {String(originalValue(shape, key) ?? '混合样式')}
                            </span>
                            <span className="change-after">
                              {String((patch as unknown as Record<string, unknown>)[key]) ||
                                '（清空）'}
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
              显示其余 {proposal.changes.length - 50} 个改动
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
                调整页序：
                {proposal.slideOrder
                  .map((id) => `原第 ${snapshot.slides.find((s) => s.id === id)!.index + 1} 页`)
                  .join(' → ')}
              </span>
            </label>
          )}
          <p className="field-help">
            改写文字可能影响框内局部格式；统一字体会替换混合样式。应用前建议备份原稿，应用后可使用
            Office 撤销。
          </p>
          <div className="button-row">
            <button
              className="primary"
              disabled={busy || outdated || written || selectedCount === 0}
              onClick={() => setConfirmed(true)}
            >
              应用所选 {selectedCount} 项改动
            </button>
            {access?.capabilities.backup && (
              <button
                className="secondary"
                disabled={busy}
                onClick={() =>
                  void run('导出原稿备份', async () => {
                    downloadBlob(await access.backup(), 'TrioBridge-原稿备份.pptx');
                    notify('已请求下载原稿备份，请确认文件已保存。');
                  })
                }
              >
                <Download size={16} />
                下载原稿备份
              </button>
            )}
          </div>
          {confirmed && !outdated && (
            <div className="ppt-confirm">
              <p>
                确认修改 {snapshot.scope === 'all' ? '整稿' : '选中页'}范围内的 {chosen.size} 个对象
                {includeOrder && proposal.slideOrder ? '及页序' : ''}？应用前会再次核对页面和内容。
              </p>
              <div className="button-row">
                <button className="primary" disabled={busy} onClick={() => void apply()}>
                  确认应用改动
                </button>
                <button className="secondary" disabled={busy} onClick={() => setConfirmed(false)}>
                  返回检查
                </button>
              </div>
            </div>
          )}
        </section>
      )}
    </div>
  );
}
