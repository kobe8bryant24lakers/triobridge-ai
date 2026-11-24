import { useState } from 'react';
import { ArrowDown, ArrowUp, Download, Presentation, Sparkles, Trash2 } from 'lucide-react';
import type { DeckContent, DeckSlide, MaterialContext } from '../presentation/types';
import { deckMessages, parseDeck } from '../presentation/content';
import {
  DEFAULT_THEME,
  downloadPresentation,
  renderDeck,
  STANDARD,
  WIDE,
} from '../presentation/render';
import { SlidePreview } from './PresentationPreview';
import type { PresentationRuntime } from './presentation-runtime';

function demoDeck(count: number): DeckContent {
  const pages: DeckSlide[] = [
    {
      title: '季度业务回顾 · 演示示例',
      layout: 'cover',
      subtitle: '固定示例，用于体验整稿生成流程',
      notes: '此为固定演示内容，未根据上传资料调用模型。',
      sources: [],
    },
    {
      title: '从服务效率到协作效率',
      layout: 'twoColumn',
      columns: [
        { heading: '本季度进展', bullets: ['客户响应流程优化', '新版入门指南上线'] },
        { heading: '下季度重点', bullets: ['完善跨团队协作', '持续跟踪客户反馈'] },
      ],
      notes: '固定演示内容。',
      sources: [],
    },
    {
      title: '三步落实行动计划',
      layout: 'process',
      steps: [
        { title: '梳理问题', detail: '汇总客户与团队反馈' },
        { title: '确定责任', detail: '分解任务与验收指标' },
        { title: '复盘改进', detail: '按周期检查实际效果' },
      ],
      notes: '固定演示内容。',
      sources: [],
    },
  ];
  return {
    title: '季度业务回顾 · 演示示例',
    slides: Array.from({ length: count }, (_, i) =>
      structuredClone(
        i === count - 1
          ? ({
              title: '下一步行动',
              layout: 'closing',
              subtitle: '把讨论结果转化为具体任务',
              notes: '固定演示内容。',
              sources: [],
            } as DeckSlide)
          : pages[i % pages.length],
      ),
    ),
  };
}
export function DeckGenerator(runtime: PresentationRuntime) {
  const { context, access, busy, run, request, notify, modelStamp } = runtime;
  const [brief, setBrief] = useState(''),
    [countInput, setCountInput] = useState('8'),
    [ratio, setRatio] = useState('wide'),
    [theme, setTheme] = useState(DEFAULT_THEME);
  const [draft, setDraft] = useState<DeckContent>(),
    [usedContext, setUsedContext] = useState<MaterialContext>(),
    [stamp, setStamp] = useState(''),
    [page, setPage] = useState(0),
    [confirmAppend, setConfirmAppend] = useState(false),
    [appended, setAppended] = useState(false);
  const count = Number(countInput),
    validCount = Number.isInteger(count) && count >= 2 && count <= 20;
  const fingerprint = JSON.stringify([
    brief,
    count,
    context.materials.map((m) => [m.id, m.name, m.text, m.image?.modelData]),
    context.includeImages,
    modelStamp,
  ]);
  const outdated = Boolean(draft && stamp !== fingerprint);
  function updatePage(patch: Partial<DeckSlide>) {
    setDraft((d) =>
      d ? { ...d, slides: d.slides.map((s, i) => (i === page ? { ...s, ...patch } : s)) } : d,
    );
    setAppended(false);
    setConfirmAppend(false);
  }
  async function generate() {
    await run('根据素材生成整套 PPT', async (signal) => {
      const messages = deckMessages(context, brief, count),
        raw = await request(messages, signal, JSON.stringify(demoDeck(count)));
      if (signal.aborted) return;
      const result = parseDeck(raw, context, count);
      setDraft(result);
      setUsedContext(context);
      setStamp(fingerprint);
      setPage(0);
      setAppended(false);
      setConfirmAppend(false);
      notify(`已生成 ${result.slides.length} 页。请预览内容与来源，再打开或下载 PPT。`);
    });
  }
  async function exportTo(target: 'download' | 'open' | 'append') {
    if (!draft || !usedContext || outdated) return;
    await run(target === 'append' ? '插入完整页面' : '构建可编辑 PPT', async (signal) => {
      // Match the current deck when appending; never silently assume a custom slide size.
      const snapshot =
        target === 'append' && access
          ? await access.read(access.capabilities.selected ? 'selected' : 'all')
          : undefined;
      if (target === 'append' && !snapshot?.size)
        throw new Error('无法确认当前页面尺寸，请下载 PPT 后手动导入。');
      const base64 = await renderDeck(
        draft,
        usedContext,
        theme,
        snapshot?.size || (ratio === 'wide' ? WIDE : STANDARD),
      );
      if (signal.aborted) return;
      if (target === 'download') {
        downloadPresentation(base64, draft.title);
        notify('已生成下载文件，请在 PowerPoint 中核对排版和数据。');
      }
      if (target === 'open' && access) {
        await access.create(base64);
        notify('已请求 PowerPoint 打开新演示稿，请查看新窗口并保存。');
      }
      if (target === 'append' && access) {
        // A failed Office response can follow a partial write. Require a new preview
        // instead of allowing a second click to duplicate pages blindly.
        setAppended(true);
        setConfirmAppend(false);
        await access.insert(base64, draft.slides.length, snapshot);
        notify(
          `已在当前演示稿末尾插入 ${draft.slides.length} 页，请检查页面。可使用 Office 撤销。`,
        );
      }
    });
  }
  const current = draft?.slides[page];
  return (
    <div className="ppt-flow">
      <section className="ppt-compose" aria-label="整稿生成要求">
        <h2>从素材生成整套 PPT</h2>
        <p className="field-help">
          说明受众、目的和重点。模型规划内容与版式，生成可编辑文字、表格、图表和图片页面。
        </p>
        <label className="field-label" htmlFor="deck-brief">
          演示要求
        </label>
        <textarea
          id="deck-brief"
          value={brief}
          maxLength={6000}
          disabled={busy}
          rows={5}
          placeholder="例如：给管理层做 8 页季度经营汇报。根据目录中的报告与销售表，突出增长、风险和下季度行动；用图表呈现月度趋势。"
          onChange={(e) => setBrief(e.target.value)}
        />
        <div className="ppt-field-grid">
          <label>
            页数
            <input
              aria-label="生成页数"
              type="number"
              min={2}
              max={20}
              value={countInput}
              disabled={busy}
              onChange={(e) => setCountInput(e.target.value)}
            />
          </label>
          <label>
            页面比例
            <select value={ratio} disabled={busy} onChange={(e) => setRatio(e.target.value)}>
              <option value="wide">宽屏 16:9</option>
              <option value="standard">标准 4:3</option>
            </select>
          </label>
        </div>
        {!validCount && <p className="notice warning">页数请输入 2–20 之间的整数。</p>}
        <details className="ppt-options">
          <summary>字体与主题</summary>
          <div className="ppt-field-grid">
            <label>
              字体
              <select
                value={theme.font}
                disabled={busy}
                onChange={(e) => setTheme({ ...theme, font: e.target.value })}
              >
                <option>Aptos</option>
                <option>Arial</option>
                <option>等线</option>
              </select>
            </label>
            <label>
              主题色
              <input
                type="color"
                value={theme.accent}
                disabled={busy}
                onChange={(e) => setTheme({ ...theme, accent: e.target.value })}
              />
            </label>
          </div>
          <p className="field-help">使用本机 Office 可用字体；缺失字体会由 Office 替换。</p>
        </details>
        <p className="field-help">
          将使用 {context.materials.length} 份已选素材
          {context.includeImages ? '，包含图片预览' : ''}。生成后可逐页检查来源并修改内容。
        </p>
        <button
          type="button"
          className="primary"
          disabled={busy || !brief.trim() || !validCount}
          onClick={() => void generate()}
        >
          <Sparkles size={17} />
          {draft ? '重新生成整套 PPT' : '生成整套 PPT'}
        </button>
      </section>
      {draft && current && usedContext && (
        <section className="ppt-result" aria-label="整稿预览">
          <div className="section-heading">
            <h2>整稿预览</h2>
            <span>{draft.slides.length} 页</span>
          </div>
          {outdated && (
            <p className="notice warning" role="status">
              素材、要求或模型已变化。请重新生成后再导出，当前预览仍保留。
            </p>
          )}
          <label className="field-label">
            文件标题
            <input
              value={draft.title}
              maxLength={100}
              disabled={busy}
              onChange={(e) => setDraft({ ...draft, title: e.target.value })}
            />
          </label>
          <div className="page-picker">
            <label>
              预览页面
              <select
                value={page}
                disabled={busy}
                onChange={(e) => setPage(Number(e.target.value))}
              >
                {draft.slides.map((s, i) => (
                  <option key={i} value={i}>
                    {i + 1}. {s.title}
                  </option>
                ))}
              </select>
            </label>
            <div className="button-row">
              {[-1, 1].map((direction) => (
                <button
                  key={direction}
                  className="icon-button"
                  aria-label={direction < 0 ? '页面前移' : '页面后移'}
                  disabled={busy || page + direction < 0 || page + direction >= draft.slides.length}
                  onClick={() => {
                    const slides = [...draft.slides];
                    [slides[page], slides[page + direction]] = [
                      slides[page + direction],
                      slides[page],
                    ];
                    setDraft({ ...draft, slides });
                    setPage(page + direction);
                    setAppended(false);
                  }}
                >
                  {direction < 0 ? <ArrowUp size={16} /> : <ArrowDown size={16} />}
                </button>
              ))}
              <button
                className="icon-button"
                aria-label="移除此页"
                disabled={busy || draft.slides.length <= 1}
                onClick={() => {
                  setDraft({ ...draft, slides: draft.slides.filter((_, i) => i !== page) });
                  setPage(Math.max(0, page - 1));
                  setAppended(false);
                }}
              >
                <Trash2 size={16} />
              </button>
            </div>
          </div>
          <SlidePreview page={current} context={usedContext} />
          <p className="field-help">
            内容与版式概览；最终排版请在 PowerPoint 中核对。来源：
            {current.sources
              .map((id) => usedContext.materials.find((m) => m.id === id)?.name)
              .join('、') || '未引用素材，需核实事实'}
          </p>
          <details className="ppt-page-editor" key={page}>
            <summary>编辑本页文字与备注</summary>
            <label>
              页面标题
              <input
                value={current.title}
                maxLength={64}
                disabled={busy}
                onChange={(e) => updatePage({ title: e.target.value })}
              />
            </label>
            {current.subtitle !== undefined && (
              <label>
                副标题
                <textarea
                  value={current.subtitle}
                  maxLength={180}
                  disabled={busy}
                  rows={2}
                  onChange={(e) => updatePage({ subtitle: e.target.value })}
                />
              </label>
            )}
            {current.bullets && (
              <label>
                要点（每行一条）
                <textarea
                  value={current.bullets.join('\n')}
                  disabled={busy}
                  rows={5}
                  onChange={(e) => updatePage({ bullets: e.target.value.split('\n') })}
                />
              </label>
            )}
            {current.columns?.map((c, i) => (
              <div key={i}>
                <label>
                  第 {i + 1} 栏标题
                  <input
                    value={c.heading}
                    maxLength={32}
                    disabled={busy}
                    onChange={(e) =>
                      updatePage({
                        columns: current.columns!.map((v, n) =>
                          n === i ? { ...v, heading: e.target.value } : v,
                        ),
                      })
                    }
                  />
                </label>
                <label>
                  栏要点（每行一条）
                  <textarea
                    value={c.bullets.join('\n')}
                    disabled={busy}
                    rows={3}
                    onChange={(e) =>
                      updatePage({
                        columns: current.columns!.map((v, n) =>
                          n === i ? { ...v, bullets: e.target.value.split('\n') } : v,
                        ),
                      })
                    }
                  />
                </label>
              </div>
            ))}
            {current.metrics?.map((m, i) => (
              <div className="ppt-field-grid" key={i}>
                {(['label', 'value', 'detail'] as const).map((k) => (
                  <label key={k}>
                    {k === 'label' ? '指标名称' : k === 'value' ? '指标值' : '说明'}
                    <input
                      value={m[k]}
                      disabled={busy}
                      maxLength={k === 'value' ? 16 : k === 'label' ? 24 : 80}
                      onChange={(e) =>
                        updatePage({
                          metrics: current.metrics!.map((v, n) =>
                            n === i ? { ...v, [k]: e.target.value } : v,
                          ),
                        })
                      }
                    />
                  </label>
                ))}
              </div>
            ))}
            {current.steps?.map((s, i) => (
              <div key={i}>
                <label>
                  步骤 {i + 1}
                  <input
                    value={s.title}
                    maxLength={24}
                    disabled={busy}
                    onChange={(e) =>
                      updatePage({
                        steps: current.steps!.map((v, n) =>
                          n === i ? { ...v, title: e.target.value } : v,
                        ),
                      })
                    }
                  />
                </label>
                <label>
                  步骤说明
                  <input
                    value={s.detail}
                    maxLength={80}
                    disabled={busy}
                    onChange={(e) =>
                      updatePage({
                        steps: current.steps!.map((v, n) =>
                          n === i ? { ...v, detail: e.target.value } : v,
                        ),
                      })
                    }
                  />
                </label>
              </div>
            ))}
            <label>
              演讲备注
              <textarea
                value={current.notes}
                maxLength={2500}
                disabled={busy}
                rows={4}
                onChange={(e) => updatePage({ notes: e.target.value })}
              />
            </label>
          </details>
          <div className="button-row ppt-export">
            <button
              className="primary"
              disabled={busy || outdated}
              onClick={() => void exportTo('download')}
            >
              <Download size={17} />
              下载完整 PPT
            </button>
            {access?.capabilities.create && (
              <button
                className="secondary"
                disabled={busy || outdated}
                onClick={() => void exportTo('open')}
              >
                <Presentation size={17} />在 PowerPoint 中打开
              </button>
            )}
            {access?.capabilities.insert && (
              <button
                className="secondary"
                disabled={busy || outdated || appended}
                onClick={() => setConfirmAppend(true)}
              >
                {appended ? '已尝试插入，请检查末尾' : '插入当前演示稿末尾'}
              </button>
            )}
          </div>
          {confirmAppend && !outdated && (
            <div className="ppt-confirm">
              <p>
                将在当前演示稿末尾新增 {draft.slides.length} 页，按当前页面尺寸生成。请先保存原稿。
              </p>
              <div className="button-row">
                <button className="primary" disabled={busy} onClick={() => void exportTo('append')}>
                  确认插入 {draft.slides.length} 页
                </button>
                <button
                  className="secondary"
                  disabled={busy}
                  onClick={() => setConfirmAppend(false)}
                >
                  返回预览
                </button>
              </div>
            </div>
          )}
        </section>
      )}
    </div>
  );
}
