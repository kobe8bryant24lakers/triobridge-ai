import { englishCount } from './presentation-runtime';
import { t } from '../i18n';
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
import {
  appendEditingTurn,
  buildEditingFollowUp,
  type EditingTurn,
} from '../core/editing-conversation';
import { EditingConversation } from './EditingConversation';

function demoDeck(count: number): DeckContent {
  const pages: DeckSlide[] = [
    {
      title: t('季度业务回顾 · 演示示例', 'Quarterly business review · Demo'),
      layout: 'cover',
      subtitle: t('固定示例，用于体验整稿生成流程', 'A sample deck to explore the workflow'),
      notes: t(
        '此为固定演示内容，未根据上传资料调用模型。',
        'This fixed example uses no model and does not use the uploaded materials.',
      ),
      sources: [],
    },
    {
      title: t('从服务效率到协作效率', 'Service and team performance'),
      layout: 'twoColumn',
      columns: [
        {
          heading: t('本季度进展', 'Progress this quarter'),
          bullets: [
            t('客户响应流程优化', 'Improved customer response times'),
            t('新版入门指南上线', 'Released the new onboarding guide'),
          ],
        },
        {
          heading: t('下季度重点', 'Next quarter'),
          bullets: [
            t('完善跨团队协作', 'Improve collaboration between teams'),
            t('持续跟踪客户反馈', 'Continue tracking customer feedback'),
          ],
        },
      ],
      notes: t('固定演示内容。', 'Fixed demo content.'),
      sources: [],
    },
    {
      title: t('三步落实行动计划', 'Three steps to put the plan into action'),
      layout: 'process',
      steps: [
        {
          title: t('梳理问题', 'Identify the issues'),
          detail: t('汇总客户与团队反馈', 'Gather feedback from customers and teams'),
        },
        {
          title: t('确定责任', 'Assign responsibility'),
          detail: t('分解任务与验收指标', 'Define tasks and acceptance criteria'),
        },
        {
          title: t('复盘改进', 'Review and improve'),
          detail: t('按周期检查实际效果', 'Check results regularly'),
        },
      ],
      notes: t('固定演示内容。', 'Fixed demo content.'),
      sources: [],
    },
  ];
  return {
    title: t('季度业务回顾 · 演示示例', 'Quarterly business review · Demo'),
    slides: Array.from({ length: count }, (_, i) =>
      structuredClone(
        i === count - 1
          ? ({
              title: t('下一步行动', 'Next steps'),
              layout: 'closing',
              subtitle: t('把讨论结果转化为具体任务', 'Turn the discussion into specific actions'),
              notes: t('固定演示内容。', 'Fixed demo content.'),
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
  const [turns, setTurns] = useState<EditingTurn[]>([]),
    [followUp, setFollowUp] = useState('');
  const count = Number(countInput),
    validCount = Number.isInteger(count) && count >= 2 && count <= 20;
  const fingerprint = JSON.stringify([
    brief,
    count,
    context.materials.map((m) => [m.id, m.name, m.text, m.tables, m.image?.modelData]),
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
    await run(
      t('根据素材生成整套 PPT', 'Generating a deck from files'),
      async (signal) => {
        const messages = deckMessages(context, brief, count),
          raw = await request(
            messages,
            signal,
            JSON.stringify(demoDeck(count)),
            'presentation-deck',
          );
        if (signal.aborted) return;
        const result = parseDeck(raw, context, count);
        setDraft(result);
        setUsedContext(context);
        setStamp(fingerprint);
        setPage(0);
        setAppended(false);
        setConfirmAppend(false);
        setTurns([{ instruction: brief, result: JSON.stringify(result) }]);
        setFollowUp('');
        notify(
          t(
            `已生成 ${result.slides.length} 页。请预览内容与来源，再打开或下载 PPT。`,
            `Generated ${result.slides.length} slides. Review their content and sources before opening or downloading the deck.`,
          ),
        );
      },
      true,
    );
  }
  async function refine() {
    if (!draft || !usedContext || outdated || busy || !followUp.trim()) return;
    const instruction = followUp.trim();
    await run(
      t('根据对话调整整套 PPT', 'Refining the deck'),
      async (signal) => {
        const current = JSON.stringify(draft);
        const messages = buildEditingFollowUp(
          deckMessages(usedContext, brief, draft.slides.length),
          turns,
          current,
          instruction,
        );
        const raw = await request(messages, signal, current, 'presentation-deck');
        if (signal.aborted) return;
        const result = parseDeck(raw, usedContext, draft.slides.length);
        setDraft(result);
        setTurns((previous) =>
          appendEditingTurn(previous, current, instruction, JSON.stringify(result)),
        );
        setFollowUp('');
        setAppended(false);
        setConfirmAppend(false);
        notify(
          runtime.demo
            ? t(
                '固定演示保留当前预览；连接真实模型后可按对话调整。',
                'The demo keeps the current preview. Connect a model to refine it through conversation.',
              )
            : t(
                '已更新整稿预览，请检查后再导出或插入。',
                'Deck preview updated. Review it before exporting or inserting.',
              ),
        );
      },
      true,
    );
  }
  async function exportTo(target: 'download' | 'open' | 'append') {
    if (!draft || !usedContext || outdated) return;
    await run(
      target === 'append'
        ? t('插入完整页面', 'Inserting slides')
        : t('构建可编辑 PPT', 'Building an editable deck'),
      async (signal) => {
        // Match the current deck when appending; never silently assume a custom slide size.
        const insertTarget =
          target === 'append' && access ? await access.prepareInsert() : undefined;
        if (target === 'append' && !insertTarget?.size)
          throw new Error(
            t(
              '无法确认当前页面尺寸，请下载 PPT 后手动导入。',
              'The slide size could not be confirmed. Download the deck and import it in PowerPoint.',
            ),
          );
        const base64 = await renderDeck(
          draft,
          usedContext,
          theme,
          insertTarget?.size || (ratio === 'wide' ? WIDE : STANDARD),
        );
        if (signal.aborted) return;
        if (target === 'download') {
          downloadPresentation(base64, draft.title);
          notify(
            t(
              '已生成下载文件，请在 PowerPoint 中核对排版和数据。',
              'The deck is ready to download. Check its layout and data in PowerPoint.',
            ),
          );
        }
        if (target === 'open' && access) {
          await access.create(base64);
          notify(
            t(
              '已请求 PowerPoint 打开新演示稿，请查看新窗口并保存。',
              'A new presentation was requested. Check the new PowerPoint window and save the file.',
            ),
          );
        }
        if (target === 'append' && access && insertTarget) {
          // A failed Office response can follow a partial write. Require a new preview
          // instead of allowing a second click to duplicate pages blindly.
          setAppended(true);
          setConfirmAppend(false);
          await access.insert(base64, draft.slides.length, insertTarget);
          notify(
            t(
              `已在当前演示稿末尾插入 ${draft.slides.length} 页，请检查页面。可使用 Office 撤销。`,
              `Appended ${englishCount(draft.slides.length, 'slide')}. Check the end of the deck. You can undo the insertion in Office.`,
            ),
          );
        }
      },
    );
  }
  const current = draft?.slides[page];
  return (
    <div className="ppt-flow">
      <section className="ppt-compose" aria-label={t('整稿生成要求', 'Presentation brief')}>
        <h2>{t('从素材生成整套 PPT', 'Build a deck from files')}</h2>
        <p className="field-help">
          {t(
            '说明受众、目的和重点。模型规划内容与版式，生成可编辑文字、表格、图表和图片页面。',
            'Describe the audience, purpose and key points. The model plans the slides and creates editable text, tables, charts and images.',
          )}
        </p>
        <label className="field-label" htmlFor="deck-brief">
          {t('演示要求', 'Presentation brief')}
        </label>
        <textarea
          id="deck-brief"
          value={brief}
          maxLength={6000}
          disabled={busy}
          rows={5}
          placeholder={t(
            '例如：给管理层做 8 页季度经营汇报。根据目录中的报告与销售表，突出增长、风险和下季度行动；用图表呈现月度趋势。',
            'Example: Create an eight-slide quarterly review for management. Use the report and sales figures to show growth, risks and next steps.',
          )}
          onChange={(e) => setBrief(e.target.value)}
        />
        <div className="ppt-field-grid">
          <label>
            {t('页数', 'Slide count')}
            <input
              aria-label={t('生成页数', 'Number of slides')}
              type="number"
              min={2}
              max={20}
              value={countInput}
              disabled={busy}
              onChange={(e) => setCountInput(e.target.value)}
            />
          </label>
          <label>
            {t('页面比例', 'Aspect ratio')}
            <select value={ratio} disabled={busy} onChange={(e) => setRatio(e.target.value)}>
              <option value="wide">{t('宽屏 16:9', 'Widescreen 16:9')}</option>
              <option value="standard">{t('标准 4:3', 'Standard 4:3')}</option>
            </select>
          </label>
        </div>
        {!validCount && (
          <p className="notice warning">
            {t('页数请输入 2–20 之间的整数。', 'Enter a whole number from 2 to 20.')}
          </p>
        )}
        <details className="ppt-options">
          <summary>{t('字体与主题', 'Font and theme')}</summary>
          <div className="ppt-field-grid">
            <label>
              {t('字体', 'Font')}
              <select
                value={theme.font}
                disabled={busy}
                onChange={(e) => setTheme({ ...theme, font: e.target.value })}
              >
                <option>Aptos</option>
                <option>Arial</option>
                <option value="等线">{t('等线', 'DengXian')}</option>
              </select>
            </label>
            <label>
              {t('主题色', 'Accent color')}
              <input
                type="color"
                value={theme.accent}
                disabled={busy}
                onChange={(e) => setTheme({ ...theme, accent: e.target.value })}
              />
            </label>
          </div>
          <p className="field-help">
            {t(
              '使用本机 Office 可用字体；缺失字体会由 Office 替换。',
              'Office substitutes fonts that are unavailable on this computer.',
            )}
          </p>
        </details>
        <p className="field-help">
          {t('将使用 ', 'Using ')}
          {context.materials.length}
          {t(' 份已选素材', context.materials.length === 1 ? ' selected file' : ' selected files')}
          {context.includeImages ? t('，包含图片预览', ', including image previews') : ''}
          {t(
            '。生成后可逐页检查来源并修改内容。',
            '. Review the sources and edit each slide after generation.',
          )}
        </p>
        <button
          type="button"
          className="primary"
          disabled={busy || !brief.trim() || !validCount}
          onClick={() => void generate()}
        >
          <Sparkles size={17} />
          {draft ? t('重新生成整套 PPT', 'Regenerate deck') : t('生成整套 PPT', 'Generate deck')}
        </button>
      </section>
      {draft && current && usedContext && (
        <section className="ppt-result" aria-label={t('整稿预览', 'Deck preview')}>
          <div className="section-heading">
            <h2>{t('整稿预览', 'Deck preview')}</h2>
            <span>
              {t(`${draft.slides.length} 页`, englishCount(draft.slides.length, 'slide'))}
            </span>
          </div>
          {outdated && (
            <p className="notice warning" role="status">
              {t(
                '素材、要求或模型已变化。请重新生成后再导出，当前预览仍保留。',
                'Sources, instructions or model changed. Regenerate before exporting. The current preview is still available.',
              )}
            </p>
          )}
          <label className="field-label">
            {t('文件标题', 'File title')}
            <input
              value={draft.title}
              maxLength={100}
              disabled={busy}
              onChange={(e) => setDraft({ ...draft, title: e.target.value })}
            />
          </label>
          <div className="page-picker">
            <label>
              {t('预览页面', 'Preview slide')}
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
                  aria-label={
                    direction < 0
                      ? t('页面前移', 'Move slide up')
                      : t('页面后移', 'Move slide down')
                  }
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
                aria-label={t('移除此页', 'Remove slide')}
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
            {t(
              '内容与版式概览；最终排版请在 PowerPoint 中核对。来源：',
              'Check the final layout in PowerPoint. Sources: ',
            )}
            {current.sources
              .map((id) => usedContext.materials.find((m) => m.id === id)?.name)
              .join(t('、', ', ')) ||
              t('未引用素材，需核实事实', 'No sources cited. Verify the facts.')}
          </p>
          <details className="ppt-page-editor" key={page}>
            <summary>{t('编辑本页文字与备注', 'Edit slide text and notes')}</summary>
            <label>
              {t('页面标题', 'Slide title')}
              <input
                value={current.title}
                maxLength={64}
                disabled={busy}
                onChange={(e) => updatePage({ title: e.target.value })}
              />
            </label>
            {current.subtitle !== undefined && (
              <label>
                {t('副标题', 'Subtitle')}
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
                {t('要点（每行一条）', 'Bullets (one per line)')}
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
                  {t(`第 ${i + 1} 栏标题`, `Column ${i + 1} heading`)}
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
                  {t('栏要点（每行一条）', 'Column bullets (one per line)')}
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
                    {k === 'label'
                      ? t('指标名称', 'Metric name')
                      : k === 'value'
                        ? t('指标值', 'Value')
                        : t('说明', 'Description')}
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
                  {t('步骤 ', 'Step ')}
                  {i + 1}
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
                  {t('步骤说明', 'Step details')}
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
              {t('演讲备注', 'Speaker notes')}
              <textarea
                value={current.notes}
                maxLength={2500}
                disabled={busy}
                rows={4}
                onChange={(e) => updatePage({ notes: e.target.value })}
              />
            </label>
          </details>
          <EditingConversation
            id="deck-conversation"
            turns={turns}
            value={followUp}
            onChange={setFollowUp}
            onSubmit={() => void refine()}
            onReset={() => {
              setTurns([]);
              setFollowUp('');
            }}
            disabled={busy}
            stale={outdated}
          />
          <div className="button-row ppt-export">
            <button
              className="primary"
              disabled={busy || outdated}
              onClick={() => void exportTo('download')}
            >
              <Download size={17} />
              {t('下载完整 PPT', 'Download full deck')}
            </button>
            {access?.capabilities.create && (
              <button
                className="secondary"
                disabled={busy || outdated}
                onClick={() => void exportTo('open')}
              >
                <Presentation size={17} />
                {t('在 PowerPoint 中打开', 'Open in PowerPoint')}
              </button>
            )}
            {access?.capabilities.insert && (
              <button
                className="secondary"
                disabled={busy || outdated || appended}
                onClick={() => setConfirmAppend(true)}
              >
                {appended
                  ? t('已尝试插入，请检查末尾', 'Insertion attempted. Check the last slide.')
                  : t('插入当前演示稿末尾', 'Append to current deck')}
              </button>
            )}
          </div>
          {confirmAppend && !outdated && (
            <div className="ppt-confirm">
              <p>
                {t(
                  `将在当前演示稿末尾新增 ${draft.slides.length} 页，按当前页面尺寸生成。请先保存原稿。`,
                  `Append ${englishCount(draft.slides.length, 'slide')} using this deck’s slide size. Save the original first.`,
                )}
              </p>
              <div className="button-row">
                <button className="primary" disabled={busy} onClick={() => void exportTo('append')}>
                  {t('确认插入 ', 'Confirm insertion of ')}
                  {t(`${draft.slides.length} 页`, englishCount(draft.slides.length, 'slide'))}
                </button>
                <button
                  className="secondary"
                  disabled={busy}
                  onClick={() => setConfirmAppend(false)}
                >
                  {t('返回预览', 'Back to preview')}
                </button>
              </div>
            </div>
          )}
        </section>
      )}
    </div>
  );
}
