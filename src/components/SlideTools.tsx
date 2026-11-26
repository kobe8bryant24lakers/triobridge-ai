import { useState } from 'react';
import { t } from '../i18n';
import type {
  BackgroundFill,
  ChartData,
  ChartKind,
  EditProposal,
  MaterialContext,
  PageEdit,
  PageSize,
  PresentationSnapshot,
  SlideBackground,
  SlideState,
} from '../presentation/types';
import { backgroundLabel, pagesForRange } from '../presentation/page-edits';
import {
  applyRecipe,
  sampleStyle,
  slideIssues,
  splitPoint,
  type EditRecipe,
  type SampleStyle,
} from '../presentation/slide-tools';
import { CHART_LABELS, chartData, tableFromText } from '../presentation/charts';
import { allTables } from '../presentation/content';
import { BackgroundPicker, fillStyle, type PickBackgroundPicture } from './SlideBackground';

const recipeKey = 'triobridge.ppt.recipes.v1';
function savedRecipes(): EditRecipe[] {
  try {
    const v = JSON.parse(localStorage.getItem(recipeKey) || '[]');
    return Array.isArray(v) ? v.filter((r) => r && typeof r.name === 'string').slice(0, 50) : [];
  } catch {
    return [];
  }
}
interface Props {
  pickPicture: PickBackgroundPicture;
  snapshot: PresentationSnapshot;
  context: MaterialContext;
  disabled: boolean;
  full: boolean;
  canImage: boolean;
  canGradient: boolean;
  demo: boolean;
  stage: (make: () => EditProposal) => Promise<void>;
  notice: (message: string) => void;
  focus: (slide: SlideState, shapeId?: string) => void;
  comment: (slide: SlideState, text: string, shapeId?: string) => void;
}
export function SlideTools({
  pickPicture,
  snapshot,
  context,
  disabled,
  full,
  canImage,
  canGradient,
  demo,
  stage,
  notice,
  focus,
  comment,
}: Props) {
  const slides = snapshot.slides;
  const [pages, setPages] = useState(''),
    [exclude, setExclude] = useState(''),
    [fill, setFill] = useState<BackgroundFill>({ kind: 'solid', color: '#F5F1E8' });
  const [useFill, setUseFill] = useState(true),
    [useTitle, setUseTitle] = useState(false),
    [color, setColor] = useState('#17365D'),
    [find, setFind] = useState(''),
    [replacement, setReplacement] = useState('');
  const [name, setName] = useState(''),
    [recipes, setRecipes] = useState(savedRecipes);
  const [reference, setReference] = useState(slides[0].id),
    [style, setStyle] = useState<SampleStyle>({
      background: true,
      title: true,
      body: true,
      alignment: false,
      geometry: false,
    });
  const [anchor, setAnchor] = useState(slides[0].id),
    [operation, setOperation] = useState<'insert' | 'split' | 'chart'>('insert'),
    [title, setTitle] = useState(''),
    [body, setBody] = useState(''),
    [object, setObject] = useState(''),
    [at, setAt] = useState(''),
    [data, setData] = useState(''),
    [tableId, setTableId] = useState(''),
    [withChart, setWithChart] = useState(false),
    [kind, setKind] = useState<ChartKind>('column'),
    [unit, setUnit] = useState('');
  const [handled, setHandled] = useState<Set<string>>(new Set());
  const [located, setLocated] = useState<{ slide: SlideState; id?: string }>();
  const targets = () => pagesForRange(slides, pages, exclude);
  const recipe = (): EditRecipe => ({
    name: name.trim() || t('页面批量编辑', 'Batch slide edit'),
    ...(useFill ? { fill } : {}),
    ...(useTitle ? { titleColor: color } : {}),
    ...(find ? { find, replacement } : {}),
  });
  function save() {
    if (!name.trim()) return;
    const entry = recipe();
    if (entry.fill?.kind === 'image') {
      const id = entry.fill.imageId;
      entry.pictureName = context.materials.find((m) => m.id === id)?.name;
    }
    const next = [entry, ...recipes.filter((r) => r.name !== name.trim())].slice(0, 50);
    try {
      localStorage.setItem(recipeKey, JSON.stringify(next));
      setRecipes(next);
      notice(
        t(
          '配方已保存，可在其他演示稿中调整参数后重用。',
          'Recipe saved. You can reuse it with different parameters in another deck.',
        ),
      );
    } catch {
      notice(t('此入口无法保存本地配方。', 'Local recipe storage is unavailable here.'));
    }
  }
  function restore(r: EditRecipe) {
    setName(r.name);
    setUseFill(!!r.fill);
    if (r.fill?.kind === 'image') {
      const matching = context.materials.filter((m) => m.image && m.name === r.pictureName);
      setFill({ ...r.fill, imageId: matching.length === 1 ? matching[0].id : '' });
    } else if (r.fill) setFill(r.fill);
    setUseTitle(!!r.titleColor);
    if (r.titleColor) setColor(r.titleColor);
    setFind(r.find || '');
    setReplacement(r.replacement || '');
  }
  const anchorSlide = slides.find((s) => s.id === anchor) || slides[0],
    ref = slides.find((s) => s.id === reference) || slides[0];
  const objects = anchorSlide.shapes.filter((s) =>
    operation === 'chart' ? s.chart : s.role === 'body' && s.text && s.text.trim().length > 1,
  );
  const shape = objects.find((s) => s.id === object) || objects[0];
  const tables = allTables(context);
  function source(kind: ChartKind) {
    const table = tables.find((t) => t.id === tableId) || tableFromText(data);
    return chartData(
      table,
      0,
      table.headers.slice(1).map((_, i) => i + 1),
      kind,
    );
  }
  function structure(): EditProposal {
    const slideId = anchorSlide.id;
    let action: PageEdit;
    if (operation === 'chart') {
      if (!shape?.chart)
        throw Error(
          t('该页没有支持刷新的原生图表。', 'This slide has no supported editable chart.'),
        );
      action = { kind: 'chart', slideId, shapeId: shape.id, data: source(shape.chart.kind) };
    } else if (operation === 'split') {
      if (!shape?.text)
        throw Error(t('该页没有可拆分的正文。', 'This slide has no body text to split.'));
      action = {
        kind: 'split',
        slideId,
        shapeId: shape.id,
        at: at ? Number(at) : splitPoint(shape.text),
        title: title.trim() || anchorSlide.title.slice(0, 145) + t('（续）', ' (continued)'),
      };
    } else {
      action = {
        kind: 'insert',
        slideId,
        title: title.trim(),
        body,
        ...(withChart ? { chart: { kind, data: source(kind), unit } } : {}),
      };
    }
    return {
      summary: t(
        '修改第 ' + (anchorSlide.index + 1) + ' 页的内容或后续页面。',
        'Edit slide ' + (anchorSlide.index + 1) + ' or the page following it.',
      ),
      changes: [],
      actions: [action],
    };
  }
  const options = slides.map((s) => (
    <option value={s.id} key={s.id}>
      {s.index + 1} · {s.title.slice(0, 50)}
    </option>
  ));
  const findings = slideIssues(snapshot),
    [limit, setLimit] = useState(12);
  return (
    <div className="ppt-slide-tools">
      <details className="ppt-quick" open>
        <summary>{t('页面背景与批量配方', 'Slide backgrounds and recipes')}</summary>
        <fieldset disabled={disabled} className="ppt-tool-fields">
          <div className="ppt-field-grid">
            <label>
              {t('应用到页码', 'Apply to slides')}
              <input
                value={pages}
                onChange={(e) => setPages(e.target.value)}
                placeholder={t('已读取的全部页，或 1,3-8', 'All slides read, or 1,3-8')}
              />
            </label>
            <label>
              {t('排除页码', 'Exclude slides')}
              <input
                value={exclude}
                onChange={(e) => setExclude(e.target.value)}
                placeholder={t('例如 1,4', 'e.g. 1,4')}
              />
            </label>
          </div>
          {slides.some((s) => s.background) && (
            <>
              <label className="checkbox-label">
                <input
                  type="checkbox"
                  checked={useFill}
                  onChange={(e) => setUseFill(e.target.checked)}
                />
                {t('修改页面背景', 'Change slide background')}
              </label>
              {useFill && (
                <BackgroundPicker
                  pickPicture={pickPicture}
                  fill={fill}
                  onChange={(f) => f && setFill(f)}
                  context={context}
                  kinds={
                    full || demo
                      ? undefined
                      : [
                          'solid',
                          ...(canGradient ? ['gradient' as const] : []),
                          ...(canImage ? ['image' as const] : []),
                          'master',
                        ]
                  }
                />
              )}
            </>
          )}
          {!full && !demo && (
            <p className="field-help">
              {t(
                '打开 PPTX 工作副本可使用双色渐变、图片背景、拆页和原图表数据。',
                'Open a PPTX working copy for two-color gradients, picture backgrounds, slide splitting and original chart data.',
              )}
            </p>
          )}
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={useTitle}
              onChange={(e) => setUseTitle(e.target.checked)}
            />
            {t('同时修改标题颜色', 'Change title color')}
          </label>
          {useTitle && (
            <label>
              {t('配方标题色', 'Recipe title color')}
              <input type="color" value={color} onChange={(e) => setColor(e.target.value)} />
            </label>
          )}
          <details className="ppt-options">
            <summary>{t('替换用词与保存配方', 'Replace wording and save recipe')}</summary>
            <div className="ppt-field-grid">
              <label>
                {t('原用词', 'Original wording')}
                <input value={find} onChange={(e) => setFind(e.target.value)} maxLength={100} />
              </label>
              <label>
                {t('新用词', 'New wording')}
                <input
                  value={replacement}
                  onChange={(e) => setReplacement(e.target.value)}
                  maxLength={200}
                />
              </label>
            </div>
            {recipes.length > 0 && (
              <label>
                {t('载入已保存配方', 'Load saved recipe')}
                <select
                  value=""
                  onChange={(e) => {
                    const r = recipes.find((r) => r.name === e.target.value);
                    if (r) restore(r);
                  }}
                >
                  <option value="">{t('选择配方', 'Choose a recipe')}</option>
                  {recipes.map((r) => (
                    <option key={r.name}>{r.name}</option>
                  ))}
                </select>
              </label>
            )}
            <label>
              {t('配方名称', 'Recipe name')}
              <input value={name} maxLength={80} onChange={(e) => setName(e.target.value)} />
            </label>
            <button className="text-button" disabled={!name.trim()} onClick={save}>
              {t('保存配方', 'Save recipe')}
            </button>
          </details>
          <button
            className="secondary"
            onClick={() =>
              void stage(() =>
                applyRecipe(
                  {
                    ...recipe(),
                    ...(!slides.some((s) => s.background) ? { fill: undefined } : {}),
                  },
                  snapshot,
                  targets(),
                ),
              )
            }
          >
            {t('预览背景与批量改动', 'Preview background and batch changes')}
          </button>
        </fieldset>
      </details>
      <details className="ppt-quick">
        <summary>{t('套用参考页样式', 'Match a reference slide')}</summary>
        <fieldset disabled={disabled} className="ppt-tool-fields">
          <label>
            {t('参考页面', 'Reference slide')}
            <select value={ref.id} onChange={(e) => setReference(e.target.value)}>
              {options}
            </select>
          </label>
          <p className="field-help">
            {t(
              '目标沿用上方页码范围。同类文本框按名称或顺序匹配。',
              'Uses the slide range above. Text boxes match by role, then name or order.',
            )}
          </p>
          {(
            [
              ['background', t('背景', 'Background')],
              ['title', t('标题文字样式', 'Title formatting')],
              ['body', t('正文文字样式', 'Body formatting')],
              ['alignment', t('段落对齐', 'Paragraph alignment')],
              ['geometry', t('文本框位置与尺寸', 'Text box position and size')],
            ] as const
          ).map(([k, label]) => (
            <label className="checkbox-label" key={k}>
              <input
                type="checkbox"
                checked={style[k]}
                onChange={(e) => setStyle({ ...style, [k]: e.target.checked })}
              />
              {label}
            </label>
          ))}
          <button
            className="secondary"
            onClick={() => void stage(() => sampleStyle(snapshot, ref, targets(), style))}
          >
            {t('预览参考页样式', 'Preview reference styles')}
          </button>
        </fieldset>
      </details>
      {full && (
        <details className="ppt-quick">
          <summary>{t('插页、拆页与原图表数据', 'Insert, split and update charts')}</summary>
          <fieldset disabled={disabled} className="ppt-tool-fields">
            <label>
              {t('目标页面', 'Target slide')}
              <select
                value={anchorSlide.id}
                onChange={(e) => {
                  setAnchor(e.target.value);
                  setObject('');
                  setAt('');
                }}
              >
                {options}
              </select>
            </label>
            <label>
              {t('页面操作', 'Slide operation')}
              <select
                value={operation}
                onChange={(e) => {
                  setOperation(e.target.value as typeof operation);
                  setObject('');
                  setAt('');
                }}
              >
                <option value="insert">{t('在此页后插入', 'Insert after this slide')}</option>
                <option value="split">{t('拆分正文到下一页', 'Split body onto next slide')}</option>
                <option value="chart">{t('刷新原有图表', 'Update existing chart')}</option>
              </select>
            </label>
            {operation !== 'insert' && (
              <label>
                {t('目标对象', 'Target object')}
                <select
                  value={shape?.id || ''}
                  onChange={(e) => {
                    setObject(e.target.value);
                    setAt('');
                  }}
                >
                  {!objects.length && (
                    <option value="">{t('此页没有可用对象', 'No matching objects')}</option>
                  )}
                  {objects.map((s) => (
                    <option key={s.id} value={s.id}>
                      {s.name} · {(s.text || s.chart?.data.series[0]?.name || '').slice(0, 35)}
                    </option>
                  ))}
                </select>
              </label>
            )}
            {operation !== 'chart' && (
              <label>
                {t('新页标题', 'New slide title')}
                <input
                  maxLength={160}
                  value={title}
                  onChange={(e) => setTitle(e.target.value)}
                  placeholder={
                    operation === 'split' ? anchorSlide.title + t('（续）', ' (continued)') : ''
                  }
                />
              </label>
            )}
            {operation === 'insert' && (
              <>
                <label>
                  {t('新页正文', 'New slide body')}
                  <textarea
                    rows={3}
                    maxLength={4000}
                    value={body}
                    onChange={(e) => setBody(e.target.value)}
                  />
                </label>
                <label className="checkbox-label">
                  <input
                    type="checkbox"
                    checked={withChart}
                    onChange={(e) => setWithChart(e.target.checked)}
                  />
                  {t('插入可编辑图表', 'Include an editable chart')}
                </label>
              </>
            )}
            {operation === 'split' && shape?.text && (
              <label>
                {t('从第几个字符后拆开', 'Split after character')}
                <input
                  type="number"
                  min={1}
                  max={shape.text.length - 1}
                  value={at || splitPoint(shape.text)}
                  onChange={(e) => setAt(e.target.value)}
                />
              </label>
            )}
            {((operation === 'insert' && withChart) || operation === 'chart') && (
              <>
                {operation === 'insert' && (
                  <div className="ppt-field-grid">
                    <label>
                      {t('插入图表类型', 'Inserted chart type')}
                      <select value={kind} onChange={(e) => setKind(e.target.value as ChartKind)}>
                        {Object.entries(CHART_LABELS).map(([k, v]) => (
                          <option key={k} value={k}>
                            {v}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label>
                      {t('图表单位', 'Chart unit')}
                      <input
                        maxLength={24}
                        value={unit}
                        onChange={(e) => setUnit(e.target.value)}
                      />
                    </label>
                  </div>
                )}
                {tables.length > 0 && (
                  <label>
                    {t('数据来源', 'Data source')}
                    <select value={tableId} onChange={(e) => setTableId(e.target.value)}>
                      <option value="">{t('粘贴 Excel 数据', 'Paste Excel data')}</option>
                      {tables.map((t) => (
                        <option value={t.id} key={t.id}>
                          {t.name}
                        </option>
                      ))}
                    </select>
                  </label>
                )}
                {!tableId && (
                  <label>
                    {t(
                      '粘贴数据（含表头，首列为分类）',
                      'Paste data with headers and categories in the first column',
                    )}
                    <textarea
                      rows={4}
                      value={data}
                      onChange={(e) => setData(e.target.value)}
                      placeholder={'Quarter\tSales\nQ1\t20\nQ2\t25'}
                    />
                  </label>
                )}
              </>
            )}
            <button
              className="secondary"
              disabled={operation === 'insert' ? !title.trim() : !shape}
              onClick={() => void stage(structure)}
            >
              {t('预览页面操作', 'Preview slide operation')}
            </button>
          </fieldset>
        </details>
      )}
      <details className="ppt-quick">
        <summary>
          {t('批注与版面检查', 'Comments and layout review')} · {findings.length}
        </summary>
        {!findings.length && (
          <p className="field-help">
            {t(
              '当前范围没有可处理的批注或版面问题。',
              'No actionable comments or layout issues in this range.',
            )}
          </p>
        )}
        {!full && (
          <p className="field-help">
            {t(
              '在工作副本中还可读取 PowerPoint 批注。',
              'Open a working copy to include PowerPoint comments.',
            )}
          </p>
        )}
        <ul className="ppt-issues">
          {findings.slice(0, limit).map((f, i) => (
            <li key={i}>
              <strong>
                {f.slide.index + 1} · {f.shape?.name || f.slide.title}
              </strong>
              <p>{f.message}</p>
              {f.commentId && (
                <label className="checkbox-label">
                  <input
                    type="checkbox"
                    checked={handled.has(f.commentId)}
                    onChange={(e) => {
                      const next = new Set(handled);
                      if (e.target.checked) next.add(f.commentId!);
                      else next.delete(f.commentId!);
                      setHandled(next);
                    }}
                  />
                  {t('本次清单已处理', 'Handled in this review')}
                </label>
              )}
              <div className="button-row">
                <button
                  disabled={disabled}
                  className="text-button"
                  onClick={() => {
                    setLocated({ slide: f.slide, id: f.shape?.id });
                    focus(f.slide, f.shape?.id);
                  }}
                >
                  {t('定位', 'Locate')}
                </button>
                {(f.fix?.actions?.length ? full : !!f.fix) && (
                  <button
                    className="text-button"
                    disabled={disabled}
                    onClick={() => void stage(() => f.fix!)}
                  >
                    {t('预览修复', 'Preview fix')}
                  </button>
                )}
                {f.type === 'comment' && (
                  <button
                    className="text-button"
                    disabled={disabled}
                    onClick={() => comment(f.slide, f.message, f.shape?.id)}
                  >
                    {t('按批注生成方案', 'Propose a comment edit')}
                  </button>
                )}
              </div>
            </li>
          ))}
        </ul>
        {findings.length > limit && (
          <button className="text-button" onClick={() => setLimit(limit + 25)}>
            {t('显示更多检查项', 'Show more findings')}
          </button>
        )}
        {located && (
          <div className="ppt-located" role="status">
            <strong>
              {located.slide.index + 1} · {located.slide.title}
            </strong>
            {snapshot.size && (
              <SlideMap
                slide={located.slide}
                focus={located.id}
                size={snapshot.size}
                background={located.slide.effectiveBackground || located.slide.background}
                image={(id) => context.materials.find((m) => m.id === id)?.image?.data}
              />
            )}
            <p>
              {located.slide.shapes
                .filter((s) => (located.id ? s.id === located.id : s.role === 'body'))
                .map((s) => s.text || s.name)
                .join('\n')
                .slice(0, 600)}
            </p>
          </div>
        )}
      </details>
    </div>
  );
}
function DataGrid({ data }: { data: ChartData }) {
  return (
    <div className="ppt-data-scroll">
      <table>
        <thead>
          <tr>
            <th>{t('分类', 'Category')}</th>
            {data.series.map((s, j) => (
              <th key={j}>{s.name}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.categories.map((c, i) => (
            <tr key={i}>
              <th>{c}</th>
              {data.series.map((s, j) => (
                <td key={j}>{s.values[i] ?? '—'}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
function SlideMap({
  slide,
  size,
  background,
  image,
  focus,
}: {
  slide: SlideState;
  size: PageSize;
  background?: SlideBackground;
  image: (id: string) => string | undefined;
  focus?: string;
}) {
  return (
    <svg
      className="ppt-slide-map"
      viewBox={`0 0 ${size.width} ${size.height}`}
      aria-hidden="true"
      style={fillStyle(background, image)}
    >
      {slide.shapes.map(
        (s) =>
          s.width > 0 &&
          s.height > 0 && (
            <svg
              key={s.id}
              x={s.left}
              y={s.top}
              width={s.width}
              height={s.height}
              overflow="hidden"
            >
              {s.id === focus && (
                <rect
                  x="1"
                  y="1"
                  width={Math.max(0, s.width - 2)}
                  height={Math.max(0, s.height - 2)}
                  fill="none"
                  stroke="#28604A"
                  strokeWidth="4"
                />
              )}
              {s.text ? (
                s.text
                  .split('\n')
                  .slice(0, 6)
                  .map((line, i) => (
                    <text
                      key={i}
                      x="1"
                      y={(i + 1) * (s.font?.size || 16) * 1.12}
                      fontSize={s.font?.size || 16}
                      fontFamily={s.font?.name || 'sans-serif'}
                      fontWeight={s.font?.bold ? 700 : 400}
                      fill={s.font?.color || '#25362D'}
                    >
                      {line}
                    </text>
                  ))
              ) : (
                <rect width="100%" height="100%" fill="#85968B" opacity={s.chart ? 0.3 : 0.14} />
              )}
            </svg>
          ),
      )}
    </svg>
  );
}
export function PageChange({
  action,
  slide,
  slides,
  image,
  size,
  current,
}: {
  action: PageEdit;
  slide: SlideState;
  slides: SlideState[];
  size: PageSize | null;
  current?: SlideState;
  image: (id: string) => string | undefined;
}) {
  if (action.kind === 'background') {
    const fill = action.fill,
      ref = fill.kind === 'reference' ? slides.find((s) => s.id === fill.sourceId) : undefined;
    const after = fill.kind === 'reference' ? ref?.background : fill;
    const before =
      slide.effectiveBackground ||
      (slide.background?.kind === 'master' ? slide.masterBackground : slide.background);
    const afterStyle =
      ref?.effectiveBackground || (after?.kind === 'master' ? slide.masterBackground : after);
    const displayable = (v: SlideBackground | undefined) =>
      v && v.kind !== 'other' && v.kind !== 'master' && (v.kind !== 'image' || !!image(v.imageId));
    const unavailable = (
      <p className="field-help">
        {t('请在 PowerPoint 中查看此背景外观。', 'View this background in PowerPoint.')}
      </p>
    );
    const sample = (v: SlideBackground | undefined) => (
      <span className="ppt-fill-sample" style={fillStyle(v, image)} aria-hidden="true" />
    );
    return (
      <>
        <div className="ppt-fill-change">
          <span>
            {sample(before)}
            {backgroundLabel(slide.background)}
          </span>
          <span aria-hidden="true">→</span>
          <span>
            {sample(afterStyle)}
            {backgroundLabel(after)}
            {ref && ' · ' + t('参考页 ', 'reference ') + (ref.index + 1)}
          </span>
        </div>
        {size && (
          <details className="ppt-map-review">
            <summary>{t('对照页面布局', 'Compare slide layouts')}</summary>
            <div className="ppt-map-pair">
              <figure>
                {displayable(before) ? (
                  <SlideMap slide={slide} size={size} background={before} image={image} />
                ) : (
                  unavailable
                )}
                <figcaption>{t('修改前', 'Before')}</figcaption>
              </figure>
              <figure>
                {displayable(afterStyle) ? (
                  <SlideMap
                    slide={current || slide}
                    size={size}
                    background={afterStyle}
                    image={image}
                  />
                ) : (
                  unavailable
                )}
                <figcaption>{t('修改后（布局示意）', 'After (layout preview)')}</figcaption>
              </figure>
            </div>
          </details>
        )}
      </>
    );
  }
  if (action.kind === 'chart') {
    const old = slide.shapes.find((s) => s.id === action.shapeId)!.chart!.data;
    return (
      <div className="ppt-data-diff">
        <div>
          <strong>{t('原数据', 'Original data')}</strong>
          <DataGrid data={old} />
        </div>
        <div>
          <strong>{t('新数据', 'New data')}</strong>
          <DataGrid data={action.data} />
        </div>
      </div>
    );
  }
  if (action.kind === 'split') {
    const original = slide.shapes.find((s) => s.id === action.shapeId)!.text!;
    return (
      <div className="ppt-page-diff">
        <strong>{slide.title}</strong>
        <p>{original.slice(0, action.at)}</p>
        <strong>{action.title}</strong>
        <p>{original.slice(action.at)}</p>
      </div>
    );
  }
  return (
    <div className="ppt-page-diff">
      <strong>{action.title}</strong>
      <p>{action.body}</p>
      {action.chart && <DataGrid data={action.chart.data} />}
    </div>
  );
}
