import type { PresentationInsertTarget } from '../presentation/types';
import { t, getUiLanguage } from '../i18n';
import { useMemo, useState } from 'react';
import { Download, Sparkles } from 'lucide-react';
import { CHART_LABELS, chartData, tableFromText } from '../presentation/charts';
import { allTables } from '../presentation/content';
import { materialMessages } from '../presentation/materials';
import { fields, json, text } from '../presentation/validation';
import { DEFAULT_THEME, downloadPresentation, renderDeck, WIDE } from '../presentation/render';
import type { ChartKind, DeckContent, MaterialContext, MaterialTable } from '../presentation/types';
import { ChartPreview } from './PresentationPreview';
import type { PresentationRuntime } from './presentation-runtime';
import {
  appendEditingTurn,
  buildEditingFollowUp,
  type EditingTurn,
} from '../core/editing-conversation';
import { EditingConversation } from './EditingConversation';

export function DataChart({
  context,
  access,
  busy,
  run,
  request,
  notify,
  modelStamp,
  demo,
}: PresentationRuntime) {
  const tables = allTables(context),
    [source, setSource] = useState('paste'),
    [pasted, setPasted] = useState(''),
    [customTitle, setTitle] = useState<string>(),
    [kind, setKind] = useState<ChartKind>('column'),
    [unit, setUnit] = useState('');
  const [category, setCategory] = useState(0),
    [series, setSeries] = useState([1]),
    [startInput, setStartInput] = useState('1'),
    [lengthInput, setLengthInput] = useState('12'),
    [reason, setReason] = useState(''),
    [confirm, setConfirm] = useState(false),
    [inserted, setInserted] = useState('');
  const [after, setAfter] = useState(''),
    [location, setLocation] = useState<{ target: PresentationInsertTarget; id: string }>();
  async function reviewInsertion() {
    if (!after || !access) {
      setLocation(undefined);
      setConfirm(true);
      return;
    }
    await run(t('核对插入页码', 'Checking insertion point'), async () => {
      const target = await access.prepareInsert(),
        page = Number(after);
      if (!Number.isInteger(page) || page < 1 || page > target.slideIds.length)
        throw Error(t('请输入当前稿中存在的页码。', 'Enter a slide number in the current deck.'));
      setLocation({ target, id: target.slideIds[page - 1] });
      setConfirm(true);
    });
  }
  const [turns, setTurns] = useState<EditingTurn[]>([]),
    [followUp, setFollowUp] = useState(''),
    [conversationStamp, setConversationStamp] = useState('');
  const title = customTitle ?? t('业务数据对比', 'Business comparison');
  const language = getUiLanguage();
  const start = Number(startInput),
    length = Number(lengthInput);
  const parsed = useMemo(() => {
    try {
      if (
        !Number.isInteger(start) ||
        start < 1 ||
        start > 1000 ||
        !Number.isInteger(length) ||
        length < 1 ||
        length > 30
      )
        throw new Error(
          t(
            '起始行请输入 1–1,000，纳入行数请输入 1–30。',
            'Enter a starting row from 1 to 1,000 and a row count from 1 to 30.',
          ),
        );
      const original =
        source === 'paste' ? tableFromText(pasted) : tables.find((t) => t.id === source);
      if (!original) throw new Error(t('请重新选择素材表格。', 'Select the source table again.'));
      if (start > original.rows.length)
        throw new Error(
          t(
            `起始数据行超过总行数 ${original.rows.length}。`,
            `The starting row exceeds the ${original.rows.length} data rows available.`,
          ),
        );
      const table: MaterialTable = {
        ...original,
        rows: original.rows.slice(start - 1, start - 1 + length),
      };
      return {
        table,
        originalRows: original.rows.length,
        data: chartData(table, category, series, kind),
        error: '',
      };
    } catch (e) {
      return { error: e instanceof Error ? e.message : t('数据无效。', 'Invalid data.') };
    }
  }, [source, pasted, context, start, length, category, series, kind, language]);
  const fingerprint = JSON.stringify([
    source,
    pasted,
    after,
    parsed.table,
    category,
    series,
    kind,
    title,
    unit,
  ]);
  const dataStamp = JSON.stringify([source, pasted, parsed.table, category, series, modelStamp]);
  const conversationOutdated = Boolean(conversationStamp && conversationStamp !== dataStamp);
  function reset() {
    setConfirm(false);
    setReason('');
  }
  function recommendationMessages() {
    return materialMessages(
      { materials: [], includeImages: false },
      t(
        '你是数据可视化助手。根据提供的数据推荐图表类型与简明标题，不修改数字。只返回 JSON：{"type":"column|bar|line|pie|doughnut","title":"≤64字符","reason":"≤300字符"}。饼图和圆环图只能一个系列、最多8个分类，不能包含负数、空值或全零；趋势适用折线，比较适用柱形或条形。',
        'Recommend a chart type and concise title for the supplied data without changing any numbers. Return only JSON with type (column, bar, line, pie or doughnut), title (up to 64 characters), and reason (up to 300 characters). Pie and doughnut charts require one complete nonnegative series, at most 8 categories and a positive total. Use a line for trends and bars or columns for comparisons.',
      ),
      { title, unit, data: parsed.data },
    );
  }
  function parseRecommendation(raw: string) {
    const result = json(raw);
    fields(result, ['type', 'title', 'reason'], t('图表建议', 'Chart recommendation'));
    chartData(parsed.table!, category, series, result.type as ChartKind);
    return {
      type: result.type as ChartKind,
      title: text(result.title, 64, t('图表标题', 'Chart title')),
      reason: text(result.reason, 300, t('推荐理由', 'Recommendation')),
    };
  }
  async function recommend() {
    if (!parsed.data) return;
    await run(
      t('根据数据推荐图表', 'Recommending a chart'),
      async (signal) => {
        const raw = await request(
          recommendationMessages(),
          signal,
          JSON.stringify({
            type: 'column',
            title: t('业务数据对比 · 演示建议', 'Business comparison · Demo recommendation'),
            reason: t(
              '固定演示：用柱形图比较各分类；真实模型会根据数据推荐。',
              'The demo compares categories with columns. A connected model can recommend a chart for your data.',
            ),
          }),
        );
        if (signal.aborted) return;
        const result = parseRecommendation(raw);
        setKind(result.type);
        setTitle(result.title);
        setReason(result.reason);
        setConfirm(false);
        setConversationStamp(dataStamp);
        setTurns([
          {
            instruction: t('根据当前数据推荐图表。', 'Recommend a chart for the current data.'),
            result: JSON.stringify(result),
          },
        ]);
        setFollowUp('');
      },
      true,
    );
  }
  async function refine() {
    if (!parsed.data || !conversationStamp || conversationOutdated || busy || !followUp.trim())
      return;
    const instruction = followUp.trim();
    await run(
      t('根据对话调整图表建议', 'Refining the chart recommendation'),
      async (signal) => {
        const current = JSON.stringify({
          type: kind,
          title,
          reason: reason || t('当前预览中的图表设置。', 'Chart settings in the current preview.'),
        });
        const messages = buildEditingFollowUp(
          recommendationMessages(),
          turns,
          current,
          instruction,
        );
        const raw = await request(messages, signal, current);
        if (signal.aborted) return;
        const result = parseRecommendation(raw);
        setKind(result.type);
        setTitle(result.title);
        setReason(result.reason);
        setConfirm(false);
        setTurns((previous) =>
          appendEditingTurn(previous, current, instruction, JSON.stringify(result)),
        );
        setFollowUp('');
        notify(
          demo
            ? t(
                '固定演示保留当前图表；连接真实模型后可按对话调整。',
                'The demo keeps the current chart. Connect a model to refine it through conversation.',
              )
            : t(
                '已更新图表预览，数据保持不变。请检查后再插入。',
                'Chart preview updated without changing the data. Review it before insertion.',
              ),
        );
      },
      true,
    );
  }
  async function exportChart(target: 'download' | 'insert') {
    if (!parsed.table || !parsed.data || !title.trim()) return;
    await run(
      target === 'insert'
        ? t('插入并核对图表页', 'Inserting and verifying the chart slide')
        : t('生成可编辑图表文件', 'Building an editable chart'),
      async (signal) => {
        const insertTarget =
          target === 'insert' && access
            ? location?.target || (await access.prepareInsert())
            : undefined;
        if (target === 'insert' && !insertTarget?.size)
          throw new Error(
            t(
              '无法确认当前页面尺寸，请下载图表页后手动导入。',
              'The slide size could not be confirmed. Download the chart slide and import it in PowerPoint.',
            ),
          );
        const chartContext: MaterialContext = {
          materials: [
            {
              id: 'chart-source',
              name: t(
                `${parsed.table!.name} · 数据行 ${start}–${start + parsed.table!.rows.length - 1}`,
                `${parsed.table!.name} · Data rows ${start}–${start + parsed.table!.rows.length - 1}`,
              ),
              kind: 'spreadsheet',
              text: '',
              tables: [parsed.table!],
            },
          ],
          includeImages: false,
        };
        const deck: DeckContent = {
          title,
          slides: [
            {
              title,
              layout: 'chart',
              chart: { type: kind, tableId: parsed.table!.id, category, series, unit },
              notes: t(
                `数据源：${chartContext.materials[0].name}。空值保持为空，不作为零。`,
                `Data source: ${chartContext.materials[0].name}. Blank cells are preserved as missing values.`,
              ),
              sources: ['chart-source'],
            },
          ],
        };
        const base64 = await renderDeck(
          deck,
          chartContext,
          DEFAULT_THEME,
          insertTarget?.size || WIDE,
        );
        if (signal.aborted) return;
        if (target === 'download') {
          downloadPresentation(base64, title);
          notify(
            t(
              '已生成图表 PPT；可在 PowerPoint 中右键图表编辑数据。',
              'Chart slide downloaded. Right-click the chart in PowerPoint to edit its data.',
            ),
          );
        } else if (access && insertTarget) {
          setConfirm(false);
          setInserted(fingerprint);
          if (location) await access.insert(base64, 1, insertTarget, location.id);
          else await access.insert(base64, 1, insertTarget);
          notify(
            location
              ? t(
                  `已在第 ${after} 页后插入并核对图表页。`,
                  `Inserted and verified the chart after slide ${after}.`,
                )
              : t(
                  '已在末尾插入 1 页原生图表，并核对页数。可右键图表编辑数据。',
                  'One editable chart slide was appended and the slide count verified. Right-click the chart to edit its data.',
                ),
          );
        }
      },
    );
  }
  return (
    <div className="ppt-flow">
      <section className="ppt-compose" aria-label={t('数据图表设置', 'Chart settings')}>
        <h2>{t('把数据变成可编辑图表', 'Create an editable chart')}</h2>
        <p className="field-help">
          {t(
            '使用素材中的表格，或粘贴 Excel 区域。图表与内嵌数据一同写入 PPT，可继续编辑。',
            'Select a source table or paste cells from Excel. The slide includes the chart and its editable data.',
          )}
        </p>
        <label>
          {t('数据来源', 'Data source')}
          <select
            value={source}
            disabled={busy}
            onChange={(e) => {
              setSource(e.target.value);
              setStartInput('1');
              setCategory(0);
              setSeries([1]);
              reset();
            }}
          >
            <option value="paste">{t('粘贴 Excel / CSV 数据', 'Paste Excel or CSV data')}</option>
            {tables.map((sourceTable) => (
              <option key={sourceTable.id} value={sourceTable.id}>
                {sourceTable.name} · {sourceTable.rows.length}
                {t(' 行 / ', ' rows / ')}
                {sourceTable.headers.length}
                {t(' 列', ' columns')}
              </option>
            ))}
          </select>
        </label>
        {source === 'paste' && (
          <label>
            {t('含表头的数据', 'Data with headers')}
            <textarea
              rows={7}
              value={pasted}
              maxLength={60000}
              disabled={busy}
              onChange={(e) => {
                setPasted(e.target.value);
                reset();
              }}
              placeholder={t(
                '月份\t华东\t华南\n1月\t128\t96\n2月\t156\t112\n3月\t184\t138',
                'Month\tEast\tSouth\nJanuary\t128\t96\nFebruary\t156\t112\nMarch\t184\t138',
              )}
            />
          </label>
        )}
        <div className="ppt-field-grid">
          <label>
            {t('起始数据行', 'Starting data row')}
            <input
              type="number"
              min={1}
              max={1000}
              value={startInput}
              disabled={busy}
              onChange={(e) => {
                setStartInput(e.target.value);
                reset();
              }}
            />
          </label>
          <label>
            {t('最多纳入行数', 'Maximum rows')}
            <input
              type="number"
              min={1}
              max={30}
              value={lengthInput}
              disabled={busy}
              onChange={(e) => {
                setLengthInput(e.target.value);
                reset();
              }}
            />
          </label>
        </div>
        <p className="field-help">
          {t(
            '数据行从表头之后算起。最多 30 个分类、6 个系列；空白单元格保留为空，百分数字符按比例解析。',
            'Count from the first row below the header. Use up to 30 categories and 6 series. Blanks remain empty. Percentages are parsed as proportions.',
          )}
        </p>
        {(() => {
          let sourceTable: MaterialTable | undefined;
          try {
            sourceTable =
              source === 'paste' ? tableFromText(pasted) : tables.find((v) => v.id === source);
          } catch {
            /* Parse error is displayed below. */
          }
          return (
            sourceTable && (
              <>
                <label>
                  {t('分类列', 'Category column')}
                  <select
                    value={category}
                    disabled={busy}
                    onChange={(e) => {
                      setCategory(Number(e.target.value));
                      setSeries(series.filter((i) => i !== Number(e.target.value)));
                      reset();
                    }}
                  >
                    {sourceTable.headers.map((h, i) => (
                      <option key={i} value={i}>
                        {h || t(`列 ${i + 1}`, `Column ${i + 1}`)}
                      </option>
                    ))}
                  </select>
                </label>
                <fieldset className="chart-series">
                  <legend>{t('数值列', 'Value columns')}</legend>
                  {sourceTable.headers.map(
                    (h, i) =>
                      i !== category && (
                        <label className="checkbox-label" key={i}>
                          <input
                            type="checkbox"
                            disabled={busy}
                            checked={series.includes(i)}
                            onChange={(e) => {
                              setSeries(
                                e.target.checked ? [...series, i] : series.filter((n) => n !== i),
                              );
                              reset();
                            }}
                          />
                          {h || t(`列 ${i + 1}`, `Column ${i + 1}`)}
                        </label>
                      ),
                  )}
                </fieldset>
              </>
            )
          );
        })()}
        <label>
          {t('图表标题', 'Chart title')}
          <input
            value={title}
            maxLength={64}
            disabled={busy}
            onChange={(e) => {
              setTitle(e.target.value);
              reset();
            }}
          />
        </label>
        <div className="ppt-field-grid">
          <label>
            {t('图表类型', 'Chart type')}
            <select
              value={kind}
              disabled={busy}
              onChange={(e) => {
                setKind(e.target.value as ChartKind);
                reset();
              }}
            >
              {Object.entries(CHART_LABELS).map(([id, name]) => (
                <option key={id} value={id}>
                  {name}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('单位', 'Unit')}
            <input
              value={unit}
              maxLength={24}
              disabled={busy}
              placeholder={t('如：万元、%', 'For example: USD, %')}
              onChange={(e) => {
                setUnit(e.target.value);
                reset();
              }}
            />
          </label>
        </div>
        <button
          className="secondary"
          disabled={busy || !parsed.data}
          onClick={() => void recommend()}
        >
          <Sparkles size={16} />
          {t('让模型推荐图表', 'Recommend a chart')}
        </button>
        {reason && <p className="field-help">{reason}</p>}
        {parsed.error && (pasted || source !== 'paste') && (
          <p className="notice warning" role="status">
            {parsed.error}
          </p>
        )}
      </section>
      {parsed.table && parsed.data && (
        <section className="ppt-result" aria-label={t('图表预览', 'Chart preview')}>
          <h2>{title || t('图表预览', 'Chart preview')}</h2>
          <ChartPreview data={parsed.data} kind={kind} unit={unit} />
          <p className="field-help">
            {t('数据行 ', 'Rows ')}
            {start}–{start + parsed.table.rows.length - 1}
            {t(' / 共 ', ' of ')}
            {parsed.originalRows}
            {t(' 行 ·', ' ·')} {parsed.data.series.length}
            {t(
              ' 个系列。图表数据来自此范围，不由模型生成。',
              ' series. Chart values come from the selected rows.',
            )}
          </p>
          <details>
            <summary>{t('核对实际图表数据', 'View chart data')}</summary>
            <div className="table-scroll">
              <table>
                <thead>
                  <tr>
                    <th>{parsed.table.headers[category]}</th>
                    {parsed.data.series.map((s, i) => (
                      <th key={i}>{s.name}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {parsed.data.categories.map((c, i) => (
                    <tr key={i}>
                      <td>{c}</td>
                      {parsed.data.series.map((s, j) => (
                        <td key={j}>{s.values[i] === null ? t('空值', 'Blank') : s.values[i]}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </details>
          {conversationStamp && (
            <EditingConversation
              id="chart-conversation"
              turns={turns}
              value={followUp}
              onChange={setFollowUp}
              onSubmit={() => void refine()}
              onReset={() => {
                setTurns([]);
                setFollowUp('');
              }}
              disabled={busy}
              stale={conversationOutdated}
            />
          )}
          <div className="button-row ppt-export">
            <button
              className="primary"
              disabled={busy || !title.trim()}
              onClick={() => void exportChart('download')}
            >
              <Download size={16} />
              {t('下载图表 PPT', 'Download chart slide')}
            </button>
            {access?.capabilities.insert && (
              <button
                className="secondary"
                disabled={busy || !title.trim() || inserted === fingerprint}
                onClick={() => void reviewInsertion()}
              >
                {inserted === fingerprint
                  ? t('已尝试插入，请检查末尾', 'Insertion attempted. Check the last slide.')
                  : t('插入可编辑图表', 'Insert editable chart')}
              </button>
            )}
          </div>
          {access?.capabilities.insert && (
            <label className="field-label">
              {t('插在第几页之后', 'Insert after slide')}
              <input
                type="number"
                min={1}
                value={after}
                disabled={busy}
                placeholder={t('默认在末尾', 'End of deck by default')}
                onChange={(e) => {
                  setAfter(e.target.value);
                  setConfirm(false);
                  setLocation(undefined);
                }}
              />
            </label>
          )}
          {confirm && (
            <div className="ppt-confirm">
              <p>
                {location
                  ? t(
                      `在第 ${after} 页后新增 1 页可编辑图表。`,
                      `Insert one editable chart after slide ${after}.`,
                    )
                  : t(
                      '在当前演示稿末尾新增 1 页图表。图表可继续编辑，原有页面保留。',
                      'Append one editable chart slide to the end of the current deck.',
                    )}
              </p>
              <div className="button-row">
                <button
                  className="primary"
                  disabled={busy}
                  onClick={() => void exportChart('insert')}
                >
                  {t('确认插入图表页', 'Confirm chart insertion')}
                </button>
                <button className="secondary" disabled={busy} onClick={() => setConfirm(false)}>
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
