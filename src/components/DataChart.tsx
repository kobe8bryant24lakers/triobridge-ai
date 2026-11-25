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

export function DataChart({ context, access, busy, run, request, notify }: PresentationRuntime) {
  const tables = allTables(context),
    [source, setSource] = useState('paste'),
    [pasted, setPasted] = useState(''),
    [title, setTitle] = useState('业务数据对比'),
    [kind, setKind] = useState<ChartKind>('column'),
    [unit, setUnit] = useState('');
  const [category, setCategory] = useState(0),
    [series, setSeries] = useState([1]),
    [startInput, setStartInput] = useState('1'),
    [lengthInput, setLengthInput] = useState('12'),
    [reason, setReason] = useState(''),
    [confirm, setConfirm] = useState(false),
    [inserted, setInserted] = useState('');
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
        throw new Error('起始行请输入 1–1,000，纳入行数请输入 1–30。');
      const original =
        source === 'paste' ? tableFromText(pasted) : tables.find((t) => t.id === source);
      if (!original) throw new Error('请重新选择素材表格。');
      if (start > original.rows.length)
        throw new Error(`起始数据行超过总行数 ${original.rows.length}。`);
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
      return { error: e instanceof Error ? e.message : '数据无效。' };
    }
  }, [source, pasted, context, start, length, category, series, kind]);
  const fingerprint = JSON.stringify([
    source,
    pasted,
    parsed.table,
    category,
    series,
    kind,
    title,
    unit,
  ]);
  function reset() {
    setConfirm(false);
    setReason('');
  }
  async function recommend() {
    if (!parsed.data) return;
    await run('根据数据推荐图表', async (signal) => {
      const raw = await request(
        materialMessages(
          { materials: [], includeImages: false },
          '你是数据可视化助手。根据提供的数据推荐图表类型与简明标题，不修改数字。只返回 JSON：{"type":"column|bar|line|pie|doughnut","title":"≤64字符","reason":"≤300字符"}。饼图和圆环图只能一个系列、最多8个分类，不能包含负数、空值或全零；趋势适用折线，比较适用柱形或条形。',
          { title, unit, data: parsed.data },
        ),
        signal,
        JSON.stringify({
          type: 'column',
          title: '业务数据对比 · 演示建议',
          reason: '固定演示：用柱形图比较各分类；真实模型会根据数据推荐。',
        }),
      );
      if (signal.aborted) return;
      const result = json(raw);
      fields(result, ['type', 'title', 'reason'], '图表建议');
      chartData(parsed.table!, category, series, result.type as ChartKind);
      setKind(result.type as ChartKind);
      setTitle(text(result.title, 64, '图表标题'));
      setReason(text(result.reason, 300, '推荐理由'));
      setConfirm(false);
    });
  }
  async function exportChart(target: 'download' | 'insert') {
    if (!parsed.table || !parsed.data || !title.trim()) return;
    await run(target === 'insert' ? '插入并核对图表页' : '生成可编辑图表文件', async (signal) => {
      const insertTarget = target === 'insert' && access ? await access.prepareInsert() : undefined;
      if (target === 'insert' && !insertTarget?.size)
        throw new Error('无法确认当前页面尺寸，请下载图表页后手动导入。');
      const chartContext: MaterialContext = {
        materials: [
          {
            id: 'chart-source',
            name: `${parsed.table!.name} · 数据行 ${start}–${start + parsed.table!.rows.length - 1}`,
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
            notes: `数据源：${chartContext.materials[0].name}。空值保持为空，不作为零。`,
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
        notify('已生成图表 PPT；可在 PowerPoint 中右键图表编辑数据。');
      } else if (access && insertTarget) {
        setConfirm(false);
        setInserted(fingerprint);
        await access.insert(base64, 1, insertTarget);
        notify('已在末尾插入 1 页原生图表，并核对页数。可右键图表编辑数据。');
      }
    });
  }
  return (
    <div className="ppt-flow">
      <section className="ppt-compose" aria-label="数据图表设置">
        <h2>把数据变成可编辑图表</h2>
        <p className="field-help">
          使用素材中的表格，或粘贴 Excel 区域。图表与内嵌数据一同写入 PPT，可继续编辑。
        </p>
        <label>
          数据来源
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
            <option value="paste">粘贴 Excel / CSV 数据</option>
            {tables.map((t) => (
              <option key={t.id} value={t.id}>
                {t.name} · {t.rows.length} 行 / {t.headers.length} 列
              </option>
            ))}
          </select>
        </label>
        {source === 'paste' && (
          <label>
            含表头的数据
            <textarea
              rows={7}
              value={pasted}
              maxLength={60000}
              disabled={busy}
              onChange={(e) => {
                setPasted(e.target.value);
                reset();
              }}
              placeholder={'月份\t华东\t华南\n1月\t128\t96\n2月\t156\t112\n3月\t184\t138'}
            />
          </label>
        )}
        <div className="ppt-field-grid">
          <label>
            起始数据行
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
            最多纳入行数
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
          数据行从表头之后算起。最多 30 个分类、6 个系列；空白单元格保留为空，百分数字符按比例解析。
        </p>
        {(() => {
          let t: MaterialTable | undefined;
          try {
            t = source === 'paste' ? tableFromText(pasted) : tables.find((v) => v.id === source);
          } catch {
            /* Parse error is displayed below. */
          }
          return (
            t && (
              <>
                <label>
                  分类列
                  <select
                    value={category}
                    disabled={busy}
                    onChange={(e) => {
                      setCategory(Number(e.target.value));
                      setSeries(series.filter((i) => i !== Number(e.target.value)));
                      reset();
                    }}
                  >
                    {t.headers.map((h, i) => (
                      <option key={i} value={i}>
                        {h || `列 ${i + 1}`}
                      </option>
                    ))}
                  </select>
                </label>
                <fieldset className="chart-series">
                  <legend>数值列</legend>
                  {t.headers.map(
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
                          {h || `列 ${i + 1}`}
                        </label>
                      ),
                  )}
                </fieldset>
              </>
            )
          );
        })()}
        <label>
          图表标题
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
            图表类型
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
            单位
            <input
              value={unit}
              maxLength={24}
              disabled={busy}
              placeholder="如：万元、%"
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
          让模型推荐图表
        </button>
        {reason && <p className="field-help">{reason}</p>}
        {parsed.error && (pasted || source !== 'paste') && (
          <p className="notice warning" role="status">
            {parsed.error}
          </p>
        )}
      </section>
      {parsed.table && parsed.data && (
        <section className="ppt-result" aria-label="图表预览">
          <h2>{title || '图表预览'}</h2>
          <ChartPreview data={parsed.data} kind={kind} unit={unit} />
          <p className="field-help">
            数据行 {start}–{start + parsed.table.rows.length - 1} / 共 {parsed.originalRows} 行 ·{' '}
            {parsed.data.series.length} 个系列。图表数据来自此范围，不由模型生成。
          </p>
          <details>
            <summary>核对实际图表数据</summary>
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
                        <td key={j}>{s.values[i] === null ? '空值' : s.values[i]}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </details>
          <div className="button-row ppt-export">
            <button
              className="primary"
              disabled={busy || !title.trim()}
              onClick={() => void exportChart('download')}
            >
              <Download size={16} />
              下载图表 PPT
            </button>
            {access?.capabilities.insert && (
              <button
                className="secondary"
                disabled={busy || !title.trim() || inserted === fingerprint}
                onClick={() => setConfirm(true)}
              >
                {inserted === fingerprint ? '已尝试插入，请检查末尾' : '插入可编辑图表'}
              </button>
            )}
          </div>
          {confirm && (
            <div className="ppt-confirm">
              <p>在当前演示稿末尾新增 1 页图表。图表可继续编辑，原有页面保留。</p>
              <div className="button-row">
                <button
                  className="primary"
                  disabled={busy}
                  onClick={() => void exportChart('insert')}
                >
                  确认插入图表页
                </button>
                <button className="secondary" disabled={busy} onClick={() => setConfirm(false)}>
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
