import type { ChartData, ChartKind, DeckSlide, MaterialContext } from '../presentation/types';
import { allTables, LAYOUT_LABELS } from '../presentation/content';
import { chartData } from '../presentation/charts';

const palette = ['#28604A', '#6F94A8', '#C59B53', '#937A9B', '#9AAB75', '#C88176'];
export function ChartPreview({
  data,
  kind,
  unit = '',
}: {
  data: ChartData;
  kind: ChartKind;
  unit?: string;
}) {
  const values = data.series.flatMap((s) => s.values).filter((v): v is number => v !== null),
    low = Math.min(0, ...values),
    high = Math.max(0, ...values),
    span = high - low || 1;
  const y = (v: number) => 220 - ((v - low) / span) * 190,
    n = data.categories.length,
    x = (i: number) => 55 + ((i + 0.5) * 450) / n;
  const horizontalX = (v: number) => 105 + ((v - low) / span) * 400;
  const label = (v: number) =>
    unit === '%'
      ? `${(v * 100).toFixed(1)}%`
      : new Intl.NumberFormat('zh-CN', { maximumFractionDigits: 2, notation: 'compact' }).format(v);
  let angle = -Math.PI / 2;
  const pie = kind === 'pie' || kind === 'doughnut';
  return (
    <figure className="chart-preview">
      <svg
        viewBox="0 0 560 290"
        role="img"
        aria-label={`${data.categories.length} 个分类、${data.series.length} 个系列的${kind === 'line' ? '折线' : pie ? '占比' : '条形'}图预览`}
      >
        {pie ? (
          data.series[0].values.map((value, i) => {
            const share = (value || 0) / values.reduce((a, b) => a + b, 0);
            const start = angle;
            angle += share * Math.PI * 2;
            const end = angle,
              large = share > 0.5 ? 1 : 0;
            if (share === 1)
              return (
                <circle key={i} cx="280" cy="140" r="104" fill={palette[i % palette.length]} />
              );
            return (
              <path
                key={i}
                d={`M280,140 L${280 + Math.cos(start) * 104},${140 + Math.sin(start) * 104} A104,104 0 ${large} 1 ${280 + Math.cos(end) * 104},${140 + Math.sin(end) * 104} Z`}
                fill={palette[i % palette.length]}
              >
                <title>
                  {data.categories[i]}：{label(value || 0)}
                </title>
              </path>
            );
          })
        ) : kind === 'bar' ? (
          <>
            {[0, 0.5, 1].map((t) => (
              <g key={t}>
                <line
                  x1={horizontalX(low + t * span)}
                  x2={horizontalX(low + t * span)}
                  y1="25"
                  y2="238"
                  stroke="#DDE5DE"
                />
                <text
                  x={horizontalX(low + t * span)}
                  y="258"
                  textAnchor="middle"
                  fontSize="11"
                  fill="#526359"
                >
                  {label(low + t * span)}
                </text>
              </g>
            ))}
            <line x1={horizontalX(0)} x2={horizontalX(0)} y1="25" y2="238" stroke="#738579" />
            {data.categories.map((category, i) => (
              <g key={i}>
                {(n <= 12 || i % Math.ceil(n / 12) === 0) && (
                  <text
                    x="98"
                    y={25 + ((i + 0.5) * 213) / n + 3}
                    textAnchor="end"
                    fontSize="10"
                    fill="#526359"
                  >
                    {category.length > 8 ? `${category.slice(0, 7)}…` : category}
                  </text>
                )}
                {data.series.map((series, j) => {
                  const v = series.values[i];
                  return v === null ? null : (
                    <rect
                      key={j}
                      x={Math.min(horizontalX(0), horizontalX(v))}
                      y={25 + (i * 213 + 22 + (j * 169) / data.series.length) / n}
                      width={Math.abs(horizontalX(v) - horizontalX(0))}
                      height={Math.max(0.5, 155 / n / data.series.length)}
                      fill={palette[j]}
                    >
                      <title>
                        {category} · {series.name}：{label(v)}
                      </title>
                    </rect>
                  );
                })}
              </g>
            ))}
          </>
        ) : (
          <>
            {[0, 0.5, 1].map((t) => (
              <g key={t}>
                <line
                  x1="55"
                  x2="505"
                  y1={y(low + t * span)}
                  y2={y(low + t * span)}
                  stroke="#DDE5DE"
                />
                <text
                  x="49"
                  y={y(low + t * span) + 4}
                  textAnchor="end"
                  fontSize="11"
                  fill="#526359"
                >
                  {label(low + t * span)}
                </text>
              </g>
            ))}
            <line x1="55" x2="505" y1={y(0)} y2={y(0)} stroke="#738579" />
            {data.series.map((series, s) => (
              <g key={s}>
                {kind === 'line' ? (
                  <>
                    {series.values.map((v, i) =>
                      v !== null && i > 0 && series.values[i - 1] !== null ? (
                        <line
                          key={i}
                          x1={x(i - 1)}
                          y1={y(series.values[i - 1]!)}
                          x2={x(i)}
                          y2={y(v)}
                          stroke={palette[s]}
                          strokeWidth="2.5"
                        />
                      ) : null,
                    )}
                    {series.values.map((v, i) =>
                      v !== null ? (
                        <circle key={i} cx={x(i)} cy={y(v)} r="3.5" fill={palette[s]}>
                          <title>
                            {data.categories[i]} · {series.name}：{label(v)}
                          </title>
                        </circle>
                      ) : null,
                    )}
                  </>
                ) : (
                  series.values.map((v, i) =>
                    v === null ? null : (
                      <rect
                        key={i}
                        x={x(i) - 170 / n + (s * 340) / n / data.series.length}
                        y={Math.min(y(v), y(0))}
                        width={Math.max(1, 320 / n / data.series.length)}
                        height={Math.max(0.6, Math.abs(y(v) - y(0)))}
                        fill={palette[s]}
                      >
                        <title>
                          {data.categories[i]} · {series.name}：{label(v)}
                        </title>
                      </rect>
                    ),
                  )
                )}
              </g>
            ))}
            {data.categories.map(
              (c, i) =>
                (n <= 12 || i % Math.ceil(n / 12) === 0) && (
                  <text key={i} x={x(i)} y="243" fontSize="10" textAnchor="middle" fill="#526359">
                    {c.length > 7 ? `${c.slice(0, 6)}…` : c}
                  </text>
                ),
            )}
          </>
        )}
        {kind === 'doughnut' && <circle cx="280" cy="140" r="58" fill="#FAFCF9" />}
      </svg>
      <figcaption>
        {pie
          ? data.categories.map((c, i) => (
              <span key={i}>
                <i style={{ backgroundColor: palette[i % palette.length] }} />
                {c} ·{' '}
                {(
                  ((data.series[0].values[i] || 0) / values.reduce((a, b) => a + b, 0)) *
                  100
                ).toFixed(1)}
                %
              </span>
            ))
          : data.series.map((s, i) => (
              <span key={i}>
                <i style={{ backgroundColor: palette[i] }} />
                {s.name}
              </span>
            ))}
      </figcaption>
    </figure>
  );
}
export function SlidePreview({ page, context }: { page: DeckSlide; context: MaterialContext }) {
  const table = allTables(context).find((t) => t.id === (page.chart?.tableId || page.tableId));
  return (
    <div className={`slide-preview layout-${page.layout}`}>
      <h3>{page.title}</h3>
      <span className="slide-layout-name">{LAYOUT_LABELS[page.layout]}</span>
      {page.subtitle && <p>{page.subtitle}</p>}
      {page.layout === 'image' && (
        <img
          src={context.materials.find((m) => m.id === page.imageId)?.image?.data}
          alt={context.materials.find((m) => m.id === page.imageId)?.name || ''}
        />
      )}
      {page.bullets?.length ? (
        <ul>
          {page.bullets.map((b, i) => (
            <li key={i}>{b}</li>
          ))}
        </ul>
      ) : null}
      {page.columns && (
        <div className="slide-columns">
          {page.columns.map((c, i) => (
            <div key={i}>
              <h4>{c.heading}</h4>
              <ul>
                {c.bullets.map((b, n) => (
                  <li key={n}>{b}</li>
                ))}
              </ul>
            </div>
          ))}
        </div>
      )}
      {page.metrics && (
        <dl className="slide-metrics">
          {page.metrics.map((m, i) => (
            <div key={i}>
              <dt>{m.label}</dt>
              <dd>{m.value}</dd>
              <p>{m.detail}</p>
            </div>
          ))}
        </dl>
      )}
      {page.steps && (
        <ol className="slide-steps">
          {page.steps.map((s, i) => (
            <li key={i}>
              <strong>{s.title}</strong>
              <p>{s.detail}</p>
            </li>
          ))}
        </ol>
      )}
      {page.chart && table && (
        <ChartPreview
          kind={page.chart.type}
          unit={page.chart.unit}
          data={chartData(table, page.chart.category, page.chart.series, page.chart.type)}
        />
      )}
      {page.tableId && table && (
        <div className="table-scroll">
          <table>
            <thead>
              <tr>
                {table.headers.map((h, i) => (
                  <th key={i}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {table.rows.map((row, i) => (
                <tr key={i}>
                  {row.map((c, j) => (
                    <td key={j}>{String(c ?? '')}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
