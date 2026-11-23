import type { ReportContent, ReportSection } from '../core/types';

export function ReportEditor({
  report,
  disabled,
  onChange,
}: {
  report: ReportContent;
  disabled: boolean;
  onChange: (value: ReportContent) => void;
}) {
  function update(index: number, change: Partial<ReportSection>) {
    onChange({
      ...report,
      sections: report.sections.map((section, i) =>
        i === index ? { ...section, ...change } : section,
      ),
    });
  }
  return (
    <div className="report-editor">
      <label>
        文档标题
        <input
          disabled={disabled}
          value={report.title}
          onChange={(e) => onChange({ ...report, title: e.target.value })}
        />
      </label>
      {report.sections.map((section, i) => (
        <fieldset className="slide-editor" key={i} disabled={disabled}>
          <legend>第 {i + 1} 节</legend>
          <label>
            第 {i + 1} 节标题
            <input
              value={section.heading}
              onChange={(e) => update(i, { heading: e.target.value })}
            />
          </label>
          <label>
            第 {i + 1} 节正文
            <textarea
              rows={5}
              value={section.paragraphs.join('\n\n')}
              onChange={(e) =>
                update(i, { paragraphs: e.target.value ? e.target.value.split(/\n\n/) : [] })
              }
            />
          </label>
          <p className="field-help">用空行分隔段落。</p>
          <label>
            第 {i + 1} 节要点
            <textarea
              rows={3}
              value={section.bullets.join('\n')}
              onChange={(e) =>
                update(i, { bullets: e.target.value ? e.target.value.split('\n') : [] })
              }
            />
          </label>
          {section.table && (
            <div
              className="table-scroll"
              tabIndex={0}
              role="region"
              aria-label={`编辑第 ${i + 1} 节表格`}
            >
              <table>
                <thead>
                  <tr>
                    {section.table.headers.map((header, c) => (
                      <th scope="col" key={c}>
                        <input
                          aria-label={`第 ${i + 1} 节第 ${c + 1} 列名称`}
                          value={header}
                          onChange={(e) =>
                            update(i, {
                              table: {
                                ...section.table!,
                                headers: section.table!.headers.map((value, col) =>
                                  col === c ? e.target.value : value,
                                ),
                              },
                            })
                          }
                        />
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {section.table.rows.map((row, r) => (
                    <tr key={r}>
                      {row.map((cell, c) => (
                        <td key={c}>
                          <textarea
                            rows={2}
                            aria-label={`第 ${i + 1} 节表格第 ${r + 1} 行第 ${c + 1} 列`}
                            value={cell}
                            onChange={(e) =>
                              update(i, {
                                table: {
                                  ...section.table!,
                                  rows: section.table!.rows.map((values, rowIndex) =>
                                    rowIndex === r
                                      ? values.map((value, col) =>
                                          col === c ? e.target.value : value,
                                        )
                                      : values,
                                  ),
                                },
                              })
                            }
                          />
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          <button
            className="text-button"
            disabled={disabled || report.sections.length === 1}
            onClick={() =>
              onChange({ ...report, sections: report.sections.filter((_, index) => index !== i) })
            }
          >
            删除第 {i + 1} 节
          </button>
        </fieldset>
      ))}
      <button
        className="secondary"
        disabled={disabled || report.sections.length >= 12}
        onClick={() =>
          onChange({
            ...report,
            sections: [
              ...report.sections,
              { heading: '新章节', paragraphs: ['待补充内容'], bullets: [] },
            ],
          })
        }
      >
        添加章节（{report.sections.length} / 12）
      </button>
    </div>
  );
}
