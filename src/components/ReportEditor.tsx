import { t } from '../i18n';
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
        {t('文档标题', 'Document title')}
        <input
          disabled={disabled}
          value={report.title}
          onChange={(e) => onChange({ ...report, title: e.target.value })}
        />
      </label>
      {report.sections.map((section, i) => (
        <fieldset className="slide-editor" key={i} disabled={disabled}>
          <legend>{t(`第 ${i + 1} 节`, `Section ${i + 1}`)}</legend>
          <label>
            {t(`第 ${i + 1} 节标题`, `Section ${i + 1} heading`)}
            <input
              value={section.heading}
              onChange={(e) => update(i, { heading: e.target.value })}
            />
          </label>
          <label>
            {t(`第 ${i + 1} 节正文`, `Section ${i + 1} text`)}
            <textarea
              rows={5}
              value={section.paragraphs.join('\n\n')}
              onChange={(e) =>
                update(i, { paragraphs: e.target.value ? e.target.value.split(/\n\n/) : [] })
              }
            />
          </label>
          <p className="field-help">
            {t('用空行分隔段落。', 'Separate paragraphs with a blank line.')}
          </p>
          <label>
            {t(`第 ${i + 1} 节要点`, `Section ${i + 1} bullets`)}
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
              aria-label={t(`编辑第 ${i + 1} 节表格`, `Edit table in section ${i + 1}`)}
            >
              <table>
                <thead>
                  <tr>
                    {section.table.headers.map((header, c) => (
                      <th scope="col" key={c}>
                        <input
                          aria-label={t(
                            `第 ${i + 1} 节第 ${c + 1} 列名称`,
                            `Section ${i + 1}, column ${c + 1} name`,
                          )}
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
                            aria-label={t(
                              `第 ${i + 1} 节表格第 ${r + 1} 行第 ${c + 1} 列`,
                              `Section ${i + 1} table, row ${r + 1}, column ${c + 1}`,
                            )}
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
            {t(`删除第 ${i + 1} 节`, `Delete section ${i + 1}`)}
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
              {
                heading: t('新章节', 'New section'),
                paragraphs: [t('待补充内容', 'Content to add')],
                bullets: [],
              },
            ],
          })
        }
      >
        {t(
          `添加章节（${report.sections.length} / 12）`,
          `Add section (${report.sections.length} / 12)`,
        )}
      </button>
    </div>
  );
}
