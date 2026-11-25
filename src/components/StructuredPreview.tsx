import { t } from '../i18n';
import { cellText, hasExplicitListNumbers } from '../core/structured-output';
import type { ReportContent, TableContent } from '../core/types';

export function ReportPreview({ report }: { report: ReportContent }) {
  return (
    <article className="report-preview">
      <h3>{report.title}</h3>
      {report.sections.map((section, i) => (
        <section key={i}>
          <h4>{section.heading}</h4>
          {section.paragraphs.map((paragraph, p) => (
            <p key={p}>{paragraph}</p>
          ))}
          {hasExplicitListNumbers(section.bullets) ? (
            <div role="list">
              {section.bullets.map((bullet, b) => (
                <p role="listitem" key={b}>
                  {bullet}
                </p>
              ))}
            </div>
          ) : (
            section.bullets.length > 0 && (
              <ul>
                {section.bullets.map((bullet, b) => (
                  <li key={b}>{bullet}</li>
                ))}
              </ul>
            )
          )}
          {section.table && (
            <div
              className="table-scroll"
              tabIndex={0}
              role="region"
              aria-label={t(`${section.heading}表格`, `${section.heading} table`)}
            >
              <table>
                <thead>
                  <tr>
                    {section.table.headers.map((header, c) => (
                      <th scope="col" key={c}>
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {section.table.rows.map((row, r) => (
                    <tr key={r}>
                      {row.map((cell, c) => (
                        <td key={c}>{cell}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </section>
      ))}
    </article>
  );
}

export function TablePreview({ table }: { table: TableContent }) {
  return (
    <div className="table-preview">
      <h3>{table.name}</h3>
      <p className="field-help">
        {t(
          '表头从 A1 开始。公式显示为表达式，写入后由 Excel 计算。',
          'Headers start in A1. Formulas are shown here and calculated in Excel after writing.',
        )}
      </p>
      <div
        className="table-scroll"
        tabIndex={0}
        role="region"
        aria-label={t('数据表预览', 'Data table preview')}
      >
        <table>
          <thead>
            <tr>
              <th scope="col">{t('行', 'Row')}</th>
              {table.headers.map((header, c) => (
                <th scope="col" key={c}>
                  {header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {table.rows.map((row, r) => (
              <tr key={r}>
                <th scope="row">{r + 2}</th>
                {row.map((cell, c) => (
                  <td key={c} className={cell && typeof cell === 'object' ? 'formula-cell' : ''}>
                    {cell === null ? (
                      <span className="muted">{t('空', 'Empty')}</span>
                    ) : (
                      cellText(cell)
                    )}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
