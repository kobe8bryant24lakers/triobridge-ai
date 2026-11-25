import { t } from '../i18n';
import { useEffect, useRef, useState } from 'react';
import { cellText, columnName } from '../core/structured-output';
import { parseFormula } from '../core/output';
import type { ColumnFormat, TableCell, TableContent } from '../core/types';

type CellType = 'text' | 'number' | 'boolean' | 'empty' | 'formula';
function cellType(cell: TableCell): CellType {
  return cell === null
    ? 'empty'
    : typeof cell === 'object'
      ? 'formula'
      : typeof cell === 'number'
        ? 'number'
        : typeof cell === 'boolean'
          ? 'boolean'
          : 'text';
}

export function TableEditor({
  table,
  disabled,
  onChange,
  onPendingChange,
}: {
  table: TableContent;
  disabled: boolean;
  onChange: (value: TableContent) => void;
  onPendingChange: (pending: boolean) => void;
}) {
  const [selected, setSelected] = useState<[number, number]>([0, 0]);
  const [value, setValue] = useState(cellText(table.rows[0][0]));
  const [type, setType] = useState<CellType>(cellType(table.rows[0][0]));
  const [error, setError] = useState('');
  const [dirty, setDirty] = useState(false);
  const cells = useRef(new Map<string, HTMLButtonElement>());
  useEffect(() => {
    onPendingChange(dirty);
    return () => onPendingChange(false);
  }, [dirty, onPendingChange]);
  function select(r: number, c: number) {
    if (dirty) {
      setError(
        t('请先应用或取消当前单元格的修改。', 'Apply or cancel your changes to this cell first.'),
      );
      return;
    }
    setSelected([r, c]);
    setValue(cellText(table.rows[r][c]));
    setType(cellType(table.rows[r][c]));
    setError('');
  }
  function apply() {
    try {
      let next: TableCell = value;
      if (type === 'empty') next = null;
      if (type === 'number') {
        if (!value.trim() || !Number.isFinite(Number(value)))
          throw new Error(
            t(
              '请输入有效数字，如 1200.5；百分比 12% 输入 0.12。',
              'Enter a valid number, such as 1200.5. For 12%, enter 0.12.',
            ),
          );
        next = Number(value);
      }
      if (type === 'boolean') {
        if (!['true', 'false'].includes(value))
          throw new Error(
            t('布尔值请输入 true 或 false。', 'Enter true or false for a Boolean value.'),
          );
        next = value === 'true';
      }
      if (type === 'formula') next = { formula: parseFormula(value) };
      if (type === 'text' && value.length > 4000)
        throw new Error(
          t('单元格文本最多 4,000 字符。', 'Cell text can contain up to 4,000 characters.'),
        );
      onChange({
        ...table,
        rows: table.rows.map((row, r) =>
          r === selected[0] ? row.map((cell, c) => (c === selected[1] ? next : cell)) : row,
        ),
      });
      setError('');
      setDirty(false);
    } catch (reason) {
      setError((reason as Error).message);
    }
  }
  return (
    <div className="table-editor">
      <label>
        {t('新工作表名称', 'New worksheet name')}
        <input
          disabled={disabled}
          value={table.name}
          onChange={(e) => onChange({ ...table, name: e.target.value })}
        />
      </label>
      <p className="field-help" id="table-keyboard-help">
        {t(
          '选择单元格后在下方修改。方向键切换单元格，Tab 进入编辑区。',
          'Select a cell to edit it below. Use arrow keys to change cells and Tab to enter the editor.',
        )}
      </p>
      <div
        className="table-scroll"
        tabIndex={0}
        role="region"
        aria-label={t('选择要编辑的单元格', 'Choose a cell to edit')}
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
                  <td key={c}>
                    <button
                      className="cell-button"
                      ref={(element) => {
                        if (element) cells.current.set(`${r}:${c}`, element);
                        else cells.current.delete(`${r}:${c}`);
                      }}
                      disabled={disabled}
                      tabIndex={selected[0] === r && selected[1] === c ? 0 : -1}
                      aria-describedby="table-keyboard-help"
                      aria-pressed={selected[0] === r && selected[1] === c}
                      aria-label={`${columnName(c)}${r + 2} ${cellText(cell) || t('空值', 'Empty cell')}`}
                      onClick={() => select(r, c)}
                      onKeyDown={(event) => {
                        const offsets: Record<string, [number, number]> = {
                          ArrowUp: [-1, 0],
                          ArrowDown: [1, 0],
                          ArrowLeft: [0, -1],
                          ArrowRight: [0, 1],
                        };
                        const offset = offsets[event.key];
                        if (!offset) return;
                        event.preventDefault();
                        const nextRow = Math.max(0, Math.min(table.rows.length - 1, r + offset[0]));
                        const nextColumn = Math.max(
                          0,
                          Math.min(table.headers.length - 1, c + offset[1]),
                        );
                        select(nextRow, nextColumn);
                        if (!dirty) cells.current.get(`${nextRow}:${nextColumn}`)?.focus();
                      }}
                    >
                      {cellText(cell) || t('空值', 'Empty cell')}
                    </button>
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <fieldset className="cell-inspector" disabled={disabled}>
        <legend>
          {t('编辑 ', 'Edit ')}
          {columnName(selected[1])}
          {selected[0] + 2} · {table.headers[selected[1]]}
        </legend>
        <label>
          {t('当前列名称', 'Column name')}
          <input
            value={table.headers[selected[1]]}
            onChange={(e) =>
              onChange({
                ...table,
                headers: table.headers.map((header, c) =>
                  c === selected[1] ? e.target.value : header,
                ),
              })
            }
          />
        </label>
        <label>
          {t('本列显示格式', 'Column display format')}
          <select
            value={table.formats[selected[1]]}
            onChange={(e) =>
              onChange({
                ...table,
                formats: table.formats.map((format, c) =>
                  c === selected[1] ? (e.target.value as ColumnFormat) : format,
                ),
              })
            }
          >
            {(
              [
                ['general', t('常规', 'General')],
                ['text', t('文本', 'Text')],
                ['integer', t('整数', 'Integer')],
                ['number', t('两位小数', 'Two decimal places')],
                ['percent', t('百分比', 'Percentage')],
              ] as const
            ).map(([format, label]) => (
              <option key={format} value={format}>
                {label}
              </option>
            ))}
          </select>
        </label>
        <label>
          {t('单元格类型', 'Cell type')}
          <select
            value={type}
            onChange={(e) => {
              setType(e.target.value as CellType);
              setDirty(true);
            }}
          >
            {[
              ['text', t('文本', 'Text')],
              ['number', t('数字', 'Number')],
              ['formula', t('公式', 'Formula')],
              ['boolean', t('布尔值', 'Boolean')],
              ['empty', t('空值', 'Empty cell')],
            ].map(([kind, label]) => (
              <option value={kind} key={kind}>
                {label}
              </option>
            ))}
          </select>
        </label>
        <label>
          {t('单元格内容', 'Cell content')}
          <textarea
            rows={3}
            value={value}
            disabled={disabled || type === 'empty'}
            onChange={(e) => {
              setValue(e.target.value);
              setDirty(true);
            }}
          />
        </label>
        {error && (
          <p className="field-error" role="alert">
            {error}
          </p>
        )}
        <div className="button-row">
          <button className="secondary" onClick={apply} disabled={disabled || !dirty}>
            {t('应用单元格修改', 'Apply Cell Changes')}
          </button>
          <button
            className="text-button"
            disabled={disabled || !dirty}
            onClick={() => {
              setDirty(false);
              setError('');
              setValue(cellText(table.rows[selected[0]][selected[1]]));
              setType(cellType(table.rows[selected[0]][selected[1]]));
            }}
          >
            {t('取消修改', 'Cancel Changes')}
          </button>
        </div>
        {dirty && (
          <p className="field-help" role="status">
            {t(
              '单元格修改尚未应用，预览与写入仍使用原值。',
              'This cell has unapplied changes. The preview and written output still use its previous value.',
            )}
          </p>
        )}
      </fieldset>
    </div>
  );
}
