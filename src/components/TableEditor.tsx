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
      setError('请先应用或取消当前单元格的修改。');
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
          throw new Error('请输入有效数字，如 1200.5；百分比 12% 输入 0.12。');
        next = Number(value);
      }
      if (type === 'boolean') {
        if (!['true', 'false'].includes(value)) throw new Error('布尔值请输入 true 或 false。');
        next = value === 'true';
      }
      if (type === 'formula') next = { formula: parseFormula(value) };
      if (type === 'text' && value.length > 4000) throw new Error('单元格文本最多 4,000 字符。');
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
        新工作表名称
        <input
          disabled={disabled}
          value={table.name}
          onChange={(e) => onChange({ ...table, name: e.target.value })}
        />
      </label>
      <p className="field-help" id="table-keyboard-help">
        选择单元格后在下方修改。方向键切换单元格，Tab 进入编辑区。
      </p>
      <div className="table-scroll" tabIndex={0} role="region" aria-label="选择要编辑的单元格">
        <table>
          <thead>
            <tr>
              <th scope="col">行</th>
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
                      aria-label={`${columnName(c)}${r + 2} ${cellText(cell) || '空值'}`}
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
                      {cellText(cell) || '空值'}
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
          编辑 {columnName(selected[1])}
          {selected[0] + 2} · {table.headers[selected[1]]}
        </legend>
        <label>
          当前列名称
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
          本列显示格式
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
                ['general', '常规'],
                ['text', '文本'],
                ['integer', '整数'],
                ['number', '两位小数'],
                ['percent', '百分比'],
              ] as const
            ).map(([format, label]) => (
              <option key={format} value={format}>
                {label}
              </option>
            ))}
          </select>
        </label>
        <label>
          单元格类型
          <select
            value={type}
            onChange={(e) => {
              setType(e.target.value as CellType);
              setDirty(true);
            }}
          >
            {[
              ['text', '文本'],
              ['number', '数字'],
              ['formula', '公式'],
              ['boolean', '布尔值'],
              ['empty', '空值'],
            ].map(([kind, label]) => (
              <option value={kind} key={kind}>
                {label}
              </option>
            ))}
          </select>
        </label>
        <label>
          单元格内容
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
            应用单元格修改
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
            取消修改
          </button>
        </div>
        {dirty && (
          <p className="field-help" role="status">
            单元格修改尚未应用，预览与写入仍使用原值。
          </p>
        )}
      </fieldset>
    </div>
  );
}
