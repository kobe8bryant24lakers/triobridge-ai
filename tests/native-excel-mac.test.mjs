import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';

const script = readFileSync(new URL('../offline/mac/office-mac.js', import.meta.url), 'utf8');
const excel = runInNewContext(script + '\nexcel', { ObjC: { import() {} } });

function fixture({ rows = 1, columns = 1, count = rows * columns } = {}) {
  const state = {
    book: 'synthetic.xlsx',
    fullName: '/qa/synthetic.xlsx',
    sheet: 'QA',
    address: '$B$6',
    value: '',
    formula: '',
    merged: false,
    array: false,
    reads: 0,
    writes: [],
  };
  const range = {
    rows: { length: rows },
    columns: { length: columns },
    cells: { length: count },
    // This is what real Mac Excel returns, unlike its Windows COM counterpart.
    countLarge: () => null,
    value() {
      state.reads++;
      return state.value;
    },
    mergeCells: () => state.merged,
    hasArray: () => state.array,
  };
  Object.defineProperty(range, 'formula', {
    get: () => () => {
      state.reads++;
      return state.formula;
    },
    set: (value) => {
      state.writes.push(value);
      state.formula = value;
    },
  });
  const app = {
    activeWorkbook: { name: () => state.book, fullName: () => state.fullName },
    activeSheet: { name: () => state.sheet },
    activeCell: range,
    selection: range,
    getAddress: () => state.address,
  };
  return {
    state,
    range,
    read: () => excel(app, { op: 'selection' }),
    target: () => excel(app, { op: 'target', kind: 'formula' }),
    write: (guard) => excel(app, { op: 'write', kind: 'formula', guard, value: '=SUM(B2:B4)' }),
  };
}

test('Mac Excel writes the confirmed single cell even when countLarge returns null', () => {
  const f = fixture();
  const target = f.target();
  assert.equal(target.label, 'Excel · QA!$B$6');
  assert.equal(target.text, '');
  f.write(target.guard);
  assert.deepEqual(f.state.writes, ['=SUM(B2:B4)']);
});

test('Mac Excel reads an ordinary rectangular range', () => {
  const f = fixture({ rows: 4, columns: 2 });
  f.state.address = '$A$1:$B$4';
  f.state.value = [
    ['项目', '金额'],
    ['甲', 120],
    ['乙', 80],
    ['丙', 50],
  ];
  f.state.formula = f.state.value;
  assert.equal(f.read().text, '项目\t金额\n甲\t120\n乙\t80\n丙\t50');
});

test('Mac Excel rejects a multiple-cell formula target before writing', () => {
  const f = fixture({ rows: 2 });
  const target = f.target();
  assert.throws(() => f.write(target.guard), /一个公式目标单元格/);
  assert.deepEqual(f.state.writes, []);
});

for (const dimensions of [
  { rows: 2001, columns: 1 },
  { rows: 1048576, columns: 16384, count: 0 },
  { rows: 1048576, columns: 16384, count: 1 },
  { rows: 1, columns: 1, count: 2001 },
]) {
  test(
    'Mac Excel bounds collection reads before fetching values: ' + JSON.stringify(dimensions),
    () => {
      const f = fixture(dimensions);
      assert.throws(() => f.read(), /2,000/);
      assert.equal(f.state.reads, 0);
    },
  );
}

for (const dimensions of [
  { rows: 0 },
  { columns: NaN },
  { count: null },
  { count: 0.5 },
  { rows: 2, columns: 2, count: 3 },
]) {
  test(
    'Mac Excel fails closed for unverified selection dimensions: ' + JSON.stringify(dimensions),
    () => {
      const f = fixture(dimensions);
      assert.throws(() => f.read(), /无法确认选区大小/);
      assert.equal(f.state.reads, 0);
    },
  );
}

for (const [field, changed] of [
  ['book', 'other.xlsx'],
  ['fullName', '/other/synthetic.xlsx'],
  ['sheet', 'Other'],
  ['address', '$C$6'],
  ['value', 'changed'],
  ['formula', '=1'],
]) {
  test('Mac Excel retains the confirmation guard for ' + field, () => {
    const f = fixture();
    const target = f.target();
    f.state[field] = changed;
    assert.throws(() => f.write(target.guard), /已发生变化/);
    assert.deepEqual(f.state.writes, []);
  });
}

for (const protectedKind of ['merged', 'array']) {
  test('Mac Excel preserves ' + protectedKind + ' cell protection', () => {
    const f = fixture();
    const target = f.target();
    f.state[protectedKind] = true;
    assert.throws(() => f.write(target.guard), /合并区域或数组公式/);
    assert.deepEqual(f.state.writes, []);
  });
}
