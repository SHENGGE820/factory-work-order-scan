// Checks the table (工廠管理) side of apps-script/Code.gs: options / board / save / remove, against a tiny in-memory spreadsheet.
import fs from 'fs';
import assert from 'assert';
const code = fs.readFileSync(new URL('../apps-script/Code.gs', import.meta.url), 'utf8');

class Sheet {
  constructor(name, rows = []) { this.name = name; this.rows = rows.map(r => [...r]); this.max = Math.max(rows.length, 20); this.cols = 26; } // a new Google sheet has 26 columns
  getMaxRows() { return this.max; }
  getMaxColumns() { return this.cols; }
  insertColumnsAfter(_, n) { this.cols += n; }
  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getDataRange() { return { getDisplayValues: () => this.rows.map(r => r.map(v => String(v ?? ''))) }; }
  setName(n) { this.name = n; }
  setFrozenRows() {}
  insertRowsAfter(_, n) { this.max += n; }
  deleteRow(n) { this.rows.splice(n - 1, 1); }
  getParent() { return { getUrl: () => 'https://docs.google.com/spreadsheets/d/db' }; }
  getRange(r, c, nr = 1, nc = 1) {
    const sh = this;
    if (r + nr - 1 > this.max || c + nc - 1 > this.cols) throw new Error('outside the sheet');
    return {
      getValues() { return Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => sh.rows[r - 1 + i]?.[c - 1 + j] ?? '')); },
      setNumberFormat() { return this; },
      setFontWeight() { return this; },
      setValues(vals) {
        vals.forEach((row, i) => { while (sh.rows.length < r - 1 + i + 1) sh.rows.push([]); row.forEach((v, j) => { sh.rows[r - 1 + i][c - 1 + j] = v; }); });
        return this;
      },
    };
  }
}
const planRows = [
  ['', '', '', '', '', '', ''],
  ['產品編號', '產品名稱', '數量', '單位', '工令單號', '品項', '預計\n料齊到\n倉庫日'],
  ['4104002200098', '2M_西洋梨型 裸包 (每包3公克)', '6,000', '包', 'V1150831005', '4粉包', '9/11'],
  ['7101002200144', '2M_西洋梨 混料(每公斤)', '18', '公斤', 'V1150831006', '7混料', '00'],
  ['重工-\n4104000000001', '重工品', '5', '包', '重工換包裝\nX1150901001', '5包裝', ''],
];
const optRows = [
  ['0', '人數', '生產機台', '預計完工日', '天 / 時', '項次拋光', '回數QC', '進度說明', '', '包裝組操作人員', '人數'],
  ['小明', '1', '240kg混料', '1', '天', '項次：', '回數：', '完工', '', '阿甲', '9'],
  ['小華', '3', '100kg混料', '3', '時', '已拋光', '已QC', '作業中', '', '', ''],
  ['小明', '4', '小粉包機-1', '', '', '', '', '', '', '', ''],
];
const sheets = { plan: new Sheet('訂單計畫單(NEW)', planRows), opts: new Sheet('選單', optRows) };
const books = {
  plan: { getSheetByName: n => n === '訂單計畫單(NEW)' ? sheets.plan : null, getSheets: () => [sheets.plan] },
  bke: { getSheetByName: n => n === '選單' ? sheets.opts : null },
};
let dbBook = null, creates = 0;
const props = { PIN: '1234', SHEET_ID: 'plan', BKE_SHEET_ID: 'bke' };
const cache = new Map();
let uuid = 0;
const g = {
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
  CacheService: { getScriptCache: () => ({ get: k => cache.get(k) ?? null, put: (k, v) => cache.set(k, v), putAll: o => Object.entries(o).forEach(([k, v]) => cache.set(k, v)), remove: k => cache.delete(k) }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  SpreadsheetApp: {
    openById: id => {
      if (id === 'db') return dbBook;
      if (books[id]) return books[id];
      throw new Error('no access to ' + id);
    },
    create: name => {
      creates++;
      const sh = new Sheet('Sheet1');
      dbBook = { getId: () => 'db', getSheets: () => [sh], getSheetByName: n => (n === sh.name ? sh : null) };
      return dbBook;
    },
  },
  Utilities: { sleep() {}, getUuid: () => 'id-' + (++uuid), formatDate: () => '2026-10-14 09:00:00' },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: t => ({ setMimeType: () => JSON.parse(t) }) },
  Logger: { log: (...a) => console.log('LOG', ...a) },
};
const f = new Function(...Object.keys(g), code + ';return {doGet,doPost,checkForm};')(...Object.values(g));
const get = q => f.doGet({ parameter: { pin: '1234', ...q } });
const post = b => f.doPost({ postData: { contents: JSON.stringify({ pin: '1234', ...b }) } });

// pin / request checks
assert.equal(f.doGet({ parameter: { action: 'board', pin: 'x' } }).error, 'bad_pin');
assert.equal(f.doPost({ postData: { contents: 'not json' } }).error, 'bad_request');
assert.equal(post({ action: 'nope' }).error, 'bad_request');

// the phone lookup gives only the four fields (plus the rework flag), never the table-only ones
const one = get({ no: 'V1150831005' }).order;
assert.deepEqual(Object.keys(one).sort(), ['code', 'name', 'no', 'qty', 'rework', 'unit']);
assert.equal(one.name, '2M_西洋梨型 裸包 (每包3公克)');

// dropdown options: people from column A, left-side headers only, green QC/拋光 columns ignored
const o = get({ action: 'options' }).options;
assert.deepEqual(o['作業人員'], ['小明', '小華']);
assert.deepEqual(o['機台'], ['240kg混料', '100kg混料', '小粉包機-1']);
assert.deepEqual(o['天/時'], ['天', '時']);
assert.deepEqual(o['進度說明'], ['完工', '作業中']);
assert.equal(Object.keys(o).some(k => /拋光|QC/.test(k)), false);

// board: every plan order in one call, with the extra plan columns, and nothing saved yet
let b = get({ action: 'board' });
assert.equal(creates, 1);
assert.deepEqual(b.rows, []);
assert.deepEqual(b.orders.map(x => x.no), ['V1150831005', 'V1150831006', 'X1150901001']);
assert.equal(b.orders[0].item, '4粉包');
assert.equal(b.orders[0].arrive, '9/11');
assert.equal(b.orders[2].rework, true);
assert.equal(b.orders[2].code, '4104000000001');

// save keyed by work-order number: plan data wins, 人數 counts the selected people
let r = post({ action: 'save', row: { 工令單號: 'v1150831005', 產品名稱: '被竄改', 機台: '小粉包機-1', 作業人員1: '小明', 作業人員2: '小華', 人數: '9', 領料: '是' } });
assert.equal(r.ok, true);
assert.equal(r.row['工令單號'], 'V1150831005');
assert.equal(r.row['產品名稱'], '2M_西洋梨型 裸包 (每包3公克)');
assert.equal(r.row['人數'], '2');
assert.equal(r.row['建立時間'], '2026-10-14 09:00:00');
const id = r.row.id;

// saving the same number again updates that row (no duplicate); untouched fields are kept; 人數 follows the people
r = post({ action: 'save', row: { 工令單號: 'V1150831005', 作業人員2: '', 進度說明: '作業中' } });
assert.equal(r.ok, true);
assert.equal(r.row.id, id);
assert.equal(r.row['人數'], '1');
assert.equal(r.row['機台'], '小粉包機-1');
assert.equal(r.row['領料'], '是');
assert.equal(r.row['進度說明'], '作業中');
r = post({ action: 'save', row: { 工令單號: 'V1150831005', 作業人員1: '' } });
assert.equal(r.row['人數'], '');

// the sort order is just another saved field and only touches that one row
r = post({ action: 'save', row: { 工令單號: 'V1150831005', 排序: '1500.5' } });
assert.equal(r.row['排序'], '1500.5');
assert.equal(r.row['機台'], '小粉包機-1'); // untouched fields stay
assert.ok(dbBook.getSheets()[0].cols >= 32, 'columns are added when the grid is too small');

// errors
assert.equal(post({ action: 'save', row: { 工令單號: 'V1150831999' } }).error, 'not_found');
assert.equal(post({ action: 'save', row: { 工令單號: 'H1150831005' } }).error, 'not_found'); // customer order numbers are not work orders
assert.equal(post({ action: 'save', row: { 工令單號: '12345' } }).error, 'bad_format');

// second order, then the board returns plan orders and saved rows together
assert.equal(post({ action: 'save', row: { 工令單號: 'V1150831006', 機台: '100kg混料' } }).ok, true);
b = get({ action: 'board' });
assert.equal(b.rows.length, 2);
assert.deepEqual(b.rows.map(x => x['工令單號']), ['V1150831005', 'V1150831006']);

// remove clears one order's fill-in data only
assert.equal(post({ action: 'remove', no: 'V1150831005' }).ok, true);
assert.deepEqual(get({ action: 'list' }).rows.map(x => x['工令單號']), ['V1150831006']);
assert.equal(post({ action: 'remove', no: 'V1150831005' }).error, 'row_missing');
assert.equal(get({ action: 'board' }).orders.length, 3);

// a saved order can still be edited after it left the plan sheet
sheets.plan.rows.splice(3, 1); cache.clear();
assert.equal(post({ action: 'save', row: { 工令單號: 'V1150831006', 備註: '單已移走' } }).ok, true);
assert.equal(post({ action: 'save', row: { 工令單號: 'V1150831006', 備註: 'x' } }).row['備註'], 'x');
assert.equal(creates, 1);

// a database made by an older version (shorter header, fewer columns) is upgraded in place
const old = dbBook.getSheets()[0];
const oldHead = old.rows[0].slice(0, 30);
old.rows = [oldHead, ...old.rows.slice(1).map(r => r.slice(0, 30))];
old.cols = 30;
const upgraded = get({ action: 'list' }).rows;
assert.equal(old.rows[0].length >= 32 && old.rows[0][31] === '排序', true);
assert.deepEqual(upgraded.map(x => x['工令單號']), ['V1150831006']);
assert.equal(upgraded[0]['排序'], '');

// the plan index is cached in chunks, so a big sheet still fits and the second call skips the sheet
cache.clear();
const big = new Sheet('訂單計畫單(NEW)', [planRows[0], planRows[1]]);
for (let i = 0; i < 800; i++) big.rows.push([`4104${String(i).padStart(9, '0')}`, '品名'.repeat(20) + i, '1,000', '包', `V${1150000000 + i}`, '4粉包', '9/1']);
books.plan = { getSheetByName: () => big, getSheets: () => [big] };
assert.equal(get({ action: 'board' }).orders.length, 800);
assert.ok(Number(cache.get('idx_n')) > 1, 'index should be split into several cache chunks');
big.rows.length = 2; // wipe the sheet: only the cache can still answer
assert.equal(get({ action: 'board' }).orders.length, 800);
f.checkForm();
console.log('form backend OK');
