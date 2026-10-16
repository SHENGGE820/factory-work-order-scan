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
  getLastColumn() { return Math.max(0, ...this.rows.map(r => r.length)); }
  getDataRange() { return { getDisplayValues: () => this.rows.map(r => r.map(v => String(v ?? ''))) }; }
  setName(n) { this.name = n; }
  setFrozenRows() {}
  insertRowsAfter(_, n) { this.max += n; }
  deleteRow(n) { this.rows.splice(n - 1, 1); }
  getParent() { return this.book || { getUrl: () => 'https://docs.google.com/spreadsheets/d/db' }; }
  getRange(r, c, nr = 1, nc = 1) {
    const sh = this;
    if (r + nr - 1 > this.max || c + nc - 1 > this.cols) throw new Error('outside the sheet');
    const at = (i, j) => sh.rows[r - 1 + i]?.[c - 1 + j] ?? '';
    return {
      getValues() { return Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => at(i, j))); },
      getDisplayValues() { return Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => String(at(i, j)))); },
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
// 計畫單平台: same columns as the plan sheet plus 領料單號, 領料, 代工廠, 交貨日, 工令完工, 上階工令單 (matched by header text)
const PH = ['日期', '', '', '', '', '產品編號', '產品名稱', '數量', '單位', '工令單號', '', '', 'BOM確單', '工令', '品項', '預計\n料齊到\n倉庫日', '備註',
  '領料單號/調撥單號', '領料', '代工廠', '', '確認交貨日', '交期YYY.MM', '交貨日', '', '銷貨單號', '', '組合', '工令完工', '上階工令單', '半成品完工'];
const prow = o => { const r = Array(PH.length).fill(''); const put = (h, v) => { r[PH.indexOf(h)] = v; };
  put('工令單號', o.no); put('產品編號', o.code || ''); put('產品名稱', o.name || ''); put('數量', o.qty || ''); put('單位', o.unit || ''); put('品項', o.item || '');
  put('預計\n料齊到\n倉庫日', o.arrive || ''); put('領料單號/調撥單號', o.pickNo || ''); put('領料', o.picked || 'FALSE'); put('代工廠', o.vendor || '');
  put('交貨日', o.due || ''); put('工令完工', o.done || 'FALSE'); put('上階工令單', o.parent || ''); put('半成品完工', o.semi || ''); return r; };
const platRows = [
  ['計畫單平台'],
  PH,
  prow({ no: 'V1150831005', code: '4104002200098', name: '2M_西洋梨型 裸包 (每包3公克)', qty: '6,000', unit: '包', item: '4粉包', arrive: '9/11', pickNo: 'W1150903013', picked: 'TRUE', vendor: '五股廠', due: '115.1001' }),
  prow({ no: 'V1150831006', code: '7101002200144', name: '2M_西洋梨 混料(每公斤)', qty: '18', unit: '公斤', item: '7混料', arrive: '00', vendor: '五股廠', done: 'TRUE', parent: 'V1150831005' }),
  prow({ no: 'X1150901001', code: '4104000000001', name: '重工品', qty: '5', unit: '包', item: '5包裝', vendor: '全宏' }),
  prow({ no: 'V1150831005', vendor: '五股廠', due: 'duplicate row is ignored' }),
  prow({ no: 'V9999999999', vendor: '五股廠', item: '7混料' }), // not in the plan sheet any more
];
const sheets = { plan: new Sheet('訂單計畫單(NEW)', planRows), opts: new Sheet('選單', optRows), plat: new Sheet('計畫單平台', platRows) };
Object.values(sheets).forEach(sh => { sh.cols = 71; }); // the real 計畫單平台 has 71 columns
const books = {
  plan: { getSheetByName: n => n === '訂單計畫單(NEW)' ? sheets.plan : null, getSheets: () => [sheets.plan] },
  bke: { getSheetByName: n => ({ '選單': sheets.opts, '計畫單平台': sheets.plat }[n] || null) },
};
let dbBook = null, creates = 0;
const props = { PIN: '1234', SHEET_ID: 'plan', BKE_SHEET_ID: 'bke' };
const cache = new Map();
let uuid = 0;
const g = {
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
  CacheService: { getScriptCache: () => ({ get: k => cache.get(k) ?? null, put: (k, v) => cache.set(k, v), putAll: o => Object.entries(o).forEach(([k, v]) => cache.set(k, v)), remove: k => cache.delete(k), removeAll: ks => ks.forEach(k => cache.delete(k)) }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  SpreadsheetApp: {
    openById: id => {
      if (id === 'db') return dbBook;
      if (books[id]) return books[id];
      throw new Error('no access to ' + id);
    },
    create: name => {
      creates++;
      const sh = new Sheet('Sheet1'), all = [sh];
      dbBook = { getId: () => 'db', getUrl: () => 'https://docs.google.com/spreadsheets/d/db', getSheets: () => all, getSheetByName: n => all.find(x => x.name === n) || null,
        insertSheet: n => { const x = new Sheet(n); x.book = dbBook; all.push(x); return x; } };
      sh.book = dbBook;
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
assert.deepEqual(Object.keys(one).sort(), ['code', 'loss', 'machine', 'name', 'no', 'people', 'prepped', 'qty', 'rework', 'unit']);
assert.equal(one.machine, ''); assert.deepEqual(one.people, []);
assert.equal(one.prepped, false);
assert.equal(one.loss, 0);
assert.equal(one.name, '2M_西洋梨型 裸包 (每包3公克)');

// dropdown options: people from column A, left-side headers only, green QC/拋光 columns ignored
const o = get({ action: 'options' }).options;
assert.deepEqual(o['作業人員'], ['小明', '小華']);
assert.deepEqual(o['機台'], ['240kg混料', '100kg混料', '小粉包機-1']);
assert.deepEqual(o['天/時'], ['天', '時']);
assert.deepEqual(o['進度說明'], ['完工', '作業中']);
assert.equal(Object.keys(o).some(k => /拋光|QC/.test(k)), false);

// board: orders come from 計畫單平台 (代工廠 = 五股廠, still in the plan sheet), with the columns the form shows automatically
let b = get({ action: 'board' });
assert.equal(creates, 1);
assert.deepEqual(b.rows, []);
assert.deepEqual(b.orders.map(x => x.no), ['V1150831005', 'V1150831006']);
assert.deepEqual(b.orders[0], { no: 'V1150831005', code: '4104002200098', name: '2M_西洋梨型 裸包 (每包3公克)', qty: '6,000', unit: '包', item: '4粉包', arrive: '9/11',
  pickNo: 'W1150903013', picked: true, due: '115.1001', done: false, parent: '', semi: '', rework: false });
assert.equal(b.orders[1].done, true);
assert.equal(b.orders[1].parent, 'V1150831005');

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
assert.equal(get({ action: 'board' }).orders.length, 2);

// the platform sheet is cached (it is big); fresh=1 reads it again
sheets.plat.rows[2][PH.indexOf('領料')] = 'FALSE';
assert.equal(get({ action: 'board' }).orders[0].picked, true, 'served from cache');
assert.equal(get({ action: 'board', fresh: '1' }).orders[0].picked, false, 'fresh read');

// without the platform sheet the board falls back to the plan sheet's orders
const bkeId = props.BKE_SHEET_ID; delete props.BKE_SHEET_ID; cache.clear();
assert.deepEqual(get({ action: 'board' }).orders.map(x => x.no), ['V1150831005', 'V1150831006', 'X1150901001']);
props.BKE_SHEET_ID = bkeId; cache.clear();

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

// phone timer: each finished stage is one segment in 工時紀錄; the order's row in the table gets the totals
const T0 = Date.UTC(2026, 9, 14, 1, 0, 0), M = 60000;
post({ action: 'save', row: { 工令單號: 'V1150831005', 機台: '小粉包機-1', 作業人員1: '小明', 作業人員2: '小華' } });
let w = post({ action: 'work', seg: { id: 's1', no: 'V1150831005', phase: 'prep', start: T0, end: T0 + 30 * M, loss: 50 } });
assert.equal(w.ok, true);
assert.equal(w.sum.allow, '180'); // 6,000 × 3%
assert.equal(post({ action: 'work', seg: { id: 's1', no: 'V1150831005', phase: 'prep', start: T0, end: T0 + 30 * M, loss: 50 } }).sum.loss, 50, 'a resent segment is not counted twice');
post({ action: 'work', seg: { id: 's2', no: 'V1150831005', phase: 'run', start: T0 + 30 * M, end: T0 + 90 * M } });
post({ action: 'work', seg: { id: 's3', no: 'V1150831005', phase: 'stop', reasonText: '等料', start: T0 + 90 * M, end: T0 + 100 * M } });
w = post({ action: 'work', seg: { id: 's4', no: 'V1150831005', phase: 'run', start: T0 + 100 * M, end: T0 + 125 * M, done: 5900 } });
assert.deepEqual(w.sum, { prep: 30, run: 85, stop: 10, loss: 50, done: 5900, allow: '180' });
const ws = dbBook.getSheetByName('工時紀錄');
assert.equal(ws.rows.length, 5);
assert.deepEqual(ws.rows[3].slice(1, 12), ['V1150831005', '2M_西洋梨型 裸包 (每包3公克)', '停機', '等料', '2026-10-14 09:00:00', '2026-10-14 09:00:00', '10', '', '', '小粉包機-1', '小明、小華']);
const tr = get({ action: 'list' }).rows.find(x => x['工令單號'] === 'V1150831005');
assert.equal(tr['前置(分)'], '30'); assert.equal(tr['上機(分)'], '85'); assert.equal(tr['停機(分)'], '10');
assert.equal(tr['損耗'], '50'); assert.equal(tr['可損耗'], '180'); assert.equal(tr['完成數量'], '5900'); assert.equal(tr['生產進度'], '5900');
assert.equal(tr['實際用時(時)'], '2'); assert.equal(tr['實際用時(分)'], '5');
assert.equal(tr['製造日期'], '2026-10-14'); assert.equal(tr['完工時間'], '2026-10-14');
assert.equal(tr['機台'], '小粉包機-1'); // the table's own fields stay
assert.equal(get({ no: 'V1150831005' }).order.machine, '小粉包機-1');
assert.deepEqual(get({ no: 'V1150831005' }).order.people, ['小明', '小華']);
// a segment that says who really worked and on which machine wins over the table's plan
post({ action: 'work', seg: { id: 's5', no: 'V1150831005', phase: 'run', start: T0 + 130 * M, end: T0 + 140 * M, people: ['小華'], machine: '小粉包機-2' } });
assert.deepEqual(ws.rows[5].slice(10, 12), ['小粉包機-2', '小華']);
const again = get({ no: 'V1150831005' }).order;
assert.equal(again.prepped, true); assert.equal(again.loss, 50);
// bad segments are refused
assert.equal(post({ action: 'work', seg: { id: 'x', no: 'V1150831005', phase: 'lunch', start: T0, end: T0 + M } }).error, 'bad_request');
assert.equal(post({ action: 'work', seg: { id: 'x', no: 'V1150831005', phase: 'run', start: T0 + M, end: T0 } }).error, 'bad_request');
assert.equal(post({ action: 'work', seg: { id: 'x', no: 'V1150831005', phase: 'prep', start: T0, end: T0 + M, loss: -3 } }).error, 'bad_request');
assert.equal(post({ action: 'work', seg: { id: 'x', no: 'V1150831999', phase: 'run', start: T0, end: T0 + M } }).error, 'not_found');

// manager report: segments of a date range plus the orders they belong to
let rp = get({ action: 'report', from: '2026-10-14', to: '2026-10-14' });
assert.equal(rp.ok, true);
assert.equal(rp.segs.length, 5);
assert.deepEqual(rp.segs[2], { no: 'V1150831005', phase: 'stop', reason: '等料', start: '2026-10-14 09:00:00', end: '2026-10-14 09:00:00', min: 10, loss: null, done: null, machine: '小粉包機-1', people: ['小明', '小華'] });
assert.deepEqual(rp.orders.V1150831005, { name: '2M_西洋梨型 裸包 (每包3公克)', code: '4104002200098', qty: 6000, unit: '包', item: '4粉包', lossAll: 50 });
assert.equal(get({ action: 'report', from: '2026-10-15', to: '2026-10-20' }).segs.length, 0);
assert.equal(get({ action: 'report', from: 'x', to: '2026-10-20' }).error, 'bad_request');
assert.equal(f.doGet({ parameter: { action: 'report', from: '2026-10-14', to: '2026-10-14' } }).ok, true); // the report needs no password
// cost settings for the profit part of the report
assert.deepEqual(rp.settings, { wage: 200, markup: 0.3, machineCost: { '*': 0 }, material: { '*': 0 }, fee: {} });
const sv = post({ action: 'settings', settings: { wage: '210', markup: 0.25, machineCost: { '*': 40, '小粉包機-1': '55' }, material: { '包': 0.4, x: 'abc' }, fee: { '4粉包': 1.2, '7混料': '' } } });
assert.equal(sv.ok, true);
assert.deepEqual(sv.settings, { wage: 210, markup: 0.25, machineCost: { '*': 40, '小粉包機-1': 55 }, material: { '包': 0.4 }, fee: { '4粉包': 1.2 } });
assert.equal(get({ action: 'report', from: '2026-10-14', to: '2026-10-14' }).settings.machineCost['小粉包機-1'], 55);
assert.equal(post({ action: 'settings', settings: { wage: -1, markup: 0.3 } }).error, 'bad_request');
assert.equal(f.doPost({ postData: { contents: JSON.stringify({ action: 'settings', settings: { wage: 1, markup: 0 } }) } }).ok, true); // nor do the settings

// staff list for the shared phone
assert.deepEqual(get({ action: 'people' }).people, ['小明', '小華']);
assert.deepEqual(get({ action: 'people' }).machines, ['240kg混料', '100kg混料', '小粉包機-1']);

// the plan index is cached in chunks, so a big sheet still fits and the second lookup skips the sheet
cache.clear();
const big = new Sheet('訂單計畫單(NEW)', [planRows[0], planRows[1]]);
for (let i = 0; i < 800; i++) big.rows.push([`4104${String(i).padStart(9, '0')}`, '品名'.repeat(20) + i, '1,000', '包', `V${1150000000 + i}`, '4粉包', '9/1']);
books.plan = { getSheetByName: () => big, getSheets: () => [big] };
assert.equal(get({ no: 'V1150000799' }).order.no, 'V1150000799');
assert.ok(Number(cache.get('idx_n')) > 1, 'index should be split into several cache chunks');
big.rows.length = 2; // wipe the sheet: only the cache can still answer
assert.equal(get({ no: 'V1150000001' }).order.no, 'V1150000001');
f.checkForm();
console.log('form backend OK');
