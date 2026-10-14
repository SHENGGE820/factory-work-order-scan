// Checks the form side of apps-script/Code.gs (options / list / save / remove) against a tiny in-memory spreadsheet.
import fs from 'fs';
import assert from 'assert';
const code = fs.readFileSync(new URL('../apps-script/Code.gs', import.meta.url), 'utf8');

class Sheet {
  constructor(name, rows = []) { this.name = name; this.rows = rows.map(r => [...r]); this.max = Math.max(rows.length, 20); this.formats = []; }
  getMaxRows() { return this.max; }
  getLastRow() { return this.rows.length; }
  getDataRange() { return { getDisplayValues: () => this.rows.map(r => r.map(v => String(v ?? ''))) }; }
  setName(n) { this.name = n; }
  setFrozenRows() {}
  insertRowsAfter(_, n) { this.max += n; }
  deleteRow(n) { this.rows.splice(n - 1, 1); }
  getParent() { return { getUrl: () => 'https://docs.google.com/spreadsheets/d/db' }; }
  getRange(r, c, nr = 1, nc = 1) {
    const sh = this;
    if (r + nr - 1 > this.max) throw new Error('outside the sheet');
    return {
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
  ['', '', '', '', ''],
  ['產品編號', '產品名稱', '數量', '單位', '工令單號'],
  ['4104002200098', '2M_西洋梨型 裸包 (每包3公克)', '6,000', '包', 'V1150831005'],
  ['7101002200144', '2M_西洋梨 混料(每公斤)', '18', '公斤', 'V1150831006'],
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
  CacheService: { getScriptCache: () => ({ get: k => cache.get(k) ?? null, put: (k, v) => cache.set(k, v), remove: k => cache.delete(k) }) },
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
assert.equal(f.doGet({ parameter: { action: 'list', pin: 'x' } }).error, 'bad_pin');
assert.equal(f.doPost({ postData: { contents: 'not json' } }).error, 'bad_request');
assert.equal(post({ action: 'nope' }).error, 'bad_request');
// the phone lookup is unchanged
assert.equal(get({ no: 'V1150831005' }).order.name, '2M_西洋梨型 裸包 (每包3公克)');

// dropdown options: people from column A, left-side headers only, green QC/拋光 columns ignored
const o = get({ action: 'options' }).options;
assert.deepEqual(o['作業人員'], ['小明', '小華']);
assert.deepEqual(o['機台'], ['240kg混料', '100kg混料', '小粉包機-1']);
assert.deepEqual(o['人數'], ['1', '3', '4']);
assert.deepEqual(o['天/時'], ['天', '時']);
assert.deepEqual(o['進度說明'], ['完工', '作業中']);
assert.deepEqual(o['預計完工日'], ['1', '3']);
assert.equal(Object.keys(o).some(k => /拋光|QC/.test(k)), false);

// empty list creates the database spreadsheet once
assert.deepEqual(get({ action: 'list' }).rows, []);
assert.equal(creates, 1);

// save: order data always comes from the plan sheet, other fields from the form
let r = post({ action: 'save', row: { 工令單號: 'v1150831005', 產品名稱: '被竄改', 機台: '小粉包機-1', 作業人員1: '小明', 人數: '1', 領料: '是', 備註: '測試' } });
assert.equal(r.ok, true);
assert.equal(r.row['工令單號'], 'V1150831005');
assert.equal(r.row['產品名稱'], '2M_西洋梨型 裸包 (每包3公克)');
assert.equal(r.row['數量'], '6,000');
assert.equal(r.row['機台'], '小粉包機-1');
assert.equal(r.row['建立時間'], '2026-10-14 09:00:00');
const id = r.row.id;

// errors
assert.equal(post({ action: 'save', row: { 工令單號: 'V1150831999' } }).error, 'not_found');
assert.equal(post({ action: 'save', row: { 工令單號: 'H1150831005' } }).error, 'not_found'); // customer order numbers are not work orders
assert.equal(post({ action: 'save', row: { 工令單號: '12345' } }).error, 'bad_format');
assert.equal(post({ action: 'save', row: { id: 'nope', 工令單號: 'V1150831005' } }).error, 'row_missing');

// edit: only the sent fields change, the rest is kept
r = post({ action: 'save', row: { id, 工令單號: 'V1150831005', 進度說明: '完工', 工令完工: '是' } });
assert.equal(r.ok, true);
assert.equal(r.row['進度說明'], '完工');
assert.equal(r.row['機台'], '小粉包機-1');
assert.equal(r.row['備註'], '測試');
assert.equal(r.row['id'], id);

// second order, then list / remove
assert.equal(post({ action: 'save', row: { 工令單號: 'V1150831006', 機台: '100kg混料' } }).ok, true);
let rows = get({ action: 'list' }).rows;
assert.equal(rows.length, 2);
assert.deepEqual(rows.map(x => x['工令單號']), ['V1150831005', 'V1150831006']);
assert.equal(post({ action: 'remove', id }).ok, true);
rows = get({ action: 'list' }).rows;
assert.deepEqual(rows.map(x => x['工令單號']), ['V1150831006']);
assert.equal(post({ action: 'remove', id }).error, 'row_missing');

// a saved row can still be edited after its order left the plan sheet
sheets.plan.rows.splice(3, 1); cache.clear();
const keep = rows[0].id;
assert.equal(post({ action: 'save', row: { id: keep, 工令單號: 'V1150831006', 備註: '單已移走' } }).ok, true);
assert.equal(creates, 1);
f.checkForm();
console.log('form backend OK');
