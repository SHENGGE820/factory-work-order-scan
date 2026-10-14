/**
 * 工令單查詢＋在途工單表單（Google Apps Script）
 *
 * 兩種放法，部署成「網頁應用程式」：
 *  1. 放在工令單試算表裡（擴充功能 → Apps Script）。
 *  2. 對試算表只有檢視權限時：在 script.google.com 另開專案，指令碼屬性加 SHEET_ID（試算表網址 /d/ 後面那串）。
 *
 * 手機 App：送來一個工令單號，只回傳四項（工令單號、產品名稱、產品編號、數量＋單位）。
 * 電腦表單（form.html）：用同一組工廠密碼，另外提供
 *   action=options  讀「BKE在途工單」的「選單」分頁，回傳下拉選項（含員工名字，所以不放在公開網頁裡）
 *   action=board    一次回傳要排的工令單（讀 BKE在途工單 的「計畫單平台」，只取代工廠＝五股廠且在訂單計畫單裡的單），加上已填的生產組資料（表格畫面用）
 *   action=list     只列出已填的生產組資料
 *   save / remove   以工令單號為鍵，存入或清除一張單的填寫內容（POST）
 * 資料存在這個帳號自己 Google 雲端的「五股廠生產組表單」試算表，第一次存檔時自動建立。
 *
 * 必要設定（專案設定 → 指令碼屬性）：PIN（工廠密碼）、SHEET_ID（訂單計畫單所在試算表）、
 * BKE_SHEET_ID（BKE在途工單試算表，只在表單要用下拉選項時需要）。這些都不要寫在這個檔案裡。
 */
const CONFIG = {
  // 只讀這個工作表；留空＝自動找出所有含「工令單號」標題的工作表。
  SHEET_NAME: '訂單計畫單(NEW)',
  HEADERS: { no: '工令單號', name: '產品名稱', code: '產品編號', qty: '數量', unit: '單位', item: '品項', arrive: '預計料齊到倉庫日' },
  HEADER_SEARCH_ROWS: 10,
  CACHE_SECONDS: 300, // 試算表改了之後，最多 5 分鐘內會反映到 App
};
const ORDER_NO = /[A-Z]\d{10}/g;
// 「BKE在途工單」各分頁的 交期、領料單號、到料時間、領料、工令完工、上階工單 原本都是用公式從「計畫單平台」依工令單號查出來的，
// 所以表格也直接讀這個分頁（欄位用表頭文字對應，不依位置）。
const PLATFORM = {
  SHEET: '計畫單平台',
  VENDOR: '五股廠',
  CACHE_SECONDS: 1200, // 這個分頁很大（五千多列），讀一次要好幾秒，所以多存一會；表格上有「重新讀取」可以略過快取
  COLS: { no: '工令單號', code: '產品編號', name: '產品名稱', qty: '數量', unit: '單位', item: '品項', arrive: '預計料齊到倉庫日',
    pickNo: '領料單號/調撥單號', picked: '領料', vendor: '代工廠', due: '交貨日', done: '工令完工', parent: '上階工令單', semi: '半成品完工' },
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  const bad = auth_(p);
  if (bad) return bad;
  if (p.action === 'options') return json_({ ok: true, options: options_() });
  if (p.action === 'list') return json_({ ok: true, rows: list_() });
  if (p.action === 'board') {
    if (p.fresh) { cacheDel_('idx'); cacheDel_('plat'); }
    return json_({ ok: true, orders: board_(), rows: list_() });
  }
  const no = String(p.no || '').trim().toUpperCase();
  if (!/^[A-Z]\d{10}$/.test(no)) return json_({ ok: false, error: 'bad_format' });
  const o = index_()[no];
  // 手機 App 只拿得到這四項（加重工標記），其餘欄位只給表格畫面（action=board）
  return json_(o ? { ok: true, order: { no: o.no, name: o.name, code: o.code, qty: o.qty, unit: o.unit, rework: o.rework } } : { ok: false, error: 'not_found' });
}

// 寫入一律用 POST（本文是 JSON 文字）：{ action:'save', pin, row:{工令單號,...} } 或 { action:'remove', pin, no }
function doPost(e) {
  let p;
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  const bad = auth_(p);
  if (bad) return bad;
  if (p.action === 'save') return json_(save_(p.row || {}));
  if (p.action === 'remove') return json_(remove_(String(p.no || '').trim().toUpperCase()));
  return json_({ ok: false, error: 'bad_request' });
}

function auth_(p) {
  const pin = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!pin) return json_({ ok: false, error: 'not_configured' });
  if (String(p.pin || '') !== pin) {
    Utilities.sleep(800);
    return json_({ ok: false, error: 'bad_pin' });
  }
  return null;
}

// 快取單筆上限約 100KB，全部工令單放不下，所以切段存放。
function cachePut_(key, text, seconds) {
  const cache = CacheService.getScriptCache();
  const size = 30000, n = Math.ceil(text.length / size);
  const all = { [key + '_n']: String(n) };
  for (let i = 0; i < n; i++) all[key + '_' + i] = text.slice(i * size, (i + 1) * size);
  try { cache.putAll(all, seconds || CONFIG.CACHE_SECONDS); } catch (err) { /* 放不進去就每次直接讀表 */ }
}
function cacheDel_(key) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(key + '_n') || 0), keys = [key + '_n'];
  for (let i = 0; i < n; i++) keys.push(key + '_' + i);
  cache.removeAll(keys);
}
function cacheGet_(key) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(key + '_n') || 0);
  if (!n) return null;
  let text = '';
  for (let i = 0; i < n; i++) {
    const part = cache.get(key + '_' + i);
    if (part == null) return null;
    text += part;
  }
  return text;
}

function index_() {
  const hit = cacheGet_('idx');
  if (hit) return JSON.parse(hit);

  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
  const sheets = CONFIG.SHEET_NAME ? [ss.getSheetByName(CONFIG.SHEET_NAME)] : ss.getSheets();
  const idx = {};
  sheets.filter(Boolean).forEach(sh => {
    const rows = sh.getDataRange().getDisplayValues();
    const h = findHeader_(rows);
    if (!h) return;
    for (let r = h.row + 1; r < rows.length; r++) {
      const row = rows[r];
      const noCell = String(row[h.col.no] || '');
      const nos = noCell.toUpperCase().match(ORDER_NO);
      if (!nos) continue;
      const codeCell = h.col.code >= 0 ? String(row[h.col.code] || '').trim() : '';
      const rework = /重工/.test(noCell) || /重工/.test(codeCell);
      const order = {
        name: h.col.name >= 0 ? String(row[h.col.name] || '').trim() : '',
        code: codeCell.replace(/^重工\s*-?\s*/, '').trim(),
        qty: h.col.qty >= 0 ? String(row[h.col.qty] || '').trim() : '',
        unit: h.col.unit >= 0 ? String(row[h.col.unit] || '').trim() : '',
        item: h.col.item >= 0 ? String(row[h.col.item] || '').trim() : '',
        arrive: h.col.arrive >= 0 ? String(row[h.col.arrive] || '').trim() : '',
        rework,
      };
      nos.forEach(no => { idx[no] = Object.assign({ no }, order); });
    }
  });
  cachePut_('idx', JSON.stringify(idx));
  return idx;
}

// 表格要排的單：計畫單平台裡「代工廠＝五股廠」且仍在訂單計畫單(NEW)的單。讀不到計畫單平台時，退回用訂單計畫單。
function board_() {
  const idx = index_();
  const plat = platform_(idx);
  if (!plat.length) return Object.keys(idx).map(k => idx[k]);
  return plat.map(o => Object.assign({}, o, { rework: !!(idx[o.no] && idx[o.no].rework) }));
}

function platform_(idx) {
  const hit = cacheGet_('plat');
  if (hit) return JSON.parse(hit);
  const id = PropertiesService.getScriptProperties().getProperty('BKE_SHEET_ID');
  const sh = id ? SpreadsheetApp.openById(id).getSheetByName(PLATFORM.SHEET) : null;
  if (!sh) return [];
  const clean = v => String(v).replace(/\s/g, '');
  const width = Math.min(sh.getLastColumn(), 40);
  const rows = sh.getRange(1, 1, sh.getLastRow(), width).getDisplayValues();
  let hr = -1;
  for (let r = 0; r < Math.min(6, rows.length); r++) if (rows[r].map(clean).indexOf(PLATFORM.COLS.no) >= 0) { hr = r; break; }
  if (hr < 0) return [];
  const head = rows[hr].map(clean), col = {};
  Object.keys(PLATFORM.COLS).forEach(k => { col[k] = head.indexOf(PLATFORM.COLS[k]); });
  const cell = (row, k) => (col[k] >= 0 ? String(row[col[k]] || '').trim() : '');
  const out = [], seen = {};
  for (let r = hr + 1; r < rows.length; r++) {
    const row = rows[r];
    const no = (cell(row, 'no').toUpperCase().match(ORDER_NO) || [])[0];
    if (!no || seen[no] || !idx[no] || cell(row, 'vendor') !== PLATFORM.VENDOR) continue;
    seen[no] = true;
    out.push({ no, code: cell(row, 'code'), name: cell(row, 'name'), qty: cell(row, 'qty'), unit: cell(row, 'unit'),
      item: cell(row, 'item'), arrive: cell(row, 'arrive'), pickNo: cell(row, 'pickNo'), picked: cell(row, 'picked') === 'TRUE',
      due: cell(row, 'due'), done: cell(row, 'done') === 'TRUE', parent: cell(row, 'parent'), semi: cell(row, 'semi') });
  }
  cachePut_('plat', JSON.stringify(out), PLATFORM.CACHE_SECONDS);
  return out;
}

function findHeader_(rows) {
  const clean = v => String(v).replace(/\s/g, '');
  for (let r = 0; r < Math.min(CONFIG.HEADER_SEARCH_ROWS, rows.length); r++) {
    const cells = rows[r].map(clean);
    const col = {};
    Object.keys(CONFIG.HEADERS).forEach(k => { col[k] = cells.indexOf(CONFIG.HEADERS[k]); });
    if (col.no >= 0) return { row: r, col };
  }
  return null;
}

// ---- 在途工單表單：五股廠生產組 ----
const DB = {
  NAME: '五股廠生產組表單',
  SHEET: '生產組',
  // 欄位依「BKE在途工單」的「五股廠(生產組)」分頁（綠色欄位除外）。
  FIELDS: ['工令單號', '產品編號', '產品名稱', '數量', '單位', '出貨點', '交期',
    '作業人員1', '作業人員2', '作業人員3', '作業人員4', '人數', '機台', '預計完工日', '天/時',
    '製造日期', '生產進度', '完工時間', '實際用時(時)', '實際用時(分)', '進度說明',
    '答交', '作業時間', '領料單號/調撥單號', '到料時間', '領料', '工令完工', '備註', '排序', '完工資訊', '分區'],
  FROM_ORDER: { '產品編號': 'code', '產品名稱': 'name', '數量': 'qty', '單位': 'unit' }, // 一律以訂單計畫單為準
  // 「選單」分頁表頭 → 下拉選項名稱；第一欄（表頭是 0）是作業人員。
  OPTION_HEADERS: { '人數': '人數', '生產機台': '機台', '預計完工日': '預計完工日', '天/時': '天/時', '進度說明': '進度說明' },
  OPTIONS_SHEET: '選單',
  TZ: 'Asia/Taipei',
};
const DB_HEAD = ['id', '建立時間', '更新時間'].concat(DB.FIELDS);

function db_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('DB_ID');
  let sh;
  if (id) {
    const ss = SpreadsheetApp.openById(id);
    sh = ss.getSheetByName(DB.SHEET) || ss.getSheets()[0];
    if (sh.getName() !== DB.SHEET) sh.setName(DB.SHEET);
  } else {
    const ss = SpreadsheetApp.create(DB.NAME);
    props.setProperty('DB_ID', ss.getId());
    sh = ss.getSheets()[0];
    sh.setName(DB.SHEET);
  }
  ensureDb_(sh);
  return sh;
}

// 新試算表只有 26 欄，這裡要的欄數更多，所以先補欄；表頭不一致（例如新增了欄位）也一併補齊。
function ensureDb_(sh) {
  const need = DB_HEAD.length;
  if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
  const have = sh.getRange(1, 1, 1, need).getValues()[0].map(String);
  if (have.join('|') !== DB_HEAD.join('|')) {
    sh.getRange(1, 1, sh.getMaxRows(), need).setNumberFormat('@'); // 全部當文字，編號不會被改成數字
    sh.getRange(1, 1, 1, need).setValues([DB_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

function rowToObj_(r) {
  const o = {};
  DB_HEAD.forEach((h, i) => { o[h] = r[i] == null ? '' : r[i]; });
  return o;
}

function list_() {
  const rows = db_().getDataRange().getDisplayValues();
  return rows.slice(1).filter(r => r[0]).map(rowToObj_);
}

function save_(row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = db_();
    const data = sh.getDataRange().getDisplayValues();
    const no = String(row['工令單號'] || '').trim().toUpperCase();
    if (!/^[A-Z]\d{10}$/.test(no)) return { ok: false, error: 'bad_format' };
    const noCol = DB_HEAD.indexOf('工令單號');
    let at = -1;
    for (let i = 1; i < data.length; i++) if (data[i][noCol] === no) { at = i; break; }
    const old = at >= 0 ? data[at] : null;
    const order = index_()[no];
    // 新的單一定要在訂單計畫單查得到；已存過的單，就算已移到完工分頁也能繼續改
    if (!order && !old) return { ok: false, error: 'not_found' };
    const now = Utilities.formatDate(new Date(), DB.TZ, 'yyyy-MM-dd HH:mm:ss');
    const out = DB_HEAD.map((h, i) => {
      if (h === 'id') return old ? old[0] : Utilities.getUuid();
      if (h === '建立時間') return old ? old[1] : now;
      if (h === '更新時間') return now;
      if (h === '工令單號') return no;
      if (order && DB.FROM_ORDER[h]) return String(order[DB.FROM_ORDER[h]] || '');
      if (Object.prototype.hasOwnProperty.call(row, h)) return String(row[h] == null ? '' : row[h]).slice(0, 500);
      return old && old[i] != null ? old[i] : '';
    });
    // 人數＝選了幾位作業人員
    const people = [1, 2, 3, 4].filter(n => String(out[DB_HEAD.indexOf('作業人員' + n)] || '').trim()).length;
    out[DB_HEAD.indexOf('人數')] = people ? String(people) : '';
    const target = at >= 0 ? at + 1 : data.length + 1;
    if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200);
    const range = sh.getRange(target, 1, 1, DB_HEAD.length);
    range.setNumberFormat('@');
    range.setValues([out]);
    return { ok: true, row: rowToObj_(out) };
  } finally {
    lock.releaseLock();
  }
}

// 清除一張單的填寫內容（訂單計畫單那邊不受影響）
function remove_(no) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = db_();
    const data = sh.getDataRange().getDisplayValues();
    const noCol = DB_HEAD.indexOf('工令單號');
    for (let i = 1; i < data.length; i++) {
      if (data[i][noCol] === no) { sh.deleteRow(i + 1); return { ok: true }; }
    }
    return { ok: false, error: 'row_missing' };
  } finally {
    lock.releaseLock();
  }
}

function options_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('opts');
  if (hit) return JSON.parse(hit);
  const id = PropertiesService.getScriptProperties().getProperty('BKE_SHEET_ID');
  const sh = id ? SpreadsheetApp.openById(id).getSheetByName(DB.OPTIONS_SHEET) : null;
  const out = {};
  if (sh) {
    const rows = sh.getDataRange().getDisplayValues();
    const head = rows[0].map(h => String(h).replace(/\s/g, ''));
    const column = c => {
      const seen = [];
      for (let r = 1; r < rows.length; r++) {
        const v = String(rows[r][c] || '').trim();
        if (v && seen.indexOf(v) < 0) seen.push(v);
      }
      return seen;
    };
    out['作業人員'] = column(0);
    Object.keys(DB.OPTION_HEADERS).forEach(h => {
      const c = head.indexOf(h);
      if (c >= 0) out[DB.OPTION_HEADERS[h]] = column(c); // 右側包裝組也有同名表頭，取最左邊那個
    });
  }
  try { cache.put('opts', JSON.stringify(out), CONFIG.CACHE_SECONDS); } catch (err) { /* ignore */ }
  return out;
}

// ---- 一次性匯入：把「五股廠(生產組)」現有的填寫內容帶進表格 ----
// 只帶「目前表格裡有的單」（計畫單平台的五股廠單），而且只帶標準型粉包機／加大型粉包機／混料三區（膠囊先不帶）。
// 已經填過的單不會被蓋掉。dry=true 只統計不寫入。在編輯器裡執行（見 CLAUDE.md），不會對外開放。
function importProdSheet_(dry) {
  const props = PropertiesService.getScriptProperties();
  const bke = SpreadsheetApp.openById(props.getProperty('BKE_SHEET_ID'));
  const sh = bke.getSheetByName('五股廠(生產組)');
  const last = sh.getLastRow(), w = Math.min(sh.getLastColumn(), 40);
  const vals = sh.getRange(1, 1, last, w).getDisplayValues();
  const bg = sh.getRange(1, 1, last, 1).getBackgrounds();
  const head = vals[0].map(h => String(h).replace(/\s/g, ''));
  const col = re => head.findIndex(h => re.test(h));
  const C = { no: head.indexOf('工令單號/銷貨單號'), mach: col(/^機台/), eta: col(/^預計完工日/), unit: col(/^天\/時/), made: col(/^製造日期/), prog: col(/^生產進度/),
    fin: col(/^完工時間/), hrs: col(/^實際用時.*giờ/), min: col(/^實際用時.*phút/), status: col(/^進度說明/), ship: head.indexOf('出貨點'),
    ans: head.indexOf('答交'), work: head.indexOf('作業時間'), note: head.indexOf('備註'), info: head.indexOf('完工資訊') };
  const idx = index_(), board = {};
  platform_(idx).forEach(o => { board[o.no] = o; });
  const existing = {}; list_().forEach(r => { existing[r['工令單號']] = true; });
  const day = v => { const m = String(v).match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/); return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : String(v).trim(); };
  const num = v => String(v).replace(/,/g, '').trim();
  const optPeople = options_()['作業人員'] || [], optMach = options_()['機台'] || [];
  const now = Utilities.formatDate(new Date(), DB.TZ, 'yyyy-MM-dd HH:mm:ss');
  const SECTION = { std: /標準型/, big: /加大型/, mix: /^混料$/, cap: /膠囊/ };
  let sec = 'cap', waiting = false;
  const st = { sheetRows: {}, inBoard: {}, imported: {}, existing: 0, dup: 0, done: 0, waiting: 0, unknownPeople: {}, unknownMach: {}, notInBoard: 0 };
  const bump = (o, k) => { o[k] = (o[k] || 0) + 1; };
  const seen = {}, out = [], sample = [];
  for (let r = 1; r < last; r++) {
    const text = vals[r].map(x => String(x).trim()).filter(Boolean);
    if (String(bg[r][0]).toLowerCase() === '#000000') {
      const t = text.join(' '); sec = Object.keys(SECTION).find(k => SECTION[k].test(t)) || 'cap'; waiting = false; continue;
    }
    if (vals[r].some(x => String(x).trim() === '等待安排')) { waiting = true; continue; }
    const no = (String(vals[r][C.no]).toUpperCase().match(/[A-Z]\d{10}/) || [])[0];
    if (!no) continue;
    bump(st.sheetRows, sec);
    if (sec === 'cap') continue;
    if (!board[no]) { st.notInBoard++; continue; }
    bump(st.inBoard, sec);
    if (seen[no]) { st.dup++; continue; }
    seen[no] = true;
    if (existing[no]) { st.existing++; continue; }
    const people = [0, 1, 2, 3].map(i => String(vals[r][i]).trim()).filter(Boolean).slice(0, 4);
    people.forEach(p => { if (optPeople.indexOf(p) < 0) bump(st.unknownPeople, p); });
    let mach = waiting ? '' : String(vals[r][C.mach]).trim();
    if (mach && optMach.indexOf(mach) < 0) bump(st.unknownMach, mach);
    const rec = { 工令單號: no, 分區: sec, 機台: mach, 預計完工日: String(vals[r][C.eta]).trim(), '天/時': String(vals[r][C.unit]).trim(), 製造日期: day(vals[r][C.made]),
      生產進度: num(vals[r][C.prog]), 完工時間: day(vals[r][C.fin]), '實際用時(時)': num(vals[r][C.hrs]), '實際用時(分)': num(vals[r][C.min]), 進度說明: String(vals[r][C.status]).trim(),
      出貨點: String(vals[r][C.ship]).trim(), 答交: String(vals[r][C.ans]).trim(), 作業時間: String(vals[r][C.work]).trim(), 備註: String(vals[r][C.note]).trim(), 完工資訊: String(vals[r][C.info]).trim(),
      排序: String(r + 1), 人數: people.length ? String(people.length) : '' };
    people.forEach((p, i) => { rec['作業人員' + (i + 1)] = p; });
    if (board[no].done || rec['進度說明'] === '完工') st.done++;
    if (waiting) st.waiting++;
    bump(st.imported, sec);
    out.push(rec);
    if (sample.length < 3) sample.push(JSON.stringify(rec));
  }
  Logger.log('@@sheet rows per section ' + JSON.stringify(st.sheetRows) + ' | not in board ' + st.notInBoard);
  Logger.log('@@in board ' + JSON.stringify(st.inBoard) + ' | to import ' + JSON.stringify(st.imported) + ' | dup ' + st.dup + ' | already saved ' + st.existing + ' | done ' + st.done + ' | waiting ' + st.waiting);
  Logger.log('@@people not in dropdown ' + JSON.stringify(st.unknownPeople) + ' | machines not in dropdown ' + JSON.stringify(st.unknownMach));
  sample.forEach(s => Logger.log('@@sample ' + s.slice(0, 300)));
  if (dry || !out.length) return out.length;
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const db = db_(), n = db.getLastRow();
    const rows = out.map(rec => DB_HEAD.map(h => {
      if (h === 'id') return Utilities.getUuid();
      if (h === '建立時間' || h === '更新時間') return now;
      const o = idx[rec['工令單號']];
      if (o && DB.FROM_ORDER[h]) return String(o[DB.FROM_ORDER[h]] || '');
      return rec[h] == null ? '' : String(rec[h]);
    }));
    if (n + rows.length > db.getMaxRows()) db.insertRowsAfter(db.getMaxRows(), n + rows.length - db.getMaxRows() + 50);
    const range = db.getRange(n + 1, 1, rows.length, DB_HEAD.length);
    range.setNumberFormat('@'); range.setValues(rows);
  } finally { lock.releaseLock(); }
  Logger.log('@@imported ' + out.length + ' rows');
  return out.length;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 在編輯器裡選這個函式按「執行」，可檢查讀到幾張單（結果只會顯示在你自己的執行紀錄）。 */
function checkSetup() {
  cacheDel_('idx');
  const idx = index_();
  const nos = Object.keys(idx);
  Logger.log('讀到 %s 張工令單。PIN 已設定：%s', nos.length, !!PropertiesService.getScriptProperties().getProperty('PIN'));
  if (nos.length) Logger.log('第一張：%s', JSON.stringify(idx[nos[0]]));
}

/** 檢查表單這一側：建立（或找到）「五股廠生產組表單」試算表，並讀「選單」下拉選項。 */
function checkForm() {
  CacheService.getScriptCache().remove('opts');
  cacheDel_('plat');
  const sh = db_();
  Logger.log('表單試算表：%s（目前 %s 筆）', sh.getParent().getUrl(), Math.max(0, sh.getLastRow() - 1));
  const o = options_();
  Logger.log('計畫單平台（五股廠）：%s 張', platform_(index_()).length);
  Logger.log('下拉選項：%s', Object.keys(o).map(k => k + ' ' + o[k].length + ' 個').join('、') || '讀不到（請設定 BKE_SHEET_ID）');
}
