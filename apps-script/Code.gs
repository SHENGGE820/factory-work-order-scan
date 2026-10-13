/**
 * 工令單查詢（Google Apps Script）
 *
 * 放在工令單試算表裡（擴充功能 → Apps Script），部署成「網頁應用程式」。
 * 手機 App 送來一個工令單號，這裡只回傳四項：工令單號、產品名稱、產品編號、數量＋單位。
 * 不提供列出全部單據的功能，也不回傳其他欄位。
 *
 * 必要設定：專案設定 → 指令碼屬性 → 新增 PIN（工廠密碼）。PIN 不要寫在這個檔案裡。
 */
const CONFIG = {
  // 留空＝自動找出所有含「工令單號」標題的工作表；也可以填單一工作表名稱。
  SHEET_NAME: '',
  HEADERS: { no: '工令單號', name: '產品名稱', code: '產品編號', qty: '數量', unit: '單位' },
  HEADER_SEARCH_ROWS: 10,
  CACHE_SECONDS: 300, // 試算表改了之後，最多 5 分鐘內會反映到 App
};
const ORDER_NO = /[A-Z]\d{10}/g;

function doGet(e) {
  const p = (e && e.parameter) || {};
  const pin = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!pin) return json_({ ok: false, error: 'not_configured' });
  if (String(p.pin || '') !== pin) {
    Utilities.sleep(800);
    return json_({ ok: false, error: 'bad_pin' });
  }
  const no = String(p.no || '').trim().toUpperCase();
  if (!/^[A-Z]\d{10}$/.test(no)) return json_({ ok: false, error: 'bad_format' });
  const order = index_()[no];
  return json_(order ? { ok: true, order } : { ok: false, error: 'not_found' });
}

function index_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('idx');
  if (hit) return JSON.parse(hit);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
        rework,
      };
      nos.forEach(no => { idx[no] = Object.assign({ no }, order); });
    }
  });
  try { cache.put('idx', JSON.stringify(idx), CONFIG.CACHE_SECONDS); } catch (err) { /* 太大放不進快取時，每次直接讀表 */ }
  return idx;
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

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 在編輯器裡選這個函式按「執行」，可檢查讀到幾張單（結果只會顯示在你自己的執行紀錄）。 */
function checkSetup() {
  CacheService.getScriptCache().remove('idx');
  const idx = index_();
  const nos = Object.keys(idx);
  Logger.log('讀到 %s 張工令單。PIN 已設定：%s', nos.length, !!PropertiesService.getScriptProperties().getProperty('PIN'));
  if (nos.length) Logger.log('第一張：%s', JSON.stringify(idx[nos[0]]));
}
