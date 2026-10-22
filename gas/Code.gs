/**
 * 平安住・專案管理系統（多案場共用後端）
 * ------------------------------------------------------------
 * 部署帳號：yuengshih@gmail.com
 * 1. 新建 Apps Script 專案，貼上本檔
 * 2. 執行一次 setup()（會建立試算表與雲端資料夾，並把龍德路登錄為外部獨立頁）
 * 3. 專案設定 → 指令碼屬性 → 新增 ADMIN_PASS（平安住管理密碼）→ 儲存指令碼屬性
 * 4. 部署 → 新增部署作業 → 網頁應用程式：執行身分「我」、存取權「所有人」
 * 5. 把 /exec 結尾的網址貼到 GitHub projects repo 的 assets/common.js 的 API
 */

const P = PropertiesService.getScriptProperties();
const SHEETS = {
  cases:   ['id', 'name', 'address', 'status', 'kind', 'url', 'api', 'config', 'state', 'sort', 'createdAt', 'updatedAt'],
  photos:  ['id', 'caseId', 'cat', 'area', 'caption', 'fileId', 'createdAt'],
  updates: ['id', 'caseId', 'date', 'stage', 'title', 'note', 'photoIds', 'createdAt'],
  docs:    ['id', 'caseId', 'type', 'title', 'date', 'files', 'createdAt'],
  log:     ['time', 'action', 'caseId', 'detail']
};
const JSON_COLS = { config: 1, state: 1, photoIds: 1, files: 1 };

/* ================= setup ================= */
function setup() {
  let ssId = P.getProperty('SHEET_ID'), ss;
  if (ssId) ss = SpreadsheetApp.openById(ssId);
  else { ss = SpreadsheetApp.create('平安住・專案管理系統'); P.setProperty('SHEET_ID', ss.getId()); }
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const head = SHEETS[name];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const def = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  if (!P.getProperty('FOLDER_ID')) {
    const f = DriveApp.createFolder('平安住・專案管理系統檔案');
    P.setProperty('FOLDER_ID', f.getId());
  }
  if (!rows_('cases').length) seedLongde_();
  if (!P.getProperty('ADMIN_PASS')) Logger.log('⚠ 請到「專案設定 → 指令碼屬性」新增 ADMIN_PASS');
  Logger.log('試算表：' + ss.getUrl());
  Logger.log('資料夾：' + DriveApp.getFolderById(P.getProperty('FOLDER_ID')).getUrl());
}

function seedLongde_() {
  const cfg = {
    client: '', netTotal: 5522000, taxRate: 0.05,
    areas: ['1F', '2F', '3F', '4F', '5F', '6F／頂樓', '外牆', '其他'],
    stages: [
      { k: 's1', n: '第一階段', d: '打除清運、窗戶訂製、臨時水電', checks: ['拆除完成查驗'] },
      { k: 's2', n: '第二階段', d: '水電、電箱、管線、鐵件', checks: ['水電隱蔽前查驗', '鐵件完成查驗'] },
      { k: 's3', n: '第三階段', d: '土水、防水、窗戶安裝、廁所完工', checks: ['防水層隱蔽前查驗', '廁所及窗戶完成查驗'] },
      { k: 's4', n: '第四階段', d: '批土、外牆防水、總驗收', checks: ['外牆防水查驗', '總驗收'] }
    ],
    pays: [
      { p: '第一期', c: '簽約時', pct: 20, need: ['__SIGN__'] },
      { p: '第二期', c: '拆除完成查驗合格', pct: 20, need: ['拆除完成查驗'] },
      { p: '第三期', c: '水電隱蔽前及鐵件完成查驗合格', pct: 20, need: ['水電隱蔽前查驗', '鐵件完成查驗'] },
      { p: '第四期', c: '廁所及窗戶完成查驗合格', pct: 20, need: ['廁所及窗戶完成查驗'] },
      { p: '第五期・尾款', c: '總驗收完成，結算清運及追加減', pct: 20, need: ['總驗收'] }
    ]
  };
  append_('cases', {
    id: 'longde', name: '龍德路住宅整修工程', address: '高雄市鼓山區龍德路106巷16-1號・1F～6F',
    status: '施工中', kind: 'external', url: 'https://yuengshih-work.github.io/longde/',
    api: 'https://script.google.com/macros/s/AKfycbxqDwrg2CxSYpC7CpoVp32XBEsFH0O0qk6fjqJc7t0faKmeeOG3yTm3SBN-DANiFT7d/exec',
    config: cfg, state: {}, sort: 1, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString()
  });
}

/* ================= HTTP ================= */
function doGet(e) {
  const q = e.parameter || {};
  try {
    const admin = isAdmin_(q.pass);
    switch (q.action) {
      case 'ping': return out_({ ok: true, time: new Date().toISOString() });
      case 'case': return out_(getCase_(q.id, admin));
      case 'hub':
        if (!admin) throw new Error('需要管理密碼');
        return out_(getHub_());
      default: throw new Error('未知的 action');
    }
  } catch (err) { return out_({ ok: false, error: err.message }); }
}

function doPost(e) {
  let b = {};
  try {
    b = JSON.parse(e.postData.contents || '{}');
    if (!isAdmin_(b.pass)) throw new Error('管理密碼錯誤');
    if (b.action === 'checkAdmin') return out_({ ok: true });
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const r = handle_(b);
      log_(b.action, b.caseId || b.id || '', '');
      return out_(Object.assign({ ok: true }, r || {}));
    } finally { lock.releaseLock(); }
  } catch (err) { return out_({ ok: false, error: err.message }); }
}

function handle_(b) {
  switch (b.action) {
    case 'saveCase': return saveCase_(b.case || {});
    case 'deleteCase': return deleteCase_(b.id);
    case 'saveState': {
      const c = findCase_(b.caseId);
      const st = Object.assign({}, c.state || {}, b.state || {});
      updateRow_('cases', 'id', b.caseId, { state: st, updatedAt: new Date().toISOString() });
      return { state: st };
    }
    case 'uploadPhoto': {
      findCase_(b.caseId);
      const f = saveFile_(b.caseId, b.data, (b.cat || 'photo') + '_' + Date.now() + '.jpg');
      const row = { id: uid_(), caseId: b.caseId, cat: b.cat || 'update', area: b.area || '', caption: b.caption || '', fileId: f.getId(), createdAt: new Date().toISOString() };
      append_('photos', row); return { id: row.id, fileId: row.fileId };
    }
    case 'deletePhoto': {
      const p = rows_('photos').find(r => r.id === b.id); if (!p) throw new Error('找不到照片');
      trash_(p.fileId); deleteRow_('photos', 'id', b.id); return {};
    }
    case 'addUpdate': {
      findCase_(b.caseId);
      const row = { id: uid_(), caseId: b.caseId, date: b.date || '', stage: b.stage || '', title: b.title || '', note: b.note || '', photoIds: b.photoIds || [], createdAt: new Date().toISOString() };
      append_('updates', row); return { id: row.id };
    }
    case 'deleteUpdate': {
      const u = rows_('updates').find(r => r.id === b.id); if (!u) throw new Error('找不到紀錄');
      const ps = rows_('photos'); (u.photoIds || []).forEach(pid => { const p = ps.find(x => x.id === pid); if (p) { trash_(p.fileId); deleteRow_('photos', 'id', pid); } });
      deleteRow_('updates', 'id', b.id); return {};
    }
    case 'addDoc': {
      findCase_(b.caseId);
      const row = { id: uid_(), caseId: b.caseId, type: b.type || '其他文件', title: b.title || '', date: b.date || '', files: [], createdAt: new Date().toISOString() };
      append_('docs', row); return { id: row.id };
    }
    case 'uploadDocFile': {
      const d = rows_('docs').find(r => r.id === b.docId); if (!d) throw new Error('找不到文件');
      const f = saveFile_(d.caseId, b.data, b.name || ('doc_' + Date.now()));
      const files = (d.files || []).concat([{ id: f.getId(), name: b.name || '', mime: f.getMimeType() }]);
      updateRow_('docs', 'id', b.docId, { files: files }); return { id: f.getId() };
    }
    case 'deleteDoc': {
      const d = rows_('docs').find(r => r.id === b.id); if (!d) throw new Error('找不到文件');
      (d.files || []).forEach(f => trash_(f.id)); deleteRow_('docs', 'id', b.id); return {};
    }
    default: throw new Error('未知的 action：' + b.action);
  }
}

/* ================= 讀取 ================= */
function getCase_(id, admin) {
  const c = rows_('cases').find(r => r.id === String(id || '').trim());
  if (!c) throw new Error('找不到這個案場（代號：' + (id || '未提供') + '）');
  if (c.kind === 'external') return { ok: true, redirect: c.url };
  return {
    ok: true, admin: admin, case: pub_(c),
    photos: rows_('photos').filter(r => r.caseId === c.id),
    updates: rows_('updates').filter(r => r.caseId === c.id),
    docs: rows_('docs').filter(r => r.caseId === c.id)
  };
}

function getHub_() {
  const cs = rows_('cases'), ups = rows_('updates'), docs = rows_('docs'), phs = rows_('photos');
  return {
    ok: true, admin: true,
    cases: cs.map(c => {
      const o = pub_(c); o.url = c.url; o.api = c.api; o.sort = c.sort;
      if (c.kind !== 'external') {
        const u = ups.filter(r => r.caseId === c.id).sort((a, b) => String(b.date).localeCompare(String(a.date)) || String(b.createdAt).localeCompare(String(a.createdAt)));
        o.latest = u[0] ? { date: u[0].date, title: u[0].title } : null;
        o.updateCount = u.length;
        o.photoCount = phs.filter(r => r.caseId === c.id).length;
        o.docs = docs.filter(r => r.caseId === c.id).map(d => ({ type: d.type, title: d.title, date: d.date }));
      }
      return o;
    })
  };
}

function pub_(c) {
  return { id: c.id, name: c.name, address: c.address, status: c.status, kind: c.kind || 'template', config: c.config || {}, state: c.state || {}, updatedAt: c.updatedAt };
}

/* ================= 案場設定 ================= */
function saveCase_(c) {
  const id = String(c.id || '').trim().toLowerCase();
  if (!/^[a-z0-9][a-z0-9-]{1,30}$/.test(id)) throw new Error('案場代號只能用英文小寫、數字與 -，2～31 字');
  if (!String(c.name || '').trim()) throw new Error('請填案場名稱');
  const now = new Date().toISOString();
  const data = {
    name: String(c.name).trim(), address: String(c.address || '').trim(), status: c.status || '準備中',
    kind: c.kind === 'external' ? 'external' : 'template', url: String(c.url || '').trim(), api: String(c.api || '').trim(),
    config: c.config || {}, sort: Number(c.sort) || 0, updatedAt: now
  };
  const exist = rows_('cases').find(r => r.id === id);
  if (c.isNew && exist) throw new Error('代號「' + id + '」已被使用');
  if (exist) updateRow_('cases', 'id', id, data);
  else append_('cases', Object.assign({ id: id, state: {}, createdAt: now }, data));
  return { id: id };
}

function deleteCase_(id) {
  const c = findCase_(id);
  if (rows_('photos').some(r => r.caseId === id) || rows_('updates').some(r => r.caseId === id) || rows_('docs').some(r => r.caseId === id))
    throw new Error('此案場已有照片、紀錄或文件，不能刪除；請改把狀態設為「已結案」');
  deleteRow_('cases', 'id', c.id); return {};
}

function findCase_(id) {
  const c = rows_('cases').find(r => r.id === id);
  if (!c) throw new Error('找不到案場：' + id);
  return c;
}

/* ================= 檔案 ================= */
function caseFolder_(caseId) {
  const root = DriveApp.getFolderById(P.getProperty('FOLDER_ID'));
  const it = root.getFoldersByName(caseId);
  return it.hasNext() ? it.next() : root.createFolder(caseId);
}
function saveFile_(caseId, dataUrl, name) {
  const m = String(dataUrl || '').match(/^data:([^;]+);base64,(.*)$/);
  if (!m) throw new Error('檔案格式錯誤');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
  const f = caseFolder_(caseId).createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f;
}
function trash_(fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { } }

/* ================= 試算表工具 ================= */
function ss_() { return SpreadsheetApp.openById(P.getProperty('SHEET_ID')); }
function sh_(name) { const s = ss_().getSheetByName(name); if (!s) throw new Error('缺少分頁 ' + name + '，請重新執行 setup()'); return s; }
function rows_(name) {
  const sh = sh_(name), v = sh.getDataRange().getValues(), head = v.shift() || [];
  return v.filter(r => r[0] !== '').map(r => {
    const o = {};
    head.forEach((h, i) => {
      let x = r[i];
      if (x instanceof Date) x = Utilities.formatDate(x, 'Asia/Taipei', h === 'date' ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss");
      if (JSON_COLS[h]) { const dv = (h === 'config' || h === 'state') ? {} : []; try { x = x ? JSON.parse(x) : dv; } catch (e) { x = dv; } if (!x) x = dv; }
      o[h] = x;
    });
    return o;
  });
}
function cell_(h, v) { return JSON_COLS[h] ? JSON.stringify(v == null ? '' : v) : (v == null ? '' : v); }
function append_(name, o) {
  const sh = sh_(name);
  sh.appendRow(SHEETS[name].map(h => cell_(h, o[h])));
  if (name !== 'log') { const r = sh.getLastRow(); const di = SHEETS[name].indexOf('date'); if (di >= 0) sh.getRange(r, di + 1).setNumberFormat('@').setValue(o.date || ''); }
}
function findRow_(name, key, val) {
  const sh = sh_(name), col = SHEETS[name].indexOf(key) + 1;
  const vals = sh.getRange(2, col, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(val)) return i + 2;
  return 0;
}
function updateRow_(name, key, val, patch) {
  const sh = sh_(name), r = findRow_(name, key, val); if (!r) throw new Error('找不到資料 ' + val);
  Object.keys(patch).forEach(h => { const c = SHEETS[name].indexOf(h); if (c >= 0) sh.getRange(r, c + 1).setValue(cell_(h, patch[h])); });
}
function deleteRow_(name, key, val) { const r = findRow_(name, key, val); if (r) sh_(name).deleteRow(r); }
function log_(action, caseId, detail) { try { sh_('log').appendRow([new Date(), action, caseId, detail]); } catch (e) { } }
function uid_() { return Utilities.getUuid().slice(0, 8) + Date.now().toString(36).slice(-4); }
function isAdmin_(p) { const a = String(P.getProperty('ADMIN_PASS') || '').trim(); return !!a && String(p || '').trim() === a; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
