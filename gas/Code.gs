/**
 * 平安住・專案管理系統（多案場共用後端）
 * ------------------------------------------------------------
 * 部署帳號：yuengshih@gmail.com
 * 1. 新建 Apps Script 專案，貼上本檔
 * 2. 執行一次 setup()（會建立試算表與雲端資料夾，並把龍德路登錄為外部獨立頁）
 * 3. 專案設定 → 指令碼屬性 → 新增 ADMIN_PASS（平安住管理密碼）→ 儲存指令碼屬性
 * 4. 部署 → 新增部署作業 → 網頁應用程式：執行身分「我」、存取權「所有人」
 * 5. 把 /exec 結尾的網址貼到 GitHub projects repo 的 assets/common.js 的 API
 *
 * 更新程式後（例如新增下包廠商專區）：貼上新版 → 再執行一次 setup()（補建新分頁，
 * 不會動到既有資料）→ 管理部署作業 → 編輯 → 版本選「新版本」→ 部署
 */

const P = PropertiesService.getScriptProperties();
const SHEETS = {
  cases:   ['id', 'name', 'address', 'status', 'kind', 'url', 'api', 'config', 'state', 'sort', 'createdAt', 'updatedAt'],
  photos:  ['id', 'caseId', 'cat', 'area', 'caption', 'fileId', 'createdAt'],
  updates: ['id', 'caseId', 'date', 'stage', 'title', 'note', 'photoIds', 'createdAt'],
  docs:    ['id', 'caseId', 'type', 'title', 'date', 'files', 'createdAt'],
  // 下包廠商專區
  vendors: ['id', 'caseId', 'name', 'trade', 'contact', 'phone', 'token', 'showMoney', 'amount', 'pays', 'note', 'active', 'sort', 'createdAt', 'updatedAt'],
  tasks:   ['id', 'caseId', 'vendorId', 'title', 'area', 'start', 'end', 'status', 'deps', 'note', 'sort', 'updatedAt'],
  notices: ['id', 'caseId', 'vendorIds', 'title', 'body', 'important', 'sort', 'createdAt', 'updatedAt'],
  vphotos: ['id', 'caseId', 'vendorId', 'taskId', 'caption', 'fileId', 'createdAt'],
  log:     ['time', 'action', 'caseId', 'detail']
};
const JSON_COLS = { config: 1, state: 1, photoIds: 1, files: 1, pays: 1, deps: 1, vendorIds: 1 };
const DATE_COLS = { date: 1, start: 1, end: 1 };   // 存成純文字 yyyy-MM-dd
const TASK_STATUS = ['未開始', '施工中', '待驗收', '已完成'];

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
      case 'vendor': return out_(vendorView_(q.t));
      case 'vadmin':
        if (!admin) throw new Error('需要管理密碼');
        return out_(vendorAdmin_(q.id));
      default: throw new Error('未知的 action');
    }
  } catch (err) { return out_({ ok: false, error: err.message }); }
}

function doPost(e) {
  let b = {};
  try {
    b = JSON.parse(e.postData.contents || '{}');
    if (/^v_/.test(b.action)) {               // 下包廠商：用專屬連結代碼，不需管理密碼
      const lock = LockService.getScriptLock(); lock.waitLock(20000);
      try { const r = vendorHandle_(b); return out_(Object.assign({ ok: true }, r || {})); }
      finally { lock.releaseLock(); }
    }
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
    case 'editUpdate': {
      const u = rows_('updates').find(r => r.id === b.id); if (!u) throw new Error('找不到紀錄');
      const rm = b.removePhotoIds || [], ps = rows_('photos');
      rm.forEach(pid => { const p = ps.find(x => x.id === pid); if (p) { trash_(p.fileId); deleteRow_('photos', 'id', pid); } });
      const ids = (u.photoIds || []).filter(x => rm.indexOf(x) < 0).concat(b.addPhotoIds || []);
      updateRow_('updates', 'id', b.id, { date: b.date || u.date, stage: b.stage == null ? u.stage : b.stage, title: b.title == null ? u.title : b.title, note: b.note == null ? u.note : b.note, photoIds: ids });
      const di = SHEETS.updates.indexOf('date'); sh_('updates').getRange(findRow_('updates', 'id', b.id), di + 1).setNumberFormat('@').setValue(b.date || u.date);
      return {};
    }
    case 'editPhoto': {
      if (!rows_('photos').some(r => r.id === b.id)) throw new Error('找不到照片');
      const patch = {}; if (b.area != null) patch.area = b.area; if (b.caption != null) patch.caption = b.caption;
      updateRow_('photos', 'id', b.id, patch); return {};
    }
    case 'editDoc': {
      const d = rows_('docs').find(r => r.id === b.id); if (!d) throw new Error('找不到文件');
      const rm = b.removeFileIds || []; rm.forEach(trash_);
      updateRow_('docs', 'id', b.id, { type: b.type || d.type, title: b.title == null ? d.title : b.title, date: b.date == null ? d.date : b.date, files: (d.files || []).filter(f => rm.indexOf(f.id) < 0) });
      const di = SHEETS.docs.indexOf('date'); sh_('docs').getRange(findRow_('docs', 'id', b.id), di + 1).setNumberFormat('@').setValue(b.date == null ? d.date : b.date);
      return {};
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
    default: {
      const r = vendorAdminHandle_(b); if (r) return r;
      throw new Error('未知的 action：' + b.action);
    }
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
      if (x instanceof Date) x = Utilities.formatDate(x, 'Asia/Taipei', DATE_COLS[h] ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss");
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
  if (name !== 'log') { const r = sh.getLastRow(); SHEETS[name].forEach((h, i) => { if (DATE_COLS[h]) sh.getRange(r, i + 1).setNumberFormat('@').setValue(o[h] || ''); }); }
}
function findRow_(name, key, val) {
  const sh = sh_(name), col = SHEETS[name].indexOf(key) + 1;
  const vals = sh.getRange(2, col, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(val)) return i + 2;
  return 0;
}
function updateRow_(name, key, val, patch) {
  const sh = sh_(name), r = findRow_(name, key, val); if (!r) throw new Error('找不到資料 ' + val);
  Object.keys(patch).forEach(h => { const c = SHEETS[name].indexOf(h); if (c < 0) return; const rg = sh.getRange(r, c + 1); if (DATE_COLS[h]) rg.setNumberFormat('@'); rg.setValue(cell_(h, patch[h])); });
}
function deleteRow_(name, key, val) { const r = findRow_(name, key, val); if (r) sh_(name).deleteRow(r); }
function log_(action, caseId, detail) { try { sh_('log').appendRow([new Date(), action, caseId, detail]); } catch (e) { } }
function uid_() { return Utilities.getUuid().slice(0, 8) + Date.now().toString(36).slice(-4); }
function isAdmin_(p) { const a = String(P.getProperty('ADMIN_PASS') || '').trim(); return !!a && String(p || '').trim() === a; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* =====================================================================
 * 下包廠商專區
 * - 每家廠商一個專屬連結代碼（token），不需密碼即可查看該案場的
 *   整體進度、3D 圖、給他的注意事項、自己的工項，並可上傳照片
 * - 金額與請款進度只在該廠商設定「顯示金額」時回傳，且只回傳他自己的
 * - 業主合約金額、成本、其他廠商金額一律不回傳
 * ===================================================================== */

function newToken_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
const bool_ = v => v === true || v === 'TRUE' || v === 'true' || v === 1;
const vp_ = c => Object.assign({ model: '', log: [], contacts: '' }, (c.state || {}).vp || {});

function findVendorByToken_(t) {
  t = String(t || '').trim();
  if (t.length < 20) throw new Error('連結不完整，請向平安住索取完整連結');
  const v = rows_('vendors').find(r => r.token === t);
  if (!v || !bool_(v.active)) throw new Error('這個連結已失效，請向平安住索取新連結');
  return v;
}

function sortTasks_(ts) {
  return ts.sort((a, b) => String(a.start || '9999').localeCompare(String(b.start || '9999')) || (+a.sort || 0) - (+b.sort || 0) || String(a.title).localeCompare(String(b.title)));
}

function vendorView_(t) {
  const v = findVendorByToken_(t);
  const c = findCase_(v.caseId);
  const vs = rows_('vendors').filter(r => r.caseId === c.id);
  const who = {}; vs.forEach(r => who[r.id] = { name: r.name, trade: r.trade });
  const tasks = sortTasks_(rows_('tasks').filter(r => r.caseId === c.id)).map(r => ({
    id: r.id, vendorId: r.vendorId, vendor: (who[r.vendorId] || {}).name || '', trade: (who[r.vendorId] || {}).trade || '',
    title: r.title, area: r.area, start: r.start, end: r.end, status: r.status || '未開始', deps: r.deps || [],
    note: r.vendorId === v.id ? r.note : '', mine: r.vendorId === v.id
  }));
  const notices = rows_('notices').filter(r => r.caseId === c.id && (!(r.vendorIds || []).length || r.vendorIds.indexOf(v.id) >= 0))
    .map(r => ({ id: r.id, title: r.title, body: r.body, important: bool_(r.important), updatedAt: r.updatedAt || r.createdAt, sort: r.sort }));
  const photos = rows_('vphotos').filter(r => r.vendorId === v.id);
  const vp = vp_(c);
  const me = { id: v.id, name: v.name, trade: v.trade, showMoney: bool_(v.showMoney) };
  if (me.showMoney) { me.amount = +v.amount || 0; me.pays = v.pays || []; }
  return { ok: true, case: { id: c.id, name: c.name, address: c.address }, me: me, tasks: tasks, notices: notices, photos: photos, vp: vp };
}

function vendorAdmin_(id) {
  const c = findCase_(String(id || '').trim());
  return {
    ok: true, admin: true, case: pub_(c), vp: vp_(c),
    vendors: rows_('vendors').filter(r => r.caseId === c.id).map(r => Object.assign(r, { showMoney: bool_(r.showMoney), active: bool_(r.active) })),
    tasks: sortTasks_(rows_('tasks').filter(r => r.caseId === c.id)),
    notices: rows_('notices').filter(r => r.caseId === c.id).map(r => Object.assign(r, { important: bool_(r.important) })),
    vphotos: rows_('vphotos').filter(r => r.caseId === c.id)
  };
}

/* ---------- 廠商上傳（token） ---------- */
function vendorHandle_(b) {
  const v = findVendorByToken_(b.t);
  switch (b.action) {
    case 'v_upload': {
      let task = null;
      if (b.taskId) {
        task = rows_('tasks').find(r => r.id === b.taskId && r.caseId === v.caseId);
        if (!task) throw new Error('找不到這個工項');
      }
      const f = saveFile_(v.caseId, b.data, 'vendor_' + v.trade + '_' + Date.now() + '.jpg');
      const row = { id: uid_(), caseId: v.caseId, vendorId: v.id, taskId: task ? task.id : '', caption: String(b.caption || '').slice(0, 200), fileId: f.getId(), createdAt: new Date().toISOString() };
      append_('vphotos', row);
      let status = '';
      if (b.done && task && task.vendorId === v.id && task.status !== '已完成') {
        updateRow_('tasks', 'id', task.id, { status: '待驗收', updatedAt: new Date().toISOString() }); status = '待驗收';
      }
      log_('v_upload', v.caseId, v.name + (task ? '・' + task.title : '') + (status ? '・回報完工' : ''));
      return { id: row.id, fileId: row.fileId, status: status };
    }
    case 'v_delete': {
      const p = rows_('vphotos').find(r => r.id === b.id && r.vendorId === v.id);
      if (!p) throw new Error('只能刪除自己上傳的照片');
      trash_(p.fileId); deleteRow_('vphotos', 'id', p.id); log_('v_delete', v.caseId, v.name); return {};
    }
    default: throw new Error('未知的 action：' + b.action);
  }
}

/* ---------- 管理（需管理密碼） ---------- */
function vendorAdminHandle_(b) {
  const now = new Date().toISOString();
  switch (b.action) {
    case 'saveVP': {
      const c = findCase_(b.caseId);
      const st = Object.assign({}, c.state || {});
      st.vp = Object.assign(vp_(c), b.vp || {});
      updateRow_('cases', 'id', c.id, { state: st, updatedAt: now });
      return { vp: st.vp };
    }
    case 'saveVendor': {
      const x = b.vendor || {}; findCase_(b.caseId);
      if (!String(x.name || '').trim()) throw new Error('請填廠商名稱');
      const data = {
        name: String(x.name).trim(), trade: String(x.trade || '').trim(), contact: String(x.contact || '').trim(), phone: String(x.phone || '').trim(),
        showMoney: !!x.showMoney, amount: +x.amount || 0, pays: Array.isArray(x.pays) ? x.pays : [], note: String(x.note || ''),
        active: x.active !== false, sort: +x.sort || 0, updatedAt: now
      };
      if (x.id) { const old = rows_('vendors').find(r => r.id === x.id && r.caseId === b.caseId); if (!old) throw new Error('找不到廠商'); updateRow_('vendors', 'id', x.id, data); return { id: x.id }; }
      const row = Object.assign({ id: uid_(), caseId: b.caseId, token: newToken_(), createdAt: now }, data);
      append_('vendors', row); return { id: row.id, token: row.token };
    }
    case 'newToken': {
      if (!rows_('vendors').some(r => r.id === b.id)) throw new Error('找不到廠商');
      const t = newToken_(); updateRow_('vendors', 'id', b.id, { token: t, updatedAt: now }); return { token: t };
    }
    case 'deleteVendor': {
      if (rows_('tasks').some(r => r.vendorId === b.id)) throw new Error('這家廠商還有工項，請先把工項改給別家或刪除');
      if (rows_('vphotos').some(r => r.vendorId === b.id)) throw new Error('這家廠商已上傳照片，不能刪除；請改成「停用連結」');
      deleteRow_('vendors', 'id', b.id); return {};
    }
    case 'saveTask': {
      const x = b.task || {}; findCase_(b.caseId);
      if (!String(x.title || '').trim()) throw new Error('請填工項名稱');
      const data = {
        vendorId: x.vendorId || '', title: String(x.title).trim(), area: String(x.area || '').trim(), start: x.start || '', end: x.end || '',
        status: TASK_STATUS.indexOf(x.status) >= 0 ? x.status : '未開始', deps: (x.deps || []).filter(d => d && d !== x.id),
        note: String(x.note || ''), sort: +x.sort || 0, updatedAt: now
      };
      if (x.id) { updateRow_('tasks', 'id', x.id, data); return { id: x.id }; }
      const row = Object.assign({ id: uid_(), caseId: b.caseId }, data); append_('tasks', row); return { id: row.id };
    }
    case 'setTaskStatus': {
      if (TASK_STATUS.indexOf(b.status) < 0) throw new Error('狀態錯誤');
      updateRow_('tasks', 'id', b.id, { status: b.status, updatedAt: now }); return {};
    }
    case 'deleteTask': {
      rows_('tasks').filter(r => (r.deps || []).indexOf(b.id) >= 0).forEach(r => updateRow_('tasks', 'id', r.id, { deps: r.deps.filter(d => d !== b.id) }));
      rows_('vphotos').filter(r => r.taskId === b.id).forEach(r => updateRow_('vphotos', 'id', r.id, { taskId: '' }));
      deleteRow_('tasks', 'id', b.id); return {};
    }
    case 'saveNotice': {
      const x = b.notice || {}; findCase_(b.caseId);
      if (!String(x.title || '').trim()) throw new Error('請填標題');
      const data = { vendorIds: x.vendorIds || [], title: String(x.title).trim(), body: String(x.body || ''), important: !!x.important, sort: +x.sort || 0, updatedAt: now };
      if (x.id) { updateRow_('notices', 'id', x.id, data); return { id: x.id }; }
      const row = Object.assign({ id: uid_(), caseId: b.caseId, createdAt: now }, data); append_('notices', row); return { id: row.id };
    }
    case 'deleteNotice': deleteRow_('notices', 'id', b.id); return {};
    case 'deleteVPhoto': {
      const p = rows_('vphotos').find(r => r.id === b.id); if (!p) throw new Error('找不到照片');
      trash_(p.fileId); deleteRow_('vphotos', 'id', b.id); return {};
    }
    default: return null;
  }
}
