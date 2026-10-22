/* 平安住・專案管理系統 共用設定與工具（總網頁與各案場頁共用） */

/* ======== 設定：貼上「專案管理系統」Apps Script 部署網址（/exec 結尾） ======== */
const API = 'https://script.google.com/macros/s/AKfycbyeiF2_oGk3Ai7vl57RqsBAHKoFXyUlTV7L1tNoNWDRyJSqLU6MPjP8KcCqat1kMjI0Ww/exec';
/* ========================================================================== */

const FIRM = '平安住股份有限公司・統編 90052437・07-9623551 分機555';
const BOOK_URL = 'https://yuengshih-work.github.io/head/';   // 記帳系統
const SIGN = '__SIGN__';
const STATUS_LIST = ['準備中', '施工中', '已完工', '已結案'];

const $ = (s, e = document) => e.querySelector(s);
const $$ = (s, e = document) => [...e.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const money = n => 'NT$' + Math.round(+n || 0).toLocaleString('en-US');
const today = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
const roc = s => { if (!s) return ''; const [y, m, d] = String(s).slice(0, 10).split('-'); return `${+y - 1911}年${+m}月${+d}日`; };
const rocS = s => { if (!s) return ''; const [y, m, d] = String(s).slice(0, 10).split('-'); return `${+y - 1911}/${+m}/${+d}`; };
const img = (id, w) => `https://drive.google.com/thumbnail?id=${encodeURIComponent(id)}&sz=w${w}`;
const configured = () => /^https:\/\/script\.google\.com\/macros\/s\/[\w-]{20,}\/exec$/.test(API);

/* ---------- 登入密碼（同一台裝置在總網頁與案場頁共用） ---------- */
const store = {
  get(k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } },
  set(k, v) { try { v ? localStorage.setItem(k, v) : localStorage.removeItem(k); } catch (e) { } }
};
let PASS = store.get('pj_pass');
function setPass(p) { PASS = p || ''; store.set('pj_pass', PASS); }

/* ---------- 後端 ---------- */
async function fetchJ(url, opt = {}, limit = 30000) {
  const ac = new AbortController(); const t = setTimeout(() => ac.abort(), limit);
  try {
    const r = await fetch(url, { ...opt, signal: ac.signal, cache: 'no-store' });
    const j = await r.json(); if (!j.ok) throw new Error(j.error || '後端回傳錯誤'); return j;
  } catch (e) {
    if (e.name === 'AbortError') throw new Error('後端超過 ' + limit / 1000 + ' 秒沒有回應');
    if (e instanceof TypeError) throw new Error('連線失敗：請確認 Apps Script 已部署為「所有人」可存取，且網址正確');
    throw e;
  } finally { clearTimeout(t); }
}
const apiGet = (params) => fetchJ(API + '?' + new URLSearchParams({ ...params, ...(PASS ? { pass: PASS } : {}) }));
const apiPost = (body, limit = 60000) => fetchJ(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ ...body, pass: PASS }) }, limit);

function busy(on, t) { const s = $('#spin'); if (!s) return; s.classList.toggle('on', on); if (t) $('#spinTxt').textContent = t; }

/* ---------- 金額與進度計算 ---------- */
// 依比例拆每期未稅與稅金；比例合計 100% 時最後一期吸收尾差，確保加總等於合約總價
function payPlan(cfg) {
  const N = Math.round(+cfg.netTotal || 0), rate = cfg.taxRate == null ? 0.05 : +cfg.taxRate, T = Math.round(N * rate);
  const ps = cfg.pays || []; const sumPct = ps.reduce((a, p) => a + (+p.pct || 0), 0);
  let an = 0, at = 0;
  return ps.map((p, i) => {
    let net = Math.round(N * (+p.pct || 0) / 100), tax = Math.round(T * (+p.pct || 0) / 100);
    if (i === ps.length - 1 && Math.abs(sumPct - 100) < 0.001) { net = N - an; tax = T - at; }
    an += net; at += tax;
    return { ...p, net, tax };
  });
}
const contractTotal = cfg => { const N = Math.round(+cfg.netTotal || 0), r = cfg.taxRate == null ? 0.05 : +cfg.taxRate; return { net: N, tax: Math.round(N * r), all: N + Math.round(N * r), rate: r }; };
// 收款紀錄：{net:日期, tax:日期}；舊格式（單一日期字串）視為兩者皆已收
const paidOf = (state, i) => { const v = ((state || {}).paid || {})[i]; if (!v) return {}; if (typeof v === 'string') return { net: v, tax: v }; return v; };
const isMainDoc = d => String(d.type || '').indexOf('主合約') === 0;

function summarize(cfg, state, docs) {
  cfg = cfg || {}; state = state || {};
  const stages = cfg.stages || [], st = state.stages || {}, checks = state.checks || {};
  const stg = k => st[k] || '未開始';
  const done = stages.filter(s => stg(s.k) === '已完成').length;
  const cur = [...stages].reverse().find(s => stg(s.k) === '施工中') || [...stages].reverse().find(s => stg(s.k) === '已完成');
  const curLabel = stages.length && done === stages.length ? '全部完工' : cur ? `${cur.n}${stg(cur.k)}` : '準備中';
  const signed = (docs || []).find(isMainDoc);
  const ready = p => (p.need || []).every(n => n === SIGN ? !!signed : !!checks[n]);
  const plan = payPlan(cfg); const tot = contractTotal(cfg);
  let got = 0; const readyList = [];
  plan.forEach((p, i) => {
    const r = paidOf(state, i); if (r.net) got += p.net; if (r.tax) got += p.tax;
    const left = (r.net ? 0 : p.net) + (r.tax ? 0 : p.tax);
    if (left > 0 && ready(p)) readyList.push({ i, p: p.p, amt: left, part: !!(r.net || r.tax) });
  });
  return { stages, stg, done, curLabel, signed, ready, plan, total: tot.all, tot, got, readyList, checks };
}

/* ---------- 圖檔 ---------- */
async function shrink(file, max = 1600) {
  const bmp = await createImageBitmap(file); const sc = Math.min(1, max / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas'); c.width = Math.round(bmp.width * sc); c.height = Math.round(bmp.height * sc);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height); return c.toDataURL('image/jpeg', 0.8);
}
async function fileData(f) {
  if (f.type === 'application/pdf') {
    if (f.size > 20 * 1024 * 1024) throw new Error('PDF 超過 20MB');
    return await new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(f); });
  }
  return await shrink(f, 2200);
}

/* ---------- 案場網址 ---------- */
function siteBase() { const u = new URL(location.href); const p = u.pathname.replace(/[^/]*$/, ''); return u.origin + p.replace(/site\/$/, ''); }
const caseUrl = c => c.kind === 'external' ? c.url : siteBase() + 'site/?id=' + encodeURIComponent(c.id);
