/* 下包廠商專區 共用工具（廠商頁與管理頁共用） */
const TASK_STATUS = ['未開始', '施工中', '待驗收', '已完成'];
const PILL = { '未開始': '', '施工中': 'go', '待驗收': 'part', '已完成': 'done' };
const statusPill = s => `<span class="pill ${PILL[s] || ''}">${esc(s || '未開始')}</span>`;
const payPill = p => p.paid ? '<span class="pill done">已付款</span>' : p.req ? '<span class="pill part">已請款・待付款</span>' : '<span class="pill">未請款</span>';
const md = s => { if (!s) return ''; const [, m, d] = String(s).slice(0, 10).split('-'); return `${+m}/${+d}`; };
const twDate = iso => { try { return new Date(iso).toLocaleDateString('sv-SE', { timeZone: 'Asia/Taipei' }); } catch (e) { return String(iso || '').slice(0, 10); } };
const addDays = (s, n) => { const d = new Date(s + 'T00:00:00'); d.setDate(d.getDate() + n); return d.toLocaleDateString('sv-SE'); };

function range(t, short) {
  const f = short ? md : rocS;
  if (!t.start && !t.end) return '日期未定';
  if (t.start && t.end && t.start !== t.end) return `${f(t.start)}～${f(t.end)}`;
  return f(t.start || t.end) + (t.start ? '' : ' 前');
}

// 進場狀態：看前置工項是否完成
function readiness(t, all) {
  if (t.status === '已完成') return null;
  if (t.status === '待驗收') return { cls: 'wait', text: '已回報完工，等待平安住驗收' };
  if (t.status === '施工中') return null;
  const deps = (t.deps || []).map(id => all.find(x => x.id === id)).filter(Boolean);
  if (!deps.length) return null;
  const left = deps.filter(d => d.status !== '已完成');
  if (!left.length) return { cls: 'ok', text: '前置工序都完成了，可以進場' };
  return { cls: 'wait', text: '要等前置工序完成才能進場：' + left.map(d => `${d.trade || d.vendor || ''}・${d.title}（${d.status}）`).join('、') };
}

function upcoming(ts, days) {
  const t0 = today(), t1 = addDays(t0, days);
  return ts.filter(t => t.status !== '已完成' && (t.status === '施工中' || t.status === '待驗收' ||
    (t.start && t.start <= t1 && (!t.end || t.end >= t0)) || (!t.start && t.end && t.end >= t0 && t.end <= t1)));
}

function byMonth(ts) {
  const g = {};
  ts.forEach(t => { const k = t.start ? t.start.slice(0, 7) : 'zzzz'; (g[k] = g[k] || []).push(t); });
  return Object.keys(g).sort().map(k => ({ label: k === 'zzzz' ? '日期未定' : `${+k.slice(0, 4) - 1911}年${+k.slice(5)}月`, items: g[k] }));
}

// 文字中的電話自動變成可撥打連結
function linkify(s) {
  return esc(s).replace(/(\+?886[-\s]?\d[\d\-\s]{7,12}\d|0\d{1,3}[-\s]?\d{3,4}[-\s]?\d{3,4}(?:\s*(?:#|分機)\s*\d{1,5})?)/g, m => {
    const n = m.replace(/(#|分機).*$/, '').replace(/[^\d+]/g, ''); return `<a href="tel:${n}">${m}</a>`;
  });
}

/* ---------- 3D 圖 ---------- */
function modelUrl(u) { try { return new URL(u, location.href).href; } catch (e) { return ''; } }
// 3D 檔放在本 repo（yuengshih-work.github.io/projects/...）時，自動讀 GitHub 上的檔案最後修改日期
async function repoFileDate(href) {
  try {
    const u = new URL(href); if (u.hostname !== 'yuengshih-work.github.io' || !u.pathname.startsWith('/projects/')) return '';
    let path = decodeURIComponent(u.pathname.slice('/projects/'.length)); if (!path || path.endsWith('/')) path += 'index.html';
    const r = await fetch(`https://api.github.com/repos/yuengshih-work/projects/commits?path=${encodeURIComponent(path)}&per_page=1`);
    if (!r.ok) return ''; const j = await r.json();
    return j[0] ? twDate(j[0].commit.committer.date) : '';
  } catch (e) { return ''; }
}
function renderModel(box, vp) {
  const href = modelUrl(vp.model || '');
  if (!href) { box.innerHTML = ''; return; }
  const log = [...(vp.log || [])].filter(x => x && x.date).sort((a, b) => String(b.date).localeCompare(String(a.date)));
  const last = log[0];
  box.innerHTML = `<div class="card model">
    <div class="mh"><b>3D 示意圖</b><span class="mdate" id="mdate">${last ? '最新更改：' + esc(roc(last.date)) : ''}</span></div>
    ${last && last.note ? `<div class="small mnote">${esc(last.note)}</div>` : ''}
    <div class="frame"><iframe src="${esc(href)}" loading="lazy" title="3D 示意圖" allowfullscreen></iframe></div>
    <a class="btn ghost block sm" href="${esc(href)}" target="_blank" rel="noopener">全螢幕開啟 3D（可旋轉、切換樓層）</a>
    ${log.length ? `<details class="mlog"><summary>更改紀錄（${log.length}）</summary>${log.map(x => `<div class="ml"><span class="muted">${esc(rocS(x.date))}</span>　${esc(x.note || '')}</div>`).join('')}</details>` : ''}
  </div>`;
  repoFileDate(href).then(d => {
    const el = box.querySelector('#mdate'); if (!el || !d) return;
    if (!last || d > String(last.date).slice(0, 10)) el.textContent = '最新更改：' + roc(d);
  });
}
