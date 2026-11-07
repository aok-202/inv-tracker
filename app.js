/* app.js: the Investment Tracker interface. Loads after spx-data.js, calc.js and xlsx.js.
   Data lives in the Excel file you open. Nothing is sent anywhere, and nothing is stored
   unless you turn on "Keep a copy on this device". */
(function () {
'use strict';

/* ---------- small DOM helpers ---------- */
const $ = (sel, root) => (root || document).querySelector(sel);
const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));
function h(tag, attrs, ...kids) {
  const el = document.createElement(tag);
  if (attrs) for (const k in attrs) {
    const v = attrs[k];
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'style') { for (const p in v) el.style.setProperty(p, v[p]); }
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const kid of kids.flat()) {
    if (kid == null || kid === false) continue;
    el.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
  }
  return el;
}
const SVGNS = 'http://www.w3.org/2000/svg';
function sv(tag, attrs, ...kids) {
  const el = document.createElementNS(SVGNS, tag);
  if (attrs) for (const k in attrs) if (attrs[k] != null) el.setAttribute(k, attrs[k]);
  for (const kid of kids.flat()) if (kid != null) el.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
  return el;
}
const ICON = {
  edit: '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M13.6 3.6l2.8 2.8L7.2 15.6H4.4v-2.8z"/><path d="M11.8 5.4l2.8 2.8"/></svg>',
  trash: '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 6h12M8 6V4.5h4V6M6 6l.8 10h6.4L14 6"/></svg>',
  chev: '<svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 3.5l4.5 4.5L6 12.5"/></svg>'
};
function iconBtn(icon, label, attrs) {
  const b = h('button', Object.assign({ type: 'button', class: 'icon-btn', 'aria-label': label, title: label }, attrs || {}));
  b.innerHTML = ICON[icon];
  return b;
}

/* ---------- locale & formatting ---------- */
const LOCALE = (navigator.languages && navigator.languages[0]) || navigator.language || 'en-US';
const LOCALE_DECIMAL_COMMA = (() => { try { return (1.5).toLocaleString(LOCALE).indexOf(',') !== -1; } catch (e) { return false; } })();
const LOCALE_DATE_FMT = /^(en|es)-(US|PH|PR|UM|AS|GU|MP|VI)$/i.test(LOCALE) || /^en$/i.test(LOCALE) ? 'mdy' : 'dmy';
const REDUCED = !!(window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches);
const IS_IOS = /iP(hone|ad|od)/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
const PWA = document.documentElement.dataset.build === 'pwa';
const nfCache = new Map();
function nf(opts) {
  const k = JSON.stringify(opts);
  let f = nfCache.get(k);
  if (!f) {
    try { f = new Intl.NumberFormat(LOCALE, opts); }
    catch (e) { const o = Object.assign({}, opts); delete o.signDisplay; delete o.notation; delete o.currencyDisplay; f = new Intl.NumberFormat(LOCALE, o); }
    nfCache.set(k, f);
  }
  return f;
}
const cur = () => (state && state.settings.cur) || 'USD';
function money(n, o) {
  o = o || {};
  if (n == null || !isFinite(n)) return '—';
  const dec = o.dec != null ? o.dec : (Math.round(n) !== n && Math.abs(n) < 100000 ? 2 : 0);
  const opts = { style: 'currency', currency: cur(), currencyDisplay: 'narrowSymbol', minimumFractionDigits: dec, maximumFractionDigits: dec };
  if (o.sign) opts.signDisplay = 'exceptZero';
  if (o.compact) { opts.notation = 'compact'; opts.minimumFractionDigits = 0; opts.maximumFractionDigits = 1; }
  return nf(opts).format(n);
}
function num(n, dec) { if (n == null || !isFinite(n)) return '—'; return nf({ minimumFractionDigits: dec, maximumFractionDigits: dec }).format(n); }
function plain(n) { return nf({ maximumFractionDigits: 2 }).format(n); }
function pct(r, o) {
  o = o || {};
  if (r == null || !isFinite(r)) return '—';
  const dec = o.dec != null ? o.dec : 1;
  return nf({ style: 'percent', minimumFractionDigits: dec, maximumFractionDigits: dec, signDisplay: o.sign === false ? 'auto' : 'exceptZero' }).format(r);
}
const dfLong = new Intl.DateTimeFormat(LOCALE, { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' });
const dfShort = new Intl.DateTimeFormat(LOCALE, { day: 'numeric', month: 'short', timeZone: 'UTC' });
const dfMonYr = new Intl.DateTimeFormat(LOCALE, { month: 'short', year: '2-digit', timeZone: 'UTC' });
const dfTime = new Intl.DateTimeFormat(LOCALE, { hour: 'numeric', minute: '2-digit' });
const fmtDay = t => (isFinite(t) ? dfLong.format(new Date(t * DAY_MS)) : '—');
const fmtDate = iso => fmtDay(isoToDay(iso));
function curSymbol() {
  try {
    const parts = nf({ style: 'currency', currency: cur(), currencyDisplay: 'narrowSymbol' }).formatToParts(1);
    const c = parts.find(p => p.type === 'currency');
    return c ? c.value : cur();
  } catch (e) { return cur(); }
}
function parseInputAmount(str) {
  const nc = numberCore(str || '');
  if (!nc) return NaN;
  const g = guessDecimalComma(nc.core);
  return parseAmount(str, g == null ? LOCALE_DECIMAL_COMMA : g);
}
const todayIso = () => { const d = new Date(); return ymdToIso(d.getFullYear(), d.getMonth() + 1, d.getDate()); };
const todayDay = () => isoToDay(todayIso());
const plural = (n, one, many) => n + ' ' + (n === 1 ? one : many);
function ago(ms) {
  const s = Math.round((Date.now() - ms) / 1000);
  if (s < 60) return 'just now';
  if (s < 3600) return plural(Math.round(s / 60), 'minute', 'minutes') + ' ago';
  if (s < 86400) return 'at ' + dfTime.format(new Date(ms));
  return 'on ' + dfLong.format(new Date(ms));
}

/* ---------- state ---------- */
function emptyState() { return { v: 1, deposits: [], value: null, history: [], settings: { cur: 'USD', spx: null }, updatedAt: 0 }; }
function newId() { return Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-5); }
function normalize(o) {
  const s = emptyState();
  if (!o || typeof o !== 'object') return s;
  if (Array.isArray(o.deposits)) {
    s.deposits = o.deposits
      .filter(d => d && typeof d.d === 'string' && isFinite(isoToDay(d.d)) && typeof d.a === 'number' && isFinite(d.a) && d.a !== 0)
      .map(d => ({ id: String(d.id || newId()).slice(0, 40), d: d.d, a: round2(d.a) }));
  }
  if (o.value && typeof o.value.a === 'number' && isFinite(o.value.a) && isFinite(isoToDay(o.value.d))) s.value = { a: round2(o.value.a), d: o.value.d };
  if (Array.isArray(o.history)) s.history = o.history.filter(x => x && isFinite(isoToDay(x.d)) && typeof x.a === 'number' && isFinite(x.a)).map(x => ({ d: x.d, a: round2(x.a) }));
  if (o.settings && typeof o.settings === 'object') {
    if (typeof o.settings.cur === 'string' && /^[A-Z]{3}$/.test(o.settings.cur)) s.settings.cur = o.settings.cur;
    const sp = o.settings.spx;
    if (sp && typeof sp.px === 'number' && sp.px > 0 && isFinite(isoToDay(sp.d))) s.settings.spx = { d: sp.d, px: sp.px };
  }
  s.updatedAt = typeof o.updatedAt === 'number' ? o.updatedAt : 0;
  return s;
}
const clone = o => JSON.parse(JSON.stringify(o));
function upsertHistory(s, iso, a) {
  s.history = s.history.filter(x => x.d !== iso);
  s.history.push({ d: iso, a });
  s.history.sort((x, y) => (x.d < y.d ? -1 : 1));
  if (s.history.length > 3000) s.history = s.history.slice(-3000);
}
function makeExample(today) {
  const s = emptyState();
  const deps = [];
  const add = (d, a) => { if (isoToDay(d) <= today) deps.push({ id: 'ex' + deps.length, d, a }); };
  const per = { 2019: 400, 2020: 500, 2021: 550, 2022: 600, 2023: 650, 2024: 750, 2025: 800, 2026: 900 };
  for (let y = 2019; y <= 2026; y++) {
    for (let m = 1; m <= 12; m++) {
      if (y === 2019 && m < 3) continue;
      if ((y === 2022 && (m === 7 || m === 8)) || (y === 2024 && m === 2)) continue;
      add(ymdToIso(y, m, 15), per[y]);
    }
  }
  add('2020-03-23', 3000); add('2021-12-20', 2000); add('2023-06-05', -2500); add('2023-12-18', 2500); add('2025-04-08', 3000);
  s.deposits = deps;
  const A = buildAnchors(SPX, null);
  const spAt = t => { let u = 0; for (const d of deps) { const dt = isoToDay(d.d); if (dt <= t) u += d.a / trAt(A, dt); } return u * trAt(A, t); };
  const factors = [0.99, 0.975, 0.985, 0.955, 0.962, 0.958, 0.961];
  [2019, 2020, 2021, 2022, 2023, 2024, 2025].forEach((y, i) => {
    const t = isoToDay(y + '-12-31');
    if (t <= today) s.history.push({ d: y + '-12-31', a: Math.round(spAt(t) * factors[i]) });
  });
  const vd = dayToIso(today);
  s.value = { a: Math.round(spAt(today) * 0.957), d: vd };
  upsertHistory(s, vd, s.value.a);
  return s;
}

/* mode: 'welcome' (nothing open), 'example' (sample numbers), 'data' (your file or a new tracker) */
let mode = 'welcome';
let state = emptyState();
const file = { name: null, appFormat: false, dirty: false, savedAt: null, openedAt: null, source: null };

/* Optional copy on this device. Off by default; when off, nothing is written to storage. */
const COPY_KEY = 'investment-tracker:copy:v1';
let storageOK = false;
try { const k = '__it_probe'; localStorage.setItem(k, '1'); localStorage.removeItem(k); storageOK = true; } catch (e) { storageOK = false; }
let keepCopy = false;
function writeCopy() {
  if (!keepCopy || !storageOK) return;
  try {
    localStorage.setItem(COPY_KEY, JSON.stringify(mode === 'data' ? { keep: true, state, file } : { keep: true }));
  } catch (e) { /* storage full or blocked: the Excel file still has everything */ }
}
function dropCopy() { if (!storageOK) return; try { localStorage.removeItem(COPY_KEY); } catch (e) { /* ignore */ } }
(() => {
  if (!storageOK) return;
  let c = null;
  try { c = JSON.parse(localStorage.getItem(COPY_KEY) || 'null'); } catch (e) { c = null; }
  if (!c || !c.keep) return;
  keepCopy = true;
  if (c.state) { state = normalize(c.state); Object.assign(file, c.file || {}); mode = 'data'; }
})();

let anchors = buildAnchors(SPX, state.settings.spx);
let anchorsKey = JSON.stringify(state.settings.spx);
function getAnchors() {
  const k = JSON.stringify(state.settings.spx);
  if (k !== anchorsKey) { anchors = buildAnchors(SPX, state.settings.spx); anchorsKey = k; }
  return anchors;
}

/* Inside the Claude app, files are saved through its downloads capability. Elsewhere it doesn't exist. */
let downloadsCap = null;
if (window.claude && typeof window.claude.use === 'function') {
  window.claude.use('downloads').then(d => { downloadsCap = d || null; }).catch(() => { downloadsCap = null; });
}

/* ---------- committing changes ---------- */
let undoSnapshot;
function snapshot() { return { state: clone(state), mode, file: Object.assign({}, file) }; }
function commitData(mutator) {
  const prev = snapshot();
  if (mode !== 'data') {
    const settings = state.settings;
    state = emptyState();
    state.settings = settings;
    if (mode === 'example') Object.assign(file, { name: null, appFormat: false, savedAt: null, source: null });
    mode = 'data';
    openYears.clear();
    resetZoom();
  }
  mutator(state);
  state.updatedAt = Date.now();
  file.dirty = true;
  writeCopy();
  render();
  return prev;
}
function commitSettings(mutator) {
  mutator(state);
  if (mode === 'data') { state.updatedAt = Date.now(); file.dirty = true; writeCopy(); }
  render();
}
function restore(snap) {
  state = normalize(snap.state);
  mode = snap.mode;
  Object.assign(file, snap.file);
  editing = null;
  writeCopy();
  render();
}

let toastTimer = null;
function toast(msg, prev) {
  $('#toast-msg').textContent = msg;
  undoSnapshot = prev;
  $('#toast-undo').hidden = prev === undefined;
  const t = $('#toast');
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), prev === undefined ? 3200 : 7000);
}

/* ---------- tabs ---------- */
const TABS = ['overview', 'deposits'];
let tab = 'overview';
const scrollPos = { overview: 0, deposits: 0 };
function setTab(t, opts) {
  opts = opts || {};
  if (!TABS.includes(t)) t = 'overview';
  if (t !== tab) scrollPos[tab] = window.scrollY;
  tab = t;
  for (const x of TABS) {
    $('#view-' + x).hidden = x !== t;
    const b = $('#tabbtn-' + x);
    b.setAttribute('aria-selected', String(x === t));
    b.tabIndex = x === t ? 0 : -1;
  }
  render();
  if (!opts.keepScroll) window.scrollTo(0, scrollPos[t] || 0);
  if (t === 'overview') requestAnimationFrame(animateGuilloche);
}

/* ---------- rendering ---------- */
let editing = null;
const openYears = new Set();
let histOpen = false;
let spxOpen = false;
const proj = { years: 10, annual: null, annualTouched: false, rate: 0.07, preset: '7' };
let lastM = null;

function render() {
  const M = computeMetrics(state, getAnchors(), todayDay());
  lastM = M;
  const sym = curSymbol();
  $$('[data-cur]').forEach(el => {
    el.textContent = sym;
    el.parentElement.style.setProperty('--cur-w', Math.max(1, sym.length) + 'ch');
  });
  document.body.dataset.mode = mode;
  renderHeader();
  renderBanner();
  renderHero(M);
  renderValueCard(M);
  renderKpis(M);
  renderBenchmark(M);
  if (tab === 'overview' && mode !== 'welcome') { renderGrowth(M); renderYears(M); renderProjection(M); }
  renderLedger(M);
  renderSettings();
}
function renderCharts() { if (lastM && tab === 'overview' && mode !== 'welcome') { renderGrowth(lastM); renderYears(lastM); renderProjection(lastM); } }

function renderHeader() {
  const status = $('#file-status');
  const save = $('#btn-save');
  save.disabled = mode === 'welcome';
  save.classList.toggle('btn-primary', mode === 'data' && file.dirty);
  if (mode === 'welcome') status.textContent = 'No file open';
  else if (mode === 'example') status.textContent = 'Example numbers · not saved';
  else {
    const name = file.appFormat && file.name ? file.name : 'New tracker';
    const st = file.dirty ? 'unsaved changes' : file.savedAt ? 'saved ' + ago(file.savedAt) : file.openedAt ? 'opened ' + ago(file.openedAt) : 'not saved yet';
    status.textContent = name + ' · ' + st;
  }
  status.dataset.state = mode === 'data' && file.dirty ? 'dirty' : 'ok';
}
function renderBanner() {
  const b = $('#banner');
  b.replaceChildren();
  if (mode === 'example') {
    b.hidden = false;
    b.append(h('p', null, h('strong', null, 'These are example numbers. '), 'Open your Excel file, or add a deposit to start your own tracker.'),
      h('div', { class: 'banner-actions' },
        h('button', { class: 'btn btn-primary btn-sm', type: 'button', 'data-action': 'open-file' }, 'Open Excel file'),
        h('button', { class: 'btn btn-sm', type: 'button', 'data-action': 'close-file' }, 'Close example')));
  } else if (mode === 'data' && file.dirty) {
    b.hidden = false;
    b.append(h('p', null, h('strong', null, 'Changes aren’t in your Excel file yet. '), keepCopy ? 'They’re kept on this device for now; save to put them in your file.' : 'Save before you close the app, or they’ll be lost.'),
      h('div', { class: 'banner-actions' }, h('button', { class: 'btn btn-primary btn-sm', type: 'button', 'data-action': 'save' }, 'Save to Excel')));
  } else b.hidden = true;
}

function renderHero(M) {
  const welcome = mode === 'welcome';
  $('#welcome').hidden = !welcome;
  $('#dashboard').hidden = welcome;
  if (welcome) return;
  $('#hero-value').textContent = M.hasValue ? money(M.V, { dec: 0 }) : '—';
  $('#hero-asof').textContent = M.hasValue ? 'As of ' + fmtDate(state.value.d) : 'Enter your current value below to see your returns.';
  const chips = $('#hero-chips');
  chips.replaceChildren();
  if (M.hasValue && M.gain != null) {
    const up = M.gain >= 0;
    chips.append(h('span', { class: 'pill ' + (up ? 'pill-good' : 'pill-bad') },
      h('span', { 'aria-hidden': 'true' }, up ? '▲' : '▼'),
      money(M.gain, { sign: true, dec: 0 }) + (M.roi != null ? ' (' + pct(M.roi) + ')' : '')));
    chips.append(h('span', { class: 'pill' }, 'on ' + money(M.netCounted, { dec: 0 }) + ' deposited'));
  }
}

function renderValueCard(M) {
  if (mode === 'welcome') return;
  const inp = $('#val-amount'), dt = $('#val-date');
  if (document.activeElement !== inp) inp.value = M.hasValue ? plain(M.V) : '';
  if (document.activeElement !== dt) dt.value = state.value ? state.value.d : todayIso();
  const hint = $('#val-hint');
  if (M.hasValue) {
    const age = M.today - M.vDay;
    hint.textContent = age <= 0 ? 'Updated today.' : age === 1 ? 'Updated yesterday.' : 'Updated ' + plural(age, 'day', 'days') + ' ago.' + (age > 31 ? ' Worth refreshing.' : '');
  } else hint.textContent = 'What is your account worth right now?';
  const warn = $('#val-warn');
  if (M.hasValue && M.after.length) {
    warn.hidden = false;
    warn.textContent = plural(M.after.length, 'deposit', 'deposits') + ' after ' + fmtDate(state.value.d) + (M.after.length === 1 ? ' isn’t' : ' aren’t') + ' counted in your returns yet. Update the value to include ' + (M.after.length === 1 ? 'it.' : 'them.');
  } else warn.hidden = true;
  const list = $('#val-history'), tog = $('#val-history-toggle');
  const hist = state.history.slice().sort((a, b) => (a.d < b.d ? 1 : -1));
  tog.hidden = hist.length === 0;
  tog.textContent = (histOpen ? 'Hide' : 'Show') + ' value history (' + hist.length + ')';
  tog.setAttribute('aria-expanded', String(histOpen));
  list.hidden = !histOpen || !hist.length;
  list.replaceChildren(...hist.map(x => h('li', null,
    h('span', { class: 'h-date' }, fmtDate(x.d)),
    h('span', { class: 'h-val' }, money(x.a, { dec: 0 })),
    mode === 'example' ? h('span') : iconBtn('trash', 'Delete the value saved for ' + fmtDate(x.d), { 'data-action': 'del-snap', 'data-d': x.d }))));
}

function tone(v) { return v == null || !isFinite(v) || v === 0 ? '' : v > 0 ? 'good' : 'bad'; }
function renderKpis(M) {
  if (mode === 'welcome') return;
  const perYr = v => (v == null ? '—' : [pct(v), h('small', null, ' /yr')]);
  const tiles = [
    { label: 'Total deposited', value: money(M.total, { dec: 0 }),
      cap: M.count ? (M.nWd ? plural(M.nDep, 'deposit', 'deposits') + ', ' + plural(M.nWd, 'withdrawal', 'withdrawals') : plural(M.nDep, 'deposit', 'deposits')) : 'No deposits yet' },
    { label: 'Average per year', value: M.avgPerYear != null ? money(M.avgPerYear, { dec: 0 }) : '—',
      cap: !M.avgYears ? 'Average of each calendar year’s total' : M.avgYears.n === 1 ? 'Only ' + M.avgYears.y0 + ' so far'
        : 'Average of ' + M.avgYears.n + ' yearly totals, ' + M.avgYears.y0 + '–' + M.avgYears.y1 },
    { label: 'Years since first deposit', value: M.years != null ? num(M.years, 2) : '—',
      cap: M.first != null ? 'Since ' + fmtDay(M.first) : 'Starts with your first deposit' },
    { label: 'Deposited in ' + M.cy, value: money(M.thisYear.sum, { dec: 0 }), cap: (M.cy - 1) + ': ' + money(M.lastYear.sum, { dec: 0 }) },
    { label: 'ROI', value: pct(M.roi), tone: tone(M.roi), cap: M.hasValue ? 'Gain ÷ deposits' : 'Needs your current value' },
    { label: 'CAGR', value: perYr(M.cagr), tone: tone(M.cagr), cap: '(Value\u00a0÷\u00a0deposits)\u200b^(1/years)\u00a0−\u00a01' },
    { label: 'Annualized return (XIRR)', value: perYr(M.xirr), tone: tone(M.xirr), cap: 'Counts how long each deposit was invested' },
    { label: 'vs S&P 500', value: M.sp.diff != null ? money(M.sp.diff, { sign: true, dec: 0 }) : '—', tone: tone(M.sp.diff),
      cap: M.sp.value != null ? 'Same deposits in the index: ' + money(M.sp.value, { dec: 0 }) : (M.sp.unavailable ? 'Index data starts in 1985' : 'Needs deposits and a value') }
  ];
  $('#kpis').replaceChildren(...tiles.map(t => h('div', { class: 'kpi' },
    h('p', { class: 'kpi-label' }, t.label),
    h('p', { class: 'kpi-value' + (t.tone ? ' ' + t.tone : '') }, t.value),
    h('p', { class: 'kpi-cap' }, t.cap))));
}

/* ---------- charts ---------- */
const measureCtx = document.createElement('canvas').getContext('2d');
const BODY_FONT = '-apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif';
function textW(str, size) { measureCtx.font = (size || 12) + 'px ' + BODY_FONT; return measureCtx.measureText(str).width; }
function niceTicks(min, max, n) {
  if (max === min) max = min + 1;
  const raw = (max - min) / Math.max(1, n);
  const p = Math.pow(10, Math.floor(Math.log10(raw)));
  const f = raw / p;
  const step = (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p;
  const lo = Math.floor(min / step + 1e-9) * step, hi = Math.ceil(max / step - 1e-9) * step;
  const ticks = [];
  for (let v = lo; v <= hi + step / 2; v += step) ticks.push(Math.abs(v) < step / 1e6 ? 0 : +v.toPrecision(12));
  return { lo, hi, ticks };
}
function nearest(arr, v) {
  let lo = 0, hi = arr.length - 1;
  if (v <= arr[0]) return 0;
  if (v >= arr[hi]) return hi;
  while (hi - lo > 1) { const mid = (lo + hi) >> 1; if (arr[mid] <= v) lo = mid; else hi = mid; }
  return v - arr[lo] < arr[hi] - v ? lo : hi;
}
function yAxis(svg, T, fmt, ml, mr, W, Y) {
  const compact = Math.max(Math.abs(T.lo), Math.abs(T.hi)) >= 10000;
  const g = sv('g', { class: 'grid' });
  T.ticks.forEach(v => {
    const y = Math.round(Y(v)) + 0.5;
    g.append(sv('line', { x1: ml, x2: W - mr, y1: y, y2: y, class: v === 0 ? 'base' : null }));
    svg.append(sv('text', { x: ml - 8, y: y + 4, 'text-anchor': 'end' }, fmt(v, compact)));
  });
  svg.prepend(g);
}
function axisPad(T, fmt) {
  const compact = Math.max(Math.abs(T.lo), Math.abs(T.hi)) >= 10000;
  return Math.ceil(Math.max(...T.ticks.map(v => textW(fmt(v, compact))))) + 14;
}
function legendItem(cls, label, value, area) {
  return h('li', null, h('i', { class: area ? 'area-key' : 'key key-' + cls }), h('span', null, label), value != null ? h('b', null, value) : null);
}
/* o.yZero false: the y-axis fits the data shown instead of starting at zero (zoomed-in charts).
   o.external: no pointer handling of its own; the caller drives show/hide through the returned api. */
function lineChart(box, o) {
  box.replaceChildren();
  const W = Math.max(260, Math.floor(box.clientWidth || 600)), H = o.height;
  let yMax = -Infinity, yMin = Infinity;
  for (const s of o.series) for (const p of s.pts) { if (p[1] > yMax) yMax = p[1]; if (p[1] < yMin) yMin = p[1]; }
  if (!isFinite(yMax)) { yMax = 0; yMin = 0; }
  let T;
  if (o.yZero === false && yMax > yMin) {
    const pad = (yMax - yMin) * 0.08;
    T = niceTicks(yMin >= 0 ? Math.max(0, yMin - pad) : yMin - pad, yMax + pad, H < 230 ? 4 : 5);
  } else T = niceTicks(Math.min(0, yMin), Math.max(0, yMax) * 1.02, H < 230 ? 4 : 5);
  const ml = axisPad(T, o.yFmt), mr = 10, mt = 12, mb = 28;
  const pw = W - ml - mr, ph = H - mt - mb;
  const X = v => ml + (o.x1 === o.x0 ? pw / 2 : ((v - o.x0) / (o.x1 - o.x0)) * pw);
  const Y = v => mt + ((T.hi - v) / (T.hi - T.lo)) * ph;
  const svg = sv('svg', { width: W, height: H, viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': o.label });
  yAxis(svg, T, o.yFmt, ml, mr, W, Y);
  for (const t of o.xTicks) {
    const x = X(t.v);
    if (x < ml - 1 || x > W - mr + 1) continue;
    const tw = textW(t.label);
    const anchor = x - tw / 2 < 2 ? 'start' : x + tw / 2 > W - 2 ? 'end' : 'middle';
    svg.append(sv('line', { x1: Math.round(x) + 0.5, x2: Math.round(x) + 0.5, y1: mt + ph, y2: mt + ph + 4, class: 'tick' }));
    svg.append(sv('text', { x: anchor === 'start' ? Math.max(0, x - 2) : anchor === 'end' ? Math.min(W, x + 2) : x, y: H - 8, 'text-anchor': anchor }, t.label));
  }
  const baseY = Y(Math.max(T.lo, 0));
  for (const s of o.series) {
    if (!s.pts.length) continue;
    let d = '';
    s.pts.forEach((p, i) => {
      const x = X(p[0]).toFixed(1), y = Y(p[1]).toFixed(1);
      if (!i) d += 'M' + x + ' ' + y;
      else if (s.type === 'step') d += 'H' + x + 'V' + y;
      else d += 'L' + x + ' ' + y;
    });
    if (s.area) svg.append(sv('path', { d: d + 'V' + baseY.toFixed(1) + 'H' + X(s.pts[0][0]).toFixed(1) + 'Z', class: 'area f-' + s.cls }));
    svg.append(sv('path', { d, class: 'line s-' + s.cls }));
    if (s.dots) for (const p of (s.dotPts || s.pts)) svg.append(sv('circle', { cx: X(p[0]).toFixed(1), cy: Y(p[1]).toFixed(1), r: 4, class: 'dot f-' + s.cls }));
    else if (s.end) { const p = s.endPt || s.pts[s.pts.length - 1]; svg.append(sv('circle', { cx: X(p[0]).toFixed(1), cy: Y(p[1]).toFixed(1), r: 4, class: 'dot f-' + s.cls })); }
  }
  const xh = sv('line', { class: 'xhair', y1: mt, y2: mt + ph, x1: -20, x2: -20, visibility: 'hidden' });
  const fg = sv('g');
  const hit = o.external ? null : sv('rect', { x: ml, y: 0, width: Math.max(1, pw), height: H - 6, fill: 'transparent', class: 'hit' });
  svg.append(xh, fg);
  if (hit) svg.append(hit);
  box.append(svg);
  const tip = h('div', { class: 'tip', hidden: true });
  box.append(tip);
  let curI = -1;
  function show(i) {
    if (!o.samples.length) return;
    i = Math.max(0, Math.min(o.samples.length - 1, i));
    curI = i;
    const x = Math.round(X(o.samples[i])) + 0.5;
    xh.setAttribute('x1', x); xh.setAttribute('x2', x); xh.setAttribute('visibility', 'visible');
    fg.replaceChildren(...o.focus(i).map(f => sv('circle', { cx: X(f.x).toFixed(1), cy: Y(f.y).toFixed(1), r: 5, class: 'dot f-' + f.cls })));
    const r = o.readout(i);
    tip.replaceChildren(h('div', { class: 'tip-title' }, r.title),
      ...r.rows.map(row => h('div', { class: 'tip-row' }, h('i', { class: 'key key-' + row.cls }), h('b', null, row.value), h('span', null, row.label))));
    tip.hidden = false;
    const tw = tip.offsetWidth;
    let left = x + 14;
    if (left + tw > W) left = x - 14 - tw;
    if (left < 0) left = Math.max(0, Math.min(W - tw, x - tw / 2));
    tip.style.left = left + 'px';
  }
  function hide() { curI = -1; xh.setAttribute('visibility', 'hidden'); fg.replaceChildren(); tip.hidden = true; }
  const idxAt = cx => {
    const rect = svg.getBoundingClientRect();
    const px = (cx - rect.left) * (W / rect.width);
    return nearest(o.samples, o.x0 + ((px - ml) / pw) * (o.x1 - o.x0));
  };
  if (hit) {
    hit.addEventListener('pointermove', e => show(idxAt(e.clientX)));
    hit.addEventListener('pointerdown', e => show(idxAt(e.clientX)));
    hit.addEventListener('pointerleave', e => { if (e.pointerType === 'mouse') hide(); });
  }
  box.tabIndex = 0;
  box.setAttribute('aria-label', o.label + '. Use the left and right arrow keys to read values.');
  box.onkeydown = e => {
    if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
      e.preventDefault();
      show(curI < 0 ? o.samples.length - 1 : curI + (e.key === 'ArrowRight' ? 1 : -1));
    } else if (e.key === 'Escape') hide();
  };
  box.onblur = hide;
  return {
    show: i => { if (o.samples.length) show(i); }, hide, idxAt,
    /* the plot area in page pixels, for turning finger positions into dates */
    plot: () => { const r = svg.getBoundingClientRect(), k = r.width / W; return { left: r.left + ml * k, width: Math.max(1, pw * k) }; }
  };
}
function roundedBar(x, w, y0, y1, r) {
  const hgt = Math.abs(y0 - y1);
  r = Math.min(r, w / 2, hgt);
  const f = v => v.toFixed(1);
  if (y1 <= y0) return 'M' + f(x) + ' ' + f(y0) + 'V' + f(y1 + r) + 'Q' + f(x) + ' ' + f(y1) + ' ' + f(x + r) + ' ' + f(y1) + 'H' + f(x + w - r) + 'Q' + f(x + w) + ' ' + f(y1) + ' ' + f(x + w) + ' ' + f(y1 + r) + 'V' + f(y0) + 'Z';
  return 'M' + f(x) + ' ' + f(y0) + 'V' + f(y1 - r) + 'Q' + f(x) + ' ' + f(y1) + ' ' + f(x + r) + ' ' + f(y1) + 'H' + f(x + w - r) + 'Q' + f(x + w) + ' ' + f(y1) + ' ' + f(x + w) + ' ' + f(y1 - r) + 'V' + f(y0) + 'Z';
}
function barChart(box, o) {
  box.replaceChildren();
  const W = Math.max(260, Math.floor(box.clientWidth || 600)), H = o.height;
  const vals = o.bars.map(b => b.value);
  const yMax = Math.max(0, ...vals, o.ref ? o.ref.value : 0), yMin = Math.min(0, ...vals);
  const T = niceTicks(yMin, yMax * 1.04, 4);
  const ml = axisPad(T, o.yFmt), mr = 8, mt = 22, mb = 28;
  const pw = W - ml - mr, ph = H - mt - mb;
  const n = o.bars.length, band = pw / Math.max(1, n), bw = Math.max(4, Math.min(24, band * 0.62));
  const Y = v => mt + ((T.hi - v) / (T.hi - T.lo)) * ph;
  const svg = sv('svg', { width: W, height: H, viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': o.label });
  yAxis(svg, T, o.yFmt, ml, mr, W, Y);
  const y0 = Y(0);
  const useShort = o.bars.some(b => textW(b.label) + 8 > band);
  const lw = Math.max(...o.bars.map(b => textW(useShort ? b.short : b.label))) + 8;
  const every = Math.max(1, Math.ceil(lw / band));
  const maxV = Math.max(...vals);
  const barEls = [];
  o.bars.forEach((b, i) => {
    const cx = ml + band * (i + 0.5), yv = Y(b.value);
    const p = Math.abs(yv - y0) >= 0.5 ? sv('path', { d: roundedBar(cx - bw / 2, bw, y0, yv, 4), class: 'bar f-' + o.cls + (b.faded ? ' faded' : '') }) : sv('g');
    svg.append(p); barEls.push(p);
    if ((n - 1 - i) % every === 0) svg.append(sv('text', { x: cx, y: H - 8, 'text-anchor': 'middle' }, useShort ? b.short : b.label));
    const vl = o.valueLabel(b.value);
    const showVal = (n <= 8 || b.value === maxV || i === n - 1) && textW(vl) + 4 <= Math.max(band, 40) && b.value !== 0;
    if (showVal) svg.append(sv('text', { x: cx, y: b.value >= 0 ? yv - 6 : yv + 15, 'text-anchor': 'middle', class: 'val' }, vl));
  });
  if (o.ref && o.ref.value > 0) {
    const y = Math.round(Y(o.ref.value)) + 0.5;
    svg.append(sv('line', { x1: ml, x2: W - mr, y1: y, y2: y, class: 'ref' }));
  }
  const tip = h('div', { class: 'tip', hidden: true });
  let curI = -1;
  function show(i) {
    i = Math.max(0, Math.min(n - 1, i)); curI = i;
    barEls.forEach((el, j) => el.classList && el.classList.toggle('dim', j !== i));
    const b = o.bars[i];
    tip.replaceChildren(h('div', { class: 'tip-title' }, b.tip.title), ...b.tip.rows.map(row => h('div', { class: 'tip-row' }, h('i', { class: 'key key-' + row.cls }), h('b', null, row.value), h('span', null, row.label))));
    tip.hidden = false;
    const cx = ml + band * (i + 0.5), tw = tip.offsetWidth;
    let left = cx + bw / 2 + 10;
    if (left + tw > W) left = cx - bw / 2 - 10 - tw;
    if (left < 0) left = Math.max(0, Math.min(W - tw, cx - tw / 2));
    tip.style.left = left + 'px';
  }
  function hide() { curI = -1; barEls.forEach(el => el.classList && el.classList.remove('dim')); tip.hidden = true; }
  o.bars.forEach((b, i) => {
    const r = sv('rect', { x: ml + band * i, y: 0, width: band, height: H, fill: 'transparent', class: 'hit' });
    r.addEventListener('pointerenter', () => show(i));
    r.addEventListener('pointerdown', () => show(i));
    r.addEventListener('pointerleave', e => { if (e.pointerType === 'mouse') hide(); });
    svg.append(r);
  });
  box.append(svg, tip);
  box.tabIndex = 0;
  box.setAttribute('aria-label', o.label + '. Use the left and right arrow keys to read values.');
  box.onkeydown = e => {
    if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') { e.preventDefault(); show(curI < 0 ? n - 1 : curI + (e.key === 'ArrowRight' ? 1 : -1)); }
    else if (e.key === 'Escape') hide();
  };
  box.onblur = hide;
}
function timeTicks(t0, t1, width) {
  const maxN = Math.max(2, Math.floor((width - 50) / 50));
  const ticks = [];
  if (t1 - t0 > 730) {
    const y0 = dayYear(t0) + (dayMonth(t0) === 1 && dayDate(t0) === 1 ? 0 : 1), y1 = dayYear(t1);
    const count = y1 - y0 + 1;
    const step = [1, 2, 5, 10, 20, 50].find(s => Math.ceil(count / s) <= maxN) || 50;
    for (let y = Math.ceil(y0 / step) * step; y <= y1; y += step) ticks.push({ v: Math.round(Date.UTC(y, 0, 1) / DAY_MS), label: String(y) });
  } else {
    let y = dayYear(t0), m = dayMonth(t0) + 1;
    if (m > 12) { m = 1; y++; }
    const months = [];
    for (;;) { const t = Math.round(Date.UTC(y, m - 1, 1) / DAY_MS); if (t > t1) break; months.push({ t, m }); m++; if (m > 12) { m = 1; y++; } }
    const step = [1, 2, 3, 6, 12].find(s => Math.ceil(months.length / s) <= maxN) || 12;
    months.forEach(mm => { if ((mm.m - 1) % step === 0) ticks.push({ v: mm.t, label: dfMonYr.format(new Date(mm.t * DAY_MS)) }); });
  }
  if (ticks.length < 2) return [{ v: t0, label: dfShort.format(new Date(t0 * DAY_MS)) }, { v: t1, label: dfShort.format(new Date(t1 * DAY_MS)) }];
  return ticks;
}
const yMoney = (v, compact) => money(v, compact ? { compact: true } : { dec: 0 });
const isWide = () => window.innerWidth >= 720;

/* ---------- the growth chart, with a window you can zoom: pinch, trackpad, or the slider below ---------- */
const MIN_SPAN = 31;
const zoom = { preset: 'all', a: null, b: null, pinEnd: false };
const gz = { G: null, M: null, api: null };
function resetZoom() { Object.assign(zoom, { preset: 'all', a: null, b: null, pinEnd: false }); }
const canZoom = G => !!G && G.end - G.start >= 2 * MIN_SPAN;
const isZoomed = (G, w) => w.a > G.start || w.b < G.end;
function clampWindow(G, a, b) {
  const full = G.end - G.start, minS = Math.min(full, MIN_SPAN);
  a = Math.round(a); b = Math.round(b);
  const span = Math.max(minS, Math.min(full, b - a));
  a = Math.max(G.start, Math.min(G.end - span, a));
  return { a, b: a + span };
}
const PRESETS = [
  { k: 'ytd', label: 'YTD', name: 'This year so far' },
  { k: '1y', label: '1Y', name: 'Last year', n: 1 },
  { k: '3y', label: '3Y', name: 'Last 3 years', n: 3 },
  { k: '5y', label: '5Y', name: 'Last 5 years', n: 5 },
  { k: 'all', label: 'All', name: 'Everything' }
];
function presetRange(G, k) {
  if (k === 'all') return { a: G.start, b: G.end };
  const p = PRESETS.find(x => x.k === k);
  if (!p) return null;
  const e = new Date(G.end * DAY_MS), y = e.getUTCFullYear();
  const a = Math.round((k === 'ytd' ? Date.UTC(y, 0, 1) : Date.UTC(y - p.n, e.getUTCMonth(), e.getUTCDate())) / DAY_MS);
  if (a <= G.start + 14 || G.end - a < MIN_SPAN) return null;
  return { a, b: G.end };
}
function currentWindow(G) {
  if (!canZoom(G)) return { a: G.start, b: G.end };
  if (zoom.preset) return presetRange(G, zoom.preset) || { a: G.start, b: G.end };
  if (zoom.a == null) return { a: G.start, b: G.end };
  return clampWindow(G, zoom.a, zoom.pinEnd ? G.end : zoom.b);
}
function setZoom(a, b) {
  const G = gz.G;
  if (!G) return;
  const w = clampWindow(G, a, b);
  Object.assign(zoom, { a: w.a, b: w.b, pinEnd: w.b >= G.end, preset: isZoomed(G, w) ? null : 'all' });
}
let zoomQueued = false;
function queueZoom(a, b) {
  setZoom(a, b);
  if (zoomQueued) return;
  zoomQueued = true;
  requestAnimationFrame(() => { zoomQueued = false; drawGrowth(); });
}
function spanText(days) {
  if (days >= 347) { const y = days / 365.25; return Math.abs(y - Math.round(y)) < 0.05 ? plural(Math.round(y), 'year', 'years') : num(y, 1) + ' years'; }
  const m = Math.round(days / 30.44);
  return m >= 2 ? m + ' months' : plural(Math.round(days), 'day', 'days');
}

function renderGrowth(M) {
  const G = growthSeries(M, state.history, getAnchors());
  gz.G = G; gz.M = M;
  if (!G) {
    gz.api = null;
    $('#growth-plot').replaceChildren(h('p', { class: 'empty' }, 'Your chart appears after your first deposit.'));
    $('#growth-legend').replaceChildren(); $('#growth-table').replaceChildren(); $('#growth-tv').hidden = true;
    $('#growth-nav').hidden = true; $('#growth-presets').replaceChildren(); $('#growth-window').textContent = '';
    $('#growth-compare').hidden = true;
    return;
  }
  $('#growth-tv').hidden = false;
  drawSpark(G);
  drawGrowth();
  growthTable(G);
}

function drawGrowth() {
  const G = gz.G, M = gz.M;
  if (!G || mode === 'welcome') return;
  const zoomable = canZoom(G), w = currentWindow(G), zoomed = isZoomed(G, w);
  const V = windowSeries(G, M.deps, getAnchors(), w.a, w.b);
  const spLabel = V.rebased ? 'S&P 500 from ' + fmtDay(V.cmp.t0) : 'S&P 500, same deposits';
  const series = [{ cls: 'dep', type: 'step', area: true, pts: V.pts.map(p => [p.t, p.dep]) }];
  const spPts = V.pts.filter(p => p.sp != null).map(p => [p.t, p.sp]);
  if (spPts.length) series.push({ cls: 'sp', type: 'line', end: true, pts: spPts });
  const lastReal = V.vin.length ? V.vin[V.vin.length - 1] : null;
  if (V.vline.length) series.push({ cls: 'value', type: 'line', pts: V.vline, dots: V.vin.length <= 36, dotPts: V.vin.map(v => [v.t, v.a]), end: !!lastReal, endPt: lastReal ? [lastReal.t, lastReal.a] : null });
  const valBy = new Map(V.vin.map(v => [v.t, v.a]));
  const lastP = V.pts[V.pts.length - 1];
  $('#growth-legend').replaceChildren(
    V.vline.length ? legendItem('value', 'Your portfolio', lastReal ? money(lastReal.a, { dec: 0 }) : null) : null,
    spPts.length ? legendItem('sp', spLabel, money(spPts[spPts.length - 1][1], { dec: 0 })) : null,
    legendItem('dep', 'Deposited', money(lastP.dep, { dec: 0 }), true));
  const plot = $('#growth-plot');
  gz.api = lineChart(plot, {
    height: isWide() ? 300 : 250, x0: w.a, x1: w.b, external: true, yZero: !zoomed,
    xTicks: timeTicks(w.a, w.b, plot.clientWidth || 600),
    series, yFmt: yMoney, samples: V.pts.map(p => p.t),
    readout: i => {
      const p = V.pts[i], rows = [];
      if (valBy.has(p.t)) rows.push({ cls: 'value', value: money(valBy.get(p.t), { dec: 0 }), label: 'Your portfolio' });
      if (p.sp != null) rows.push({ cls: 'sp', value: money(p.sp, { dec: 0 }), label: spLabel });
      rows.push({ cls: 'dep', value: money(p.dep, { dec: 0 }), label: 'Deposited' });
      return { title: fmtDay(p.t), rows };
    },
    focus: i => {
      const p = V.pts[i], f = [{ cls: 'dep', x: p.t, y: p.dep }];
      if (p.sp != null) f.push({ cls: 'sp', x: p.t, y: p.sp });
      if (valBy.has(p.t)) f.push({ cls: 'value', x: p.t, y: valBy.get(p.t) });
      return f;
    },
    label: 'Line chart' + (zoomed ? ' from ' + fmtDay(w.a) + ' to ' + fmtDay(w.b) : '') + ': money deposited, your portfolio value, and ' + (V.rebased ? 'the S&P 500 starting from your value on ' + fmtDay(V.cmp.t0) : 'the same deposits in the S&P 500') + ', over time'
  });
  $('#growth-chart').classList.toggle('zoomed', zoomed);
  $('#growth-nav').hidden = !zoomable;
  if (zoomable) navLayout(G, w);
  renderPresets(G, zoomable);
  $('#growth-window').textContent = !zoomable ? '' : zoomed ? fmtDay(w.a) + ' – ' + fmtDay(w.b) + ' · ' + spanText(w.b - w.a) : 'Pinch the chart or drag the handles to look at a shorter period.';
  renderCompare(V, zoomed);
}

function renderPresets(G, zoomable) {
  const box = $('#growth-presets');
  if (!zoomable) { box.replaceChildren(); return; }
  box.replaceChildren(...PRESETS.filter(p => presetRange(G, p.k)).map(p =>
    h('button', { type: 'button', 'data-action': 'zoom-preset', 'data-k': p.k, 'aria-label': p.name, 'aria-pressed': String(zoom.preset === p.k) }, p.label)));
}

function renderCompare(V, zoomed) {
  const box = $('#growth-compare');
  if (!zoomed) { box.hidden = true; box.replaceChildren(); return; }
  box.hidden = false;
  const c = V.cmp;
  const addLink = () => (mode === 'example' ? null : h('button', { type: 'button', class: 'link', 'data-action': 'open-import', 'data-target': 'values' }, 'Add many values'));
  if (c.problem) {
    const msg = c.problem === 'nosp' ? 'The S&P 500 data here starts in January 1985, so this period can’t be compared with it.'
      : c.fromStart ? 'To compare this period with the S&P 500, the app needs your portfolio value on a date inside it.'
        : 'To compare this period with the S&P 500, the app needs your portfolio value on at least two dates inside it.';
    box.replaceChildren(h('p', { class: 'note' }, msg + ' ', c.problem === 'nosp' ? null : addLink()));
    return;
  }
  const row = (cls, name, gain, ret) => h('tr', null,
    h('td', null, h('i', { class: 'key key-' + cls }), name),
    h('td', null, money(gain, { sign: true, dec: 0 })),
    h('td', null, ret != null ? pct(ret) : '—'));
  const ahead = c.diff >= 0;
  const flows = Math.abs(c.net) < 0.005 ? 'with no deposits in between' : c.net > 0 ? 'then get the same ' + money(c.net, { dec: 0 }) + ' of deposits' : 'then have the same ' + money(-c.net, { dec: 0 }) + ' taken out';
  box.replaceChildren(
    h('p', { class: 'cmp-title' }, (c.fromStart ? 'Your first deposit, ' + fmtDay(c.t0) : fmtDay(c.t0)) + ' to ' + fmtDay(c.t1)),
    h('div', { class: 'table-scroll' }, h('table', { class: 'num-table cmp-table' },
      h('thead', null, h('tr', null, h('th', { scope: 'col' }, h('span', { class: 'sr-only' }, 'Investment')), h('th', { scope: 'col' }, 'Gain'), h('th', { scope: 'col' }, 'Return'))),
      h('tbody', null, row('value', 'Your portfolio', c.gainV, c.retV), row('sp', 'S&P 500', c.gainS, c.retS)))),
    h('p', { class: 'verdict' },
      h('span', { class: 'pill ' + (ahead ? 'pill-good' : 'pill-bad') }, h('span', { 'aria-hidden': 'true' }, ahead ? '▲' : '▼'),
        money(Math.abs(c.diff), { dec: 0 }) + (ahead ? ' ahead of' : ' behind') + ' the S&P 500')),
    h('p', { class: 'fine' }, c.fromStart
      ? 'Gains after deposits. The S&P 500 side gets the same deposits on the same days. Return counts each deposit from the day it went in.'
      : 'Gains after deposits. Both start from your ' + money(c.V0, { dec: 0 }) + ' on ' + fmtDay(c.t0) + ' and ' + flows + '. Return counts each deposit from the day it went in.'));
}

function growthTable(G) {
  const valBy = new Map(G.vals.map(v => [v.t, v.a]));
  const rowsT = new Set();
  for (const p of G.pts) if (dayMonth(p.t) === 12 && dayDate(p.t) === 31) rowsT.add(p.t);
  for (const v of G.vals) rowsT.add(v.t);
  rowsT.add(G.end);
  const byT = new Map(G.pts.map(p => [p.t, p]));
  const trs = [...rowsT].filter(t => byT.has(t)).sort((a, b) => b - a).map(t => {
    const p = byT.get(t);
    return h('tr', null, h('td', null, fmtDay(t)), h('td', null, money(p.dep, { dec: 0 })),
      h('td', null, valBy.has(t) ? money(valBy.get(t), { dec: 0 }) : '—'), h('td', null, p.sp != null ? money(p.sp, { dec: 0 }) : '—'));
  });
  $('#growth-table').replaceChildren(h('table', { class: 'num-table' },
    h('thead', null, h('tr', null, h('th', { scope: 'col' }, 'Date'), h('th', { scope: 'col' }, 'Deposited'), h('th', { scope: 'col' }, 'Your portfolio'), h('th', { scope: 'col' }, 'S&P 500, same deposits'))),
    h('tbody', null, trs)));
}

/* The small overview behind the window slider: everything, start to end. */
function drawSpark(G) {
  const svg = $('#growth-spark');
  const full = Math.max(1, G.end - G.start);
  let max = 0, min = 0;
  for (const p of G.pts) { if (p.dep > max) max = p.dep; if (p.dep < min) min = p.dep; }
  for (const v of G.vals) { if (v.a > max) max = v.a; if (v.a < min) min = v.a; }
  if (max === min) max = min + 1;
  const X = t => (((t - G.start) / full) * 1000).toFixed(1), Y = v => (94 - ((v - min) / (max - min)) * 84).toFixed(1);
  let dep = '';
  G.pts.forEach((p, i) => { dep += i ? 'H' + X(p.t) + 'V' + Y(p.dep) : 'M' + X(p.t) + ' ' + Y(p.dep); });
  const kids = [sv('path', { d: dep + 'V' + Y(0) + 'H' + X(G.pts[0].t) + 'Z', class: 'sp-area' })];
  if (G.vals.length > 1) kids.push(sv('path', { d: G.vals.map((v, i) => (i ? 'L' : 'M') + X(v.t) + ' ' + Y(v.a)).join(''), class: 'sp-line' }));
  svg.replaceChildren(...kids);
}

/* ---- the window slider ---- */
const navTrack = $('#nav-track'), navWin = $('#nav-win'), navHL = $('#nav-h-l'), navHR = $('#nav-h-r');
function navLayout(G, w) {
  const full = G.end - G.start, minS = Math.min(full, MIN_SPAN);
  const l = ((w.a - G.start) / full) * 100, r = ((w.b - G.start) / full) * 100;
  navWin.style.left = l + '%';
  navWin.style.width = (r - l) + '%';
  $('#nav-shade-l').style.width = l + '%';
  $('#nav-shade-r').style.left = r + '%';
  navWin.classList.toggle('narrow', (navTrack.clientWidth * (w.b - w.a)) / full < 48);
  const setSlider = (el, min, max, now) => {
    el.setAttribute('aria-valuemin', String(Math.round(min)));
    el.setAttribute('aria-valuemax', String(Math.round(max)));
    el.setAttribute('aria-valuenow', String(Math.round(now)));
    el.setAttribute('aria-valuetext', fmtDay(now));
  };
  setSlider(navHL, G.start, w.b - minS, w.a);
  setSlider(navHR, w.a + minS, G.end, w.b);
}
let navDrag = null;
navTrack.addEventListener('pointerdown', e => {
  const G = gz.G;
  if (!canZoom(G) || (e.pointerType === 'mouse' && e.button !== 0)) return;
  const rect = navTrack.getBoundingClientRect(), full = G.end - G.start;
  let w = currentWindow(G), kind;
  const handle = e.target.closest('.nav-handle');
  if (handle) kind = navWin.classList.contains('narrow') ? 'either' : handle === navHL ? 'l' : 'r';
  else if (e.target.closest('#nav-win')) kind = 'pan';
  else {
    /* a tap beside the window moves it there */
    const c = G.start + ((e.clientX - rect.left) / rect.width) * full, span = w.b - w.a;
    setZoom(c - span / 2, c + span / 2);
    drawGrowth();
    w = currentWindow(G);
    kind = 'pan';
  }
  navDrag = { id: e.pointerId, kind, x0: e.clientX, a0: w.a, b0: w.b, k: full / Math.max(1, rect.width), minS: Math.min(full, MIN_SPAN) };
  try { navTrack.setPointerCapture(e.pointerId); } catch (err) { /* fine without capture */ }
});
navTrack.addEventListener('pointermove', e => {
  const d = navDrag;
  if (!d || e.pointerId !== d.id) return;
  const dx = e.clientX - d.x0, dd = dx * d.k;
  if (d.kind === 'either') { if (Math.abs(dx) < 3) return; d.kind = dx < 0 ? 'l' : 'r'; }
  const G = gz.G;
  if (d.kind === 'l') queueZoom(Math.max(G.start, Math.min(d.b0 - d.minS, d.a0 + dd)), d.b0);
  else if (d.kind === 'r') queueZoom(d.a0, Math.min(G.end, Math.max(d.a0 + d.minS, d.b0 + dd)));
  else { const span = d.b0 - d.a0, a = Math.max(G.start, Math.min(G.end - span, d.a0 + dd)); queueZoom(a, a + span); }
});
const endNavDrag = e => { if (navDrag && e.pointerId === navDrag.id) navDrag = null; };
navTrack.addEventListener('pointerup', endNavDrag);
navTrack.addEventListener('pointercancel', endNavDrag);
function navKey(e, which) {
  const G = gz.G;
  if (!canZoom(G)) return;
  const w = currentWindow(G), minS = Math.min(G.end - G.start, MIN_SPAN);
  const step = { ArrowLeft: -30, ArrowDown: -30, ArrowRight: 30, ArrowUp: 30, PageDown: -365, PageUp: 365 }[e.key];
  let { a, b } = w;
  if (which === 'l') {
    if (e.key === 'Home') a = G.start; else if (e.key === 'End') a = b - minS; else if (step) a = Math.max(G.start, Math.min(b - minS, a + step)); else return;
  } else {
    if (e.key === 'Home') b = a + minS; else if (e.key === 'End') b = G.end; else if (step) b = Math.min(G.end, Math.max(a + minS, b + step)); else return;
  }
  e.preventDefault();
  setZoom(a, b);
  drawGrowth();
}
navHL.addEventListener('keydown', e => navKey(e, 'l'));
navHR.addEventListener('keydown', e => navKey(e, 'r'));

/* ---- gestures on the chart: one finger reads values, two fingers zoom and slide ----
   The touch layer stays in place while the chart under it is redrawn, so a pinch never loses its fingers. */
const gTouch = $('#growth-touch');
const ptrs = new Map();
let gest = null, holdUntilClear = false, lastTap = null, gestureStart = null, lastPtrType = 'mouse';
function zoomAround(clientX) {
  const G = gz.G;
  if (!canZoom(G) || !gz.api) return;
  const w = currentWindow(G);
  if (isZoomed(G, w)) { resetZoom(); drawGrowth(); return; }
  const P = gz.api.plot(), f = Math.max(0, Math.min(1, (clientX - P.left) / P.width));
  const c = w.a + f * (w.b - w.a), span = Math.max(MIN_SPAN, (w.b - w.a) / 3);
  setZoom(c - f * span, c - f * span + span);
  drawGrowth();
}
function startPinch() {
  const [p, q] = [...ptrs.values()];
  const P = gz.api.plot(), w = currentWindow(gz.G), mx = (p.x + q.x) / 2;
  gest = { kind: 'pinch', P, d0: Math.max(24, Math.hypot(p.x - q.x, p.y - q.y)), a0: w.a, b0: w.b,
    dm: w.a + ((mx - P.left) / P.width) * (w.b - w.a) };
  gz.api.hide();
}
function movePinch() {
  const [p, q] = [...ptrs.values()];
  const d = Math.max(24, Math.hypot(p.x - q.x, p.y - q.y)), mx = (p.x + q.x) / 2;
  const span = ((gest.b0 - gest.a0) * gest.d0) / d;
  const a = gest.dm - (span * (mx - gest.P.left)) / gest.P.width;
  queueZoom(a, a + span);
}
gTouch.addEventListener('pointerdown', e => {
  lastPtrType = e.pointerType;
  if (!gz.api || (e.pointerType === 'mouse' && e.button !== 0)) return;
  ptrs.set(e.pointerId, { x: e.clientX, y: e.clientY, x0: e.clientX, y0: e.clientY, t0: e.timeStamp, type: e.pointerType });
  if (e.pointerType !== 'touch') { try { gTouch.setPointerCapture(e.pointerId); } catch (err) { /* ok */ } }
  if (ptrs.size === 2 && canZoom(gz.G)) { startPinch(); return; }
  if (ptrs.size > 1) return;
  const G = gz.G, w = G ? currentWindow(G) : null;
  gest = e.pointerType === 'mouse' && G && isZoomed(G, w) ? { kind: 'drag', x0: e.clientX, a0: w.a, b0: w.b, moved: false, P: gz.api.plot() } : { kind: 'scrub' };
  gz.api.show(gz.api.idxAt(e.clientX));
});
gTouch.addEventListener('pointermove', e => {
  if (!gz.api) return;
  const p = ptrs.get(e.pointerId);
  if (!p) { if (e.pointerType === 'mouse') gz.api.show(gz.api.idxAt(e.clientX)); return; }
  p.x = e.clientX; p.y = e.clientY;
  if (gest && gest.kind === 'pinch') { if (ptrs.size >= 2) movePinch(); return; }
  if (holdUntilClear || ptrs.size > 1) return;
  if (gest && gest.kind === 'drag') {
    if (Math.abs(e.clientX - gest.x0) > 3) gest.moved = true;
    if (gest.moved) {
      const span = gest.b0 - gest.a0, a = gest.a0 - ((e.clientX - gest.x0) / gest.P.width) * span;
      gz.api.hide();
      queueZoom(a, a + span);
      return;
    }
  }
  gz.api.show(gz.api.idxAt(e.clientX));
});
function endPointer(e) {
  const p = ptrs.get(e.pointerId);
  ptrs.delete(e.pointerId);
  if (gest && gest.kind === 'pinch') {
    if (ptrs.size < 2) { gest = null; holdUntilClear = ptrs.size > 0; }
    return;
  }
  if (!ptrs.size) {
    const wasHold = holdUntilClear;
    holdUntilClear = false;
    if (e.type === 'pointerup' && p && p.type === 'touch' && !wasHold && gest && gest.kind === 'scrub' &&
        Math.hypot(e.clientX - p.x0, e.clientY - p.y0) < 12 && e.timeStamp - p.t0 < 320) {
      /* double tap: zoom in around that spot, or back out to everything */
      if (lastTap && e.timeStamp - lastTap.t < 360 && Math.abs(e.clientX - lastTap.x) < 36) { lastTap = null; zoomAround(e.clientX); }
      else lastTap = { t: e.timeStamp, x: e.clientX };
    }
    gest = null;
  }
}
gTouch.addEventListener('pointerup', endPointer);
gTouch.addEventListener('pointercancel', endPointer);
gTouch.addEventListener('pointerleave', e => { if (e.pointerType === 'mouse' && !ptrs.size && gz.api) gz.api.hide(); });
/* Browsers also report a double tap as a double click; the touch code above already handled it. */
gTouch.addEventListener('dblclick', e => { if (lastPtrType === 'mouse') zoomAround(e.clientX); });
/* Keep the page itself from zooming when two fingers are on the chart (iPhone Safari needs both of these). */
const twoFingers = e => { if (e.touches && e.touches.length > 1 && canZoom(gz.G)) e.preventDefault(); };
gTouch.addEventListener('touchstart', twoFingers, { passive: false });
gTouch.addEventListener('touchmove', twoFingers, { passive: false });
/* Safari's trackpad pinch on a Mac arrives as gesture events; on iPhone the pointer code above handles it. */
gTouch.addEventListener('gesturestart', e => {
  e.preventDefault();
  if (ptrs.size > 1 || !gz.api || !canZoom(gz.G)) { gestureStart = null; return; }
  const w = currentWindow(gz.G), P = gz.api.plot(), x = e.clientX || P.left + P.width / 2;
  gestureStart = { w, P, x, dm: w.a + ((x - P.left) / P.width) * (w.b - w.a) };
});
gTouch.addEventListener('gesturechange', e => {
  e.preventDefault();
  if (!gestureStart || ptrs.size > 1 || (gest && gest.kind === 'pinch') || !(e.scale > 0)) return;
  const g = gestureStart, span = (g.w.b - g.w.a) / e.scale, a = g.dm - (span * (g.x - g.P.left)) / g.P.width;
  queueZoom(a, a + span);
});
gTouch.addEventListener('gestureend', e => { e.preventDefault(); gestureStart = null; });
/* Trackpad pinch in Chrome and Firefox (ctrl + wheel), and sideways swipes to slide a zoomed chart. */
gTouch.addEventListener('wheel', e => {
  const G = gz.G;
  if (!gz.api || !canZoom(G)) return;
  const w = currentWindow(G), span = w.b - w.a, P = gz.api.plot();
  const unit = e.deltaMode === 1 ? 16 : e.deltaMode === 2 ? P.width : 1;
  if (e.ctrlKey) {
    e.preventDefault();
    const f = Math.max(0, Math.min(1, (e.clientX - P.left) / P.width)), dm = w.a + f * span;
    const ns = span * Math.exp(Math.max(-1, Math.min(1, e.deltaY * unit * 0.01)));
    queueZoom(dm - f * ns, dm - f * ns + ns);
  } else if (Math.abs(e.deltaX) > Math.abs(e.deltaY) && isZoomed(G, w)) {
    e.preventDefault();
    const dd = ((e.deltaX * unit) / P.width) * span;
    queueZoom(w.a + dd, w.b + dd);
  }
}, { passive: false });

function renderBenchmark(M) {
  if (mode === 'welcome') return;
  const body = $('#bm-body'), sp = M.sp;
  $('#spx-form').hidden = !spxOpen;
  $('#spx-clear').hidden = !state.settings.spx;
  const parts = [];
  if (!M.count) {
    parts.push(h('p', { class: 'empty' }, 'Add deposits to see what the same money would have done in the S&P 500.'));
  } else if (sp.unavailable === 'early') {
    parts.push(h('p', { class: 'empty' }, 'The S&P 500 data here starts in January 1985, and your first deposit is earlier, so this comparison isn’t available.'));
  } else if (sp.value != null) {
    const when = M.hasValue ? 'on ' + fmtDay(sp.endT) : 'today';
    parts.push(h('p', { class: 'lede' }, 'If each deposit had gone into the S&P 500 on the same day, with dividends reinvested, it would be worth ', h('strong', null, money(sp.value, { dec: 0 })), ' ' + when + '.'));
    const max = Math.max(sp.value, M.hasValue ? M.V : 0, 1);
    const row = (cls, name, amt, rate) => h('div', { class: 'vs-row' },
      h('span', { class: 'vs-name' }, h('i', { class: 'key key-' + cls }), name),
      h('span', { class: 'vs-amt' }, money(amt, { dec: 0 })),
      h('div', { class: 'vs-track', 'aria-hidden': 'true' }, h('i', { class: 't-' + cls, style: { width: Math.max(1, (Math.max(0, amt) / max) * 100).toFixed(1) + '%' } })),
      h('span', { class: 'vs-rate' }, rate));
    const vs = h('div', { class: 'vs' });
    if (M.hasValue) vs.append(row('value', 'Your portfolio', M.V, M.xirr != null ? 'Annualized ' + pct(M.xirr) + ' a year (XIRR)' : 'Annualized return needs at least a month of history'));
    vs.append(row('sp', 'S&P 500, same deposits', sp.value, sp.xirr != null ? 'Annualized ' + pct(sp.xirr) + ' a year (XIRR)' : 'Index up ' + pct(sp.periodReturn) + ' since your first deposit'));
    parts.push(vs);
    if (M.hasValue && sp.diff != null) {
      const ahead = sp.diff >= 0;
      const gap = M.xirr != null && sp.xirr != null ? (M.xirr - sp.xirr) * 100 : null;
      parts.push(h('p', { class: 'verdict' },
        h('span', { class: 'pill ' + (ahead ? 'pill-good' : 'pill-bad') }, h('span', { 'aria-hidden': 'true' }, ahead ? '▲' : '▼'), money(Math.abs(sp.diff), { dec: 0 }) + (ahead ? ' ahead' : ' behind')),
        h('span', null, gap != null ? (Math.abs(gap) < 0.05 ? 'Level with the index.' : num(Math.abs(gap), 1) + ' percentage points a year ' + (gap >= 0 ? 'better' : 'worse') + ' than the index.') : '')));
    }
    if (sp.stale) parts.push(h('p', { class: 'warn-note mt-12' }, 'Your value is dated after the last S&P close stored here (' + fmtDay(sp.lastDay) + '), so the S&P side stops on that day. Add a newer close below for a like-for-like comparison.'));
  } else {
    parts.push(h('p', { class: 'empty' }, 'Add deposits and your current value to compare.'));
  }
  parts.push(h('p', { class: 'fine' },
    'S&P 500 total return (dividends reinvested), month-end data from 1985 to the close of ' + fmtDay(sp.lastDay) + ' at ' + num(sp.lastPx, 2) + (sp.user ? ', which you added' : '') + '. ',
    cur() !== 'USD' ? 'Assumes your deposits were in US dollars; currency moves aren’t included. ' : '',
    spxOpen || mode === 'example' ? null : h('button', { type: 'button', class: 'link', 'data-action': 'edit-spx' }, 'Add a newer close')));
  body.replaceChildren(...parts);
}

function renderYears(M) {
  const box = $('#years-chart'), tb = $('#years-tbody'), lg = $('#years-legend');
  lg.replaceChildren();
  if (M.byYear.length > 1 && M.avgPerYear > 0) lg.append(h('li', null, h('i', { class: 'key dash' }), h('span', null, 'Average per year'), h('b', null, money(M.avgPerYear, { dec: 0 }))));
  if (M.byYear.some(r => r.ytd)) lg.append(h('li', null, h('span', null, 'Lighter bar: ' + M.cy + ' so far')));
  if (!M.byYear.length) {
    box.replaceChildren(h('p', { class: 'empty' }, 'Yearly totals appear after your first deposit.'));
    tb.replaceChildren();
    return;
  }
  const bars = M.byYear.map(r => ({
    label: String(r.year), short: '’' + String(r.year).slice(2), value: r.sum, faded: r.ytd,
    tip: { title: r.ytd ? r.year + ' (year to date)' : String(r.year), rows: [{ cls: 'dep', value: money(r.sum, { dec: 0 }), label: plural(r.n, 'entry', 'entries') }] }
  }));
  barChart(box, {
    height: isWide() ? 240 : 210, bars, cls: 'dep', yFmt: yMoney,
    valueLabel: v => money(v, { compact: Math.abs(v) >= 10000, dec: 0 }),
    ref: M.avgPerYear > 0 && M.byYear.length > 1 ? { value: M.avgPerYear } : null,
    label: 'Bar chart: money deposited each year' + (M.byYear.some(r => r.ytd) ? '; the current year is lighter because it is not finished' : '')
  });
  tb.replaceChildren(...M.byYear.slice().reverse().map(r => h('tr', { class: r.ytd ? 'current' : null },
    h('td', null, r.ytd ? r.year + ' YTD' : String(r.year)), h('td', null, money(r.sum, { dec: 0 })), h('td', null, String(r.n)), h('td', null, money(r.cum, { dec: 0 })))));
}

function renderProjection(M) {
  const v0 = M.hasValue ? M.V : Math.max(0, M.total);
  if (!proj.annualTouched) proj.annual = M.avgPerYear > 0 ? Math.round(M.avgPerYear / 100) * 100 : 6000;
  const ai = $('#p-annual');
  if (document.activeElement !== ai) ai.value = plain(proj.annual);
  $('#p-years').value = proj.years;
  $('#p-years-out').textContent = plural(proj.years, 'year', 'years');
  $('#p-rate').value = (proj.rate * 100).toFixed(1);
  $('#p-rate-out').textContent = pct(proj.rate, { sign: false, dec: (proj.rate * 1000) % 10 ? 1 : 0 });
  const presets = [];
  if (M.xirr != null && M.xirr > -0.5 && M.xirr < 0.5) presets.push({ k: 'xirr', label: 'Your XIRR ' + pct(M.xirr, { sign: false }), r: M.xirr });
  presets.push({ k: '10', label: 'S&P long-run 10%', r: 0.10 }, { k: '7', label: 'Middle 7%', r: 0.07 }, { k: '4', label: 'Cautious 4%', r: 0.04 });
  $('#p-presets').replaceChildren(...presets.map(p => h('button', { type: 'button', 'data-action': 'proj-preset', 'data-rate': String(p.r), 'data-k': p.k, 'aria-pressed': String(proj.preset === p.k) }, p.label)));
  const pts = projectSeries(v0, proj.annual, proj.rate, proj.years);
  const end = pts[pts.length - 1];
  const added = proj.annual * proj.years, growth = end.v - v0 - added;
  $('#proj-result').replaceChildren('In ' + (M.cy + proj.years) + ' you could have ', h('strong', null, money(end.v, { dec: 0 })));
  $('#proj-break').textContent = money(v0, { dec: 0 }) + (M.hasValue ? ' today' : ' deposited so far') + ' + ' + money(added, { dec: 0 }) + ' in new deposits + ' + money(growth, { dec: 0 }) + ' of growth.';
  $('#proj-legend').replaceChildren(legendItem('value', 'Projected value', money(end.v, { dec: 0 })), legendItem('dep', 'Without growth', money(end.base, { dec: 0 }), true));
  const box = $('#proj-chart');
  const maxTicks = Math.max(2, Math.floor(((box.clientWidth || 600) - 60) / 52));
  const step = [1, 2, 5, 10].find(st => Math.floor(proj.years / st) + 1 <= maxTicks) || 10;
  const ticks = [];
  for (let y = 0; y <= proj.years; y += step) ticks.push({ v: y, label: String(M.cy + y) });
  lineChart(box, {
    height: isWide() ? 220 : 190, x0: 0, x1: proj.years, xTicks: ticks,
    series: [{ cls: 'dep', type: 'line', area: true, pts: pts.map(p => [p.y, p.base]) }, { cls: 'value', type: 'line', end: true, pts: pts.map(p => [p.y, p.v]) }],
    yFmt: yMoney, samples: pts.map(p => p.y),
    readout: i => ({ title: String(M.cy + pts[i].y), rows: [{ cls: 'value', value: money(pts[i].v, { dec: 0 }), label: 'Projected value' }, { cls: 'dep', value: money(pts[i].base, { dec: 0 }), label: 'Without growth' }] }),
    focus: i => [{ cls: 'dep', x: pts[i].y, y: pts[i].base }, { cls: 'value', x: pts[i].y, y: pts[i].v }],
    label: 'Line chart: projected portfolio value by year, compared with the same deposits without growth'
  });
}

function renderLedger(M) {
  const list = $('#ledger-list');
  if (!M.deps.length) {
    list.replaceChildren(h('p', { class: 'empty' }, mode === 'welcome' ? 'Open your Excel file to see your deposits here, or add your first one above to start a new tracker.' : 'No deposits yet. Add one above, or bring in many at once from Excel.'));
    return;
  }
  const groups = new Map();
  for (let i = M.deps.length - 1; i >= 0; i--) {
    const d = M.deps[i], y = dayYear(d.t);
    if (!groups.has(y)) groups.set(y, []);
    groups.get(y).push(d);
  }
  if (!openYears.size) openYears.add(groups.keys().next().value);
  const frag = document.createDocumentFragment();
  for (const [y, items] of groups) {
    const total = sumOf(items, d => d.a), nWd = items.filter(d => d.a < 0).length, nDep = items.length - nWd;
    const det = h('details', { class: 'yr' });
    if (openYears.has(y) || (editing && items.some(d => d.id === editing.id))) det.open = true;
    det.addEventListener('toggle', () => { if (det.open) openYears.add(y); else openYears.delete(y); });
    const summary = h('summary', null,
      h('span', { class: 'yr-name' }, String(y)),
      h('span', { class: 'yr-meta' }, [nDep ? plural(nDep, 'deposit', 'deposits') : '', nWd ? plural(nWd, 'withdrawal', 'withdrawals') : ''].filter(Boolean).join(', ')),
      h('span', { class: 'yr-sum' }, money(total, { dec: 0 })));
    const chev = h('span', { class: 'chev', 'aria-hidden': 'true' });
    chev.innerHTML = ICON.chev;
    summary.append(chev);
    det.append(summary);
    const ul = h('ul', { class: 'rows' });
    for (const d of items) ul.append(editing && editing.id === d.id ? editRow(d) : viewRow(d));
    det.append(ul);
    frag.append(det);
  }
  list.replaceChildren(frag);
  if (editing && editing.focus) {
    const el = document.getElementById('edit-amount');
    if (el) { el.focus(); el.select(); }
    editing.focus = false;
  }
}
function viewRow(d) {
  return h('li', { class: 'row' },
    h('span', { class: 'r-date' }, fmtDay(d.t)),
    h('span', { class: 'r-amt' + (d.a < 0 ? ' neg' : '') }, money(d.a, {}), d.a < 0 ? h('span', { class: 'r-tag' }, 'out') : null),
    mode === 'example' ? h('span') : iconBtn('edit', 'Edit the ' + money(Math.abs(d.a), {}) + ' entry on ' + fmtDay(d.t), { 'data-action': 'edit-row', 'data-id': d.id }));
}
function editRow(d) {
  const li = h('li', { class: 'row editing' });
  const dateIn = h('input', { type: 'date', id: 'edit-date', value: editing.d });
  const amtIn = h('input', { type: 'text', id: 'edit-amount', inputmode: 'decimal', autocomplete: 'off', value: editing.a });
  dateIn.addEventListener('input', () => { editing.d = dateIn.value; });
  amtIn.addEventListener('input', () => { editing.a = amtIn.value; });
  const onKey = e => { if (e.key === 'Enter') { e.preventDefault(); saveEdit(); } else if (e.key === 'Escape') { e.preventDefault(); editing = null; render(); } };
  dateIn.addEventListener('keydown', onKey); amtIn.addEventListener('keydown', onKey);
  const sym = curSymbol();
  li.append(
    h('div', { class: 'row2' },
      h('label', { class: 'field', for: 'edit-date' }, h('span', { class: 'field-label' }, 'Date'), dateIn),
      h('label', { class: 'field', for: 'edit-amount' }, h('span', { class: 'field-label' }, 'Amount'),
        h('span', { class: 'money', style: { '--cur-w': Math.max(1, sym.length) + 'ch' } }, h('span', { class: 'cur' }, sym), amtIn))),
    h('p', { class: 'edit-hint' }, 'Use a minus sign for a withdrawal, like -500.'),
    h('p', { class: 'field-error', id: 'edit-error', role: 'alert' }),
    h('div', { class: 'edit-actions' },
      h('button', { type: 'button', class: 'btn btn-sm btn-primary', 'data-action': 'save-row' }, 'Save'),
      h('button', { type: 'button', class: 'btn btn-sm', 'data-action': 'cancel-row' }, 'Cancel'),
      h('span', { class: 'spacer' }),
      h('button', { type: 'button', class: 'btn btn-sm btn-danger', 'data-action': 'delete-row', 'data-id': d.id }, 'Delete')));
  return li;
}
function saveEdit() {
  if (!editing) return;
  const iso = editing.d, a = parseInputAmount(String(editing.a));
  const err = $('#edit-error');
  if (!isFinite(isoToDay(iso))) { if (err) err.textContent = 'Choose a date.'; return; }
  if (!isFinite(a) || a === 0) { if (err) err.textContent = 'Enter an amount, like 500, or -500 for a withdrawal.'; return; }
  const id = editing.id;
  editing = null;
  const prev = commitData(s => { const d = s.deposits.find(x => x.id === id); if (d) { d.d = iso; d.a = round2(a); } });
  openYears.add(dayYear(isoToDay(iso)));
  toast('Entry updated', prev);
}

/* ---------- settings ---------- */
const CURRENCIES = ['USD', 'EUR', 'GBP', 'AED', 'SAR', 'QAR', 'KWD', 'BHD', 'OMR', 'EGP', 'JOD', 'LBP', 'TRY', 'CAD', 'AUD', 'NZD', 'CHF', 'SEK', 'NOK', 'DKK', 'PLN', 'JPY', 'CNY', 'HKD', 'SGD', 'INR', 'PKR', 'ZAR', 'BRL', 'MXN'];
function renderSettings() {
  const sel = $('#set-cur');
  if (!sel.options.length) {
    let names = null;
    try { names = new Intl.DisplayNames([LOCALE], { type: 'currency' }); } catch (e) { names = null; }
    for (const c of CURRENCIES) sel.append(h('option', { value: c }, c + (names ? ' · ' + names.of(c) : '')));
  }
  if (!CURRENCIES.includes(cur())) sel.append(h('option', { value: cur() }, cur()));
  sel.value = cur();
  const keep = $('#set-keep');
  keep.checked = keepCopy;
  keep.disabled = !storageOK;
  $('#set-keep-note').textContent = !storageOK
    ? 'This browser blocks storage, so nothing can be kept here. Your Excel file has everything.'
    : keepCopy
      ? 'On: the app keeps a copy of your numbers on this device so it opens without picking the file. The copy is never uploaded. Turn it off to erase it.'
      : 'Off: nothing is stored on this device. The app forgets your numbers when it closes, and your Excel file is the only copy.';
  const fname = $('#set-file');
  fname.textContent = mode === 'welcome' ? 'No file open.' : mode === 'example' ? 'Showing example numbers.' : (file.appFormat && file.name ? file.name : 'A new tracker, not saved to a file yet') + (file.dirty ? ' · unsaved changes' : '');
  $('#set-close').hidden = mode === 'welcome';
  $('#set-save').disabled = mode === 'welcome';
  $('#set-spx').textContent = 'S&P 500 total return (dividends reinvested) from 1985 to the close of ' + fmtDate(SPX.last.d) + ' at ' + num(SPX.last.px, 2) + '. ' +
    (PWA ? 'A daily job on your GitHub site adds each new close (from FRED, the St. Louis Fed’s data service), and the app picks it up whenever it opens online. Your phone never contacts those sites.'
         : 'For later dates, use “Add a newer close” on the Overview.');
}

/* ---------- the engraved pattern behind the value ---------- */
let gProgress = REDUCED ? 1 : 0;
function drawGuilloche() {
  const cv = $('#guilloche');
  if (!cv || !cv.getContext || cv.offsetParent === null) return;
  const W = Math.max(1, Math.round(cv.clientWidth)), H = Math.max(1, Math.round(cv.clientHeight));
  const dpr = Math.min(2, window.devicePixelRatio || 1);
  if (cv.width !== W * dpr || cv.height !== H * dpr) { cv.width = W * dpr; cv.height = H * dpr; }
  const ctx = cv.getContext('2d');
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, W, H);
  const css = getComputedStyle(document.documentElement);
  const c1 = css.getPropertyValue('--accent').trim() || '#2F5FD0';
  const c2 = css.getPropertyValue('--brass').trim() || '#C07A0C';
  const dark = css.getPropertyValue('color-scheme').trim() === 'dark';
  const R = Math.min(H * 0.82, W * 0.46, 260);
  const cx = W - R * 0.3, cy = H * 0.5;
  const bands = [
    { r0: 0.30, A: 0.05, m: 10, n: 16, c: c1 },
    { r0: 0.57, A: 0.045, m: 18, n: 18, c: c1 },
    { r0: 0.85, A: 0.038, m: 26, n: 20, c: c2 }
  ];
  const steps = 640, lim = Math.max(2, Math.floor(steps * gProgress));
  ctx.lineWidth = 0.6;
  for (const b of bands) {
    ctx.strokeStyle = b.c;
    ctx.globalAlpha = dark ? 0.30 : 0.22;
    for (let k = 0; k < b.n; k++) {
      const ph = (k / b.n) * ((2 * Math.PI) / b.m) * 2;
      ctx.beginPath();
      for (let i = 0; i <= lim; i++) {
        const th = (i / steps) * 2 * Math.PI;
        const r = R * (b.r0 + b.A * Math.sin(b.m * th + ph) + b.A * 0.4 * Math.sin((b.m / 2) * th - ph * 3));
        const x = cx + r * Math.cos(th), y = cy + r * Math.sin(th);
        if (i) ctx.lineTo(x, y); else ctx.moveTo(x, y);
      }
      ctx.stroke();
    }
  }
  ctx.globalAlpha = 1;
}
let gAnimated = false;
function animateGuilloche() {
  if (gAnimated || mode === 'welcome' || tab !== 'overview') { drawGuilloche(); return; }
  gAnimated = true;
  if (REDUCED) { gProgress = 1; drawGuilloche(); return; }
  const t0 = performance.now(), dur = 1800;
  const step = now => {
    const p = Math.min(1, (now - t0) / dur);
    gProgress = 1 - Math.pow(1 - p, 3);
    drawGuilloche();
    if (p < 1) requestAnimationFrame(step);
  };
  requestAnimationFrame(step);
}

/* ---------- opening and saving Excel files ---------- */
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const readOpts = () => ({ localeFmt: LOCALE_DATE_FMT, localeDecimalComma: LOCALE_DECIMAL_COMMA });
function openErrorMessage(e) {
  const c = e && e.code;
  if (c === 'notzip') return 'That file isn’t an .xlsx workbook. If it’s an old .xls file or password-protected, open it in Excel and save a copy as .xlsx without a password.';
  if (c === 'nodecompress') return 'This device can’t unpack Excel files. Update iOS, or copy the cells in Excel and use “Add many at once”.';
  if (c === 'encrypted') return 'That workbook is password-protected. Save a copy without the password, then open it here.';
  return 'That file couldn’t be read. Try saving it again as .xlsx, or copy the cells and use “Add many at once”.';
}
async function readAnyFile(f) {
  const name = f.name || 'file';
  const head = new Uint8Array(await f.slice(0, 4).arrayBuffer());
  const isZip = head[0] === 0x50 && head[1] === 0x4B;
  if (isZip) return { kind: 'xlsx', name, wb: await XLSXLite.read(await f.arrayBuffer()) };
  if (/\.(csv|tsv|txt)$/i.test(name) || (f.type && /^text\//.test(f.type))) return { kind: 'text', name, text: await f.text() };
  const e = new Error('notzip'); e.code = 'notzip'; throw e;
}
async function openFile(f) {
  if (!f) return;
  let got;
  try { got = await readAnyFile(f); } catch (e) { toast(openErrorMessage(e)); return; }
  if (got.kind === 'xlsx' && isAppWorkbook(got.wb)) {
    if (mode === 'data' && file.dirty) {
      const ok = await confirmBox('Open ' + got.name + '?', 'Your unsaved changes in the tracker that’s open now will be lost.', 'Open file', true);
      if (!ok) return;
    }
    const r = readAppWorkbook(got.wb, readOpts());
    const prev = snapshot();
    state = emptyState();
    state.deposits = r.deposits.map(d => ({ id: newId(), d: d.d, a: d.a }));
    state.history = r.history;
    state.value = r.value;
    state.settings = r.settings;
    state.updatedAt = Date.now();
    mode = 'data';
    Object.assign(file, { name: got.name, appFormat: true, dirty: false, savedAt: null, openedAt: Date.now(), source: got.name });
    editing = null; openYears.clear(); resetZoom(); proj.annualTouched = false;
    writeCopy();
    setTab('overview');
    toast('Opened ' + got.name + ': ' + plural(r.deposits.length, 'deposit', 'deposits') + (r.skipped ? '. ' + plural(r.skipped, 'row', 'rows') + ' couldn’t be read and ' + (r.skipped === 1 ? 'was' : 'were') + ' left out' : ''), prev);
    return;
  }
  /* Someone's own spreadsheet or a CSV: show what was found before using it. */
  openImportForFile(got);
}

function excelSerial(iso) { return isoToDay(iso) + 25569; }
function buildWorkbook(st, M) {
  const deps = st.deposits.slice().sort((a, b) => (a.d < b.d ? -1 : a.d > b.d ? 1 : 0));
  const hist = st.history.slice();
  if (st.value && !hist.some(x => x.d === st.value.d)) hist.push({ d: st.value.d, a: st.value.a });
  hist.sort((a, b) => (a.d < b.d ? -1 : 1));
  const today = todayIso();
  const S = (v, style) => ({ v, style });
  const F = (f, v, style) => ({ f, v, style });
  const first = deps.length ? excelSerial(deps[0].d) : null;
  const vDate = M.hasValue ? excelSerial(st.value.d) : null;
  const yearsNow = first != null ? (Math.max(excelSerial(today), vDate || 0) - first) / 365.25 : null;
  /* Filled in once the yearly table's rows are known: the average of those yearly totals. */
  const avgRow = [S('Average deposit per calendar year'), ''];
  const summary = [
    [S('Investment summary', 'title')],
    [S('The figures in this column update by themselves in Excel. The Investment Tracker app reads the Deposits and Values sheets.', 'note')],
    [],
    [S('Current value'), F('IFERROR(INDEX(Values!B:B,MATCH(MAX(Values!A:A),Values!A:A,0)),"")', M.hasValue ? M.V : '', 'money')],
    [S('Value as of'), F('IF(COUNT(Values!A:A)>0,MAX(Values!A:A),"")', vDate != null ? vDate : '', 'date')],
    [S('Total deposited'), F('SUM(Deposits!B:B)', round2(M.total), 'money')],
    [S('First deposit'), F('IF(COUNT(Deposits!A:A)>0,MIN(Deposits!A:A),"")', first != null ? first : '', 'date')],
    [S('Years since first deposit'), F('IF(B7="","",(MAX(TODAY(),N(B5))-B7)/365.25)', yearsNow != null ? yearsNow : '', 'decimal')],
    avgRow,
    [S('Deposits up to the value date'), F('IF(B5="","",SUMIFS(Deposits!B:B,Deposits!A:A,"<="&B5))', M.hasValue ? round2(M.netCounted) : '', 'money')],
    [S('Gain'), F('IF(B5="","",B4-B10)', M.gain != null ? M.gain : '', 'money')],
    [S('ROI'), F('IF(N(B10)>0,B11/B10,"")', M.roi != null ? M.roi : '', 'percent')],
    [S('CAGR'), F('IF(AND(N(B10)>0,N(B4)>0,N(B5)>N(B7)),(B4/B10)^(365.25/(B5-B7))-1,"")', M.cagr != null ? M.cagr : '', 'percent')],
    [],
    [S('From the app when you saved on ' + fmtDate(today), 'header')],
    [S('Annualized return (XIRR)'), M.xirr != null ? S(M.xirr, 'percent') : ''],
    [S('S&P 500, same deposits'), M.sp.value != null ? S(round2(M.sp.value), 'money') : ''],
    [S('Difference vs S&P 500'), M.sp.diff != null ? S(round2(M.sp.diff), 'money') : ''],
    [],
    [S('Year', 'header'), S('Deposited', 'header'), S('Cumulative', 'header')]
  ];
  const y0 = summary.length + 1;
  M.byYear.forEach((r, i) => {
    const row = y0 + i;
    summary.push([r.year,
      F('SUMIFS(Deposits!$B:$B,Deposits!$A:$A,">="&DATE(A' + row + ',1,1),Deposits!$A:$A,"<"&DATE(A' + row + '+1,1,1))', r.sum, 'money'),
      F('SUM($B$' + y0 + ':B' + row + ')', r.cum, 'money')]);
  });
  if (M.byYear.length) avgRow[1] = F('IFERROR(AVERAGE(B' + y0 + ':B' + (y0 + M.byYear.length - 1) + '),"")', M.avgPerYear, 'money');
  const settings = [
    [S('Setting', 'header'), S('Value', 'header')],
    [S('Currency'), st.settings.cur],
    [S('S&P 500 close date'), st.settings.spx ? S(excelSerial(st.settings.spx.d), 'date') : ''],
    [S('S&P 500 close level'), st.settings.spx ? S(st.settings.spx.px, 'money') : ''],
    [],
    [S('File format'), FORMAT_MARK],
    [S('Keep the date in column A and the amount in column B on the Deposits and Values sheets. Withdrawals are negative amounts.', 'note')]
  ];
  return {
    active: 0,
    sheets: [
      { name: 'Summary', cols: [32, 18, 16], rows: summary },
      { name: 'Deposits', cols: [16, 16], freeze: 1, rows: [[S('Date', 'header'), S('Amount', 'header')]].concat(deps.map(d => [S(excelSerial(d.d), 'date'), S(d.a, 'money')])) },
      { name: 'Values', cols: [16, 18], freeze: 1, rows: [[S('Date', 'header'), S('Portfolio value', 'header')]].concat(hist.map(x => [S(excelSerial(x.d), 'date'), S(x.a, 'money')])) },
      { name: 'Settings', cols: [24, 22], rows: settings }
    ]
  };
}
function saveFileName() {
  if (file.appFormat && file.name && /\.xlsx$/i.test(file.name)) return file.name;
  return 'Investments.xlsx';
}
function saveToExcel() {
  if (mode === 'welcome') return;
  if (mode === 'example') { toast('These are example numbers. Open your file or add a deposit first.'); return; }
  const M = lastM || computeMetrics(state, getAnchors(), todayDay());
  let bytes;
  try { bytes = XLSXLite.write(buildWorkbook(state, M)); } catch (e) { toast('The Excel file couldn’t be built. Nothing was changed.'); return; }
  const name = saveFileName();
  const blob = new Blob([bytes], { type: XLSX_MIME });
  deliver(blob, name).then(ok => {
    if (!ok) return;
    Object.assign(file, { name, appFormat: true, dirty: false, savedAt: Date.now() });
    writeCopy();
    render();
  });
}
/* Hands the file to you: the Claude app's save prompt, the iPhone share sheet, or a normal download. */
async function deliver(blob, name) {
  if (downloadsCap) {
    try { await downloadsCap.save({ filename: name, data: blob }); toast('Saved ' + name + '.'); return true; }
    catch (e) {
      const c = e && e.code;
      if (c === 'declined') return false;
      if (c === 'rate_limited') { toast('A save is already waiting for your answer.'); return false; }
    }
  }
  if (IS_IOS && typeof File === 'function' && navigator.canShare) {
    const f = new File([blob], name, { type: XLSX_MIME });
    if (navigator.canShare({ files: [f] })) {
      try {
        await navigator.share({ files: [f] });
        toast('Saved. If Files asks, choose Replace to update your old copy.');
        return true;
      } catch (e) {
        if (e && e.name === 'AbortError') return false;
      }
    }
  }
  const url = URL.createObjectURL(blob);
  const a = h('a', { href: url, download: name, class: 'sr-only' });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
  toast('Downloaded ' + name + '.');
  return true;
}

/* ---------- import dialog: add many deposits, or read someone's own workbook ---------- */
/* target: 'deposits' (the ledger) or 'values' (your portfolio's value on past dates) */
const imp = { mode: 'add', target: 'deposits', tab: 'paste', fmt: 'auto', dateCol: null, amountCol: null, delim: null, src: null, sheetIdx: 0, res: null, skipDup: true, toAdd: [], openName: '', guess: null };
const IMP_TEXT = {
  deposits: { title: 'Add many deposits', label: 'In Excel, select your date and amount columns, copy them, then paste here.', placeholder: 'Date\tAmount\n15/01/2024\t500\n15/02/2024\t500\n20/03/2024\t1,250',
    drop: 'Excel (.xlsx) or CSV. The deposits in it are added to your ledger.', replace: 'Replace my existing deposits' },
  values: { title: 'Add many values', label: 'In Excel, select a date column and your portfolio value on each date, copy them, then paste here.', placeholder: 'Date\tValue\n31/01/2024\t10,250\n29/02/2024\t10,610\n31/03/2024\t11,020',
    drop: 'Excel (.xlsx) or CSV with a date column and a value column. The values are added to your value history.', replace: 'Replace my saved values' }
};
function importTabs() { return imp.target === 'values' ? ['paste', 'file'] : ['paste', 'file', 'recurring']; }
const dlgImport = $('#dlg-import');
function openDialog(d) { if (typeof d.showModal === 'function') { if (!d.open) d.showModal(); } else d.setAttribute('open', ''); }
function closeDialog(d) { if (typeof d.close === 'function') { if (d.open) d.close(); } else d.removeAttribute('open'); }
function resetImport() { Object.assign(imp, { fmt: 'auto', dateCol: null, amountCol: null, delim: null, res: null, toAdd: [], sheetIdx: 0 }); }
function setImportTarget(target) {
  target = target === 'values' ? 'values' : 'deposits';
  if (target !== imp.target) { $('#imp-text').value = ''; $('#imp-file').value = ''; imp.src = null; }
  imp.target = target;
  const T = IMP_TEXT[target];
  $('#imp-title').textContent = T.title;
  $('#imp-text-label').textContent = T.label;
  $('#imp-text').placeholder = T.placeholder;
  $('#imp-drop-note').textContent = T.drop;
  $('#imp-replace-label').textContent = T.replace;
  $('#tab-recurring').hidden = target === 'values';
}
function openImport(tabName, target) {
  resetImport();
  setImportTarget(target);
  imp.mode = 'add'; imp.src = null; imp.guess = null;
  if (!importTabs().includes(tabName)) tabName = 'paste';
  $('#imp-tabs').hidden = false;
  $('#imp-drop').hidden = false;
  $('#imp-value').hidden = true;
  $('#imp-file-status').textContent = '';
  $('#imp-replace').checked = false;
  if (!$('#rc-start').value) { const t = todayDay(); $('#rc-start').value = dayToIso(t - 365); $('#rc-end').value = dayToIso(t); }
  setImportTab(tabName || 'paste');
  openDialog(dlgImport);
  if (imp.tab === 'paste') setTimeout(() => $('#imp-text').focus(), 30);
}
function openImportForFile(got) {
  resetImport();
  setImportTarget('deposits');
  imp.mode = 'open';
  imp.openName = got.name;
  if (got.kind === 'xlsx') {
    imp.src = { kind: 'sheets', sheets: got.wb.sheets };
    imp.sheetIdx = Math.max(0, bestDepositSheet(got.wb, readOpts()));
    imp.guess = findCurrentValue(got.wb);
  } else {
    imp.src = { kind: 'text', text: got.text };
    imp.guess = null;
  }
  $('#imp-title').textContent = 'Open ' + got.name;
  $('#imp-tabs').hidden = true;
  $('#imp-drop').hidden = true;
  $('#imp-value').hidden = false;
  $('#imp-file-status').textContent = 'This reads your file on this device. Nothing in it is changed.';
  $('#imp-val-amount').value = imp.guess ? plain(imp.guess.value) : (mode === 'data' && state.value ? plain(state.value.a) : '');
  $('#imp-val-date').value = todayIso();
  $('#imp-value-note').textContent = imp.guess ? 'Found next to “' + imp.guess.label + '” on the ' + imp.guess.sheet + ' sheet. Check it, or type today’s value.' : 'I didn’t find a current value in the file. Type what your portfolio is worth today, or leave it blank.';
  setImportTab('file');
  openDialog(dlgImport);
}
function setImportTab(tabName) {
  imp.tab = tabName;
  for (const t of ['paste', 'file', 'recurring']) {
    $('#tab-' + t).setAttribute('aria-selected', String(t === tabName));
    $('#tab-' + t).tabIndex = t === tabName ? 0 : -1;
    $('#panel-' + t).hidden = t !== tabName;
  }
  analyzeImport();
}
function importOpts() {
  return { fmt: imp.fmt, localeFmt: LOCALE_DATE_FMT, localeDecimalComma: LOCALE_DECIMAL_COMMA, dateCol: imp.dateCol, amountCol: imp.amountCol, delim: imp.delim };
}
function analyzeImport() {
  if (imp.tab === 'paste') {
    const text = $('#imp-text').value;
    imp.res = text.trim() ? parseTextTable(text, importOpts()) : null;
  } else if (imp.tab === 'file') {
    const src = imp.src;
    if (src && src.kind === 'text') imp.res = parseTextTable(src.text, importOpts());
    else if (src && src.kind === 'sheets') imp.res = src.sheets.length ? analyzeRows(src.sheets[imp.sheetIdx].rows, importOpts()) : { problem: 'empty', items: [], errors: [], skipped: 0 };
    else imp.res = null;
  } else {
    const amt = parseInputAmount($('#rc-amount').value);
    const dates = recurringDates($('#rc-start').value, $('#rc-end').value, $('#rc-freq').value);
    imp.res = { recurring: true, items: isFinite(amt) && amt !== 0 ? dates.map(d => ({ d, a: round2(amt) })) : [], errors: [], skipped: 0, amt, dates };
  }
  renderImportPreview();
}
const colName = (i, header) => { const L = String.fromCharCode(65 + (i % 26)); return header && header[i] ? String(header[i]).slice(0, 20) + ' (' + L + ')' : 'Column ' + L; };
/* Column pickers and date-order controls shared by the deposits and values previews. */
function columnOptions(r, amountLabel) {
  const out = [];
  const opts = h('div', { class: 'pv-opts' });
  const mkSel = (id, label, value, onChange) => {
    const sel = h('select', { id });
    for (let i = 0; i < r.ncol; i++) sel.append(h('option', { value: String(i) }, colName(i, r.header)));
    sel.value = String(value);
    sel.addEventListener('change', () => onChange(+sel.value));
    return h('label', { class: 'field', for: id }, h('span', { class: 'field-label' }, label), sel);
  };
  opts.append(mkSel('imp-dcol', 'Dates from', r.dateCol, v => { imp.dateCol = v; if (imp.amountCol === v) imp.amountCol = null; imp.delim = r.delim || null; analyzeImport(); }));
  opts.append(mkSel('imp-acol', amountLabel, r.amountCol, v => { imp.amountCol = v; if (imp.dateCol === v) imp.dateCol = null; imp.delim = r.delim || null; analyzeImport(); }));
  if (r.hasNumDates) {
    const fsel = h('select', { id: 'imp-fmt' }, h('option', { value: 'dmy' }, 'DD/MM/YYYY'), h('option', { value: 'mdy' }, 'MM/DD/YYYY'));
    fsel.value = r.fmt;
    fsel.addEventListener('change', () => { imp.fmt = fsel.value; analyzeImport(); });
    opts.append(h('label', { class: 'field', for: 'imp-fmt' }, h('span', { class: 'field-label' }, r.fmtDetected && imp.fmt === 'auto' ? 'Dates read as (detected)' : 'Dates read as'), fsel));
  }
  out.push(opts);
  if (r.fmtAmbiguous && imp.fmt === 'auto') out.push(h('p', { class: 'warn-note' }, 'Every date here could be read either way, so I assumed ' + (r.fmt === 'mdy' ? 'month / day / year' : 'day / month / year') + '. Check the dates below and switch if they look wrong.'));
  if (r.fmtConflict) out.push(h('p', { class: 'warn-note' }, 'Some dates only make sense as day / month and others as month / day. Rows that don’t fit the chosen order are skipped.'));
  return out;
}
function renderImportPreview() {
  if (imp.target === 'values') { renderValuesPreview(); return; }
  const pv = $('#imp-preview'), go = $('#imp-go');
  const replaceWrap = $('#imp-replace-wrap');
  replaceWrap.hidden = imp.mode === 'open' || mode !== 'data' || !state.deposits.length;
  const replace = imp.mode === 'open' || (!replaceWrap.hidden && $('#imp-replace').checked);
  pv.replaceChildren();
  const r = imp.res;
  if (imp.mode === 'open' && imp.src && imp.src.kind === 'sheets' && imp.src.sheets.length > 1) {
    const ssel = h('select', { id: 'imp-sheet' });
    imp.src.sheets.forEach((s, i) => ssel.append(h('option', { value: String(i) }, s.name + (s.hidden ? ' (hidden)' : ''))));
    ssel.value = String(imp.sheetIdx);
    ssel.addEventListener('change', () => { imp.sheetIdx = +ssel.value; imp.dateCol = null; imp.amountCol = null; analyzeImport(); });
    pv.append(h('label', { class: 'field', for: 'imp-sheet' }, h('span', { class: 'field-label' }, 'Deposits are on the sheet'), ssel));
  }
  if (!r) { go.disabled = true; go.textContent = imp.mode === 'open' ? 'Use these deposits' : 'Add deposits'; return; }
  if (r.recurring) {
    if (!isFinite(r.amt) || r.amt === 0) pv.append(h('p', { class: 'note' }, 'Enter an amount to preview the schedule.'));
    else if (!r.dates.length) pv.append(h('p', { class: 'field-error' }, 'The last date needs to be on or after the first.'));
  }
  if (r.problem && !r.recurring) {
    const msg = { empty: 'Nothing to read yet.', nodate: 'I couldn’t find a column of dates' + (imp.mode === 'open' ? ' on this sheet. Pick another sheet above, if there is one.' : '. Check that the date column came along when you copied.'), noamount: 'I found dates but no column of amounts next to them.' }[r.problem];
    pv.append(h('p', { class: 'field-error' }, msg));
  }
  const existing = new Set(replace || mode !== 'data' ? [] : state.deposits.map(d => d.d + '|' + d.a.toFixed(2)));
  let dups = 0;
  const items = r.items.map(it => { const dup = existing.has(it.d + '|' + it.a.toFixed(2)); if (dup) dups++; return Object.assign({ dup }, it); });
  const toAdd = items.filter(it => !(it.dup && imp.skipDup));
  imp.toAdd = toAdd;
  if (!r.recurring && r.ncol > 1 && !r.problem) pv.append(...columnOptions(r, 'Amounts from'));
  if (items.length || r.errors.length) {
    const sumTxt = [h('strong', null, plural(toAdd.length, 'entry', 'entries')), ' ready' + (imp.mode === 'open' ? '' : ' to add')];
    if (toAdd.length) sumTxt.push(', totaling ' + money(sumOf(toAdd, x => x.a), { dec: 0 }));
    sumTxt.push('.');
    const extra = [];
    if (dups) extra.push(plural(dups, 'row matches', 'rows match') + ' an entry already in your ledger');
    if (r.errors.length) extra.push(plural(r.errors.length, 'row', 'rows') + ' couldn’t be read');
    if (r.header) extra.push('the header row was skipped');
    if (r.skipped) extra.push(plural(r.skipped, 'row', 'rows') + ' without a date or amount ' + (r.skipped === 1 ? 'was' : 'were') + ' skipped');
    pv.append(h('p', { class: 'pv-summary' }, ...sumTxt, extra.length ? ' ' + extra.join('; ').replace(/^./, c => c.toUpperCase()) + '.' : ''));
    if (dups) {
      const cb = h('input', { type: 'checkbox', id: 'imp-skipdup' });
      cb.checked = imp.skipDup;
      cb.addEventListener('change', () => { imp.skipDup = cb.checked; renderImportPreview(); });
      pv.append(h('label', { class: 'check', for: 'imp-skipdup' }, cb, 'Skip rows that are already in my ledger'));
    }
    const rowsAll = items.map(it => ({ line: it.line, d: fmtDate(it.d), a: money(it.a, {}), st: it.dup ? (imp.skipDup ? 'Already in ledger, skipped' : 'Already in ledger') : '', cls: it.dup ? 'dup' : '' }))
      .concat(r.errors.map(e => ({ line: e.line, d: String(e.raw[0] == null ? '' : e.raw[0]), a: String(e.raw[1] == null ? '' : e.raw[1]), st: e.reason === 'date' ? 'Couldn’t read the date' : 'Couldn’t read the amount', cls: 'err' })));
    if (!r.recurring) rowsAll.sort((x, y) => (x.line || 0) - (y.line || 0));
    const shown = rowsAll.slice(0, 120);
    pv.append(h('div', { class: 'table-scroll' }, h('table', { class: 'num-table pv-table' },
      h('thead', null, h('tr', null, h('th', { scope: 'col' }, 'Date'), h('th', { scope: 'col' }, 'Amount'), h('th', { scope: 'col' }, ''))),
      h('tbody', null, shown.map(x => h('tr', null, h('td', null, x.d), h('td', null, x.a), h('td', { class: 'st ' + x.cls }, x.st)))))));
    if (rowsAll.length > shown.length) pv.append(h('p', { class: 'pv-more' }, 'Showing the first ' + shown.length + ' of ' + rowsAll.length + ' rows.'));
  }
  go.disabled = toAdd.length === 0;
  go.textContent = toAdd.length ? (imp.mode === 'open' ? 'Use these ' + plural(toAdd.length, 'entry', 'entries') : (replace ? 'Replace with ' : 'Add ') + plural(toAdd.length, 'entry', 'entries')) : (imp.mode === 'open' ? 'Use these deposits' : 'Add deposits');
}
/* Values: one per date. A date that's already saved gets the new value; the latest date becomes your current value. */
function renderValuesPreview() {
  const pv = $('#imp-preview'), go = $('#imp-go');
  const replaceWrap = $('#imp-replace-wrap');
  replaceWrap.hidden = mode !== 'data' || !state.history.length;
  const replace = !replaceWrap.hidden && $('#imp-replace').checked;
  pv.replaceChildren();
  imp.toAdd = [];
  go.disabled = true;
  go.textContent = 'Add values';
  const r = imp.res;
  if (!r) return;
  if (r.problem) {
    pv.append(h('p', { class: 'field-error' }, { empty: 'Nothing to read yet.', nodate: 'I couldn’t find a column of dates. Check that the date column came along when you copied.', noamount: 'I found dates but no column of values next to them.' }[r.problem]));
    return;
  }
  if (r.ncol > 1) pv.append(...columnOptions(r, 'Values from'));
  const saved = new Map();
  if (!replace && mode === 'data') { for (const x of state.history) saved.set(x.d, x.a); if (state.value && !saved.has(state.value.d)) saved.set(state.value.d, state.value.a); }
  const lastIdx = new Map();
  r.items.forEach((it, i) => lastIdx.set(it.d, i));
  const toAdd = [], rows = [];
  let nRepl = 0, nSame = 0, nTwice = 0, nNeg = 0;
  r.items.forEach((it, i) => {
    let st = '', cls = '';
    if (it.a < 0) { st = 'A value can’t be negative'; cls = 'err'; nNeg++; }
    else if (lastIdx.get(it.d) !== i) { st = 'Same date again further down, skipped'; cls = 'dup'; nTwice++; }
    else if (saved.has(it.d) && Math.abs(saved.get(it.d) - it.a) < 0.005) { st = 'Already saved'; cls = 'dup'; nSame++; }
    else {
      if (saved.has(it.d)) { st = 'Replaces ' + money(saved.get(it.d), { dec: 0 }); cls = 'repl'; nRepl++; }
      toAdd.push(it);
    }
    rows.push({ line: it.line, d: fmtDate(it.d), a: money(it.a, {}), st, cls });
  });
  for (const e of r.errors) rows.push({ line: e.line, d: String(e.raw[0] == null ? '' : e.raw[0]), a: String(e.raw[1] == null ? '' : e.raw[1]), st: e.reason === 'date' ? 'Couldn’t read the date' : 'Couldn’t read the value', cls: 'err' });
  rows.sort((x, y) => (x.line || 0) - (y.line || 0));
  imp.toAdd = toAdd;
  if (rows.length) {
    const extra = [];
    if (nRepl) extra.push(plural(nRepl, 'replaces a value', 'replace values') + ' you already saved for that date');
    if (nSame) extra.push(plural(nSame, 'is', 'are') + ' already saved');
    if (nTwice) extra.push(plural(nTwice, 'date appears', 'dates appear') + ' twice, so the lower row is used');
    if (nNeg) extra.push(plural(nNeg, 'row has', 'rows have') + ' a negative value');
    if (r.errors.length) extra.push(plural(r.errors.length, 'row', 'rows') + ' couldn’t be read');
    if (r.header) extra.push('the header row was skipped');
    if (r.skipped) extra.push(plural(r.skipped, 'row', 'rows') + ' without a date or value ' + (r.skipped === 1 ? 'was' : 'were') + ' skipped');
    pv.append(h('p', { class: 'pv-summary' }, h('strong', null, plural(toAdd.length, 'value', 'values')), ' ready to add.', extra.length ? ' ' + extra.join('; ').replace(/^./, c => c.toUpperCase()) + '.' : ''));
    if (toAdd.length) {
      const newest = toAdd.reduce((m, x) => (x.d > m.d ? x : m), toAdd[0]);
      const curD = replace || mode !== 'data' || !state.value ? '' : state.value.d;
      if (newest.d >= curD) pv.append(h('p', { class: 'note' }, 'Your current value becomes ' + money(newest.a, { dec: 0 }) + ', from ' + fmtDate(newest.d) + ', the latest date here.'));
    }
    const shown = rows.slice(0, 160);
    pv.append(h('div', { class: 'table-scroll' }, h('table', { class: 'num-table pv-table' },
      h('thead', null, h('tr', null, h('th', { scope: 'col' }, 'Date'), h('th', { scope: 'col' }, 'Value'), h('th', { scope: 'col' }, ''))),
      h('tbody', null, shown.map(x => h('tr', null, h('td', null, x.d), h('td', null, x.a), h('td', { class: 'st ' + x.cls }, x.st)))))));
    if (rows.length > shown.length) pv.append(h('p', { class: 'pv-more' }, 'Showing the first ' + shown.length + ' of ' + rows.length + ' rows.'));
  }
  go.disabled = !toAdd.length;
  if (toAdd.length) go.textContent = (replace ? 'Replace with ' : 'Add ') + plural(toAdd.length, 'value', 'values');
}
async function runValuesImport(items) {
  const replace = !$('#imp-replace-wrap').hidden && $('#imp-replace').checked;
  if (replace) {
    closeDialog(dlgImport);
    const ok = await confirmBox('Replace your saved values?', 'Your ' + plural(state.history.length, 'saved value', 'saved values') + ' will be removed and replaced with ' + plural(items.length, 'value', 'values') + ' from this import.', 'Replace', true);
    if (!ok) { openDialog(dlgImport); return; }
  }
  const prev = commitData(s => {
    const byDate = new Map();
    if (!replace) { for (const x of s.history) byDate.set(x.d, x.a); if (s.value && !byDate.has(s.value.d)) byDate.set(s.value.d, s.value.a); }
    for (const it of items) byDate.set(it.d, it.a);
    s.history = [...byDate].map(([d, a]) => ({ d, a })).sort((x, y) => (x.d < y.d ? -1 : 1)).slice(-3000);
    const last = s.history[s.history.length - 1];
    s.value = last ? { a: last.a, d: last.d } : null;
  });
  closeDialog(dlgImport);
  $('#imp-text').value = '';
  $('#imp-file').value = '';
  imp.src = null;
  $('#imp-file-status').textContent = '';
  toast((replace ? 'Replaced with ' : 'Added ') + plural(items.length, 'value', 'values'), prev);
}
async function runImport() {
  const items = imp.toAdd || [];
  if (!items.length) return;
  if (imp.target === 'values') { await runValuesImport(items); return; }
  if (imp.mode === 'open') {
    let value = null;
    const va = parseInputAmount($('#imp-val-amount').value), vd = $('#imp-val-date').value;
    if ($('#imp-val-amount').value.trim()) {
      if (!isFinite(va) || va < 0) { $('#imp-val-error').textContent = 'Enter the value as a number, like 25,000, or leave it blank.'; return; }
      if (!isFinite(isoToDay(vd))) { $('#imp-val-error').textContent = 'Choose the date of this value.'; return; }
      value = { a: round2(va), d: vd };
    }
    $('#imp-val-error').textContent = '';
    if (mode === 'data' && file.dirty) {
      closeDialog(dlgImport);
      const ok = await confirmBox('Replace what’s open?', 'Your unsaved changes in the tracker that’s open now will be lost.', 'Replace', true);
      if (!ok) { openDialog(dlgImport); return; }
    }
    const prev = snapshot();
    const settings = state.settings;
    state = emptyState();
    state.settings = mode === 'data' ? settings : Object.assign(emptyState().settings, { cur: settings.cur });
    state.deposits = items.map(it => ({ id: newId(), d: it.d, a: it.a }));
    if (value) { state.value = value; state.history = [{ d: value.d, a: value.a }]; }
    state.updatedAt = Date.now();
    mode = 'data';
    Object.assign(file, { name: null, appFormat: false, dirty: true, savedAt: null, openedAt: Date.now(), source: imp.openName });
    closeDialog(dlgImport);
    editing = null; openYears.clear(); resetZoom(); proj.annualTouched = false;
    writeCopy();
    setTab('overview');
    toast('Loaded ' + plural(items.length, 'entry', 'entries') + ' from ' + imp.openName + '. Save to create your tracker file.', prev);
    return;
  }
  const replace = !$('#imp-replace-wrap').hidden && $('#imp-replace').checked;
  if (replace) {
    closeDialog(dlgImport);
    const ok = await confirmBox('Replace your deposits?', 'Your ' + plural(state.deposits.length, 'existing entry', 'existing entries') + ' will be removed and replaced with ' + plural(items.length, 'entry', 'entries') + ' from this import.', 'Replace', true);
    if (!ok) { openDialog(dlgImport); return; }
  }
  const prev = commitData(s => {
    const add = items.map(it => ({ id: newId(), d: it.d, a: it.a }));
    s.deposits = replace ? add : s.deposits.concat(add);
  });
  closeDialog(dlgImport);
  $('#imp-text').value = '';
  $('#imp-file').value = '';
  imp.src = null;
  $('#imp-file-status').textContent = '';
  openYears.clear();
  toast((replace ? 'Replaced with ' : 'Added ') + plural(items.length, 'entry', 'entries'), prev);
}
async function handleAddFile(f) {
  imp.src = null; resetImport();
  if (!f) { analyzeImport(); return; }
  $('#imp-file-status').textContent = 'Reading ' + f.name + '…';
  try {
    const got = await readAnyFile(f);
    if (got.kind === 'xlsx') {
      const want = imp.target === 'values' ? 'values' : 'deposits';
      const wb = isAppWorkbook(got.wb) ? { sheets: got.wb.sheets.filter(s => String(s.name).trim().toLowerCase() === want) } : got.wb;
      imp.src = { kind: 'sheets', sheets: wb.sheets.length ? wb.sheets : got.wb.sheets };
      imp.sheetIdx = Math.max(0, bestDepositSheet({ sheets: imp.src.sheets }, readOpts()));
      $('#imp-file-status').textContent = f.name + (imp.src.sheets.length > 1 ? ' · using the “' + imp.src.sheets[imp.sheetIdx].name + '” sheet' : '');
    } else {
      imp.src = { kind: 'text', text: got.text };
      $('#imp-file-status').textContent = f.name;
    }
  } catch (e) {
    imp.src = null;
    $('#imp-file-status').textContent = openErrorMessage(e);
  }
  analyzeImport();
}

/* ---------- confirm dialog ---------- */
const dlgConfirm = $('#dlg-confirm');
function confirmBox(title, body, okLabel, danger) {
  return new Promise(resolve => {
    $('#cf-title').textContent = title;
    $('#cf-body').textContent = body;
    const ok = $('#cf-ok');
    ok.textContent = okLabel || 'Confirm';
    ok.className = 'btn ' + (danger ? 'btn-danger' : 'btn-primary');
    dlgConfirm.returnValue = '';
    const onClose = () => { dlgConfirm.removeEventListener('close', onClose); resolve(dlgConfirm.returnValue === 'ok'); };
    dlgConfirm.addEventListener('close', onClose);
    openDialog(dlgConfirm);
  });
}

/* ---------- events ---------- */
const actions = {
  'open-file': () => { closeDialog($('#dlg-settings')); const f = $('#open-file'); f.value = ''; f.click(); },
  'save': () => { closeDialog($('#dlg-settings')); saveToExcel(); },
  'settings': () => { renderSettings(); openDialog($('#dlg-settings')); },
  'close-settings': () => closeDialog($('#dlg-settings')),
  'new-tracker': () => {
    const prev = snapshot();
    state = emptyState(); mode = 'data';
    Object.assign(file, { name: null, appFormat: false, dirty: false, savedAt: null, openedAt: null, source: null });
    resetZoom();
    writeCopy();
    setTab('deposits');
    toast('New tracker started. Add your deposits, then save to Excel.', prev);
    setTimeout(() => $('#add-amount').focus(), 60);
  },
  'show-example': () => { state = makeExample(todayDay()); mode = 'example'; Object.assign(file, { name: null, appFormat: false, dirty: false, savedAt: null, openedAt: null, source: null }); openYears.clear(); resetZoom(); setTab('overview'); },
  'close-file': async () => {
    if (mode === 'data' && file.dirty) {
      closeDialog($('#dlg-settings'));
      const ok = await confirmBox('Close without saving?', 'Changes that aren’t in your Excel file will be lost.', 'Close file', true);
      if (!ok) return;
    }
    const settings = { cur: state.settings.cur, spx: null };
    state = emptyState(); state.settings = settings; mode = 'welcome';
    Object.assign(file, { name: null, appFormat: false, dirty: false, savedAt: null, openedAt: null, source: null });
    editing = null; openYears.clear(); resetZoom();
    writeCopy();
    closeDialog($('#dlg-settings'));
    setTab('overview');
  },
  'erase-all': async () => {
    closeDialog($('#dlg-settings'));
    const ok = await confirmBox('Erase everything from this app?', 'This clears the numbers in the app' + (keepCopy ? ' and the copy kept on this device' : '') + '. Your Excel file isn’t touched.', 'Erase', true);
    if (!ok) return;
    dropCopy(); keepCopy = false;
    state = emptyState(); mode = 'welcome';
    Object.assign(file, { name: null, appFormat: false, dirty: false, savedAt: null, openedAt: null, source: null });
    editing = null; openYears.clear(); resetZoom();
    setTab('overview');
    toast('Erased from this app.');
  },
  'open-import': btn => openImport(btn.dataset.tab, btn.dataset.target),
  'zoom-preset': btn => { resetZoom(); zoom.preset = btn.dataset.k; drawGrowth(); },
  'focus-value': () => { setTab('overview'); const el = $('#val-amount'); setTimeout(() => el.focus(), 50); },
  'toggle-history': () => { histOpen = !histOpen; renderValueCard(lastM); },
  'del-snap': btn => {
    const d = btn.dataset.d;
    const prev = commitData(s => {
      s.history = s.history.filter(x => x.d !== d);
      if (s.value && s.value.d === d) { const rest = s.history.slice().sort((a, b) => (a.d < b.d ? 1 : -1)); s.value = rest.length ? { a: rest[0].a, d: rest[0].d } : null; }
    });
    toast('Saved value removed', prev);
  },
  'edit-row': btn => {
    const d = state.deposits.find(x => x.id === btn.dataset.id);
    if (!d) return;
    editing = { id: d.id, d: d.d, a: plain(d.a), focus: true };
    renderLedger(lastM);
  },
  'cancel-row': () => { editing = null; renderLedger(lastM); },
  'save-row': () => saveEdit(),
  'delete-row': btn => {
    const id = btn.dataset.id;
    const d = state.deposits.find(x => x.id === id);
    editing = null;
    const prev = commitData(s => { s.deposits = s.deposits.filter(x => x.id !== id); });
    toast('Deleted ' + (d ? money(d.a, {}) + ' on ' + fmtDate(d.d) : 'entry'), prev);
  },
  'edit-spx': () => {
    spxOpen = true; renderBenchmark(lastM);
    const t = todayDay(), last = isoToDay(SPX.last.d);
    $('#spx-date').value = state.settings.spx ? state.settings.spx.d : dayToIso(Math.max(t - 1, last + 1));
    $('#spx-level').value = state.settings.spx ? plain(state.settings.spx.px) : '';
    $('#spx-error').textContent = '';
    setTimeout(() => $('#spx-level').focus(), 30);
  },
  'cancel-spx': () => { spxOpen = false; renderBenchmark(lastM); },
  'clear-spx': () => { spxOpen = false; commitSettings(s => { s.settings.spx = null; }); toast('Back to the built-in S&P data'); },
  'proj-preset': btn => { proj.rate = +btn.dataset.rate; proj.preset = btn.dataset.k; renderProjection(lastM); },
  'undo': () => {
    if (undoSnapshot === undefined) return;
    const snap = undoSnapshot;
    undoSnapshot = undefined;
    restore(snap);
    if (mode === 'data') { file.dirty = true; writeCopy(); render(); }
    $('#toast').classList.remove('show');
  }
};
document.addEventListener('click', e => {
  const btn = e.target.closest('[data-action]');
  if (!btn || btn.disabled) return;
  const fn = actions[btn.dataset.action];
  if (fn) { e.preventDefault(); fn(btn, e); }
});

$$('.tabbar [role="tab"]').forEach(b => {
  b.addEventListener('click', () => setTab(b.dataset.tab));
  b.addEventListener('keydown', e => {
    if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
      e.preventDefault();
      const next = tab === 'overview' ? 'deposits' : 'overview';
      setTab(next); $('#tabbtn-' + next).focus();
    }
  });
});

let kind = 'dep';
function setKind(k) {
  kind = k;
  $('#kind-dep').setAttribute('aria-checked', String(k === 'dep'));
  $('#kind-wd').setAttribute('aria-checked', String(k === 'wd'));
  $('#kind-dep').tabIndex = k === 'dep' ? 0 : -1;
  $('#kind-wd').tabIndex = k === 'wd' ? 0 : -1;
  $('#add-submit').textContent = k === 'dep' ? 'Add deposit' : 'Add withdrawal';
}
$('#kind-dep').addEventListener('click', () => setKind('dep'));
$('#kind-wd').addEventListener('click', () => setKind('wd'));
$('.seg').addEventListener('keydown', e => {
  if (['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(e.key)) { e.preventDefault(); setKind(kind === 'dep' ? 'wd' : 'dep'); $('#kind-' + kind).focus(); }
});
setKind('dep');

$('#add-date').value = todayIso();
$('#add-form').addEventListener('submit', e => {
  e.preventDefault();
  const iso = $('#add-date').value, err = $('#add-error');
  let a = parseInputAmount($('#add-amount').value);
  if (!isFinite(isoToDay(iso))) { err.textContent = 'Choose the date of the deposit.'; $('#add-date').focus(); return; }
  if (!isFinite(a) || a === 0) { err.textContent = 'Enter an amount, like 500.'; $('#add-amount').focus(); return; }
  err.textContent = '';
  a = round2(Math.abs(a) * (kind === 'wd' || a < 0 ? -1 : 1));
  const prev = commitData(s => { s.deposits.push({ id: newId(), d: iso, a }); });
  openYears.add(dayYear(isoToDay(iso)));
  renderLedger(lastM);
  $('#add-amount').value = '';
  toast('Added ' + money(a, {}) + ' on ' + fmtDate(iso), prev);
  $('#add-amount').focus();
});
$('#add-amount').addEventListener('input', () => { $('#add-error').textContent = ''; });

function saveValue() {
  const raw = $('#val-amount').value.trim();
  const iso = $('#val-date').value || todayIso();
  const err = $('#val-error');
  if (!raw) { err.textContent = state.value ? '' : 'Enter what your portfolio is worth, like 25,000.'; return; }
  const a = parseInputAmount(raw);
  if (!isFinite(a) || a < 0) { err.textContent = 'Enter the value as a number, like 25,000.'; return; }
  if (!isFinite(isoToDay(iso))) { err.textContent = 'Choose the date of this value.'; return; }
  err.textContent = '';
  if (mode === 'data' && state.value && state.value.a === round2(a) && state.value.d === iso) return;
  const prev = commitData(s => { s.value = { a: round2(a), d: iso }; upsertHistory(s, iso, round2(a)); });
  toast('Value updated: ' + money(a, { dec: 0 }) + ' on ' + fmtDate(iso), prev);
}
$('#value-form').addEventListener('submit', e => { e.preventDefault(); saveValue(); $('#val-amount').blur(); });
$('#val-amount').addEventListener('focus', e => setTimeout(() => e.target.select(), 0));
$('#val-amount').addEventListener('input', () => { $('#val-error').textContent = ''; });

$('#spx-form').addEventListener('submit', e => {
  e.preventDefault();
  const iso = $('#spx-date').value, px = parseInputAmount($('#spx-level').value), err = $('#spx-error');
  if (!(isoToDay(iso) > isoToDay(SPX.last.d))) { err.textContent = 'Use a date after ' + fmtDate(SPX.last.d) + ', the last close already stored.'; return; }
  if (isoToDay(iso) > todayDay()) { err.textContent = 'That date is in the future.'; return; }
  if (!(px >= 1000 && px <= 100000)) { err.textContent = 'Enter the index level, like 7,400.'; return; }
  spxOpen = false;
  commitSettings(s => { s.settings.spx = { d: iso, px: round2(px) }; });
  toast('S&P 500 updated to ' + num(px, 2) + ' on ' + fmtDate(iso));
});

$('#set-cur').addEventListener('change', e => { const c = e.target.value; commitSettings(s => { s.settings.cur = c; }); });
$('#set-keep').addEventListener('change', e => {
  keepCopy = e.target.checked && storageOK;
  if (keepCopy) writeCopy(); else dropCopy();
  renderSettings();
  toast(keepCopy ? 'A copy is now kept on this device.' : 'The copy on this device was erased.');
});

$('#p-years').addEventListener('input', e => { proj.years = +e.target.value; renderProjection(lastM); });
$('#p-rate').addEventListener('input', e => { proj.rate = +e.target.value / 100; proj.preset = ''; renderProjection(lastM); });
$('#p-annual').addEventListener('input', e => { const v = parseInputAmount(e.target.value); if (isFinite(v) && v >= 0) { proj.annual = v; proj.annualTouched = true; renderProjection(lastM); } });
$('#p-annual').addEventListener('blur', e => { e.target.value = plain(proj.annual); });

$$('#imp-tabs [role="tab"]').forEach(t => {
  t.addEventListener('click', () => setImportTab(t.dataset.tab));
  t.addEventListener('keydown', e => {
    const order = importTabs();
    if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
      e.preventDefault();
      const i = (order.indexOf(imp.tab) + (e.key === 'ArrowRight' ? 1 : order.length - 1)) % order.length;
      setImportTab(order[i]); $('#tab-' + order[i]).focus();
    }
  });
});
let pasteTimer = null;
$('#imp-text').addEventListener('input', () => { clearTimeout(pasteTimer); Object.assign(imp, { dateCol: null, amountCol: null, delim: null, fmt: 'auto' }); pasteTimer = setTimeout(analyzeImport, 180); });
$('#imp-file').addEventListener('change', e => handleAddFile(e.target.files && e.target.files[0]));
['rc-amount', 'rc-start', 'rc-end'].forEach(id => $('#' + id).addEventListener('input', analyzeImport));
$('#rc-freq').addEventListener('change', analyzeImport);
$('#imp-replace').addEventListener('change', renderImportPreview);
$('#imp-go').addEventListener('click', runImport);
$('#imp-form').addEventListener('submit', e => { if (e.submitter && e.submitter.value === 'cancel') return; e.preventDefault(); });
$('#open-file').addEventListener('change', e => openFile(e.target.files && e.target.files[0]));

/* resize, theme, offline */
let resizeTimer = null, lastW = window.innerWidth;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => { drawGuilloche(); if (Math.abs(window.innerWidth - lastW) > 2) { lastW = window.innerWidth; renderCharts(); } }, 120);
});
new MutationObserver(() => drawGuilloche()).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme', 'class'] });
if (window.matchMedia) { const mq = matchMedia('(prefers-color-scheme: dark)'); if (mq.addEventListener) mq.addEventListener('change', drawGuilloche); }
document.addEventListener('visibilitychange', () => { if (!document.hidden) renderHeader(); });
setInterval(() => { if (!document.hidden) renderHeader(); }, 60000);
if (PWA && 'serviceWorker' in navigator && window.isSecureContext) {
  window.addEventListener('load', () => { navigator.serviceWorker.register('sw.js').catch(() => { /* runs online-only without it */ }); });
}

/* Always opens on Overview: current value and KPIs first. */
setTab('overview', { keepScroll: true });
})();
