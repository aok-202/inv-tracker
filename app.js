/* build 2026-11-12.1 */
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
/* Replace an element's children; anything missing (null, false) is skipped rather than printed. */
function fill(el, ...kids) { el.replaceChildren(...kids.filter(k => k != null && k !== false)); }
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
  chev: '<svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 3.5l4.5 4.5L6 12.5"/></svg>',
  up: '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 14l4.5-4.5 3 3L17 6"/><path d="M12.5 6H17v4.5"/></svg>',
  down: '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6l4.5 4.5 3-3L17 14"/><path d="M12.5 14H17V9.5"/></svg>',
  cal: '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="4.5" width="14" height="12.5" rx="2"/><path d="M3 8.5h14M7 3v3M13 3v3"/></svg>',
  pie: '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 4a7 7 0 1 0 7 7H9z"/><path d="M12 2.7A6.6 6.6 0 0 1 17.3 8H12z"/></svg>',
  pace: '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="10" cy="10" r="7"/><path d="M10 6.8v6.4M6.8 10h6.4"/></svg>',
  peak: '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2.5 16l5-8.5 3 4.5 2-3 5 7z"/></svg>',
  goal: '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="10" cy="10" r="7"/><circle cx="10" cy="10" r="3"/></svg>',
  clock: '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="10" cy="10" r="7"/><path d="M10 6v4.2l2.8 1.8"/></svg>'
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
/* Numbers on screen carry a true minus sign, which lines up with the digits (a hyphen sits low and short). */
const MINUS = '\u2212';
const tm = s => s.replace(/-/g, MINUS);
/* "Hide amounts" shows dots in place of money (percentages stay). It's never applied while an import is being checked. */
let hideAmounts = false;
const importOpen = () => { const d = document.getElementById('dlg-import'); return !!(d && d.open); };
function money(n, o) {
  o = o || {};
  if (n == null || !isFinite(n)) return '—';
  if (hideAmounts && !o.show && !importOpen()) return (o.sign && n ? (n > 0 ? '+' : MINUS) : n < 0 ? MINUS : '') + curSymbol() + (o.compact ? '\u2022\u2022' : '\u2022\u2022\u2022\u2022');
  const dec = o.dec != null ? o.dec : (Math.round(n) !== n && Math.abs(n) < 100000 ? 2 : 0);
  const opts = { style: 'currency', currency: cur(), currencyDisplay: 'narrowSymbol', minimumFractionDigits: dec, maximumFractionDigits: dec };
  if (o.sign) opts.signDisplay = 'exceptZero';
  if (o.compact) { opts.notation = 'compact'; opts.minimumFractionDigits = 0; opts.maximumFractionDigits = 1; }
  return tm(nf(opts).format(n));
}
function num(n, dec) { if (n == null || !isFinite(n)) return '—'; return tm(nf({ minimumFractionDigits: dec, maximumFractionDigits: dec }).format(n)); }
function plain(n) { return nf({ maximumFractionDigits: 2 }).format(n); }
function pct(r, o) {
  o = o || {};
  if (r == null || !isFinite(r)) return '—';
  const dec = o.dec != null ? o.dec : 1;
  return tm(nf({ style: 'percent', minimumFractionDigits: dec, maximumFractionDigits: dec, signDisplay: o.sign === false ? 'auto' : 'exceptZero' }).format(r));
}
const dfLong = new Intl.DateTimeFormat(LOCALE, { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' });
const dfShort = new Intl.DateTimeFormat(LOCALE, { day: 'numeric', month: 'short', timeZone: 'UTC' });
const dfMonYr = new Intl.DateTimeFormat(LOCALE, { month: 'short', year: '2-digit', timeZone: 'UTC' });
const dfTime = new Intl.DateTimeFormat(LOCALE, { hour: 'numeric', minute: '2-digit' });
const dfMonthYear = new Intl.DateTimeFormat(LOCALE, { month: 'short', year: 'numeric', timeZone: 'UTC' });
const dfMonthNarrow = new Intl.DateTimeFormat(LOCALE, { month: 'narrow', timeZone: 'UTC' });
const fmtMonth = (y, m) => dfMonthYear.format(new Date(Date.UTC(y, m - 1, 15)));
/* the gap between two percentages, in points: +1.2 pts */
const ptsDiff = d => (d == null || !isFinite(d) ? '—' : (d > 0 ? '+' : d < 0 ? MINUS : '') + num(Math.abs(d) * 100, 1) + ' pts');
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
function emptyState() { return { v: 1, deposits: [], value: null, history: [], settings: { cur: 'USD', spx: null, goal: null, plan: null }, updatedAt: 0 }; }
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
    const g = o.settings.goal;
    if (typeof g === 'number' && isFinite(g) && g > 0) s.settings.goal = round2(g);
    const pl = o.settings.plan;
    if (pl && typeof pl === 'object') {
      const yrs = Math.round(+pl.years), rate = +pl.rate;
      s.settings.plan = {
        years: yrs >= 1 && yrs <= 40 ? yrs : 10,
        annual: typeof pl.annual === 'number' && isFinite(pl.annual) && pl.annual >= 0 ? round2(pl.annual) : null,
        rate: rate >= 0 && rate <= 0.15 ? rate : 0.07
      };
    }
  }
  s.updatedAt = typeof o.updatedAt === 'number' ? o.updatedAt : 0;
  return s;
}
const clone = o => JSON.parse(JSON.stringify(o));
function upsertHistory(s, iso, a) {
  s.history = s.history.filter(x => x.d !== iso);
  s.history.push({ d: iso, a });
  s.history.sort((x, y) => (x.d < y.d ? -1 : 1));
  if (s.history.length > 20000) s.history = s.history.slice(-20000);
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
  /* An invented portfolio. It moves with the market each month, a little better in some months and a
     little worse in others, and keeps a value for every month end so each tab has something to show. */
  const A = buildAnchors(SPX, null);
  const flows = deps.map(d => ({ t: isoToDay(d.d), a: d.a })).sort((x, y) => x.t - y.t);
  let v = 0, prev = flows[0].t, fi = 0, k = 0;
  const step = t => {
    const wobble = 0.011 * Math.sin(k * 1.7) + 0.006 * Math.sin(k * 0.37 + 1) - 0.0006;
    const span = Math.max(1, t - prev);
    const rate = (trAt(A, t) / trAt(A, prev)) * (1 + wobble * (span / 30.44));
    v *= rate;
    while (fi < flows.length && flows[fi].t <= t) { const f = flows[fi++]; v += f.a * Math.pow(rate, (t - f.t) / span); }
    prev = t; k++;
  };
  let y = dayYear(flows[0].t), m = dayMonth(flows[0].t);
  for (;;) {
    const t = Math.round(Date.UTC(y, m, 0) / DAY_MS);
    if (t >= today) break;
    step(t);
    s.history.push({ d: dayToIso(t), a: Math.round(v) });
    m++; if (m > 12) { m = 1; y++; }
  }
  step(today);
  s.value = { a: Math.round(v), d: dayToIso(today) };
  upsertHistory(s, s.value.d, s.value.a);
  s.settings.goal = 250000;
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
    const prefs = { hide: hideAmounts };
    localStorage.setItem(COPY_KEY, JSON.stringify(mode === 'data' ? { keep: true, prefs, state, file } : { keep: true, prefs }));
  } catch (e) { /* storage full or blocked: the Excel file still has everything */ }
}
function dropCopy() { if (!storageOK) return; try { localStorage.removeItem(COPY_KEY); } catch (e) { /* ignore */ } }
(() => {
  if (!storageOK) return;
  let c = null;
  try { c = JSON.parse(localStorage.getItem(COPY_KEY) || 'null'); } catch (e) { c = null; }
  if (!c || !c.keep) return;
  keepCopy = true;
  if (c.prefs && c.prefs.hide) hideAmounts = true;
  if (c.state) { state = normalize(c.state); Object.assign(file, c.file || {}); mode = 'data'; }
})();

let anchors = buildAnchors(SPX, state.settings.spx);
let anchorsKey = JSON.stringify(state.settings.spx);
/* New S&P 500 prices, taken in without reloading the page. The text is this site's own
   spx-data.js, fetched by the offline helper (sw.js) on request: the page itself can't download. */
let spxChecked = 0, spxBusy = false, swChecked = Date.now();
const spxCanCheck = () => PWA && 'serviceWorker' in navigator && !!navigator.serviceWorker.controller;
function applySpx(text) {
  const m = /const SPX = (\{[\s\S]*?\});/.exec(text || '');
  if (!m) return false;
  let d;
  try { d = JSON.parse(m[1]); } catch (e) { return false; }
  if (!d || !Array.isArray(d.tr) || !d.tr.length || !d.last || !/^\d{4}-\d{2}-\d{2}$/.test(d.last.d) || !(d.last.px > 0) || !(d.last.tr > 0)) return false;
  if (!(d.last.d > SPX.last.d)) return false;
  SPX.tr = d.tr;
  SPX.last = d.last;
  if (d.div != null) SPX.div = d.div;
  anchorsKey = null;
  return true;
}
function checkSpx(manual) {
  if (!spxCanCheck()) {
    if (manual) toast('Close the app fully and open it again to load the latest S&P 500 prices.', undefined, 6000);
    return;
  }
  if (spxBusy) return;
  spxBusy = true;
  if (manual) { if (lastM && mode !== 'welcome') renderBenchmark(lastM); renderSettings(); }
  const ch = new MessageChannel();
  let finished = false;
  const finish = reply => {
    if (finished) return;
    finished = true;
    spxBusy = false;
    spxChecked = Date.now();
    const updated = !!(reply && reply.ok && applySpx(reply.text));
    if (updated) {
      render();
      toast('S&P 500 prices updated to the close of ' + fmtDay(isoToDay(SPX.last.d)) + '.', undefined, 4500);
    } else {
      if (lastM && mode !== 'welcome') renderBenchmark(lastM);
      renderSettings();
      if (manual) toast(reply && reply.ok ? 'No newer S&P 500 prices yet. The latest close here is ' + fmtDay(isoToDay(SPX.last.d)) + '.'
        : 'Couldn’t check for new S&P 500 prices. Make sure you’re online, then try again.', undefined, 5000);
    }
  };
  setTimeout(() => finish(null), 20000);
  ch.port1.onmessage = e => finish(e.data);
  try { navigator.serviceWorker.controller.postMessage({ type: 'spx-refresh' }, [ch.port2]); } catch (e) { finish(null); }
}
function getAnchors() {
  const k = JSON.stringify(state.settings.spx);
  if (k !== anchorsKey) { anchors = buildAnchors(SPX, state.settings.spx); anchorsKey = k; }
  return anchors;
}

/* Where the page runs inside an app that offers its own save prompt, files go through that. Otherwise this stays empty. */
let downloadsCap = null;

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
function toast(msg, prev, ms) {
  $('#toast-msg').textContent = msg;
  undoSnapshot = prev;
  $('#toast-undo').hidden = prev === undefined;
  const t = $('#toast');
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), ms || (prev === undefined ? 3200 : 7000));
}

/* ---------- tabs ---------- */
const TABS = ['overview', 'deposits', 'performance', 'plan'];
let tab = 'overview';
const scrollPos = {};
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
}

/* ---------- rendering ---------- */
let editing = null;
const openYears = new Set();
let histOpen = false, histAll = false, valueOpen = false;
let spxOpen = false;
const proj = { years: 10, annual: null, annualTouched: false, rate: 0.07, preset: '7' };
const perf = { preset: 'all' };
let lastM = null, lastG = null, lastP = null;

function render() {
  const A = getAnchors();
  const M = computeMetrics(state, A, todayDay());
  lastM = M;
  lastG = mode === 'welcome' ? null : growthSeries(M, state.history, A);
  lastP = lastG ? perfSeries(lastG, M.deps, A) : null;
  const sym = curSymbol();
  $$('[data-cur]').forEach(el => {
    el.textContent = sym;
    el.parentElement.style.setProperty('--cur-w', Math.max(1, sym.length) + 'ch');
  });
  document.body.dataset.mode = mode;
  document.body.dataset.tab = tab;
  renderHeader();
  renderBanner();
  const welcome = mode === 'welcome';
  $('#welcome').hidden = !welcome;
  $('#dashboard').hidden = welcome;
  if (tab === 'overview') {
    if (!welcome) {
      $('#overview-more').hidden = !M.count && !M.hasValue;
      renderGrowth(M); renderValueForm(M); renderKpis(M); renderHighlights(M); renderYears(M);
    }
  }
  else if (tab === 'deposits') renderLedger(M);
  else if (tab === 'performance') renderPerformance(M);
  else renderPlan(M);
  renderSettings();
}

function renderHeader() {
  const status = $('#file-status'), save = $('#btn-save');
  const dirty = mode === 'data' && file.dirty;
  save.disabled = mode !== 'data';
  save.classList.toggle('dirty', dirty);
  if (mode === 'welcome') status.textContent = 'No file open';
  else if (mode === 'example') status.textContent = 'Example numbers';
  else {
    const name = file.appFormat && file.name ? file.name : 'New tracker';
    const st = file.dirty ? 'unsaved changes' : file.savedAt ? 'saved ' + ago(file.savedAt) : file.openedAt ? 'opened ' + ago(file.openedAt) : 'not saved yet';
    status.textContent = name + ', ' + st;
  }
  status.dataset.state = dirty ? 'dirty' : 'ok';
  const hb = $('#btn-hide'), label = hideAmounts ? 'Show amounts' : 'Hide amounts';
  hb.setAttribute('aria-pressed', String(hideAmounts));
  hb.setAttribute('aria-label', label);
  hb.title = label;
}
function renderBanner() {
  const b = $('#banner');
  fill(b);
  if (mode === 'example') {
    b.hidden = false;
    b.dataset.kind = 'example';
    b.append(h('p', null, h('strong', null, 'These are example numbers. '), 'Open your own file when you’re ready.'),
      h('div', { class: 'banner-actions' },
        h('button', { class: 'btn btn-primary btn-sm', type: 'button', 'data-action': 'open-file' }, 'Open Excel file'),
        h('button', { class: 'btn btn-sm', type: 'button', 'data-action': 'close-file' }, 'Close example')));
  } else if (mode === 'data' && file.dirty) {
    b.hidden = false;
    b.dataset.kind = 'dirty';
    b.append(h('p', null, h('strong', null, 'Not saved to Excel yet. '), keepCopy ? 'Your changes are kept on this device for now.' : 'Save before you close the app, or your changes are lost.'),
      h('div', { class: 'banner-actions' }, h('button', { class: 'btn btn-primary btn-sm', type: 'button', 'data-action': 'save' }, 'Save to Excel')));
  } else b.hidden = true;
}

/* ---------- the headline: your value, and what changed over the period on the chart ---------- */
function changeParts(gain, ret, label) {
  const up = gain >= 0;
  return [h('span', { class: 'chg ' + (up ? 'good' : 'bad') },
    h('span', { class: 'tri', 'aria-hidden': 'true' }, up ? '▲' : '▼'),
    h('span', null, money(gain, { sign: true, dec: 0 })),
    ret != null ? h('span', { class: 'chg-pct' }, pct(ret)) : null),
    h('span', { class: 'chg-label' }, label)];
}
/* The big figure shrinks to stay on one line, whatever the amount or the phone. */
function fitFigure(el, min, box) {
  min = min || 30;
  el.style.fontSize = '';
  const avail = (box || el).getBoundingClientRect().width;
  if (!avail) return;
  const text = document.createRange();
  text.selectNodeContents(el);
  let size = parseFloat(getComputedStyle(el).fontSize);
  while (text.getBoundingClientRect().width > avail - 0.25 && size > min) { size -= size > 34 ? 2 : 1; el.style.fontSize = size + 'px'; }
}
/* Figures in a grid of tiles shrink the same way, so a long amount is never broken across lines. */
function fitTiles(box) { for (const el of box.querySelectorAll('.kpi-value')) fitFigure(el, 14); }
/* read: a point on the chart under your finger ({ t, value, dep }), or null for today's numbers. */
function paintHero(read) {
  const M = lastM;
  if (!M || mode === 'welcome') return;
  const valEl = $('#hero-value'), chips = $('#hero-chips'), meta = $('#hero-asof');
  const G = gz.G, V = gz.V;
  $('#hero').classList.toggle('reading', !!read);
  valEl.classList.toggle('none', read ? read.value == null : !M.hasValue);
  fill(chips);
  if (read) {
    valEl.textContent = read.value != null ? money(read.value, { dec: 0 }) : 'No value saved';
    if (read.value != null) chips.append(...changeParts(read.value - read.dep, read.dep > 0 ? (read.value - read.dep) / read.dep : null, 'on ' + money(read.dep, { dec: 0 }) + ' deposited'));
    else chips.append(h('span', { class: 'chg-label' }, money(read.dep, { dec: 0 }) + ' deposited by then'));
    meta.textContent = fmtDay(read.t);
  } else {
    valEl.textContent = M.hasValue ? money(M.V, { dec: 0 }) : 'No value yet';
    const c = G && V && M.hasValue && isZoomed(G, V) ? V.cmp : null;
    const latest = G && G.vals.length ? G.vals[G.vals.length - 1].t : null;
    if (c && c.problem) chips.append(h('span', { class: 'chg-label' }, 'This period needs saved values inside it to show a change.'));
    else if (c && !(c.fromStart && c.t1 === latest)) {
      const label = c.t1 === latest ? 'since ' + fmtDay(c.t0) : c.fromStart ? 'up to ' + fmtDay(c.t1) : fmtDay(c.t0) + ' to ' + fmtDay(c.t1);
      chips.append(...changeParts(c.gainV, c.retV, label));
    } else if (M.hasValue && M.gain != null) chips.append(...changeParts(M.gain, M.roi, 'all time'));
    if (M.hasValue) {
      const age = M.today - M.vDay;
      meta.textContent = 'As of ' + fmtDate(state.value.d) + (age > 31 ? ', ' + age + ' days ago' : '');
    } else meta.textContent = M.count ? 'Add your portfolio’s value to see your returns.' : 'Add your deposits, then your portfolio’s value, to see your returns.';
  }
  fitFigure(valEl);
}

/* ---------- updating the value (opens under the headline) ---------- */
function renderValueForm(M) {
  const form = $('#value-form'), openBtn = $('#val-open');
  if (mode === 'example') valueOpen = false;
  form.hidden = !valueOpen;
  openBtn.hidden = valueOpen || mode === 'example';
  /* nothing here yet: the first step is deposits, on the next tab */
  const fresh = !M.count && !M.hasValue;
  openBtn.textContent = fresh ? 'Add deposits' : M.hasValue ? 'Update value' : 'Add value';
  openBtn.dataset.action = fresh ? 'go-deposits' : 'open-value';
  const warn = $('#val-warn');
  if (M.hasValue && M.after.length) {
    warn.hidden = false;
    warn.textContent = plural(M.after.length, 'deposit', 'deposits') + ' after ' + fmtDate(state.value.d) + (M.after.length === 1 ? ' isn’t' : ' aren’t') + ' counted in your returns yet. Update the value to include ' + (M.after.length === 1 ? 'it.' : 'them.');
  } else warn.hidden = true;
  if (!valueOpen) return;
  const list = $('#val-history'), tog = $('#val-history-toggle');
  const hist = state.history.slice().sort((a, b) => (a.d < b.d ? 1 : -1));
  tog.hidden = hist.length === 0;
  tog.textContent = (histOpen ? 'Hide' : 'Show') + ' value history (' + hist.length + ')';
  tog.setAttribute('aria-expanded', String(histOpen));
  list.hidden = !histOpen || !hist.length;
  /* the latest two years of month ends, or everything when asked; a daily history can run to thousands */
  const shown = histAll ? hist : hist.slice(0, 24);
  fill(list, ...shown.map(x => h('li', null,
    h('span', { class: 'h-date' }, fmtDate(x.d)),
    h('span', { class: 'h-val' }, money(x.a, { dec: 0 })),
    iconBtn('trash', 'Delete the value saved for ' + fmtDate(x.d), { 'data-action': 'del-snap', 'data-d': x.d }))),
    shown.length < hist.length ? h('li', { class: 'hist-more' }, h('button', { type: 'button', class: 'link', 'data-action': 'hist-all' }, 'Show all ' + hist.length + ' values')) : null);
  paintValueDelta();
}
/* As you type a new value: how it compares with the one saved now. */
function paintValueDelta() {
  const el = $('#val-delta');
  const a = parseInputAmount($('#val-amount').value);
  if (!state.value || mode !== 'data' || !isFinite(a) || a < 0 || round2(a) === state.value.a) { el.textContent = ''; el.className = 'val-delta'; return; }
  const d = a - state.value.a;
  el.className = 'val-delta ' + (d >= 0 ? 'good' : 'bad');
  el.textContent = (d >= 0 ? '▲ ' : '▼ ') + money(Math.abs(d), { dec: 0, show: true }) + (d >= 0 ? ' more' : ' less') + ' than on ' + fmtDate(state.value.d) + (state.value.a > 0 ? ' (' + pct(d / state.value.a) + ')' : '');
}

function tone(v) { return v == null || !isFinite(v) || v === 0 ? '' : v > 0 ? 'good' : 'bad'; }
function kpiTiles(tiles) {
  return tiles.map(t => h('div', { class: 'kpi' },
    h('p', { class: 'kpi-label' }, t.label),
    h('p', { class: 'kpi-value fig' + (t.tone ? ' ' + t.tone : '') + (t.value === '—' ? ' none' : '') }, t.value),
    h('p', { class: 'kpi-cap' }, t.cap)));
}
function renderKpis(M) {
  const perYr = v => (v == null ? '—' : [pct(v), h('small', null, ' a year')]);
  fill($('#kpis'), ...kpiTiles([
    { label: 'Total deposited', value: money(M.total, { dec: 0 }),
      cap: M.count ? (M.nWd ? plural(M.nDep, 'deposit', 'deposits') + ', ' + plural(M.nWd, 'withdrawal', 'withdrawals') : plural(M.nDep, 'deposit', 'deposits')) : 'No deposits yet' },
    { label: 'Average per year', value: M.avgPerYear != null ? money(M.avgPerYear, { dec: 0 }) : '—',
      cap: !M.avgYears ? 'Average of each calendar year’s total' : M.avgYears.n === 1 ? 'Only ' + M.avgYears.y0 + ' so far'
        : 'Average of ' + M.avgYears.n + ' yearly totals, ' + M.avgYears.y0 + '–' + M.avgYears.y1 },
    { label: 'Years since first deposit', value: M.years != null ? num(M.years, 2) : '—',
      cap: M.first != null ? 'Since ' + fmtDay(M.first) : 'Starts with your first deposit' },
    { label: 'Deposited in ' + M.cy, value: money(M.thisYear.sum, { dec: 0 }), cap: (M.cy - 1) + ': ' + money(M.lastYear.sum, { dec: 0 }) },
    { label: 'ROI', value: pct(M.roi), tone: tone(M.roi), cap: M.hasValue ? 'Gain ÷ deposits' : 'Needs your current value' },
    { label: 'CAGR', value: perYr(M.cagr), tone: tone(M.cagr), cap: '(Value ÷ deposits)​^(1/years) − 1' },
    { label: 'Yearly return (XIRR)', value: perYr(M.xirr), tone: tone(M.xirr), cap: 'Counts how long each deposit was invested' },
    { label: 'Against the S&P 500', value: M.sp.diff != null ? money(M.sp.diff, { sign: true, dec: 0 }) : '—', tone: tone(M.sp.diff),
      cap: M.sp.value != null ? 'Same deposits in the index: ' + money(M.sp.value, { dec: 0 }) : (M.sp.unavailable ? 'Index data starts in 1985' : 'Needs deposits and a value') }
  ]));
  fitTiles($('#kpis'));
}

/* ---------- highlights: a few plain sentences about where you stand ---------- */
const B = t => h('b', null, t);
function monthAhead(m) { const d = new Date(); return dfMonthYear.format(new Date(Date.UTC(d.getFullYear(), d.getMonth() + m, 15))); }
function planAnnual(M) {
  if (proj.annualTouched && proj.annual != null) return proj.annual;
  return M.avgPerYear > 0 ? Math.round(M.avgPerYear / 100) * 100 : 6000;
}
function buildHighlights(M) {
  const out = [], G = lastG, P = lastP;
  const add = (icon, tn, ...parts) => out.push({ icon, tone: tn, parts });
  if (!G || !M.hasValue || !G.vals.length) return out;
  const vals = G.vals, last = vals[vals.length - 1];
  const age = M.today - M.vDay;
  if (age > 31) add('clock', 'warn', 'Your value is ', B(age + ' days'), ' old. Update it to bring every number here up to date.');
  if (vals.length >= 2) {
    const p = vals[vals.length - 2], net = flowsBetween(M.deps, p.t, last.t), gain = last.a - p.a - net;
    const flow = Math.abs(net) < 0.005 ? '' : net > 0 ? (gain >= 0 ? ', on top of ' : ', after ') + money(net, { dec: 0 }) + ' deposited' : ', after ' + money(-net, { dec: 0 }) + ' taken out';
    add(gain >= 0 ? 'up' : 'down', gain >= 0 ? 'good' : 'bad', gain >= 0 ? 'Up ' : 'Down ', B(money(Math.abs(gain), { dec: 0 })), ' since ' + fmtDay(p.t) + flow + '.');
  }
  if (P) {
    const ytd = perfPeriods(P).find(r => r.k === 'ytd');
    if (ytd) {
      const net = flowsBetween(M.deps, ytd.t0, ytd.t1), gain = P.pts[ytd.j].a - P.pts[ytd.i].a - net;
      add('cal', 'plain', dayYear(ytd.t1) + ' so far: your investments ' + (gain >= 0 ? 'gained ' : 'lost '), B(money(Math.abs(gain), { dec: 0 })),
        ' (' + pct(ytd.you) + ')' + (ytd.sp != null ? ', while the S&P 500 returned ' + pct(ytd.sp) : '') + '.');
    }
  }
  if (M.gain > 0 && M.V > 0) add('pie', 'plain', 'Growth makes up ', B(pct(M.gain / M.V, { sign: false, dec: 0 })), ' of your portfolio. The rest is money you paid in.');
  const jan1 = Math.round(Date.UTC(M.cy, 0, 1) / DAY_MS), frac = (M.today - jan1 + 1) / 365.25;
  if (M.thisYear.n >= 2 && M.thisYear.sum > 0 && M.lastYear.sum > 0 && frac > 0.2 && frac < 0.97) {
    const pace = Math.round(M.thisYear.sum / frac / 100) * 100;
    add('pace', 'plain', 'At this year’s pace you’ll deposit about ', B(money(pace, { dec: 0 })), ' in ' + M.cy + '. Last year was ' + money(M.lastYear.sum, { dec: 0 }) + '.');
  }
  if (vals.length >= 4 && vals.every(v => v.a <= last.a)) add('peak', 'good', 'Your portfolio is at its ', B('highest saved value'), '.');
  const goal = state.settings.goal;
  if (goal) {
    if (M.V >= goal) add('goal', 'good', 'You’ve reached your goal of ', B(money(goal, { dec: 0 })), '.');
    else {
      const m = monthsToGoal(M.V, planAnnual(M), proj.rate, goal);
      add('goal', 'plain', B(pct(M.V / goal, { sign: false, dec: 0 })), ' of the way to your ' + money(goal, { dec: 0 }) + ' goal' + (m != null ? '. At your planned pace you’d get there around ' + monthAhead(m) : '') + '.');
    }
  }
  return out.slice(0, 5);
}
function renderHighlights(M) {
  const items = buildHighlights(M);
  $('#sec-highlights').hidden = !items.length;
  fill($('#highlights'), ...items.map(it => {
    const ic = h('span', { class: 'hl-icon ' + it.tone, 'aria-hidden': 'true' });
    ic.innerHTML = ICON[it.icon];
    return h('li', null, ic, h('p', null, ...it.parts));
  }));
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
/* Round-number ticks inside [lo, hi]. The range stays as given, so a chart isn't padded out to the next round number. */
function insideTicks(lo, hi, n) {
  if (hi === lo) hi = lo + 1;
  const eps = (hi - lo) * 1e-9;
  return { lo, hi, ticks: niceTicks(lo, hi, n).ticks.filter(v => v >= lo - eps && v <= hi + eps) };
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
/* A line chart. Series types: 'line' and 'step' (optionally with an area down to zero), 'band' (a shaded
   range, pts [x, low, high]) and 'wedge' (a filled outline, pts [x, y] around it).
   o.yZero false: the y-axis fits the data shown instead of starting at zero (zoomed-in charts).
   o.external: no pointer handling of its own; the caller drives show/hide through the returned api.
   o.tip false: no floating readout; o.onRead(i) is told which sample is being read (-1 when none).
   o.refs: horizontal reference lines [{ value, label, inRange }]. */
function lineChart(box, o) {
  fill(box);
  const W = Math.max(260, Math.floor(box.clientWidth || 600)), H = o.height;
  let yMax = -Infinity, yMin = Infinity;
  const see = v => { if (v > yMax) yMax = v; if (v < yMin) yMin = v; };
  for (const s of o.series) for (const p of s.pts) { see(p[1]); if (s.type === 'band') see(p[2]); }
  for (const r of o.refs || []) if (r.inRange) see(r.value);
  if (!isFinite(yMax)) { yMax = 0; yMin = 0; }
  let lo, hi;
  if (o.yZero === false && yMax > yMin) {
    const pad = (yMax - yMin) * 0.08;
    lo = yMin >= 0 ? Math.max(0, yMin - pad) : yMin - pad; hi = yMax + pad;
  } else {
    lo = Math.min(0, yMin); hi = Math.max(0, yMax);
    const pad = (hi - lo || 1) * 0.07;
    hi += pad; if (lo < 0) lo -= pad;
  }
  const T = insideTicks(lo, hi, H < 200 ? 4 : 5);
  const ml = axisPad(T, o.yFmt) + (o.padL || 0), mr = 10 + (o.padR || 0), mt = 12, mb = 28;
  const pw = W - ml - mr, ph = H - mt - mb;
  const X = v => ml + (o.x1 === o.x0 ? pw / 2 : ((v - o.x0) / (o.x1 - o.x0)) * pw);
  const Y = v => mt + ((T.hi - v) / (T.hi - T.lo)) * ph;
  const xy = (x, y) => X(x).toFixed(1) + ' ' + Y(y).toFixed(1);
  const svg = sv('svg', { width: W, height: H, viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': o.label });
  yAxis(svg, T, o.yFmt, ml, mr, W, Y);
  svg.append(sv('line', { x1: ml, x2: W - mr, y1: Math.round(mt + ph) + 0.5, y2: Math.round(mt + ph) + 0.5, class: 'axis' }));
  let lastRight = -Infinity;
  for (const t of o.xTicks) {
    const x = X(t.v);
    if (x < ml - 1 || x > W - mr + 1) continue;
    const tw = textW(t.label);
    const anchor = x - tw / 2 < 2 ? 'start' : x + tw / 2 > W - 2 ? 'end' : 'middle';
    const tx = anchor === 'start' ? Math.max(0, x - 2) : anchor === 'end' ? Math.min(W, x + 2) : x;
    const left = anchor === 'start' ? tx : anchor === 'end' ? tx - tw : tx - tw / 2;
    svg.append(sv('line', { x1: Math.round(x) + 0.5, x2: Math.round(x) + 0.5, y1: mt + ph, y2: mt + ph + 4, class: 'tick' }));
    if (left < lastRight + 8) continue;   /* a label that would run into the one before it is left out */
    lastRight = left + tw;
    svg.append(sv('text', { x: tx, y: H - 8, 'text-anchor': anchor }, t.label));
  }
  const baseY = Y(Math.max(T.lo, Math.min(T.hi, 0)));
  for (const s of o.series) {
    if (!s.pts.length) continue;
    if (s.type === 'band') {
      let d = '';
      s.pts.forEach((p, i) => { d += (i ? 'L' : 'M') + xy(p[0], p[2]); });
      for (let i = s.pts.length - 1; i >= 0; i--) d += 'L' + xy(s.pts[i][0], s.pts[i][1]);
      svg.append(sv('path', { d: d + 'Z', class: 'band f-' + s.cls }));
      continue;
    }
    if (s.type === 'wedge') {
      svg.append(sv('path', { d: s.pts.map((p, i) => (i ? 'L' : 'M') + xy(p[0], p[1])).join('') + 'Z', class: 'wedge f-' + s.cls }));
      continue;
    }
    let d = '';
    s.pts.forEach((p, i) => {
      const x = X(p[0]).toFixed(1), y = Y(p[1]).toFixed(1);
      if (!i) d += 'M' + x + ' ' + y;
      else if (s.type === 'step') d += 'H' + x + 'V' + y;
      else d += 'L' + x + ' ' + y;
    });
    if (s.area) svg.append(sv('path', { d: d + 'V' + baseY.toFixed(1) + 'H' + X(s.pts[0][0]).toFixed(1) + 'Z', class: 'area f-' + s.cls }));
    svg.append(sv('path', { d, class: 'line s-' + s.cls + (s.thin ? ' thin' : '') + (s.dash ? ' dash' : '') }));
    if (s.dots) for (const p of (s.dotPts || s.pts)) svg.append(sv('circle', { cx: X(p[0]).toFixed(1), cy: Y(p[1]).toFixed(1), r: 3.5, class: 'dot f-' + s.cls }));
    if (s.end) { const p = s.endPt || s.pts[s.pts.length - 1]; svg.append(sv('circle', { cx: X(p[0]).toFixed(1), cy: Y(p[1]).toFixed(1), r: 4.5, class: 'dot end f-' + s.cls })); }
  }
  for (const r of o.refs || []) {
    if (r.value < T.lo || r.value > T.hi) continue;
    const y = Math.round(Y(r.value)) + 0.5;
    svg.append(sv('line', { x1: ml, x2: W - mr, y1: y, y2: y, class: 'ref' }));
    if (r.label) svg.append(sv('text', { x: W - mr - 2, y: y - 6, 'text-anchor': 'end', class: 'ref-label' }, r.label));
  }
  const xh = sv('line', { class: 'xhair', y1: mt, y2: mt + ph, x1: -20, x2: -20, visibility: 'hidden' });
  const fg = sv('g');
  const hit = o.external ? null : sv('rect', { x: ml, y: 0, width: Math.max(1, pw), height: H - 6, fill: 'transparent', class: 'hit' });
  svg.append(xh, fg);
  if (hit) svg.append(hit);
  box.append(svg);
  const tip = o.tip === false ? null : h('div', { class: 'tip', hidden: true });
  if (tip) box.append(tip);
  let curI = -1;
  function show(i) {
    if (!o.samples.length) return;
    i = Math.max(0, Math.min(o.samples.length - 1, i));
    const changed = i !== curI;
    curI = i;
    const x = Math.round(X(o.samples[i])) + 0.5;
    xh.setAttribute('x1', x); xh.setAttribute('x2', x); xh.setAttribute('visibility', 'visible');
    fill(fg, ...o.focus(i).map(f => sv('circle', { cx: X(f.x).toFixed(1), cy: Y(f.y).toFixed(1), r: 5.5, class: 'dot focus f-' + f.cls })));
    if (o.onRead && changed) o.onRead(i);
    if (!tip) return;
    const r = o.readout(i);
    fill(tip, h('div', { class: 'tip-title' }, r.title),
      ...r.rows.map(row => h('div', { class: 'tip-row' }, h('i', { class: 'key key-' + row.cls }), h('b', null, row.value), h('span', null, row.label))));
    tip.hidden = false;
    const tw = tip.offsetWidth;
    let left = x + 14;
    if (left + tw > W) left = x - 14 - tw;
    if (left < 0) left = Math.max(0, Math.min(W - tw, x - tw / 2));
    tip.style.left = left + 'px';
  }
  function hide() {
    const was = curI;
    curI = -1;
    xh.setAttribute('visibility', 'hidden');
    fill(fg);
    if (tip) tip.hidden = true;
    if (o.onRead && was >= 0) o.onRead(-1);
  }
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
  fill(box);
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
    fill(tip, h('div', { class: 'tip-title' }, b.tip.title), ...b.tip.rows.map(row => h('div', { class: 'tip-row' }, h('i', { class: 'key key-' + row.cls }), h('b', null, row.value), h('span', null, row.label))));
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
  const G = lastG;
  gz.G = G; gz.M = M;
  if (!G) {
    gz.api = null; gz.V = null;
    $('#growth').hidden = true;
    paintHero(null);
    return;
  }
  $('#growth').hidden = false;
  drawSpark(G);
  drawGrowth();
  growthTable(G);
}

/* The outline of the space between your value line and what you'd paid in: along the value line,
   then back along the deposit steps. That space is your gain (or loss, where the line dips under). */
function wedgePath(vline, pts) {
  const a = vline[0][0], b = vline[vline.length - 1][0];
  let level = 0;
  for (const p of pts) { if (p.t > b) break; level = p.dep; }
  const out = vline.slice();
  out.push([b, level]);
  for (let k = pts.length - 1; k >= 0; k--) {
    const p = pts[k];
    if (p.t > b) continue;
    if (p.t <= a) break;
    const before = k > 0 ? pts[k - 1].dep : 0;
    out.push([p.t, level], [p.t, before]);
    level = before;
  }
  out.push([a, level]);
  return out;
}

function drawGrowth() {
  const G = gz.G, M = gz.M;
  if (!G || mode === 'welcome') return;
  const zoomable = canZoom(G), w = currentWindow(G), zoomed = isZoomed(G, w);
  const V = windowSeries(G, M.deps, getAnchors(), w.a, w.b);
  gz.V = V;
  const plot = $('#growth-plot');
  const spPts = V.pts.filter(p => p.sp != null).map(p => [p.t, p.sp]);
  const lastReal = V.vin.length ? V.vin[V.vin.length - 1] : null;
  /* A value for every day would make a scribble: draw the last of each week or month instead (whichever fits),
     always with the first and latest. Zooming in brings the detail back. The numbers still use every value. */
  const drawn = thinValues(V.vin, Math.max(24, Math.floor((plot.clientWidth || 600) / 6)));
  let vline = V.vline;
  if (drawn !== V.vin) {
    const head = V.vline.length && (!V.vin.length || V.vline[0][0] < V.vin[0].t) ? [V.vline[0]] : [];
    const tail = V.vline.length && (!V.vin.length || V.vline[V.vline.length - 1][0] > lastReal.t) ? [V.vline[V.vline.length - 1]] : [];
    vline = head.concat(drawn.map(v => [v.t, v.a]), tail);
  }
  const series = [{ cls: 'dep', type: 'step', area: true, pts: V.pts.map(p => [p.t, p.dep]) }];
  if (vline.length > 1) series.push({ cls: 'value', type: 'wedge', pts: wedgePath(vline, V.pts) });
  if (spPts.length) series.push({ cls: 'sp', type: 'line', thin: true, end: true, pts: spPts });
  if (vline.length) series.push({ cls: 'value', type: 'line', pts: vline, dots: drawn.length <= 14, dotPts: drawn.map(v => [v.t, v.a]), end: !!lastReal, endPt: lastReal ? [lastReal.t, lastReal.a] : null });
  const valBy = new Map(V.vin.map(v => [v.t, v.a]));
  /* Reading the chart stops at your own values (the points on the line), so every stop shows your
     portfolio next to the S&P 500 and your deposits. With no values in view it stops at deposits and month ends. */
  const ptBy = new Map(V.pts.map(p => [p.t, p]));
  const stops = drawn.length >= 2 ? drawn.map(v => v.t).filter(t => ptBy.has(t)) : V.pts.map(p => p.t);
  const lastP = V.pts[V.pts.length - 1];
  /* The legend carries the numbers: for the end of the period, or for the point being read. */
  const paintLegend = p => {
    const spV = p ? p.sp : spPts.length ? spPts[spPts.length - 1][1] : null;
    const youV = p ? (valBy.has(p.t) ? valBy.get(p.t) : null) : lastReal ? lastReal.a : null;
    const box = $('#growth-legend'), depV = (p || lastP).dep;
    /* three figures share the row: when full amounts won't fit a narrow phone, all three go short (12.3M) */
    let fmt = v => (v == null ? '—' : money(v, { dec: 0 }));
    const colW = ((box.clientWidth || 340) - 24) / 3;
    if (!isWide() && [youV, spV, depV].some(v => textW(fmt(v), 16) * 1.12 > colW)) fmt = v => (v == null ? '—' : money(v, { compact: true }));
    fill(box,
      vline.length ? legendItem('value', 'You', fmt(youV)) : null,
      spV != null ? legendItem('sp', 'S&P 500', fmt(spV)) : null,
      legendItem('dep', 'Deposited', fmt(depV), true));
  };
  gz.api = lineChart(plot, {
    height: isWide() ? 320 : 230, x0: w.a, x1: w.b, external: true, tip: false, yZero: !zoomed, padL: isWide() ? 0 : 8, padR: isWide() ? 0 : 6,
    xTicks: timeTicks(w.a, w.b, plot.clientWidth || 600),
    series, yFmt: yMoney, samples: stops,
    focus: i => {
      const p = ptBy.get(stops[i]), f = [{ cls: 'dep', x: p.t, y: p.dep }];
      if (p.sp != null) f.push({ cls: 'sp', x: p.t, y: p.sp });
      if (valBy.has(p.t)) f.push({ cls: 'value', x: p.t, y: valBy.get(p.t) });
      return f;
    },
    onRead: i => {
      const p = i < 0 ? null : ptBy.get(stops[i]);
      paintLegend(p);
      paintHero(p ? { t: p.t, value: valBy.has(p.t) ? valBy.get(p.t) : null, dep: p.dep } : null);
    },
    label: 'Line chart' + (zoomed ? ' from ' + fmtDay(w.a) + ' to ' + fmtDay(w.b) : '') + ': money deposited, your portfolio value, and ' + (V.rebased ? 'the S&P 500 starting from your value on ' + fmtDay(V.cmp.t0) : 'the same deposits in the S&P 500') + ', over time'
  });
  $('#growth-chart').classList.toggle('zoomed', zoomed);
  $('#growth-nav').hidden = !zoomable;
  if (zoomable) navLayout(G, w);
  renderPresets(G, zoomable);
  const spNote = spPts.length && !V.rebased ? 'S&P 500: your deposits, put into the index on the same days. ' : '';
  $('#growth-window').textContent = zoomed ? fmtDay(w.a) + ' to ' + fmtDay(w.b) + ' (' + spanText(w.b - w.a) + ')'
    : spNote + (zoomable ? 'Pinch the chart or drag the handles for a shorter period.' : '');
  renderCompare(V, zoomed);
  paintLegend(null);
  paintHero(null);
}

function renderPresets(G, zoomable) {
  const box = $('#growth-presets');
  if (!zoomable) { fill(box); return; }
  fill(box, ...PRESETS.filter(p => presetRange(G, p.k)).map(p =>
    h('button', { type: 'button', 'data-action': 'zoom-preset', 'data-k': p.k, 'aria-label': p.name, 'aria-pressed': String(zoom.preset === p.k) }, p.label)));
}

function renderCompare(V, zoomed) {
  const box = $('#growth-compare');
  if (!zoomed) { box.hidden = true; fill(box); return; }
  box.hidden = false;
  const c = V.cmp;
  const addLink = () => (mode === 'example' ? null : h('button', { type: 'button', class: 'link', 'data-action': 'open-import', 'data-target': 'values' }, 'Add many values'));
  if (c.problem) {
    const msg = c.problem === 'nosp' ? 'The S&P 500 data here starts in January 1985, so this period can’t be compared with it.'
      : c.fromStart ? 'To compare this period with the S&P 500, the app needs your portfolio value on a date inside it.'
        : 'To compare this period with the S&P 500, the app needs your portfolio value on at least two dates inside it.';
    fill(box, h('p', { class: 'note' }, msg + ' ', c.problem === 'nosp' ? null : addLink()));
    return;
  }
  const row = (cls, name, gain, ret) => h('tr', null,
    h('td', null, h('i', { class: 'key key-' + cls }), name),
    h('td', null, money(gain, { sign: true, dec: 0 })),
    h('td', null, ret != null ? pct(ret) : '—'));
  const ahead = c.diff >= 0;
  const flows = Math.abs(c.net) < 0.005 ? 'with no deposits in between' : c.net > 0 ? 'then get the same ' + money(c.net, { dec: 0 }) + ' of deposits' : 'then have the same ' + money(-c.net, { dec: 0 }) + ' taken out';
  fill(box,
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
  fill($('#growth-table'), h('table', { class: 'num-table' },
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
  fill(svg, ...kids);
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

/* ---- gestures on the chart ----
   Fingers: one finger reads values and never moves the chart; two fingers zoom and slide it.
   Fingers are followed with touch events, whose list of fingers on the screen is always complete,
   so a touch that turns into a page scroll can't leave a "finger" behind. Mouse and pen use pointer
   events: hovering reads values, dragging slides a zoomed chart, double-clicking zooms.
   The touch layer stays in place while the chart under it is redrawn, so a pinch never loses its fingers. */
const gTouch = $('#growth-touch');
let tg = null, touchesDown = 0, mdrag = null, gestureStart = null, lastPtrType = 'mouse';
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
const touchById = (list, id) => { for (let i = 0; i < list.length; i++) if (list[i].identifier === id) return list[i]; return null; };
function startTouchPinch(t1, t2) {
  const P = gz.api.plot(), w = currentWindow(gz.G), mx = (t1.clientX + t2.clientX) / 2;
  tg = { kind: 'pinch', id1: t1.identifier, id2: t2.identifier, P, a0: w.a, b0: w.b,
    d0: Math.max(24, Math.hypot(t1.clientX - t2.clientX, t1.clientY - t2.clientY)),
    dm: w.a + ((mx - P.left) / P.width) * (w.b - w.a) };
  gz.api.hide();
}
gTouch.addEventListener('touchstart', e => {
  touchesDown = e.touches.length;
  if (!gz.api) return;
  if (e.touches.length >= 2) {
    /* two fingers zoom the chart, never the page */
    if (canZoom(gz.G)) { e.preventDefault(); startTouchPinch(e.touches[0], e.touches[1]); }
    else tg = { kind: 'hold' };
    return;
  }
  const t = e.touches[0];
  tg = { kind: 'read', id: t.identifier, x0: t.clientX, y0: t.clientY, locked: false };
  gz.api.show(gz.api.idxAt(t.clientX));
}, { passive: false });
gTouch.addEventListener('touchmove', e => {
  if (!tg || !gz.api) return;
  if (tg.kind === 'pinch') {
    if (e.cancelable) e.preventDefault();
    const t1 = touchById(e.touches, tg.id1), t2 = touchById(e.touches, tg.id2);
    if (!t1 || !t2) { if (e.touches.length >= 2) startTouchPinch(e.touches[0], e.touches[1]); return; }
    const d = Math.max(24, Math.hypot(t1.clientX - t2.clientX, t1.clientY - t2.clientY)), mx = (t1.clientX + t2.clientX) / 2;
    const span = ((tg.b0 - tg.a0) * tg.d0) / d, a = tg.dm - (span * (mx - tg.P.left)) / tg.P.width;
    queueZoom(a, a + span);
    return;
  }
  if (tg.kind !== 'read') { if (e.touches.length > 1 && e.cancelable) e.preventDefault(); return; }
  const t = touchById(e.touches, tg.id);
  if (!t) return;
  const dx = t.clientX - tg.x0, dy = t.clientY - tg.y0;
  if (!tg.locked) {
    /* mostly up or down is scrolling the page: put the readout away and let the page move */
    if (Math.abs(dy) > 10 && Math.abs(dy) > Math.abs(dx)) { tg = { kind: 'scroll' }; gz.api.hide(); return; }
    if (Math.abs(dx) > 4) tg.locked = true;
  }
  /* sideways: keep reading values, and hold the page still while the finger moves */
  if (tg.locked && e.cancelable) e.preventDefault();
  gz.api.show(gz.api.idxAt(t.clientX));
}, { passive: false });
function touchEnd(e) {
  touchesDown = e.touches.length;
  if (!tg) return;
  if (!e.touches.length) {
    /* after a drag across the chart the headline returns to today's numbers; a tap keeps its reading until the next touch */
    if (tg.kind === 'read' && tg.locked && gz.api) gz.api.hide();
    tg = null;
    return;
  }
  /* a finger lifted mid-pinch: keep pinching while two are down, otherwise wait until the last one lifts */
  if (tg.kind === 'pinch' && e.touches.length < 2) tg = { kind: 'hold' };
}
gTouch.addEventListener('touchend', touchEnd);
gTouch.addEventListener('touchcancel', touchEnd);
document.addEventListener('visibilitychange', () => { tg = null; touchesDown = 0; mdrag = null; });
/* a touch anywhere else puts the readout away */
document.addEventListener('touchstart', e => { if (gz.api && !gTouch.contains(e.target)) gz.api.hide(); }, { passive: true });

gTouch.addEventListener('pointerdown', e => {
  lastPtrType = e.pointerType;
  if (e.pointerType === 'touch' || !gz.api || (e.pointerType === 'mouse' && e.button !== 0)) return;
  const w = currentWindow(gz.G);
  mdrag = e.pointerType === 'mouse' && isZoomed(gz.G, w) ? { id: e.pointerId, x0: e.clientX, a0: w.a, b0: w.b, moved: false, P: gz.api.plot() } : null;
  if (mdrag) { try { gTouch.setPointerCapture(e.pointerId); } catch (err) { /* fine without capture */ } }
  gz.api.show(gz.api.idxAt(e.clientX));
});
gTouch.addEventListener('pointermove', e => {
  if (e.pointerType === 'touch' || !gz.api) return;
  if (mdrag && e.pointerId === mdrag.id) {
    if (Math.abs(e.clientX - mdrag.x0) > 3) mdrag.moved = true;
    if (mdrag.moved) {
      const span = mdrag.b0 - mdrag.a0, a = mdrag.a0 - ((e.clientX - mdrag.x0) / mdrag.P.width) * span;
      gz.api.hide();
      queueZoom(a, a + span);
      return;
    }
  }
  gz.api.show(gz.api.idxAt(e.clientX));
});
const endMouse = e => { if (mdrag && e.pointerId === mdrag.id) mdrag = null; };
gTouch.addEventListener('pointerup', endMouse);
gTouch.addEventListener('pointercancel', endMouse);
gTouch.addEventListener('pointerleave', e => { if (e.pointerType === 'mouse' && !mdrag && gz.api) gz.api.hide(); });
/* Double-click zooms with a mouse only. A double tap does nothing, so tapping twice to read values can't zoom by surprise. */
gTouch.addEventListener('dblclick', e => { if (lastPtrType === 'mouse') zoomAround(e.clientX); });
/* Safari's trackpad pinch on a Mac arrives as gesture events. On an iPhone the touch code above does the
   zooming, so these only stop the page from zooming. */
gTouch.addEventListener('gesturestart', e => {
  e.preventDefault();
  if (touchesDown > 0 || !gz.api || !canZoom(gz.G)) { gestureStart = null; return; }
  const w = currentWindow(gz.G), P = gz.api.plot(), x = e.clientX || P.left + P.width / 2;
  gestureStart = { w, P, x, dm: w.a + ((x - P.left) / P.width) * (w.b - w.a) };
});
gTouch.addEventListener('gesturechange', e => {
  e.preventDefault();
  if (!gestureStart || touchesDown > 0 || !(e.scale > 0)) return;
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
    if (M.hasValue) vs.append(row('value', 'Your portfolio', M.V, M.xirr != null ? pct(M.xirr) + ' a year (XIRR)' : 'A yearly return needs at least a month of history'));
    vs.append(row('sp', 'S&P 500, same deposits', sp.value, sp.xirr != null ? pct(sp.xirr) + ' a year (XIRR)' : 'Index up ' + pct(sp.periodReturn) + ' since your first deposit'));
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
    spxCanCheck() ? h('button', { type: 'button', class: 'link', 'data-action': 'check-spx' }, spxBusy ? 'Checking…' : 'Check for new prices') : null,
    spxCanCheck() && !(spxOpen || mode === 'example') ? ' ' : null,
    spxOpen || mode === 'example' ? null : h('button', { type: 'button', class: 'link', 'data-action': 'edit-spx' }, 'Add a newer close')));
  fill(body, ...parts);
}

function renderYears(M) {
  const box = $('#years-chart'), tb = $('#years-tbody'), lg = $('#years-legend');
  fill(lg);
  if (M.byYear.length > 1 && M.avgPerYear > 0) lg.append(h('li', null, h('i', { class: 'key dash' }), h('span', null, 'Average per year'), h('b', null, money(M.avgPerYear, { dec: 0 }))));
  if (M.byYear.some(r => r.ytd)) lg.append(h('li', null, h('span', null, 'Lighter bar: ' + M.cy + ' so far')));
  if (!M.byYear.length) {
    fill(box, h('p', { class: 'empty' }, 'Yearly totals appear after your first deposit.'));
    $('#years-table').hidden = true;
    fill(tb);
    fill($('#years-tfoot'));
    return;
  }
  $('#years-table').hidden = false;
  const bars = M.byYear.map(r => ({
    label: String(r.year), short: '’' + String(r.year).slice(2), value: r.sum, faded: r.ytd,
    tip: { title: r.ytd ? r.year + ' so far' : String(r.year), rows: [{ cls: 'dep', value: money(r.sum, { dec: 0 }), label: plural(r.n, 'entry', 'entries') }] }
  }));
  barChart(box, {
    height: isWide() ? 240 : 210, bars, cls: 'dep', yFmt: yMoney,
    valueLabel: v => money(v, { compact: Math.abs(v) >= 10000, dec: 0 }),
    ref: M.avgPerYear > 0 && M.byYear.length > 1 ? { value: M.avgPerYear } : null,
    label: 'Bar chart: money deposited each year' + (M.byYear.some(r => r.ytd) ? '; the current year is lighter because it is not finished' : '')
  });
  fill(tb, ...M.byYear.slice().reverse().map(r => h('tr', { class: r.ytd ? 'current' : null },
    h('td', null, r.ytd ? r.year + ' so far' : String(r.year)), h('td', null, money(r.sum, { dec: 0 })), h('td', null, String(r.n)), h('td', null, money(r.cum, { dec: 0 })))));
  fill($('#years-tfoot'), h('tr', { class: 'total' }, h('td', null, 'Total'), h('td', null, money(M.total, { dec: 0 })), h('td', null, String(M.count)), h('td', null, '')));
}

/* ---------- Performance: your investments against the S&P 500, with deposits taken out of the picture ---------- */
const PERF_PRESETS = [
  { k: 'ytd', label: 'YTD', name: 'This year so far' },
  { k: '1y', label: '1Y', name: 'Past year' },
  { k: '3y', label: '3Y', name: 'Past 3 years' },
  { k: '5y', label: '5Y', name: 'Past 5 years' },
  { k: '10y', label: '10Y', name: 'Past 10 years' },
  { k: 'all', label: 'All', name: 'Since your first saved value' }
];
const diffCell = (you, sp) => (you == null || sp == null ? h('td', null, '—') : h('td', { class: tone(you - sp) }, ptsDiff(you - sp)));
const pctCell = v => h('td', null, v == null ? '—' : pct(v));

function renderPerformance(M) {
  const empty = $('#perf-empty'), body = $('#perf-body'), P = lastP;
  $('#benchmark').hidden = mode === 'welcome';
  if (mode === 'welcome') {
    empty.hidden = false; body.hidden = true;
    $('#perf-sub').textContent = '';
    fill(empty, h('h3', null, 'No file open'), h('p', null, 'Open your Excel file to see how your investments have performed.'),
      h('div', { class: 'btn-row' }, h('button', { class: 'btn btn-primary', type: 'button', 'data-action': 'open-file' }, 'Open Excel file')));
    return;
  }
  renderBenchmark(M);
  if (!P) {
    empty.hidden = false; body.hidden = true;
    $('#perf-sub').textContent = '';
    fill(empty,
      h('h3', null, 'Performance needs your value on more than one date'),
      h('p', null, 'Each time you update your value, the app keeps the earlier one. With two or more saved values it can measure how your investments did between them, apart from what you paid in.'),
      mode === 'example' ? null : h('div', { class: 'btn-row' }, h('button', { class: 'btn btn-primary', type: 'button', 'data-action': 'open-import', 'data-target': 'values' }, 'Add many values')),
      h('p', { class: 'note' }, 'To fill in the past in one go, paste dates and values from Excel or from a broker’s export.'));
    return;
  }
  empty.hidden = true; body.hidden = false;
  const periods = perfPeriods(P), years = perfYears(P), months = perfMonths(P), stats = perfStats(P, months);
  $('#perf-sub').textContent = 'How your investments did, whatever you paid in and when. Measured between your ' + P.pts.length + ' saved values since ' + fmtDay(P.t0) + '.';
  renderPerfChart(P, periods);
  renderPerfTable(P, periods);
  renderPerfYears(years);
  renderPerfStats(stats);
  renderPerfMonths(months);
}

function renderPerfChart(P, periods) {
  if (!periods.some(r => r.k === perf.preset)) perf.preset = 'all';
  const row = periods.find(r => r.k === perf.preset);
  fill($('#perf-presets'), ...PERF_PRESETS.filter(p => periods.some(r => r.k === p.k)).map(p =>
    h('button', { type: 'button', 'data-action': 'perf-preset', 'data-k': p.k, 'aria-label': p.name, 'aria-pressed': String(perf.preset === p.k) }, p.label)));
  const box = $('#perf-chart'), base = P.pts[row.i], hasSp = base.sp != null;
  const drawn = thinValues(P.pts.slice(row.i), Math.max(24, Math.floor((box.clientWidth || 600) / 6)));
  const you = drawn.map(p => [p.t, p.idx / base.idx - 1]);
  const sp = hasSp ? drawn.map(p => [p.t, p.sp / base.sp - 1]) : [];
  fill($('#perf-legend'), legendItem('value', 'You', pct(row.you)), hasSp ? legendItem('sp', 'S&P 500', pct(row.sp)) : null);
  lineChart(box, {
    height: isWide() ? 280 : 220, x0: base.t, x1: P.t1, xTicks: timeTicks(base.t, P.t1, box.clientWidth || 600),
    series: (hasSp ? [{ cls: 'sp', type: 'line', thin: true, end: true, pts: sp }] : []).concat([{ cls: 'value', type: 'line', end: true, dots: you.length <= 14, pts: you }]),
    yFmt: v => pct(v, { dec: 0 }), samples: drawn.map(p => p.t),
    readout: i => ({ title: fmtDay(drawn[i].t), rows: [{ cls: 'value', value: pct(you[i][1]), label: 'You' }].concat(hasSp ? [{ cls: 'sp', value: pct(sp[i][1]), label: 'S&P 500' }] : []) }),
    focus: i => [{ cls: 'value', x: you[i][0], y: you[i][1] }].concat(hasSp ? [{ cls: 'sp', x: sp[i][0], y: sp[i][1] }] : []),
    label: 'Line chart: your return and the S&P 500’s since ' + fmtDay(base.t)
  });
  $('#perf-fine').textContent = 'Total return since ' + fmtDay(base.t) + '. The S&P 500 includes reinvested dividends. Between your saved values the lines run straight, because what happened in between isn’t known.';
}

function renderPerfTable(P, periods) {
  const label = r => {
    if (r.k === 'ytd') return dayYear(r.t1) + ' so far';
    if (r.k === 'all') return 'All time';
    return r.n === 1 ? 'Past year' : 'Past ' + r.n + ' years';
  };
  const rows = [];
  for (const r of periods) {
    const yearly = r.days >= 365 && r.k !== '1y';
    rows.push(h('tr', null, h('td', null, label(r), yearly ? h('small', { class: 'sub' }, 'a year') : null),
      pctCell(yearly ? r.youAnn : r.you), pctCell(yearly ? r.spAnn : r.sp), diffCell(yearly ? r.youAnn : r.you, yearly ? r.spAnn : r.sp)));
    if (r.k === 'all' && yearly) rows.push(h('tr', null, h('td', null, label(r), h('small', { class: 'sub' }, 'in total')), pctCell(r.you), pctCell(r.sp), diffCell(r.you, r.sp)));
  }
  fill($('#perf-tbody'), ...rows);
}

/* Two bars side by side for each group: yours and the index's. */
function pairBars(box, o) {
  fill(box);
  const W = Math.max(260, Math.floor(box.clientWidth || 600)), H = o.height;
  const vals = [];
  for (const g of o.groups) { if (g.a != null) vals.push(g.a); if (g.b != null) vals.push(g.b); }
  const T = niceTicks(Math.min(0, ...vals), Math.max(0, ...vals) * 1.04, 4);
  const ml = axisPad(T, o.yFmt), mr = 8, mt = 12, mb = 28;
  const pw = W - ml - mr, ph = H - mt - mb;
  const n = o.groups.length, band = pw / Math.max(1, n), bw = Math.max(3, Math.min(16, band * 0.3)), gap = Math.min(3, bw * 0.25);
  const Y = v => mt + ((T.hi - v) / (T.hi - T.lo)) * ph;
  const svg = sv('svg', { width: W, height: H, viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': o.label });
  yAxis(svg, T, o.yFmt, ml, mr, W, Y);
  const y0 = Y(0);
  const useShort = o.groups.some(g => textW(g.label) + 6 > band);
  const lw = Math.max(...o.groups.map(g => textW(useShort ? g.short : g.label))) + 6;
  const every = Math.max(1, Math.ceil(lw / band));
  const els = [];
  o.groups.forEach((g, i) => {
    const cx = ml + band * (i + 0.5), grp = sv('g', { class: 'pair' + (g.faded ? ' faded' : '') });
    const bar = (v, x, cls) => { if (v != null && Math.abs(Y(v) - y0) >= 0.5) grp.append(sv('path', { d: roundedBar(x, bw, y0, Y(v), 3), class: 'bar f-' + cls })); };
    bar(g.a, cx - gap / 2 - bw, 'value');
    bar(g.b, cx + gap / 2, 'sp');
    svg.append(grp); els.push(grp);
    if ((n - 1 - i) % every === 0) svg.append(sv('text', { x: cx, y: H - 8, 'text-anchor': 'middle' }, useShort ? g.short : g.label));
  });
  const tip = h('div', { class: 'tip', hidden: true });
  let curI = -1;
  function show(i) {
    i = Math.max(0, Math.min(n - 1, i)); curI = i;
    els.forEach((el, j) => el.classList.toggle('dim', j !== i));
    const g = o.groups[i];
    fill(tip, h('div', { class: 'tip-title' }, g.title), ...g.rows.map(row => h('div', { class: 'tip-row' }, h('i', { class: 'key key-' + row.cls }), h('b', null, row.value), h('span', null, row.label))));
    tip.hidden = false;
    const cx = ml + band * (i + 0.5), tw = tip.offsetWidth;
    let left = cx + bw + 10;
    if (left + tw > W) left = cx - bw - 10 - tw;
    if (left < 0) left = Math.max(0, Math.min(W - tw, cx - tw / 2));
    tip.style.left = left + 'px';
  }
  function hide() { curI = -1; els.forEach(el => el.classList.remove('dim')); tip.hidden = true; }
  o.groups.forEach((g, i) => {
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

function renderPerfYears(years) {
  const sec = $('#perf-years'), shown = years.filter(r => !r.missing);
  sec.hidden = !shown.length;
  if (!shown.length) return;
  const name = r => r.year + (r.ytd ? ' so far' : r.partial ? ' from ' + dfShort.format(new Date(r.t0 * DAY_MS)) : '');
  fill($('#perf-years-legend'), legendItem('value', 'You'), legendItem('sp', 'S&P 500'));
  pairBars($('#perf-years-chart'), {
    height: isWide() ? 240 : 200, yFmt: v => pct(v, { dec: 0 }),
    groups: shown.map(r => ({
      label: String(r.year), short: '’' + String(r.year).slice(2), a: r.you, b: r.sp, faded: r.ytd || r.partial, title: name(r),
      rows: [{ cls: 'value', value: pct(r.you), label: 'You' }].concat(r.sp != null ? [{ cls: 'sp', value: pct(r.sp), label: 'S&P 500' }] : [])
    })),
    label: 'Bar chart: your return and the S&P 500’s for each calendar year'
  });
  fill($('#perf-years-tbody'), ...years.slice().reverse().map(r => r.missing
    ? h('tr', null, h('td', null, String(r.year)), h('td', { colspan: '3', class: 'gap' }, 'No saved value near a year end'))
    : h('tr', { class: r.ytd ? 'current' : null }, h('td', null, r.ytd ? r.year + ' so far' : String(r.year), r.partial && !r.ytd ? h('small', { class: 'sub' }, 'from ' + dfShort.format(new Date(r.t0 * DAY_MS))) : null),
      pctCell(r.you), pctCell(r.sp), diffCell(r.you, r.sp))));
}

function renderPerfStats(st) {
  const tiles = [];
  const mo = r => fmtMonth(r.y, r.m);
  if (st.n >= 3) {
    tiles.push({ label: 'Best month', value: pct(st.best.you), tone: tone(st.best.you), cap: mo(st.best) + (st.best.sp != null ? '; the S&P 500 did ' + pct(st.best.sp) : '') });
    tiles.push({ label: 'Worst month', value: pct(st.worst.you), tone: tone(st.worst.you), cap: mo(st.worst) + (st.worst.sp != null ? '; the S&P 500 did ' + pct(st.worst.sp) : '') });
  }
  tiles.push(st.dd
    ? { label: 'Biggest drop', value: pct(st.dd.depth), tone: 'bad', cap: fmtDay(st.dd.peakT) + ' to ' + fmtDay(st.dd.troughT) + (st.dd.recT ? ', made back by ' + fmtDay(st.dd.recT) : ', not made back yet') }
    : { label: 'Biggest drop', value: 'None', cap: 'No fall between your saved values so far' });
  if (st.n >= 3) tiles.push({ label: 'Months that went up', value: st.up + ' of ' + st.n, cap: st.cmp ? 'Ahead of the S&P 500 in ' + st.ahead + ' of ' + st.cmp : 'Full months with a value at each end' });
  else tiles.push(st.high.atHigh
    ? { label: 'Against its high', value: 'At the high', tone: 'good', cap: 'Your latest value is the best so far' }
    : { label: 'Against its high', value: pct(st.high.below), tone: 'bad', cap: 'Below the high of ' + fmtDay(st.high.t) });
  fill($('#perf-stats'), ...kpiTiles(tiles));
  fitTiles($('#perf-stats'));
}

/* A year-by-month table of returns, each cell shaded by how far up or down the month went. */
function renderPerfMonths(months) {
  const box = $('#perf-heat'), read = $('#perf-heat-read'), fine = $('#perf-heat-fine');
  if (months.length < 3) {
    fill(box);
    fill(read, 'Monthly returns appear once you have values saved near the ends of three or more months. ',
      mode === 'example' ? null : h('button', { type: 'button', class: 'link', 'data-action': 'open-import', 'data-target': 'values' }, 'Add many values'));
    fine.textContent = 'A broker’s export of daily or monthly values fills this in at once.';
    return;
  }
  const by = new Map(months.map(r => [r.y * 12 + r.m, r]));
  const y1 = months[months.length - 1].y, y0 = months[0].y;
  const wide = (box.parentElement.clientWidth || 320) >= 520;
  const say = r => {
    fill(read, B(fmtMonth(r.y, r.m) + (r.mtd ? ' so far' : '')), ': you ' + pct(r.you) + (r.sp != null ? ', S&P 500 ' + pct(r.sp) : '') + '.');
  };
  const head = h('div', { class: 'heat-row heat-head', 'aria-hidden': 'true' }, h('span'));
  for (let m = 1; m <= 12; m++) head.append(h('span', null, dfMonthNarrow.format(new Date(Date.UTC(2001, m - 1, 15)))));
  const rows = [head];
  for (let y = y1; y >= y0; y--) {
    const row = h('div', { class: 'heat-row' }, h('span', { class: 'heat-year', 'aria-hidden': 'true' }, String(y)));
    for (let m = 1; m <= 12; m++) {
      const r = by.get(y * 12 + m);
      if (!r) { row.append(h('span', { class: 'heat-cell none' })); continue; }
      const p = r.you * 100, strength = Math.min(1, Math.abs(p) / 8);
      const txt = (wide && Math.abs(p) < 9.95 ? num(Math.abs(p), 1) : String(Math.round(Math.abs(p))));
      const cell = h('button', { type: 'button', class: 'heat-cell ' + (p >= 0 ? 'up' : 'dn') + (r.mtd ? ' mtd' : ''),
        'aria-label': fmtMonth(r.y, r.m) + (r.mtd ? ' so far' : '') + ': ' + pct(r.you), style: { '--a': (0.1 + 0.5 * strength).toFixed(3) } },
        (p < 0 && Math.round(Math.abs(p) * (wide ? 10 : 1)) !== 0 ? MINUS : '') + txt);
      cell.addEventListener('click', () => say(r));
      cell.addEventListener('pointerenter', e => { if (e.pointerType === 'mouse') say(r); });
      cell.addEventListener('focus', () => say(r));
      row.append(cell);
    }
    rows.push(row);
  }
  fill(box, ...rows);
  say(months[months.length - 1]);
  fine.textContent = 'Each cell is your return for that month, in percent. Tap one to compare it with the S&P 500. A month needs a saved value near both of its ends.';
}

/* ---------- Plan: a goal, and a look ahead ---------- */
/* The plan's settings travel with your file (Settings sheet). */
function syncPlanFromState() {
  const pl = state.settings.plan;
  Object.assign(proj, { years: 10, annual: null, annualTouched: false, rate: 0.07, preset: '7' });
  if (!pl) return;
  proj.years = pl.years;
  proj.rate = pl.rate;
  proj.preset = { 0.1: '10', 0.07: '7', 0.04: '4' }[pl.rate] || '';
  if (pl.annual != null) { proj.annual = pl.annual; proj.annualTouched = true; }
}
function storePlan() {
  if (mode !== 'data') return;
  state.settings.plan = { years: proj.years, annual: proj.annualTouched ? proj.annual : null, rate: proj.rate };
  writeCopy();
}
const GOAL_STEPS = [5e3, 1e4, 2.5e4, 5e4, 7.5e4, 1e5, 1.5e5, 2e5, 2.5e5, 3e5, 4e5, 5e5, 7.5e5, 1e6, 1.5e6, 2e6, 3e6, 5e6, 7.5e6, 1e7, 2e7, 5e7, 1e8];
function goalPicks(v0) {
  const out = GOAL_STEPS.filter(g => g > v0 * 1.05).slice(0, 4);
  return out.length ? out : [Math.ceil((v0 * 1.5) / 1e7) * 1e7];
}
function ringSvg(frac, label) {
  const R = 46, C = 2 * Math.PI * R;
  const svg = sv('svg', { viewBox: '0 0 112 112', width: 112, height: 112, role: 'img', 'aria-label': label ? label + ' of the way to your goal' : 'No goal set' },
    sv('circle', { cx: 56, cy: 56, r: R, class: 'ring-track' }));
  if (frac > 0) svg.append(sv('circle', { cx: 56, cy: 56, r: R, class: 'ring-fill' + (frac >= 1 ? ' done' : ''), 'stroke-dasharray': (Math.max(0.005, frac) * C).toFixed(1) + ' ' + C.toFixed(1), transform: 'rotate(-90 56 56)' }));
  svg.append(sv('text', { x: 56, y: 56, 'text-anchor': 'middle', 'dominant-baseline': 'central', class: 'ring-text' }, label || '—'));
  return svg;
}

function renderPlan(M) {
  const has = mode !== 'welcome';
  $('#plan-empty').hidden = has;
  $('#plan-body').hidden = !has;
  if (!has) {
    fill($('#plan-empty'), h('h3', null, 'No file open'), h('p', null, 'Open your Excel file to set a goal and look ahead.'),
      h('div', { class: 'btn-row' }, h('button', { class: 'btn btn-primary', type: 'button', 'data-action': 'open-file' }, 'Open Excel file')));
    return;
  }
  renderGoal(M);
  renderProjection(M);
}

function renderGoal(M) {
  const goal = state.settings.goal, v0 = M.hasValue ? M.V : Math.max(0, M.total);
  const inp = $('#goal-amount');
  if (document.activeElement !== inp) inp.value = goal ? plain(goal) : '';
  fill($('#goal-presets'),
    ...goalPicks(v0).map(g => h('button', { type: 'button', 'data-action': 'goal-preset', 'data-v': String(g), 'aria-pressed': String(goal === g) }, money(g, { dec: 0, compact: g >= 1e6 }))),
    goal ? h('button', { type: 'button', 'data-action': 'clear-goal' }, 'No goal') : null);
  const ring = $('#goal-ring'), text = $('#goal-text'), paceEl = $('#goal-pace');
  fill(paceEl);
  if (!goal) {
    fill(ring, ringSvg(0, null));
    fill(text, h('p', { class: 'goal-lead' }, 'Set a target'), h('p', { class: 'note' }, 'Type an amount or pick one, and the app shows how far along you are and when you might get there.'));
    return;
  }
  fill(ring, ringSvg(Math.max(0, Math.min(1, v0 / goal)), pct(v0 / goal, { sign: false, dec: 0 })));
  const figEl = h('span', { class: 'goal-fig' }, money(v0, { dec: 0 }));
  const parts = [h('p', { class: 'goal-lead' }, figEl, h('small', null, 'of ' + money(goal, { dec: 0 })))];
  if (v0 >= goal) parts.push(h('p', { class: 'note' }, 'You’ve reached this goal. Set the next one whenever you like.'));
  else {
    const annual = planAnnual(M);
    const m = monthsToGoal(v0, annual, proj.rate, goal);
    const slow = monthsToGoal(v0, annual, Math.max(0, proj.rate - 0.03), goal), fast = monthsToGoal(v0, annual, proj.rate + 0.03, goal);
    const pace = 'At ' + money(annual, { dec: 0 }) + ' a year in deposits and a ' + pct(proj.rate, { sign: false, dec: (proj.rate * 1000) % 10 ? 1 : 0 }) + ' yearly return';
    parts.push(h('p', { class: 'goal-left' }, money(goal - v0, { dec: 0 }) + ' to go'));
    if (m == null) paceEl.append(pace + ', this goal isn’t reached within 100 years. Try a higher deposit or return under Looking ahead.');
    else paceEl.append(pace + ', you’d get there around ', B(monthAhead(m)),
      fast != null && slow != null && fast !== slow ? ' (between ' + monthAhead(fast) + ' and ' + monthAhead(slow) + ' if returns run three points higher or lower).' : '.');
  }
  fill(text, ...parts);
  fitFigure(figEl, 17, text);
}

function renderProjection(M) {
  const v0 = M.hasValue ? M.V : Math.max(0, M.total);
  if (!proj.annualTouched) proj.annual = planAnnual(M);
  const ai = $('#p-annual');
  if (document.activeElement !== ai) ai.value = plain(proj.annual);
  $('#p-years').value = proj.years;
  $('#p-years-out').textContent = plural(proj.years, 'year', 'years');
  $('#p-rate').value = (proj.rate * 100).toFixed(1);
  $('#p-rate-out').textContent = pct(proj.rate, { sign: false, dec: (proj.rate * 1000) % 10 ? 1 : 0 });
  const presets = [];
  if (M.xirr != null && M.xirr > -0.5 && M.xirr < 0.5) presets.push({ k: 'xirr', label: 'Your XIRR ' + pct(M.xirr, { sign: false }), r: M.xirr });
  presets.push({ k: '10', label: 'S&P long-run 10%', r: 0.10 }, { k: '7', label: 'Middle 7%', r: 0.07 }, { k: '4', label: 'Cautious 4%', r: 0.04 });
  fill($('#p-presets'), ...presets.map(p => h('button', { type: 'button', 'data-action': 'proj-preset', 'data-rate': String(p.r), 'data-k': p.k, 'aria-pressed': String(proj.preset === p.k) }, p.label)));
  const pts = projectSeries(v0, proj.annual, proj.rate, proj.years);
  const lo = projectSeries(v0, proj.annual, Math.max(0, proj.rate - 0.03), proj.years), hi = projectSeries(v0, proj.annual, proj.rate + 0.03, proj.years);
  const end = pts[pts.length - 1], endLo = lo[lo.length - 1].v, endHi = hi[hi.length - 1].v;
  const added = proj.annual * proj.years, growth = end.v - v0 - added;
  fill($('#proj-result'), 'In ' + (M.cy + proj.years) + ' you could have ', h('strong', { class: 'fig' }, money(end.v, { dec: 0 })));
  $('#proj-break').textContent = money(v0, { dec: 0 }) + (M.hasValue ? ' today' : ' deposited so far') + ' + ' + money(added, { dec: 0 }) + ' in new deposits + ' + money(growth, { dec: 0 }) + ' of growth. With returns three points lower or higher: ' + money(endLo, { dec: 0 }) + ' to ' + money(endHi, { dec: 0 }) + '.';
  const goal = state.settings.goal;
  const showGoal = !!goal && goal > v0 && goal <= endHi * 1.5;
  fill($('#proj-legend'), legendItem('value', 'Projected value', money(end.v, { dec: 0 })), legendItem('dep', 'Without growth', money(end.base, { dec: 0 }), true),
    showGoal ? h('li', null, h('i', { class: 'key dash' }), h('span', null, 'Goal'), h('b', null, money(goal, { dec: 0 }))) : null);
  const box = $('#proj-chart');
  const maxTicks = Math.max(2, Math.floor(((box.clientWidth || 600) - 60) / 52));
  const step = [1, 2, 5, 10].find(st => Math.floor(proj.years / st) + 1 <= maxTicks) || 10;
  const ticks = [];
  for (let y = 0; y <= proj.years; y += step) ticks.push({ v: y, label: String(M.cy + y) });
  lineChart(box, {
    height: isWide() ? 260 : 210, x0: 0, x1: proj.years, xTicks: ticks,
    series: [{ cls: 'dep', type: 'line', area: true, pts: pts.map(p => [p.y, p.base]) },
      { cls: 'value', type: 'band', pts: pts.map((p, i) => [p.y, lo[i].v, hi[i].v]) },
      { cls: 'value', type: 'line', end: true, pts: pts.map(p => [p.y, p.v]) }],
    refs: showGoal ? [{ value: goal, inRange: true }] : [],
    yFmt: yMoney, samples: pts.map(p => p.y),
    readout: i => ({ title: String(M.cy + pts[i].y), rows: [{ cls: 'value', value: money(pts[i].v, { dec: 0 }), label: 'Projected value' }, { cls: 'value', value: money(lo[i].v, { dec: 0, compact: true }) + ' to ' + money(hi[i].v, { dec: 0, compact: true }), label: 'Range' }, { cls: 'dep', value: money(pts[i].base, { dec: 0 }), label: 'Without growth' }] }),
    focus: i => [{ cls: 'dep', x: pts[i].y, y: pts[i].base }, { cls: 'value', x: pts[i].y, y: pts[i].v }],
    label: 'Line chart: projected portfolio value by year with a range around it, compared with the same deposits without growth'
  });
}

function renderLedger(M) {
  const list = $('#ledger-list');
  $('#dep-sub').textContent = M.count ? plural(M.count, 'entry', 'entries') + ', ' + money(M.total, { dec: 0 }) + ' in total' : 'Every deposit and withdrawal, newest first.';
  if (!M.deps.length) {
    fill(list, h('p', { class: 'empty' }, mode === 'welcome' ? 'Open your Excel file to see your deposits here, or add your first one above to start a new tracker.' : 'No deposits yet. Add one above, or bring in many at once from Excel.'));
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
  fill(list, frag);
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
/* The dollar, euro and pound first, then the rest in alphabetical order. A file in any other currency keeps it. */
const CURRENCIES = ['USD', 'EUR', 'GBP', 'AED', 'ARS', 'AUD', 'BHD', 'BRL', 'CAD', 'CHF', 'CLP', 'CNY', 'COP', 'CZK', 'DKK', 'EGP', 'HKD', 'HUF', 'IDR', 'ILS', 'INR', 'JOD', 'JPY', 'KES', 'KRW', 'KWD', 'LBP', 'MAD', 'MXN', 'MYR', 'NGN', 'NOK', 'NZD', 'OMR', 'PHP', 'PKR', 'PLN', 'QAR', 'RON', 'SAR', 'SEK', 'SGD', 'THB', 'TRY', 'TWD', 'VND', 'ZAR'];
function renderSettings() {
  const sel = $('#set-cur');
  if (!sel.options.length) {
    let names = null;
    try { names = new Intl.DisplayNames([LOCALE], { type: 'currency' }); } catch (e) { names = null; }
    for (const c of CURRENCIES) sel.append(h('option', { value: c }, c + (names ? ', ' + names.of(c) : '')));
  }
  if (!CURRENCIES.includes(cur()) && !Array.from(sel.options).some(o => o.value === cur())) sel.append(h('option', { value: cur() }, cur()));
  sel.value = cur();
  for (const id of ['set-keep', 'wel-keep']) { const el = $('#' + id); el.checked = keepCopy; el.disabled = !storageOK; }
  $('#wel-keep-wrap').hidden = !storageOK;
  $('#set-keep-note').textContent = !storageOK
    ? 'This browser blocks storage, so nothing can be kept here. Your Excel file has everything.'
    : keepCopy
      ? 'On: the app keeps a copy of your numbers on this device so it opens without picking the file. The copy is never uploaded. Turn it off to erase it.'
      : 'Off: nothing is stored on this device. The app forgets your numbers when it closes, and your Excel file is the only copy.';
  $('#set-file').textContent = mode === 'welcome' ? 'No file open.' : mode === 'example' ? 'Showing example numbers.' : (file.appFormat && file.name ? file.name : 'A new tracker, not saved to a file yet') + (file.dirty ? ', with unsaved changes' : '');
  $('#set-close').hidden = mode === 'welcome';
  $('#set-save').disabled = mode !== 'data';
  fill($('#set-spx'), 'S&P 500 total return (dividends reinvested) from 1985 to the close of ' + fmtDate(SPX.last.d) + ' at ' + num(SPX.last.px, 2) + '. ' +
    (PWA ? 'A daily job on the site this app is served from adds each new close, and the app picks it up whenever it opens online. Your phone never contacts a price service itself. '
         : 'For later dates, use “Add a newer close” on the Performance tab. '),
    spxCanCheck() ? h('button', { type: 'button', class: 'link', 'data-action': 'check-spx' }, spxBusy ? 'Checking…' : 'Check for new prices') : null);
  const ver = document.querySelector('meta[name="app-version"]'), verEl = document.getElementById('set-version');
  if (verEl) verEl.textContent = ver ? 'App version ' + ver.content.replace(/^build /, '') : '';
}

/* ---------- opening and saving Excel files ---------- */
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const readOpts = () => ({ localeFmt: LOCALE_DATE_FMT, localeDecimalComma: LOCALE_DECIMAL_COMMA, want: imp.target, today: todayIso() });
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
    resetView(); syncPlanFromState();
    writeCopy();
    setTab('overview');
    toast('Opened ' + got.name + ': ' + plural(r.deposits.length, 'deposit', 'deposits') + (r.skipped ? '. ' + plural(r.skipped, 'row', 'rows') + ' couldn’t be read and ' + (r.skipped === 1 ? 'was' : 'were') + ' left out' : ''), prev, r.skipped ? 7000 : 4500);
    return;
  }
  /* Someone's own spreadsheet or a CSV: show what was found before using it. */
  openImportForFile(got);
}

function excelSerial(iso) { return isoToDay(iso) + 25569; }
function buildWorkbook(st, M, template) {
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
    [S(template
      ? 'Type your deposits on the Deposits sheet (date, amount) and what your portfolio was worth on the Values sheet (date, value), then open this file in the Investment Tracker app. The app fills in the rest of this sheet when you save.'
      : 'The figures in this column update by themselves in Excel. The Investment Tracker app reads the Deposits and Values sheets.', 'note')],
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
    [S(template ? 'Filled in by the app when you save' : 'From the app when you saved on ' + fmtDate(today), 'header')],
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
    [S('Goal'), st.settings.goal ? S(st.settings.goal, 'money') : ''],
    [S('Plan: years ahead'), st.settings.plan ? st.settings.plan.years : ''],
    [S('Plan: deposits per year'), st.settings.plan && st.settings.plan.annual != null ? S(st.settings.plan.annual, 'money') : ''],
    [S('Plan: return per year'), st.settings.plan ? S(st.settings.plan.rate, 'percent') : ''],
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
/* An empty workbook in the app's own layout, for starting a file by hand in Excel or Numbers. */
function downloadTemplate() {
  const st = emptyState();
  st.settings.cur = cur();
  let bytes;
  try { bytes = XLSXLite.write(buildWorkbook(st, computeMetrics(st, getAnchors(), todayDay()), true)); } catch (e) { toast('The template couldn’t be built.'); return; }
  deliver(new Blob([bytes], { type: XLSX_MIME }), 'Investments template.xlsx', true);
}
/* Hands the file to you: the host app's save prompt where there is one, the iPhone share sheet, or a normal download. */
async function deliver(blob, name, template) {
  if (downloadsCap) {
    try { await downloadsCap.save({ filename: name, data: blob }); toast((template ? 'Template saved as ' : 'Saved ') + name + '.'); return true; }
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
        toast(template ? 'Template ready. Fill it in with Excel or Numbers, then open it here.' : 'Saved. If Files asks, choose Replace to update your old copy.', undefined, 5000);
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
  toast('Downloaded ' + name + '.' + (template ? ' Fill it in, then open it here.' : ''));
  return true;
}

/* ---------- import dialog: add many deposits, or read someone's own workbook ---------- */
/* target: 'deposits' (the ledger) or 'values' (your portfolio's value on past dates) */
/* keep: for a long run of daily values, keep the last of each 'month', each 'week', or 'all' of them */
const imp = { mode: 'add', target: 'deposits', tab: 'paste', fmt: 'auto', dateCol: null, amountCol: null, delim: null, src: null, sheetIdx: 0, res: null, skipDup: true, toAdd: [], openName: '', guess: null, keep: 'month', plan: null };
const IMP_TEXT = {
  deposits: { title: 'Add many deposits', label: 'In Excel, select your date and amount columns, copy them, then paste here.', placeholder: 'Date\tAmount\n15/01/2024\t500\n15/02/2024\t500\n20/03/2024\t1,250',
    drop: 'Excel (.xlsx) or CSV. The deposits in it are added to your ledger.', replace: 'Replace my existing deposits' },
  values: { title: 'Add many values', label: 'Paste a date column and your portfolio’s value on each date from Excel. A broker’s CSV export of daily or monthly values can be pasted whole.', placeholder: 'Date\tValue\n31/01/2024\t10,250\n29/02/2024\t10,610\n31/03/2024\t11,020',
    drop: 'Excel (.xlsx) or CSV with a date column and a value column, or a broker export. The values are added to your value history.', replace: 'Replace my saved values' }
};
function importTabs() { return imp.target === 'values' ? ['paste', 'file'] : ['paste', 'file', 'recurring']; }
const dlgImport = $('#dlg-import');
function openDialog(d) { if (typeof d.showModal === 'function') { if (!d.open) d.showModal(); } else d.setAttribute('open', ''); }
function closeDialog(d) { if (typeof d.close === 'function') { if (d.open) d.close(); } else d.removeAttribute('open'); }
function resetImport() { Object.assign(imp, { fmt: 'auto', dateCol: null, amountCol: null, delim: null, res: null, toAdd: [], sheetIdx: 0, keep: 'month', plan: null }); }
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
  return { fmt: imp.fmt, localeFmt: LOCALE_DATE_FMT, localeDecimalComma: LOCALE_DECIMAL_COMMA, dateCol: imp.dateCol, amountCol: imp.amountCol, delim: imp.delim, want: imp.target, today: todayIso() };
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
/* What was read from a broker export, so the dates are never a surprise. */
function brokerNote(b) {
  const kinds = { month: 'Each month', quarter: 'Each quarter' };
  let t = 'Read from the “' + b.section + '” section of your broker export.';
  if (b.kind && kinds[b.kind]) {
    t += ' ' + kinds[b.kind] + ' is dated its last day';
    t += b.capped ? '; the ' + b.kind + ' that isn’t over yet is dated ' + (b.cap === todayIso() ? 'today' : fmtDate(b.cap) + ', when the report ends') + '.' : '.';
  }
  return h('p', { class: 'note' }, t);
}
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
  fill(pv);
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
    const msg = { empty: 'Nothing to read yet.', nodate: 'I couldn’t find a column of dates' + (imp.mode === 'open' ? ' on this sheet. Pick another sheet above, if there is one.' : '. Check that the date column came along when you copied.'), noamount: 'I found dates but no column of amounts next to them.', broker: 'This looks like a broker export, but it has no deposits section to read. To add your portfolio values from it, use Add many values on the Overview tab.' }[r.problem];
    pv.append(h('p', { class: 'field-error' }, msg));
  }
  const existing = new Set(replace || mode !== 'data' ? [] : state.deposits.map(d => d.d + '|' + d.a.toFixed(2)));
  let dups = 0;
  const items = r.items.map(it => { const dup = existing.has(it.d + '|' + it.a.toFixed(2)); if (dup) dups++; return Object.assign({ dup }, it); });
  const toAdd = items.filter(it => !(it.dup && imp.skipDup));
  imp.toAdd = toAdd;
  if (r.broker && !r.problem) pv.append(brokerNote(r.broker));
  if (!r.recurring && r.ncol > 1 && !r.problem) pv.append(...columnOptions(r, 'Amounts from'));
  if (items.length || r.errors.length) {
    const sumTxt = [h('strong', null, plural(toAdd.length, 'entry', 'entries')), ' ready' + (imp.mode === 'open' ? '' : ' to add')];
    if (toAdd.length) sumTxt.push(', totaling ' + money(sumOf(toAdd, x => x.a), { dec: 0 }));
    sumTxt.push('.');
    const extra = [];
    if (dups) extra.push(plural(dups, 'row matches', 'rows match') + ' an entry already in your ledger');
    if (r.errors.length) extra.push(plural(r.errors.length, 'row', 'rows') + ' couldn’t be read');
    if (r.header && !r.broker) extra.push('the header row was skipped');
    if (r.skipped && !r.broker) extra.push(plural(r.skipped, 'row', 'rows') + ' without a date or amount ' + (r.skipped === 1 ? 'was' : 'were') + ' skipped');
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
/* Values: one per date. A date that's already saved gets the new value; the latest date becomes your current value.
   A long run of daily values (a broker export) is thinned to one per month or week, so the chart stays readable. */
const savedValues = s => {
  const out = s.history.slice();
  if (s.value && !out.some(x => x.d === s.value.d)) out.push({ d: s.value.d, a: s.value.a });
  return out;
};
const KEEP = [
  { k: 'month', label: 'The last value of each month' },
  { k: 'week', label: 'The last value of each week' },
  { k: 'all', label: 'Every value' }
];
function renderValuesPreview() {
  const pv = $('#imp-preview'), go = $('#imp-go');
  const replaceWrap = $('#imp-replace-wrap');
  replaceWrap.hidden = mode !== 'data' || !state.history.length;
  const replace = !replaceWrap.hidden && $('#imp-replace').checked;
  fill(pv);
  imp.toAdd = [];
  imp.plan = null;
  go.disabled = true;
  go.textContent = 'Add values';
  const r = imp.res;
  if (!r) return;
  if (r.problem) {
    pv.append(h('p', { class: 'field-error' }, { empty: 'Nothing to read yet.', nodate: 'I couldn’t find a column of dates. Check that the date column came along when you copied.', noamount: 'I found dates but no column of values next to them.',
      broker: 'This looks like a broker export, but I couldn’t find a section with a date and a NAV or total value column in it' + (r.broker && r.broker.sections.length ? ' (it has ' + r.broker.sections.slice(0, 6).map(s => '“' + s + '”').join(', ') + ')' : '') + '.' }[r.problem]));
    return;
  }
  if (r.broker) pv.append(brokerNote(r.broker));
  if (r.ncol > 1) pv.append(...columnOptions(r, 'Values from'));
  const saved = replace || mode !== 'data' ? [] : savedValues(state);
  const plan = planValues(r.items, saved, imp.keep);
  imp.plan = plan;
  if (plan.dense) {
    const ksel = h('select', { id: 'imp-keep' }, ...KEEP.map(o => h('option', { value: o.k }, o.label)));
    ksel.value = imp.keep;
    ksel.addEventListener('change', () => { imp.keep = ksel.value; renderValuesPreview(); });
    pv.append(h('label', { class: 'field', for: 'imp-keep' }, h('span', { class: 'field-label' }, 'This has a value for nearly every day. Keep'), ksel));
  }
  const rows = plan.kept.map(v => ({ line: v.line, d: fmtDate(v.d), a: money(v.a, {}),
    st: v.st === 'same' ? 'Already saved' : v.st === 'repl' ? 'Replaces ' + money(v.old, { dec: 0 }) : '', cls: v.st === 'same' ? 'dup' : v.st === 'repl' ? 'repl' : '' }));
  for (const it of plan.neg) rows.push({ line: it.line, d: fmtDate(it.d), a: money(it.a, {}), st: 'A value can’t be negative', cls: 'err' });
  for (const e of r.errors) rows.push({ line: e.line, d: String(e.raw[0] == null ? '' : e.raw[0]), a: String(e.raw[1] == null ? '' : e.raw[1]), st: e.reason === 'date' ? 'Couldn’t read the date' : 'Couldn’t read the value', cls: 'err' });
  rows.sort((x, y) => (x.line || 0) - (y.line || 0));
  const changes = plan.nNew + plan.nRepl;
  imp.toAdd = plan.kept.filter(v => v.st !== 'same');
  if (rows.length) {
    const lead = [];
    if (plan.dense && imp.keep !== 'all') {
      lead.push('Found ' + plural(plan.found, 'daily value', 'daily values') + ' from ' + fmtDate(plan.range[0]) + ' to ' + fmtDate(plan.range[1]) + '. Keeping ',
        h('strong', null, plural(plan.kept.length, 'value', 'values')), ': the last of each ' + (imp.keep === 'week' ? 'week' : 'month') + ', plus the first and the latest.');
    } else lead.push(h('strong', null, plural(plan.kept.length, 'value', 'values')), ' found.');
    const extra = [];
    if (saved.length) {
      if (plan.nNew) extra.push(plural(plan.nNew, 'is', 'are') + ' new');
      if (plan.nRepl) extra.push(plural(plan.nRepl, 'replaces a value', 'replace values') + ' you saved for that date');
      if (plan.nSame) extra.push(plural(plan.nSame, 'is', 'are') + ' already saved');
      if (plan.removed.length) extra.push(plural(plan.removed.length, 'other saved value', 'other saved values') + ' in that period ' + (plan.removed.length === 1 ? 'is' : 'are') + ' replaced too');
    }
    if (plan.twice.length) extra.push(plural(plan.twice.length, 'date appears', 'dates appear') + ' twice, so the lower row is used');
    if (plan.neg.length) extra.push(plural(plan.neg.length, 'row has', 'rows have') + ' a negative value');
    if (r.errors.length) extra.push(plural(r.errors.length, 'row', 'rows') + ' couldn’t be read');
    if (r.header && !r.broker) extra.push('the header row was skipped');
    if (r.skipped && !r.broker) extra.push(plural(r.skipped, 'row', 'rows') + ' without a date or value ' + (r.skipped === 1 ? 'was' : 'were') + ' skipped');
    pv.append(h('p', { class: 'pv-summary' }, ...lead, extra.length ? ' ' + extra.join('; ').replace(/^./, c => c.toUpperCase()) + '.' : ''));
    const newest = plan.kept.length ? plan.kept[plan.kept.length - 1] : null;
    const curD = saved.length && state.value ? state.value.d : '';
    if (newest && newest.st !== 'same' && newest.d >= curD) pv.append(h('p', { class: 'note' }, 'Your current value becomes ' + money(newest.a, { dec: 0 }) + ', from ' + fmtDate(newest.d) + ', the latest date here.'));
    const shown = rows.slice(0, 160);
    pv.append(h('div', { class: 'table-scroll' }, h('table', { class: 'num-table pv-table' },
      h('thead', null, h('tr', null, h('th', { scope: 'col' }, 'Date'), h('th', { scope: 'col' }, 'Value'), h('th', { scope: 'col' }, ''))),
      h('tbody', null, shown.map(x => h('tr', null, h('td', null, x.d), h('td', null, x.a), h('td', { class: 'st ' + x.cls }, x.st)))))));
    if (rows.length > shown.length) pv.append(h('p', { class: 'pv-more' }, 'Showing the first ' + shown.length + ' of ' + rows.length + ' rows.'));
  }
  const n = changes || (plan.removed.length ? plan.kept.length : 0);
  go.disabled = !n;
  if (n) go.textContent = (replace ? 'Replace with ' : changes ? 'Add ' : 'Save ') + plural(replace ? plan.kept.length : n, 'value', 'values');
  else if (plan.kept.length) go.textContent = 'All already saved';
}
async function runValuesImport() {
  const replace = !$('#imp-replace-wrap').hidden && $('#imp-replace').checked;
  const r = imp.res, keep = imp.keep;
  if (!r || !imp.plan) return;
  const count = replace ? imp.plan.kept.length : imp.plan.nNew + imp.plan.nRepl;
  if (replace) {
    closeDialog(dlgImport);
    const ok = await confirmBox('Replace your saved values?', 'Your ' + plural(state.history.length, 'saved value', 'saved values') + ' will be removed and replaced with ' + plural(imp.plan.kept.length, 'value', 'values') + ' from this import.', 'Replace', true);
    if (!ok) { openDialog(dlgImport); return; }
  }
  const prev = commitData(s => {
    s.history = planValues(r.items, replace ? [] : savedValues(s), keep).history.slice(-20000);
    const last = s.history[s.history.length - 1];
    s.value = last ? { a: last.a, d: last.d } : null;
  });
  closeDialog(dlgImport);
  $('#imp-text').value = '';
  $('#imp-file').value = '';
  imp.src = null;
  $('#imp-file-status').textContent = '';
  toast(count ? (replace ? 'Replaced with ' : 'Added ') + plural(count, 'value', 'values') : 'Value history tidied', prev);
}
async function runImport() {
  if (imp.target === 'values') { await runValuesImport(); return; }
  const items = imp.toAdd || [];
  if (!items.length) return;
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
    resetView(); syncPlanFromState();
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
const noFile = () => Object.assign(file, { name: null, appFormat: false, dirty: false, savedAt: null, openedAt: null, source: null });
/* Back to a clean view whenever a different set of numbers is loaded. */
function resetView() { editing = null; valueOpen = false; histOpen = false; histAll = false; spxOpen = false; openYears.clear(); resetZoom(); perf.preset = 'all'; }
function setGoal(v) {
  const g = v && isFinite(v) && v > 0 ? round2(v) : null;
  if (mode === 'welcome') return;
  if (g === (state.settings.goal || null)) { renderPlan(lastM); return; }
  commitSettings(s => { s.settings.goal = g; });
  toast(g ? 'Goal set: ' + money(g, { dec: 0 }) : 'Goal removed');
}
function setKeep(on) {
  keepCopy = !!on && storageOK;
  if (keepCopy) writeCopy(); else dropCopy();
  render();
  toast(keepCopy ? 'A copy is now kept on this device.' : 'The copy on this device was erased.');
}

const actions = {
  'open-file': () => { closeDialog($('#dlg-settings')); const f = $('#open-file'); f.value = ''; f.click(); },
  'save': () => { closeDialog($('#dlg-settings')); saveToExcel(); },
  'template': () => { closeDialog($('#dlg-settings')); downloadTemplate(); },
  'settings': () => { renderSettings(); openDialog($('#dlg-settings')); },
  'toggle-hide': () => { hideAmounts = !hideAmounts; writeCopy(); render(); toast(hideAmounts ? 'Amounts hidden' : 'Amounts shown'); },
  'new-tracker': () => {
    const prev = snapshot();
    state = emptyState(); mode = 'data';
    noFile(); resetView(); syncPlanFromState();
    writeCopy();
    setTab('deposits');
    toast('New tracker started. Add your deposits, then save to Excel.', prev);
    setTimeout(() => $('#add-amount').focus(), 60);
  },
  'show-example': () => { state = makeExample(todayDay()); mode = 'example'; noFile(); resetView(); syncPlanFromState(); setTab('overview'); },
  'close-file': async () => {
    if (mode === 'data' && file.dirty) {
      closeDialog($('#dlg-settings'));
      const ok = await confirmBox('Close without saving?', 'Changes that aren’t in your Excel file will be lost.', 'Close file', true);
      if (!ok) return;
    }
    const keepCur = state.settings.cur;
    state = emptyState(); state.settings.cur = keepCur; mode = 'welcome';
    noFile(); resetView(); syncPlanFromState();
    writeCopy();
    closeDialog($('#dlg-settings'));
    setTab('overview');
  },
  'erase-all': async () => {
    closeDialog($('#dlg-settings'));
    const ok = await confirmBox('Erase everything from this app?', 'This clears the numbers in the app' + (keepCopy ? ' and the copy kept on this device' : '') + '. Your Excel file isn’t touched.', 'Erase', true);
    if (!ok) return;
    dropCopy(); keepCopy = false; hideAmounts = false;
    state = emptyState(); mode = 'welcome';
    noFile(); resetView(); syncPlanFromState();
    setTab('overview');
    toast('Erased from this app.');
  },
  'open-import': btn => openImport(btn.dataset.tab, btn.dataset.target),
  'zoom-preset': btn => { resetZoom(); zoom.preset = btn.dataset.k; drawGrowth(); },
  'perf-preset': btn => { perf.preset = btn.dataset.k; if (lastP) renderPerfChart(lastP, perfPeriods(lastP)); },
  'go-deposits': () => setTab('deposits'),
  'open-value': () => {
    if (mode === 'welcome' || mode === 'example') return;
    valueOpen = true;
    const el = $('#val-amount');
    el.value = state.value ? plain(state.value.a) : '';
    $('#val-date').value = todayIso();
    $('#val-error').textContent = '';
    renderValueForm(lastM);
    el.focus();
    el.select();
  },
  'close-value': () => { valueOpen = false; renderValueForm(lastM); },
  'toggle-history': () => { histOpen = !histOpen; histAll = false; renderValueForm(lastM); },
  'hist-all': () => { histAll = true; renderValueForm(lastM); },
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
  'check-spx': () => checkSpx(true),
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
  'proj-preset': btn => { proj.rate = +btn.dataset.rate; proj.preset = btn.dataset.k; storePlan(); renderPlan(lastM); },
  'goal-preset': btn => setGoal(+btn.dataset.v),
  'clear-goal': () => setGoal(null),
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
  b.addEventListener('click', () => {
    /* tapping the tab you're on goes back to the top, as iPhone apps do */
    if (b.dataset.tab === tab) window.scrollTo({ top: 0, behavior: REDUCED ? 'auto' : 'smooth' });
    else setTab(b.dataset.tab);
  });
  b.addEventListener('keydown', e => {
    if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
      e.preventDefault();
      const next = TABS[(TABS.indexOf(tab) + (e.key === 'ArrowRight' ? 1 : TABS.length - 1)) % TABS.length];
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
  $('#entry-title').textContent = k === 'dep' ? 'Add a deposit' : 'Add a withdrawal';
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

/* Your current value is always the latest saved one; an older date is kept as history. */
function saveValue() {
  const raw = $('#val-amount').value.trim();
  const iso = $('#val-date').value || todayIso();
  const err = $('#val-error');
  if (!raw) { err.textContent = 'Enter what your portfolio is worth, like 25,000.'; return; }
  const a = parseInputAmount(raw);
  if (!isFinite(a) || a < 0) { err.textContent = 'Enter the value as a number, like 25,000.'; return; }
  if (!isFinite(isoToDay(iso))) { err.textContent = 'Choose the date of this value.'; return; }
  if (isoToDay(iso) > todayDay()) { err.textContent = 'That date is in the future. Choose today or earlier.'; return; }
  err.textContent = '';
  valueOpen = false;
  if (mode === 'data' && state.value && state.value.a === round2(a) && state.value.d === iso) { renderValueForm(lastM); return; }
  let latest = iso;
  const prev = commitData(s => {
    upsertHistory(s, iso, round2(a));
    const last = s.history[s.history.length - 1];
    s.value = { a: last.a, d: last.d };
    latest = last.d;
  });
  toast(latest === iso ? 'Value saved: ' + money(a, { dec: 0 }) + ' on ' + fmtDate(iso)
    : 'Saved for ' + fmtDate(iso) + '. Your current value is still the one from ' + fmtDate(latest) + '.', prev, latest === iso ? undefined : 8000);
}
$('#value-form').addEventListener('submit', e => { e.preventDefault(); saveValue(); });
$('#value-form').addEventListener('keydown', e => { if (e.key === 'Escape') { valueOpen = false; renderValueForm(lastM); $('#val-open').focus(); } });
$('#val-amount').addEventListener('input', () => { $('#val-error').textContent = ''; paintValueDelta(); });

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
$('#set-keep').addEventListener('change', e => setKeep(e.target.checked));
$('#wel-keep').addEventListener('change', e => setKeep(e.target.checked));

$('#goal-form').addEventListener('submit', e => { e.preventDefault(); $('#goal-amount').blur(); });
$('#goal-amount').addEventListener('change', e => {
  const raw = e.target.value.trim(), err = $('#goal-error');
  if (!raw) { err.textContent = ''; setGoal(null); return; }
  const v = parseInputAmount(raw);
  if (!isFinite(v) || v <= 0) { err.textContent = 'Enter the amount you’re aiming for, like 250,000.'; return; }
  err.textContent = '';
  setGoal(v);
});
$('#p-years').addEventListener('input', e => { proj.years = +e.target.value; storePlan(); renderPlan(lastM); });
$('#p-rate').addEventListener('input', e => { proj.rate = +e.target.value / 100; proj.preset = ''; storePlan(); renderPlan(lastM); });
$('#p-annual').addEventListener('input', e => { const v = parseInputAmount(e.target.value); if (isFinite(v) && v >= 0) { proj.annual = v; proj.annualTouched = true; storePlan(); renderPlan(lastM); } });
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
/* amounts are shown in full while an import is being checked; hide them again when it closes */
dlgImport.addEventListener('close', () => { if (hideAmounts) render(); });
$('#open-file').addEventListener('change', e => openFile(e.target.files && e.target.files[0]));

/* resize, and keeping "saved 5 minutes ago" current */
let resizeTimer = null, lastW = window.innerWidth;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => { if (Math.abs(window.innerWidth - lastW) > 2) { lastW = window.innerWidth; render(); } }, 120);
});
document.addEventListener('visibilitychange', () => { if (!document.hidden) renderHeader(); });
setInterval(() => { if (!document.hidden) renderHeader(); }, 60000);
if (PWA && 'serviceWorker' in navigator && window.isSecureContext) {
  window.addEventListener('load', () => { navigator.serviceWorker.register('sw.js').catch(() => { /* runs online-only without it */ }); });
  /* A Home Screen app isn't reloaded when you go back to it: the phone keeps the page paused in the
     background. So new S&P prices are fetched when you return (and at start, in case a slow
     connection meant the saved copy was used), and taken in without reloading the page. */
  navigator.serviceWorker.ready.then(() => setTimeout(() => checkSpx(false), 4000)).catch(() => {});
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) return;
    if (Date.now() - spxChecked > 5 * 60000) checkSpx(false);
    if (Date.now() - swChecked > 30 * 60000) {
      swChecked = Date.now();
      navigator.serviceWorker.getRegistration().then(r => r && r.update()).catch(() => {});
    }
  });
  window.addEventListener('online', () => checkSpx(false));
  /* A new version of the app arrived while this page was open: it's used from the next launch. */
  const hadController = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (hadController) toast('The app has been updated. Close it fully (swipe it away) and open it again to use the new version.', undefined, 9000);
  });
}

/* Always opens on Overview: your value and key figures first. */
syncPlanFromState();
setTab('overview', { keepScroll: true });
})();
