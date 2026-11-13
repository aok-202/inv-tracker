/* build 2026-11-13.2 */
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
  swap: '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 7h11M12 4l3 3-3 3M16 13H5M8 10l-3 3 3 3"/></svg>',
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

/* ---------- motion ----------
   Things move when you do something: open a file, change tab, pick a period. Renders caused by anything else
   (typing, an undo, new prices arriving) draw the finished picture at once. `fresh` names what plays next:
   'tab' (everything on the tab being shown), or one part of it ('growth', 'perf', 'board', 'health'). */
const fresh = new Set();
const plays = k => !REDUCED && (fresh.has('tab') || fresh.has(k));
/* redraw one part with its motion, in answer to a tap */
function replay(k, fn) { fresh.add(k); try { fn(); } finally { fresh.clear(); } }
const EASE = 'cubic-bezier(.2,.8,.2,1)';
function anim(el, frames, opts) {
  try { return el.animate(frames, Object.assign({ easing: EASE, fill: 'backwards' }, opts)); } catch (e) { return null; }
}
/* A chart draws itself: lines run left to right, filled areas and points fade in behind them. */
function drawIn(svg) {
  let k = 0;
  svg.querySelectorAll('path.area, path.wedge, path.band').forEach(p => anim(p, [{ opacity: 0, offset: 0 }], { duration: 700, delay: 260, easing: 'ease-out' }));
  svg.querySelectorAll('path.line').forEach(p => {
    if (p.classList.contains('dash')) { anim(p, [{ opacity: 0, offset: 0 }], { duration: 500 }); return; }
    let len = 0;
    try { len = p.getTotalLength(); } catch (e) { len = 0; }
    if (!(len > 0)) return;
    p.style.strokeDasharray = len + ' ' + len;
    const a = anim(p, [{ strokeDashoffset: len }, { strokeDashoffset: 0 }], { duration: 950, delay: 70 * k++, easing: 'cubic-bezier(.3,.6,.2,1)' });
    const done = () => { p.style.strokeDasharray = ''; };
    if (a) { a.onfinish = done; a.oncancel = done; } else done();
  });
  svg.querySelectorAll('circle.dot').forEach(c => anim(c, [{ opacity: 0, transform: 'scale(.3)', offset: 0 }], { duration: 320, delay: 880, easing: 'cubic-bezier(.3,1.5,.5,1)' }));
  svg.querySelectorAll('line.ref, text.ref-label').forEach(c => anim(c, [{ opacity: 0, offset: 0 }], { duration: 400, delay: 700 }));
}
function growBars(svg) {
  svg.querySelectorAll('path.bar').forEach((b, i) => anim(b, [{ transform: 'scaleY(0)', offset: 0 }], { duration: 560, delay: Math.min(i * 30, 420) }));
  svg.querySelectorAll('text.val').forEach(t => anim(t, [{ opacity: 0, offset: 0 }], { duration: 300, delay: 520 }));
}
/* A number that counts up to its value. */
function countUp(el, to, fmt, ms) {
  if (REDUCED || !(Math.abs(to) > 0)) { el.textContent = fmt(to); return; }
  const t0 = performance.now(), dur = ms || 700;
  const step = now => {
    const k = Math.min(1, (now - t0) / dur), e = 1 - Math.pow(1 - k, 3);
    el.textContent = fmt(to * e);
    if (k < 1 && el.isConnected) requestAnimationFrame(step); else el.textContent = fmt(to);
  };
  el.textContent = fmt(0);
  requestAnimationFrame(step);
}
/* The headline figure is an odometer: each digit is a strip of 0 to 9 that rolls to its number. Characters are
   matched from the right, so the ones stay the ones when the amount gains or loses a digit. */
function setOdometer(el, text, opts) {
  opts = opts || {};
  const chars = Array.from(text), old = el._od || [], out = [];
  const plainText = !/\d/.test(text) || opts.plain;
  if (plainText) { el._od = null; el.textContent = text; return; }
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i], digit = ch >= '0' && ch <= '9', prev = old[old.length - (chars.length - i)];
    let node;
    if (prev && prev.digit === digit && (digit || prev.ch === ch)) node = prev.node;
    else if (digit) {
      node = h('span', { class: 'od-c od-d' }, h('span', { class: 'od-s' }, '0\n1\n2\n3\n4\n5\n6\n7\n8\n9'));
      if (!opts.instant && !REDUCED) node.firstChild.style.setProperty('--d', opts.from === 'zero' ? '0' : ch);
    } else node = h('span', { class: 'od-c' }, ch);
    out.push({ node, ch, digit });
  }
  if (!el._od) el.textContent = '';
  el.replaceChildren(...out.map(o => o.node));
  el._od = out;
  const set = () => { for (const o of out) if (o.digit) o.node.firstChild.style.setProperty('--d', o.ch); };
  if (opts.from === 'zero' && !REDUCED) { el.offsetWidth; requestAnimationFrame(set); } else set();
  el.classList.toggle('quick', !!opts.quick);
}

/* ---------- state ---------- */
function emptyState() { return { v: 1, deposits: [], value: null, history: [], settings: { cur: 'USD', spx: null, goal: null, plan: null, cmp: ['SPX', null, null], main: 'SPX' }, updatedAt: 0 }; }
const BENCH_ID = /^[A-Z0-9._^-]{1,12}$/;
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
    /* up to three things to compare with, each keeping its place (and so its colour), and which one is the main one */
    const c = o.settings.cmp;
    if (Array.isArray(c)) {
      const seen = new Set();
      const slots = [0, 1, 2].map(i => { const id = typeof c[i] === 'string' && BENCH_ID.test(c[i]) && !seen.has(c[i]) ? c[i] : null; if (id) seen.add(id); return id; });
      if (slots.some(Boolean)) s.settings.cmp = slots;
    }
    const mn = o.settings.main;
    s.settings.main = typeof mn === 'string' && s.settings.cmp.includes(mn) ? mn : s.settings.cmp.find(Boolean);
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
  s.settings.cmp = ['SPX', 'VT', 'MIX'];
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

/* ---------- what your portfolio can be compared with ----------
   The S&P 500 ships with the app (spx-data.js). Everything else comes from bench-data.js, which the daily data
   job fills in; until it has, the list below shows those as "on the way". */
const BENCH_DATA = (typeof BENCH !== 'undefined' && BENCH && Array.isArray(BENCH.series)) ? BENCH : { v: 1, updated: null, series: [] };
const CATALOG = {
  SPX: { short: 'S&P 500', name: 'S&P 500', group: 'US stocks', kind: 'stock', note: 'The 500 largest US companies' },
  QQQ: { short: 'Nasdaq-100', name: 'Nasdaq-100', group: 'US stocks', kind: 'stock', note: 'Large US technology and growth companies (QQQ fund)' },
  VTI: { short: 'US market', name: 'US total market', group: 'US stocks', kind: 'stock', note: 'Every listed US company (VTI fund)' },
  DIA: { short: 'Dow Jones', name: 'Dow Jones', group: 'US stocks', kind: 'stock', note: '30 large US companies (DIA fund)' },
  IWM: { short: 'US small caps', name: 'US small companies', group: 'US stocks', kind: 'stock', note: 'The Russell 2000 (IWM fund)' },
  VT: { short: 'World', name: 'World stocks', group: 'World', kind: 'stock', note: 'Companies in every market, the US included (VT fund)' },
  VEA: { short: 'Dev. ex-US', name: 'Developed ex-US', group: 'World', kind: 'stock', note: 'Europe, Japan and other developed markets (VEA fund)' },
  VWO: { short: 'Emerging', name: 'Emerging markets', group: 'World', kind: 'stock', note: 'China, India, Brazil and others (VWO fund)' },
  MIX: { short: '60/40', name: '60/40 mix', group: 'Bonds, cash and more', kind: 'mix', note: '60% S&P 500 and 40% US bonds, rebalanced every month' },
  AGG: { short: 'US bonds', name: 'US bonds', group: 'Bonds, cash and more', kind: 'bond', note: 'Government and company bonds (AGG fund)' },
  CASH: { short: 'Cash', name: 'Cash', group: 'Bonds, cash and more', kind: 'cash', note: '3-month US Treasury bills' },
  CPI: { short: 'Inflation', name: 'US inflation', group: 'Bonds, cash and more', kind: 'prices', note: 'Consumer prices: what it takes just to keep up' },
  GLD: { short: 'Gold', name: 'Gold', group: 'Bonds, cash and more', kind: 'other', note: 'The gold price (GLD fund)' },
  BTC: { short: 'Bitcoin', name: 'Bitcoin', group: 'Bonds, cash and more', kind: 'other', note: 'The bitcoin price in US dollars' }
};
const GROUPS = ['US stocks', 'World', 'Bonds, cash and more', 'Other'];
/* Stock and bond funds are the field your return is ranked in; cash, prices, gold and bitcoin are yardsticks beside it. */
const isPeer = b => b.kind === 'stock' || b.kind === 'bond' || b.kind === 'mix';
let benchRev = 0, benchCache = null;
function validSeries(x) {
  return !!x && typeof x.id === 'string' && BENCH_ID.test(x.id) && x.id !== 'SPX' && x.id !== 'MIX' && /^\d{4}-\d{2}$/.test(x.start || '') && Array.isArray(x.tr) && x.tr.length >= 2 &&
    x.tr.every(v => typeof v === 'number' && v > 0) && !!x.last && /^\d{4}-\d{2}-\d{2}$/.test(x.last.d || '') && x.last.tr > 0;
}
function benchList() {
  const key = benchRev + '|' + JSON.stringify(state.settings.spx);
  if (benchCache && benchCache.key === key) return benchCache.list;
  const list = [], byId = new Map();
  const add = (id, A, extra) => {
    const c = CATALOG[id] || {};
    const b = Object.assign({ id, name: c.name || (extra && extra.name) || id, short: c.short || null, group: c.group || 'Other', kind: c.kind || 'stock', note: c.note || (extra && extra.code) || '',
      A: A && A.length >= 2 ? A : null }, extra ? { priceOnly: !!extra.priceOnly, src: extra.src } : {});
    b.avail = !!b.A;
    b.startT = b.avail ? b.A[0].t : null;
    b.lastT = b.avail ? b.A[b.A.length - 1].t : null;
    list.push(b); byId.set(id, b);
  };
  add('SPX', buildAnchors(SPX, state.settings.spx), { src: 'spx' });
  const stored = new Map(BENCH_DATA.series.filter(validSeries).map(x => [x.id, x]));
  for (const id of Object.keys(CATALOG)) {
    if (id === 'SPX') continue;
    if (id === 'MIX') { const agg = byId.get('AGG') || (stored.has('AGG') ? { A: seriesAnchors(stored.get('AGG')) } : null); add('MIX', agg && agg.A ? blendAnchors([{ A: byId.get('SPX').A, w: 0.6 }, { A: agg.A, w: 0.4 }]) : null, { src: 'mix' }); continue; }
    const x = stored.get(id);
    add(id, x ? seriesAnchors(x) : null, x || null);
  }
  for (const [id, x] of stored) if (!CATALOG[id]) add(id, seriesAnchors(x), x);
  /* keep the catalog's reading order: US stocks, world, then the rest */
  list.sort((a, b) => GROUPS.indexOf(a.group) - GROUPS.indexOf(b.group));
  benchCache = { key, list, byId };
  return list;
}
const benchById = id => { benchList(); return benchCache.byId.get(id) || null; };
/* The ones picked for the charts, in their places (a place keeps its colour), and the main one among them. */
function cmpSlots() {
  const s = state.settings.cmp.map(id => { const b = id ? benchById(id) : null; return b && b.avail ? b : null; });
  /* a file can name picks whose data hasn't arrived yet: until it does, the S&P 500 stands in */
  if (!s.some(Boolean)) s[0] = benchById('SPX');
  return s;
}
function mainBench() {
  const slots = cmpSlots();
  const m = slots.find(b => b && b.id === state.settings.main) || slots.find(Boolean);
  return m || benchById('SPX');
}
const slotOf = id => cmpSlots().findIndex(b => b && b.id === id);
/* series classes: s0, s1, s2 are the three places; the main index keeps its place's colour everywhere */
const slotCls = id => 's' + Math.max(0, slotOf(id));
const mainName = () => mainBench().name;
const theMain = () => { const n = mainName(); return /^(S&P|Nasdaq|Dow)/.test(n) ? 'the ' + n : n; };

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
  benchRev++;
  return true;
}
/* The other indexes and funds, from this site's own bench-data.js. Taken when it holds something newer or more. */
function applyBench(text) {
  const m = /const BENCH = (\{[\s\S]*?\});/.exec(text || '');
  if (!m) return false;
  let d;
  try { d = JSON.parse(m[1]); } catch (e) { return false; }
  if (!d || !Array.isArray(d.series)) return false;
  const next = d.series.filter(validSeries);
  const stamp = list => list.map(x => x.id + ':' + x.tr.length + ':' + x.last.d).sort().join(',');
  if (!next.length || stamp(next) === stamp(BENCH_DATA.series.filter(validSeries))) return false;
  BENCH_DATA.series = next;
  BENCH_DATA.updated = d.updated || null;
  benchRev++;
  return true;
}
function checkSpx(manual) {
  if (!spxCanCheck()) {
    if (manual) toast('Close the app fully and open it again to load the latest prices.', undefined, 6000);
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
    const had = BENCH_DATA.series.length;
    const newSpx = !!(reply && reply.ok && applySpx(reply.text));
    const newBench = !!(reply && reply.ok && applyBench(reply.bench));
    if (newSpx || newBench) {
      render();
      toast(!had && newBench ? 'More indexes and funds to compare with have arrived. Pick them under Performance.'
        : newSpx ? 'Prices updated to the close of ' + fmtDay(isoToDay(SPX.last.d)) + '.' : 'Index and fund prices updated.', undefined, 4500);
    } else {
      if (lastM && mode !== 'welcome') renderBenchmark(lastM);
      renderSettings();
      if (manual) toast(reply && reply.ok ? 'No newer prices yet. The latest S&P 500 close here is ' + fmtDay(isoToDay(SPX.last.d)) + '.'
        : 'Couldn’t check for new prices. Make sure you’re online, then try again.', undefined, 5000);
    }
  };
  setTimeout(() => finish(null), 20000);
  ch.port1.onmessage = e => finish(e.data);
  try { navigator.serviceWorker.controller.postMessage({ type: 'data-refresh' }, [ch.port2]); } catch (e) { finish(null); }
}
function getAnchors() { return mainBench().A; }

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
  const prev = tab;
  if (t !== tab) scrollPos[tab] = window.scrollY;
  tab = t;
  for (const x of TABS) {
    $('#view-' + x).hidden = x !== t;
    const b = $('#tabbtn-' + x);
    b.setAttribute('aria-selected', String(x === t));
    b.tabIndex = x === t ? 0 : -1;
  }
  $('#tab-ind').style.setProperty('--i', String(TABS.indexOf(t)));
  if (!opts.quiet) fresh.add('tab');
  render();
  if (!opts.keepScroll) window.scrollTo(0, scrollPos[t] || 0);
  /* the new tab slides in from the side it sits on */
  if (t !== prev && !REDUCED && !opts.quiet) {
    const dir = TABS.indexOf(t) > TABS.indexOf(prev) ? 1 : -1;
    anim($('#view-' + t), [{ opacity: 0, transform: 'translateX(' + dir * 26 + 'px)' }, { opacity: 1, transform: 'none' }], { duration: 300, fill: 'none' });
    const ic = $('#tabbtn-' + t + ' svg');
    if (ic) anim(ic, [{ transform: 'scale(.82)' }, { transform: 'scale(1.12)', offset: 0.55 }, { transform: 'none' }], { duration: 380, fill: 'none' });
  }
  syncBar();
}
/* The top bar: when the page's own headline has scrolled away, the bar shows a small copy of it. */
let barAt = Infinity;
function syncBar() {
  const bar = $('#topbar'), y = window.scrollY;
  bar.classList.toggle('scrolled', y > 6);
  bar.classList.toggle('compact', y > barAt);
}
function measureBar() {
  const head = tab === 'overview' ? (mode === 'welcome' ? null : $('#hero-fig')) : $('#view-' + tab + ' .view-title');
  if (!head || !head.offsetParent) { barAt = Infinity; syncBar(); return; }
  barAt = head.getBoundingClientRect().bottom + window.scrollY - $('#topbar').offsetHeight - 4;
  const title = $('#bar-mini-title'), sub = $('#bar-mini-sub');
  if (tab === 'overview') {
    title.textContent = $('#hero-value').textContent;
    const pc = lastM && lastM.hasValue && lastM.roi != null ? pct(lastM.roi) : '';
    sub.textContent = pc;
    sub.className = 'bar-mini-sub ' + (pc ? (lastM.roi >= 0 ? 'good' : 'bad') : '');
  } else { title.textContent = head.textContent; sub.textContent = ''; }
  syncBar();
}
window.addEventListener('scroll', syncBar, { passive: true });

/* ---------- rendering ---------- */
let editing = null;
const openYears = new Set();
let histOpen = false, histAll = false, valueOpen = false;
let spxOpen = false;
const proj = { years: 10, annual: null, annualTouched: false, rate: 0.07, preset: '7' };
const perf = { preset: 'all' };
let lastM = null, lastG = null, lastP = null, lastB = null;
const board = { mode: 'money', period: 'all' };

/* Everything there is to compare with, worked out once per render: the same deposits in each index,
   and each index's return between your saved values. */
function computeBench(M, G) {
  const main = mainBench();
  const rows = benchList().filter(b => b.avail).map(b => {
    const isMain = b.id === main.id;
    const P = G ? (isMain ? lastP : perfSeries(G, M.deps, b.A)) : null;
    return { b, what: isMain ? M.sp : sameDeposits(M, b.A), P, periods: P ? perfPeriods(P) : [] };
  });
  return { main, rows, byId: new Map(rows.map(r => [r.b.id, r])) };
}

function render() {
  const A = getAnchors();
  const M = computeMetrics(state, A, todayDay());
  lastM = M;
  lastG = mode === 'welcome' ? null : growthSeries(M, state.history, A);
  lastP = lastG ? perfSeries(lastG, M.deps, A) : null;
  lastB = mode === 'welcome' ? null : computeBench(M, lastG);
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
      $('#growth-more').hidden = !lastG;
      renderGrowth(M); renderValueForm(M); renderVerdict(M); renderKpis(M); renderHighlights(M); renderYears(M);
    }
  }
  else if (tab === 'deposits') { renderLedger(M); slideThumb($('#kind-dep').parentElement); }
  else if (tab === 'performance') renderPerformance(M);
  else renderPlan(M);
  renderSettings();
  if (dlgPicker.open) renderPicker();
  fresh.clear();
  measureBar();
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
  document.body.classList.toggle('has-banner', !b.hidden);
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
  const fig = $('#hero-fig'), valText = $('#hero-value'), chips = $('#hero-chips'), meta = $('#hero-asof'), stage = $('#hero');
  const G = gz.G, V = gz.V;
  stage.classList.toggle('reading', !!read);
  fill(chips);
  let text, trend = null;
  if (read) {
    text = read.value != null ? money(read.value, { dec: 0 }) : 'No value saved';
    if (read.value != null) {
      trend = read.value - read.dep;
      chips.append(...changeParts(read.value - read.dep, read.dep > 0 ? (read.value - read.dep) / read.dep : null, 'on ' + money(read.dep, { dec: 0 }) + ' deposited'));
    } else chips.append(h('span', { class: 'chg-label' }, money(read.dep, { dec: 0 }) + ' deposited by then'));
    meta.textContent = fmtDay(read.t);
  } else {
    text = M.hasValue ? money(M.V, { dec: 0 }) : 'No value yet';
    const c = G && V && M.hasValue && isZoomed(G, V) ? V.cmp : null;
    const latest = G && G.vals.length ? G.vals[G.vals.length - 1].t : null;
    if (c && c.problem) chips.append(h('span', { class: 'chg-label' }, 'This period needs saved values inside it to show a change.'));
    else if (c && !(c.fromStart && c.t1 === latest)) {
      const label = c.t1 === latest ? 'since ' + fmtDay(c.t0) : c.fromStart ? 'up to ' + fmtDay(c.t1) : fmtDay(c.t0) + ' to ' + fmtDay(c.t1);
      chips.append(...changeParts(c.gainV, c.retV, label));
      trend = c.gainV;
    } else if (M.hasValue && M.gain != null) { chips.append(...changeParts(M.gain, M.roi, 'all time')); trend = M.gain; }
    if (M.hasValue) {
      const age = M.today - M.vDay;
      meta.textContent = 'As of ' + fmtDate(state.value.d) + (age > 31 ? ', ' + age + ' days ago' : '');
    } else meta.textContent = M.count ? 'Add your portfolio’s value to see your returns.' : 'Add your deposits, then your portfolio’s value, to see your returns.';
  }
  const none = read ? read.value == null : !M.hasValue;
  valText.textContent = text;
  fig.parentElement.classList.toggle('none', none);
  /* the light behind the chart follows the result: up or down over what is shown */
  stage.dataset.trend = trend == null ? 'flat' : trend >= 0 ? 'up' : 'down';
  setOdometer(fig, text, { plain: none || hideAmounts, quick: !!read, from: fig._od ? null : 'zero' });
  fitFigure(fig, 30, fig.parentElement);
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
    { label: 'Against ' + theMain(), value: M.sp.diff != null ? money(M.sp.diff, { sign: true, dec: 0 }) : '—', tone: tone(M.sp.diff),
      cap: M.sp.value != null ? 'Same deposits in it: ' + money(M.sp.value, { dec: 0 }) : (M.sp.unavailable ? 'Its data starts after your first deposit' : 'Needs deposits and a value') }
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
        ' (' + pct(ytd.you) + ')' + (ytd.sp != null ? ', while ' + theMain() + ' returned ' + pct(ytd.sp) : '') + '.');
    }
  }
  /* what else the same deposits could have become, among the stock and bond funds there is data for */
  const Bn = lastB;
  if (Bn) {
    const alts = Bn.rows.filter(r => isPeer(r.b) && r.what.value != null).sort((a, b) => b.what.value - a.what.value);
    if (alts.length >= 3) {
      const top = alts[0];
      if (top.what.value > M.V) add('swap', 'plain', 'The same deposits in ', B(top.b.name), ' would be worth ' + money(top.what.value, { dec: 0 }) + ', ' + money(top.what.value - M.V, { dec: 0 }) + ' more than you have. It did best of the ' + alts.length + ' funds here.');
      else add('swap', 'good', B('None of the ' + alts.length + ' funds'), ' here would have made more of your deposits.');
    }
  }
  /* money-weighted against time-weighted: the gap is what the timing of your deposits did */
  const whole = P ? perfPeriods(P).find(r => r.k === 'all') : null;
  if (whole && whole.youAnn != null && M.xirr != null && M.first != null && Math.abs(P.t0 - M.first) <= 45 && P.t1 === M.vDay) {
    const d = M.xirr - whole.youAnn;
    if (Math.abs(d) >= 0.005) add('clock', d > 0 ? 'good' : 'plain', 'When you paid in ', B((d > 0 ? 'added ' : 'cost ') + num(Math.abs(d) * 100, 1) + ' points a year'),
      ': your money earned ' + pct(M.xirr) + ' a year, the investments themselves ' + pct(whole.youAnn) + '.');
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
  return out.slice(0, 6);
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
    svg.append(sv('path', { d, class: 'line s-' + s.cls + (s.thin ? ' thin' : '') + (s.dash ? ' dash' : '') + (s.glow ? ' glow' : '') }));
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
  if (o.animate) drawIn(svg);
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
    const p = Math.abs(yv - y0) >= 0.5 ? sv('path', { d: roundedBar(cx - bw / 2, bw, y0, yv, 4), class: 'bar f-' + o.cls + (b.faded ? ' faded' : '') + (b.value < 0 ? ' neg' : '') }) : sv('g');
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
  if (o.animate) growBars(svg);
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
  drawGrowth(plays('growth'));
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

function drawGrowth(play) {
  const G = gz.G, M = gz.M;
  if (!G || mode === 'welcome') return;
  const zoomable = canZoom(G), w = currentWindow(G), zoomed = isZoomed(G, w);
  const main = mainBench();
  const V = windowSeries(G, M.deps, main.A, w.a, w.b);
  gz.V = V;
  /* The other picks are sampled on the same days, so their points line up with the main one's by position. */
  const picks = [{ b: main, pts: V.pts }].concat(cmpSlots().filter(b => b && b.id !== main.id)
    .map(b => ({ b, pts: windowSeries(Object.assign({}, G, { spOK: G.start >= b.startT }), M.deps, b.A, w.a, w.b).pts })));
  picks.sort((x, y) => slotOf(x.b.id) - slotOf(y.b.id));
  const lines = picks.map(x => ({ b: x.b, cls: slotCls(x.b.id), at: new Map(x.pts.map(p => [p.t, p.sp])), pts: x.pts.filter(p => p.sp != null).map(p => [p.t, p.sp]) })).filter(x => x.pts.length);
  const plot = $('#growth-plot');
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
  for (const l of lines) series.push({ cls: l.cls, type: 'line', thin: true, end: true, pts: l.pts });
  if (vline.length) series.push({ cls: 'value', type: 'line', glow: true, pts: vline, dots: drawn.length <= 14, dotPts: drawn.map(v => [v.t, v.a]), end: !!lastReal, endPt: lastReal ? [lastReal.t, lastReal.a] : null });
  const valBy = new Map(V.vin.map(v => [v.t, v.a]));
  /* Reading the chart stops at your own values (the points on the line), so every stop shows your portfolio
     next to the indexes and your deposits. With no values in view it stops at deposits and month ends. */
  const ptBy = new Map(V.pts.map(p => [p.t, p]));
  const stops = drawn.length >= 2 ? drawn.map(v => v.t).filter(t => ptBy.has(t)) : V.pts.map(p => p.t);
  const lastP = V.pts[V.pts.length - 1];
  /* The legend carries the numbers: for the end of the period, or for the point being read. */
  const paintLegend = p => {
    const youV = p ? (valBy.has(p.t) ? valBy.get(p.t) : null) : lastReal ? lastReal.a : null;
    const box = $('#growth-legend'), depV = (p || lastP).dep;
    const bm = lines.map(l => ({ l, v: p ? l.at.get(p.t) : l.pts[l.pts.length - 1][1] }));
    /* the figures share a row: when full amounts won't fit a narrow phone, they all go short (12.3M) */
    let fmt = v => (v == null ? '—' : money(v, { dec: 0 }));
    const colW = ((box.clientWidth || 340) - 24) / 3;
    if (!isWide() && [youV, depV].concat(bm.map(x => x.v)).some(v => textW(fmt(v), 16) * 1.12 > colW)) fmt = v => (v == null ? '—' : money(v, { compact: true }));
    fill(box,
      vline.length ? legendItem('value', 'You', fmt(youV)) : null,
      ...bm.map(x => legendItem(x.l.cls, x.l.b.short || x.l.b.name, fmt(x.v))),
      legendItem('dep', 'Deposited', fmt(depV), true));
  };
  gz.api = lineChart(plot, {
    height: isWide() ? 330 : 236, x0: w.a, x1: w.b, external: true, tip: false, yZero: !zoomed, padL: isWide() ? 0 : 8, padR: isWide() ? 0 : 6, animate: play,
    xTicks: timeTicks(w.a, w.b, plot.clientWidth || 600),
    series, yFmt: yMoney, samples: stops,
    focus: i => {
      const p = ptBy.get(stops[i]), f = [{ cls: 'dep', x: p.t, y: p.dep }];
      for (const l of lines) { const v = l.at.get(p.t); if (v != null) f.push({ cls: l.cls, x: p.t, y: v }); }
      if (valBy.has(p.t)) f.push({ cls: 'value', x: p.t, y: valBy.get(p.t) });
      return f;
    },
    onRead: i => {
      const p = i < 0 ? null : ptBy.get(stops[i]);
      paintLegend(p);
      paintHero(p ? { t: p.t, value: valBy.has(p.t) ? valBy.get(p.t) : null, dep: p.dep } : null);
    },
    label: 'Line chart' + (zoomed ? ' from ' + fmtDay(w.a) + ' to ' + fmtDay(w.b) : '') + ': money deposited, your portfolio value, and ' +
      (V.rebased ? lines.map(l => l.b.name).join(', ') + ' starting from your value on ' + fmtDay(V.cmp.t0) : 'the same deposits in ' + (lines.map(l => l.b.name).join(', ') || 'an index')) + ', over time'
  });
  $('#growth-chart').classList.toggle('zoomed', zoomed);
  $('#growth-nav').hidden = !zoomable;
  if (zoomable) navLayout(G, w);
  renderPresets(G, zoomable);
  const names = lines.map(l => l.b.short || l.b.name);
  const spNote = lines.length && !V.rebased ? (names.length > 1 ? 'The other lines: ' : names[0] + ': ') + 'your deposits, put into ' + (names.length > 1 ? 'each one' : 'it') + ' on the same days. ' : '';
  $('#growth-window').textContent = zoomed ? fmtDay(w.a) + ' to ' + fmtDay(w.b) + ' (' + spanText(w.b - w.a) + ')'
    : spNote + (zoomable ? 'Pinch the chart or drag the handles for a shorter period.' : '');
  renderCompare(V, zoomed);
  paintLegend(null);
  paintHero(null);
}

/* Period buttons with a pill that slides to the one in use. */
const thumbAt = new WeakMap();
function slideThumb(box) {
  const btn = box.querySelector('[aria-pressed="true"], [aria-checked="true"]');
  let th = box.querySelector(':scope > .thumb');
  if (!btn || !btn.offsetWidth) { if (th) th.remove(); thumbAt.delete(box); return; }
  const made = !th;
  if (made) { th = h('span', { class: 'thumb', 'aria-hidden': 'true' }); box.prepend(th); }
  const now = { x: btn.offsetLeft, y: btn.offsetTop, w: btn.offsetWidth, h: btn.offsetHeight }, was = thumbAt.get(box);
  const place = r => { th.style.width = r.w + 'px'; th.style.height = r.h + 'px'; th.style.transform = 'translate(' + r.x + 'px,' + r.y + 'px)'; };
  /* a pill made just now starts where the last one stood, then slides */
  if (made) { th.style.transition = 'none'; place(was && !REDUCED ? was : now); th.offsetWidth; th.style.transition = ''; }
  place(now);
  thumbAt.set(box, now);
}

function renderPresets(G, zoomable) {
  const box = $('#growth-presets');
  if (!zoomable) { fill(box); thumbAt.delete(box); return; }
  fill(box, ...PRESETS.filter(p => presetRange(G, p.k)).map(p =>
    h('button', { type: 'button', 'data-action': 'zoom-preset', 'data-k': p.k, 'aria-label': p.name, 'aria-pressed': String(zoom.preset === p.k) }, p.label)));
  slideThumb(box);
}

function renderCompare(V, zoomed) {
  const box = $('#growth-compare');
  if (!zoomed) { box.hidden = true; fill(box); return; }
  box.hidden = false;
  const c = V.cmp;
  const addLink = () => (mode === 'example' ? null : h('button', { type: 'button', class: 'link', 'data-action': 'open-import', 'data-target': 'values' }, 'Add many values'));
  if (c.problem) {
    const mb = mainBench(), nm = theMain();
    const msg = c.problem === 'nosp' ? 'The data for ' + nm + ' starts in ' + dfMonthYear.format(new Date(mb.startT * DAY_MS)) + ', so this period can’t be compared with it.'
      : c.fromStart ? 'To compare this period with ' + nm + ', the app needs your portfolio value on a date inside it.'
        : 'To compare this period with ' + nm + ', the app needs your portfolio value on at least two dates inside it.';
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
      h('tbody', null, row('value', 'Your portfolio', c.gainV, c.retV), row(slotCls(mainBench().id), mainName(), c.gainS, c.retS)))),
    h('p', { class: 'verdict' },
      h('span', { class: 'pill ' + (ahead ? 'pill-good' : 'pill-bad') }, h('span', { 'aria-hidden': 'true' }, ahead ? '▲' : '▼'),
        money(Math.abs(c.diff), { dec: 0 }) + (ahead ? ' ahead of ' : ' behind ') + theMain())),
    h('p', { class: 'fine' }, c.fromStart
      ? 'Gains after deposits. The other side gets the same deposits on the same days. Return counts each deposit from the day it went in.'
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
    h('thead', null, h('tr', null, h('th', { scope: 'col' }, 'Date'), h('th', { scope: 'col' }, 'Deposited'), h('th', { scope: 'col' }, 'Your portfolio'), h('th', { scope: 'col' }, mainName() + ', same deposits'))),
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

/* Where the prices come from and how fresh they are. */
function renderBenchmark(M) {
  if (mode === 'welcome') return;
  const body = $('#bm-body'), spx = benchById('SPX'), lastA = spx.A[spx.A.length - 1], main = mainBench();
  $('#spx-form').hidden = !spxOpen;
  $('#spx-clear').hidden = !state.settings.spx;
  const parts = [];
  const others = benchList().filter(b => b.avail && b.id !== 'SPX' && b.id !== 'MIX');
  if (M.sp.stale) parts.push(h('p', { class: 'warn-note' }, 'Your value is dated after the last close stored here for ' + theMain() + ' (' + fmtDay(M.sp.lastDay) + '), so its side of each comparison stops on that day.' + (main.id === 'SPX' ? ' Add a newer close below for a like-for-like comparison.' : '')));
  if (M.sp.unavailable === 'early') parts.push(h('p', { class: 'warn-note' }, 'The data for ' + theMain() + ' starts in ' + dfMonthYear.format(new Date(main.startT * DAY_MS)) + ', after your first deposit, so the same-deposits comparison isn’t available for it.'));
  parts.push(h('p', { class: 'fine first' },
    'S&P 500 total return (dividends reinvested), month-end data from 1985 to the close of ' + fmtDay(lastA.t) + ' at ' + num(lastA.px, 2) + (lastA.user ? ', which you added' : '') + '. ',
    spxCanCheck() ? h('button', { type: 'button', class: 'link', 'data-action': 'check-spx' }, spxBusy ? 'Checking…' : 'Check for new prices') : null,
    spxCanCheck() && !(spxOpen || mode === 'example') ? ' ' : null,
    spxOpen || mode === 'example' ? null : h('button', { type: 'button', class: 'link', 'data-action': 'edit-spx' }, 'Add a newer close')));
  if (others.length) {
    const newest = Math.max(...others.map(b => b.lastT)), late = others.filter(b => b.lastT < newest - 45 && b.kind !== 'prices');
    parts.push(h('p', { class: 'fine' }, others.length + ' more indexes and funds, to ' + fmtDay(newest) + '. Funds stand in for their indexes: their prices include dividends and running costs. ' +
      (others.some(b => b.id === 'CPI') ? 'Inflation is the US consumer price index, which comes out a few weeks after each month ends. ' : '') +
      (late.length ? 'Not updated lately: ' + late.map(b => b.name).join(', ') + '. ' : '') +
      (others.some(b => b.priceOnly) ? 'Without dividends (price only): ' + others.filter(b => b.priceOnly).map(b => b.name).join(', ') + '. ' : '')));
  } else if (PWA) {
    parts.push(h('p', { class: 'fine' }, 'The other indexes and funds (Nasdaq-100, world stocks, bonds, gold, cash, inflation and more) haven’t arrived yet. They come with the daily data update on the site this app is served from, and the app picks them up the next time it opens online.'));
  }
  parts.push(h('p', { class: 'fine' }, 'Everything is in US dollars' + (cur() !== 'USD' ? '; your amounts are taken to be dollars too, so currency moves aren’t included' : '') + '. The app itself never contacts a price service.'));
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

/* ---------- how it's doing: one reading of the numbers, shared by Overview and Performance ---------- */
const ordinal = n => { const s = ['th', 'st', 'nd', 'rd'], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); };
/* Where you stand among the stock and bond funds over a period (k: 'all', '1y', ...), by return. */
function rankIn(B, k) {
  const mine = B.byId.get(B.main.id).periods.find(r => r.k === k);
  if (!mine) return null;
  const yearly = mine.youAnn != null && k !== '1y';
  const list = [];
  for (const row of B.rows) {
    const r = row.periods.find(x => x.k === k), v = r ? (yearly ? r.spAnn : r.sp) : null;
    if (v != null) list.push({ b: row.b, v, peer: isPeer(row.b) });
  }
  list.push({ you: true, v: yearly ? mine.youAnn : mine.you, peer: true });
  list.sort((a, b) => b.v - a.v);
  const peers = list.filter(x => x.peer);
  return { pos: peers.findIndex(x => x.you) + 1, of: peers.length, list, peers, yearly, row: mine };
}
function computeHealth() {
  const P = lastP, B = lastB;
  if (!P || !B) return null;
  const mainRow = B.byId.get(B.main.id);
  const all = mainRow.periods.find(r => r.k === 'all'), year = mainRow.periods.find(r => r.k === '1y');
  const months = perfMonths(P), years = perfYears(P), stats = perfStats(P, months);
  const allOf = id => { const row = B.byId.get(id); const r = row && row.b.id !== B.main.id ? row.periods.find(x => x.k === 'all') : null; return r && r.sp != null ? r : null; };
  let cash = null;
  const cashRow = B.byId.get('CASH');
  if (cashRow && cashRow.P) { cash = new Map(); for (const r of perfMonths(cashRow.P)) if (!r.mtd && r.sp != null) cash.set(r.y * 12 + r.m, r.sp); }
  const risk = all && all.sp != null ? riskStats(months, cash) : null;
  const rank = rankIn(B, 'all');
  const ddBm = deepestDrop(P.pts, p => p.sp);
  const checks = all && all.sp != null ? healthChecks({ all, year: year && year.sp != null ? year : null, rank: rank ? { pos: rank.pos, of: rank.of } : null,
    cash: allOf('CASH'), prices: allOf('CPI'), risk, ddYou: stats.dd, ddBm, pts: P.pts.length, years }) : [];
  const n = checks.length, k = checks.filter(c => c.pass).length, r = n ? k / n : 0;
  const verdict = n < 3 ? null : { n, k, tone: r >= 0.6 ? 'good' : r >= 0.4 ? 'warn' : 'bad',
    title: r >= 0.85 ? 'Doing very well' : r >= 0.6 ? 'Doing well' : r >= 0.4 ? 'A mixed picture' : 'Falling behind' };
  return { checks, verdict, all, year, months, years, stats, risk, rank, ddBm };
}
function marketLine(H) {
  const c = H.checks.find(x => x.id === 'market');
  if (!c) return '';
  const d = c.you - c.bm, since = dfMonthYear.format(new Date(c.t0 * DAY_MS));
  if (Math.abs(d) < 0.0005) return 'Level with ' + theMain() + ' since ' + since + '.';
  return (d > 0 ? 'Ahead of ' : 'Behind ') + theMain() + ' by ' + num(Math.abs(d) * 100, 1) + (c.yearly ? ' points a year' : ' points') + ' since ' + since + '.';
}
function checkText(c) {
  const nm = theMain(), per = c.yearly ? ' a year' : '';
  switch (c.id) {
    case 'market': return { title: 'Beats ' + nm, detail: pct(c.you) + per + ' against ' + pct(c.bm) + ', since ' + fmtDay(c.t0) };
    case 'recent': return { title: 'Beats it over the last 12 months', detail: pct(c.you) + ' against ' + pct(c.bm) };
    case 'field': return { title: 'In the top half of the funds', detail: ordinal(c.you) + ' of ' + c.bm + ', counting you and the stock and bond funds here' };
    case 'cash': return { title: 'Beats cash', detail: pct(c.you) + per + ' against ' + pct(c.bm) + ' in Treasury bills' };
    case 'prices': return { title: 'Beats inflation', detail: pct(c.you) + per + ' against ' + pct(c.bm) + ' for US prices' };
    case 'risk': return { title: 'Earns its risk', detail: num(c.you, 2) + ' of return for each point of swing, against ' + num(c.bm, 2) };
    case 'drop': return { title: 'Falls no harder than ' + nm, detail: 'Deepest drop ' + (c.you ? pct(c.you) : 'none') + ' against ' + (c.bm ? pct(c.bm) : 'none') };
    default: return { title: 'Ahead in most years', detail: 'Ahead in ' + c.you + ' of ' + c.bm + ' calendar years' };
  }
}
/* A ring that fills to k of n, with the count in the middle. */
function scoreRing(k, n, tn, play) {
  const R = 30, C = 2 * Math.PI * R, frac = n ? k / n : 0;
  const fillEl = sv('circle', { cx: 38, cy: 38, r: R, class: 'ring-fill ' + tn, 'stroke-dasharray': C.toFixed(1), 'stroke-dashoffset': (C * (1 - frac)).toFixed(1), transform: 'rotate(-90 38 38)' });
  const num0 = sv('tspan', { class: 'ring-k' }, String(k)), svg = sv('svg', { viewBox: '0 0 76 76', width: 76, height: 76, role: 'img', 'aria-label': k + ' of ' + n + ' checks passed' },
    sv('circle', { cx: 38, cy: 38, r: R, class: 'ring-track' }), fillEl,
    sv('text', { x: 38, y: 39, 'text-anchor': 'middle', 'dominant-baseline': 'central', class: 'ring-text' }, num0, sv('tspan', { class: 'ring-n' }, '/' + n)));
  if (play) {
    anim(fillEl, [{ strokeDashoffset: C }, { strokeDashoffset: C * (1 - frac) }], { duration: 900, delay: 150 });
    countUp(num0, k, v => String(Math.round(v)), 800);
  }
  return svg;
}
const ICON_CHECK = '<svg viewBox="0 0 20 20" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4.5 10.5l3.5 3.5 7.5-8"/></svg>';
const ICON_CROSS = '<svg viewBox="0 0 20 20" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M5.5 5.5l9 9M14.5 5.5l-9 9"/></svg>';
function markOf(pass) { const m = h('span', { class: 'mark ' + (pass ? 'yes' : 'no'), 'aria-hidden': 'true' }); m.innerHTML = pass ? ICON_CHECK : ICON_CROSS; return m; }
/* rows of a leaderboard: a place, a name, a bar and a figure; yours stands out */
function boardRow(x, i, o) {
  const li = h('li', { class: 'board-row' + (x.you ? ' you' : '') + (x.peer === false ? ' aside' : ''), style: { '--i': String(i) } },
    h('span', { class: 'board-pos' }, x.pos != null ? String(x.pos) : ''),
    h('span', { class: 'board-name' }, x.you ? 'You' : (x.b.short && o.short ? x.b.short : x.b.name), x.dot ? h('i', { class: 'dot-key key-' + x.dot }) : null),
    h('span', { class: 'board-val' }, o.fmt(x.v)),
    h('span', { class: 'board-track', 'aria-hidden': 'true' }, h('i', { class: 'board-bar' + (x.v < 0 ? ' neg' : ''), style: { '--l': (o.left(x.v) * 100).toFixed(2) + '%', '--w': (o.width(x.v) * 100).toFixed(2) + '%' } })),
    x.note ? h('span', { class: 'board-note ' + (x.noteTone || '') }, x.note) : null);
  return li;
}
function boardScale(vals) {
  const lo = Math.min(0, ...vals), hi = Math.max(0, ...vals), span = hi - lo || 1;
  return { left: v => (Math.min(v, 0) - lo) / span, width: v => Math.max(0.012, Math.abs(v) / span) };
}

/* Overview: the verdict in one card, with the three places around yours. */
function renderVerdict(M) {
  const sec = $('#verdict-sec'), box = $('#verdict'), B = lastB;
  const H = computeHealth(), play = plays('health');
  if (!B || !M.hasValue || !M.count) { sec.hidden = true; return; }
  const more = h('button', { type: 'button', class: 'btn btn-block', 'data-action': 'go-performance' }, 'Open the full comparison');
  if (!H || !H.verdict) {
    /* not enough saved values for the full check: say where the same deposits would stand instead */
    const sp = M.sp;
    if (sp.diff == null) { sec.hidden = true; return; }
    sec.hidden = false;
    const ahead = sp.diff >= 0;
    fill(box, h('div', { class: 'verdict-top' }, markOf(ahead),
      h('div', null, h('h3', { class: 'verdict-title' }, (ahead ? 'Ahead of ' : 'Behind ') + theMain()),
        h('p', { class: 'verdict-line' }, 'The same deposits in ' + theMain() + ' would be worth ' + money(sp.value, { dec: 0 }) + ': ' + money(Math.abs(sp.diff), { dec: 0 }) + (ahead ? ' less' : ' more') + ' than you have.'))),
      h('p', { class: 'note' }, 'Save your value on a few more dates and a fuller check appears here: returns, risk and how you rank.'), more);
    return;
  }
  sec.hidden = false;
  const v = H.verdict, rk = H.rank;
  const parts = [h('div', { class: 'verdict-top' }, h('div', { class: 'score ' + v.tone }, scoreRing(v.k, v.n, v.tone, play)),
    h('div', null, h('h3', { class: 'verdict-title' }, v.title), h('p', { class: 'verdict-line' }, marketLine(H) + ' ' + v.k + ' of ' + v.n + ' checks passed.')))];
  if (rk && rk.peers.length >= 2) {
    const i = rk.pos - 1, a = Math.max(0, Math.min(rk.peers.length - 3, i - 1));
    const near = rk.peers.slice(a, a + 3), sc = boardScale(near.map(x => x.v));
    parts.push(h('ol', { class: 'board mini' + (play ? ' play' : '') }, ...near.map((x, j) => boardRow(Object.assign({ pos: a + j + 1, dot: x.b && slotOf(x.b.id) >= 0 ? slotCls(x.b.id) : null }, x), j,
      { fmt: val => pct(val) + (rk.yearly ? ' a year' : ''), left: sc.left, width: sc.width, short: true }))));
    parts.push(h('p', { class: 'fine' }, 'Return ' + (rk.yearly ? 'a year ' : '') + 'since ' + fmtDay(rk.row.t0) + ', with deposits taken out of the picture. ' + ordinal(rk.pos) + ' of ' + rk.of + ' counting you and the stock and bond funds.'));
  }
  parts.push(more);
  fill(box, ...parts);
}

/* ---------- Performance ---------- */
const PERF_PRESETS = [
  { k: 'ytd', label: 'YTD', name: 'This year so far' },
  { k: '1y', label: '1Y', name: 'Past year' },
  { k: '3y', label: '3Y', name: 'Past 3 years' },
  { k: '5y', label: '5Y', name: 'Past 5 years' },
  { k: '10y', label: '10Y', name: 'Past 10 years' },
  { k: 'all', label: 'All', name: 'Since your first saved value' }
];
const diffCell = (you, sp) => (you == null || sp == null ? h('td', null, '—') : h('td', { class: tone(you - sp) }, ptsDiff(you - sp)));
const pctCell = (v, cls) => h('td', cls ? { class: cls } : null, v == null ? '—' : pct(v));
/* the picks with data over your saved values, in their places: [{ b, cls, row }] */
function pickRows() {
  const B = lastB;
  return cmpSlots().filter(Boolean).map(b => ({ b, cls: slotCls(b.id), row: B.byId.get(b.id) })).filter(x => x.row);
}

function renderPerformance(M) {
  const empty = $('#perf-empty'), body = $('#perf-body'), top = $('#perf-top'), P = lastP;
  $('#benchmark').hidden = mode === 'welcome';
  $('#cmp-bar').hidden = mode === 'welcome';
  if (mode === 'welcome') {
    empty.hidden = false; body.hidden = true; top.hidden = true;
    $('#perf-sub').textContent = '';
    fill(empty, h('h3', null, 'No file open'), h('p', null, 'Open your Excel file to see how your investments have performed.'),
      h('div', { class: 'btn-row' }, h('button', { class: 'btn btn-primary', type: 'button', 'data-action': 'open-file' }, 'Open Excel file')));
    return;
  }
  renderCmpBar();
  renderBenchmark(M);
  top.hidden = !M.count;
  renderBoard(M);
  if (!P) {
    empty.hidden = false; body.hidden = true;
    $('#health').hidden = true;
    $('#perf-sub').textContent = '';
    fill(empty,
      h('h3', null, 'Performance needs your value on more than one date'),
      h('p', null, 'Each time you update your value, the app keeps the earlier one. With two or more saved values it can measure how your investments did between them, apart from what you paid in.'),
      mode === 'example' ? null : h('div', { class: 'btn-row' }, h('button', { class: 'btn btn-primary', type: 'button', 'data-action': 'open-import', 'data-target': 'values' }, 'Add many values')),
      h('p', { class: 'note' }, 'To fill in the past in one go, paste dates and values from Excel or from a broker’s export.'));
    return;
  }
  empty.hidden = true; body.hidden = false;
  const H = computeHealth();
  $('#perf-sub').textContent = 'How your investments did, whatever you paid in and when. Measured between your ' + P.pts.length + ' saved values since ' + fmtDay(P.t0) + '.';
  renderHealth(H);
  const picks = pickRows(), periods = lastB.byId.get(lastB.main.id).periods;
  renderPerfChart(P, periods, picks);
  renderPerfTable(periods, picks);
  renderPerfRisk(H);
  renderPerfYears(H.years, picks);
  renderPerfStats(H.stats);
  renderPerfMonths(H.months);
}

/* What you're comparing with, as chips; tapping opens the list to choose from. */
function renderCmpBar() {
  const main = mainBench(), waiting = !BENCH_DATA.series.length;
  fill($('#cmp-bar'), h('span', { class: 'cmp-label' }, 'Against'),
    ...cmpSlots().filter(Boolean).map(b => h('button', { type: 'button', class: 'cmp-chip', 'data-action': 'open-picker', 'aria-label': 'Compared with ' + b.name + (b.id === main.id ? ', the main one' : '') + '. Change' },
      h('i', { class: 'dot-key key-' + slotCls(b.id) }), b.short || b.name, b.id === main.id && cmpSlots().filter(Boolean).length > 1 ? h('span', { class: 'star', 'aria-hidden': 'true' }, '★') : null)),
    h('button', { type: 'button', class: 'cmp-chip add', 'data-action': 'open-picker' }, waiting ? (PWA ? 'More on the way' : 'More in the installed app') : 'Change'));
}

function renderHealth(H) {
  const sec = $('#health'), play = plays('health');
  if (!H.verdict) {
    sec.hidden = !H.checks.length;
    if (!H.checks.length) return;
  }
  sec.hidden = false;
  const v = H.verdict;
  fill($('#health-top'), v ? h('div', { class: 'score ' + v.tone }, scoreRing(v.k, v.n, v.tone, play)) : null,
    h('div', null, h('h4', { class: 'verdict-title' }, v ? v.title : 'A first reading'),
      h('p', { class: 'verdict-line' }, marketLine(H) + (v ? ' ' + v.k + ' of ' + v.n + ' checks passed.' : ' More checks appear as your saved values build up.'))));
  const list = $('#health-list');
  list.classList.toggle('play', play);
  fill(list, ...H.checks.map((c, i) => { const t = checkText(c); return h('li', { class: c.pass ? 'yes' : 'no', style: { '--i': String(i) } }, markOf(c.pass), h('div', null, h('p', { class: 'check-title' }, t.title, h('span', { class: 'sr-only' }, c.pass ? ': yes. ' : ': no. ')), h('p', { class: 'check-detail' }, t.detail))); }));
  $('#health-fine').textContent = 'A reading of your saved values against public indexes and funds, with ' + theMain() + ' as the main yardstick. It describes the past and isn’t advice.';
}

/* Everything on one ladder: by what the same deposits would be worth, or by return over a period. */
function renderBoard(M) {
  const B = lastB, P = lastP, list = $('#board-list'), play = plays('board') || plays('health');
  if (!P) board.mode = 'money';
  $$('#board-seg button').forEach(b => { b.setAttribute('aria-checked', String(b.dataset.k === board.mode)); b.disabled = b.dataset.k === 'rate' && !P; });
  slideThumb($('#board-seg'));
  const pres = $('#board-presets'), lede = $('#board-lede'), fine = $('#board-fine');
  let rows = [], fmt, note = '';
  if (board.mode === 'money') {
    pres.hidden = true; fill(pres);
    for (const r of B.rows) if (r.what.value != null) rows.push({ b: r.b, v: r.what.value, peer: isPeer(r.b) });
    if (M.hasValue) rows.push({ you: true, v: M.V, peer: true });
    fmt = v => money(v, { dec: 0 });
    for (const x of rows) if (!x.you && M.hasValue) { const d = M.V - x.v; x.note = (d >= 0 ? 'You’re ahead by ' : 'You’re behind by ') + money(Math.abs(d), { dec: 0 }); x.noteTone = d >= 0 ? 'good' : 'bad'; }
    fill(lede, 'What your deposits would be worth ' + (M.hasValue ? 'on ' + fmtDay(M.vDay) : 'today') + ' if each one had gone into these instead, on the same day.');
    note = 'Counts every deposit from its own date, so it answers “would I have more money?”. Dividends are reinvested. ';
  } else {
    const mine = B.byId.get(B.main.id).periods;
    if (!mine.some(r => r.k === board.period)) board.period = 'all';
    pres.hidden = false;
    fill(pres, ...PERF_PRESETS.filter(p => mine.some(r => r.k === p.k)).map(p => h('button', { type: 'button', 'data-action': 'board-period', 'data-k': p.k, 'aria-label': p.name, 'aria-pressed': String(board.period === p.k) }, p.label)));
    slideThumb(pres);
    const rk = rankIn(B, board.period);
    rows = rk.list.slice();
    fmt = v => pct(v) + (rk.yearly ? ' a year' : '');
    const you = rows.find(x => x.you);
    for (const x of rows) if (!x.you) { const d = you.v - x.v; x.note = Math.abs(d) < 0.0005 ? 'Level with you' : (d > 0 ? 'You’re ahead by ' : 'You’re behind by ') + num(Math.abs(d) * 100, 1) + ' pts'; x.noteTone = d >= 0 ? 'good' : 'bad'; }
    fill(lede, 'Return ' + (rk.yearly ? 'a year ' : '') + 'from ' + fmtDay(rk.row.t0) + ' to ' + fmtDay(rk.row.t1) + ', measured between your saved values so deposits don’t count as performance.');
    note = 'A like-for-like test of the investments themselves. ';
  }
  rows.sort((a, b) => b.v - a.v);
  let pos = 0;
  for (const x of rows) x.pos = x.peer ? ++pos : null;
  for (const x of rows) if (x.b && slotOf(x.b.id) >= 0 && cmpSlots()[slotOf(x.b.id)]) x.dot = slotCls(x.b.id);
  const you = rows.find(x => x.you), peers = rows.filter(x => x.peer);
  if (you && peers.length > 1) you.note = ordinal(you.pos) + ' of ' + peers.length + (board.mode === 'money' ? ' by value' : ' by return');
  const sc = boardScale(rows.map(x => x.v));
  list.classList.toggle('play', play);
  /* the field first, in order; then the yardsticks that aren't given a place */
  const o = { fmt, left: sc.left, width: sc.width }, ranked = rows.filter(x => x.peer !== false), aside = rows.filter(x => x.peer === false);
  fill(list, ...ranked.map((x, i) => boardRow(x, i, o)),
    aside.length ? h('li', { class: 'board-sep', style: { '--i': String(ranked.length) } }, 'For scale') : null,
    ...aside.map((x, i) => boardRow(x, ranked.length + 1 + i, o)));
  const waiting = !BENCH_DATA.series.length;
  fill(fine, note + (rows.some(x => x.peer === false) ? 'Cash, inflation, gold and bitcoin are shown for scale and aren’t given a place. ' : ''),
    waiting ? 'Only the S&P 500 is here so far: more indexes and funds arrive with the app’s next daily data update.' : null);
}

function renderPerfChart(P, periods, picks) {
  if (!periods.some(r => r.k === perf.preset)) perf.preset = 'all';
  const row = periods.find(r => r.k === perf.preset);
  const pres = $('#perf-presets');
  fill(pres, ...PERF_PRESETS.filter(p => periods.some(r => r.k === p.k)).map(p =>
    h('button', { type: 'button', 'data-action': 'perf-preset', 'data-k': p.k, 'aria-label': p.name, 'aria-pressed': String(perf.preset === p.k) }, p.label)));
  slideThumb(pres);
  const box = $('#perf-chart'), i0 = row.i;
  const drawn = thinValues(P.pts.slice(i0), Math.max(24, Math.floor((box.clientWidth || 600) / 6)));
  const base = P.pts[i0];
  const you = drawn.map(p => [p.t, p.idx / base.idx - 1]);
  /* each pick's index at the same saved values; one that starts later begins where its data does */
  const lines = picks.map(x => {
    const at = new Map(x.row.P.pts.map(p => [p.t, p.sp])), b0 = at.get(base.t);
    const first = drawn.find(p => at.get(p.t) != null);
    const ref = b0 != null ? { t: base.t, v: b0, lift: 0 } : first ? { t: first.t, v: at.get(first.t), lift: first.idx / base.idx - 1 } : null;
    const pts = ref ? drawn.filter(p => p.t >= ref.t && at.get(p.t) != null).map(p => [p.t, (1 + ref.lift) * (at.get(p.t) / ref.v) - 1]) : [];
    const per = x.row.periods.find(r => r.k === perf.preset);
    return { b: x.b, cls: x.cls, pts, by: new Map(pts), total: b0 != null && per ? per.sp : null };
  }).filter(l => l.pts.length > 1);
  fill($('#perf-legend'), legendItem('value', 'You', pct(row.you)), ...lines.map(l => legendItem(l.cls, l.b.short || l.b.name, l.total != null ? pct(l.total) : null)));
  lineChart(box, {
    height: isWide() ? 290 : 226, x0: base.t, x1: P.t1, xTicks: timeTicks(base.t, P.t1, box.clientWidth || 600), animate: plays('perf'),
    series: lines.map(l => ({ cls: l.cls, type: 'line', thin: true, end: true, pts: l.pts })).concat([{ cls: 'value', type: 'line', end: true, dots: you.length <= 14, pts: you }]),
    yFmt: v => pct(v, { dec: 0 }), samples: drawn.map(p => p.t),
    readout: i => ({ title: fmtDay(drawn[i].t), rows: [{ cls: 'value', value: pct(you[i][1]), label: 'You' }].concat(lines.filter(l => l.by.has(drawn[i].t)).map(l => ({ cls: l.cls, value: pct(l.by.get(drawn[i].t)), label: l.b.short || l.b.name }))) }),
    focus: i => [{ cls: 'value', x: you[i][0], y: you[i][1] }].concat(lines.filter(l => l.by.has(drawn[i].t)).map(l => ({ cls: l.cls, x: drawn[i].t, y: l.by.get(drawn[i].t) }))),
    label: 'Line chart: your return since ' + fmtDay(base.t) + (lines.length ? ' against ' + lines.map(l => l.b.name).join(', ') : '')
  });
  $('#perf-fine').textContent = 'Total return since ' + fmtDay(base.t) + ', dividends reinvested. Between your saved values the lines run straight, because what happened in between isn’t known.' +
    (lines.some(l => l.total == null) ? ' A line that starts later begins level with you on the day its data starts.' : '') +
    (picks.length ? ' In the tables, a filled dot marks where you came out ahead and a ring where you were behind.' : '');
}

/* Periods down the side, you and each pick across. With one pick there is room for the gap as well. */
function matrixHead(el, first, picks) {
  const one = picks.length === 1;
  fill(el, h('tr', null, h('th', { scope: 'col' }, first), h('th', { scope: 'col' }, h('i', { class: 'dot-key key-value' }), 'You'),
    ...picks.map(x => h('th', { scope: 'col', class: 'vs' }, h('i', { class: 'dot-key key-' + x.cls }), x.b.short || x.b.name)), one ? h('th', { scope: 'col' }, 'Difference') : null));
}
function renderPerfTable(periods, picks) {
  const tight = picks.length > 1 && !isWide();
  const label = r => {
    if (r.k === 'ytd') return tight ? 'This year' : dayYear(r.t1) + ' so far';
    if (r.k === 'all') return 'All time';
    return tight ? (r.n === 1 ? '1 year' : r.n + ' years') : r.n === 1 ? 'Past year' : 'Past ' + r.n + ' years';
  };
  const show = v => (tight && Math.abs(v) >= 0.9995 ? pct(v, { dec: 0 }) : pct(v));
  matrixHead($('#perf-thead'), 'Period', picks);
  const one = picks.length === 1, rows = [];
  const line = (r, yearly, sub) => {
    const youV = yearly ? r.youAnn : r.you;
    const cells = picks.map(x => { const pr = x.row.periods.find(q => q.k === r.k), v = pr ? (yearly ? pr.spAnn : pr.sp) : null; return { v, td: h('td', { class: v == null ? null : 'vs ' + (youV >= v ? 'won' : 'lost') }, v == null ? '—' : show(v)) }; });
    rows.push(h('tr', null, h('td', null, label(r), sub ? h('small', { class: 'sub' }, sub) : null), h('td', { class: 'me' }, youV == null ? '—' : show(youV)), ...cells.map(c => c.td), one ? diffCell(youV, cells[0].v) : null));
  };
  for (const r of periods) {
    const yearly = r.days >= 365 && r.k !== '1y';
    line(r, yearly, yearly ? 'a year' : null);
    if (r.k === 'all' && yearly) line(r, false, 'in total');
  }
  fill($('#perf-tbody'), ...rows);
}

function renderPerfRisk(H) {
  const sec = $('#perf-risk'), k = H.risk;
  sec.hidden = !k;
  if (!k) return;
  const nm = mainName(), tiles = [
    { label: 'Swings', value: [pct(k.volYou, { sign: false }), h('small', null, ' a year')], cap: nm + ': ' + pct(k.volBm, { sign: false }) + '. How far monthly returns spread out' },
    { label: 'Return for the risk', value: num(k.ratioYou, 2), tone: k.ratioYou >= k.ratioBm ? 'good' : 'bad', cap: nm + ': ' + num(k.ratioBm, 2) + '. Yearly return' + (k.annCash != null ? ' above cash' : '') + ' for each point of swing' }
  ];
  if (k.beta != null) tiles.push({ label: 'Moves with ' + nm, value: [num(k.beta, 2), h('small', null, '×')], cap: 'When it moves 1%, you tend to move ' + num(k.beta, 2) + '%. In step ' + pct(Math.max(0, k.corr), { sign: false, dec: 0 }) + ' of the way' });
  if (k.up != null && k.down != null) tiles.push({ label: 'Its good and bad months', value: [pct(k.up, { sign: false, dec: 0 }), h('small', null, ' up, '), pct(k.down, { sign: false, dec: 0 }), h('small', null, ' down')],
    tone: k.up >= k.down ? 'good' : '', cap: 'Your average month as a share of its average, in its ' + k.nUp + ' rising and ' + k.nDown + ' falling months' });
  fill($('#perf-risk-tiles'), ...kpiTiles(tiles));
  fitTiles($('#perf-risk-tiles'));
  $('#perf-risk-fine').textContent = 'From the ' + k.n + ' full months you have a value at each end of.';
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
    const bar = (v, x, cls) => { if (v != null && Math.abs(Y(v) - y0) >= 0.5) grp.append(sv('path', { d: roundedBar(x, bw, y0, Y(v), 3), class: 'bar f-' + cls + (v < 0 ? ' neg' : '') })); };
    bar(g.a, cx - gap / 2 - bw, 'value');
    bar(g.b, cx + gap / 2, o.clsB || 's0');
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
  if (o.animate) growBars(svg);
  box.tabIndex = 0;
  box.setAttribute('aria-label', o.label + '. Use the left and right arrow keys to read values.');
  box.onkeydown = e => {
    if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') { e.preventDefault(); show(curI < 0 ? n - 1 : curI + (e.key === 'ArrowRight' ? 1 : -1)); }
    else if (e.key === 'Escape') hide();
  };
  box.onblur = hide;
}

function renderPerfYears(years, picks) {
  const sec = $('#perf-years'), shown = years.filter(r => !r.missing);
  sec.hidden = !shown.length;
  if (!shown.length) return;
  const main = mainBench(), mcls = slotCls(main.id), mname = main.short || main.name;
  const name = r => r.year + (r.ytd ? ' so far' : r.partial ? ' from ' + dfShort.format(new Date(r.t0 * DAY_MS)) : '');
  fill($('#perf-years-legend'), legendItem('value', 'You'), legendItem(mcls, mname));
  pairBars($('#perf-years-chart'), {
    height: isWide() ? 240 : 200, yFmt: v => pct(v, { dec: 0 }), clsB: mcls, animate: plays('perf'),
    groups: shown.map(r => ({
      label: String(r.year), short: '’' + String(r.year).slice(2), a: r.you, b: r.sp, faded: r.ytd || r.partial, title: name(r),
      rows: [{ cls: 'value', value: pct(r.you), label: 'You' }].concat(r.sp != null ? [{ cls: mcls, value: pct(r.sp), label: mname }] : [])
    })),
    label: 'Bar chart: your return and ' + main.name + '’s for each calendar year'
  });
  matrixHead($('#perf-years-thead'), 'Year', picks);
  const one = picks.length === 1;
  const others = picks.map(x => new Map(perfYears(x.row.P).filter(r => !r.missing).map(r => [r.year, r.sp])));
  fill($('#perf-years-tbody'), ...years.slice().reverse().map(r => r.missing
    ? h('tr', null, h('td', null, String(r.year)), h('td', { colspan: String(1 + picks.length + (one ? 1 : 0)), class: 'gap' }, 'No saved value near a year end'))
    : h('tr', { class: r.ytd ? 'current' : null }, h('td', null, r.ytd ? r.year + ' so far' : String(r.year), r.partial && !r.ytd ? h('small', { class: 'sub' }, 'from ' + dfShort.format(new Date(r.t0 * DAY_MS))) : null),
      h('td', { class: 'me' }, pct(r.you)), ...others.map(m => { const v = m.get(r.year); return h('td', { class: v == null ? null : 'vs ' + (r.you >= v ? 'won' : 'lost') }, v == null ? '—' : pct(v)); }),
      one ? diffCell(r.you, others[0].get(r.year)) : null)));
}

function renderPerfStats(st) {
  const tiles = [];
  const mo = r => fmtMonth(r.y, r.m);
  if (st.n >= 3) {
    tiles.push({ label: 'Best month', value: pct(st.best.you), tone: tone(st.best.you), cap: mo(st.best) + (st.best.sp != null ? '; ' + theMain() + ' did ' + pct(st.best.sp) : '') });
    tiles.push({ label: 'Worst month', value: pct(st.worst.you), tone: tone(st.worst.you), cap: mo(st.worst) + (st.worst.sp != null ? '; ' + theMain() + ' did ' + pct(st.worst.sp) : '') });
  }
  tiles.push(st.dd
    ? { label: 'Biggest drop', value: pct(st.dd.depth), tone: 'bad', cap: fmtDay(st.dd.peakT) + ' to ' + fmtDay(st.dd.troughT) + (st.dd.recT ? ', made back by ' + fmtDay(st.dd.recT) : ', not made back yet') }
    : { label: 'Biggest drop', value: 'None', cap: 'No fall between your saved values so far' });
  if (st.n >= 3) tiles.push({ label: 'Months that went up', value: st.up + ' of ' + st.n, cap: st.cmp ? 'Ahead of ' + theMain() + ' in ' + st.ahead + ' of ' + st.cmp : 'Full months with a value at each end' });
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
    fill(read, B(fmtMonth(r.y, r.m) + (r.mtd ? ' so far' : '')), ': you ' + pct(r.you) + (r.sp != null ? ', ' + mainName() + ' ' + pct(r.sp) : '') + '.');
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
        'aria-label': fmtMonth(r.y, r.m) + (r.mtd ? ' so far' : '') + ': ' + pct(r.you), style: { '--a': (0.1 + 0.5 * strength).toFixed(3), '--i': String((y1 - y) * 2 + m) } },
        (p < 0 && Math.round(Math.abs(p) * (wide ? 10 : 1)) !== 0 ? MINUS : '') + txt);
      cell.addEventListener('click', () => say(r));
      cell.addEventListener('pointerenter', e => { if (e.pointerType === 'mouse') say(r); });
      cell.addEventListener('focus', () => say(r));
      row.append(cell);
    }
    rows.push(row);
  }
  box.classList.toggle('play', plays('perf'));
  fill(box, ...rows);
  say(months[months.length - 1]);
  fine.textContent = 'Each cell is your return for that month, in percent. Tap one to compare it with ' + theMain() + '. A month needs a saved value near both of its ends.';
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
  const nOther = benchList().filter(b => b.avail && b.id !== 'SPX' && b.id !== 'MIX').length;
  fill($('#set-spx'), 'S&P 500 total return (dividends reinvested) from 1985 to the close of ' + fmtDate(SPX.last.d) + ' at ' + num(SPX.last.px, 2) + (nOther ? ', and ' + nOther + ' more indexes and funds' : '') + '. ' +
    (PWA ? 'A daily job on the site this app is served from adds new prices, and the app picks them up whenever it opens online. Your phone never contacts a price service itself. '
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
    state.settings = normalize({ settings: r.settings }).settings;
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
  const spx = sameDeposits(M, benchById('SPX').A);
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
    [S('S&P 500, same deposits'), spx.value != null ? S(round2(spx.value), 'money') : ''],
    [S('Difference vs S&P 500'), spx.diff != null ? S(round2(spx.diff), 'money') : ''],
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
    [S('Compare with'), st.settings.cmp.map(id => id || '-').join(',')],
    [S('Main comparison'), st.settings.main || ''],
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
function openDialog(d) {
  if (d._closing) { d._closing.cancel(); d._closing = null; d.classList.remove('closing'); }
  d.style.transform = '';
  if (typeof d.showModal === 'function') { if (!d.open) d.showModal(); } else d.setAttribute('open', '');
}
/* Sheets leave the way they came: down off the bottom on a phone, fading on a wide screen. `now` skips the motion. */
function closeDialog(d, value, now) {
  if (typeof d.close !== 'function') { d.removeAttribute('open'); return; }
  if (!d.open) return;
  if (now || REDUCED || !d.classList.contains('sheet') || d._closing) { if (d._closing) return; d.close(value); return; }
  const phone = window.innerWidth <= 640, from = d.style.transform || 'translateY(0)';
  d.classList.add('closing');
  const a = anim(d, phone ? [{ transform: from }, { transform: 'translateY(104%)' }] : [{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'translateY(10px) scale(.98)' }],
    { duration: phone ? 240 : 160, easing: 'cubic-bezier(.4,0,.9,.6)', fill: 'forwards' });
  const done = () => { d._closing = null; d.classList.remove('closing'); d.style.transform = ''; if (a) a.cancel(); if (d.open) d.close(value); };
  if (!a) { done(); return; }
  d._closing = a;
  a.onfinish = done;
}
for (const d of $$('dialog.sheet')) {
  d.addEventListener('cancel', e => { e.preventDefault(); closeDialog(d); });
  d.addEventListener('click', e => { if (e.target === d) closeDialog(d); });
  const form = d.querySelector('form');
  if (form) form.addEventListener('submit', e => { if (e.submitter && e.submitter.value === 'cancel') { e.preventDefault(); closeDialog(d, 'cancel'); } });
  /* pull the sheet down by its top edge to put it away */
  const head = d.querySelector('.dlg-head');
  let drag = null;
  head.addEventListener('touchstart', e => { if (e.touches.length === 1 && window.innerWidth <= 640) drag = { y0: e.touches[0].clientY, t0: e.timeStamp, dy: 0 }; }, { passive: true });
  head.addEventListener('touchmove', e => {
    if (!drag) return;
    drag.dy = Math.max(0, e.touches[0].clientY - drag.y0);
    d.style.transition = 'none';
    d.style.transform = 'translateY(' + drag.dy + 'px)';
  }, { passive: true });
  const end = e => {
    if (!drag) return;
    const fast = drag.dy / Math.max(1, e.timeStamp - drag.t0) > 0.55;
    d.style.transition = '';
    if (drag.dy > 110 || (fast && drag.dy > 30)) closeDialog(d);
    else d.style.transform = '';
    drag = null;
  };
  head.addEventListener('touchend', end);
  head.addEventListener('touchcancel', end);
}
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
/* what to say when no column reads as dates: a file can't have left its dates behind, a paste can */
function noDateMsg(open) {
  if (open) return 'I couldn’t find a column of dates on this sheet. Pick another sheet above, if there is one.';
  const many = imp.tab === 'file' && imp.src && imp.src.kind === 'sheets' && imp.src.sheets.length > 1;
  return imp.tab === 'file'
    ? 'I couldn’t find a column of dates in this file' + (many ? ' on this sheet' : '') + '. Dates can look like 27/12/2023, 2023-12-27 or 20231227.'
    : 'I couldn’t find a column of dates. Check that the date column came along when you copied.';
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
    const msg = { empty: 'Nothing to read yet.', nodate: noDateMsg(imp.mode === 'open'), noamount: 'I found dates but no column of amounts next to them.', broker: 'This looks like a broker export, but it has no deposits section to read. To add your portfolio values from it, use Add many values on the Overview tab.' }[r.problem];
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
    pv.append(h('p', { class: 'field-error' }, { empty: 'Nothing to read yet.', nodate: noDateMsg(false), noamount: 'I found dates but no column of values next to them.',
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

/* ---------- choosing what to compare with ----------
   Up to three at a time. Each keeps its place, and so its colour, until it is taken off; one of them is the
   main one, the yardstick for the health check and the headline comparison. */
const dlgPicker = $('#dlg-picker');
let pickChanged = false;
function pickSay(msg) {
  const el = $('#pick-note');
  el.textContent = msg || 'Choose up to three. The starred one is your main yardstick: the health check and the headline comparison use it.';
  el.classList.toggle('warn', !!msg);
  if (msg) anim(el, [{ transform: 'translateX(0)' }, { transform: 'translateX(-5px)' }, { transform: 'translateX(4px)' }, { transform: 'translateX(0)' }], { duration: 260, fill: 'none', easing: 'ease-out' });
}
function renderPicker() {
  const list = benchList(), slots = cmpSlots(), on = slots.filter(Boolean), main = mainBench();
  const first = lastM && lastM.first != null ? lastM.first : null;
  const groups = [];
  for (const g of GROUPS) {
    const inG = list.filter(b => b.group === g);
    if (!inG.length) continue;
    groups.push(h('div', { class: 'pick-group', role: 'group', 'aria-label': g }, h('p', { class: 'pick-head' }, g),
      ...inG.map(b => {
        const picked = on.includes(b), isMain = picked && b.id === main.id;
        const tag = !b.avail ? (PWA ? 'On the way' : 'Installed app') : first != null && b.startT > first + 31 ? 'From ' + dfMonthYear.format(new Date(b.startT * DAY_MS)) : null;
        const box = h('span', { class: 'pick-box' + (picked ? ' f-' + slotCls(b.id) : ''), 'aria-hidden': 'true' });
        if (picked) box.innerHTML = ICON_CHECK;
        return h('div', { class: 'pick-row' + (picked ? ' on' : '') + (b.avail ? '' : ' off') },
          h('button', { type: 'button', class: 'pick-btn', role: 'checkbox', 'aria-checked': String(picked), 'aria-disabled': b.avail ? null : 'true', 'data-action': 'pick-toggle', 'data-id': b.id },
            box, h('span', { class: 'pick-text' }, h('b', null, b.name, tag ? h('i', { class: 'pick-tag' }, tag) : null), h('small', null, b.note + (b.priceOnly ? '. Without dividends' : '')))),
          picked && on.length > 1 ? h('button', { type: 'button', class: 'pick-star' + (isMain ? ' on' : ''), 'data-action': 'pick-main', 'data-id': b.id, 'aria-pressed': String(isMain), 'aria-label': isMain ? b.name + ' is the main comparison' : 'Make ' + b.name + ' the main comparison' },
            h('span', { 'aria-hidden': 'true' }, '★'), isMain ? 'Main' : null) : null);
      })));
  }
  fill($('#pick-list'), ...groups);
  const waiting = list.some(b => !b.avail);
  $('#pick-fine').textContent = 'Funds stand in for the indexes they follow, with dividends reinvested, in US dollars. ' +
    (!waiting ? '' : PWA ? 'The ones marked “On the way” arrive with the app’s daily data update. ' : 'This preview carries the S&P 500 only: the others come with the installed app’s daily data update. ') + 'The app itself never contacts a price service.';
}
function togglePick(id) {
  const b = benchById(id);
  if (!b) return;
  if (!b.avail) { pickSay(PWA ? b.name + ' hasn’t arrived yet. It comes with the app’s daily data update.' : b.name + ' comes with the installed app’s daily data update.'); return; }
  const slots = cmpSlots(), i = slots.findIndex(x => x && x.id === id), n = slots.filter(Boolean).length;
  if (i >= 0 && n === 1) { pickSay('Keep at least one to compare with.'); return; }
  if (i < 0 && n >= 3) { pickSay('Three at a time. Take one off to add ' + b.name + '.'); return; }
  pickChanged = true; pickSay('');
  commitSettings(s => {
    const next = slots.map(x => (x ? x.id : null));
    if (i >= 0) next[i] = null; else next[next.indexOf(null)] = id;
    s.settings.cmp = next;
    if (!next.includes(s.settings.main)) s.settings.main = next.find(Boolean);
  });
}
/* the charts behind the sheet redraw with the new lines once it is out of the way */
dlgPicker.addEventListener('close', () => { if (pickChanged) { pickChanged = false; fresh.add('tab'); render(); } });

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
function resetView() { editing = null; valueOpen = false; histOpen = false; histAll = false; spxOpen = false; openYears.clear(); resetZoom(); perf.preset = 'all'; board.mode = 'money'; board.period = 'all'; for (const k in scrollPos) delete scrollPos[k]; }
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
  'show-example': () => { state = makeExample(todayDay()); mode = 'example'; noFile(); resetView(); syncPlanFromState(); setTab('overview'); window.scrollTo(0, 0); },
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
  'zoom-preset': btn => { resetZoom(); zoom.preset = btn.dataset.k; drawGrowth(!REDUCED); },
  'perf-preset': btn => { perf.preset = btn.dataset.k; if (lastP && lastB) replay('perf', () => renderPerfChart(lastP, lastB.byId.get(lastB.main.id).periods, pickRows())); },
  'board-mode': btn => { if (btn.dataset.k === board.mode) return; board.mode = btn.dataset.k; replay('board', () => renderBoard(lastM)); },
  'board-period': btn => { board.period = btn.dataset.k; replay('board', () => renderBoard(lastM)); },
  'go-performance': () => setTab('performance'),
  'open-picker': () => { if (mode === 'welcome') return; pickChanged = false; pickSay(''); renderPicker(); openDialog(dlgPicker); },
  'pick-toggle': btn => togglePick(btn.dataset.id),
  'pick-main': btn => {
    const id = btn.dataset.id;
    if (mainBench().id === id || slotOf(id) < 0) return;
    pickChanged = true; pickSay('');
    commitSettings(s => { s.settings.main = id; });
  },
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
  slideThumb($('#kind-dep').parentElement);
}
$('#kind-dep').addEventListener('click', () => setKind('dep'));
$('#kind-wd').addEventListener('click', () => setKind('wd'));
$('#kind-dep').parentElement.addEventListener('keydown', e => {
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
  fresh.add('growth'); fresh.add('health');
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
