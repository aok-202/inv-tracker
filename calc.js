/* calc.js: dates, returns (CAGR, XIRR), the S&P 500 comparison, and reading pasted or
   uploaded tables. No network, no storage; plain functions only. */
const DAY_MS = 86400000;
const pad2 = n => String(n).padStart(2, '0');
const round2 = n => Math.round(n * 100) / 100;

function ymdToIso(y, m, d) {
  y = +y; m = +m; d = +d;
  if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return null;
  if (y < 1900 || y > 2200 || m < 1 || m > 12 || d < 1) return null;
  const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
  if (d > dim) return null;
  return y + '-' + pad2(m) + '-' + pad2(d);
}
function isoToDay(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
  if (!m) return NaN;
  return Math.round(Date.UTC(+m[1], +m[2] - 1, +m[3]) / DAY_MS);
}
function dayToIso(n) { return new Date(n * DAY_MS).toISOString().slice(0, 10); }
function dayYear(n) { return new Date(n * DAY_MS).getUTCFullYear(); }
function dayMonth(n) { return new Date(n * DAY_MS).getUTCMonth() + 1; }
function dayDate(n) { return new Date(n * DAY_MS).getUTCDate(); }

/* Anchor points for the S&P total-return index: month ends, the last close, and an optional newer close the user typed. */
function buildAnchors(spx, user) {
  const A = spx.tr.map((tr, i) => {
    const k = 11 + i, y = 1984 + Math.floor(k / 12), m = (k % 12) + 1;
    return { t: Math.round(Date.UTC(y, m, 0) / DAY_MS), tr };
  });
  const lastT = isoToDay(spx.last.d);
  A.push({ t: lastT, tr: spx.last.tr, px: spx.last.px, d: spx.last.d });
  if (user && user.px > 0 && isoToDay(user.d) > lastT) {
    const dt = isoToDay(user.d) - lastT;
    A.push({ t: isoToDay(user.d), tr: spx.last.tr * (user.px / spx.last.px) * Math.pow(1 + spx.div, dt / 365), px: user.px, d: user.d, user: true });
  }
  return A;
}
function trAt(A, t) {
  if (!(t >= A[0].t)) return null;
  const n = A.length;
  if (t >= A[n - 1].t) return A[n - 1].tr;
  let lo = 0, hi = n - 1;
  while (hi - lo > 1) { const mid = (lo + hi) >> 1; if (A[mid].t <= t) lo = mid; else hi = mid; }
  const a = A[lo], b = A[hi];
  return a.tr * Math.pow(b.tr / a.tr, (t - a.t) / (b.t - a.t));
}

/* XIRR, Excel convention (365-day years). flows: [{t: day number, a: amount}], money in is negative. */
function xirr(flows) {
  if (!flows || flows.length < 2) return null;
  let neg = false, pos = false, t0 = Infinity, t1 = -Infinity;
  for (const f of flows) { if (f.a < 0) neg = true; if (f.a > 0) pos = true; if (f.t < t0) t0 = f.t; if (f.t > t1) t1 = f.t; }
  if (!neg || !pos || t1 <= t0) return null;
  const ys = flows.map(f => (f.t - t0) / 365);
  const npv = r => { let s = 0; for (let i = 0; i < flows.length; i++) s += flows[i].a / Math.pow(1 + r, ys[i]); return s; };
  let lo = -0.99, hi = 10, flo = npv(lo), fhi = npv(hi);
  if (!isFinite(flo) || !isFinite(fhi)) return null;
  if (flo * fhi > 0) {
    let found = false;
    for (const cand of [50, 200, 1000]) { const fc = npv(cand); if (isFinite(fc) && flo * fc <= 0) { hi = cand; fhi = fc; found = true; break; } }
    if (!found) return null;
  }
  for (let i = 0; i < 300; i++) {
    const mid = (lo + hi) / 2, fm = npv(mid);
    if (fm === 0) return mid;
    if (flo * fm < 0) { hi = mid; fhi = fm; } else { lo = mid; flo = fm; }
    if (hi - lo < 1e-11) break;
  }
  return (lo + hi) / 2;
}

const sumOf = (arr, f) => arr.reduce((s, x) => s + f(x), 0);

function computeMetrics(st, A, today) {
  const deps = st.deposits
    .map(d => ({ id: d.id, d: d.d, t: isoToDay(d.d), a: d.a }))
    .filter(d => isFinite(d.t) && isFinite(d.a))
    .sort((x, y) => x.t - y.t || (x.id < y.id ? -1 : 1));
  const hasValue = !!(st.value && isFinite(st.value.a) && isFinite(isoToDay(st.value.d)));
  const vDay = hasValue ? isoToDay(st.value.d) : today;
  const V = hasValue ? st.value.a : null;
  const M = { deps, hasValue, V, vDay, today };
  M.count = deps.length;
  M.nDep = deps.filter(d => d.a > 0).length;
  M.nWd = M.count - M.nDep;
  M.total = sumOf(deps, d => d.a);
  M.first = deps.length ? deps[0].t : null;
  M.last = deps.length ? deps[deps.length - 1].t : null;
  const endDay = Math.max(today, vDay);
  M.years = M.first != null ? Math.max(0, (endDay - M.first) / 365.25) : null;
  M.largest = deps.reduce((m, d) => (d.a > (m ? m.a : 0) ? d : m), null);

  const counted = deps.filter(d => d.t <= vDay);
  M.counted = counted;
  M.after = deps.filter(d => d.t > vDay);
  M.netCounted = sumOf(counted, d => d.a);
  M.retYears = counted.length ? (vDay - counted[0].t) / 365.25 : null;
  M.gain = M.roi = M.cagr = M.xirr = null;
  if (hasValue && counted.length) {
    M.gain = V - M.netCounted;
    M.roi = M.netCounted > 0 ? M.gain / M.netCounted : null;
    if (M.netCounted > 0 && V > 0 && M.retYears >= 1 / 12) M.cagr = Math.pow(V / M.netCounted, 1 / M.retYears) - 1;
    if (M.retYears >= 1 / 12) {
      const flows = counted.map(d => ({ t: d.t, a: -d.a }));
      flows.push({ t: vDay, a: V });
      M.xirr = xirr(flows);
    }
  }

  const lastA = A[A.length - 1];
  M.sp = { lastDay: lastA.t, lastPx: lastA.px, lastD: lastA.d, user: !!lastA.user, value: null, xirr: null, diff: null };
  if (counted.length && counted[0].t >= A[0].t) {
    let units = 0;
    for (const d of counted) units += d.a / trAt(A, d.t);
    const endT = vDay;
    M.sp.endT = endT;
    M.sp.value = units * trAt(A, endT);
    M.sp.stale = endT > lastA.t + 3;
    M.sp.periodReturn = trAt(A, endT) / trAt(A, counted[0].t) - 1;
    if ((endT - counted[0].t) / 365.25 >= 1 / 12) {
      const flows = counted.map(d => ({ t: d.t, a: -d.a }));
      flows.push({ t: endT, a: M.sp.value });
      M.sp.xirr = xirr(flows);
    }
    if (hasValue) M.sp.diff = V - M.sp.value;
  } else if (counted.length) {
    M.sp.unavailable = 'early';
  }

  const cy = dayYear(today);
  const years = new Map();
  for (const d of deps) {
    const y = dayYear(d.t);
    let r = years.get(y);
    if (!r) years.set(y, (r = { year: y, sum: 0, n: 0 }));
    r.sum += d.a; r.n++;
  }
  M.byYear = [];
  if (deps.length) {
    const y0 = dayYear(deps[0].t), y1 = Math.max(dayYear(deps[deps.length - 1].t), cy);
    let cum = 0;
    for (let y = y0; y <= y1; y++) {
      const r = years.get(y) || { year: y, sum: 0, n: 0 };
      cum += r.sum;
      M.byYear.push({ year: y, sum: round2(r.sum), n: r.n, cum: round2(cum), ytd: y === cy });
    }
  }
  /* Average per year = the average of the calendar-year totals above (every year from your first
     deposit's year to this one, including any year with nothing in it). */
  M.avgPerYear = M.byYear.length ? M.total / M.byYear.length : null;
  M.avgYears = M.byYear.length ? { n: M.byYear.length, y0: M.byYear[0].year, y1: M.byYear[M.byYear.length - 1].year } : null;
  M.cy = cy;
  M.thisYear = years.get(cy) || { sum: 0, n: 0 };
  M.lastYear = years.get(cy - 1) || { sum: 0, n: 0 };
  return M;
}

/* Samples for the growth chart: every deposit, every saved value, every month end. */
function growthSeries(M, history, A) {
  const deps = M.deps;
  if (!deps.length) return null;
  const start = deps[0].t;
  const end = Math.max(M.hasValue ? M.vDay : M.today, deps[deps.length - 1].t);
  const days = new Set([start, end]);
  for (const d of deps) days.add(d.t);
  const snaps = (history || [])
    .map(x => ({ t: isoToDay(x.d), a: x.a }))
    .filter(x => isFinite(x.t) && isFinite(x.a) && x.t >= start && x.t <= end);
  for (const s of snaps) days.add(s.t);
  let y = dayYear(start), mo = dayMonth(start);
  for (;;) {
    const me = Math.round(Date.UTC(y, mo, 0) / DAY_MS);
    if (me > end) break;
    if (me >= start) days.add(me);
    mo++; if (mo > 12) { mo = 1; y++; }
  }
  const xs = [...days].sort((a, b) => a - b);
  const spOK = start >= A[0].t;
  const pts = [];
  let i = 0, cum = 0, units = 0;
  for (const t of xs) {
    while (i < deps.length && deps[i].t <= t) {
      cum += deps[i].a;
      if (spOK) units += deps[i].a / trAt(A, deps[i].t);
      i++;
    }
    pts.push({ t, dep: cum, sp: spOK ? units * trAt(A, t) : null });
  }
  const vmap = new Map();
  for (const s of snaps.sort((a, b) => a.t - b.t)) vmap.set(s.t, s.a);
  if (M.hasValue && M.vDay >= start) vmap.set(M.vDay, M.V);
  const vals = [...vmap.entries()].map(([t, a]) => ({ t, a })).sort((a, b) => a.t - b.t);
  return { pts, vals, start, end, spOK };
}

/* Your portfolio's value on day t, read off the straight line between the saved values around it. */
function valueAt(vals, t) {
  const n = vals.length;
  if (!n || t < vals[0].t || t > vals[n - 1].t) return null;
  if (n === 1) return vals[0].a;
  let lo = 0, hi = n - 1;
  while (hi - lo > 1) { const mid = (lo + hi) >> 1; if (vals[mid].t <= t) lo = mid; else hi = mid; }
  const p = vals[lo], q = vals[hi];
  return q.t === p.t ? q.a : p.a + ((q.a - p.a) * (t - p.t)) / (q.t - p.t);
}

/* Your portfolio against the S&P 500 over a window of days [a, b].
   It starts from a value that's known: nothing before your first deposit, or else your first
   saved value inside the window. It ends at the last saved value inside the window. The S&P 500
   side starts with the same money and gets the same deposits on the same days.
   Return is the gain divided by the money at work, counting each deposit from its date (Modified Dietz). */
function compareWindow(G, deps, A, a, b) {
  const fromStart = a <= G.start;
  const vin = G.vals.filter(v => v.t >= a && v.t <= b);
  const out = { fromStart, t0: null, V0: null };
  if (fromStart) { out.t0 = G.start; out.V0 = 0; }
  else if (vin.length) { out.t0 = vin[0].t; out.V0 = vin[0].a; }
  else { out.problem = 'novalue'; return out; }
  const tr0 = trAt(A, out.t0);
  if (tr0 == null) { out.problem = 'nosp'; out.t0 = null; return out; }
  const end = vin.length ? vin[vin.length - 1] : null;
  if (!end || end.t <= out.t0) { out.problem = vin.length ? 'onevalue' : 'novalue'; return out; }
  const t1 = end.t, span = t1 - out.t0;
  let units = out.V0 / tr0, net = 0, wsum = 0;
  for (const d of deps) {
    if (d.t > t1) break;
    if (fromStart ? d.t < out.t0 : d.t <= out.t0) continue;
    units += d.a / trAt(A, d.t);
    net += d.a;
    wsum += (d.a * (t1 - d.t)) / span;
  }
  const S1 = units * trAt(A, t1);
  Object.assign(out, { t1, V1: end.a, S1, net, gainV: end.a - out.V0 - net, gainS: S1 - out.V0 - net, diff: end.a - S1 });
  const base = out.V0 + wsum;
  out.retV = base > 0 ? out.gainV / base : null;
  out.retS = base > 0 ? out.gainS / base : null;
  return out;
}

/* What the growth chart draws for a window of days [a, b]. When the window starts after your first
   deposit, the S&P 500 line restarts from your value at the start of the comparison. */
function windowSeries(G, deps, A, a, b) {
  const cmp = compareWindow(G, deps, A, a, b);
  const rebased = !cmp.fromStart && cmp.t0 != null;
  const set = new Set([a, b]);
  for (const p of G.pts) if (p.t > a && p.t < b) set.add(p.t);
  for (const v of G.vals) if (v.t > a && v.t < b) set.add(v.t);
  const days = [...set].sort((x, y) => x - y);
  const pts = [];
  let i = 0, cum = 0, uAll = 0, uReb = rebased ? cmp.V0 / trAt(A, cmp.t0) : 0;
  for (const t of days) {
    while (i < deps.length && deps[i].t <= t) {
      const d = deps[i++];
      cum += d.a;
      if (G.spOK) uAll += d.a / trAt(A, d.t);
      if (rebased && d.t > cmp.t0) uReb += d.a / trAt(A, d.t);
    }
    const sp = rebased ? (t >= cmp.t0 ? uReb * trAt(A, t) : null) : G.spOK ? uAll * trAt(A, t) : null;
    pts.push({ t, dep: cum, sp });
  }
  const vin = G.vals.filter(v => v.t >= a && v.t <= b);
  const vline = vin.map(v => [v.t, v.a]);
  const va = valueAt(G.vals, a), vb = valueAt(G.vals, b);
  if (va != null && !(vin.length && vin[0].t === a)) vline.unshift([a, va]);
  if (vb != null && !(vin.length && vin[vin.length - 1].t === b)) vline.push([b, vb]);
  return { a, b, pts, vline, vin, cmp, rebased };
}

function projectSeries(v0, annual, rate, years) {
  const rm = Math.pow(1 + rate, 1 / 12) - 1, pm = annual / 12;
  let v = v0, base = v0;
  const pts = [{ y: 0, v, base }];
  for (let m = 1; m <= years * 12; m++) {
    v = v * (1 + rm) + pm;
    base += pm;
    if (m % 12 === 0) pts.push({ y: m / 12, v, base });
  }
  return pts;
}

function recurringDates(startIso, endIso, freq) {
  const s = isoToDay(startIso), e = isoToDay(endIso);
  if (!(s <= e)) return [];
  const out = [];
  if (freq === 'weekly' || freq === 'biweekly') {
    const step = freq === 'weekly' ? 7 : 14;
    for (let t = s; t <= e && out.length < 3000; t += step) out.push(dayToIso(t));
    return out;
  }
  const months = { monthly: 1, quarterly: 3, yearly: 12 }[freq] || 1;
  const [y, m, d] = startIso.split('-').map(Number);
  for (let k = 0; out.length < 3000; k++) {
    const mm = m - 1 + k * months, yy = y + Math.floor(mm / 12), mo = (mm % 12) + 1;
    const dim = new Date(Date.UTC(yy, mo, 0)).getUTCDate();
    const iso = ymdToIso(yy, mo, Math.min(d, dim));
    if (!iso || isoToDay(iso) > e) break;
    out.push(iso);
  }
  return out;
}

/* ---- reading pasted or uploaded tables ---- */
const MON3 = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
function monthFromName(s) { return MON3[String(s).toLowerCase().slice(0, 3)] || 0; }
function fixYear(ys) { let y = +ys; if (String(ys).length === 2) y += y < 70 ? 2000 : 1900; return y; }
function serialToIso(n) { if (!(n > 0 && n < 200000)) return null; return dayToIso(Math.floor(n) - 25569); }

function dateParts(raw) {
  if (raw instanceof Date) {
    if (isNaN(raw.getTime())) return null;
    const t = new Date(raw.getTime() + 43200000);
    return { kind: 'date', y: t.getFullYear(), m: t.getMonth() + 1, d: t.getDate() };
  }
  if (raw == null || typeof raw === 'number' || typeof raw === 'boolean') return null;
  let s = String(raw).trim();
  if (!s || s.length > 40) return null;
  s = s.replace(/^(mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?,?\s+/i, '');
  s = s.replace(/(?:[T\s]+|\s*,\s*)\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?\s*(?:[ap]\.?m\.?)?\s*(?:z|utc|gmt|[+-]\d{2}:?\d{2})?$/i, '');
  let m;
  if ((m = /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})\.?$/.exec(s))) return { kind: 'iso', y: +m[1], m: +m[2], d: +m[3] };
  /* 20231201, as broker exports write dates. Only a real date counts, so a large amount like 12345678 stays a number. */
  if ((m = /^(\d{4})(\d{2})(\d{2})$/.exec(s))) {
    const y = +m[1], mo = +m[2], d = +m[3];
    return y >= 1900 && y <= 2100 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31 ? { kind: 'iso', y, m: mo, d } : null;
  }
  if ((m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4}|\d{2})\.?$/.exec(s))) return { kind: 'num', a: +m[1], b: +m[2], y: fixYear(m[3]) };
  if ((m = /^(\d{1,2})(?:st|nd|rd|th)?[\s\-\/.]*([a-z]{3,10})\.?[\s\-\/.,]*(\d{4}|\d{2})$/i.exec(s))) {
    const mo = monthFromName(m[2]); if (mo) return { kind: 'name', y: fixYear(m[3]), m: mo, d: +m[1] };
  }
  if ((m = /^([a-z]{3,10})\.?[\s\-\/.]*(\d{1,2})(?:st|nd|rd|th)?,?[\s\-\/.,]+(\d{4}|\d{2})$/i.exec(s))) {
    const mo = monthFromName(m[1]); if (mo) return { kind: 'name', y: fixYear(m[3]), m: mo, d: +m[2] };
  }
  if ((m = /^([a-z]{3,10})\.?[\s\-\/.,]*(\d{4})$/i.exec(s))) {
    const mo = monthFromName(m[1]); if (mo) return { kind: 'name', y: +m[2], m: mo, d: 1 };
  }
  return null;
}
function resolveDate(p, fmt) {
  if (!p) return null;
  if (p.kind === 'num') return fmt === 'mdy' ? ymdToIso(p.y, p.a, p.b) : ymdToIso(p.y, p.b, p.a);
  return ymdToIso(p.y, p.m, p.d);
}
function numberCore(raw) {
  let s = String(raw).trim();
  if (!s) return null;
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1).trim(); }
  s = s.replace(/[−‒–—﹣－]/g, '-');
  const fd = s.search(/\d/);
  if (fd < 0) return null;
  let ld = -1;
  for (let i = s.length - 1; i >= 0; i--) if (s[i] >= '0' && s[i] <= '9') { ld = i; break; }
  const pre = s.slice(0, fd), post = s.slice(ld + 1);
  if (/[a-z]{4,}/i.test(pre + post)) return null;
  if (pre.includes('-') || post.includes('-')) neg = !neg;
  const core = s.slice(fd, ld + 1).replace(/[\s  '’]/g, '');
  if (!/^[\d.,]+$/.test(core)) return null;
  return { neg, core };
}
/* true = decimal comma, false = decimal point, null = can't tell from this value alone */
function guessDecimalComma(core) {
  const dots = (core.match(/\./g) || []).length, commas = (core.match(/,/g) || []).length;
  if (dots && commas) return core.lastIndexOf(',') > core.lastIndexOf('.');
  if (commas > 1) return false;
  if (dots > 1) return true;
  if (commas === 1) return core.split(',')[1].length === 3 ? null : true;
  if (dots === 1) return core.split('.')[1].length === 3 ? null : false;
  return null;
}
function parseAmount(raw, decimalComma) {
  if (typeof raw === 'number') return isFinite(raw) ? raw : NaN;
  if (raw == null || raw instanceof Date) return NaN;
  const nc = numberCore(raw);
  if (!nc) return NaN;
  let dc = decimalComma;
  if (dc == null) { const g = guessDecimalComma(nc.core); dc = g == null ? false : g; }
  let core = nc.core;
  if (dc) {
    if ((core.match(/,/g) || []).length > 1) return NaN;
    core = core.replace(/\./g, '').replace(',', '.');
  } else {
    core = core.replace(/,/g, '');
  }
  if ((core.match(/\./g) || []).length > 1 || core === '.') return NaN;
  const v = parseFloat(core);
  if (!isFinite(v)) return NaN;
  return nc.neg ? -v : v;
}
function csvSplit(line, delim) {
  const out = []; let cur = '', q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (q) {
      if (ch === '"') { if (line[i + 1] === '"') { cur += '"'; i++; } else q = false; }
      else cur += ch;
    } else if (ch === '"') q = true;
    else if (ch === delim) { out.push(cur.trim()); cur = ''; }
    else cur += ch;
  }
  out.push(cur.trim());
  return out;
}
function wsSplit(line) {
  const t = line.trim();
  const parts = t.split(/\s{2,}/);
  if (parts.length > 1) return parts;
  const m = /^(.*\S)\s+(\S+)$/.exec(t);
  return m ? [m[1], m[2]] : [t];
}
const isBlank = c => c == null || (typeof c === 'string' && c.trim() === '');

/* ---- broker exports in sections, like Interactive Brokers' CSV reports ----
   Every line starts with a section name and a row kind, then the fields:
     Allocation by Asset Class,Header,Date,Equities,Cash,NAV
     Allocation by Asset Class,Data,202312,9500.00,250.00,9750.00
   Only Header and Data rows are read; other kinds (MetaInfo, Total, Notes) are left alone. */
const monthEndIso = (y, m) => dayToIso(Math.round(Date.UTC(y, m, 0) / DAY_MS));
/* A period label made into a date: 202312 is the last day of December 2023, 2023Q4 is 31 Dec 2023,
   20231229 is that day. A month or quarter that hasn't ended yet is dated `cap` (the report's end, or today). */
function periodIso(raw, cap) {
  const s = typeof raw === 'number' ? String(raw) : String(raw == null ? '' : raw).trim();
  let m, iso = null;
  const okY = y => y >= 1900 && y <= 2100;
  if ((m = /^(\d{4})(\d{2})(\d{2})$/.exec(s)) && okY(+m[1])) return ymdToIso(+m[1], +m[2], +m[3]);
  if (((m = /^(\d{4})(\d{2})$/.exec(s)) || (m = /^(\d{4})[-\/.](\d{1,2})$/.exec(s))) && okY(+m[1]) && +m[2] >= 1 && +m[2] <= 12) iso = monthEndIso(+m[1], +m[2]);
  else if ((m = /^(\d{4})\s*[-\s]?\s*q([1-4])$/i.exec(s)) && okY(+m[1])) iso = monthEndIso(+m[1], +m[2] * 3);
  else if ((m = /^q([1-4])\s*[-\s]?\s*(\d{4})$/i.exec(s)) && okY(+m[2])) iso = monthEndIso(+m[2], +m[1] * 3);
  else return null;
  return cap && iso > cap ? cap : iso;
}
const NAV_NAMES = [/^(total )?nav( in base currency)?$/i, /^(net asset value|ending nav|end(ing)? value|total value|portfolio value|account value)$/i, /^(total|value|net liquidation( value)?|balance|market value)$/i];
const AMOUNT_NAMES = [/^(amount|net amount|amount in base( currency)?)$/i, /^(deposits?|deposits? ?(and|&) ?withdrawals?|net deposits?)$/i];
function brokerTable(rows, want, today) {
  const kindOf = r => (r.length >= 3 ? String(r[1] == null ? '' : r[1]).trim().toLowerCase() : '');
  let marked = 0;
  for (const r of rows) if (kindOf(r) === 'header' || kindOf(r) === 'data') marked++;
  if (marked < 2 || marked < rows.length * 0.5) return null;
  /* each Header row starts a table; Data rows belong to the latest table of their section */
  const tables = [], latest = new Map();
  for (const r of rows) {
    const k = kindOf(r), name = String(r[0] == null ? '' : r[0]).trim();
    if (k === 'header') { const t = { section: name, header: r.slice(2).map(c => (c == null ? '' : String(c).trim())), data: [] }; tables.push(t); latest.set(name, t); }
    else if (k === 'data') {
      let t = latest.get(name);
      if (!t) { t = { section: name, header: null, data: [] }; tables.push(t); latest.set(name, t); }
      t.data.push(r.slice(2));
    }
  }
  /* the report's end date, when the export says it ("January 1, 2023 - November 6, 2026") */
  let periodEnd = null;
  const rangeRe = /(\d{8}|\d{4}-\d{2}-\d{2}|[a-z]{3,9}\.? \d{1,2},? \d{4})\s*(?:-|–|—|to)\s*(\d{8}|\d{4}-\d{2}-\d{2}|[a-z]{3,9}\.? \d{1,2},? \d{4})/i;
  for (const r of rows) {
    for (const c of r) {
      const m = typeof c === 'string' ? rangeRe.exec(c) : null;
      if (!m) continue;
      const iso = resolveDate(dateParts(m[2]), 'mdy');
      if (iso && iso <= today && (!periodEnd || iso > periodEnd)) periodEnd = iso;
    }
  }
  const cap = periodEnd && periodEnd < today ? periodEnd : today;
  const names = want === 'values' ? NAV_NAMES : AMOUNT_NAMES;
  const prefer = want === 'values' ? /allocation by asset class/i : /deposit|withdraw/i;
  let best = null, bestScore = -1;
  for (const t of tables) {
    if (!t.data.length) continue;
    /* deposits only ever come from a deposits section, so dividends or fees are never read as deposits */
    if (want !== 'values' && !/deposit|withdraw|transfer/i.test(t.section)) continue;
    let dateCol = -1, valCol = -1, level = 9;
    if (t.header) {
      dateCol = t.header.findIndex(h => /^date$/i.test(h));
      if (dateCol < 0) dateCol = t.header.findIndex(h => /date|period|month/i.test(h));
      names.forEach((re, lv) => { if (level > lv) { const i = t.header.findIndex((h, j) => j !== dateCol && re.test(h)); if (i >= 0) { valCol = i; level = lv; } } });
      /* plain names like Total or Value only count in a section about holdings, not, say, returns */
      if (want === 'values' && level === 2 && !/allocation|nav|net asset|value|balance|account/i.test(t.section)) valCol = -1;
    } else if (want === 'values' && prefer.test(t.section)) {
      dateCol = 0; valCol = t.data.reduce((n, r) => Math.max(n, r.length), 0) - 1; level = 5;
    }
    if (dateCol < 0 || valCol < 0 || valCol === dateCol) continue;
    const score = (prefer.test(t.section) ? 1000 : 0) + (9 - level) * 100 + Math.min(99, t.data.length);
    if (score > bestScore) { bestScore = score; best = { t, dateCol, valCol }; }
  }
  const sections = [...new Set(tables.map(t => t.section))];
  if (!best) return { problem: 'broker', sections };
  let months = 0, quarters = 0, capped = 0;
  const out = best.t.data.map(r => {
    const row = r.slice(), raw = String(row[best.dateCol] == null ? '' : row[best.dateCol]).trim();
    const iso = periodIso(raw, cap);
    if (iso) {
      if (/q/i.test(raw)) quarters++; else if (!/^\d{8}$/.test(raw)) months++;
      if (iso === cap && periodIso(raw, null) !== cap) capped++;
      row[best.dateCol] = iso;
    }
    return row;
  });
  if (best.t.header) out.unshift(best.t.header);
  const kind = quarters > months ? 'quarter' : months ? 'month' : null;
  return { rows: out, dateCol: best.dateCol, valCol: best.valCol, section: best.t.section, sections, kind, capped, cap };
}

function analyzeRows(rows, opts) {
  opts = opts || {};
  rows = (rows || []).filter(r => Array.isArray(r) && r.some(c => !isBlank(c)));
  if (!opts.plain) {
    const today = opts.today || dayToIso(Math.floor(Date.now() / DAY_MS));
    const bt = brokerTable(rows, opts.want === 'values' ? 'values' : 'deposits', today);
    if (bt && bt.problem) {
      return { problem: 'broker', broker: { sections: bt.sections }, items: [], errors: [], skipped: 0, ncol: 0, rowsUsed: rows.length, header: null, dateCol: -1, amountCol: -1 };
    }
    if (bt) {
      const res = analyzeRows(bt.rows, Object.assign({}, opts, { plain: true, preferAmountCol: bt.valCol, preferDateCol: bt.dateCol }));
      res.broker = { section: bt.section, sections: bt.sections, kind: bt.kind, capped: bt.capped, cap: bt.cap };
      return res;
    }
  }
  const ncol = rows.reduce((n, r) => Math.max(n, r.length), 0);
  const res = { rowsUsed: rows.length, ncol, items: [], errors: [], skipped: 0, header: null, dateCol: -1, amountCol: -1, firstRow: rows[0] || null };
  if (!rows.length || !ncol) { res.problem = 'empty'; return res; }
  const stats = [];
  for (let c = 0; c < ncol; c++) {
    let ne = 0, dates = 0, nums = 0, serial = 0;
    for (const r of rows) {
      const v = r[c];
      if (isBlank(v)) continue;
      ne++;
      if (dateParts(v)) { dates++; continue; }
      const n = typeof v === 'number' ? v : parseAmount(v, null);
      if (isFinite(n)) { nums++; if (Number.isInteger(n) && n >= 25000 && n <= 60000) serial++; }
    }
    stats.push({ ne, dates, nums, serial });
  }
  res.stats = stats;
  let dateCol = Number.isInteger(opts.dateCol) && opts.dateCol >= 0 && opts.dateCol < ncol ? opts.dateCol : -1;
  if (dateCol < 0 && Number.isInteger(opts.preferDateCol) && opts.preferDateCol < ncol) {
    const s = stats[opts.preferDateCol];
    if (s && s.dates > 0 && s.dates >= s.ne * 0.5) dateCol = opts.preferDateCol;
  }
  if (dateCol < 0) {
    let best = 0;
    stats.forEach((s, c) => { if (s.dates > best && s.dates >= s.ne * 0.5) { best = s.dates; dateCol = c; } });
  }
  if (dateCol < 0) {
    stats.forEach((s, c) => { if (dateCol < 0 && s.serial >= 1 && s.serial >= s.ne * 0.6 && s.serial >= s.nums * 0.8) dateCol = c; });
  }
  const serialMode = dateCol >= 0 && stats[dateCol].dates === 0 && stats[dateCol].serial > 0;
  /* a heading row that names the column (NAV for values, Amount for deposits) decides between number columns */
  if (!Number.isInteger(opts.preferAmountCol) && dateCol >= 0 && typeof rows[0][dateCol] === 'string' && !dateParts(rows[0][dateCol])) {
    const hdr = rows[0].map(c => (c == null ? '' : String(c).trim()));
    const names = opts.want === 'values' ? NAV_NAMES : AMOUNT_NAMES;
    for (const re of names) { const i = hdr.findIndex((x, j) => j !== dateCol && re.test(x)); if (i >= 0) { opts = Object.assign({}, opts, { preferAmountCol: i }); break; } }
  }
  let amountCol = Number.isInteger(opts.amountCol) && opts.amountCol >= 0 && opts.amountCol < ncol && opts.amountCol !== dateCol ? opts.amountCol : -1;
  if (amountCol < 0 && dateCol >= 0) {
    /* Count numbers only on rows that have a date, so totals or notes beside the table don't win. Ties go to the column nearest the dates. */
    const dateRows = rows.filter(r => {
      const v = r[dateCol];
      if (!serialMode) return !!dateParts(v);
      const n = typeof v === 'number' ? v : parseAmount(v, false);
      return isFinite(n) && Number.isInteger(n) && n >= 25000 && n <= 60000;
    });
    let best = 0, bestDist = Infinity;
    const counts = [];
    for (let c = 0; c < ncol; c++) {
      if (c === dateCol) continue;
      let n = 0;
      for (const r of dateRows) {
        const v = r[c];
        if (isBlank(v)) continue;
        const x = typeof v === 'number' ? v : (dateParts(v) ? NaN : parseAmount(v, null));
        if (isFinite(x)) n++;
      }
      counts[c] = n;
      const dist = Math.abs(c - dateCol);
      if (n > best || (n === best && n > 0 && dist < bestDist)) { best = n; bestDist = dist; amountCol = c; }
    }
    /* a column named for what we want (NAV in a broker export) beats the nearest one */
    const pc = opts.preferAmountCol;
    if (Number.isInteger(pc) && pc !== dateCol && (counts[pc] || 0) >= Math.max(1, dateRows.length * 0.3)) { amountCol = pc; best = counts[pc]; }
    if (best < Math.max(1, dateRows.length * 0.3)) amountCol = -1;
  }
  res.dateCol = dateCol; res.amountCol = amountCol; res.serialMode = serialMode;
  if (dateCol < 0) { res.problem = 'nodate'; return res; }
  if (amountCol < 0) { res.problem = 'noamount'; return res; }

  /* Rows that only make sense one way round (31/01 is day first, 01/31 is month first) decide the order.
     A date that's impossible either way (31/13) is just an unreadable row. The larger side wins a disagreement. */
  let nA = 0, nB = 0, anyNum = false;
  for (const r of rows) {
    const p = dateParts(r[dateCol]);
    if (p && p.kind === 'num') { anyNum = true; if (p.a > 12 && p.b <= 12) nA++; else if (p.b > 12 && p.a <= 12) nB++; }
  }
  const detected = nA || nB ? (nA >= nB ? 'dmy' : 'mdy') : null;
  const fmt = opts.fmt === 'dmy' || opts.fmt === 'mdy' ? opts.fmt : (detected || opts.localeFmt || 'dmy');
  res.fmt = fmt; res.fmtDetected = detected; res.fmtAmbiguous = anyNum && !detected; res.fmtConflict = nA > 0 && nB > 0; res.hasNumDates = anyNum;

  let vc = 0, vd = 0;
  for (const r of rows) {
    const v = r[amountCol];
    if (typeof v !== 'string') continue;
    const nc = numberCore(v);
    if (!nc) continue;
    const g = guessDecimalComma(nc.core);
    if (g === true) vc++; else if (g === false && /[.,]/.test(nc.core)) vd++;
  }
  const dc = vc > vd ? true : vd > vc ? false : !!opts.localeDecimalComma;
  res.decimalComma = dc;

  rows.forEach((r, i) => {
    const dv = r[dateCol], av = r[amountCol];
    if (isBlank(dv) && isBlank(av)) return;
    let iso;
    if (serialMode) { const n = typeof dv === 'number' ? dv : parseAmount(dv, false); iso = isFinite(n) ? serialToIso(n) : null; }
    else iso = resolveDate(dateParts(dv), fmt);
    const amt = parseAmount(av, dc);
    if (!iso && !isFinite(amt)) {
      if (i === 0) res.header = r.map(c => (c == null ? '' : String(c).trim()));
      else res.skipped++;
      return;
    }
    if (!iso) {
      if (typeof dv === 'string' && !/\d/.test(dv)) { res.skipped++; return; }
      res.errors.push({ line: i + 1, reason: 'date', raw: [dv, av] }); return;
    }
    if (!isFinite(amt)) { res.errors.push({ line: i + 1, reason: 'amount', raw: [dv, av] }); return; }
    if (amt === 0) { res.skipped++; return; }
    res.items.push({ d: iso, a: round2(amt), line: i + 1 });
  });
  return res;
}

/* How imported portfolio values join the saved ones, one value per date.
   A dense list, like a broker's daily values, is thinned to the last value of each month or week
   (keep: 'month', 'week' or 'all'), always keeping its first and latest values. A dense list then
   stands for its whole period: saved values inside that period give way to it, so importing a
   longer export later replaces the earlier one instead of piling up extra points. */
function planValues(items, saved, keep) {
  const out = { neg: [], twice: [], found: 0, dense: false, kept: [], removed: [], range: null, nNew: 0, nRepl: 0, nSame: 0 };
  const lastIdx = new Map();
  items.forEach((it, i) => lastIdx.set(it.d, i));
  const valid = [];
  items.forEach((it, i) => {
    if (it.a < 0) out.neg.push(it);
    else if (lastIdx.get(it.d) !== i) out.twice.push(it);
    else valid.push(it);
  });
  valid.sort((x, y) => (x.d < y.d ? -1 : x.d > y.d ? 1 : 0));
  out.found = valid.length;
  if (valid.length >= 20) {
    const gaps = [];
    for (let i = 1; i < valid.length; i++) gaps.push(isoToDay(valid[i].d) - isoToDay(valid[i - 1].d));
    gaps.sort((x, y) => x - y);
    out.dense = gaps[gaps.length >> 1] <= 8;
  }
  let kept = valid;
  if (out.dense && keep !== 'all') {
    const key = keep === 'week' ? d => Math.floor((isoToDay(d) + 3) / 7) : d => d.slice(0, 7);
    const last = new Map();
    for (const v of valid) last.set(key(v.d), v);
    const pick = new Set(last.values());
    pick.add(valid[0]);
    pick.add(valid[valid.length - 1]);
    kept = valid.filter(v => pick.has(v));
  }
  if (out.dense) out.range = [valid[0].d, valid[valid.length - 1].d];
  const inRange = d => !!out.range && d >= out.range[0] && d <= out.range[1];
  const savedBy = new Map(saved.map(s => [s.d, s.a]));
  out.kept = kept.map(v => {
    const old = savedBy.get(v.d);
    const st = old == null ? 'new' : Math.abs(old - v.a) < 0.005 ? 'same' : 'repl';
    if (st === 'new') out.nNew++; else if (st === 'same') out.nSame++; else out.nRepl++;
    return Object.assign({}, v, { st, old });
  });
  const keptDates = new Set(kept.map(v => v.d));
  out.removed = saved.filter(s => inRange(s.d) && !keptDates.has(s.d));
  const byDate = new Map();
  for (const s of saved) if (!inRange(s.d)) byDate.set(s.d, s.a);
  for (const v of kept) byDate.set(v.d, v.a);
  out.history = [...byDate].map(([d, a]) => ({ d, a })).sort((x, y) => (x.d < y.d ? -1 : x.d > y.d ? 1 : 0));
  out.changes = out.nNew + out.nRepl + out.removed.length;
  return out;
}

/* Fewer points for drawing a long, dense value history: the last value of each week, month,
   quarter or year (the finest of those that fits in maxPts), plus the first and the last. */
function thinValues(vals, maxPts) {
  if (vals.length <= maxPts) return vals;
  const ym = t => { const d = new Date(t * DAY_MS); return [d.getUTCFullYear(), d.getUTCMonth()]; };
  const keys = [t => Math.floor((t + 3) / 7), t => { const [y, m] = ym(t); return y * 12 + m; },
    t => { const [y, m] = ym(t); return y * 4 + Math.floor(m / 3); }, t => ym(t)[0]];
  let out = vals;
  for (const key of keys) {
    const last = new Map();
    for (const v of vals) last.set(key(v.t), v);
    const pick = new Set(last.values());
    pick.add(vals[0]);
    pick.add(vals[vals.length - 1]);
    out = vals.filter(v => pick.has(v));
    if (out.length <= maxPts) break;
  }
  return out;
}

function parseTextTable(text, opts) {
  opts = opts || {};
  const lines = String(text || '').replace(/\r\n?/g, '\n').split('\n').filter(l => l.trim() !== '');
  if (!lines.length) return { problem: 'empty', items: [], errors: [], skipped: 0 };
  const cands = [];
  if (opts.delim) cands.push(opts.delim);
  else {
    if (lines.some(l => l.includes('\t'))) cands.push('\t');
    if (lines.some(l => l.includes(';'))) cands.push(';');
    if (lines.some(l => l.includes(','))) cands.push(',');
    cands.push('ws');
  }
  let best = null;
  for (const dl of cands) {
    const table = lines.map(l => (dl === 'ws' ? wsSplit(l) : csvSplit(l, dl)));
    const r = analyzeRows(table, opts);
    r.delim = dl; r.table = table;
    const rank = x => (x.broker ? 1e9 : 0) + x.items.length;
    if (!best || rank(r) > rank(best)) best = r;
    if (dl === '\t' && r.items.length >= lines.length * 0.8) break;
  }
  return best;
}

/* ---- workbooks: the tracker's own file format, and finding data in other people's sheets ---- */
const FORMAT_MARK = 'Investment Tracker 1';
const sheetByName = (wb, name) => wb.sheets.find(s => String(s.name).trim().toLowerCase() === name);

function isAppWorkbook(wb) {
  const set = sheetByName(wb, 'settings');
  if (!set) return false;
  return set.rows.some(r => String(r[0] == null ? '' : r[0]).trim().toLowerCase() === 'file format' && /^investment tracker/i.test(String(r[1] == null ? '' : r[1])));
}

/* Reads a workbook this app saved (Deposits, Values and Settings sheets). Rows edited by hand in Excel are read the same way as an import. */
function readAppWorkbook(wb, opts) {
  opts = opts || {};
  const out = { deposits: [], history: [], value: null, settings: { cur: 'USD', spx: null }, skipped: 0 };
  const base = { localeFmt: opts.localeFmt, localeDecimalComma: opts.localeDecimalComma, dateCol: 0, amountCol: 1 };
  const dep = sheetByName(wb, 'deposits');
  if (dep) {
    const r = analyzeRows(dep.rows, base);
    out.deposits = r.items.map(it => ({ d: it.d, a: it.a }));
    out.skipped += r.errors.length;
  }
  const val = sheetByName(wb, 'values');
  if (val) {
    const r = analyzeRows(val.rows, base);
    const byDate = new Map();
    for (const it of r.items) byDate.set(it.d, it.a);
    out.history = [...byDate.entries()].map(([d, a]) => ({ d, a })).sort((x, y) => (x.d < y.d ? -1 : 1));
    out.skipped += r.errors.length;
    if (out.history.length) out.value = Object.assign({}, out.history[out.history.length - 1]);
  }
  const set = sheetByName(wb, 'settings');
  if (set) {
    const kv = new Map();
    for (const row of set.rows) if (row[0] != null && row[0] !== '') kv.set(String(row[0]).trim().toLowerCase(), row[1]);
    const c = String(kv.get('currency') || '').trim().toUpperCase();
    if (/^[A-Z]{3}$/.test(c)) out.settings.cur = c;
    const sd = resolveDate(dateParts(kv.get('s&p 500 close date')), 'dmy');
    const sp = parseAmount(kv.get('s&p 500 close level'), null);
    if (sd && isFinite(sp) && sp > 0) out.settings.spx = { d: sd, px: sp };
  }
  return out;
}

/* Looks for a "current value" cell in someone's own workbook: a label such as
   "Current portfolio value" with a number to its right or just below it. */
function findCurrentValue(wb) {
  let best = null;
  const numAt = v => {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string' && /\d/.test(v) && !dateParts(v)) { const n = parseAmount(v, null); return isFinite(n) ? n : null; }
    return null;
  };
  for (const sh of wb.sheets) {
    sh.rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        if (typeof cell !== 'string' || cell.length > 60) return;
        const s = cell.toLowerCase();
        if (!/(value|worth|balance)/.test(s)) return;
        let score = 0;
        if (/current|today|latest|now/.test(s)) score += 3;
        if (/portfolio|account|investment/.test(s)) score += 2;
        if (/market|total/.test(s)) score += 1;
        if (/deposit|contribut|average|annual|year|roi|cagr|gain|profit|return/.test(s)) score -= 4;
        if (score <= 0) return;
        let n = null;
        for (let k = 1; k <= 3 && n == null; k++) n = numAt(row[ci + k]);
        for (let k = 1; k <= 2 && n == null; k++) { const below = sh.rows[ri + k]; n = below ? numAt(below[ci]) : null; }
        if (n != null && n > 0 && (!best || score > best.score)) best = { value: round2(n), label: cell.trim(), sheet: sh.name, score };
      });
    });
  }
  return best;
}

/* Picks the sheet whose rows look most like a list of deposits. */
function bestDepositSheet(wb, opts) {
  let best = -1, bestN = -1;
  wb.sheets.forEach((sh, i) => {
    const r = analyzeRows(sh.rows, opts);
    const n = r.items.length - (sh.hidden ? 0.5 : 0);
    if (n > bestN) { bestN = n; best = i; }
  });
  return best;
}
if (typeof module !== 'undefined') module.exports = { ymdToIso, isoToDay, dayToIso, dayYear, buildAnchors, trAt, xirr, computeMetrics, growthSeries, valueAt, compareWindow, windowSeries, planValues, thinValues, projectSeries, recurringDates, dateParts, resolveDate, parseAmount, guessDecimalComma, numberCore, analyzeRows, parseTextTable, serialToIso, isAppWorkbook, readAppWorkbook, findCurrentValue, bestDepositSheet, FORMAT_MARK };
