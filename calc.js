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
  M.avgPerYear = M.years > 0 ? M.total / M.years : null;
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
  if ((m = /^(\d{4})(\d{2})(\d{2})$/.exec(s))) return { kind: 'iso', y: +m[1], m: +m[2], d: +m[3] };
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

function analyzeRows(rows, opts) {
  opts = opts || {};
  rows = (rows || []).filter(r => Array.isArray(r) && r.some(c => !isBlank(c)));
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
  if (dateCol < 0) {
    let best = 0;
    stats.forEach((s, c) => { if (s.dates > best && s.dates >= s.ne * 0.5) { best = s.dates; dateCol = c; } });
  }
  if (dateCol < 0) {
    stats.forEach((s, c) => { if (dateCol < 0 && s.serial >= 1 && s.serial >= s.ne * 0.6 && s.serial >= s.nums * 0.8) dateCol = c; });
  }
  const serialMode = dateCol >= 0 && stats[dateCol].dates === 0 && stats[dateCol].serial > 0;
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
    for (let c = 0; c < ncol; c++) {
      if (c === dateCol) continue;
      let n = 0;
      for (const r of dateRows) {
        const v = r[c];
        if (isBlank(v)) continue;
        const x = typeof v === 'number' ? v : (dateParts(v) ? NaN : parseAmount(v, null));
        if (isFinite(x)) n++;
      }
      const dist = Math.abs(c - dateCol);
      if (n > best || (n === best && n > 0 && dist < bestDist)) { best = n; bestDist = dist; amountCol = c; }
    }
    if (best < Math.max(1, dateRows.length * 0.3)) amountCol = -1;
  }
  res.dateCol = dateCol; res.amountCol = amountCol; res.serialMode = serialMode;
  if (dateCol < 0) { res.problem = 'nodate'; return res; }
  if (amountCol < 0) { res.problem = 'noamount'; return res; }

  let anyA = false, anyB = false, anyNum = false;
  for (const r of rows) {
    const p = dateParts(r[dateCol]);
    if (p && p.kind === 'num') { anyNum = true; if (p.a > 12) anyA = true; if (p.b > 12) anyB = true; }
  }
  const detected = anyA && !anyB ? 'dmy' : anyB && !anyA ? 'mdy' : null;
  const fmt = opts.fmt === 'dmy' || opts.fmt === 'mdy' ? opts.fmt : (detected || opts.localeFmt || 'dmy');
  res.fmt = fmt; res.fmtDetected = detected; res.fmtAmbiguous = anyNum && !detected; res.fmtConflict = anyA && anyB; res.hasNumDates = anyNum;

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
    if (!best || r.items.length > best.items.length) best = r;
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
if (typeof module !== 'undefined') module.exports = { ymdToIso, isoToDay, dayToIso, dayYear, buildAnchors, trAt, xirr, computeMetrics, growthSeries, projectSeries, recurringDates, dateParts, resolveDate, parseAmount, guessDecimalComma, numberCore, analyzeRows, parseTextTable, serialToIso, isAppWorkbook, readAppWorkbook, findCurrentValue, bestDepositSheet, FORMAT_MARK };
