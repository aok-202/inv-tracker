#!/usr/bin/env python3
"""Keeps the app's market data current: spx-data.js (the S&P 500) and bench-data.js (the other
indexes and funds you can compare against).

Runs on GitHub's servers through .github/workflows/update-sp500.yml, never on your phone.
It reads public prices only: FRED (the St. Louis Fed's free data service) and Yahoo Finance.
It never sees your Excel file or anything about you, and it needs no passwords or keys.

S&P 500: new closes are chained onto the stored total-return index with dividends added at the
yearly rate stored in spx-data.js (div). Month ends are recorded once a month is complete.

Everything else: month-end levels of a total-return index (100 at the first month stored), built
from Yahoo Finance's dividend-adjusted closes, or from FRED for inflation and cash. When FRED can't
be reached, cash comes from Yahoo Finance's 13-week Treasury bill rate (^IRX) and inflation from the
Bureau of Labor Statistics' public data service (the same index FRED republishes). Months already
stored are never rewritten; new months are chained onto the last one stored. If a source can't be
reached, that series is simply left as it is until the next run, and the reason is written into
bench-data.js ("problems"), so it can be read in the repository without opening the job's log.

To compare against another index or fund, add a line to BENCHMARKS below.

  python3 update_spx.py                          fetch and update both files
  python3 update_spx.py --csv FILE [--source N]  S&P 500 from a FRED-style CSV (date,close); nothing else is fetched
  python3 update_spx.py --bench-dir DIR          read the other series from files in DIR instead of the internet
  python3 update_spx.py --bench-only             leave spx-data.js alone
"""
import calendar, csv, datetime as dt, io, json, re, signal, sys, time, urllib.error, urllib.parse, urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
# next to this script (the usual place), or one folder up if it lives in scripts/
ROOT = HERE if (HERE / 'spx-data.js').exists() else HERE.parent
DATA = ROOT / 'spx-data.js'
BENCH = ROOT / 'bench-data.js'
FRED = 'https://fred.stlouisfed.org/graph/fredgraph.csv?id='
YAHOO = 'https://query1.finance.yahoo.com/v8/finance/chart/'
HEADERS = {'User-Agent': 'Mozilla/5.0 (compatible; investment-tracker data refresh)'}
BROWSER = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
           'Accept': 'text/csv,text/plain,*/*'}
BLS = 'https://api.bls.gov/publicAPI/v1/timeseries/data/'
STATE = {'fred_down': None}   # why FRED failed earlier in this run, if it did, so it isn't waited on again
EASTERN = dt.timezone(dt.timedelta(hours=-5))

# What the app can be compared against, besides the S&P 500: (id, Yahoo Finance symbol, name).
# Funds stand in for their indexes because their prices include dividends, fees and all.
BENCHMARKS = [
    ('QQQ', 'QQQ', 'Nasdaq-100'),
    ('VTI', 'VTI', 'US total market'),
    ('DIA', 'DIA', 'Dow Jones'),
    ('IWM', 'IWM', 'US small companies'),
    ('VT', 'VT', 'World stocks'),
    ('VEA', 'VEA', 'Developed markets outside the US'),
    ('VWO', 'VWO', 'Emerging markets'),
    ('AGG', 'AGG', 'US bonds'),
    ('GLD', 'GLD', 'Gold'),
    ('BTC', 'BTC-USD', 'Bitcoin'),
]
# From FRED: (id, FRED series, name, kind). 'level' is an index to follow; 'yield' is a yearly rate to compound.
FRED_SERIES = [
    ('CASH', 'DGS3MO', 'Cash (3-month US Treasury bills)', 'yield'),
    ('CPI', 'CPIAUCNS', 'US inflation', 'level'),
]
# Where to turn when FRED can't be reached: the same measure from another public source.
FALLBACKS = {
    'CASH': ('yahoo-yield', '^IRX'),          # 13-week Treasury bill rate, in percent
    'CPI': ('bls', 'CUUR0000SA0'),            # CPI-U, all items, not seasonally adjusted: FRED's CPIAUCNS
}
TIME_BUDGET = 170   # seconds for everything in bench-data.js
JOB_LIMIT = 240     # GitHub stops the whole job at 5 minutes (getting the files and saving them included)
STARTED = time.time()


class OutOfTime(Exception):
    pass


def within(until, fn, /, *args, **kw):
    """Runs fn, stopped at the deadline whatever the network does (a server that never answers, or answers a
    byte at a time, included). So a stuck price service can never make the job overrun and lose the S&P 500 update."""
    left = until - time.time()
    if left < 1:
        raise OutOfTime('out of time')

    def ring(signum, frame):
        raise OutOfTime('out of time')
    try:
        before = signal.signal(signal.SIGALRM, ring)
    except (AttributeError, ValueError):   # no alarm clock here (Windows, or not the main thread): timeouts only
        return fn(*args, **kw)
    signal.setitimer(signal.ITIMER_REAL, left)
    try:
        return fn(*args, **kw)
    finally:
        signal.setitimer(signal.ITIMER_REAL, 0)
        signal.signal(signal.SIGALRM, before)


def fetch(url, timeout=30, headers=None):
    with urllib.request.urlopen(urllib.request.Request(url, headers=headers or HEADERS), timeout=timeout) as r:
        return r.read().decode('utf-8')


# Yahoo Finance answers "too many requests" (429) when asked too fast or too often from one address, so the
# fund prices are asked for slowly, from either of its two addresses, with a wait and another try after a refusal.
YAHOO_HOSTS = ['https://query1.finance.yahoo.com', 'https://query2.finance.yahoo.com']
# The same plain request as the S&P 500 one, which Yahoo answers from GitHub's servers. (Dressing it up as a
# browser is what bot filters look for, so it isn't.)
YAHOO_HEADERS = HEADERS


def yahoo_get(path, deadline=None):
    err = None
    for wait in (0, 8, 20):
        if wait:
            if deadline and time.time() + wait > deadline:
                break
            time.sleep(wait)
        for host in YAHOO_HOSTS:
            left = deadline - time.time() if deadline else 25
            if left < 1:
                raise OutOfTime('out of time')
            try:
                return fetch(host + '/v8/finance/chart/' + path, timeout=min(25, left), headers=YAHOO_HEADERS)
            except OutOfTime:
                raise
            except urllib.error.HTTPError as e:
                err = e
                if e.code not in (403, 429, 500, 502, 503, 504):
                    raise   # a wrong symbol or address won't get better by waiting
            except Exception as e:
                err = e
    raise err or ValueError('no answer')


def brief(e):
    """An error as a short, plain line that is safe to store in a data file."""
    text = re.sub(r"[^A-Za-z0-9 .,:;()'/_%^-]+", ' ', str(e) or type(e).__name__)   # no braces: the file's own markers
    return re.sub(r'\s+', ' ', text).strip()[:160] or type(e).__name__


def fred_csv(series, timeout=30):
    """One FRED series as CSV text. After a refusal it asks once more, as a browser would."""
    try:
        return fetch(FRED + series, timeout=timeout)
    except urllib.error.HTTPError as e:
        if e.code not in (403, 406, 429):
            raise
        return fetch(FRED + series, timeout=timeout, headers=BROWSER)


def parse_csv(text, digits=2):
    out = {}
    for row in csv.reader(io.StringIO(text)):
        if len(row) >= 2 and re.fullmatch(r'\d{4}-\d{2}-\d{2}', row[0].strip()) and row[1].strip() not in ('', '.'):
            try:
                out[row[0].strip()] = round(float(row[1]), digits)
            except ValueError:
                pass
    return out


def yahoo_closes():
    res = json.loads(fetch(YAHOO + '%5EGSPC?range=1mo&interval=1d'))['chart']['result'][0]
    now = dt.datetime.now(EASTERN)
    out = {}
    for t, c in zip(res['timestamp'], res['indicators']['quote'][0]['close']):
        if c is None:
            continue
        day = dt.datetime.fromtimestamp(t, EASTERN).date()
        if day == now.date() and now.hour < 17:
            continue  # today's bar isn't a final close yet
        out[day.isoformat()] = round(float(c), 2)
    return out


def month_end(index):
    """Month end for position `index` in the stored S&P list (0 = Dec 1984)."""
    k = 11 + index
    y, m = 1984 + k // 12, k % 12 + 1
    return dt.date(y, m, calendar.monthrange(y, m)[1])


def update_spx(argv):
    text = DATA.read_text(encoding='utf-8')
    found = re.search(r'const SPX = (\{.*?\});', text, re.S)
    if not found:
        sys.exit('spx-data.js not recognized; nothing changed.')
    spx = json.loads(found.group(1))

    if '--csv' in argv:
        closes = parse_csv(Path(argv[argv.index('--csv') + 1]).read_text(encoding='utf-8'))
        source = argv[argv.index('--source') + 1] if '--source' in argv else 'CSV file'
    else:
        closes, used = {}, []
        try:
            closes.update(parse_csv(fred_csv('SP500')))
            used.append('FRED')
        except Exception as e:
            STATE['fred_down'] = brief(e)
            print('FRED unavailable:', e)
        try:
            newest = max(closes) if closes else ''
            extra = {d: v for d, v in yahoo_closes().items() if d > newest}
            if extra:
                closes.update(extra)
                used.append('Yahoo Finance')
        except Exception as e:
            print('Yahoo Finance unavailable:', e)
        if not closes:
            print('No S&P 500 price source could be reached; spx-data.js left as it is.')
            return False
        source = ' + '.join(used)

    last = spx['last']
    days = sorted(d for d in closes if d > last['d'])
    if not days:
        print('S&P 500 already up to date through', last['d'])
        return True
    div = float(spx['div'])
    cur_d, cur_px, cur_tr = dt.date.fromisoformat(last['d']), float(last['px']), float(last['tr'])
    next_end = month_end(len(spx['tr']))
    months = 0
    for ds in days:
        d = dt.date.fromisoformat(ds)
        while d > next_end:  # that month is over: store its month-end value
            spx['tr'].append(round(cur_tr * (1 + div) ** ((next_end - cur_d).days / 365), 2))
            months += 1
            next_end = month_end(len(spx['tr']))
        px = closes[ds]
        cur_tr *= (px / cur_px) * (1 + div) ** ((d - cur_d).days / 365)
        cur_d, cur_px = d, px
    spx['last'] = {'d': cur_d.isoformat(), 'px': cur_px, 'tr': round(cur_tr, 2)}

    text = text[:found.start(1)] + json.dumps(spx, separators=(',', ':')) + text[found.end(1):]
    text = re.sub(r'Last close: [^*]*\*/', 'Last close: %s (%s). */' % (cur_d.isoformat(), source), text, count=1)
    DATA.write_text(text, encoding='utf-8')
    print('S&P 500: added %d closes and %d month ends. Data now runs to %s at %.2f (%s).' % (len(days), months, cur_d, cur_px, source))
    return True


# ---------------------------------------------------------------------------------------------
# The other indexes and funds
# ---------------------------------------------------------------------------------------------

def ym_add(ym, n):
    k = ym[0] * 12 + ym[1] - 1 + n
    return (k // 12, k % 12 + 1)


def ym_text(ym):
    return '%04d-%02d' % ym


def ym_parse(s):
    return (int(s[:4]), int(s[5:7]))


def month_last_day(ym):
    return dt.date(ym[0], ym[1], calendar.monthrange(ym[0], ym[1])[1])


def yahoo_daily(symbol, source_dir=None, recent=False, deadline=None):
    """Dividend-adjusted daily closes for one symbol: ([(date, level)], adjusted?). Raises if it can't be read.
    recent: the last two years are enough (the earlier months are already stored)."""
    q = urllib.parse.quote(symbol, safe='')
    if source_dir:
        texts = [(Path(source_dir) / (symbol + '.json')).read_text(encoding='utf-8')]
    else:
        texts = None
        tail = '&interval=1d&includeAdjustedClose=true'
        urls = [q + '?range=max' + tail, q + '?period1=0&period2=%d' % (int(time.time()) + 86400) + tail]
        if recent:
            urls.insert(0, q + '?range=2y' + tail)
    best, err = None, None
    for i in range(len(texts) if texts else len(urls)):
        try:
            res = json.loads(texts[i] if texts else yahoo_get(urls[i], deadline))['chart']['result'][0]
            stamps = res.get('timestamp') or []
            ind = res['indicators']
            adj = (ind.get('adjclose') or [{}])[0].get('adjclose')
            close = ind['quote'][0]['close']
            levels, adjusted = (adj, True) if adj and len(adj) == len(stamps) else (close, False)
            meta = res.get('meta') or {}
            offset = int(meta.get('gmtoffset') or 0)
            always_open = str(meta.get('instrumentType', '')).upper() == 'CRYPTOCURRENCY'
            now = dt.datetime.now(dt.timezone.utc).replace(tzinfo=None) + dt.timedelta(seconds=offset)
            rows = {}
            for t, v in zip(stamps, levels):
                if v is None or not v > 0:
                    continue
                day = (dt.datetime(1970, 1, 1) + dt.timedelta(seconds=int(t) + offset)).date()
                if day == now.date() and (always_open or now.hour < 17):
                    continue  # today's bar isn't a final close yet
                if day > now.date():
                    continue
                rows[day] = float(v)
            out = sorted(rows.items())
            if len(out) < 2:
                raise ValueError('fewer than two closes')
            # a full daily history is best; a coarser answer (weekly, monthly bars) still gives month ends
            gaps = sorted((b[0] - a[0]).days for a, b in zip(out, out[1:]))
            daily = gaps[len(gaps) // 2] <= 4
            if best is None or (daily and not best[2]):
                best = (out, adjusted, daily)
            if daily:
                break
        except OutOfTime:
            raise
        except urllib.error.HTTPError as e:
            if e.code in (403, 429):
                raise   # refused, even after waiting: asking another way won't help
            err = e
        except Exception as e:  # try the next address
            err = e
    if best is None:
        raise err or ValueError('no data')
    return best[0], best[1]


def months_from_daily(rows):
    """[(date, level)] -> (completed months [((y, m), level)], latest (date, level)).
    A month is complete once there is a close in a later month."""
    by = {}
    for day, v in rows:
        by[(day.year, day.month)] = v      # the last close of each month wins
    keys = sorted(by)
    months, prev = [], None
    for ym in keys[:-1]:
        while prev is not None and ym_add(prev, 1) < ym:   # a month with no closes at all: carry the level
            prev = ym_add(prev, 1)
            months.append((prev, months[-1][1]))
        months.append((ym, by[ym]))
        prev = ym
    return months, rows[-1]


def fred_levels(series, kind, source_dir=None, deadline=None):
    """A FRED series as daily-style rows [(date, level)] that months_from_daily understands.
    'level': the index itself, dated at each month's last day (FRED dates a month by its first day).
    'yield': a yearly rate in percent, compounded day by day into a level."""
    text = (Path(source_dir) / (series + '.csv')).read_text(encoding='utf-8') if source_dir else \
        fred_csv(series, timeout=max(1, min(30, deadline - time.time())) if deadline else 30)
    obs = sorted((dt.date.fromisoformat(d), v) for d, v in parse_csv(text, digits=6).items())
    obs = [(d, v) for d, v in obs if d >= dt.date(1984, 12, 1)]
    if len(obs) < 2:
        raise ValueError('too few observations')
    if kind == 'level':
        return monthly_rows([((d.year, d.month), float(v)) for d, v in obs])
    return compound(obs)


def monthly_rows(items):
    """[((y, m), level)] for whole months -> rows dated at each month's last day, plus a marker in the following
    month: every month listed is complete."""
    rows = [(month_last_day(ym), float(v)) for ym, v in sorted(items)]
    rows.append((rows[-1][0] + dt.timedelta(days=1), rows[-1][1]))
    return rows


def compound(obs):
    """[(date, yearly rate in percent)] -> [(date, level)]: each rate earned until the next date."""
    rows, level = [(obs[0][0], 100.0)], 100.0
    for (d0, y0), (d1, _) in zip(obs, obs[1:]):
        level *= (1 + y0 / 100.0) ** ((d1 - d0).days / 365.0)
        rows.append((d1, level))
    return rows


def yahoo_yield_levels(symbol, source_dir=None, recent=False, deadline=None):
    """A rate quoted in percent on Yahoo Finance (^IRX: 13-week Treasury bills), compounded day by day."""
    rows, _ = yahoo_daily(symbol, source_dir, recent=recent, deadline=deadline)
    obs = [(d, v) for d, v in rows if d >= dt.date(1984, 12, 1) and 0 < v < 30]
    if len(obs) < 2:
        raise ValueError('too few observations')
    return compound(obs)


def bls_levels(series, source_dir=None, recent=False, deadline=None):
    """US inflation from the Bureau of Labor Statistics' public data service, month by month.
    Without a key it answers ten years per request."""
    if source_dir:
        docs = [json.loads((Path(source_dir) / (series + '.json')).read_text(encoding='utf-8'))]
    else:
        docs, this_year = [], dt.date.today().year
        y0 = this_year - 2 if recent else 1985
        while y0 <= this_year:
            y1 = min(y0 + 9, this_year)
            left = deadline - time.time() if deadline else 30
            if left < 1:
                raise OutOfTime('out of time')
            body = json.dumps({'seriesid': [series], 'startyear': str(y0), 'endyear': str(y1)}).encode('utf-8')
            req = urllib.request.Request(BLS, data=body, headers=dict(HEADERS, **{'Content-Type': 'application/json'}))
            with urllib.request.urlopen(req, timeout=min(30, left)) as r:
                docs.append(json.loads(r.read().decode('utf-8')))
            y0 = y1 + 1
    months = {}
    for doc in docs:
        if doc.get('status') != 'REQUEST_SUCCEEDED':
            raise ValueError('BLS said: ' + ' '.join(str(m) for m in (doc.get('message') or [doc.get('status')])))
        for item in doc['Results']['series'][0]['data']:
            period = str(item.get('period', ''))
            if not re.fullmatch(r'M(0[1-9]|1[0-2])', period):
                continue   # M13 is the year's average
            try:
                months[(int(item['year']), int(period[1:]))] = float(item['value'])
            except (KeyError, TypeError, ValueError):
                continue   # '-': a month with no figure (October 2025 was never published)
    if len(months) < 2:
        raise ValueError('too few months')
    return monthly_rows(months.items())


def merge_series(old, months, latest, digits=4):
    """Chain new months onto what is stored. Returns (tr list, start 'YYYY-MM', last dict) or None if it can't be joined."""
    if not months:
        return None
    by = dict(months)
    if old and old.get('tr') and old.get('start'):
        start = ym_parse(old['start'])
        stored_last = ym_add(start, len(old['tr']) - 1)
        if stored_last not in by:
            return None    # the answer doesn't reach back to what's stored: leave it alone
        ref_level, ref_tr = by[stored_last], float(old['tr'][-1])
        tr = list(old['tr'])
        for ym, v in months:
            if ym > stored_last:
                tr.append(round(ref_tr * v / ref_level, digits))
    else:
        start = months[0][0]
        base = months[0][1]
        tr = [round(100.0 * v / base, digits) for _, v in months]
    end_ym, end_level = months[-1]
    # tr[-1] belongs to the last month that is complete in this answer or stored from before, whichever is later
    n_end = (end_ym[0] * 12 + end_ym[1]) - (start[0] * 12 + start[1])
    if n_end != len(tr) - 1:
        return None
    last = {'d': latest[0].isoformat(), 'tr': round(tr[-1] * latest[1] / end_level, digits)}
    if latest[0] <= month_last_day(end_ym):   # nothing newer than the last complete month yet
        last = {'d': month_last_day(end_ym).isoformat(), 'tr': tr[-1]}
    return tr, ym_text(start), last


def update_bench(argv):
    if not BENCH.exists():
        print('bench-data.js is not in the repository, so only the S&P 500 is kept current.')
        return
    source_dir = argv[argv.index('--bench-dir') + 1] if '--bench-dir' in argv else None
    text = BENCH.read_text(encoding='utf-8')
    found = re.search(r'const BENCH = (\{.*?\});', text, re.S)
    if not found:
        print('bench-data.js not recognized; left as it is.')
        return
    try:
        bench = json.loads(found.group(1))
    except ValueError:
        print('bench-data.js could not be read; left as it is.')
        return
    stored = {s.get('id'): s for s in bench.get('series', []) if isinstance(s, dict)}
    began, got, notes, refused, problems = time.time(), {}, [], False, {}
    deadline = min(began + TIME_BUDGET, STARTED + JOB_LIMIT)

    def keep(sid):
        if sid in stored:
            got[sid] = stored[sid]

    # FRED first: it is quick and dependable, so a slow or refusing Yahoo can't use up its time
    jobs = [(sid, name, 'fred', ser, kind) for sid, ser, name, kind in FRED_SERIES] + [(sid, name, 'yahoo', sym, None) for sid, sym, name in BENCHMARKS]
    for n, (sid, name, where, code, kind) in enumerate(jobs):
        if time.time() > deadline - 1:
            notes.append('%s: out of time, left for the next run' % sid)
            problems[sid] = 'out of time'
            keep(sid)
            continue
        if where == 'yahoo' and refused:
            notes.append('%s: left for the next run (Yahoo Finance is turning requests from here away for now)' % sid)
            problems[sid] = 'not asked: Yahoo Finance turned an earlier request away'
            keep(sid)
            continue
        try:
            if where == 'yahoo':
                rows, adjusted = within(deadline, yahoo_daily, code, source_dir, recent=sid in stored, deadline=deadline)
                src = 'Yahoo Finance'
            else:
                rows, adjusted, tried = None, True, []
                if STATE['fred_down'] and not source_dir:
                    tried.append('FRED: ' + STATE['fred_down'])
                else:
                    try:
                        rows, src = within(deadline, fred_levels, code, kind, source_dir, deadline), 'FRED'
                    except OutOfTime:
                        raise
                    except Exception as e:
                        tried.append('FRED: ' + brief(e))
                if rows is None and sid in FALLBACKS:
                    how, other = FALLBACKS[sid]
                    try:
                        if how == 'yahoo-yield':
                            rows, src = within(deadline, yahoo_yield_levels, other, source_dir, sid in stored, deadline), 'Yahoo Finance (' + other + ')'
                        else:
                            rows, src = within(deadline, bls_levels, other, source_dir, sid in stored, deadline), 'BLS'
                        code = other
                    except OutOfTime:
                        raise
                    except Exception as e:
                        tried.append(other + ': ' + brief(e))
                if rows is None:
                    raise ValueError('; '.join(tried))
            months, latest = months_from_daily(rows)
            merged = merge_series(stored.get(sid), months, latest)
            if merged is None:
                notes.append('%s: the new data did not join onto what is stored; left as it is' % sid)
                problems[sid] = 'new data did not join onto what is stored'
                keep(sid)
                continue
            tr, start, last = merged
            if kind == 'level':   # a monthly index: the newest figure is the last month end itself
                last = {'d': month_last_day(ym_add(ym_parse(start), len(tr) - 1)).isoformat(), 'tr': tr[-1]}
            entry = {'id': sid, 'code': code, 'name': name, 'start': start, 'tr': tr, 'last': last, 'src': src}
            if not adjusted:
                entry['priceOnly'] = True
            got[sid] = entry
            before = stored.get(sid)
            if before != entry:
                notes.append('%s: %d month ends from %s, latest %s' % (sid, len(tr), start, last['d']))
        except Exception as e:
            notes.append('%s: could not be updated (%s)' % (sid, e))
            problems[sid] = brief(e)
            keep(sid)
            if where == 'yahoo' and isinstance(e, urllib.error.HTTPError) and e.code in (403, 429):
                refused = True
        if where == 'yahoo' and not source_dir and n < len(jobs) - 1:
            time.sleep(1.5)   # one polite request at a time
    # the file keeps one order whatever order the work was done in; anything stored that is no longer listed stays
    order = [sid for sid, _, _ in BENCHMARKS] + [sid for sid, _, _, _ in FRED_SERIES]
    out = [got[sid] for sid in order if sid in got] + [s for sid, s in stored.items() if sid not in order]

    dates = [s['last']['d'] for s in out if s.get('last')]
    new = {'v': 1, 'updated': max(dates) if dates else None, 'series': out}
    if problems:   # what went wrong this run, so it can be seen in the repository without the job's log
        new['problems'] = problems
    for line in notes:
        print(' ', line)
    if new == bench:
        print('Other indexes and funds: nothing new.')
        return
    text = text[:found.start(1)] + json.dumps(new, separators=(',', ':')) + text[found.end(1):]
    BENCH.write_text(text, encoding='utf-8')
    print('Other indexes and funds: %d series stored, newest data %s.' % (len(out), new['updated']))


def main(argv):
    if '--bench-only' not in argv:
        update_spx(argv)
    if '--csv' not in argv or '--bench-dir' in argv:
        try:
            update_bench(argv)
        except Exception as e:   # never let the extras stop the S&P 500 update from being saved
            print('Other indexes and funds: skipped (%s).' % e)


if __name__ == '__main__':
    main(sys.argv[1:])
