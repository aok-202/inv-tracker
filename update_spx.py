#!/usr/bin/env python3
"""Adds the latest S&P 500 closes to spx-data.js.

Runs on GitHub's servers through .github/workflows/update-sp500.yml, never on your phone.
It reads public index prices only: FRED (the St. Louis Fed's free data service), plus
Yahoo Finance to fill in the newest day if FRED hasn't published it yet. It never sees
your Excel file or anything about you, and it needs no passwords or keys.

New closes are chained onto the stored total-return index with dividends added at the
yearly rate stored in spx-data.js (div). Month ends are recorded once a month is complete.

  python3 scripts/update_spx.py                         fetch and update
  python3 scripts/update_spx.py --csv FILE [--source N]  use a FRED-style CSV instead (date,close)
"""
import calendar, csv, datetime as dt, io, json, re, sys, urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
# next to this script (the usual place), or one folder up if it lives in scripts/
DATA = HERE / 'spx-data.js' if (HERE / 'spx-data.js').exists() else HERE.parent / 'spx-data.js'
FRED = 'https://fred.stlouisfed.org/graph/fredgraph.csv?id=SP500'
YAHOO = 'https://query1.finance.yahoo.com/v8/finance/chart/%5EGSPC?range=1mo&interval=1d'
HEADERS = {'User-Agent': 'Mozilla/5.0 (compatible; investment-tracker data refresh)'}
EASTERN = dt.timezone(dt.timedelta(hours=-5))


def fetch(url):
    with urllib.request.urlopen(urllib.request.Request(url, headers=HEADERS), timeout=30) as r:
        return r.read().decode('utf-8')


def parse_csv(text):
    out = {}
    for row in csv.reader(io.StringIO(text)):
        if len(row) >= 2 and re.fullmatch(r'\d{4}-\d{2}-\d{2}', row[0].strip()) and row[1].strip() not in ('', '.'):
            out[row[0].strip()] = round(float(row[1]), 2)
    return out


def yahoo_closes():
    res = json.loads(fetch(YAHOO))['chart']['result'][0]
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
    """Month end for position `index` in the stored list (0 = Dec 1984)."""
    k = 11 + index
    y, m = 1984 + k // 12, k % 12 + 1
    return dt.date(y, m, calendar.monthrange(y, m)[1])


def main(argv):
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
            closes.update(parse_csv(fetch(FRED)))
            used.append('FRED')
        except Exception as e:
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
            sys.exit('No price source could be reached; nothing changed.')
        source = ' + '.join(used)

    last = spx['last']
    days = sorted(d for d in closes if d > last['d'])
    if not days:
        print('Already up to date through', last['d'])
        return
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
    print('Added %d closes and %d month ends. Data now runs to %s at %.2f (%s).' % (len(days), months, cur_d, cur_px, source))


if __name__ == '__main__':
    main(sys.argv[1:])
