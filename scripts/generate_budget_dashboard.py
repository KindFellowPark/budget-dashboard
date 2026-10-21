#!/usr/bin/env python3
"""
Budget dashboard generator.

Reads three sources from one folder and writes a single HTML page:

  spending_tracker.xlsx     the office's own tracker (required). Two tabs:
                            "Account Summary" (the budget and committed figures
                            the office keeps by hand) and "Pipeline" (spending
                            that is planned but not yet in the finance system).
  budget_snapshot*.xlsx     finance-system export, one row per account, with the
                            adjusted budget (budget after mid-year transfers).
  transaction_detail.csv    finance-system export, one row per transaction line
  (or .xlsx)                (YTD = spent, ENC = encumbered, OBD = original budget).

Merge rules, in order of authority:
  Budget     snapshot adjusted budget, else the tracker's budget.
  Committed  transaction detail (YTD + ENC), else snapshot (YTD + ENC),
             else the tracker's committed figure.
Accounts below 4000 (payroll and fringe) are left out everywhere.

Checks the page reports:
  - accounts whose committed + planned spending exceeds the budget
  - accounts where the finance system's budget differs from the tracker
  - accounts where the finance system's committed figure differs from the
    tracker by more than $50
  - untracked activity: finance-system figures on account codes the tracker
    does not list

Exit codes: 0 written; 3 a required source is missing or ambiguous;
1 a source is present but cannot be parsed. Nothing is written on failure.
"""

import argparse
import csv
import html
import os
import sys
from datetime import date

import openpyxl

EXCLUDE_BELOW = 4000          # payroll and fringe accounts are not operating budget
MISMATCH_TOLERANCE = 50.00    # dollars; smaller differences are timing noise
WATCH_PCT, LIMIT_PCT = 75, 95

EXIT_OK, EXIT_PARSE, EXIT_MISSING = 0, 1, 3

SNAPSHOT_HEADERS = ['Account', 'Type', 'Title', 'Adj Budget', 'YTD Activity',
                    'Commitments', 'Available Balance']
TXN_COLUMNS = {'account': 0, 'type': 4, 'doc': 5, 'description': 6, 'fund': 8,
               'date': 11, 'field': 12, 'amount': 13}


class SourceMissing(Exception):
    exit_code = EXIT_MISSING


class SourceParseError(Exception):
    exit_code = EXIT_PARSE


# ── 1. Finding sources ─────────────────────────────────────────────────────

def find_by_prefix(folder, prefix, extensions):
    """Case-insensitive prefix match. The finance system controls export
    filenames and has renamed them before, so an exact name would silently
    drop a renamed export. More than one match is an error, never a guess."""
    if not os.path.isdir(folder):
        raise SourceMissing(f'source folder not found: {folder}')
    hits = sorted(f for f in os.listdir(folder)
                  if f.lower().startswith(prefix) and f.lower().endswith(extensions)
                  and not f.startswith('~$'))
    if len(hits) > 1:
        raise SourceMissing(f'{len(hits)} files match "{prefix}*" in {folder}: '
                            f'{", ".join(hits)}. Remove the stale one.')
    return os.path.join(folder, hits[0]) if hits else None


def locate_sources(folder):
    tracker = find_by_prefix(folder, 'spending_tracker', ('.xlsx',))
    snapshot = find_by_prefix(folder, 'budget_snapshot', ('.xlsx',))
    txns = find_by_prefix(folder, 'transaction_detail', ('.xlsx', '.csv'))
    if tracker is None:
        raise SourceMissing(f'spending_tracker.xlsx not found in {folder}')
    if snapshot is None and txns is None:
        raise SourceMissing(f'no finance-system export found in {folder} '
                            '(need budget_snapshot*.xlsx or transaction_detail.*)')
    return {'tracker': tracker, 'snapshot': snapshot, 'transactions': txns}


# ── 2. Loaders: rows in file order, no aggregation ─────────────────────────

def _account(value, where):
    """Account code as int. Returns None for total/blank rows; raises on junk."""
    s = str(value).strip().strip("'") if value is not None else ''
    if not s or s.lower().startswith('total'):
        return None
    try:
        return int(float(s))
    except ValueError:
        raise SourceParseError(f'{where}: account "{s}" is not a number')


def _amount(value, where):
    if value is None or value == '':
        return 0.0
    try:
        return float(str(value).replace(',', '').replace('$', '').strip())
    except ValueError:
        raise SourceParseError(f'{where}: amount "{value}" is not a number')


def load_snapshot(path):
    """Budget snapshot export. Rows 1-2 are query parameters, row 3 the
    headers, row 4 onward one row per account. Returns (rows, excluded_count)."""
    ws = openpyxl.load_workbook(path, data_only=True).active
    name = os.path.basename(path)
    header = [str(c or '').strip() for c in next(ws.iter_rows(min_row=3, max_row=3, values_only=True))]
    if header[:len(SNAPSHOT_HEADERS)] != SNAPSHOT_HEADERS:
        raise SourceParseError(f'{name}: row 3 headers changed; expected {SNAPSHOT_HEADERS}, got {header}')
    rows, excluded = [], 0
    for i, r in enumerate(ws.iter_rows(min_row=4, values_only=True), start=4):
        where = f'{name} row {i}'
        acct = _account(r[0], where)
        if acct is None:
            continue
        if acct < EXCLUDE_BELOW:
            excluded += 1
            continue
        rows.append({'account': acct, 'title': str(r[2] or '').strip(),
                     'adj_budget': round(_amount(r[3], where), 2),
                     'ytd': round(_amount(r[4], where), 2),
                     'enc': round(_amount(r[5], where), 2),
                     'available': round(_amount(r[6], where), 2)})
    return rows, excluded


def load_transactions(path):
    """Transaction detail export (CSV or XLSX). Rows 1-2 are metadata, row 3
    the headers. Amounts stay unrounded so totals add up the way the finance
    system adds them. Returns (rows, excluded_count)."""
    name = os.path.basename(path)
    if path.lower().endswith('.csv'):
        with open(path, newline='', encoding='utf-8-sig') as f:
            raw = list(csv.reader(f))
    else:
        ws = openpyxl.load_workbook(path, data_only=True).active
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
    if len(raw) < 3:
        raise SourceParseError(f'{name}: fewer than 3 rows; not a transaction export')
    header = [str(c or '').strip() for c in raw[2]]
    for key, want in (('account', 'Account'), ('field', 'Field'), ('amount', 'Amount')):
        col = TXN_COLUMNS[key]
        if len(header) <= col or header[col] != want:
            raise SourceParseError(f'{name}: expected "{want}" in column {col + 1} of row 3')
    rows, excluded = [], 0
    for i, r in enumerate(raw[3:], start=4):
        if not r or all(c in (None, '') for c in r):
            continue
        where = f'{name} row {i}'
        acct = _account(r[TXN_COLUMNS['account']], where)
        if acct is None:
            continue
        field = str(r[TXN_COLUMNS['field']] or '').strip().strip("'")
        if field not in ('YTD', 'ENC', 'OBD'):
            continue
        if acct < EXCLUDE_BELOW:
            excluded += 1
            continue
        cell = lambda k: str(r[TXN_COLUMNS[k]] or '').strip().strip("'")
        rows.append({'account': acct, 'field': field,
                     'amount': _amount(r[TXN_COLUMNS['amount']], where),
                     'type': cell('type'), 'doc': cell('doc'),
                     'description': cell('description'), 'fund': cell('fund'),
                     'date': cell('date')})
    return rows, excluded


def load_tracker(path):
    """The office's spending tracker. Account Summary: row 1 title, row 2
    headers, row 3 onward Code | Account | Budget | Committed. Pipeline: row 1
    headers, then Project | Vendor | Type | Account | Amount This Year | Status."""
    wb = openpyxl.load_workbook(path, data_only=True)
    name = os.path.basename(path)
    for tab in ('Account Summary', 'Pipeline'):
        if tab not in wb.sheetnames:
            raise SourceParseError(f'{name}: tab "{tab}" not found')
    summary = []
    for i, r in enumerate(wb['Account Summary'].iter_rows(min_row=3, values_only=True), start=3):
        if not isinstance(r[0], (int, float)):
            continue  # blank or total row
        where = f'{name} Account Summary row {i}'
        summary.append({'code': int(r[0]), 'name': str(r[1] or '').strip(),
                        'budget': _amount(r[2], where), 'committed': _amount(r[3], where)})
    pipeline = []
    for i, r in enumerate(wb['Pipeline'].iter_rows(min_row=2, values_only=True), start=2):
        if not r[0]:
            continue
        where = f'{name} Pipeline row {i}'
        pipeline.append({'project': str(r[0]).strip(), 'vendor': str(r[1] or '').strip(),
                         'type': str(r[2] or '').strip(),
                         'account': _account(r[3], where) or 0,
                         'amount': _amount(r[4], where), 'status': str(r[5] or '').strip()})
    return summary, pipeline


# ── 3. Merge ───────────────────────────────────────────────────────────────

def group_transactions(rows):
    """Per-account totals. committed is rounded from the raw YTD + ENC sum,
    before the parts are rounded, so it matches the finance system to the cent."""
    if rows is None:
        return None
    totals = {}
    for row in rows:
        t = totals.setdefault(row['account'], {'ytd': 0.0, 'enc': 0.0, 'fund': ''})
        t['fund'] = row['fund']
        if row['field'] in ('YTD', 'ENC'):
            t[row['field'].lower()] += row['amount']
    return {acct: {'ytd': round(v['ytd'], 2), 'enc': round(v['enc'], 2),
                   'committed': round(v['ytd'] + v['enc'], 2), 'fund': v['fund']}
            for acct, v in totals.items()}


def status_for(pct, over):
    if over:
        return 'over', 'Over budget'
    if pct >= LIMIT_PCT:
        return 'limit', 'Near limit'
    if pct >= WATCH_PCT:
        return 'watch', 'Watch'
    return 'ok', 'On track'


def merge(summary, pipeline, snapshot_rows, txn_rows):
    """Apply the precedence rules and compute every check. snapshot_rows or
    txn_rows may be None when that export was not supplied."""
    snap = {r['account']: r for r in snapshot_rows} if snapshot_rows is not None else None
    txns = group_transactions(txn_rows)

    accounts = []
    for row in summary:
        code = row['code']
        if code < EXCLUDE_BELOW:
            continue
        s = (snap or {}).get(code)
        t = (txns or {}).get(code)

        budget = s['adj_budget'] if s else row['budget']
        budget_delta = round(budget - row['budget'], 2) if s and abs(budget - row['budget']) > 0.01 else None

        if t:
            committed, committed_src = t['committed'], 'transaction detail'
        elif s:
            committed, committed_src = round(s['ytd'] + s['enc'], 2), 'budget snapshot'
        else:
            committed, committed_src = row['committed'], 'tracker'
        committed_delta = round(committed - row['committed'], 2) if t else None

        planned = round(sum(p['amount'] for p in pipeline if p['account'] == code), 2)
        exposure = round(committed + planned, 2)
        pct = round(exposure / budget * 100, 1) if budget > 0 else 0.0
        over = budget > 0 and exposure > budget
        level, label = status_for(pct, over)
        accounts.append({
            'code': code, 'name': row['name'], 'budget': budget,
            'budget_tracker': row['budget'], 'budget_delta': budget_delta,
            'committed': committed, 'committed_source': committed_src,
            'committed_tracker': row['committed'], 'committed_delta': committed_delta,
            'planned': planned, 'exposure': exposure, 'remaining': round(budget - exposure, 2),
            'pct': pct, 'over': over, 'level': level, 'status': label,
        })

    tracked = {a['code'] for a in accounts}
    untracked, seen = [], set()
    for acct, v in sorted((txns or {}).items()):
        if acct < EXCLUDE_BELOW or acct in tracked:
            continue
        if v['ytd'] != 0 or v['enc'] != 0:
            s = (snap or {}).get(acct)
            untracked.append({'code': acct, 'name': s['title'] if s else '',
                              'ytd': v['ytd'], 'enc': v['enc'], 'committed': v['committed'],
                              'adj_budget': s['adj_budget'] if s else None,
                              'available': s['available'] if s else None})
            seen.add(acct)
    for acct, s in sorted((snap or {}).items()):
        if acct < EXCLUDE_BELOW or acct in tracked or acct in seen:
            continue
        committed = round(s['ytd'] + s['enc'], 2)
        if committed != 0 or s['adj_budget'] != 0:
            untracked.append({'code': acct, 'name': s['title'], 'ytd': s['ytd'], 'enc': s['enc'],
                              'committed': committed, 'adj_budget': s['adj_budget'],
                              'available': s['available']})

    return {
        'accounts': accounts,
        'pipeline': sorted(pipeline, key=lambda p: (p['account'], p['project'])),
        'over_budget': [a for a in accounts if a['over']],
        'budget_changes': [a for a in accounts if a['budget_delta'] is not None],
        'mismatches': [a for a in accounts if a['committed_delta'] is not None
                       and abs(a['committed_delta']) > MISMATCH_TOLERANCE],
        'untracked': untracked,
        'totals': {k: round(sum(a[k] for a in accounts), 2)
                   for k in ('budget', 'committed', 'planned', 'remaining')},
    }


def build(folder):
    paths = locate_sources(folder)
    summary, pipeline = load_tracker(paths['tracker'])
    snapshot_rows, snap_excl = load_snapshot(paths['snapshot']) if paths['snapshot'] else (None, 0)
    txn_rows, txn_excl = load_transactions(paths['transactions']) if paths['transactions'] else (None, 0)
    view = merge(summary, pipeline, snapshot_rows, txn_rows)
    view['sources'] = {k: os.path.basename(v) if v else None for k, v in paths.items()}
    view['excluded_lines'] = snap_excl + txn_excl
    return view


# ── 4. Render ──────────────────────────────────────────────────────────────

def money(n):
    return ('-$' if n < 0 else '$') + f'{abs(n):,.0f}'


def signed(n):
    return ('+' if n > 0 else '-' if n < 0 else '') + money(abs(n))


e = html.escape

CSS = """
:root{--navy:#1F3A5F;--slate:#4A5568;--gold:#B7791F;--red:#B42318;--amber:#8A5A00;
--green:#1E6B3A;--line:#D5DBE3;--soft:#F5F7FA;}
*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif;
color:#1A202C;background:#fff;font-size:15px;line-height:1.5}
header{background:var(--navy);color:#fff;padding:20px 32px;border-bottom:4px solid var(--gold)}
header h1{margin:0;font-size:24px}
header p{margin:4px 0 0;font-size:14px;color:#E2E8F0}
main{max-width:1180px;margin:0 auto;padding:24px 32px 40px}
h2{color:var(--navy);font-size:19px;margin:32px 0 10px;padding-bottom:4px;border-bottom:2px solid var(--line)}
h3{color:var(--slate);font-size:16px;margin:20px 0 6px}
.alert{border-left:6px solid var(--red);background:#FDF2F1;padding:12px 16px;margin:0 0 8px}
.alert h2{margin:0 0 4px;border:0;padding:0;color:var(--red);font-size:17px}
.alert p{margin:0}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:20px 0 0;padding:0}
.kpis div{border:1px solid var(--line);border-top:4px solid var(--navy);padding:12px 16px;margin:0}
.kpis dt{font-size:13px;color:var(--slate);font-weight:600}
.kpis dd{margin:2px 0 0;font-size:24px;font-weight:700;font-variant-numeric:tabular-nums}
table{width:100%;border-collapse:collapse;margin:6px 0 4px}
caption{text-align:left;color:var(--slate);font-size:14px;padding:0 0 6px}
th,td{padding:7px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:middle}
thead th{background:var(--soft);color:var(--navy);font-size:13px;border-bottom:2px solid var(--navy)}
tbody th{font-weight:600}
.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.bar{display:inline-block;width:90px;height:8px;background:#E2E8F0;vertical-align:middle;margin-right:8px;position:relative}
.bar span{position:absolute;left:0;top:0;bottom:0}
.tag{display:inline-block;padding:1px 8px;border-radius:10px;font-size:13px;font-weight:600;white-space:nowrap}
.ok{color:var(--green)} .ok.tag{background:#E7F3EC}
.watch{color:var(--amber)} .watch.tag{background:#FDF3DC}
.limit{color:var(--red)} .limit.tag{background:#FDE7E5}
.over{color:#fff} .over.tag{background:var(--red)}
.fill-ok{background:var(--green)} .fill-watch{background:var(--gold)}
.fill-limit,.fill-over{background:var(--red)}
.neg{color:var(--red);font-weight:700}
.note{color:var(--slate);font-size:14px;margin:4px 0 0}
.good{color:var(--green);font-weight:600}
footer{color:var(--slate);font-size:13px;border-top:1px solid var(--line);margin-top:36px;padding-top:10px}
@media print{header{print-color-adjust:exact;-webkit-print-color-adjust:exact}}
"""


def table(caption, headers, rows, numeric_from=2):
    """headers: list of labels; rows: list of cell lists (already escaped HTML).
    The first cell of each row is a row header. Columns at index >= numeric_from
    are right-aligned unless the label starts with '!'."""
    cols = []
    for i, h in enumerate(headers):
        cls = ' class="num"' if i >= numeric_from and not h.startswith('!') else ''
        cols.append((h.lstrip('!'), cls))
    head = ''.join(f'<th scope="col"{c}>{e(h)}</th>' for h, c in cols)
    body = []
    for r in rows:
        cells = [f'<th scope="row">{r[0]}</th>']
        cells += [f'<td{cols[i][1]}>{v}</td>' for i, v in enumerate(r[1:], start=1)]
        body.append('<tr>' + ''.join(cells) + '</tr>')
    return (f'<table><caption>{e(caption)}</caption><thead><tr>{head}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody></table>')


def render(view, office, fiscal_year, as_of):
    t, src = view['totals'], view['sources']
    parts = []

    over = view['over_budget']
    if over:
        items = ', '.join(f'{a["code"]} {e(a["name"])} ({money(a["remaining"])})' for a in over)
        parts.append(f'<section class="alert" aria-labelledby="alert-h"><h2 id="alert-h">'
                     f'{len(over)} account{"s" if len(over) > 1 else ""} over budget</h2>'
                     f'<p>Committed plus planned spending is above the budget for: {items}.</p></section>')

    parts.append('<section aria-labelledby="sum-h"><h2 id="sum-h">Summary</h2><dl class="kpis">'
                 + ''.join(f'<div><dt>{label}</dt><dd class="{"neg" if v < 0 else ""}">{money(v)}</dd></div>'
                           for label, v in (('Budget', t['budget']), ('Committed', t['committed']),
                                            ('Planned, not yet committed', t['planned']),
                                            ('Remaining after planned', t['remaining'])))
                 + '</dl></section>')

    rows = []
    for a in view['accounts']:
        width = min(a['pct'], 100)
        rows.append([
            e(str(a['code'])), e(a['name']), money(a['budget']), money(a['committed']),
            money(a['planned']),
            f'<span class="{"neg" if a["remaining"] < 0 else ""}">{money(a["remaining"])}</span>',
            f'<span class="bar" aria-hidden="true"><span class="fill-{a["level"]}" style="width:{width}%">'
            f'</span></span>{a["pct"]:.0f}%',
            f'<span class="tag {a["level"]}">{a["status"]}</span>'])
    parts.append('<section aria-labelledby="acct-h"><h2 id="acct-h">Accounts</h2>'
                 + table('Each account the office tracks. Used is committed plus planned, as a share of budget.',
                         ['Code', 'Account', 'Budget', 'Committed', 'Planned', 'Remaining', 'Used', '!Status'],
                         rows)
                 + f'<p class="note">Status: On track under {WATCH_PCT}% used, Watch from {WATCH_PCT}%, '
                   f'Near limit from {LIMIT_PCT}%, Over budget above 100%.</p></section>')

    checks = ['<section aria-labelledby="chk-h"><h2 id="chk-h">Checks against the finance system</h2>']
    checks.append('<h3>Budget changes</h3>')
    if view['budget_changes']:
        checks.append(table('The finance system shows a different budget than the tracker, usually a mid-year '
                            'transfer. The dashboard uses the finance system figure.',
                            ['Code', 'Account', 'Tracker budget', 'Finance system budget', 'Difference'],
                            [[e(str(a['code'])), e(a['name']), money(a['budget_tracker']), money(a['budget']),
                              signed(a['budget_delta'])] for a in view['budget_changes']]))
    elif src['snapshot']:
        checks.append('<p class="good">Every tracked budget matches the finance system.</p>')
    else:
        checks.append('<p class="note">No budget snapshot supplied; budgets come from the tracker.</p>')

    checks.append('<h3>Committed figures that disagree</h3>')
    if view['mismatches']:
        checks.append(table(f'The finance system\'s committed total differs from the tracker by more than '
                            f'{money(MISMATCH_TOLERANCE)}. The dashboard uses the finance system figure.',
                            ['Code', 'Account', 'Tracker', 'Finance system', 'Difference'],
                            [[e(str(a['code'])), e(a['name']), money(a['committed_tracker']),
                              money(a['committed']), signed(a['committed_delta'])]
                             for a in view['mismatches']]))
    else:
        checks.append(f'<p class="good">No committed figure is off by more than {money(MISMATCH_TOLERANCE)}.</p>')

    checks.append('<h3>Untracked activity</h3>')
    if view['untracked']:
        checks.append(table('Account codes with finance-system activity that the tracker does not list. '
                            'Each one is either a miscoded charge or an account to add to the tracker.',
                            ['Code', 'Account', 'Budget', 'Spent', 'Encumbered', 'Committed', 'Available'],
                            [[e(str(u['code'])), e(u['name']) or '<em>No title in export</em>',
                              money(u['adj_budget']) if u['adj_budget'] is not None else 'n/a',
                              money(u['ytd']), money(u['enc']), money(u['committed']),
                              (f'<span class="{"neg" if u["available"] < 0 else ""}">{money(u["available"])}</span>'
                               if u['available'] is not None else 'n/a')]
                             for u in view['untracked']]))
    else:
        checks.append('<p class="good">No activity on untracked accounts.</p>')
    checks.append(f'<p class="note">{view["excluded_lines"]} payroll and fringe lines (accounts below '
                  f'{EXCLUDE_BELOW}) were left out of every figure on this page.</p></section>')
    parts.append(''.join(checks))

    if view['pipeline']:
        parts.append('<section aria-labelledby="pipe-h"><h2 id="pipe-h">Planned spending</h2>'
                     + table('Spending in the tracker that is not yet committed in the finance system.',
                             ['Project', 'Vendor', 'Account', 'Amount', '!Status'],
                             [[e(p['project']), e(p['vendor']), e(str(p['account'])), money(p['amount']),
                               e(p['status'])] for p in view['pipeline']], numeric_from=3)
                     + '</section>')

    used = [f'{label}: {e(src[k])}' for k, label in
            (('tracker', 'tracker'), ('snapshot', 'budget snapshot'), ('transactions', 'transaction detail'))
            if src[k]]
    committed_from = ('transaction detail' if src['transactions'] else 'budget snapshot')
    title = f'{office} budget, {fiscal_year}'
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(title)}</title>
<style>{CSS}</style>
</head>
<body>
<header>
<h1>{e(title)}</h1>
<p>As of {e(as_of)}. Committed figures from the finance system's {committed_from}.</p>
</header>
<main>
{''.join(parts)}
<footer>Sources: {'; '.join(used)}. Generated by generate_budget_dashboard.py.</footer>
</main>
</body>
</html>
"""


def main(argv=None):
    ap = argparse.ArgumentParser(description='Render a one-page budget dashboard from a spending '
                                             'tracker and finance-system exports.')
    ap.add_argument('--sources', required=True, help='folder holding the tracker and exports')
    ap.add_argument('--output', required=True, help='HTML file to write')
    ap.add_argument('--office', default='Office of Communications')
    ap.add_argument('--fiscal-year', default='2026-27')
    ap.add_argument('--as-of', default=date.today().isoformat(), help='date shown on the page')
    args = ap.parse_args(argv)
    try:
        view = build(args.sources)
    except (SourceMissing, SourceParseError) as err:
        print(f'FATAL: {err}. Nothing was written.', file=sys.stderr)
        return err.exit_code
    except Exception as err:  # an unreadable workbook is a parse failure, not a crash
        print(f'FATAL: could not read a source: {err!r}. Nothing was written.', file=sys.stderr)
        return EXIT_PARSE
    page = render(view, args.office, args.fiscal_year, args.as_of)
    with open(args.output, 'w', encoding='utf-8') as f:
        f.write(page)
    print(f'Wrote {args.output}: {len(view["accounts"])} accounts, '
          f'{len(view["over_budget"])} over budget, {len(view["untracked"])} untracked, '
          f'{len(view["mismatches"])} committed mismatches.')
    return EXIT_OK


if __name__ == '__main__':
    sys.exit(main())
