#!/usr/bin/env python3
"""
Writes the fictional sample sources for Cedar Ridge College's Office of
Communications into sample/sources/. Every account, vendor and figure is
invented. Running it twice produces the same figures.

The sample is built so every check on the dashboard has something to show:
  5120  digital advertising goes over budget once planned spending is added
  5120, 5130  the finance system's budget differs from the tracker (transfers)
  5410  the tracker's committed figure is behind the finance system
  5220  appears only in the budget snapshot, so committed falls back to it
  5650, 4510  finance-system activity on accounts the tracker does not list
  2110, 3110, 3410  payroll and fringe lines that must be left out
"""

import csv
import os
import random
import sys
from datetime import date, datetime, timedelta

import openpyxl

FUND, ORG = 'GEN01', '2410'

# code, name, tracker budget, finance-system adjusted budget, spent target, encumbered target
ACCOUNTS = [
    (4310, 'Printing and copy supplies',       18000, 18000,  6420.18,  1250.00),
    (4320, 'Event supplies',                   12500, 12500,  4875.40,   900.00),
    (4330, 'Promotional items',                 9000,  9000,  3310.75,  2400.00),
    (4410, 'Photo equipment, non-capital',      6000,  6000,  1189.99,     0.00),
    (5110, 'Printing services',                42000, 42000, 17840.00,  9600.00),
    (5120, 'Digital advertising',              55000, 50000, 31250.00, 16400.00),
    (5130, 'Print and outdoor advertising',    30000, 33000, 14500.00,  6000.00),
    (5140, 'Radio and streaming advertising',  24000, 24000,  9800.00,  7200.00),
    (5210, 'Software subscriptions',           16500, 16500, 12960.00,  1440.00),
    (5220, 'Website hosting and domains',       4800,  4800,  1620.00,  1080.00),
    (5310, 'Postage and mailing',               7500,  7500,  2214.36,     0.00),
    (5410, 'Conference and training',           8000,  8000,  2640.00,     0.00),
    (5510, 'Event venues and catering',        20000, 20000,  8935.50,  3500.00),
    (5610, 'Photo and video services',         15000, 15000,  5400.00,  4200.00),
    (5710, 'Memberships and dues',              3200,  3200,  2875.00,     0.00),
]
TRACKER_COMMITTED_OVERRIDE = {5410: 2100.00}   # the tracker missed a registration
SNAPSHOT_ONLY = {5220}                          # no transaction lines for this account

# code, title, adjusted budget, spent, encumbered (not in the tracker)
UNTRACKED = [
    (5650, 'Translation services',      0,    1450.00, 0.00),
    (4510, 'Furniture, non-capital', 2000,       0.00, 0.00),
]
PAYROLL = [
    (2110, 'Classified salaries', 182000, 60666.64),
    (3110, 'Retirement benefits',  41000, 13666.68),
    (3410, 'Health benefits',      36000, 12000.00),
]

VENDORS = {
    4310: ['Copper Leaf Office Supply'], 4320: ['Ridgeway Party Rental', 'Copper Leaf Office Supply'],
    4330: ['Summit Promo Goods'], 4410: ['Lakeside Camera'],
    5110: ['Cedar Valley Printing', 'Harbor Lane Press'], 5120: ['Brightline Digital', 'Northgate Media'],
    5130: ['Northgate Outdoor', 'Valley Weekly News'], 5140: ['Riverbend Radio Group', 'Westward Streaming'],
    5210: ['Design suite license', 'Survey tool license', 'Scheduling app license'],
    5220: ['Hillcrest Hosting'], 5310: ['Postal permit account'], 5410: ['Regional comms conference'],
    5510: ['Pinecrest Catering', 'Ridgeway Event Hall'], 5610: ['Oakmont Studio'],
    5710: ['College marketing association'], 5650: ['Clearword Translation'],
    2110: ['Payroll'], 3110: ['Payroll'], 3410: ['Payroll'],
}

PIPELINE = [
    ('Spring enrollment campaign, social ads', 'Brightline Digital', 'PO', 5120, 4000.00, 'Quote received'),
    ('Spring enrollment campaign, billboards', 'Northgate Outdoor', 'PO', 5130, 6500.00, 'Requisition drafted'),
    ('Commencement program printing', 'Cedar Valley Printing', 'PO', 5110, 7800.00, 'Quote received'),
    ('Commencement photography', 'Oakmont Studio', 'Contract', 5610, 2500.00, 'Contract in review'),
    ('Open house catering', 'Pinecrest Catering', 'Direct pay', 5510, 3200.00, 'Not started'),
    ('Annual report mailing', 'Postal permit account', 'Direct pay', 5310, 1850.00, 'Not started'),
    ('Spring recruitment giveaways', 'Summit Promo Goods', 'PO', 4330, 2100.00, 'Quote received'),
]

FY_START = date(2026, 7, 1)
AS_OF = date(2026, 10, 15)


def split(total, parts, rng):
    """Split a total into `parts` positive amounts that sum exactly to it."""
    if total == 0:
        return []
    cents = round(total * 100)
    cuts = sorted(rng.sample(range(1, cents), parts - 1)) if parts > 1 and cents > parts else []
    edges = [0] + cuts + [cents]
    return [(edges[i + 1] - edges[i]) / 100 for i in range(len(edges) - 1)]


def txn_lines(code, spent, enc, rng):
    lines = []
    if spent == 0 and enc == 0:
        return lines
    vendors = VENDORS[code]
    for i, amt in enumerate(split(spent, min(4, max(1, int(spent // 2500) + 1)), rng)):
        d = FY_START + timedelta(days=rng.randrange(0, (AS_OF - FY_START).days))
        doc = 'J' if code < 4000 else 'I'
        lines.append([code, 'YTD', amt, 'JV' if code < 4000 else 'INV', f'{doc}{rng.randrange(100000, 999999)}',
                      vendors[i % len(vendors)], d])
    for i, amt in enumerate(split(enc, 2 if enc > 3000 else 1, rng)):
        d = FY_START + timedelta(days=rng.randrange(0, (AS_OF - FY_START).days))
        lines.append([code, 'ENC', amt, 'PO', f'P{rng.randrange(1000000, 9999999)}',
                      vendors[i % len(vendors)], d])
    return lines


def write_snapshot(path, rows):
    """rows: (code, title, adj_budget, ytd, enc). Layout matches the export:
    two parameter rows, a header row, then one row per account."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Budget Snapshot'
    ws.append([f'Fund: {FUND}   Org: {ORG}'])
    ws.append([f'Fiscal year 2027   Through: {AS_OF.isoformat()}'])
    ws.append(['Account', 'Type', 'Title', 'Adj Budget', 'YTD Activity', 'Commitments', 'Available Balance'])
    for code, title, adj, ytd, enc in rows:
        ws.append([str(code), 'L' if code < 4000 else 'E', title, adj, ytd, enc, round(adj - ytd - enc, 2)])
    _save(wb, path)


TXN_HEADERS = ['Account', 'Account Title', 'Org', 'Program', 'Type', 'Document', 'Description',
               'Commit Indicator', 'Fund', 'Activity', 'Location', 'Transaction Date', 'Field',
               'Amount', 'Running Total']


def write_transactions_csv(path, lines, titles):
    """lines: [code, field, amount, type, doc, description, date]. Two metadata
    rows, a header row, then one row per transaction line (15 columns)."""
    running = {}
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow([f'Transaction detail   Fund: {FUND}   Org: {ORG}'] + [''] * 14)
        w.writerow([f'Fiscal year 2027   Through: {AS_OF.isoformat()}'] + [''] * 14)
        w.writerow(TXN_HEADERS)
        for code, field, amt, typ, doc, desc, d in lines:
            running[code] = round(running.get(code, 0) + amt, 2)
            w.writerow([code, titles.get(code, ''), ORG, '6600', typ, doc, desc, 'U', FUND, '', '',
                        d.isoformat(), field, f'{amt:.2f}', f'{running[code]:.2f}'])


def write_tracker(path, summary, pipeline):
    """summary: (code, name, budget, committed); pipeline: PIPELINE-shaped tuples."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Account Summary'
    ws.append(['Office of Communications spending tracker, 2026-27'])
    ws.append(['Code', 'Account', 'Budget', 'Committed'])
    for row in summary:
        ws.append(list(row))
    ws.append(['Total', '', sum(r[2] for r in summary), round(sum(r[3] for r in summary), 2)])
    wp = wb.create_sheet('Pipeline')
    wp.append(['Project', 'Vendor', 'Type', 'Account', 'Amount This Year', 'Status'])
    for row in pipeline:
        wp.append(list(row))
    _save(wb, path)


def _save(wb, path):
    wb.properties.creator = 'Jordan Lee'
    wb.properties.lastModifiedBy = 'Jordan Lee'
    wb.properties.created = wb.properties.modified = datetime(2026, 10, 15, 8, 0, 0)
    wb.save(path)


def main(out_dir):
    rng = random.Random(2027)
    os.makedirs(out_dir, exist_ok=True)
    titles = {c: n for c, n, *_ in ACCOUNTS}
    titles.update({c: n for c, n, *_ in UNTRACKED})
    titles.update({c: n for c, n, *_ in PAYROLL})

    lines = []
    for code, _, _, _, spent, enc in ACCOUNTS:
        if code not in SNAPSHOT_ONLY:
            lines += txn_lines(code, spent, enc, rng)
    for code, _, _, spent, enc in UNTRACKED:
        lines += txn_lines(code, spent, enc, rng)
    for code, _, _, spent in PAYROLL:
        lines += txn_lines(code, spent, 0, rng)
    lines.sort(key=lambda l: (l[0], l[6], l[4]))
    write_transactions_csv(os.path.join(out_dir, 'transaction_detail.csv'), lines, titles)

    snap = [(c, n, adj, spent, enc) for c, n, _, adj, spent, enc in ACCOUNTS]
    snap += [(c, n, adj, spent, enc) for c, n, adj, spent, enc in UNTRACKED]
    snap += [(c, n, adj, spent, 0) for c, n, adj, spent in PAYROLL]
    write_snapshot(os.path.join(out_dir, 'budget_snapshot_current.xlsx'), sorted(snap))

    summary = [(c, n, budget, TRACKER_COMMITTED_OVERRIDE.get(c, round(spent + enc, 2)))
               for c, n, budget, _, spent, enc in ACCOUNTS]
    write_tracker(os.path.join(out_dir, 'spending_tracker.xlsx'), summary, PIPELINE)
    print(f'Wrote sample sources to {out_dir}')


if __name__ == '__main__':
    here = os.path.dirname(os.path.abspath(__file__))
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, '..', 'sample', 'sources'))
