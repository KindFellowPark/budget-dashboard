"""Tests for the merge rules, the payroll exclusion, the untracked sweep and
loud failure on missing or broken sources."""
import os
import subprocess
import sys
from datetime import date

import pytest

import generate_budget_dashboard as gen
import make_sample_data as sample

SCRIPT = os.path.join(os.path.dirname(__file__), '..', 'scripts', 'generate_budget_dashboard.py')


def tracker(code, budget=10000.0, committed=1000.0, name='Printing services'):
    return {'code': code, 'name': name, 'budget': budget, 'committed': committed}


def snap(code, adj_budget, ytd, enc, title='From snapshot'):
    return {'account': code, 'title': title, 'adj_budget': adj_budget, 'ytd': ytd, 'enc': enc,
            'available': round(adj_budget - ytd - enc, 2)}


def txn(code, field, amount):
    return {'account': code, 'field': field, 'amount': amount, 'type': 'INV', 'doc': 'I1',
            'description': '', 'fund': 'GEN01', 'date': '2026-09-01'}


# ── merge precedence ──────────────────────────────────────────────────────

def test_transaction_detail_wins_for_committed():
    view = gen.merge([tracker(5110)], [], [snap(5110, 10000, 900, 100)],
                     [txn(5110, 'YTD', 2000), txn(5110, 'ENC', 500), txn(5110, 'OBD', 10000)])
    a = view['accounts'][0]
    assert a['committed'] == 2500            # YTD + ENC from transactions; OBD ignored
    assert a['committed_source'] == 'transaction detail'


def test_snapshot_used_when_account_missing_from_transactions():
    view = gen.merge([tracker(5220)], [], [snap(5220, 4800, 1620, 1080)], [txn(5110, 'YTD', 50)])
    a = view['accounts'][0]
    assert a['committed'] == 2700
    assert a['committed_source'] == 'budget snapshot'
    assert a['committed_delta'] is None      # mismatch check only runs on transaction detail


def test_tracker_used_when_no_export_covers_account():
    view = gen.merge([tracker(5310, committed=777)], [], None, None)
    assert view['accounts'][0]['committed'] == 777
    assert view['accounts'][0]['committed_source'] == 'tracker'


def test_snapshot_adjusted_budget_wins_and_change_is_reported():
    view = gen.merge([tracker(5120, budget=55000)], [], [snap(5120, 50000, 0, 0)], None)
    a = view['accounts'][0]
    assert a['budget'] == 50000
    assert a['budget_delta'] == -5000
    assert view['budget_changes'] == [a]


def test_committed_is_rounded_after_summing():
    # 0.005 + 0.005 rounds to 0.01 only if summed before rounding
    view = gen.merge([tracker(5110)], [], None, [txn(5110, 'YTD', 0.005), txn(5110, 'ENC', 0.005)])
    assert view['accounts'][0]['committed'] == 0.01


def test_mismatch_over_tolerance_is_flagged_and_small_one_is_not():
    view = gen.merge([tracker(5410, committed=2100), tracker(5510, committed=1000)], [], None,
                     [txn(5410, 'YTD', 2640), txn(5510, 'YTD', 1040)])
    assert [a['code'] for a in view['mismatches']] == [5410]


def test_over_budget_counts_planned_spending():
    pipeline = [{'project': 'Ads', 'vendor': 'V', 'type': 'PO', 'account': 5120,
                 'amount': 4000, 'status': ''}]
    view = gen.merge([tracker(5120, budget=50000)], pipeline, None, [txn(5120, 'YTD', 47650)])
    a = view['accounts'][0]
    assert a['over'] and a['status'] == 'Over budget' and a['remaining'] == -1650
    assert view['over_budget'] == [a]


# ── exclusion rule ────────────────────────────────────────────────────────

def test_payroll_and_fringe_excluded_by_loaders(tmp_path):
    rows = [(2110, 'Salaries', 100000, 30000, 0), (3410, 'Health', 9000, 3000, 0),
            (5110, 'Printing services', 10000, 2000, 0)]
    sample.write_snapshot(tmp_path / 'budget_snapshot_x.xlsx', rows)
    lines = [[2110, 'YTD', 30000, 'JV', 'J1', 'Payroll', date(2026, 8, 1)],
             [5110, 'YTD', 2000, 'INV', 'I1', 'Vendor', date(2026, 8, 1)]]
    sample.write_transactions_csv(tmp_path / 'transaction_detail.csv', lines, {})
    snap_rows, snap_excl = gen.load_snapshot(str(tmp_path / 'budget_snapshot_x.xlsx'))
    txn_rows, txn_excl = gen.load_transactions(str(tmp_path / 'transaction_detail.csv'))
    assert [r['account'] for r in snap_rows] == [5110] and snap_excl == 2
    assert [r['account'] for r in txn_rows] == [5110] and txn_excl == 1


def test_payroll_never_appears_as_untracked_or_tracked():
    view = gen.merge([tracker(3110)], [], [snap(2110, 1000, 500, 0)], [txn(3410, 'YTD', 900)])
    assert view['accounts'] == [] and view['untracked'] == []


# ── untracked sweep ───────────────────────────────────────────────────────

def test_untracked_sweep_finds_both_exports_without_duplicates():
    view = gen.merge(
        [tracker(5110)], [],
        [snap(5110, 10000, 0, 0), snap(5650, 0, 1450, 0, 'Translation'),
         snap(4510, 2000, 0, 0, 'Furniture'), snap(4520, 0, 0, 0, 'Dormant')],
        [txn(5110, 'YTD', 10), txn(5650, 'YTD', 1450)])
    found = {u['code']: u for u in view['untracked']}
    assert set(found) == {5650, 4510}        # 4520 has no budget and no activity
    assert found[5650]['name'] == 'Translation' and found[5650]['available'] == -1450
    assert found[4510]['adj_budget'] == 2000


def test_untracked_from_transactions_without_snapshot():
    view = gen.merge([tracker(5110)], [], None, [txn(5990, 'ENC', 300)])
    assert view['untracked'][0]['code'] == 5990 and view['untracked'][0]['adj_budget'] is None


# ── loud failure ──────────────────────────────────────────────────────────

def run(sources, out):
    return subprocess.run([sys.executable, SCRIPT, '--sources', str(sources), '--output', str(out)],
                          capture_output=True, text=True)


def test_missing_tracker_exits_3_and_writes_nothing(tmp_path):
    sample.main(str(tmp_path))
    os.remove(tmp_path / 'spending_tracker.xlsx')
    out = tmp_path / 'dash.html'
    r = run(tmp_path, out)
    assert r.returncode == 3 and 'FATAL' in r.stderr and not out.exists()


def test_no_finance_exports_exits_3(tmp_path):
    sample.main(str(tmp_path))
    os.remove(tmp_path / 'budget_snapshot_current.xlsx')
    os.remove(tmp_path / 'transaction_detail.csv')
    r = run(tmp_path, tmp_path / 'dash.html')
    assert r.returncode == 3 and 'no finance-system export' in r.stderr


def test_two_snapshot_files_exit_3_instead_of_guessing(tmp_path):
    sample.main(str(tmp_path))
    (tmp_path / 'BUDGET_SNAPSHOT_old.xlsx').write_bytes((tmp_path / 'budget_snapshot_current.xlsx').read_bytes())
    r = run(tmp_path, tmp_path / 'dash.html')
    assert r.returncode == 3 and 'Remove the stale one' in r.stderr


def test_changed_export_layout_exits_1(tmp_path):
    sample.main(str(tmp_path))
    p = tmp_path / 'transaction_detail.csv'
    text = p.read_text().replace(',Field,Amount,', ',Ledger,Amount,')
    p.write_text(text)
    out = tmp_path / 'dash.html'
    r = run(tmp_path, out)
    assert r.returncode == 1 and not out.exists()


def test_sample_builds_with_every_flag_showing(tmp_path):
    sample.main(str(tmp_path))
    view = gen.build(str(tmp_path))
    assert [a['code'] for a in view['over_budget']] == [5120]
    assert {u['code'] for u in view['untracked']} == {5650, 4510}
    assert [a['code'] for a in view['mismatches']] == [5410]
    assert {a['code'] for a in view['budget_changes']} == {5120, 5130}
    assert view['excluded_lines'] > 0
