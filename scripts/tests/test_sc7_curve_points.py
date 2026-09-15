"""SC7 curve-point (re)assignment tests.

Mock-first (no Excel, no PLAXIS): the fake book/sheet/range trio emulates
the xlwings surface SC7 touches (sheets[...].range(...).value). Pure
helpers are tested directly; planting is tested against a fake
(g_i, g_o) pair recording addcurvepoint calls.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from sc7_curve_points import (
    CURVE_DATA_END,
    CURVE_DATA_START,
    SC7_CODE_VERSION,
    _bar,
    _main_io_inner,
    plant_points,
    read_curve_point_rows,
    round_num,
    safe_num,
    safe_str,
)


# ── Fakes ──

class _FakeRange:
    def __init__(self, sheet, addr):
        self._sheet = sheet
        self._addr = addr

    @property
    def value(self):
        if isinstance(self._addr, tuple):
            # Rows are 1-based like xlwings; grid columns start at B=2,
            # so sheet column c → grid index c-2.
            (r1, c1), (r2, c2) = self._addr
            return [list(row[c1 - 2:c2 - 1]) + [None] * max(
                0, (c2 - c1 + 1) - len(row[c1 - 2:c2 - 1]))
                for row in self._sheet._grid[r1 - 1:r2]]
        return self._sheet._cells.get(self._addr)

    @value.setter
    def value(self, v):
        self._sheet._cells[self._addr] = v
        self._sheet._writes.append((self._addr, v))


class _FakeSheet:
    def __init__(self, name, grid=None):
        self.name = name
        self._grid = grid or []
        self._cells = {}
        self._writes = []

    def range(self, a, b=None):
        if b is None:
            return _FakeRange(self, a)
        return _FakeRange(self, (a, b))


class _FakeBook:
    def __init__(self, curve_grid=None, main_cells=None):
        self.sheets = {
            'output': _FakeSheet('output', curve_grid),
            'main': _FakeSheet('main'),
        }
        for addr, v in (main_cells or {}).items():
            self.sheets['main']._cells[addr] = v
        self.fullname = r'D:\fake\EXAMPLE 3.xlsm'


def _grid(rows):
    """Sheet-shaped grid: 8 blank header rows, then the B9:F108 block.

    The fake range() slices 1-based like xlwings, so rows 9-108 must sit
    at grid indices 8-107 — without the 8 leading blanks every row reads
    8 off (the bug this helper fixes).
    """
    g = [[None] * 5 for _ in range(CURVE_DATA_START - 1)]
    g += [list(r) + [None] * max(0, 5 - len(r)) for r in rows]
    g += [[None] * 5 for _ in range(CURVE_DATA_END + 1 - len(g))]
    return g


class _FakeGI:
    def __init__(self, fail=None):
        self.calls = []
        self.fail = fail

    def gotostages(self):
        self.calls.append('gotostages')
        if self.fail == 'gotostages':
            raise RuntimeError('boom')

    def selectmeshpoints(self):
        self.calls.append('selectmeshpoints')
        if self.fail == 'selectmeshpoints':
            raise RuntimeError('boom')


class _FakeGO:
    def __init__(self, fail_rows=()):
        self.planted = []
        self.updated = False
        self.fail_rows = set(fail_rows)

    def addcurvepoint(self, ptype, coords):
        if coords in self.fail_rows:
            raise RuntimeError('bad point')
        self.planted.append((ptype, coords))

    def update(self):
        self.updated = True


# ── Tests ──

class TestBanner:
    def test_version(self):
        assert SC7_CODE_VERSION == 'v0.8.3-sc7'


class TestHelpers:
    def test_safe_str(self):
        assert safe_str(None) == ''
        assert safe_str('  a ') == 'a'
        assert safe_str(None, 'd') == 'd'

    def test_safe_num(self):
        assert safe_num('3.5') == 3.5
        assert safe_num(None) is None
        assert safe_num('x') is None

    def test_round_num(self):
        assert round_num(3.14159) == 3.142
        assert round_num('x') == 'x'

    def test_bar(self):
        assert _bar(0.0).endswith('0%')
        assert _bar(1.0).endswith('100%')
        assert _bar(2.0).endswith('100%')


class TestReadRows:
    def test_two_rows_in_order(self):
        wb = _FakeBook(_grid([
            ['curve_node_1', 'curve_node', 55, 7.601, 'node'],
            ['curve_stresspoint_1', 'curve_stresspoint', 32.244, -3.107,
             'stresspoint'],
        ]))
        rows, skipped = read_curve_point_rows(wb)
        assert len(rows) == 2 and not skipped
        assert rows[0]['name'] == 'curve_node_1'
        assert rows[0]['excel_row'] == 9
        assert rows[1]['excel_row'] == 10
        assert rows[1]['y'] == -3.107

    def test_blank_rows_silent(self):
        wb = _FakeBook(_grid([]))
        rows, skipped = read_curve_point_rows(wb)
        assert rows == [] and skipped == []

    def test_bad_type_skipped(self):
        wb = _FakeBook(_grid([
            ['x', 'curve_node', 1, 2, 'bogus'],
        ]))
        rows, skipped = read_curve_point_rows(wb)
        assert rows == [] and len(skipped) == 1
        assert skipped[0][0] == 9

    def test_bad_xy_skipped(self):
        wb = _FakeBook(_grid([
            ['x', 'curve_node', 'a', 2, 'node'],
        ]))
        rows, skipped = read_curve_point_rows(wb)
        assert rows == [] and len(skipped) == 1

    def test_user_row_kept(self):
        wb = _FakeBook(_grid([
            ['my_extra', '', 10, 20, 'node'],
        ]))
        rows, skipped = read_curve_point_rows(wb)
        assert len(rows) == 1 and not skipped

    def test_case_insensitive_type(self):
        wb = _FakeBook(_grid([
            ['n', 'curve_node', 1, 2, 'Node'],
        ]))
        rows, _ = read_curve_point_rows(wb)
        assert rows[0]['type'] == 'node'


class TestPlant:
    def test_order_and_counts(self):
        g_i, g_o = _FakeGI(), _FakeGO()
        rows = [
            {'name': 'a', 'x': 1, 'y': 2, 'type': 'node',
             'excel_row': 9},
            {'name': 'b', 'x': 3, 'y': 4, 'type': 'stresspoint',
             'excel_row': 10},
        ]
        n, s, fails = plant_points(g_i, g_o, rows)
        assert (n, s, fails) == (1, 1, [])
        assert g_o.planted == [('node', (1, 2)),
                               ('stresspoint', (3, 4))]
        assert g_o.updated
        assert g_i.calls == ['gotostages', 'selectmeshpoints']

    def test_per_row_failure_continues(self):
        g_i, g_o = _FakeGI(), _FakeGO(fail_rows={(9, 9)})
        rows = [
            {'name': 'a', 'x': 9, 'y': 9, 'type': 'node',
             'excel_row': 9},
            {'name': 'b', 'x': 1, 'y': 2, 'type': 'node',
             'excel_row': 10},
        ]
        n, s, fails = plant_points(g_i, g_o, rows)
        assert n == 1 and len(fails) == 1 and fails[0][0] == 9
        assert g_o.updated

    def test_gotostages_failure(self):
        g_i, g_o = _FakeGI(fail='gotostages'), _FakeGO()
        n, s, fails = plant_points(g_i, g_o, [
            {'name': 'a', 'x': 1, 'y': 2, 'type': 'node',
             'excel_row': 9}])
        assert (n, s) == (0, 0) and len(fails) == 1


class TestMainInner:
    def test_no_rows_done_text(self):
        wb = _FakeBook(_grid([]))
        _main_io_inner(wb)
        assert wb.sheets['output']._cells['L2'].startswith(
            'sc7: done — no curve-point rows')

    def test_needs_plaxis_without_it(self):
        # No plxscripting here → connection fails → FAILED progress.
        wb = _FakeBook(_grid([
            ['curve_node_1', 'curve_node', 55, 7.601, 'node'],
        ]), main_cells={'V5': 10000, 'V6': 'pw'})
        _main_io_inner(wb)
        assert wb.sheets['output']._cells['L2'].startswith('sc7: FAILED')

    def test_progress_never_raises(self):
        wb = _FakeBook(_grid([]))
        del wb.sheets['output']
        _main_io_inner(wb)  # must not raise


class TestLayoutGuard:
    """L2 must stay a short single line — long wrapped text stretches
    the cell-anchored buttons (row 2 blew to 110pt in live Ex3)."""

    def test_multiline_failed_collapses(self):
        from sc7_curve_points import _set_progress
        wb = _FakeBook(_grid([]))
        _set_progress(
            wb, "sc7: FAILED — line one\nline two\nline three "
                + "x" * 200)
        cell = wb.sheets['output']._cells['L2']
        assert '\n' not in cell and len(cell) <= 90

    def test_failed_hint_is_short(self):
        wb = _FakeBook(_grid([
            ['curve_node_1', 'curve_node', 55, 7.601, 'node'],
        ]), main_cells={'V5': 10000, 'V6': 'pw'})
        _main_io_inner(wb)  # no plxscripting → connect-fail path
        cell = wb.sheets['output']._cells['L2']
        assert cell.startswith('sc7: FAILED')
        assert '\n' not in cell and len(cell) <= 90
