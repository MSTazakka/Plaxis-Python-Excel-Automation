"""SC5 column-C [refine f] mesh refinement tests.

Covers the v0.8.5 soil-polygon refinement gesture against the probe-proven
route (dead_plan/probe_polygon8_refine.ipynb):

1. _parse_refine_bracket(): case-insensitive parse, strip-clean, non-refine
   brackets pass through, garbage/zero/negative/out-of-range/conflicting
   values fail LOUD (ValueError, never silent skip).
2. read_soil_polygons(): single tagged row promotes the whole polygon_no
   group; identical sibling values agree; differing values fail fast;
   stripped layer never carries the bracket downstream.
3. collect_refinements() + gen_refine_code(): sorted output, single-level
   setproperties('CoarsenessFactor', f) emission (the _set equivalent —
   nested attribute writes are blind to the mesher), tag-substring census,
   no mesh call inside refinement cells.
4. _write_refine_echo(): group-scoped, idempotent (re-press converges).

Run with: python -m pytest tests/test_sc5_refine.py -v
"""
import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS_DIR))
sys.modules.setdefault("xlwings", MagicMock())
sys.modules.setdefault("ezdxf", MagicMock())

from sc5_structural_to_ipynb import (
    _parse_refine_bracket,
    _write_refine_echo,
    collect_refinements,
    gen_refine_code,
    read_soil_polygons,
)


# ── 1. bracket parser ──

def test_parse_basic_and_strip():
    clean, factor = _parse_refine_bracket('volume_profile[refine 0.2]')
    assert clean == 'volume_profile'
    assert factor == pytest.approx(0.2)


def test_parse_case_insensitive_and_spaced():
    clean, factor = _parse_refine_bracket('cut_4[REFINE   0.5]')
    assert clean == 'cut_4'
    assert factor == pytest.approx(0.5)


def test_parse_no_bracket_passthrough():
    assert _parse_refine_bracket('cut_4') == ('cut_4', None)
    assert _parse_refine_bracket('') == ('', None)


def test_parse_non_refine_bracket_untouched():
    # plate custom_wall convention must never be disturbed
    text = 'plate_1[combi_wall - OD1200]'
    assert _parse_refine_bracket(text) == (text, None)


def test_parse_refinement_word_not_a_tag():
    # \b guard: 'refinement' is a different word, passes through
    text = 'layer[refinement 0.2]'
    assert _parse_refine_bracket(text) == (text, None)


@pytest.mark.parametrize('bad', [
    'x[refine abc]', 'x[refine]', 'x[refine ]', 'x[refine 0]',
    'x[refine -0.2]', 'x[refine 1.5]', 'x[refine 0.01]',
])
def test_parse_garbage_fails_loud(bad):
    with pytest.raises(ValueError):
        _parse_refine_bracket(bad, row_label='str_2D!C250')


def test_parse_boundary_values_accepted():
    _, lo = _parse_refine_bracket('x[refine 0.05]')
    _, hi = _parse_refine_bracket('x[refine 1.0]')
    assert lo == pytest.approx(0.05)
    assert hi == pytest.approx(1.0)


def test_parse_duplicate_same_value_agrees():
    clean, factor = _parse_refine_bracket('x[refine 0.2] tail [refine 0.2]')
    assert clean == 'x tail'
    assert factor == pytest.approx(0.2)


def test_parse_duplicate_conflict_fails_loud():
    with pytest.raises(ValueError):
        _parse_refine_bracket('x[refine 0.2][refine 0.3]')


# ── fake sheet harness ──

def _row(name, layer, x1, y1, x2, y2, ptype, pno, mat='', mata=''):
    return [name, layer, x1, y1, x2, y2, ptype, pno, mat, mata]


class FakeSheet:
    """Minimal xlwings sheet stub: block reads + cell get/set."""

    def __init__(self, rows, max_rows=50):
        self._rows = rows  # list of B..K row lists, starting at row 235
        self._cells = {(233, 9): max_rows}
        self.writes = {}

    def range(self, *args):
        box = self

        class _R:
            @property
            def value(self):
                if len(args) == 2 and isinstance(args[0], int):
                    return box._cells.get((args[0], args[1]))
                (r1, _), (r2, _) = args[0], args[1]
                out = []
                for r in range(r1, r2 + 1):
                    i = r - 235
                    out.append(list(box._rows[i]) if 0 <= i < len(box._rows)
                               else [None] * 10)
                return out

            @value.setter
            def value(self, v):
                box.writes[(args[0], args[1])] = v
                box._cells[(args[0], args[1])] = v

        return _R()


def _apply_writes_to_rows(sheet):
    for (r, c), v in sheet.writes.items():
        if c == 3:
            sheet._rows[r - 235][1] = v


def _poly12_rows(layer_fn):
    # Polygon 12, two segments; tag only the first row
    return [
        _row('sp12_a', layer_fn(0), 0.0, 0.0, 1.0, 0.0, 'cut', 12),
        _row('sp12_b', layer_fn(1), 1.0, 0.0, 1.0, 1.0, 'cut', 12),
        _row('sp9_a', 'cut_1', 5.0, 5.0, 6.0, 5.0, 'cut', 9),
    ]


# ── 2. reader promotion ──

def test_reader_promotes_group_and_strips_layer():
    sheet = FakeSheet(_poly12_rows(
        lambda i: 'volume_profile[refine 0.2]' if i == 0 else 'volume_profile'))
    polys = read_soil_polygons(sheet)
    p12 = next(p for p in polys if p['polygon_no'] == 12)
    p9 = next(p for p in polys if p['polygon_no'] == 9)
    assert p12['refine_factor'] == pytest.approx(0.2)
    assert p9['refine_factor'] is None
    assert p12['layer'] == 'volume_profile'  # bracket stripped pre-layer-logic
    assert '[' not in p12['layer']


def test_reader_identical_siblings_agree():
    sheet = FakeSheet(_poly12_rows(lambda i: 'cut_4[refine 0.3]'))
    polys = read_soil_polygons(sheet)
    p12 = next(p for p in polys if p['polygon_no'] == 12)
    assert p12['refine_factor'] == pytest.approx(0.3)


def test_reader_conflict_fails_fast():
    sheet = FakeSheet(_poly12_rows(
        lambda i: f'cut_4[refine {0.2 if i == 0 else 0.4}]'))
    with pytest.raises(ValueError):
        read_soil_polygons(sheet)


def test_reader_garbage_fails_loud():
    sheet = FakeSheet(_poly12_rows(
        lambda i: 'cut_4[refine abc]' if i == 0 else 'cut_4'))
    with pytest.raises(ValueError):
        read_soil_polygons(sheet)


def test_reader_tag_without_polygon_fails_loud():
    rows = [_row('', 'x[refine 0.2]', 0.0, 0.0, 1.0, 0.0, 'cut', 0)]
    with pytest.raises(ValueError):
        read_soil_polygons(FakeSheet(rows))


def test_reader_tag_without_polygon_no_fails_loud():
    rows = [_row('sp_x', 'x[refine 0.2]', 0.0, 0.0, 1.0, 0.0, 'cut', 0)]
    with pytest.raises(ValueError):
        read_soil_polygons(FakeSheet(rows))


# ── 3. collector + emitter ──

def test_collect_sorted():
    polys = [{'polygon_no': 12, 'refine_factor': 0.2},
             {'polygon_no': 3, 'refine_factor': None},
             {'polygon_no': 8, 'refine_factor': 0.5}]
    assert collect_refinements(polys) == [(8, 0.5), (12, 0.2)]
    assert collect_refinements([]) == []
    assert collect_refinements(None) == []


def test_collect_dedupes_shared_polygon_no():
    # two spoly_name groups sharing polygon_no → one emitted block
    polys = [{'polygon_no': 12, 'refine_factor': 0.2},
             {'polygon_no': 12, 'refine_factor': 0.2}]
    assert collect_refinements(polys) == [(12, 0.2)]


def test_emitter_uses_set_equivalent_route():
    lines = gen_refine_code([(8, 0.2)])
    src = '\n'.join(lines)
    assert 'g_i.Polygons' in src  # mesh-mode tag census
    assert "_Polygon_8_" in src
    assert "setproperties('CoarsenessFactor', 0.2)" in src  # single level
    assert 'CoarsenessFactor.CoarsenessFactor' not in src  # nested is dead
    assert 'g_i.mesh(' not in src  # no mesh call inside refinement
    assert 'gotomesh' not in src  # caller owns the mode switch
    assert 'Parent' not in src  # CutObjects expose no Parent
    assert 'get_equivalent' not in src  # census replaced the partial mapping


def test_emitter_slice_count_reported():
    lines = gen_refine_code([(12, 0.3)])
    src = '\n'.join(lines)
    assert '_refn_12 += 1' in src and 'slice(s)' in src


def test_emitter_namespaced_two_polygons():
    lines = gen_refine_code([(8, 0.2), (12, 0.3)])
    src = '\n'.join(lines)
    assert '_reftag_8' in src and '_reftag_12' in src
    assert '_refn_8' in src and '_refn_12' in src
    # no shared _tag/_n loop variable that later blocks would overwrite
    assert '\n_tag = ' not in src and '\n_n = 0' not in src
    compile(src, '<emitted>', 'exec')  # emitted code is valid Python


def _census_match(src, names, p_no):
    """Replay the emitted membership condition against fake cluster names."""
    ns = {'_refhead_%d' % p_no: 'Polygon_%d_' % p_no,
          '_reftag_%d' % p_no: '_Polygon_%d_' % p_no}
    cond = [ln.strip() for ln in src.splitlines()
            if ln.strip().startswith('if _refnm.startswith')][0][3:]
    cond = cond[:-1] if cond.endswith(':') else cond
    return [n for n in names
            if eval(cond, {}, dict(ns, _refnm=n))]  # noqa: S307 — test-only eval


def test_emitter_dual_detection_covers_both_families():
    lines = gen_refine_code([(9, 0.2)])
    src = '\n'.join(lines)
    assert 'startswith' in src  # standalone arm present
    names = ['Polygon_9_2',  # standalone above-surface cluster
             'BoreholePolygon_1_Polygon_9_3_1',  # buried fragment
             'Polygon_19_1',  # trailing-digit near-miss
             'BoreholePolygon_1_Polygon_19_1',  # buried near-miss
             'Polygon_90_1']  # prefix near-miss
    assert _census_match(src, names, 9) == [
        'Polygon_9_2', 'BoreholePolygon_1_Polygon_9_3_1']


def test_emitter_polygon1_immune_to_borehole_prefix():
    # 'BoreholePolygon_1_…' carries 'e' (not '_') before 'Polygon' —
    # the substring arm must not fire, only true Polygon_1_* names match.
    lines = gen_refine_code([(1, 0.2)])
    src = '\n'.join(lines)
    names = ['BoreholePolygon_1_Polygon_8_1_1',  # another polygon's fragment
             'BoreholePolygon_1_2',  # bare borehole cluster, not polygon 1
             'Polygon_1_3',  # genuine standalone
             'BoreholePolygon_1_Polygon_1_2_1']  # genuine buried fragment
    assert _census_match(src, names, 1) == [
        'Polygon_1_3', 'BoreholePolygon_1_Polygon_1_2_1']


# ── 4. echo-back ──

def test_echo_group_scoped_and_idempotent():
    sheet = FakeSheet(_poly12_rows(
        lambda i: 'volume_profile[refine 0.2]' if i == 0 else 'volume_profile'))
    polys = read_soil_polygons(sheet)
    refined = collect_refinements(polys)
    assert refined == [(12, 0.2)]
    _write_refine_echo(sheet, polys, refined)
    _apply_writes_to_rows(sheet)
    c_vals = [r[1] for r in sheet._rows]
    # both polygon-12 rows carry the suffix now …
    assert c_vals[0] == 'volume_profile [refine 0.2]'
    assert c_vals[1] == 'volume_profile [refine 0.2]'
    # … while polygon 9 is untouched
    assert c_vals[2] == 'cut_1'
    # second press converges (no stacking)
    _write_refine_echo(sheet, polys, refined)
    _apply_writes_to_rows(sheet)
    assert sheet._rows[0][1] == 'volume_profile [refine 0.2]'
    assert sheet._rows[1][1] == 'volume_profile [refine 0.2]'
    assert sheet._rows[0][1].count('[refine') == 1


def test_echo_no_refinements_no_writes():
    sheet = FakeSheet(_poly12_rows(lambda i: 'cut_4'))
    _write_refine_echo(sheet, read_soil_polygons(sheet), [])
    assert sheet.writes == {}
