"""Groundwater USDA emission tests (v0.8.5 groundwater fix).

Covers the all-USDA groundwater contract shared by sc3/sc5 (/sc2 internal):

1. _normalise_gw_mode(): blank/None/NaN -> 'automatic'; truncated
   'by_grain_s' cell aliases to 'by_grain_size'; whitespace stripped.
2. _gw_frac_is_filled(): None/NaN/''/whitespace -> False (omit the
   fraction line so PLAXIS falls back to the class default); explicit
   0 -> True (0 still emits — PLAXIS distinguishes 0 from default).
3. gen_soil_mat() (sc5): automatic writes no fraction lines;
   by_grain_size/manual write each fraction line ONLY when that cell is
   filled; manual always carries USDA class + permeability lines;
   class string is the SOIL_TYPE_MAP value verbatim (case-safe).
4. Byte-identical regression: filled by_grain_size emits exactly the
   pre-patch block (no behaviour change for filled AE/AF layers).

Run with: python -m pytest tests/test_groundwater_usda.py -v
"""
import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS_DIR))
sys.modules.setdefault("xlwings", MagicMock())
sys.modules.setdefault("ezdxf", MagicMock())

from sc3_PythonPlaxis import _gw_frac_is_filled as sc3_filled
from sc3_PythonPlaxis import _normalise_gw_mode as sc3_mode
from sc5_structural_to_ipynb import _gw_frac_is_filled as sc5_filled
from sc5_structural_to_ipynb import _normalise_gw_mode as sc5_mode
from sc5_structural_to_ipynb import gen_soil_mat


def _base(**kw):
    d = dict(soil_layers="L1", type="clay-HS", gunsat=18, gsat=20,
             E50_ref=10000, Eoed_ref=10000, Eur_ref=30000, m=0.8, c=5,
             phi=25, psi=0, pref=100, Rinter=0.7, k_ver=0.01,
             k_hor_ratio=1.5, mode="automatic",
             clay_frac=None, silt_frac=None)
    d.update(kw)
    return d


# ── 1. mode normalisation (both modules) ──

@pytest.mark.parametrize("fn", [sc3_mode, sc5_mode])
@pytest.mark.parametrize("inp,exp", [
    ("by_grain_s", "by_grain_size"),
    ("by_grain_size", "by_grain_size"),
    ("manual", "manual"),
    ("automatic", "automatic"),
    ("", "automatic"),
    (None, "automatic"),
    (float("nan"), "automatic"),
    ("  manual  ", "manual"),
])
def test_mode_normalisation(fn, inp, exp):
    assert fn(inp) == exp


# ── 2. blank gate (both modules) ──

@pytest.mark.parametrize("fn", [sc3_filled, sc5_filled])
@pytest.mark.parametrize("inp,exp", [
    (None, False),
    (float("nan"), False),
    ("", False),
    ("   ", False),
    (0, True),
    (0.0, True),
    (12.5, True),
    ("7", True),
])
def test_frac_blank_gate(fn, inp, exp):
    assert fn(inp) is exp


# ── 3. gen_soil_mat emission ──

def _src(**kw):
    return "\n".join(gen_soil_mat(1, _base(**kw)))


def test_automatic_no_fractions():
    s = _src(mode="automatic")
    assert 'GroundwaterClassificationType.set("USDA")' in s
    assert 'GroundwaterSoilClassUSDA.set("Clay")' in s
    assert "ClayFraction" not in s and "SiltFraction" not in s
    assert "GwUseDefaults.set(True)" in s
    assert "From grain size distribution" in s


def test_by_grain_size_filled():
    s = _src(mode="by_grain_size", clay_frac=10, silt_frac=20)
    assert "ClayFraction.set(10)" in s
    assert "SiltFraction.set(20)" in s
    assert "GwUseDefaults.set(True)" in s


def test_by_grain_size_both_blank_no_fractions():
    s = _src(mode="by_grain_size", clay_frac=None, silt_frac="")
    assert "ClayFraction" not in s and "SiltFraction" not in s
    assert 'GroundwaterSoilClassUSDA.set("Clay")' in s
    assert "GwUseDefaults.set(True)" in s


def test_truncated_cell_alias():
    s = _src(mode="by_grain_s", clay_frac=10, silt_frac=20)
    assert "ClayFraction.set(10)" in s
    assert "SiltFraction.set(20)" in s


def test_manual_filled_usda_plus_perms():
    s = _src(mode="manual", clay_frac=5, silt_frac=15)
    assert 'GroundwaterClassificationType.set("USDA")' in s
    assert 'GroundwaterSoilClassUSDA.set("Clay")' in s
    assert "ClayFraction.set(5)" in s
    assert "SiltFraction.set(15)" in s
    assert "GwUseDefaults.set(False)" in s
    assert "PermHorizontalPrimary.set(" in s
    assert "PermVertical.set(0.01)" in s


def test_manual_blank_no_fractions():
    s = _src(mode="manual", clay_frac=None, silt_frac=None)
    assert 'GroundwaterSoilClassUSDA.set("Clay")' in s
    assert "ClayFraction" not in s and "SiltFraction" not in s
    assert "GwUseDefaults.set(False)" in s


def test_half_filled_emits_only_filled():
    s = _src(mode="by_grain_size", clay_frac=10, silt_frac=None)
    assert "ClayFraction.set(10)" in s
    assert "SiltFraction" not in s


def test_explicit_zero_still_emits():
    s = _src(mode="by_grain_size", clay_frac=0, silt_frac=0)
    assert "ClayFraction.set(0)" in s
    assert "SiltFraction.set(0)" in s


def test_class_casing_is_map_value_verbatim():
    s = _src(type="sand-HS", mode="manual")
    assert 'GroundwaterSoilClassUSDA.set("Sand")' in s


# ── 4. byte-identical regression for filled layers ──

def test_filled_grain_block_byte_identical():
    lines = gen_soil_mat(1, _base(mode="by_grain_size",
                                  clay_frac=10, silt_frac=20))
    gw = [x for x in lines if "Groundwater" in x or "Fraction" in x
          or "GwUse" in x or "GwDefaults" in x]
    assert gw == [
        'L1.GroundwaterClassificationType.set("USDA")',
        'L1.GroundwaterSoilClassUSDA.set("Clay")',
        'L1.ClayFraction.set(10)',
        'L1.SiltFraction.set(20)',
        'L1.GwUseDefaults.set(True)',
        'L1.GwDefaultsMethod.set("From grain size distribution")',
    ]
