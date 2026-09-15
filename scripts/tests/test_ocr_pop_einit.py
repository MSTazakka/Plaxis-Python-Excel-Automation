"""OCR/POP/einit emission tests (v0.8.5 overconsolidation fix).

Contract (HS/HSS/SSC): each of OCR/POP/einit emits independently iff its
own cell is filled (same blank gate as the groundwater fractions). Blank
omits that pair; filled emits the input value. SSC lambda/kappa/mu always
emit. POP stays trailing so a both-filled row keeps POP last.

Run with: python -m pytest tests/test_ocr_pop_einit.py -v
"""
import sys
from pathlib import Path
from unittest.mock import MagicMock

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS_DIR))
sys.modules.setdefault("xlwings", MagicMock())
sys.modules.setdefault("ezdxf", MagicMock())

from sc5_structural_to_ipynb import gen_soil_mat


def _base(**kw):
    d = dict(soil_layers="L1", type="sand-HS", gunsat=18, gsat=20,
             E50_ref=10000, Eoed_ref=10000, Eur_ref=30000, m=0.8, c=5,
             phi=25, psi=0, pref=100, Rinter=0.7, k_ver=0.01,
             k_hor_ratio=1.5, mode="automatic",
             clay_frac=None, silt_frac=None, OCR=None, POP=None,
             einit=None, **{k: None for k in ("lambda", "kappa", "mu")})
    d.update(kw)
    return d


def _src(**kw):
    return "\n".join(gen_soil_mat(1, _base(**kw)))


def test_blank_hs_omits_all_three():
    s = _src()
    assert '"OCR"' not in s
    assert '"POP"' not in s
    assert '"eInit"' not in s


def test_ocr_filled_only_ocr_emits():
    # User's exact case: OCR=1, POP/einit blank -> only "OCR" 1.
    s = _src(OCR=1)
    assert '"OCR", 1' in s
    assert '"POP"' not in s
    assert '"eInit"' not in s


def test_all_three_filled():
    s = _src(OCR=1, POP=50, einit=0.8)
    assert '"OCR", 1' in s
    assert '"POP", 50' in s
    assert '"eInit", 0.8' in s


def test_blank_hss_omits():
    s = _src(type="sand-HSS")
    assert '"OCR"' not in s
    assert '"POP"' not in s
    assert '"eInit"' not in s


def test_hss_mixed_emits_independently():
    s = _src(type="sand-HSS", OCR=1, einit=1.5)
    assert '"OCR", 1' in s
    assert '"eInit", 1.5' in s
    assert '"POP"' not in s


def test_ssc_blank_ocr_filled_pop_omits_ocr():
    # User's screenshot case: blank OCR + POP=12 -> no raw-nan OCR pair.
    s = _src(type="clay-SSC", POP=12, einit=1.2)
    assert '"OCR"' not in s
    assert 'nan' not in s
    assert '"POP", 12' in s
    assert '"eInit", 1.2' in s
    assert '"lambdaModified", 0' in s
    assert '"kappaModified", 0' in s
    assert '"muModified", 0' in s


def test_ssc_blank_all_three_omits():
    s = _src(type="clay-SSC")
    assert '"OCR"' not in s
    assert '"POP"' not in s
    assert '"eInit"' not in s
    assert 'nan' not in s
    assert '"lambdaModified", 0' in s
    assert '"kappaModified", 0' in s
    assert '"muModified", 0' in s


def test_ssc_both_filled_pop_trailing():
    s = _src(type="clay-SSC", OCR=1, POP=50, einit=0.8)
    assert '"OCR", 1' in s
    assert '"POP", 50' in s
    assert '"eInit", 0.8' in s
    assert s.index('"POP"') > s.index('"OCR"')
