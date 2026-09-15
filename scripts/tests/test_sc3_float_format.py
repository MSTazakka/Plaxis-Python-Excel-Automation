"""SC3 float formatting tests (v0.8.5 float-residue fix).

Contract: SC3 soilmat values go through round_num (mirrored from SC5) so
binary residue like 1.5000000000000004 never reaches the notebook.
Quoted strings pass through untouched; unformattable values fall back
to str, matching the old output.

Run with: python -m pytest tests/test_sc3_float_format.py -v
"""
import sys
from pathlib import Path
from unittest.mock import MagicMock

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS_DIR))
sys.modules.setdefault("xlwings", MagicMock())
sys.modules.setdefault("ezdxf", MagicMock())

from sc3_PythonPlaxis import round_num, _fmt_param


def test_residue_float_rounds_clean():
    # User's live case: 1.5 arriving as 1.5000000000000004.
    assert round_num(1.5000000000000004) == "1.5"
    assert _fmt_param(1.5000000000000004) == "1.5"


def test_integers_stay_clean():
    assert round_num(12.0) == "12"
    assert _fmt_param(12.0) == "12"


def test_genuine_decimals_kept_to_3dp():
    assert round_num(0.0164) == "0.016"
    assert round_num(0.082) == "0.082"


def test_quoted_strings_pass_through():
    assert _fmt_param('"Soft Soil Creep"') == '"Soft Soil Creep"'
    assert _fmt_param('"OCR"') == '"OCR"'


def test_unformattable_falls_back_to_str():
    assert _fmt_param(float("nan")) == "nan"
