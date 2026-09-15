"""
sc7_curve_points.py — Curve-point (re)assignment [v0.8.3]
==========================================================
Standalone script (no chain imports — SC1-SC6/SC8/SC9 untouched).

After a PLAXIS remesh, mesh nodes and stress points get new internal
identities, so previously planted curve points are stale. SC7 walks the
existing output-table block (output!B9:F108, SC4-produced, user may append
rows below) and re-plants every valid row as a fresh Output curve point,
top to bottom, exactly as the rows sit. Row order is preserved because
SC8 binds registry rows to live CurvePoints purely by order within type.

 planting order (verbatim from SC5's proven generator):
  viewmesh → new_server → gotostages → selectmeshpoints →
  one addcurvepoint per row → update. No calculation is emitted.

Additive only: existing CurvePoints are never cleared (removal API
unverified; a remesh normally wipes them anyway).

Called from Excel via:
  RunPython "import scripts.sc7_curve_points as l; l.main()"

On each SC7 press:
  1. Reads the curve-point table (B:F rows 9-108; invalid rows warn+skip).
  2. Connects to PLAXIS Input → opens saved model → opens mesh Output.
  3. Plants one addcurvepoint per valid row, in row order.
  4. Writes curve_<stem>.ipynb beside the workbook (run record + replay).
"""

from pathlib import Path
import time
import traceback

import xlwings as xw


# ── Code version banner (diagnostics) ──
# Printed on every run so the console proves which code actually executed —
# PLAXIS caches imported modules in memory, so a stale cache silently runs
# old logic after a .py edit. Bump on every SC7 logic change.
SC7_CODE_VERSION = 'v0.8.3-sc7'

# ── Curve-point table block (mirrors SC4/SC5 constants, duplicated ──
# deliberately: this script is standalone and imports nothing from the chain).
CURVE_SHEET = 'output'
CURVE_DATA_START = 9
CURVE_DATA_END = 108
CURVE_COL_START = 2    # B
CURVE_COL_END = 6      # F
CURVE_VALID_TYPES = ('node', 'stresspoint')

# ── Locator cells (same cells SC8 uses, read-only) ──
MAIN_SHEET = 'main'
PORT_CELL = 'V5'
PASSWORD_CELL = 'V6'
MODEL_CELL = 'L4'      # stem only, .p2dx appended at runtime

# ── Progress cell (shared with SC8, last-run-wins) ──
# output!L2 carries the run status with an sc7: prefix. Same convention as
# the SC6/SC9 shared cell: whoever ran last owns the text. Never raises.
PROGRESS_SHEET = 'output'
PROGRESS_CELL = 'L2'


# ═══════════════════════════════════════════════════════════════════════
#  Helpers (duplicated from chain conventions — standalone by design)
# ═══════════════════════════════════════════════════════════════════════

def safe_str(val, default=''):
    if val is None:
        return default
    return str(val).strip()


def safe_num(val):
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def round_num(val, decimals=3):
    try:
        return round(float(val), decimals)
    except (TypeError, ValueError):
        return val


def _set_progress(wb, msg):
    """Write *msg* to output!L2 as a short single line. Never raises.

    Layout guard: raw exceptions contain newlines — with wrap on they
    blow row 2 tall and stretch the cell-anchored buttons (Curve Points
    / Plaxis output / Start Over). So collapse to one line, cap length,
    force wrap off, and pin row 2 height. Full detail already goes to
    the console + MsgBox + curve_ notebook.
    """
    try:
        cell = wb.sheets[PROGRESS_SHEET].range(PROGRESS_CELL)
        if not msg:
            cell.value = None
            return
        text = str(msg).replace('\r', ' ').replace('\n', ' ')
        text = ' '.join(text.split())
        if len(text) > 90:
            text = text[:87] + '...'
        cell.value = text
        try:
            cell.wrap_text = False
        except Exception:
            pass
        try:
            wb.sheets[PROGRESS_SHEET].range('2:2').row_height = 15.75
        except Exception:
            pass
    except Exception:
        pass


# ── Tier 1 text progress bar (single cell, no layout impact) ──
PROGRESS_BAR_WIDTH = 10


def _bar(frac):
    """Return '██████░░░░ 55%' for frac in [0, 1] (clamped)."""
    frac = max(0.0, min(1.0, float(frac)))
    filled = int(round(frac * PROGRESS_BAR_WIDTH))
    return ('█' * filled + '░' * (PROGRESS_BAR_WIDTH - filled)
            + f' {int(round(frac * 100))}%')


# ═══════════════════════════════════════════════════════════════════════
#  Table read (pure Excel — no PLAXIS needed)
# ═══════════════════════════════════════════════════════════════════════

def read_curve_point_rows(wb):
    """Read curve-point rows from output!B9:F108.

    Returns (rows, skipped) where rows is a list of dicts
    {name, cad_layer, x, y, type, excel_row} in sheet order and skipped
    is a list of (excel_row, reason) for non-blank rows that fail
    validation. Blank rows (no name and no cad_layer) pass silently.
    Never raises — returns ([], []) when the sheet/block is unreadable.
    """
    try:
        sheet = wb.sheets[CURVE_SHEET]
        block = sheet.range(
            (CURVE_DATA_START, CURVE_COL_START),
            (CURVE_DATA_END, CURVE_COL_END),
        ).value
    except Exception:
        return [], []
    if block is None or not isinstance(block, list):
        return [], []

    rows = []
    skipped = []
    for idx, row_vals in enumerate(block):
        excel_row = CURVE_DATA_START + idx
        if not isinstance(row_vals, list):
            row_vals = [row_vals]
        vals = list(row_vals) + [None] * max(0, 5 - len(row_vals))
        name = safe_str(vals[0])
        cad_layer = safe_str(vals[1])
        if not name and not cad_layer:
            continue
        ptype = safe_str(vals[4]).lower()
        if ptype not in CURVE_VALID_TYPES:
            skipped.append(
                (excel_row,
                 f"type '{safe_str(vals[4])}' not node/stresspoint"))
            continue
        x = safe_num(vals[2])
        y = safe_num(vals[3])
        if x is None or y is None:
            skipped.append((excel_row, 'non-numeric x/y'))
            continue
        rows.append({'name': name or f'(row {excel_row})',
                     'cad_layer': cad_layer,
                     'x': x, 'y': y, 'type': ptype,
                     'excel_row': excel_row})
    return rows, skipped


# ═══════════════════════════════════════════════════════════════════════
#  PLAXIS connection (mesh Output for curve-point selection)
# ═══════════════════════════════════════════════════════════════════════

def connect_plaxis(wb):
    """Connect to PLAXIS Input, open the saved model, return mesh Output.

    Returns (s_i, g_i, s_o, g_o). Opens Output on the generated mesh
    (viewmesh) — curve points are planted on mesh geometry, so a mesh must
    exist; when viewmesh fails the error says so. Raises on failure with a
    message the caller prints and shows as FAILED progress.
    """
    from plxscripting.easy import new_server

    main_sheet = wb.sheets[MAIN_SHEET]
    plx_port = int(main_sheet.range(PORT_CELL).value)
    plx_password = str(main_sheet.range(PASSWORD_CELL).value)
    print(f"sc7: port={plx_port}, password={'*' * len(plx_password)}")

    try:
        _full = str(wb.fullname)
    except Exception:
        _full = ''
    model_cell_value = safe_str(
        wb.sheets[CURVE_SHEET].range(MODEL_CELL).value)
    if model_cell_value:
        if model_cell_value.lower().endswith('.p2dx'):
            model_filename = model_cell_value
        else:
            model_filename = f"{model_cell_value}.p2dx"
    elif _full:
        model_filename = f"{Path(_full).stem}.p2dx"
    else:
        raise FileNotFoundError(
            "Cannot derive the model filename — output!L4 is blank and "
            "the workbook path is unavailable.")

    if not _full:
        raise FileNotFoundError(
            "Workbook path unavailable — save the workbook first.")
    model_path = Path(_full).parent / model_filename
    print(f"sc7: model file = {model_path}")
    if not model_path.exists():
        raise FileNotFoundError(
            f"Model not found: {model_path} — the .p2dx file must sit in "
            "the same folder as this workbook.")

    print("sc7: connecting to PLAXIS Input...")
    s_i, g_i = new_server('localhost', port=plx_port, password=plx_password)

    try:
        n_existing = len(g_i.Phases)
        print(f"sc7: model already loaded ({n_existing} phases)")
    except Exception:
        n_existing = 0
    if n_existing == 0:
        print("sc7: opening model file...")
        s_i.open(str(model_path))
        print("sc7: model opened")
    else:
        print("sc7: model already open, skipping s_i.open()")

    # Mesh mode first: without gotomesh() the session may sit in structures
    # or staged-construction mode, where the mesh (and mesh Output) is not
    # viewable. This only switches the view — it never re-meshes.
    try:
        g_i.gotomesh()
        print("sc7: mesh mode (gotomesh)")
    except Exception as e:
        print(f"sc7: gotomesh() failed: {e}")
        raise RuntimeError(
            f"Could not enter mesh mode ({e}) — open the model in PLAXIS "
            "and switch to mesh mode manually, then press Curve Points "
            "again.")

    # Mesh Output (curve points live on mesh geometry — a mesh must exist).
    output_port = None
    last_err = None
    for attempt in range(1, 4):
        try:
            output_port = g_i.viewmesh()
            print(f"sc7: mesh Output port = {output_port} "
                  f"(attempt {attempt})")
            break
        except Exception as e:
            last_err = e
            print(f"sc7: viewmesh() attempt {attempt} failed: {e}")
            time.sleep(3)
    if output_port is None or output_port == 0:
        raise RuntimeError(
            f"Mesh Output did not open after 3 tries ({last_err}) — "
            "has the model been meshed? (SC5 mesh generation, str_2D!K3).")

    s_o, g_o = new_server('localhost', port=output_port,
                          password=plx_password)
    last_probe_err = None
    for attempt in range(1, 4):
        try:
            list(g_o.Phases)
            print(f"sc7: mesh Output ready (attempt {attempt})")
            break
        except Exception as e:
            last_probe_err = e
            print(f"sc7: mesh Output probe attempt {attempt} failed: {e}")
            time.sleep(3)
    else:
        raise RuntimeError(
            f"Mesh Output not ready after 3 tries ({last_probe_err}).")
    return s_i, g_i, s_o, g_o


# ═══════════════════════════════════════════════════════════════════════
#  Planting (SC5's proven call order, verbatim)
# ═══════════════════════════════════════════════════════════════════════

def plant_points(g_i, g_o, rows):
    """Plant one curve point per row, in row order. Additive — never clears.

    Sequence: gotostages → selectmeshpoints → one addcurvepoint per row →
    update. Returns (n_node, n_sp, failures) where failures is a list of
    (excel_row, error) for rows whose addcurvepoint raised. Never raises
    itself — per-row try/except, mirroring the SC8 per-spec guard.
    """
    try:
        g_i.gotostages()
    except Exception as e:
        print(f"sc7: gotostages() failed: {e}")
        return 0, 0, [(-1, f"gotostages: {e}")]
    try:
        g_i.selectmeshpoints()
    except Exception as e:
        print(f"sc7: selectmeshpoints() failed: {e}")
        return 0, 0, [(-1, f"selectmeshpoints: {e}")]

    n_node = n_sp = 0
    failures = []
    for r in (rows or []):
        try:
            g_o.addcurvepoint(r['type'],
                              (round_num(r['x']), round_num(r['y'])))
            if r['type'] == 'node':
                n_node += 1
            else:
                n_sp += 1
        except Exception as e:
            failures.append((r.get('excel_row', '?'), str(e)[:120]))
    try:
        g_o.update()
    except Exception as e:
        print(f"sc7: update() failed: {e}")
        failures.append((-1, f"update: {e}"))
    return n_node, n_sp, failures


# ═══════════════════════════════════════════════════════════════════════
#  Transcript notebook (troubleshooting artifact, never fatal)
# ═══════════════════════════════════════════════════════════════════════

def _write_sc7_transcript_notebook(wb, rows, skipped, n_node, n_sp,
                                   failures):
    """Write the SC7 run transcript as ``curve_<stem>.ipynb``.

    Beside the workbook — deliberately NOT the SC5 Input notebook name
    and NOT SC8's ``output_<stem>.ipynb``, so reruns overwrite only this
    transcript. Payload one: a markdown run record (timestamp, code
    version, rows as read, planted counts, skips/failures). Payload two:
    replay cells — re-read the table (no PLAXIS) and re-fire the planting
    (PLAXIS), so a future session redoes the assignment without
    re-pressing the button. Returns the path string, '' when skipped
    (no fullname). Never raises — the caller wraps it.
    """
    try:
        from datetime import datetime as _dt
    except Exception:
        _dt = None
    try:
        import json as _json
    except Exception:
        return ''
    try:
        _full = safe_str(getattr(wb, 'fullname', ''))
    except Exception:
        _full = ''
    if not _full:
        return ''
    _parent = Path(_full).parent
    _stem = Path(_full).stem
    _nb_path = _parent / f"curve_{_stem}.ipynb"
    _stamp = _dt.now().isoformat(timespec='seconds') if _dt else '?'

    def _md(lines):
        return {'cell_type': 'markdown', 'metadata': {},
                'source': lines}

    def _code(lines):
        return {'cell_type': 'code', 'execution_count': None,
                'metadata': {}, 'outputs': [],
                'source': lines}

    cells = []
    cells.append(_md([
        '# SC7 curve-point run transcript\n',
        f'- workbook: `{_stem}`\n',
        f'- run: `{_stamp}`\n',
        f'- code: `{SC7_CODE_VERSION}`\n',
        f'- planted: {n_node} node + {n_sp} stresspoint '
        f'= {n_node + n_sp} of {len(rows or [])} valid row(s)\n',
    ]))
    _rec = ['## Curve rows as read (output!B9:F108, sheet order)\n']
    for r in (rows or []):
        _rec.append(
            f"- `{r.get('name', '')}` — {r.get('type', '')} @ "
            f"({r.get('x')}, {r.get('y')}) "
            f"[row {r.get('excel_row', '?')}]\n")
    if skipped:
        _rec.append('\n## Skipped rows (non-blank, invalid)\n')
        for _er, _why in skipped:
            _rec.append(f'- row {_er}: {_why}\n')
    if failures:
        _rec.append('\n## Planting failures\n')
        for _er, _why in failures:
            _rec.append(f'- row {_er}: {_why}\n')
    _rec.append(
        '\n## Notes\n'
        '- Additive only: existing CurvePoints are never cleared.\n'
        '- Row order is the planting order — SC8 binds registry rows to '
        'live CurvePoints by order within type, so keep this order.\n'
        '- Reopen this workbook + rerun SC7 before trusting a replay '
        'cell — PLAXIS state (saved model, mesh) may have moved on.\n')
    cells.append(_md(_rec))
    cells.append(_code([
        '# Replay 0 — attach (same book this run used)\n',
        'import xlwings as xw\n',
        f'wb = xw.Book(r"{_full}")\n',
        'from scripts import sc7_curve_points as sc7\n',
        'print("code:", sc7.SC7_CODE_VERSION)\n',
    ]))
    cells.append(_code([
        '# Replay 1 — re-read the table (no PLAXIS needed)\n',
        'rows, skipped = sc7.read_curve_point_rows(wb)\n',
        'print(f"{len(rows)} valid, {len(skipped)} skipped")\n',
        'for r in rows:\n',
        '    print(r["excel_row"], r["name"], r["type"], r["x"], r["y"])\n',
    ]))
    cells.append(_code([
        '# Replay 2 — re-fire the planting (needs PLAXIS + meshed model)\n',
        's_i, g_i, s_o, g_o = sc7.connect_plaxis(wb)\n',
        'n_node, n_sp, failures = sc7.plant_points(g_i, g_o, rows)\n',
        'print(f"planted {n_node} node + {n_sp} stresspoint, '
        '{len(failures)} failures")\n',
    ]))
    _nb = {'nbformat': 4, 'nbformat_minor': 5, 'metadata': {
        'kernelspec': {'display_name': 'Python 3',
                       'language': 'python', 'name': 'python3'},
        'language_info': {'name': 'python', 'version': '3'}},
        'cells': cells}
    with open(str(_nb_path), 'w', encoding='utf-8') as _f:
        _json.dump(_nb, _f, indent=1, ensure_ascii=False)
    return str(_nb_path)


# ═══════════════════════════════════════════════════════════════════════
#  Entry point (SC8 styling: stdout capture + error MsgBox)
# ═══════════════════════════════════════════════════════════════════════

def main():
    """SC7 curve-point (re)assignment entry point.

    Called from Excel via:
      RunPython "import scripts.sc7_curve_points as l; l.main()"
    """
    import sys
    import io

    # Capture stdout so errors are visible even if VBA swallows them.
    _orig_stdout = sys.stdout
    _buf = io.StringIO()
    sys.stdout = _buf

    try:
        try:
            wb = xw.Book.caller()
        except Exception:
            _cwd = Path.cwd()
            _xl_files = list(_cwd.glob('*.xlsm'))
            if not _xl_files:
                _xl_files = list(_cwd.parent.glob('*.xlsm'))
            if not _xl_files:
                raise FileNotFoundError(f"No .xlsm file found in {_cwd}")
            wb = xw.Book(str(_xl_files[0]))

        _main_io(wb)

    except Exception as e:
        traceback.print_exception(type(e), e, e.__traceback__)
    finally:
        sys.stdout = _orig_stdout
        output = _buf.getvalue()
        if output:
            print(output)
        # Show errors in a MsgBox so VBA doesn't swallow them.
        if 'Error' in output or 'Exception' in output or 'Traceback' in output:
            try:
                import win32gui
                # Truncate for MsgBox limit
                msg = output[-2000:] if len(output) > 2000 else output
                win32gui.MessageBox(0, msg, "SC7 Error", 0x10)
            except Exception:
                pass


def _main_io(wb):
    """Core SC7 logic — separated from main() for testability."""
    print(f"sc7: starting curve-point (re)assignment "
          f"[{SC7_CODE_VERSION}]...")
    _set_progress(wb, "sc7: WORKING — reading curve table…")
    try:
        return _main_io_inner(wb)
    finally:
        # Never leave a stale WORKING banner: crash → FAILED, else done
        # text is already set by the success path.
        try:
            cur = wb.sheets[PROGRESS_SHEET].range(PROGRESS_CELL).value
        except Exception:
            cur = None
        if isinstance(cur, str) and 'WORKING' in cur.upper():
            _set_progress(wb, "sc7: FAILED — see console for detail")


def _main_io_inner(wb):
    """Core SC7 logic — separated from main() for testability."""

    # ── 1. Read curve-point table ──
    rows, skipped = read_curve_point_rows(wb)
    if not rows:
        print("sc7: no valid curve-point rows in output!B9:F108")
        _set_progress(wb, "sc7: done — no curve-point rows")
        return
    for _er, _why in skipped:
        print(f"sc7: row {_er} skipped ({_why})")
    _n_node_src = sum(1 for r in rows if r['type'] == 'node')
    _n_sp_src = sum(1 for r in rows if r['type'] == 'stresspoint')
    print(f"sc7: {len(rows)} row(s) "
          f"({_n_node_src} node, {_n_sp_src} stresspoint)"
          + (f", {len(skipped)} skipped" if skipped else ""))

    # ── 2. Connect to PLAXIS (mesh Output) and plant ──
    _set_progress(wb, "sc7: WORKING — connecting to PLAXIS…")
    try:
        s_i, g_i, s_o, g_o = connect_plaxis(wb)
    except Exception as e:
        print(f"sc7: PLAXIS connection failed: {e}")
        traceback.print_exc()
        # Short L2 hint only — full exception is in console + MsgBox.
        # (The raw text carries newlines and stretched the buttons.)
        if 'mesh' in str(e).lower():
            _set_progress(
                wb, "sc7: FAILED — mesh not open (meshed? SC5 K3), "
                    "see console")
        else:
            _set_progress(wb, "sc7: FAILED — connect failed, see console")
        return

    _set_progress(wb, f"sc7: WORKING — planting {_bar(0.10)}")
    n_node, n_sp, failures = plant_points(g_i, g_o, rows)
    for _er, _why in failures:
        print(f"sc7: planting failed at row {_er} ({_why})")

    if failures and (n_node + n_sp) == 0:
        _set_progress(wb, "sc7: FAILED — nothing planted, see console")
        return
    _note = (f", {len(failures)} failed" if failures else "")
    _set_progress(
        wb, f"sc7: done — {n_node + n_sp} point(s) "
            f"({n_node} node, {n_sp} stresspoint){_note} {_bar(1.0)}")
    print(f"sc7: planted {n_node} node + {n_sp} stresspoint{_note}")

    # ── 3. Transcript notebook (troubleshooting artifact, never fatal) ──
    # curve_<stem>.ipynb beside the workbook: run record + replay cells.
    try:
        _nb_path = _write_sc7_transcript_notebook(
            wb, rows, skipped, n_node, n_sp, failures)
        if _nb_path:
            print(f"sc7: transcript notebook: {_nb_path}")
    except Exception as _e:
        print(f"sc7: transcript notebook skipped ({_e})")


if __name__ == '__main__':
    main()
