"""
sc6_staged_construction.py — Assign Sequence → .ipynb [v0.7.9]
============================================================
Reads 'assign_sequence' and 'str_2D' Excel tabs via xlwings and generates
staged construction Python code, appending to the workbook-scoped .ipynb.

Phase boundary: row 1 in assign_sequence.
  - Row 1: calc type (DeformCalcType)
  - Row 2: time interval
  - Row 3: upd_mesh (yes/no)
  - Row 6: phase name (Identify)
  - Rows 7+: markers: A, D, D,Dry, A,Dry, Dry, Interp

Dewatering logic (soil polygons, v0.6.17):
  1. Read str_2D into SOURCE_BOUNDS (bounding rectangles per polygon_no).
  2. Build excavation stacks via X-range overlap (union-find), NOT Xmid.
     This handles zigzag profiles where polygons share left boundary
     but have different right bounds.
  3. For each D,Dry / A,Dry / Dry marker on a soil polygon, find target
     polygon's stack: target + above → Dry, below → Interpolate.
     Multiple Dry targets in same stack → use deepest, warn user.
  4. Water conditions emitted via set_water_for_source_polygon() which
     matches PLAXIS polygons by name tag (_Polygon_N_) or BBox fallback.
     Stacks are SOURCE_BOUNDS stacks (polygon_no-based), not runtime stacks.

Tunnel water conditions (v0.6.17):
  Tunnel opening/liner rows (row_type starting with 'tunnel_') use direct
  WaterConditions calls on the BBox-discovered proxy variables, NOT
  SOURCE_BOUNDS excavation stacks.

Tunnel identification (v0.6.16):
  After gotostages(), BBox identification code scans g_i.SoilPolygons,
  matches by distance to tunnel center, sorts by BBox area,
  and assigns tunnel_N_opening (smallest) and tunnel_N_liners (rest).

Called from Excel via:
  RunPython "import scripts.sc6_staged_construction as l; l.main()"
"""
from pathlib import Path
import json
import re
from collections import defaultdict

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

import xlwings as xw

try:
    from .notebook_paths import notebook_path_for_workbook
except ImportError:  # pragma: no cover - supports direct script execution
    from notebook_paths import notebook_path_for_workbook

SECTION_TAG = '# ── Staged Construction (from sc6) ──'
TUNNEL_ID_TAG = '# ── Tunnel identification (from sc5) ──'
SOIL_POLY_ID_TAG = '# ── Soil polygon identification (from sc6) ──'

CALC_TYPE_NORM = {
    'k0procedure': 'k0procedure', 'k0': 'k0procedure',
    'fieldstress': 'fieldstress', 'field': 'fieldstress',
    'gravityloading': 'gravityloading', 'gravity': 'gravityloading',
    'flowonly': 'flowonly', 'flow': 'flowonly',
    'plastic': 'plastic',
    'consolidation': 'consolidation', 'consolid': 'consolidation',
    'safety': 'safety', 'dynamic': 'dynamic',
}


# ── Progress cell (SC6, mirrors SC8 pattern) ──
# assign_sequence!B5 is blank in the template (verified 2026-09-08 on
# v0.7.14 EXAMPLE 2: B5/D2/D4 blank, SC5 table starts at B6). Shared
# with SC9: SC6 hands B5 to SC9 in the FULL chain (scFULL button).
# Written live during staged-construction generation. main() disables
# screen_updating for speed, so each write briefly re-enables it so the
# cell repaints, then restores the previous state. Never raises —
# progress must not break runs.
PROGRESS_SHEET = 'assign_sequence'
PROGRESS_CELL = 'B5'
PROGRESS_BAR_WIDTH = 10
_SPIN_CHARS = ('|', '/', '-', '\\')
_spin_idx = 0


def _spin():
    """Return next spinner character (hybrid indeterminate pulse)."""
    global _spin_idx
    ch = _SPIN_CHARS[_spin_idx % len(_SPIN_CHARS)]
    _spin_idx += 1
    return ch


def _bar(frac):
    """Return '██████░░░░ 55%' for frac in [0, 1] (clamped)."""
    frac = max(0.0, min(1.0, float(frac)))
    filled = int(round(frac * PROGRESS_BAR_WIDTH))
    return ('█' * filled + '░' * (PROGRESS_BAR_WIDTH - filled)
            + f' {int(round(frac * 100))}%')


def _set_progress(wb, msg):
    """Write *msg* to assign_sequence!B5. False/None clears. Never raises."""
    try:
        sheet = wb.sheets[PROGRESS_SHEET]
        try:
            app = wb.app
            prev = app.screen_updating
            app.screen_updating = True
            sheet.range(PROGRESS_CELL).value = (
                msg if msg else None)
            app.screen_updating = prev
        except Exception:
            sheet.range(PROGRESS_CELL).value = (
                msg if msg else None)
    except Exception:
        pass


# ── Helpers ──────────────────────────────────────────────────────────

def normalize_calc_type(raw_value):
    if raw_value is None or str(raw_value).strip() == '':
        return 'plastic'
    n = str(raw_value).strip().lower().replace(' ', '').replace('_', '').replace('-', '')
    return CALC_TYPE_NORM.get(n, 'plastic')


def _normalize_bool(raw_value, default=False):
    if raw_value is None:
        return default
    v = str(raw_value).strip().lower()
    if v in ('yes', 'true', '1', 'y'):
        return True
    if v in ('no', 'false', '0', 'n'):
        return False
    return default


def _parse_marker(raw_value):
    """Parse marker cell -> (action, water, prestress, user_pressure, steady_state, warn).
    action: 'activate','deactivate',None
    water: 'Dry','Dewatering','Interp',None
    prestress: float or None (absolute kN value from PXXX) — N2N/FEA only
    user_pressure: positive float or None (kPa magnitude from UXXX) — tunnel only
    steady_state: True if SS marker present
    warn: str or None (warning message for malformed U)"""
    if raw_value is None:
        return None, None, None, None, False, None
    parts = [p.strip() for p in str(raw_value).strip().split(',')]
    action, water, prestress, user_pressure, warn = None, None, None, None, None
    steady_state = False
    raw_str = str(raw_value).strip()
    for p in parts:
        pu = p.upper()
        if pu in ('A', 'ACTIVATE'):
            action = 'activate'
        elif pu in ('D', 'DEACTIVATE'):
            action = 'deactivate'
        elif pu == 'SS':
            steady_state = True
        elif pu == 'DRY':
            water = 'Dry'
        elif pu in ('DEWATERING', 'DEWATER'):
            water = 'Dewatering'
        elif pu == 'INTERP':
            water = 'Interp'
        elif pu.startswith('U'):
            num_str = pu[1:]
            if num_str == '' or not all(c.isdigit() or c == '.' for c in num_str):
                warn = f'Malformed user pressure marker "{raw_str}"'
            else:
                try:
                    user_pressure = float(num_str)
                    if user_pressure <= 0:
                        user_pressure = None
                        warn = f'Non-positive user pressure rejected: "{raw_str}"'
                except ValueError:
                    warn = f'Malformed user pressure marker "{raw_str}"'
        elif pu.startswith('P'):
            num_str = pu[1:]
            if num_str == '' or not all(c.isdigit() or c == '.' for c in num_str):
                warn = f'Malformed prestress marker "{raw_str}"'
            else:
                try:
                    prestress = float(num_str)
                    if prestress < 0:
                        prestress = None
                        warn = f'Negative prestress rejected: "{raw_str}"'
                except ValueError:
                    warn = f'Malformed prestress marker "{raw_str}"'
    return action, water, prestress, user_pressure, steady_state, warn


def _is_output_paint_label(label):
    """True for structural labels painted green in the output table.

    Structural elements only (ordinary plates, tunnels, geogrids,
    embedded beams, anchors) — never soil polygons. Mirrors the SC5
    output-table gate plus tunnel_volume liner proxies.
    """
    text = str(label or '').strip().lower()
    if text.startswith(('plate_', 'embeddedbeam_', 'nodetonodeanchor_',
                        'fixedendanchor_', 'geogrid_')):
        return True
    import re as _re
    return _re.match(r'^tunnel_\d+_liners\[\d+\]$', text) is not None


def _paint_output_activation_window(wb, phase_cols, marker_block):
    """Paint output-table arg cells green from first A to last used phase.

    Runs at SC6 press time off the already-read assign data: for each
    structural row, the first phase column holding an 'A' marker starts
    the window; the last phase column in assign_sequence ends it.
    Matching output rows (by PLAXIS label, column L) get their arg
    cells (M onward) filled green across that span — data is
    subtractable there. Rows with no A stay unpainted. Only fill and
    clear inside the table footprint; arg strings never change.
    Never raises.
    """
    try:
        import re as _re
        try:
            out_sheet = wb.sheets['output']
        except Exception:
            return
        if not phase_cols or not marker_block:
            return
        last_assign_col = max(col for col, *_ in phase_cols)
        # Output phase columns sit exactly +1 right of assign phase
        # columns (assign L=12 ↔ output M=13 — same rule as SC8's
        # phase filter: {c + 1}).
        # Output table footprint: header row 8, labels in K:L (11:12),
        # phase args from M (13) onward.
        try:
            head = out_sheet.range((8, 11), (8, 60)).value
        except Exception:
            return
        if head is None:
            return
        head = head if isinstance(head, list) else [head]
        out_labels = {}
        try:
            rows = out_sheet.range((9, 11), (1008, 12)).value
        except Exception:
            return
        if rows is None:
            return
        if rows and not isinstance(rows[0], list):
            rows = [rows]
        for i, rv in enumerate(rows):
            rv = rv if isinstance(rv, list) else [rv]
            lab = str(rv[1] if len(rv) > 1 else '' or '').strip()
            if lab:
                out_labels.setdefault(lab, 9 + i)
        GREEN = 0xC6EFCE  # Excel light-green fill (BGR)
        for row_offset, row_vals in enumerate(marker_block):
            if not isinstance(row_vals, list):
                continue
            label_raw = row_vals[2] if len(row_vals) > 2 else None
            if label_raw is None or str(label_raw).strip() == '':
                continue
            label = str(label_raw).strip()
            if not _is_output_paint_label(label):
                continue
            first_a = None
            for col, _pn, *_rest in phase_cols:
                idx = col - 2
                if idx >= len(row_vals):
                    continue
                action, *_ = _parse_marker(row_vals[idx])
                if action == 'activate':
                    first_a = col
                    break
            if first_a is None:
                continue
            out_row = out_labels.get(label)
            if out_row is None:
                continue
            c0 = first_a + 1
            c1 = last_assign_col + 1
            try:
                rng = out_sheet.range((out_row, c0), (out_row, c1))
                rng.api.Interior.Color = GREEN
            except Exception:
                continue
        print(f"sc6: output activation windows painted "
              f"(first-A → last phase, structural rows only)")
    except Exception:
        pass


def _is_prestress_supported_label(label):
    """Return whether a staged P marker is valid for this element type.

    In PLAXIS 2D, staged prestress is supported for Node-to-Node anchors
    and Fixed-End anchors.  The prestress target is the generated ``_1``
    proxy (for example, ``NodeToNodeAnchor_1_1``).
    """
    text = str(label).strip()
    return text.startswith(('NodeToNodeAnchor_', 'FixedEndAnchor_'))


def _emit_prestress_lines(label, phase, prestress):
    """Return staged prestress assignments for a supported anchor label.

    Marker values are positive magnitudes; PLAXIS receives the negative
    force value used by this project.  The ``_1`` target suffix is required
    by the current PLAXIS API proxy for N2N and fixed-end anchors.
    """
    if prestress is None or not _is_prestress_supported_label(label):
        return []
    ps_lbl = f'{label}_1'
    return [
        f'g_i.{ps_lbl}.AdjustPrestress[{phase}] = True',
        f'g_i.{ps_lbl}.PrestressForce[{phase}] = {-prestress:g}',
    ]


def _emit_tunnel_user_pressure(label, phase, pressure):
    """Emit direct proxy commands for a user-defined tunnel pressure.

    The marker stores a positive pressure magnitude; PLAXIS Pref uses the
    negative sign convention in this model.
    """
    if pressure is None or pressure <= 0:
        return []
    return [
        f'{label}.Soil.WaterConditions.Conditions[{phase}] = "User-defined"',
        f'{label}.Soil.WaterConditions.Pref[{phase}] = {-pressure:g}',
    ]


def _extract_polygon_number(label):
    """Extract polygon number from label. Accepts 'Polygon_21', '21', 21."""
    text = str(label).strip()
    m = re.search(r'(\d+)', text)
    return int(m.group(1)) if m else None


def _clean_mat_var_name(tunnel_name):
    """Convert tunnel name string to a valid Python variable name for material.

    Examples:
        'tunnel3_primary_THK=0.5m'  -> 'tunnel3_primary_THK0_5m'
        'tunnel_secondary_THK=0.3m' -> 'tunnel_secondary_THK0_3m'
    """
    name = tunnel_name.replace('=', '').replace('.', '_')
    return name


def _is_tunnel_row_type(row_type):
    """Return True if the assign_sequence row type is a tunnel object
    that uses BBox-discovered proxy variables, not SOURCE_BOUNDS polygon_no.

    Tunnel row types: tunnel, tunnel_opening, tunnel_plate, tunnel_volume"""
    return row_type.startswith('tunnel')



# ── SOURCE_BOUNDS ────────────────────────────────────────────────────

def _find_spoly_header_row(sheet, scan_from=100, scan_to=300):
    """Dynamically find the soil polygon header row by scanning column H
    for 'spoly_type'. Returns the data start row (header_row + 1),
    or None if not found."""
    for r in range(scan_from, scan_to + 1):
        val = sheet.range(r, 8).value  # column H
        if val is not None and str(val).strip().lower() == 'spoly_type':
            return r + 1
    return None


def _parse_polygon_no(value):
    """Safely extract polygon number from a cell value.
    Accepts numeric values, 'Polygon_N', or plain 'N'.
    Returns int or None."""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        num = float(text)
        return int(num) if num == int(num) else None
    except (TypeError, ValueError):
        match = re.search(r'(\d+)', text)
        return int(match.group(1)) if match else None


def _build_source_bounds_pandas(wb):
    """Build SOURCE_BOUNDS from str_2D using pandas.
    Returns {polygon_no: (xmin, xmax, ymin, ymax)}."""
    sheet = wb.sheets['str_2D']

    # Dynamically find soil polygon data start row
    data_start = _find_spoly_header_row(sheet)
    if data_start is None:
        print("sc6 WARNING: Could not find 'spoly_type' header in str_2D column H")
        return {}

    last_row = sheet.range(data_start, 2).end('down').row
    if last_row < data_start:
        return {}

    data = sheet.range(f'B{data_start}:I{last_row}').options(ndim=2).value
    if not data:
        return {}

    columns = ['spoly_name', 'cad_layer', 'x1', 'y1', 'x2', 'y2',
               'spoly_type', 'polygon_no']
    df = pd.DataFrame(data, columns=columns)

    # Clean: drop rows with no polygon_no, safely parse mixed values
    df = df.dropna(subset=['polygon_no'])
    df['polygon_no'] = df['polygon_no'].apply(_parse_polygon_no)
    df = df.dropna(subset=['polygon_no'])
    df['polygon_no'] = df['polygon_no'].astype(int)

    # Convert coordinates to float
    for col in ['x1', 'y1', 'x2', 'y2']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna(subset=['x1', 'y1', 'x2', 'y2'])

    # Aggregate bounds and types per polygon
    bounds = {}
    types = {}
    for pno, grp in df.groupby('polygon_no'):
        xmin = min(grp['x1'].min(), grp['x2'].min())
        xmax = max(grp['x1'].max(), grp['x2'].max())
        ymin = min(grp['y1'].min(), grp['y2'].min())
        ymax = max(grp['y1'].max(), grp['y2'].max())
        bounds[int(pno)] = (float(xmin), float(xmax),
                            float(ymin), float(ymax))
        _t = str(grp['spoly_type'].iloc[0]).strip().lower()
        types[int(pno)] = 'str_volume' if _t == 'none' else _t
    return bounds, types

# ── Tunnel data reader ──────────────────────────────────────────────

def _find_tunnel_header_row(sheet, scan_from=57, scan_to=77):
    """Find the tunnel section header row by scanning column C for cad_layer."""
    for r in range(scan_from, scan_to + 1):
        val = sheet.range(r, 3).value
        if val is not None and str(val).strip().lower() == 'cad_layer':
            return r
    return None


def _read_tunnel_data(wb):
    """Read tunnel rows from str_2D and deduplicate by CAD layer."""
    sheet = wb.sheets['str_2D']
    hdr = _find_tunnel_header_row(sheet) or 66
    col_map = {}
    for c in range(1, 20):
        value = sheet.range(hdr, c).value
        if value is not None:
            col_map[str(value).strip().lower()] = c

    def col(name):
        return col_map.get(name.lower())

    required = [col('cad_layer'), col('x1'), col('y1'), col('y2'),
                col('liner_model'), col('liner_position')]
    if not all(required):
        print('sc6: Could not find all tunnel columns in str_2D')
        return []

    cad_col, x_col, y_col, r_col, model_col, pos_col = required
    tunnels = {}
    for row in range(hdr + 1, hdr + 21):
        cad = sheet.range(row, cad_col).value
        if cad is None or not str(cad).strip():
            continue
        key = str(cad).strip().lower()
        x = sheet.range(row, x_col).value
        y = sheet.range(row, y_col).value
        radius = sheet.range(row, r_col).value
        if x is None or y is None or radius is None:
            continue
        item = tunnels.setdefault(key, {
            'x': float(x), 'y': float(y), 'r': float(radius),
            'has_str_volume': False,
        })
        model = str(sheet.range(row, model_col).value or '').lower()
        if 'str_volume' in model:
            item['has_str_volume'] = True
    return sorted(tunnels.values(), key=lambda item: item['x'])



def _build_source_bounds_xlwings(wb):
    """Fallback: build SOURCE_BOUNDS via xlwings cell reads.

    Optimized: single block read for all polygon data.
    Returns (bounds, types) tuple."""
    sheet = wb.sheets['str_2D']
    polys = defaultdict(lambda: {
        'xmin': float('inf'), 'xmax': float('-inf'),
        'ymin': float('inf'), 'ymax': float('-inf'),
        'type': '',
    })

    # Dynamically find soil polygon data start row
    data_start = _find_spoly_header_row(sheet)
    if data_start is None:
        print("sc6 WARNING: Could not find 'spoly_type' header in str_2D column H")
        return {}

    # Block read: columns B-J for up to 200 rows (one COM call)
    block = sheet.range((data_start, 2), (data_start + 199, 10)).value
    if block is None:
        return {}
    if not isinstance(block, list):
        block = [block]

    for row_vals in block:
        if not isinstance(row_vals, list):
            row_vals = [row_vals] * 9
        source = row_vals[0]  # B: spoly_name (index 0)
        if source is None or str(source).strip() == '':
            continue
        raw_label = row_vals[7]  # I: polygon_no (index 7)
        plaxis_no = _parse_polygon_no(raw_label)
        coords = [row_vals[c] for c in (2, 3, 4, 5)]  # D,E,F,G = indices 2,3,4,5

        if plaxis_no is not None and all(v is not None for v in coords):
            x1, y1, x2, y2 = map(float, coords)
            rec = polys[plaxis_no]
            rec['xmin'] = min(rec['xmin'], x1, x2)
            rec['xmax'] = max(rec['xmax'], x1, x2)
            rec['ymin'] = min(rec['ymin'], y1, y2)
            rec['ymax'] = max(rec['ymax'], y1, y2)
            rec['type'] = str(row_vals[6] or '').strip().lower()  # H: spoly_type
            if rec['type'] == 'none':
                # Explicit no-interface edge: same typing as 'str_volume'
                # so excavation-stack/dewatering membership never changes.
                rec['type'] = 'str_volume'

    bounds = {pno: (r['xmin'], r['xmax'], r['ymin'], r['ymax'])
              for pno, r in polys.items()}
    types = {pno: r['type'] for pno, r in polys.items()}
    return bounds, types


# ── Excavation stacks (X-range overlap, union-find) ─────────────────

def _identify_excavation_stacks(source_bounds):
    """Identify excavation stacks from SOURCE_BOUNDS using X-range overlap.

    Two polygons belong to the same stack if their X-ranges intersect.
    This handles zigzag profiles where Xmid values differ but X-ranges
    overlap (e.g. all polygons share left boundary, different right bounds).

    Union-find groups overlapping polygons. Each group is sorted top-to-bottom
    by Y-centroid. Only groups with 2+ polygons are returned.

    Returns list of lists: [[21,22,6,8,...], [4,5,7,...], ...]
    """
    if not source_bounds:
        return []

    pnos = sorted(source_bounds.keys())

    # Union-find
    parent = {p: p for p in pnos}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    # Merge polygons with overlapping X-ranges
    for i, pno_a in enumerate(pnos):
        xa_min, xa_max = source_bounds[pno_a][0], source_bounds[pno_a][1]
        for pno_b in pnos[i + 1:]:
            xb_min, xb_max = source_bounds[pno_b][0], source_bounds[pno_b][1]
            overlap = max(0.0, min(xa_max, xb_max) - max(xa_min, xb_min))
            if overlap > 1e-6:
                union(pno_a, pno_b)

    # Group by root, sort each group top-to-bottom by Y-centroid
    groups = defaultdict(list)
    for pno in pnos:
        groups[find(pno)].append(pno)

    stacks = []
    for members in groups.values():
        if len(members) >= 2:
            members.sort(
                key=lambda p: (source_bounds[p][2] + source_bounds[p][3]) / 2.0,
                reverse=True)  # top-to-bottom (ymax descending)
            stacks.append(members)

    return stacks


# ── Horizontal level neighbors (additional dewatering propagation) ─────

def _identify_horizontal_level_neighbors(target_no, source_bounds,
                                         source_types=None,
                                         tolerance=1e-6):
    """Return polygons in the same connected horizontal excavation level.

    This is an additional pass; it does not replace the vertical stack logic.
    A polygon is connected when its Y interval overlaps the target level and
    its X interval touches/overlaps another member. Connectivity is transitive
    so split left/middle/right geometry is treated as one level.

    When source_types is provided, only excavation-type polygons (cut,
    soil_replacement, str_volume) are considered. Fill polygons are excluded
    to prevent Dry from propagating into ground-surface fills.
    """
    if target_no not in source_bounds:
        return []

    EXCAVATION_TYPES = {'cut', 'soil_replacement', 'str_volume'}

    def _touches_x(a, b):
        return (min(a[1], b[1]) - max(a[0], b[0])) >= -tolerance

    def _overlaps_y(a, b):
        return (min(a[3], b[3]) - max(a[2], b[2])) >= -tolerance

    connected = {target_no}
    changed = True
    while changed:
        changed = False
        members = list(connected)
        for candidate, candidate_bounds in source_bounds.items():
            if candidate in connected:
                continue
            # Filter: only excavation-type polygons
            if source_types is not None:
                ctype = source_types.get(candidate, '')
                if ctype not in EXCAVATION_TYPES:
                    continue
            if any(_touches_x(source_bounds[member], candidate_bounds)
                   and _overlaps_y(source_bounds[member], candidate_bounds)
                   for member in members):
                connected.add(candidate)
                changed = True

    return sorted(connected)


# ── Water condition calls ────────────────────────────────────────────


def _emit_set_water_helper():
    """Generate runtime helper that maps source polygon numbers to PLAXIS polygons.

    The old SC5/SC6 implementation deliberately used the PLAXIS polygon label
    created from str_2D column I.  This is more reliable than trying to rebuild
    excavation corridors from all runtime SoilPolygons.  The BBox fallback is
    retained for surface/intermediate polygons whose generated name has no
    ``_Polygon_N_`` token.
    """
    return [
        '# Source geometry bounds, calculated from str_2D.',
        '# Format: polygon_no: (xmin, xmax, ymin, ymax)',
        'SOURCE_BOUNDS = {}',
        '',
        'def set_water_for_source_polygon(polygon_no, phase, condition):',
        '    """Set a water condition on the real PLAXIS polygon for polygon_no."""',
        '    tolerance = 1e-6',
        '    tag = f"_Polygon_{polygon_no}_"',
        '    xmin, xmax, ymin, ymax = SOURCE_BOUNDS[polygon_no]',
        '    matched = []',
        '    fallback_matched = []',
        '',
        '    # Column I is used by SC5 when naming the generated PLAXIS polygon.',
        '    for poly in g_i.Polygons:',
        '        real_name = poly.Name.value',
        '        if tag in real_name:',
        '            poly.Soil.WaterConditions.Conditions[phase] = condition',
        '            matched.append(real_name)',
        '',
        '    # Surface/intermediate polygons can have names such as BoreholePolygon_N_M.',
        '    if not matched:',
        '        for poly in g_i.Polygons:',
        '            bb = poly.BoundingBox',
        '            try:',
        '                pxmin, pxmax = float(bb.xMin.value), float(bb.xMax.value)',
        '                pymin, pymax = float(bb.yMin.value), float(bb.yMax.value)',
        '            except AttributeError:',
        '                pxmin, pxmax = float(bb.xMin), float(bb.xMax)',
        '                pymin, pymax = float(bb.yMin), float(bb.yMax)',
        '            if (pxmin >= xmin - tolerance and pxmax <= xmax + tolerance',
        '                    and pymin >= ymin - tolerance and pymax <= ymax + tolerance):',
        '                poly.Soil.WaterConditions.Conditions[phase] = condition',
        '                fallback_matched.append(poly.Name.value)',
        '',
        '    names = matched or fallback_matched',
        '    if names:',
        '        mode = "name" if matched else "bbox fallback"',
        '        print(f"Polygon_{polygon_no}: {condition} ({mode}, {len(names)} slice(s))")',
        '        for name in names:',
        '            print(f"  {name}")',
        '    else:',
        '        print(f"Polygon_{polygon_no}: NOT FOUND '
        'for x=[{xmin},{xmax}], y=[{ymin},{ymax}]")',
        '',
    ]


def _emit_source_bounds_dict(source_bounds):
    """Emit the SOURCE_BOUNDS dictionary assignment."""
    lines = ['SOURCE_BOUNDS = {']
    for pno in sorted(source_bounds.keys()):
        xmin, xmax, ymin, ymax = source_bounds[pno]
        lines.append(f'    {pno}: ({xmin}, {xmax}, {ymin}, {ymax}),')
    lines.append('}')
    return lines


def _emit_water_calls(target_no, phase_var, stack):
    """Emit set_water_for_source_polygon calls for one stack.

    Uses Y-centroid comparison: all polygons at or above the target's
    Y-centroid get Dry, those below get Interpolate.  This correctly
    handles side-by-side polygons at the same excavation level that
    share the same Y-centroid but have different positions in the stack.

    Returns list of code lines and count of polygons processed.
    """
    lines = []
    target_yc = (source_bounds_global[target_no][2]
                 + source_bounds_global[target_no][3]) / 2.0

    for no in stack:
        if no not in source_bounds_global:
            lines.append(
                f'# WARNING: Polygon_{no} not in SOURCE_BOUNDS, skipped')
            continue
        no_yc = (source_bounds_global[no][2]
                 + source_bounds_global[no][3]) / 2.0
        condition = 'Dry' if no_yc >= target_yc - 1e-6 else 'Interpolate'
        lines.append(
            f'set_water_for_source_polygon({no}, {phase_var}, "{condition}")')

    return lines, len(stack)


def _emit_horizontal_dry_extension(target_no, phase_var, stack):
    """No-op: horizontal extension removed.

    The Y-centroid comparison in _emit_water_calls already handles
    side-by-side polygons at the same excavation level within a stack.
    The broad horizontal neighbor search was causing Dry to propagate
    into fill polygons outside the excavation walls.
    """
    return [], 0


# Module-level reference set by main() for _emit_water_calls
source_bounds_global = {}
source_types_global = {}


# ── Notebook cell helpers ────────────────────────────────────────────

def _make_cell(source, cell_type='code'):
    if isinstance(source, list):
        source = '\n'.join(source)
    return {'cell_type': cell_type, 'metadata': {},
            'source': [source], 'outputs': [], 'execution_count': None}


def _write_diagnostic_sheet(wb, phase_cols, scan_log, phase_actions, warnings):
    """Write SC6 diagnostic data into the existing ``DIAGNOSTIC`` worksheet.

    Called from ``_main_io`` so the user can inspect it after running from
    an Excel button.  The sheet must already exist (created manually or by
    VBA).  All writes are batched into one block write per section.
    """
    import time as _time

    try:
        diag = wb.sheets('DIAGNOSTIC')
    except Exception:
        return  # sheet not found — skip silently

    diag.clear_contents()

    # ── Section 1: header ──
    ts = _time.strftime('%Y-%m-%d %H:%M:%S')
    diag.range((1, 1), (1, 2)).value = [['Workbook', Path(wb.fullname).name]]
    diag.range((2, 1), (2, 2)).value = [['Module', __file__]]
    diag.range((3, 1), (3, 2)).value = [['Timestamp', ts]]

    # ── Section 2: phase discovery ──
    diag.range((5, 1), (5, 1)).value = [['PHASE DISCOVERY']]
    r = 6
    diag.range((r, 1), (r, 4)).value = [['col', 'name', 'calc_type', 'default_name']]
    r += 1
    for col, pname, ct, ti, um, ident, is_def in phase_cols:
        diag.range((r, 1), (r, 4)).value = [[col, pname, ct, is_def]]
        r += 1

    # ── Section 3: scan log (every parsed cell) ──
    r += 1
    diag.range((r, 1), (r, 8)).value = [['SCAN LOG']]
    r += 1
    diag.range((r, 1), (r, 8)).value = [
        ['row', 'col', 'label', 'pno', 'raw_val',
         'action', 'prestress', 'supported']]
    r += 1
    if scan_log:
        # batch write in chunks of 200 to avoid COM overhead
        chunk = 200
        for i in range(0, len(scan_log), chunk):
            batch = scan_log[i:i + chunk]
            diag.range((r, 1), (r + len(batch) - 1, 8)).value = batch
            r += len(batch)

    # ── Section 4: phase_actions summary ──
    r += 1
    diag.range((r, 1), (r, 4)).value = [['PHASE ACTIONS (for codegen)']]
    r += 1
    diag.range((r, 1), (r, 5)).value = [['phase', 'action', 'label', 'prestress', 'row_type']]
    r += 1
    pa_rows = []
    for pname, acts in phase_actions.items():
        for action, label, prestress, row_type, row_name in acts:
            pa_rows.append([pname, action, label,
                            prestress if prestress is not None else '',
                            row_type])
    if pa_rows:
        chunk = 200
        for i in range(0, len(pa_rows), chunk):
            batch = pa_rows[i:i + chunk]
            diag.range((r, 1), (r + len(batch) - 1, 5)).value = batch
            r += len(batch)

    # ── Section 5: warnings ──
    r += 1
    diag.range((r, 1), (r, 1)).value = [['WARNINGS']]
    r += 1
    if warnings:
        for w in warnings:
            diag.range((r, 1), (r, 1)).value = [[w]]
            r += 1
    else:
        diag.range((r, 1), (r, 1)).value = [['(none)']]

    # ── Section 6: supported-prestress summary ──
    r += 1
    diag.range((r, 1), (r, 4)).value = [['PRESTRESS SUPPORTED ROWS']]
    r += 1
    diag.range((r, 1), (r, 5)).value = [
        ['phase', 'label', 'prestress', 'row_type', 'supported']]
    r += 1
    ps_rows = []
    for pname, acts in phase_actions.items():
        for action, label, prestress, row_type, row_name in acts:
            if prestress is not None:
                ps_rows.append([
                    pname, label, prestress, row_type,
                    _is_prestress_supported_label(label)])
    if ps_rows:
        diag.range((r, 1), (r + len(ps_rows) - 1, 5)).value = ps_rows
        r += len(ps_rows)
    else:
        diag.range((r, 1), (r, 1)).value = [['(none)']]

    print(f"sc6: diagnostic sheet written ({r - 1} rows)")


# ── Main ─────────────────────────────────────────────────────────────

def main():
    global source_bounds_global, source_types_global

    try:
        wb = xw.Book.caller()
    except Exception:
        _cwd = Path.cwd()
        _xl = list(_cwd.glob('*.xlsm')) or list(_cwd.parent.glob('*.xlsm'))
        if not _xl:
            raise FileNotFoundError(f"No .xlsm file found in {_cwd}")
        wb = xw.Book(str(_xl[0]))

    # ── Disable screen updating during I/O ──
    app = wb.app
    previous_screen_updating = app.screen_updating
    app.screen_updating = False
    try:
        _main_io(wb)
    finally:
        app.screen_updating = previous_screen_updating


def _main_io(wb):
    """SC6 entry — progress wrapper (mirrors SC8 pattern).

    Sets the WORKING banner on assign_sequence!B5; crash → FAILED,
    success path sets the done text. Never raises.
    """
    _set_progress(wb, f"sc6: {_spin()} WORKING — reading geometry…")
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
            _set_progress(
                wb, "sc6: FAILED — see console for detail")


def _main_io_inner(wb):
    """Core logic for SC6 — separated from main() for screen_updating wrapper."""
    global source_bounds_global, source_types_global

    # ── 0. Write diagnostic sheet FIRST, before any crash can occur ──
    try:
        _write_diagnostic_sheet(wb, [], [], {}, [])
    except Exception:
        pass  # best-effort: don't let diagnostic failure block SC6

    # ── 1. Build SOURCE_BOUNDS from str_2D ──
    _set_progress(wb, f"sc6: {_spin()} WORKING — reading str_2D {_bar(0.05)}")
    if HAS_PANDAS:
        source_bounds = _build_source_bounds_pandas(wb)
    else:
        source_bounds = _build_source_bounds_xlwings(wb)
    # Both builders return (bounds, types) tuple
    if isinstance(source_bounds, tuple):
        source_bounds, source_types = source_bounds
    else:
        source_types = {}
    source_bounds_global = source_bounds
    source_types_global = source_types

    # ── 2. Identify excavation stacks ──
    _set_progress(wb, f"sc6: {_spin()} WORKING — building stacks {_bar(0.15)}")
    stacks = _identify_excavation_stacks(source_bounds)
    _set_progress(
        wb,
        f"sc6: {_spin()} WORKING — {len(stacks)} stack(s) {_bar(0.25)}")

    # ── 3. Read assign_sequence ──
    sheet = wb.sheets['assign_sequence']

    # ── 3a. Block read: phase header rows 1-6, columns L through BD (12-56) ──
    # This replaces 50+ individual COM calls with a single block read.
    phase_block = sheet.range((1, 12), (6, 61)).value
    if phase_block is None:
        phase_block = []

    # Discover phases from the block data
    # phase_block[i] corresponds to row i+1, col j+12 = phase_block[i][j]
    phase_cols = []
    if phase_block and isinstance(phase_block[0], list):
        n_cols = len(phase_block[0])
        _auto_phase_idx = 0  # counter for auto-generated phase names
        for col_idx in range(n_cols):
            raw_calc = phase_block[0][col_idx] if phase_block[0][col_idx] is not None else None
            if raw_calc is None or str(raw_calc).strip() == '':
                continue
            # Row 6 is index 5 in block — user-supplied phase name (Identify).
            # When blank, PLAXIS uses its own default name ("Phase N").
            row6_vals = phase_block[5] if len(phase_block) > 5 else [None] * n_cols
            raw_pname = row6_vals[col_idx] if col_idx < len(row6_vals) else None
            pname = str(raw_pname).strip() if raw_pname is not None else ''
            _auto_phase_idx += 1
            if not pname:
                # Blank → use auto-generated name; code gen will skip .Identify
                pname = f'Phase_{_auto_phase_idx}'
                _pname_is_default = True
            else:
                _pname_is_default = False
            ct = normalize_calc_type(raw_calc)
            # Row 2 = index 1, Row 3 = index 2, Row 4 = index 3
            ti = phase_block[1][col_idx] if len(phase_block) > 1 and col_idx < len(phase_block[1]) else None
            um = _normalize_bool(phase_block[2][col_idx] if len(phase_block) > 2 and col_idx < len(phase_block[2]) else None)
            raw_ident = phase_block[3][col_idx] if len(phase_block) > 3 and col_idx < len(phase_block[3]) else None
            ident = str(raw_ident).strip() if raw_ident is not None else ''
            actual_col = 12 + col_idx  # actual Excel column number
            phase_cols.append((actual_col, pname, ct,
                               float(ti) if ti is not None else None, um, ident,
                               _pname_is_default))

    if not phase_cols:
        print("sc6: No phases found in assign_sequence")
        _set_progress(wb, "sc6: done — no phases found")
        return
    _set_progress(
        wb,
        f"sc6: {_spin()} WORKING — {len(phase_cols)} phase(s) {_bar(0.35)}")

    # ── 3b. Block read: marker rows (rows 7 through 213, all columns) ──
    # Scan column D to find last row with data, then block read.
    # We read columns B..BD (2-56) to cover all phase columns + row metadata.
    max_scan_row = 213
    marker_block = sheet.range((7, 2), (max_scan_row, 61)).value
    if marker_block is None:
        marker_block = []
    if isinstance(marker_block, list) and marker_block and not isinstance(marker_block[0], list):
        marker_block = [marker_block]

    # Scan markers from block data in-memory (no more COM calls)
    warnings = []
    phase_actions = {}   # pname -> [(action, label, prestress, row_type, row_name)]
    phase_water = {}     # pname -> [(pno, water)]  (soil polygons only)
    phase_tunnel_water = {}  # pname -> [(label, water, pressure, row_type, row_name)]
    phase_global_water = {}  # pname -> [water-level proxy labels]
    phase_ss = {}  # pname -> True when SS marker is present
    phase_direct_dry = {}  # pname -> [(pno)] — direct Dry, no stack propagation
    phase_direct_interp = {}  # pname -> [(pno)] — direct Interpolate, no stack propagation
    has_water = False
    scan_log = []

    for row_offset, row_vals in enumerate(marker_block):
        if not isinstance(row_vals, list):
            row_vals = [row_vals] * 60
        actual_row = 7 + row_offset

        # Column D (index 2 in block) = label
        label_raw = row_vals[2] if len(row_vals) > 2 else None
        if label_raw is None or str(label_raw).strip() == '':
            # v0.7.7: blank column D is a section separator, not end-of-table.
            # Only break when the entire row is empty across all columns,
            # otherwise skip this row and continue scanning.
            if all(v is None or (isinstance(v, str) and v.strip() == '')
                   for v in row_vals):
                break
            continue
        label = str(label_raw).strip()
        # Column C (index 1) = row_type, Column B (index 0) = row_name
        row_type = str(row_vals[1] if len(row_vals) > 1 else None or '').strip()
        # Column K is the only source for staged material assignment.
        # Normalize spreadsheet blanks and placeholder text defensively:
        # blank K must never become setmaterial(..., None).
        row_name = str(row_vals[0] if len(row_vals) > 0 else None or '').strip()
        if row_name.lower() in ('none', 'nan', 'null'):
            row_name = ''
        pno = _extract_polygon_number(label)

        for col, pname, _, _, _, _, _ in phase_cols:
            # block index: col is 1-based Excel column; block starts at col 2, so index = col - 2
            block_idx = col - 2
            if block_idx >= len(row_vals):
                continue
            val = row_vals[block_idx]
            if val is None:
                continue
            action, water, prestress, user_pressure, steady_state, p_warn = _parse_marker(val)
            scan_log.append([
                actual_row, col, label, pno, str(val), action or '',
                prestress if prestress is not None else '',
                _is_prestress_supported_label(label)])
            print(f"  scan row={actual_row} col={col} label={label!r} "
                  f"pno={pno} val={val!r} -> action={action} water={water} "
                  f"prestress={prestress} user_pressure={user_pressure} "
                  f"steady_state={steady_state}")
            if p_warn:
                warnings.append(p_warn)
            # global_waterlevel: A in phase → setglobalwaterlevel, not activate
            if row_type == 'global_waterlevel' and action == 'activate':
                phase_global_water.setdefault(pname, []).append(label)
            if steady_state:
                phase_ss[pname] = True
            if row_type == 'global_waterlevel' and action == 'activate':
                continue
            if action is not None:
                phase_actions.setdefault(pname, []).append((action, label, prestress, row_type, row_name))
            elif prestress is not None:
                # P without A/D — prestress only (element must already be active)
                phase_actions.setdefault(pname, []).append(('prestress', label, prestress, row_type, row_name))
            if water is not None or user_pressure is not None:
                if _is_tunnel_row_type(row_type):
                    # Tunnel water is always applied directly to the existing proxy.
                    phase_tunnel_water.setdefault(pname, []).append(
                        (label, water, user_pressure, row_type, row_name))
                else:
                    # U markers are tunnel-specific: pressure on a non-tunnel
                    # row is rejected so it never silently no-ops.
                    if user_pressure is not None:
                        warnings.append(
                            f"User pressure {user_pressure:g} kPa ignored on "
                            f"non-tunnel row '{label}' (type '{row_type}') "
                            f"- U markers are tunnel-only")
                    if water is not None and pno is not None:
                        if water == 'Dewatering':
                            # Stack-resolved dewatering (deepest target rule)
                            phase_water.setdefault(pname, []).append((pno, water))
                            has_water = True
                        elif water == 'Dry':
                            # Direct dry — no stack propagation
                            phase_direct_dry.setdefault(pname, []).append(pno)
                            has_water = True
                        elif water == 'Interp':
                            # Direct interpolate — no stack propagation
                            phase_direct_interp.setdefault(pname, []).append(pno)
                            has_water = True

    _set_progress(wb, f"sc6: {_spin()} WORKING — scanning markers {_bar(0.50)}")

    # ── Diagnostic: write to sc6_diagnostic worksheet ──
    _write_diagnostic_sheet(wb, phase_cols, scan_log, phase_actions, warnings)

    # ── 3c. Paint output-table activation windows (first A → last
    # phase, structural rows only) — reuses the block reads above, no
    # new COM round-trips for data.
    _set_progress(wb, f"sc6: {_spin()} WORKING — painting output windows {_bar(0.55)}")
    _paint_output_activation_window(wb, phase_cols, marker_block)

    # ── 4. Resolve water per phase per stack ──
    #    For each phase: find deepest Dry target per stack.
    #    Multiple Dry targets in same stack = warn.
    phase_stack_water = {}  # pname -> [(target_no, stack)]

    for pname, markers in phase_water.items():
        stack_targets = {}  # stack_index -> [(target_no, stack)]

        for pno, cond in markers:
            if cond != 'Dewatering':
                continue
            # Find which stack this polygon belongs to
            for si, stk in enumerate(stacks):
                if pno in stk:
                    stack_targets.setdefault(si, []).append((pno, stk))
                    break
            else:
                warnings.append(
                    f"Polygon_{pno} not in any excavation stack "
                    f"(phase {pname})")

        for si, targets in stack_targets.items():
            stack = targets[0][1]
            if len(targets) > 1:
                # Multiple Dry targets in same stack — use deepest
                target_nos = [t[0] for t in targets]
                deepest = max(target_nos,
                              key=lambda n: stack.index(n))
                warnings.append(
                    f"Multiple Dry targets in same stack "
                    f"(phase {pname}): {target_nos}. "
                    f"Using deepest Polygon_{deepest}")
            else:
                deepest = targets[0][0]
            phase_stack_water.setdefault(pname, []).append(
                (deepest, stack))

    _set_progress(wb, f"sc6: {_spin()} WORKING — resolving water {_bar(0.60)}")

    # ── 5. Load / clean existing notebook ──
    tunnel_id_cells = []  # v0.6.16: captured from SC5
    nb_path = notebook_path_for_workbook(wb)
    if nb_path.exists():
        with open(nb_path, 'r', encoding='utf-8') as f:
            notebook = json.load(f)
        cells = notebook.get('cells', [])
        new_cells = []
        tunnel_id_cells = []  # v0.6.16: capture SC5's tunnel identification
        skip = False
        for cell in cells:
            src = ''.join(cell.get('source', []))
            # v0.6.16: Capture SC5 tunnel identification cells FIRST,
            # before skip logic. These cells contain g_i.* calls but must
            # be preserved and relocated after gotostages, not deleted.
            if TUNNEL_ID_TAG in src:
                tunnel_id_cells.append(cell)
                continue
            if SECTION_TAG in src:
                skip = True
                continue
            if skip:
                # Skip the staged construction section (all g_i.* cells
                # and gotostages)
                if ('g_i.' in src or 'gotostages' in src
                        or 'EXCAVATION' in src or 'SOURCE_BOUNDS' in src
                        or 'set_water_for_source_polygon' in src):
                    continue
                # Also skip markdown cells that look like phase headers
                if cell.get('cell_type') == 'markdown':
                    continue
                skip = False
            new_cells.append(cell)
        notebook['cells'] = new_cells
    else:
        raise FileNotFoundError(
            f"sc6: Notebook '{nb_path.name}' not found for "
            f"workbook '{Path(wb.fullname).name}'. Run SC3 first."
        )

    # ── 6. Generate staged construction ──
    out = [_make_cell(SECTION_TAG, 'markdown')]
    out.append(_make_cell('g_i.gotostages()'))

    # ── 6a. Tunnel identification (BBox scan — relocated from SC5) ──
    # v0.6.16: SC5 generates the identification code during structural
    # generation. SC6 captures it from the notebook and places it right
    # after gotostages(), where SoilPolygons from generatetunnel exist.
    if tunnel_id_cells:
        out.extend(tunnel_id_cells)
        print(f"sc6: relocated {len(tunnel_id_cells)} tunnel identification cell(s) after gotostages")

    # ── 6b. Emit set_water_for_source_polygon helper + SOURCE_BOUNDS ──
    # v0.6.17: restored from v0.6.16 — name-based mapping via PLAXIS label
    # (column I) is more reliable than runtime BBox scanning.
    if has_water:
        water_helper_lines = _emit_set_water_helper()
        out.append(_make_cell(water_helper_lines))
        bounds_lines = _emit_source_bounds_dict(source_bounds)
        out.append(_make_cell(bounds_lines))
        print(f"sc6: emitted set_water_for_source_polygon helper + SOURCE_BOUNDS ({len(source_bounds)} polygons)")

    prev_var = None
    total_actions = 0
    total_water = 0

    n_phases = len(phase_cols)
    for _pi, (col, pname, ct, ti, um, ident,
              pname_is_default) in enumerate(phase_cols):
        # Codegen spans 60–95%: honest per-phase advance on the same scale.
        frac = 0.60 + 0.35 * ((_pi + 1) / max(n_phases, 1))
        _set_progress(
            wb,
            f"sc6: {_spin()} WORKING — {pname} "
            f"({_pi + 1}/{n_phases}) {_bar(frac)}")
        acts = phase_actions.get(pname, [])
        total_actions += len(acts)

        if pname == 'Initial':
            lines = [
                '# ── Initial Phase ──',
                'g_i.setcurrentphase(g_i.InitialPhase)',
                f'g_i.InitialPhase.DeformCalcType = "{ct}"',
            ]
            if ti is not None:
                lines.append(f'g_i.InitialPhase.TimeInterval = {ti}')
            if phase_ss.get(pname):
                lines.append(
                    'g_i.InitialPhase.PorePresCalcType = '
                    '"Steady state groundwater flow"')
            for act, lbl, prestress, row_type, row_name in acts:
                if lbl.startswith('tunnel_'):
                    lines.append(f'g_i.{act}({lbl}, g_i.InitialPhase)')
                else:
                    lines.append(f'g_i.{act}(g_i.{lbl}, g_i.InitialPhase)')
                # Structural volume liners, soil replacements, and
                # volume_profile polygons need explicit material
                # assignment when activated.
                if act == 'activate' and row_type in ('tunnel_volume', 'soil_replacement', 'str_volume') and row_name:
                    if row_type == 'tunnel_volume' and lbl.startswith('tunnel_'):
                        lines.append(f'{lbl}.setmaterial(g_i.InitialPhase, {row_name})')
                    else:
                        lines.append(f'g_i.{lbl}.setmaterial(g_i.InitialPhase, {row_name})')
                # Staged prestress: N2N anchors and FEA only.
                lines.extend(_emit_prestress_lines(
                    lbl, 'g_i.InitialPhase', prestress))

            # v0.6.18: Soil polygon water conditions in Initial phase.
            sw = phase_stack_water.get(pname, [])
            if sw:
                lines.append('')
                lines.append('# ── Soil polygon water conditions ──')
                for target_no, stack in sw:
                    lines.append(
                        f'# Water conditions: target Polygon_{target_no} '
                        f'in stack')
                    water_lines, count = _emit_water_calls(
                        target_no, 'g_i.InitialPhase', stack)
                    lines.extend(water_lines)
                    total_water += count
                    extra_lines, extra_count = _emit_horizontal_dry_extension(
                        target_no, 'g_i.InitialPhase', stack)
                    lines.extend(extra_lines)
                    total_water += extra_count

            # v0.6.18: Tunnel water conditions in Initial phase.
            tw = phase_tunnel_water.get(pname, [])
            if tw:
                lines.append('')
                lines.append('# ── Tunnel water conditions (direct) ──')
                for tun_label, tun_water, tun_pressure, tun_type, tun_name in tw:
                    if tun_water is not None:
                        lines.append(
                            f'{tun_label}.Soil.WaterConditions.Conditions'
                            f'[g_i.InitialPhase] = "{tun_water}"')
                    if tun_pressure is not None:
                        lines.extend(_emit_tunnel_user_pressure(
                            tun_label, 'g_i.InitialPhase', tun_pressure))
                total_water += len(tw)

            # v0.7.6: Direct Dry + Interp in InitialPhase.
            dd = phase_direct_dry.get(pname, [])
            di = phase_direct_interp.get(pname, [])
            if dd or di:
                lines.append('')
                lines.append('# ── Direct water conditions ──')
                for pno in dd:
                    lines.append(
                        f'set_water_for_source_polygon({pno}, g_i.InitialPhase, "Dry")')
                for pno in di:
                    lines.append(
                        f'set_water_for_source_polygon({pno}, g_i.InitialPhase, "Interpolate")')
                total_water += len(dd) + len(di)

            # Global water level assignments in InitialPhase.
            gw = phase_global_water.get(pname, [])
            if gw:
                lines.append('')
                lines.append('# ── Global water level ──')
                if len(gw) > 1:
                    warnings.append(
                        f'Multiple global water levels in phase {pname}: '
                        f'{gw}; last assignment is effective')
                for water_label in gw:
                    lines.append(
                        f'g_i.setglobalwaterlevel(g_i.{water_label}, '
                        f'g_i.InitialPhase)')
                total_water += len(gw)

            out.append(_make_cell(lines))
            prev_var = 'g_i.InitialPhase'
        else:
            vn = pname.replace(' ', '_')
            lines = [
                f'# ── {pname} ──',
                f'{vn} = g_i.phase({prev_var})',
                f'g_i.setcurrentphase({vn})',
            ]
            # Only emit .Identify when user provided a custom name.
            # Blank cells → auto-generated Phase_N; skip .Identify so
            # PLAXIS uses its own default "Phase N" label instead.
            if not pname_is_default:
                lines.append(f'{vn}.Identify = "{pname}"')
            if ident:
                lines.append(f'{vn}.Identification = {ident!r}')
            lines.append(f'{vn}.DeformCalcType = "{ct}"')
            if ti is not None:
                lines.append(f'{vn}.TimeInterval = {ti}')
            if phase_ss.get(pname):
                lines.append(
                    f'{vn}.PorePresCalcType = '
                    '"Steady state groundwater flow"')
            if um:
                lines.append(f'{vn}.Deform.UseUpdatedMesh = True')
                lines.append(f'{vn}.Deform.UseUpdatedWaterPressures = True')
            lines.append(f'{vn}.MaxStepsStored = 50')

            # Activation / deactivation / prestress calls
            for act, lbl, prestress, row_type, row_name in acts:
                if act == 'prestress':
                    # Standalone P marker — prestress only, no activate/deactivate.
                    lines.extend(_emit_prestress_lines(lbl, vn, prestress))
                else:
                    # v0.6.16: Proxy labels (tunnel_N_opening, tunnel_N_liners[i])
                    # are Python variables, not g_i attributes.
                    if lbl.startswith('tunnel_'):
                        lines.append(f'g_i.{act}({lbl}, {vn})')
                    else:
                        lines.append(f'g_i.{act}(g_i.{lbl}, {vn})')
                # Structural volume liners, soil replacements, and
                # volume_profile polygons need explicit material
                # assignment when activated.
                if act == 'activate' and row_type in ('tunnel_volume', 'soil_replacement', 'str_volume') and row_name:
                    if row_type == 'tunnel_volume' and lbl.startswith('tunnel_'):
                        lines.append(f'{lbl}.setmaterial({vn}, {row_name})')
                    else:
                        lines.append(f'g_i.{lbl}.setmaterial({vn}, {row_name})')
                # v0.7.7: Combined A,P or D,P markers — emit prestress
                # alongside activation/deactivation for N2N anchors and FEA.
                if act != 'prestress':
                    lines.extend(_emit_prestress_lines(lbl, vn, prestress))

            # v0.6.17: Soil polygon water conditions via set_water_for_source_polygon.
            # Each source polygon_no maps to its PLAXIS polygon(s) by name tag
            # (e.g. _Polygon_6_) or BBox fallback, then Dry/Interpolate is
            # determined by position in the source stack.
            sw = phase_stack_water.get(pname, [])
            if sw:
                lines.append('')
                lines.append('# ── Soil polygon water conditions ──')
                for target_no, stack in sw:
                    lines.append(
                        f'# Water conditions: target Polygon_{target_no} '
                        f'in stack')
                    water_lines, count = _emit_water_calls(
                        target_no, vn, stack)
                    lines.extend(water_lines)
                    total_water += count
                    extra_lines, extra_count = _emit_horizontal_dry_extension(
                        target_no, vn, stack)
                    lines.extend(extra_lines)
                    total_water += extra_count

            # v0.6.17: Tunnel water conditions via direct proxy variable calls.
            # These bypass SOURCE_BOUNDS entirely — the proxy variables
            # (tunnel_N_opening, tunnel_N_liners[i]) are BBox-discovered
            # PLAXIS objects, not soil polygons with polygon_no.
            tw = phase_tunnel_water.get(pname, [])
            if tw:
                lines.append('')
                lines.append('# ── Tunnel water conditions (direct) ──')
                for tun_label, tun_water, tun_pressure, tun_type, tun_name in tw:
                    if tun_water is not None:
                        lines.append(
                            f'{tun_label}.Soil.WaterConditions.Conditions'
                            f'[{vn}] = "{tun_water}"')
                    if tun_pressure is not None:
                        lines.extend(_emit_tunnel_user_pressure(
                            tun_label, vn, tun_pressure))
                total_water += len(tw)

            # v0.7.6: Direct Dry + Interp — no stack propagation.
            dd = phase_direct_dry.get(pname, [])
            di = phase_direct_interp.get(pname, [])
            if dd or di:
                lines.append('')
                lines.append('# ── Direct water conditions ──')
                for pno in dd:
                    lines.append(
                        f'set_water_for_source_polygon({pno}, {vn}, "Dry")')
                for pno in di:
                    lines.append(
                        f'set_water_for_source_polygon({pno}, {vn}, "Interpolate")')
                total_water += len(dd) + len(di)

            # Global water level assignments: A selects the phase-wide level.
            gw = phase_global_water.get(pname, [])
            if gw:
                lines.append('')
                lines.append('# ── Global water level ──')
                if len(gw) > 1:
                    warnings.append(
                        f'Multiple global water levels in phase {pname}: '
                        f'{gw}; last assignment is effective')
                for water_label in gw:
                    lines.append(f'g_i.setglobalwaterlevel(g_i.{water_label}, {vn})')
                total_water += len(gw)

            out.append(_make_cell(lines))
            prev_var = vn

    out.append(_make_cell(
        f'# Staged construction: {len(phase_cols)} phases, '
        f'{total_actions} actions, {total_water} water calls',
        'markdown'))

    notebook['cells'].extend(out)
    _set_progress(wb, f"sc6: {_spin()} WORKING — writing notebook {_bar(0.97)}")
    with open(nb_path, 'w', encoding='utf-8') as f:
        json.dump(notebook, f, indent=1, ensure_ascii=False)

    # ── 7. Warnings → all in B4 (pairs with B5 progress) ──
    all_warnings = []
    first_um = next((pn for _, pn, _, _, um, _, _ in phase_cols if um), None)
    if first_um:
        all_warnings.append(
            f"⚠ upd_mesh=yes on {first_um} -> all subsequent phases "
            f"also have updated mesh (PLAXIS chaining)")
    if warnings:
        all_warnings.extend(f"⚠ {w}" for w in warnings)
        for w in warnings:
            print(f"  WARNING: {w}")
    if all_warnings:
        sheet.range('B4').value = ' | '.join(all_warnings)
        sheet.range('B4').api.Font.Color = 0x0000FF
    else:
        sheet.range('B4').value = None

    # ── 8. Summary ──
    print(f"sc6: Staged construction -> {nb_path.name}")
    print(f"    Phases: {len(phase_cols)}, Actions: {total_actions}")
    print(f"    SOURCE_BOUNDS: {len(source_bounds)} polygons")
    print(f"    Excavation stacks: {len(stacks)}")
    for si, stk in enumerate(stacks):
        print(f"      Stack {si}: {stk}")
    for _, pn, ct, _, _, _, _ in phase_cols:
        acts = phase_actions.get(pn, [])
        pw = phase_water.get(pn, [])
        n_a = sum(1 for a, *_ in acts if a == 'activate')
        n_d = sum(1 for a, *_ in acts if a == 'deactivate')
        n_p = sum(1 for a, _, p, *_ in acts if p is not None)
        n_w = len(pw)
        d = f"{n_a}A {n_d}D"
        if n_p:
            d += f" {n_p}P"
        if n_w:
            d += f" {n_w}W"
        print(f"    {pn} ({ct}): {d}")
    print("sc6: Done.")
    print(f"sc6: MODULE_ID = {__file__}")
    _set_progress(
        wb,
        f"sc6: done — {len(phase_cols)} phases, "
        f"{total_actions} actions {_bar(1.0)}")


if __name__ == '__main__':
    main()
