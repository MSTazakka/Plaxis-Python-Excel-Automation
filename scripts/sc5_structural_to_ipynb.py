"""
sc5_structural_to_ipynb.py — Str_2D Tab → .ipynb (Button 2) [v0.7.9]
============================================================
Reads the 'str_2D' Excel tab via xlwings and appends PLAXIS 2D
Python API cells for plate, N2N anchor, embedded beam, strut,
pile, and line load elements to the .ipynb file created by sc3.

Called from Excel via:
  RunPython "import scripts.sc5_structural_to_ipynb as l; l.main()"
"""

from pathlib import Path
from math import pi, isnan, sqrt
import json
import re
import xlwings as xw

try:
    from .notebook_paths import notebook_path_for_workbook
except ImportError:  # pragma: no cover - supports direct script execution
    from notebook_paths import notebook_path_for_workbook


# ── PLAXIS API section markers ──
SECTION_TAG = '# ── Structural elements (from sc5) ──'
TUNNEL_TAG = '# ── Tunnel elements (from sc5) ──'
TUNNEL_ID_TAG = '# ── Tunnel identification (from sc5) ──'
SOIL_POLY_ID_TAG = '# ── Soil polygon identification (from sc5) ──'
SOIL_MAT_TAG = '# ── Soil materials (from sc5) ──'
DRAIN_TAG = '# ── PVD Drains ──'
LOAD_TAG = '# ── Line Loads ──'
POINT_LOAD_TAG = '# ── Point Loads ──'
WATERLEVEL_TAG = '# ── Water Levels (from sc5) ──'
GWFLOW_BC_TAG = '# ── Groundwater Flow Boundaries (from sc5) ──'
MERGE_TAG = '# ── Merge equivalents (from sc5) ──'


# ── Progress cell (SC5, mirrors SC6/SC9 pattern) ──
# str_2D!D4 is the SC5 banner cell (own sheet — no clash with the shared
# assign_sequence!B5 that SC6 hands to SC9 in the FULL chain). Written
# live during structural generation. main() disables screen_updating for
# speed, so each write briefly re-enables it so the cell repaints, then
# restores the previous state. Never raises — progress must not break runs.
PROGRESS_SHEET = 'str_2D'
PROGRESS_CELL = 'D4'
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
    """Write *msg* to str_2D!D4. False/None clears. Never raises."""
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


def _tunnel_component_name(tunnel_num, position, thk, is_tbm=False):
    """Return the display name for one generated tunnel component.

    TBM names deliberately describe the physical component while ordinary
    tunnel names retain the established naming convention.
    """
    if is_tbm:
        tbm_labels = {
            'primary': 'TBM',
            'grout': 'grout',
            'secondary': 'MainLiner',
        }
        position_label = tbm_labels.get(position, position)
        prefix = f'tunnel{tunnel_num}_{position_label}'
    else:
        prefix = (f'tunnel{tunnel_num}_{position}'
                  if tunnel_num is not None else f'tunnel_{position}')
    suffix = f'THK={thk}m' if thk > 0 else 'THK=?'
    return f'{prefix}_{suffix}'


def _clean_mat_var_name(tunnel_name):
    """Convert tunnel name string to a valid Python variable name for material.

    Examples:
        'tunnel3_primary_THK=0.5m'  -> 'tunnel3_primary_THK0_5m'
        'tunnel_secondary_THK=0.3m' -> 'tunnel_secondary_THK0_3m'
    """
    name = tunnel_name.replace('=', '').replace('.', '_')
    return name

# ── Soil type mapping (same as sc3/sc2) ──
SOIL_TYPE_MAP = {
    'clay-HS':       ('Hardening Soil',  'Undrained A',  False, False, 'Clay'),
    'sand-HS':       ('Hardening Soil',  'Drained',      False, False, 'Sand'),
    'silt-HS':       ('Hardening Soil',  'Undrained A',  False, False, 'Silt'),
    'clay-HSS':      ('HS Small',        'Undrained A',  True,  False, 'Clay'),
    'sand-HSS':      ('HS Small',        'Drained',      True,  False, 'Sand'),
    'silt-HSS':      ('HS Small',        'Undrained A',  True,  False, 'Silt'),
    'clay-SSC':      ('Soft Soil Creep', 'Undrained A',  False, True,  'Clay'),
    'sand-SSC':      ('Soft Soil Creep', 'Drained',      False, True,  'Sand'),
    'silt-SSC':      ('Soft Soil Creep', 'Undrained A',  False, True,  'Silt'),
    'linear-elastic': ('Linear Elastic', 'Non-porous',   False, False, None),
}


def safe_zero(val):
    """Convert xlwings value to float, returning 0 if None, NaN, or empty."""
    if val is None:
        return 0
    try:
        if isnan(val):
            return 0
    except (TypeError, ValueError):
        pass
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0


def _gw_frac_is_filled(val):
    """True iff a clay_frac/silt_frac cell is genuinely filled.

    Blank means None, NaN, or empty/whitespace string. Explicit 0 is
    filled and still emits (PLAXIS distinguishes 0 from class default).
    """
    if val is None:
        return False
    try:
        if isnan(val):
            return False
    except (TypeError, ValueError):
        pass
    if isinstance(val, str) and val.strip() == '':
        return False
    return True


def _normalise_gw_mode(mode):
    """Normalise groundwater mode; alias truncated 'by_grain_s' cells."""
    if mode is None:
        return 'automatic'
    try:
        if isnan(mode):
            return 'automatic'
    except (TypeError, ValueError):
        pass
    m = safe_str(mode)
    if m == '' or m.lower() == 'nan':
        return 'automatic'
    if m == 'by_grain_s':
        return 'by_grain_size'
    return m


def fmt(val):
    """Format a value for Python code output — strings in quotes, numbers as-is."""
    if isinstance(val, str):
        return f'"{val}"'
    return str(val)


def make_cell(source, cell_type='code'):
    # Jupyter source must be a single string (with \n) or list of strings each ending with \n
    if isinstance(source, list):
        source = '\n'.join(source)
    return {
        'cell_type': cell_type,
        'metadata': {},
        'source': [source],
        'outputs': [],
        'execution_count': None,
    }


def safe_num(val, default=0):
    """Convert xlwings value to float, returning default if None or empty."""
    if val is None:
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def safe_str(val, default=''):
    if val is None:
        return default
    return str(val).strip()


def round_num(val, decimals=3):
    """Format a number for PLAXIS: integers stay clean, floats rounded to max decimals."""
    try:
        f = float(val)
    except (TypeError, ValueError):
        return str(val)
    if f == int(f) and abs(f) < 1e15:
        return str(int(f))
    rounded = round(f, decimals)
    # Strip trailing zeros but keep at least one decimal for pure floats
    s = f'{rounded:.{decimals}f}'.rstrip('0').rstrip('.')
    if '.' in s:
        return s
    # It was something like 1234.000 -> 1234
    return str(int(rounded))


def gen_tunnel_contraction_code(tunnel_var, c_ref):
    """Apply one boundary contraction to every slice segment of a tunnel.

    A generated circular tunnel consists of multiple slice segments (e.g.
    two half-circles), so the same C value must be attached to each segment
    to represent a single uniform boundary contraction.  The command must
    run before ``generatethicklining``/``generatetunnel``.  C_ref comes from
    the primary row only; the secondary liner receives no separate
    contraction.
    """
    c_value = safe_num(c_ref)
    if c_value <= 0:
        return []
    return [
        f'for _slice in {tunnel_var}.SliceSegments[:]:',
        f'    g_i.contraction(_slice, "C", {round_num(c_value)})',
    ]


# ── Read section data from str_2D ──

def read_section(ws, start_row, max_rows=6):
    """
    Read headers from start_row (column C onwards) and data rows below.
    Returns (data_list, col_map) where:
      - data_list: list of dicts, one per non-empty row
      - col_map: {header_name: column_number} for writing back to Excel
    Column B is element_name (auto-generated by sc5), so non-empty check
    uses column C (cad_layer) which is populated by sc4 or user input.

    Optimized: block reads reduce per-cell COM round-trips.
    """
    last_col = ws.range(start_row, 3).end('right').column

    # Block read: headers row (one COM call for all columns B..last_col)
    header_vals = ws.range((start_row, 2), (start_row, last_col)).value
    if not isinstance(header_vals, list):
        header_vals = [header_vals]
    headers = [safe_str(h) for h in header_vals]

    # Build header_name → column_number mapping
    col_map = {}
    for c, h in enumerate(headers, 2):
        if h:
            col_map[h] = c

    # Block read: data rows (one COM call for all data cells)
    if max_rows < 1:
        return [], col_map
    data_start = start_row + 1
    data_end = start_row + max_rows
    block = ws.range((data_start, 2), (data_end, last_col)).value
    if block is None:
        return [], col_map
    if not isinstance(block, list):
        block = [block]

    # Filter non-empty rows and build dicts from block data
    data = []
    for row_vals in block:
        if not isinstance(row_vals, list):
            row_vals = [row_vals] * (last_col - 1)
        # Column C = index 1 in our block (block starts at col B = index 0)
        cad_layer_idx = 1  # column C is index 1 in block (B=0, C=1)
        if cad_layer_idx < len(row_vals):
            first = row_vals[cad_layer_idx]
        else:
            first = None
        if first is None or str(first).strip() == '':
            continue
        row_dict = {}
        for h, c in col_map.items():
            col_idx = c - 2  # convert 1-based column to 0-based block index
            if col_idx < len(row_vals):
                row_dict[h] = row_vals[col_idx]
            else:
                row_dict[h] = None
        data.append(row_dict)
    return data, col_map


def write_orange(ws, row, col_map, header, value):
    """Write a value to an orange column by header name. Silent no-op if header not found."""
    c = col_map.get(header)
    if c is not None:
        ws.range(row, c).value = value




# ── Geogrid input table (str_2D rows 211:230; header row 210) ──
GEOGRID_HEADER_ROW = 210
GEOGRID_DATA_START = 211
GEOGRID_DATA_END = 230


def read_geogrids(sheet, header_row=GEOGRID_HEADER_ROW, max_rows=20):
    """Read the current geogrid table using its B:R headers.

    SC4 owns only C:G. SC5 reads all B:R fields but does not yet emit
    unverified PLAXIS geogrid commands.
    """
    header_vals = sheet.range((header_row, 2), (header_row, 18)).value
    if not isinstance(header_vals, list):
        header_vals = [header_vals]
    headers = [safe_str(value) for value in header_vals]
    col_map = {header: index for index, header in enumerate(headers, 2) if header}

    data = sheet.range(
        (GEOGRID_DATA_START, 2),
        (GEOGRID_DATA_START + max_rows - 1, 18),
    ).value
    if data is None:
        return []
    if not isinstance(data, list):
        data = [data]

    rows = []
    for offset, values in enumerate(data):
        if not isinstance(values, list):
            values = [values] * len(headers)
        cad_layer = safe_str(values[col_map.get('cad_layer', 3) - 2]) \
            if col_map.get('cad_layer', 3) - 2 < len(values) else ''
        if not cad_layer:
            continue
        row = {'_row': GEOGRID_DATA_START + offset}
        for header, col in col_map.items():
            index = col - 2
            row[header] = values[index] if index < len(values) else None
        rows.append(row)
    return rows


def geogrid_display_name(geogrid, index):
    """Return sequential geogridN_<DXF bracket text> name for column B."""
    layer = safe_str(geogrid.get('cad_layer'))
    match = re.search(r'\[([^\]]+)\]', layer)
    if match:
        return f'geogrid{index}_{match.group(1)}'
    return f'geogrid{index}'


def calculate_geogrid_ea(geogrid):
    """Calculate EA1 and EA2 in kN/m from E in MPa and area in m²/m."""
    e_mpa = safe_num(geogrid.get('gg_E(Mpa)'))
    a1 = safe_num(geogrid.get('gg_A1(m2)'))
    a2 = safe_num(geogrid.get('gg_A2(m2)'))
    ea1 = e_mpa * 1000.0 * a1 if e_mpa > 0 and a1 > 0 else 0
    ea2 = e_mpa * 1000.0 * a2 if e_mpa > 0 and a2 > 0 else 0
    return ea1, ea2


def _is_yes(value):
    return safe_str(value).lower() in ('yes', 'true', '1', 'y')


def _normalize_isotropic(value, default='yes'):
    """Normalize Excel yes/no or Boolean isotropic input to yes/no."""
    text = safe_str(value).lower()
    if text in ('yes', 'true', '1', 'y'):
        return 'yes'
    if text in ('no', 'false', '0', 'n'):
        return 'no'
    return default


def _first_present(data, *keys):
    """Return the first present, nonblank value for compatible headers.
    Falls back to case-insensitive key search if none of the supplied
    keys match exactly — this handles both ``Np1Tens`` and ``Np1tens``."""
    for key in keys:
        value = data.get(key)
        if value is not None and safe_str(value) not in ('', 'N/A'):
            return value
    # Case-insensitive fallback: find any key whose lowered form matches.
    for key in keys:
        target = key.lower()
        for dk, dv in data.items():
            if safe_str(dk).lower() == target and safe_str(dv) not in ('', 'N/A'):
                return dv
    return None


def validate_geogrid_row(geogrid):
    """Return geogrid-specific warning strings; Mp is not applicable."""
    warnings = []
    x1, y1 = geogrid.get('x1'), geogrid.get('y1')
    x2, y2 = geogrid.get('x2'), geogrid.get('y2')
    if x1 is None or y1 is None or x2 is None or y2 is None:
        warnings.append('MISSING geogrid coordinates')
    elif safe_num(x1) == safe_num(x2) and safe_num(y1) == safe_num(y2):
        warnings.append('INVALID zero-length geogrid')

    e_mpa = safe_num(geogrid.get('gg_E(Mpa)'))
    a1 = safe_num(geogrid.get('gg_A1(m2)'))
    a2 = safe_num(geogrid.get('gg_A2(m2)'))
    if e_mpa <= 0:
        warnings.append('MISSING gg_E(Mpa)')
    if a1 <= 0:
        warnings.append('MISSING gg_A1(m2)')
    if a2 <= 0:
        warnings.append('MISSING gg_A2(m2)')
    if _is_yes(geogrid.get('isotropic')) and a1 > 0 and a2 > 0 and a1 != a2:
        warnings.append('MISMATCH isotropic geogrid requires gg_A2 = gg_A1')

    model = safe_str(geogrid.get('gg_model'), 'elastic').lower()
    if model == 'elastic':
        geogrid['Np1(kN/m)'] = 'N/A'
        geogrid['Np2(kN/m)'] = 'N/A'
    elif _is_yes(geogrid.get('isotropic')):
        geogrid['Np2(kN/m)'] = 'N/A'

    if model != 'elastic' and safe_num(geogrid.get('Np1(kN/m)')) <= 0:
        warnings.append('MISSING Np1(kN/m)')
    if model != 'elastic' and not _is_yes(geogrid.get('isotropic')):
        if safe_num(geogrid.get('Np2(kN/m)')) <= 0:
            warnings.append('MISSING Np2(kN/m)')

    interface = safe_str(geogrid.get('interface')).lower()
    if interface not in ('positive', 'negative', 'both', 'none'):
        warnings.append('MISSING or INVALID interface')
    return warnings


def calc_plate(d):
    """Calculate plate EA1, EI, EA2 from yellow input cells.
    Returns (ea1, ei, ea2) in kN/m and kNm²/m."""
    e_mpa = safe_num(d.get('plate_E(Mpa)'))
    a1 = safe_num(d.get('plate_A1(m2)'))
    a2 = safe_num(d.get('plate_A2(m2)'))
    thk = safe_num(d.get('plate_THK(m)'))

    ea1 = e_mpa * 1000 * a1 if e_mpa > 0 and a1 > 0 else 0
    h = a2 if a2 > 0 else thk
    ei = e_mpa * 1000 * (h ** 3) / 12 if e_mpa > 0 and h > 0 else 0
    ea2 = e_mpa * 1000 * a2 if e_mpa > 0 and a2 > 0 else 0

    return ea1, ei, ea2


def calc_n2n_ea(d):
    """Calculate N2N anchor EA from yellow input cells. Returns EA in kN."""
    e_mpa = safe_num(d.get('nn_E(Mpa)'))
    a = safe_num(d.get('nn_A(m2)'))
    return e_mpa * 1000 * a if e_mpa > 0 and a > 0 else 0


def gen_geogrid_mat(idx, d):
    """Generate PLAXIS V2025 geogrid material code.

    EA1 is the confirmed V2025 geogridmat property. EA2 and tensile
    capacities are emitted only when the corresponding input is active.
    Geogrids have no plate bending property such as Mp or EI.
    """
    name = safe_str(d.get('geogrid_name'), f'geogrid_{idx}')
    model = safe_str(d.get('gg_model'), 'elastic').lower()
    material_type = 'Elastoplastic' if model == 'elastoplastic' else 'Elastic'
    ea1 = safe_num(d.get('gg_EA1'))
    ea2 = safe_num(d.get('gg_EA2'))
    iso = _normalize_isotropic(d.get('isotropic'))

    props = [
        '"Identification"', f'"{name}"',
        '"MaterialType"', f'"{material_type}"',
        '"EA1"', round_num(ea1),
    ]
    if iso == 'no':
        props.extend(['"isotropic"', 'False', '"EA2"', round_num(ea2)])
    if model == 'elastoplastic':
        np1 = safe_num(d.get('Np1(kN/m)'))
        np2 = safe_num(d.get('Np2(kN/m)'))
        if np1 > 0:
            props.extend(['"Np1"', round_num(np1)])
        if np2 > 0 and iso == 'no':
            props.extend(['"Np2"', round_num(np2)])

    return [
        f'# ── Geogrid material {idx}: {name} ({material_type}) ──',
        f'geogrid_mat_{idx} = g_i.geogridmat({", ".join(props)})',
    ]


def gen_geogrid_geom(idx, d, mat_idx=None):
    """Generate one line and one V2025 geogrid assigned to its material."""
    if mat_idx is None:
        mat_idx = idx
    x1 = safe_num(d.get('x1'))
    y1 = safe_num(d.get('y1'))
    x2 = safe_num(d.get('x2'))
    y2 = safe_num(d.get('y2'))
    return [
        f'gg_line_{idx} = g_i.line(({round_num(x1)}, {round_num(y1)}), ({round_num(x2)}, {round_num(y2)}))[-1]',
        f'geogrid_{idx} = g_i.geogrid(gg_line_{idx}, "Material", geogrid_mat_{mat_idx})',
    ]


def geogrid_interface_labels(index, d, positive_index, negative_index):
    """Return the PLAXIS global names for this geogrid's interfaces.

    PLAXIS numbers positive and negative interfaces globally across every
    interface-producing element; the geogrid index is only used for the
    underlying line/local Python variable and is not part of the PLAXIS name.
    """
    interface_type = safe_str(d.get('interface'), 'none').lower() or 'none'
    positive_label = None
    negative_label = None
    if interface_type in ('positive', 'both'):
        positive_label = f'PositiveInterface_{positive_index}'
    if interface_type in ('negative', 'both'):
        negative_label = f'NegativeInterface_{negative_index}'
    return positive_label, negative_label


def gen_geogrid_interface(idx, d):
    """Generate interfaces for one geogrid's own underlying line."""
    interface_type = safe_str(d.get('interface'), 'none').lower()
    if not interface_type:
        interface_type = 'none'

    lines = []
    if interface_type in ('positive', 'both'):
        lines.extend([
            f'gg_int_pos_{idx} = g_i.posinterface(gg_line_{idx})',
            f'gg_int_pos_{idx}.ActiveInFlow = True',
        ])
    if interface_type in ('negative', 'both'):
        lines.extend([
            f'gg_int_neg_{idx} = g_i.neginterface(gg_line_{idx})',
            f'gg_int_neg_{idx}.ActiveInFlow = True',
        ])
    return lines


# ── custom_wall helpers (v0.7.4) ──

def custom_wall_plate_name(cad_layer):
    """Extract bracket content from cad_layer and normalize for PLAXIS material name.

    Example: plate_1[combi_wall - OD1200] → combi_wall_OD1200
    Hyphens removed; whitespace runs collapse to a single underscore.
    Returns None if no usable bracket content is found.
    """
    if not cad_layer:
        return None
    m = re.search(r'\[([^\]]*)\]', str(cad_layer))
    if m is None:
        return None
    text = m.group(1).replace('-', '')
    text = re.sub(r'\s+', '_', text).strip('_')
    return text if text else None


def resolve_plate_stiffness(d):
    """Resolve EA1, EI, EA2 for a plate row dict.

    For custom_wall: read EI/EA1/EA2 directly from the user-typed values in
    plate_EI / plate_EA1 / plate_EA2 columns (same columns as ordinary plates).
    For all other types: calculate from yellow E/A1/A2/THK inputs.
    Returns (ea1, ei, ea2).
    """
    is_custom = safe_str(d.get('plate_type'), '').lower() == 'custom_wall'
    if is_custom:
        ei = safe_num(d.get('plate_EI'))
        ea1 = safe_num(d.get('plate_EA1'))
        ea2 = safe_num(d.get('plate_EA2'))
        return ea1, ei, ea2
    return calc_plate(d)


def gen_plate_mat(idx, d, plate_material_index=0):
    """Generate plate material code using setproperties().
    Includes warnings for missing inputs and unit comments.  Plate colours
    are assigned by unique material index and cycle through the fixed palette."""
    mat_name = safe_str(d.get('plate_name'), f'pl_mat_{idx}')
    model = safe_str(d.get('plate_model'), 'elastic').lower()
    iso = _normalize_isotropic(d.get('isotropic'))
    punch = safe_str(d.get('prevent_punching'), 'no').lower()
    nu = safe_num(d.get('v_nu'), 0.15)
    w = safe_num(d.get('plate_w(kN/m/m)'))
    mat_type = 2 if model == 'elastoplastic' else 1

    ea1 = safe_num(d.get('plate_EA1'))
    ei = safe_num(d.get('plate_EI'))

    lines = [f"# ── Plate {idx}: {mat_name} ({model}, {'isotropic' if iso == 'yes' else 'anisotropic'}) ──"]

    # Unit comment
    lines.append(f'# Units: EA1(kN/m) EI(kNm²/m) StructNu(-) w(kN/m/m) Mp(kNm/m) Np1Tens(kN/m) Np2Tens(kN/m) EA2(kN/m)')

    # Warnings
    is_custom = safe_str(d.get('plate_type'), '').lower() == 'custom_wall'
    if is_custom:
        missing = []
        if ei == 0: missing.append('plate_EI')
        if ea1 == 0: missing.append('plate_EA1')
        if iso == 'no' and safe_num(d.get('plate_EA2')) == 0:
            missing.append('plate_EA2')
        if missing:
            lines.append(f'print("WARNING Plate {idx}: custom_wall missing: {", ".join(missing)}")')
    elif model == 'elastoplastic':
        mp_val = safe_str(d.get('Mp(kNm/m)'))
        np1_val = safe_str(_first_present(d, 'Np1(kN/m)', 'Np1tens(kN/m)'))
        np2_val = safe_str(_first_present(d, 'Np2(kN/m)', 'Np2tens(kN/m)'))
        missing = []
        if mp_val in ('', 'N/A', '0'): missing.append('Mp')
        if np1_val in ('', 'N/A', '0'): missing.append('Np1Tens')
        if np2_val in ('', 'N/A', '0'): missing.append('Np2Tens')
        if missing:
            lines.append(f'print("WARNING Plate {idx}: elastoplastic but missing: {", ".join(missing)}")')
    if iso == 'no' and not is_custom:
        a2_val = safe_str(d.get('plate_A2(m2)'))
        if a2_val in ('', 'N/A', '0'):
            lines.append(f'print("WARNING Plate {idx}: anisotropic but plate_A2 is empty")')

    # Build all properties in one setproperties call
    props = [
        '"Identification"', f'"{mat_name}"',
        '"MaterialType"', str(mat_type),
        '"EA1"', round_num(ea1),
        '"EI"', round_num(ei),
        '"w"', round_num(w),
        '"PreventPunching"', 'True' if punch == 'yes' else 'False',
    ]

    # Isotropic → add StructNu; Anisotropic → add isotropic=False + EA2
    if iso == 'yes':
        props.extend(['"StructNu"', round_num(nu)])
    else:
        if is_custom:
            ea2_val = safe_num(d.get('plate_EA2'))
        else:
            e_mpa = safe_num(d.get('plate_E(Mpa)'))
            a2 = safe_num(d.get('plate_A2(m2)'))
            ea2_val = e_mpa * 1000 * a2 if e_mpa > 0 and a2 > 0 else 0
        props.extend(['"isotropic"', 'False', '"EA2"', round_num(ea2_val)])

    # Elastoplastic → add Mp, Np1Tens, Np2Tens
    if model == 'elastoplastic':
        mp_val = d.get('Mp(kNm/m)')
        np1_val = _first_present(d, 'Np1(kN/m)', 'Np1tens(kN/m)')
        np2_val = _first_present(d, 'Np2(kN/m)', 'Np2tens(kN/m)')
        if safe_str(mp_val) not in ('', 'N/A'):
            props.extend(['"Mp"', round_num(safe_num(mp_val))])
        if safe_str(np1_val) not in ('', 'N/A'):
            props.extend(['"Np1Tens"', round_num(safe_num(np1_val))])
        if safe_str(np2_val) not in ('', 'N/A'):
            props.extend(['"Np2Tens"', round_num(safe_num(np2_val))])

    lines.append(f'pl_mat_{idx} = g_i.platemat()')
    lines.append(f'pl_mat_{idx}.setproperties({", ".join(props)})')
    rgb = compute_plate_color(plate_material_index)
    lines.append(f'pl_mat_{idx}.setcolour({rgb[0]}, {rgb[1]}, {rgb[2]})')
    return lines


def gen_plate_geom(idx, d, mat_idx=None):
    if mat_idx is None:
        mat_idx = idx
    x1 = safe_num(d.get('x1'))
    y1 = safe_num(d.get('y1'))
    x2 = safe_num(d.get('x2'))
    y2 = safe_num(d.get('y2'))
    return [
        f'pl_{idx} = g_i.plate({round_num(x1)}, {round_num(y1)}, {round_num(x2)}, {round_num(y2)})',
        f'for p in pl_{idx}: p.Material = pl_mat_{mat_idx}',
    ]


def gen_n2n_mat(idx, d):
    """Generate N2N anchor material code using setproperties()."""
    model = safe_str(d.get('nn_model'), 'elastic').lower()
    spacing = safe_num(d.get('L_spacing(m)'), 1)
    anchor_name = safe_str(d.get('nn_name'), f'nn_mat_{idx}')
    mat_type = 2 if model == 'elastoplastic' else 1

    ea = safe_num(d.get('nn_EA'))

    props = [
        '"Identification"', f'"{anchor_name}"',
        '"MaterialType"', str(mat_type),
        '"EA"', round_num(ea),
        '"LSpacing"', round_num(spacing),
    ]

    if model == 'elastoplastic':
        ft_val = d.get('Fmax_tens')
        fc_val = d.get('Fmax_comp')
        if safe_str(ft_val) not in ('', 'N/A'):
            props.extend(['"FmaxTens"', round_num(safe_num(ft_val))])
        if safe_str(fc_val) not in ('', 'N/A'):
            props.extend(['"FmaxComp"', round_num(safe_num(fc_val))])

    props_str = ', '.join(props)

    lines = [f"# ── N2N Anchor {idx}: {anchor_name} ({model}) ──"]
    lines.append(f'# Units: EA(kN) LSpacing(m) FmaxTens(kN) FmaxComp(kN)')
    lines.append(f'nn_mat_{idx} = g_i.anchormat()')
    lines.append(f'nn_mat_{idx}.setproperties({props_str})')
    return lines


def gen_n2n_geom(idx, d, mat_idx=None):
    if mat_idx is None:
        mat_idx = idx
    x1 = safe_num(d.get('x1'))
    y1 = safe_num(d.get('y1'))
    x2 = safe_num(d.get('x2'))
    y2 = safe_num(d.get('y2'))
    return [
        f'nn_{idx} = g_i.n2nanchor({round_num(x1)}, {round_num(y1)}, {round_num(x2)}, {round_num(y2)})',
        f'for a in nn_{idx}: a.Material = nn_mat_{mat_idx}',
    ]


def gen_anc_mat(idx, d, embedded_beam_index=0):
    """Generate embedded beam material code using setproperties().
    PLAXIS needs E (modulus in kN/m²) and Diameter — it calculates A, I, EA, EI internally.
    ``embedded_beam_index`` is shared across piles, nails, and grout anchors."""
    model = safe_str(d.get('anc_model'), 'elastic').lower()
    spacing = safe_num(d.get('L_spacing(m)'), 1)
    e_mpa = safe_num(d.get('anc_E(Mpa)'))
    w = safe_num(d.get('anc_w(kN/m/m)'))
    dia = safe_num(d.get('diameter(m)'))
    res = safe_str(d.get('resistance_model'), 'linear').lower()
    t_start = safe_num(d.get('Tskin_start'))
    t_end = safe_num(d.get('Tskin_end'))
    anc_name = safe_str(d.get('anc_name'), f'anc_mat_{idx}')
    mat_type = 2 if model == 'elastoplastic' else 1

    # AxialSkinResistance: 0=linear, 2=layer_dependent
    res_val = 0 if res == 'linear' else 2

    # Convert E from MPa to kN/m² for PLAXIS
    e_knm2 = e_mpa * 1000

    props = [
        '"Identification"', f'"{anc_name}"',
        '"MaterialType"', str(mat_type),
        '"E"', round_num(e_knm2),
        '"Diameter"', round_num(dia),
        '"LSpacing"', round_num(spacing),
        '"Gamma"', round_num(w),
        '"AxialSkinResistance"', str(res_val),
        '"TSkinStartMax"', round_num(t_start),
        '"TSkinEndMax"', round_num(t_end),
    ]

    if model == 'elastoplastic':
        props.extend(['"FMax"', '0'])
        mp_val = d.get('Mp(kNm/m)')
        np_val = d.get('Nptens(kN/m)')
        if safe_str(mp_val) not in ('', 'N/A'):
            props.extend(['"Mp"', round_num(safe_num(mp_val))])
        if safe_str(np_val) not in ('', 'N/A'):
            props.extend(['"Nptens"', round_num(safe_num(np_val))])

    lines = [f"# ── Embedded Beam {idx}: {anc_name} ({model}, {res}) ──"]
    lines.append(f'# Units: E(kN/m²) Diameter(m) LSpacing(m) Gamma(kN/m/m) TSkinStartMax(kN/m) TSkinEndMax(kN/m) Mp(kNm/m) Nptens(kN/m) FMax(kN)')
    lines.append(f'anc_mat_{idx} = g_i.embeddedbeammat()')
    lines.append(f'anc_mat_{idx}.setproperties({", ".join(props)})')
    rgb = compute_embedded_beam_color(embedded_beam_index)
    lines.append(f'anc_mat_{idx}.setcolour({rgb[0]}, {rgb[1]}, {rgb[2]})')
    return lines


def gen_anc_geom(idx, d, mat_idx=None):
    if mat_idx is None:
        mat_idx = idx
    x1 = safe_num(d.get('x1'))
    y1 = safe_num(d.get('y1'))
    x2 = safe_num(d.get('x2'))
    y2 = safe_num(d.get('y2'))
    return [
        f'anc_{idx} = g_i.embeddedbeam({round_num(x1)}, {round_num(y1)}, {round_num(x2)}, {round_num(y2)})',
        f'for a in anc_{idx}: a.Material = anc_mat_{mat_idx}',
    ]


# ── Strut functions (same PLAXIS type as N2N anchor) ──

def gen_strut_mat(idx, d):
    """Generate strut material code — same PLAXIS type as N2N anchor (g_i.anchormat())."""
    model = safe_str(d.get('nn_model'), 'elastic').lower()
    spacing = safe_num(d.get('L_spacing(m)'), 1)
    anchor_name = safe_str(d.get('nn_name'), f'strut_mat_{idx}')
    mat_type = 2 if model == 'elastoplastic' else 1
    ea = safe_num(d.get('nn_EA'))
    props = [
        '"Identification"', f'"{anchor_name}"',
        '"MaterialType"', str(mat_type),
        '"EA"', round_num(ea),
        '"LSpacing"', round_num(spacing),
    ]
    if model == 'elastoplastic':
        ft_val = d.get('Fmax_tens')
        fc_val = d.get('Fmax_comp')
        if safe_str(ft_val) not in ('', 'N/A'):
            props.extend(['"FmaxTens"', round_num(safe_num(ft_val))])
        if safe_str(fc_val) not in ('', 'N/A'):
            props.extend(['"FmaxComp"', round_num(safe_num(fc_val))])
    props_str = ', '.join(props)
    lines = [f"# ── Strut {idx}: {anchor_name} ({model}) ──"]
    lines.append(f'strut_mat_{idx} = g_i.anchormat()')
    lines.append(f'strut_mat_{idx}.setproperties({props_str})')
    return lines


def gen_strut_geom(idx, d, mat_idx=None):
    """Generate strut geometry code — uses strut_ prefix for geometry, strut_mat_ for material."""
    if mat_idx is None:
        mat_idx = idx
    x1 = safe_num(d.get('x1'))
    y1 = safe_num(d.get('y1'))
    x2 = safe_num(d.get('x2'))
    y2 = safe_num(d.get('y2'))
    return [
        f'strut_{idx} = g_i.n2nanchor({round_num(x1)}, {round_num(y1)}, {round_num(x2)}, {round_num(y2)})',
        f'for a in strut_{idx}: a.Material = strut_mat_{mat_idx}',
    ]


# ── FEA (fixed end anchor) functions — same material as N2N, point geometry ──

def gen_fea_mat(idx, d):
    """Generate FEA material code — same anchormat as N2N anchor."""
    model = safe_str(d.get('nn_model'), 'elastic').lower()
    spacing = safe_num(d.get('L_spacing(m)'), 1)
    anchor_name = safe_str(d.get('fe_name'), f'fea_mat_{idx}')
    mat_type = 2 if model == 'elastoplastic' else 1

    ea = safe_num(d.get('nn_EA'))

    props = [
        '"Identification"', f'"{anchor_name}"',
        '"MaterialType"', str(mat_type),
        '"EA"', round_num(ea),
        '"LSpacing"', round_num(spacing),
    ]

    if model == 'elastoplastic':
        ft_val = d.get('Fmax_tens')
        fc_val = d.get('Fmax_comp')
        if safe_str(ft_val) not in ('', 'N/A'):
            props.extend(['"FmaxTens"', round_num(safe_num(ft_val))])
        if safe_str(fc_val) not in ('', 'N/A'):
            props.extend(['"FmaxComp"', round_num(safe_num(fc_val))])

    props_str = ', '.join(props)

    lines = [f"# ── FEA {idx}: {anchor_name} ({model}) ──"]
    lines.append(f'# Units: EA(kN) LSpacing(m) FmaxTens(kN) FmaxComp(kN)')
    lines.append(f'fea_mat_{idx} = g_i.anchormat()')
    lines.append(f'fea_mat_{idx}.setproperties({props_str})')
    return lines


def gen_fea_geom(idx, d, mat_idx=None):
    """Generate FEA geometry — point element with Direction_x property."""
    if mat_idx is None:
        mat_idx = idx
    x1 = safe_num(d.get('x1'))
    y1 = safe_num(d.get('y1'))
    direction_x = safe_num(d.get('Direction_x'))
    lines = [
        f'fea_{idx} = g_i.fixedendanchor({round_num(x1)}, {round_num(y1)})',
        f'for a in fea_{idx}: a.Material = fea_mat_{mat_idx}',
        f'for a in fea_{idx}: a.Direction_x = {round_num(direction_x)}',
    ]
    return lines


# ── Pile functions (PLAXIS embedded beam with PredefinedCrossSectionType) ──

def gen_pile_mat(idx, d, embedded_beam_index=0):
    """Generate pile material code using setproperties().
    Sets PredefinedCrossSectionType and correct geometry properties per pile type.
    solid_circular → "Diameter" only
    circular_tube  → "Diameter" + "Thickness"
    square         → "Width" only
    ``embedded_beam_index`` is shared across piles, nails, and grout anchors."""
    model = safe_str(d.get('anc_model'), 'elastic').lower()
    spacing = safe_num(d.get('L_spacing(m)'), 1)
    e_mpa = safe_num(d.get('anc_E(Mpa)'))
    w = safe_num(d.get('anc_w(kN/m/m)'))
    res = safe_str(d.get('resistance_model'), 'linear').lower()
    t_start = safe_num(d.get('Tskin_start'))
    t_end = safe_num(d.get('Tskin_end'))
    pile_name = safe_str(d.get('pile_name'), f'pile_mat_{idx}')
    mat_type = 2 if model == 'elastoplastic' else 1
    pile_type = safe_str(d.get('anc_type'), 'solid_circular').lower()

    res_val = 0 if res == 'linear' else 2
    e_knm2 = e_mpa * 1000

    # Cross-section type mapping
    cs_map = {
        'solid_circular': 'Solid circular beam',
        'circular_tube':  'Circular tube',
        'square':         'Solid square beam',
    }
    cs_type = cs_map.get(pile_type, 'Solid circular beam')

    lines = [f"# ── Pile {idx}: {pile_name} ({pile_type}, {model}, {res}) ──"]

    # pile_THK warnings
    thk_val = safe_str(d.get('pile_THK(m)'))
    if pile_type == 'circular_tube':
        if thk_val in ('', 'N/A', '0', 'None'):
            lines.append(f'print("WARNING Pile {idx}: circular_tube but pile_THK is empty")')
    else:
        if thk_val not in ('', 'N/A', '0', 'None', 'only for circular tube'):
            lines.append(f'print("WARNING Pile {idx}: {pile_type} does not use pile_THK")')

    # Elastoplastic warnings
    if model == 'elastoplastic':
        mp_val = safe_str(d.get('Mp(kNm/m)'))
        np_val = safe_str(d.get('Nptens(kN/m)'))
        missing = []
        if mp_val in ('', 'N/A', '0', 'None'):
            missing.append('Mp')
        if np_val in ('', 'N/A', '0', 'None'):
            missing.append('Nptens')
        if missing:
            lines.append(f'print("WARNING Pile {idx}: elastoplastic but missing: {", ".join(missing)}")')

    # Base properties
    props = [
        '"Identification"', f'"{pile_name}"',
        '"MaterialType"', str(mat_type),
        '"PredefinedCrossSectionType"', f'"{cs_type}"',
        '"E"', round_num(e_knm2),
        '"LSpacing"', round_num(spacing),
        '"Gamma"', round_num(w),
        '"AxialSkinResistance"', str(res_val),
        '"TSkinStartMax"', round_num(t_start),
        '"TSkinEndMax"', round_num(t_end),
    ]

    # Geometry per pile type
    dia = safe_num(d.get('diameter/width(m)'))
    thk_raw = safe_str(d.get('pile_THK(m)'))
    thk = safe_num(thk_raw)

    if pile_type == 'solid_circular':
        props.extend(['"Diameter"', round_num(dia)])
    elif pile_type == 'circular_tube':
        props.extend(['"Diameter"', round_num(dia), '"Thickness"', round_num(thk)])
    elif pile_type == 'square':
        props.extend(['"Width"', round_num(dia)])

    # Fmax — applies to all pile models (elastic and elastoplastic)
    fmax_val = d.get('Fmax')
    if safe_str(fmax_val) not in ('', 'N/A', 'None'):
        props.extend(['"FMax"', round_num(safe_num(fmax_val))])

    # Elastoplastic strength
    if model == 'elastoplastic':
        mp_val = d.get('Mp(kNm/m)')
        np_val = d.get('Nptens(kN/m)')
        if safe_str(mp_val) not in ('', 'N/A', 'None'):
            props.extend(['"Mp"', round_num(safe_num(mp_val))])
        if safe_str(np_val) not in ('', 'N/A', 'None'):
            props.extend(['"Nptens"', round_num(safe_num(np_val))])

    lines.append(f'pile_mat_{idx} = g_i.embeddedbeammat()')
    lines.append(f'pile_mat_{idx}.setproperties({", ".join(props)})')
    rgb = compute_embedded_beam_color(embedded_beam_index)
    lines.append(f'pile_mat_{idx}.setcolour({rgb[0]}, {rgb[1]}, {rgb[2]})')
    return lines


def gen_pile_geom(idx, d, mat_idx=None):
    if mat_idx is None:
        mat_idx = idx
    x1 = safe_num(d.get('x1'))
    y1 = safe_num(d.get('y1'))
    x2 = safe_num(d.get('x2'))
    y2 = safe_num(d.get('y2'))
    return [
        f'pile_{idx} = g_i.embeddedbeam({round_num(x1)}, {round_num(y1)}, {round_num(x2)}, {round_num(y2)})',
        f'for p in pile_{idx}: p.Material = pile_mat_{mat_idx}',
    ]


def gen_plate_interface(idx, d):
    """Generate interface code for a plate element.
    Checks 'interface' column: none/positive/negative/both.
    Uses pl_{idx}[2] to target the Line element only.
    Also sets ActiveInFlow = True for each interface.
    WARNING: g_i.plate() returns [Point_1, Point_2, Line_1].
    If PLAXIS version changes list order, [2] may break."""
    plate_type = safe_str(d.get('plate_type'), 'plate').lower()
    interface_type = safe_str(d.get('interface'), 'none').lower()

    lines = []
    if interface_type not in ('positive', 'negative', 'both'):
        return lines

    if interface_type in ('positive', 'both'):
        lines.append(f'pl_int_pos_{idx} = g_i.posinterface(pl_{idx}[2])')
        lines.append(f'pl_int_pos_{idx}.ActiveInFlow = True')

    if interface_type in ('negative', 'both'):
        lines.append(f'pl_int_neg_{idx} = g_i.neginterface(pl_{idx}[2])')
        lines.append(f'pl_int_neg_{idx}.ActiveInFlow = True')

    return lines


def _is_volume_profile_type(value):
    """Return True for volume types: default, interface modes, or 'none'.

    'none' is an explicit per-edge no-interface marker — still a
    structural volume (polygon built, material assigned); only the
    interface planner skips it, exactly like 'str_volume'.
    """
    mode = safe_str(value).lower()
    return mode in ('str_volume', 'none') or mode.startswith('str_vol_')


def _volume_extension(value):
    """Extract EX direction from a volume-profile type, if present."""
    mode = safe_str(value).lower()
    for direction in ('both', 'top', 'bot', 'left', 'right'):
        if mode.endswith(f'ex{direction}'):
            return f'EX{direction}'
    return None


def _parse_volume_interface_mode(value):
    """Return (side, extended) for a str_2D volume-profile type."""
    mode = safe_str(value).lower()
    if mode in ('', 'str_volume', 'none'):
        return None, False
    if mode == 'str_vol_posinterface':
        return 'positive', False
    if mode == 'str_vol_neginterface':
        return 'negative', False
    if mode == 'str_vol_bothinterface':
        return 'both', False
    if mode.startswith('str_vol_pos_ex'):
        return 'positive', True
    if mode.startswith('str_vol_neg_ex'):
        return 'negative', True
    return None, False


def _classify_segment_orientation(x1, y1, x2, y2, tolerance=1e-6):
    """Classify a volume-profile segment using normalized SC4 coordinates.

    SC4 removes AutoCAD coordinate drift before writing the geometry to
    ``str_2D``.  Keep only a numerical-comparison tolerance here; it is not
    intended to correct geometry or redefine diagonal segments.
    """
    dx = abs(float(x2) - float(x1))
    dy = abs(float(y2) - float(y1))
    if dx <= tolerance and dy > tolerance:
        return 'vertical'
    if dy <= tolerance and dx > tolerance:
        return 'horizontal'
    return 'diagonal'


def _compute_extended_endpoints(x1, y1, x2, y2, extension, extended,
                                length=0.3):
    """Extend eligible horizontal/vertical endpoints by a fixed 0.3 m."""
    if not extended:
        return x1, y1, x2, y2
    orientation = _classify_segment_orientation(x1, y1, x2, y2)
    if orientation == 'vertical':
        low = min(y1, y2)
        high = max(y1, y2)
        if extension in ('EXtop', 'EXboth'):
            if y1 >= y2:
                y1 = high + length
            else:
                y2 = high + length
        if extension in ('EXbot', 'EXboth'):
            if y1 <= y2:
                y1 = low - length
            else:
                y2 = low - length
    elif orientation == 'horizontal':
        left = min(x1, x2)
        right = max(x1, x2)
        if extension in ('EXleft', 'EXboth'):
            if x1 <= x2:
                x1 = left - length
            else:
                x2 = left - length
        if extension in ('EXright', 'EXboth'):
            if x1 >= x2:
                x1 = right + length
            else:
                x2 = right + length
    return x1, y1, x2, y2


def _plan_volume_profile_interfaces(polygons, positive_index, negative_index):
    """Plan volume-profile interfaces and consume global sign counters.

    Each segment row in str_2D has its own H value. The interface mode is
    taken from the segment's type, not the polygon's type. A segment with
    'str_volume' or blank type receives no interface.
    """
    records = []
    ext_counter = 0
    for poly in polygons:
        if not _is_volume_profile_type(poly.get('type')):
            continue
        for segment_index, segment in enumerate(poly.get('segments', []), 1):
            seg_type = safe_str(segment.get('type')).lower()
            side, extended = _parse_volume_interface_mode(seg_type)
            if side is None:
                continue
            extension = _volume_extension(seg_type) or ''
            x1, y1 = segment['x1'], segment['y1']
            x2, y2 = segment['x2'], segment['y2']
            ex_x1, ex_y1, ex_x2, ex_y2 = _compute_extended_endpoints(
                x1, y1, x2, y2, extension, extended,
            )
            ext_counter += 1
            if side in ('positive', 'both'):
                positive_index += 1
                records.append({
                    'volume_name': poly['name'],
                    'segment_index': ext_counter,
                    'side': 'positive', 'label': f'PositiveInterface_{positive_index}',
                    'x1': ex_x1, 'y1': ex_y1, 'x2': ex_x2, 'y2': ex_y2,
                })
            if side in ('negative', 'both'):
                negative_index += 1
                records.append({
                    'volume_name': poly['name'],
                    'segment_index': ext_counter,
                    'side': 'negative', 'label': f'NegativeInterface_{negative_index}',
                    'x1': ex_x1, 'y1': ex_y1, 'x2': ex_x2, 'y2': ex_y2,
                })
    return records, positive_index, negative_index

def gen_vol_interface_code(records, positive_index=0, negative_index=0):
    """Generate volume-profile interface calls from planned segment records."""
    lines = []
    grouped = {}
    for record in records:
        grouped.setdefault(record['segment_index'], []).append(record)
    for idx, segment_records in grouped.items():
        record = segment_records[0]
        lines.append(
            f'vol_int_line_{idx} = g_i.line('
            f'({round_num(record["x1"])}, {round_num(record["y1"])}), '
            f'({round_num(record["x2"])}, {round_num(record["y2"])}))[-1]'
        )
        sides = set()
        for rec in segment_records:
            if rec['side'] == 'both':
                sides.update(('positive', 'negative'))
            else:
                sides.add(rec['side'])
        if 'positive' in sides:
            lines.append(f'vol_int_pos_{idx} = g_i.posinterface(vol_int_line_{idx})')
        if 'negative' in sides:
            lines.append(f'vol_int_neg_{idx} = g_i.neginterface(vol_int_line_{idx})')
    return lines



def _is_inheritable_blank(value):
    """Return True for empty inputs and warning text from a prior SC5 run."""
    if value is None:
        return True
    text = safe_str(value)
    return text in ('', 'None') or text.startswith((
        'MISSING ',
        'INVALID ',
        'MISMATCH ',
    ))


def apply_inheritance(data, layer_header, inherit_headers):
    """Group rows by DXF layer name (cad_layer), first occurrence = parent template.
    All elements on the same DXF layer share one material, regardless of geometry.
    Later rows on same layer inherit blank yellow cells from parent.
    Only the first element on a layer is the parent (green); all others are children.
    Returns list of parent row indices (for coloring green)."""
    parents = {}  # layer_name → parent row dict
    parent_rows = []  # indices of parent rows (for coloring green)

    for i, d in enumerate(data):
        layer_name = safe_str(d.get(layer_header), '').lower()
        if not layer_name:
            continue

        if layer_name not in parents:
            # First element on this DXF layer = parent (defines material)
            parents[layer_name] = d
            parent_rows.append(i)
        else:
            # Child: inherit blank values from parent on same layer
            parent = parents[layer_name]
            for h in inherit_headers:
                if _is_inheritable_blank(d.get(h)):
                    d[h] = parent.get(h)

    return parent_rows


def na_label(is_child):
    """Return N/A for parent rows (truly not applicable), None for children (leave blank)."""
    return None if is_child else 'N/A'


# ── Soil profile reader and material generator ──

# User-curated fixed colour palettes.  Layers are assigned top-to-bottom:
# index 0 = first layer, index 1 = second layer, etc.  If more layers than
# palette entries, colours cycle.  Silt keeps PLAXIS defaults.
SOIL_COLOUR_PALETTES = {
    'clay': [
        (243, 149, 165),  # pink
        (166, 242, 235),  # mint
        (104, 240, 25),   # bright green
        (250, 36, 25),    # bright red
    ],
    'sand': [
        (198, 182, 169),  # beige
        (221, 186, 14),   # golden yellow
        (142, 67, 69),    # dusty red-brown
        (101, 16, 11),    # maroon
    ],
    'linear-elastic': [
        (179, 195, 204),  # Concrete 1
        (105, 124, 134),  # Concrete 2
        (93, 108, 203),   # Concrete 3
    ],
}

# One shared palette for all PLAXIS embedded-beam materials: piles, nails,
# and grout anchors.  The caller supplies the global sequence index.
EMBEDDED_BEAM_COLOUR_PALETTE = [
    (199, 82, 143),   # Rose Pink
    (100, 200, 230),  # Cool Blue
    (230, 140, 60),   # Warm Orange
    (60, 160, 80),    # Forest Green
    (240, 220, 100),  # Clear Yellow
    (170, 140, 200),  # Lavender
]


def compute_embedded_beam_color(sequence_index=0):
    """Return the shared O(1) colour for a pile, nail, or grout anchor."""
    try:
        index = max(int(sequence_index), 0) % len(EMBEDDED_BEAM_COLOUR_PALETTE)
    except (TypeError, ValueError):
        index = 0
    return EMBEDDED_BEAM_COLOUR_PALETTE[index]


PLATE_COLOUR_PALETTE = [
    (0, 0, 255),      # Original Blue
    (0, 160, 0),      # Bold Green
    (128, 0, 128),    # Bold Purple
    (230, 115, 0),    # Bold Orange
    (220, 0, 0),      # Bold Red
]


def compute_plate_color(sequence_index=0):
    """Return the O(1) colour for a unique plate material, cycling as needed."""
    try:
        index = max(int(sequence_index), 0) % len(PLATE_COLOUR_PALETTE)
    except (TypeError, ValueError):
        index = 0
    return PLATE_COLOUR_PALETTE[index]


def _soil_colour_family(soil_type):
    """Return the visualization palette family for a soil type."""
    text = safe_str(soil_type, 'sand-HS').lower()
    if text == 'linear-elastic':
        return 'linear-elastic'
    return text.split('-', 1)[0]


def compute_soil_layer_color(soil_type, family_index=0, family_count=1,
                              layer_name=''):
    """Return a user-curated RGB colour for one str_2D soil material.

    Sand and clay use the fixed palette in top-to-bottom order.  Silt and
    unknown families return None so PLAXIS retains its default colour.
    ``family_count`` and ``layer_name`` are accepted for API compatibility;
    assignment is intentionally a fast O(1) palette lookup.
    """
    family = _soil_colour_family(soil_type)
    palette = SOIL_COLOUR_PALETTES.get(family)
    if not palette:
        return None
    try:
        index = max(int(family_index), 0) % len(palette)
    except (TypeError, ValueError):
        index = 0
    return palette[index]


def read_soil_profile(sheet, header_row=6, max_rows=10):
    """Read soil profile from str_2D rows 7-16 (header at row 6).
    Returns list of dicts, one per layer with a valid soil type.
    Uses column-insensitive header matching (same pattern as read_section).

    Optimized: block reads for header and data."""
    last_col = sheet.range(header_row, 28).end('right').column

    # Block read: header row (one COM call)
    header_vals = sheet.range((header_row, 2), (header_row, last_col)).value
    if not isinstance(header_vals, list):
        header_vals = [header_vals]
    headers = [safe_str(h) for h in header_vals]

    col_map = {}
    for c, h in enumerate(headers, 2):
        if h:
            col_map[h] = c

    # Block read: all data rows (one COM call)
    if max_rows < 1:
        return [], col_map
    data_start = header_row + 1
    data_end = header_row + max_rows
    block = sheet.range((data_start, 2), (data_end, last_col)).value
    if block is None:
        return [], col_map
    if not isinstance(block, list):
        block = [block]

    # Process block in-memory
    data = []
    type_col_idx = col_map.get('type', 2) - 2  # convert to 0-based block index
    soil_layers_idx = col_map.get('soil_layers', 3) - 2
    for row_vals in block:
        if not isinstance(row_vals, list):
            row_vals = [row_vals] * (last_col - 1)
        soil_type = row_vals[type_col_idx] if type_col_idx < len(row_vals) else None
        if soil_type is None or str(soil_type).strip() == '':
            continue
        if str(soil_type).strip() not in SOIL_TYPE_MAP:
            continue
        row_dict = {}
        for h, c in col_map.items():
            col_idx = c - 2
            if col_idx < len(row_vals):
                row_dict[h] = row_vals[col_idx]
            else:
                row_dict[h] = None
        data.append(row_dict)

    return data, col_map


def gen_soil_mat(layer_idx, d, is_3d=False, family_index=0, family_count=1):
    """Generate PLAXIS soil material code following sc3 pattern.
    Uses g_i.soilmat() Python API format.
    Returns list of code lines for one notebook cell.
    family_index/family_count control the material colour variation within
    the soil type family (e.g. 3 clay layers get light→dark shading)."""
    soil_name = safe_str(d.get('soil_layers'), f'Soil_{layer_idx}')
    name_formatted = "".join(e for e in soil_name.title() if e.isalnum())
    soil_type = safe_str(d.get('type'), 'sand-HS')

    model_name, drainage, is_hss, is_ssc, usda_class = SOIL_TYPE_MAP.get(
        soil_type, ('Hardening Soil', 'Drained', False, False, 'Sand')
    )

    # ── Linear Elastic: 5-param shortcut (non-porous) ──
    if soil_type == 'linear-elastic':
        # Parse E and nu from name, e.g. "Concrete_E30GPa_nu0.15".
        # PLAXIS stiffness input is kN/m² = kPa: 1 GPa = 1,000,000 kPa.
        e_match = re.search(r'E([\d.]+)GPa', soil_name, re.IGNORECASE)
        nu_match = re.search(r'(?:nu|v)([\d.]+)', soil_name, re.IGNORECASE)
        name_e_ref = (float(e_match.group(1)) * 1_000_000
                      if e_match else 0)
        name_nu = float(nu_match.group(1)) if nu_match else 0
        ERef = safe_num(d.get('E50_ref')) or name_e_ref or 30_000_000
        nu = safe_num(d.get('psi')) or name_nu or 0.15
        gammaUnsat = safe_num(d.get('gunsat')) or 25

        _clr = compute_soil_layer_color(soil_type, family_index, family_count,
                                        layer_name=soil_name)
        mat_params = [
            '"Identification"', fmt(soil_name),
            '"SoilModel"', fmt(model_name),
            '"DrainageType"', fmt(drainage),
            '"gammaUnsat"', round_num(gammaUnsat),
            '"ERef"', round_num(ERef),
            '"nu"', round_num(nu),
        ]
        params_str = ", ".join(str(p) for p in mat_params)
        result = [
            f'# Layer {layer_idx}: {soil_name} ({soil_type})',
            f'{name_formatted} = g_i.soilmat({params_str})',
        ]
        if _clr:
            result.append(f'{name_formatted}.setcolour({_clr[0]}, {_clr[1]}, {_clr[2]})')
        return result

    # ── HS / HSS / SSC models ──
    gammaUnsat = safe_num(d.get('gunsat'))
    gammaSat = safe_num(d.get('gsat'))
    E50Ref = safe_num(d.get('E50_ref'))
    EoedRef = safe_num(d.get('Eoed_ref'))
    EurRef = safe_num(d.get('Eur_ref'))
    PowerM = safe_num(d.get('m'))
    cRef = safe_num(d.get('c'))
    phi = safe_num(d.get('phi'))
    psi = safe_num(d.get('psi'))
    pref = safe_num(d.get('pref'))
    Rinter = safe_num(d.get('Rinter'))
    # Validate Rinter: must be 0.01 to 1, default to 1.0 if empty/invalid
    if Rinter <= 0 or Rinter > 1:
        Rinter = 1.0  # default: full interface strength
    # NOTE: clay_frac/silt_frac are kept raw here — a genuinely blank cell
    # means "omit the fraction line so PLAXIS falls back to the USDA class
    # default" (see _gw_frac_is_filled in the groundwater block).
    ClayFraction = d.get('clay_frac')
    SiltFraction = d.get('silt_frac')
    mode = _normalise_gw_mode(d.get('mode'))
    k_ver = safe_num(d.get('k_ver'))
    k_ratio = d.get('k_hor_ratio')
    if k_ratio is None or safe_str(k_ratio).strip() == '':
        k_ratio = 1.5
    else:
        k_ratio = safe_num(k_ratio, 1.5)

    # Build soilmat params
    mat_params = [
        '"Identification"', fmt(soil_name),
        '"SoilModel"', fmt(model_name),
        '"DrainageType"', fmt(drainage),
        '"gammaUnsat"', round_num(gammaUnsat),
        '"gammaSat"', round_num(gammaSat),
    ]

    # HS/HSS overconsolidation: each of OCR/POP/einit emits independently
    # iff its own cell is filled (same blank gate as the groundwater
    # fractions). Blank omits that pair; filled emits the input value.
    # SSC below gates eInit/OCR/POP the same way (blank omits, so no
    # raw-nan NameError); lambda/kappa/mu always emit; POP stays last
    # so a both-filled row keeps POP trailing.
    ocr_pep = []
    if _gw_frac_is_filled(d.get('OCR')):
        ocr_pep.append(('"OCR"', round_num(safe_num(d.get('OCR')))))
    if _gw_frac_is_filled(d.get('POP')):
        ocr_pep.append(('"POP"', round_num(safe_num(d.get('POP')))))
    if _gw_frac_is_filled(d.get('einit')):
        ocr_pep.append(('"eInit"', round_num(safe_num(d.get('einit')))))

    if is_hss:
        mat_params.extend([
            '"E50Ref"', round_num(E50Ref), '"EoedRef"', round_num(EoedRef), '"EurRef"', round_num(EurRef),
            '"PowerM"', round_num(PowerM), '"pRef"', round_num(pref),
            '"G0Ref"', round_num(safe_num(d.get('G0_ref'))),
            '"gamma07"', round_num(safe_num(d.get('g07'))),
        ])
        for _k, _v in ocr_pep:
            mat_params.extend([_k, _v])
    elif is_ssc:
        if _gw_frac_is_filled(d.get('einit')):
            mat_params.extend(['"eInit"', round_num(safe_num(d.get('einit')))])
        if _gw_frac_is_filled(d.get('OCR')):
            mat_params.extend(['"OCR"', round_num(safe_num(d.get('OCR')))])
        mat_params.extend([
            '"lambdaModified"', round_num(safe_zero(d.get('lambda'))),
            '"kappaModified"', round_num(safe_zero(d.get('kappa'))),
            '"muModified"', round_num(safe_zero(d.get('mu'))),
        ])
        if _gw_frac_is_filled(d.get('POP')):
            mat_params.extend(['"POP"', round_num(safe_num(d.get('POP')))])
    else:
        mat_params.extend([
            '"E50Ref"', round_num(E50Ref), '"EoedRef"', round_num(EoedRef), '"EurRef"', round_num(EurRef),
            '"PowerM"', round_num(PowerM), '"pRef"', round_num(pref),
        ])
        for _k, _v in ocr_pep:
            mat_params.extend([_k, _v])

    mat_params.extend([
        '"cRef"', round_num(cRef), '"phi"', round_num(phi), '"psi"', round_num(psi),
        '"InterfaceStrengthDetermination"', '"Manual"', '"Rinter"', round_num(Rinter),
    ])

    params_str = ", ".join(str(p) for p in mat_params)

    lines = [
        f'# Layer {layer_idx}: {soil_name} ({soil_type})',
        f'{name_formatted} = g_i.soilmat({params_str})',
    ]

    # ── Material colours (v0.7.8, palette-based) ──
    _clr = compute_soil_layer_color(soil_type, family_index, family_count)
    if _clr:
        lines.append(f'{name_formatted}.setcolour({_clr[0]}, {_clr[1]}, {_clr[2]})')

    # Groundwater per mode — all modes are USDA. The class string comes from
    # the SOIL_TYPE_MAP 5th element only (capitalized Clay/Sand/Silt — never
    # sliced from the soil-type prefix; PLAXIS .set() is case-sensitive).
    # A genuinely blank fraction cell omits that fraction line so PLAXIS
    # falls back to the class default; an explicit 0 still emits.
    if mode == 'automatic':
        lines.append(f'{name_formatted}.GroundwaterClassificationType.set("USDA")')
        lines.append(f'{name_formatted}.GroundwaterSoilClassUSDA.set("{usda_class}")')
        lines.append(f'{name_formatted}.GwUseDefaults.set(True)')
        lines.append(f'{name_formatted}.GwDefaultsMethod.set("From grain size distribution")')
    elif mode == 'by_grain_size':
        lines.append(f'{name_formatted}.GroundwaterClassificationType.set("USDA")')
        lines.append(f'{name_formatted}.GroundwaterSoilClassUSDA.set("{usda_class}")')
        if _gw_frac_is_filled(ClayFraction):
            lines.append(f'{name_formatted}.ClayFraction.set({ClayFraction})')
        if _gw_frac_is_filled(SiltFraction):
            lines.append(f'{name_formatted}.SiltFraction.set({SiltFraction})')
        lines.append(f'{name_formatted}.GwUseDefaults.set(True)')
        lines.append(f'{name_formatted}.GwDefaultsMethod.set("From grain size distribution")')
    elif mode == 'manual':
        lines.append(f'{name_formatted}.GroundwaterClassificationType.set("USDA")')
        lines.append(f'{name_formatted}.GroundwaterSoilClassUSDA.set("{usda_class}")')
        if _gw_frac_is_filled(ClayFraction):
            lines.append(f'{name_formatted}.ClayFraction.set({ClayFraction})')
        if _gw_frac_is_filled(SiltFraction):
            lines.append(f'{name_formatted}.SiltFraction.set({SiltFraction})')
        lines.append(f'{name_formatted}.GwUseDefaults.set(False)')
        lines.append(f'{name_formatted}.PermHorizontalPrimary.set({k_ver * k_ratio})')
        if is_3d:
            lines.append(f'{name_formatted}.PermHorizontalSecondary.set({k_ver * k_ratio})')
        lines.append(f'{name_formatted}.PermVertical.set({k_ver})')

    return lines


def gen_soil_profile_code(soil_data, is_3d=False):
    """Generate notebook cells for soil material definitions.
    Returns list of notebook cells."""
    cells = []

    if not soil_data:
        return cells

    cells.append(make_cell(SOIL_MAT_TAG, cell_type='markdown'))

    # Count layers by palette family so each family receives a full light→dark
    # ramp while retaining the original family identity.
    family_counts = {}
    for d in soil_data:
        family = _soil_colour_family(d.get('type'))
        family_counts[family] = family_counts.get(family, 0) + 1
    family_indices = {}

    for i, d in enumerate(soil_data, 1):
        family = _soil_colour_family(d.get('type'))
        family_index = family_indices.get(family, 0)
        family_indices[family] = family_index + 1
        lines = gen_soil_mat(
            i, d, is_3d=is_3d,
            family_index=family_index,
            family_count=family_counts[family],
        )
        cells.append(make_cell(lines))

    return cells


# ── Soil polygon functions (v0.6) ──

REFINE_BRACKET_RE = re.compile(r'\[\s*refine\b([^\]]*)\]', re.IGNORECASE)
REFINE_FACTOR_MIN = 0.05
REFINE_FACTOR_MAX = 1.0


def _parse_refine_bracket(raw, row_label=''):
    """Split a str_2D column-C value into (clean_text, factor_or_None).

    Gesture: ``volume_profile[refine 0.2]`` — case-insensitive, factor in
    [0.05, 1.0].  Any bracket carrying the ``refine`` keyword with a
    non-numeric, zero, negative, or out-of-range value raises ValueError
    LOUDLY (never a silent skip).  Brackets without the keyword pass
    through untouched.
    """
    text = safe_str(raw)
    hits = REFINE_BRACKET_RE.findall(text)
    if not hits:
        return text, None
    factors = []
    for num_raw in hits:
        num = safe_str(num_raw)
        try:
            factor = float(num)
        except (TypeError, ValueError):
            raise ValueError(
                f"sc5: bad [refine ...] tag {raw!r} at {row_label}: "
                f"value {num!r} is not a number")
        if not (REFINE_FACTOR_MIN <= factor <= REFINE_FACTOR_MAX):
            raise ValueError(
                f"sc5: bad [refine ...] tag {raw!r} at {row_label}: "
                f"factor {factor:g} outside [{REFINE_FACTOR_MIN:g}, {REFINE_FACTOR_MAX:g}]")
        factors.append(factor)
    if len(set(factors)) > 1:
        raise ValueError(
            f"sc5: conflicting [refine] tags {raw!r} at {row_label}: "
            f"{' vs '.join(f'{f:g}' for f in dict.fromkeys(factors))} — keep one value")
    clean = REFINE_BRACKET_RE.sub('', text).strip()
    clean = re.sub(r'\s{2,}', ' ', clean)
    return clean, factors[0]


def read_soil_polygons(sheet, start_row=235):
    """Read soil polygon data from the current str_2D table (header row 234,
    data starting at row 235). Groups consecutive rows with same
    spoly_name as one polygon. Returns list of dicts, one per polygon.
    Each dict has: name, layer, type, vertices [(x,y),...], y_centroid,
    material (column J, used at creation) and material_after (column K,
    used at staged construction setmaterial for soil_replacement).
    Max rows from cell I233, default 100.

    Optimized: single block read for all polygon data."""
    from collections import OrderedDict

    # Read max rows from I233; default 100
    _raw_max = sheet.range(233, 9).value  # I233
    try:
        max_rows = max(50, int(_raw_max))
    except (TypeError, ValueError):
        max_rows = 100

    # Block read: all columns B-K for all data rows (one COM call)
    block = sheet.range((start_row, 2), (start_row + max_rows - 1, 11)).value
    if block is None:
        return []
    if not isinstance(block, list):
        block = [block]

    # Process block data in-memory (no more COM calls)
    raw = []
    for _idx, row_vals in enumerate(block):
        _excel_row = start_row + _idx
        _row_label = f'str_2D!C{_excel_row}'
        if not isinstance(row_vals, list):
            row_vals = [row_vals] * 10  # B..K = 10 columns
        # Column-C [refine f] gesture: parse LOUD, then strip BEFORE any
        # layer logic.  Promotion to the whole polygon_no group happens
        # after grouping below.  Parser-only: no preservation (an SC4
        # reload wipes column C; re-press SC5 silently drops refinement).
        layer_raw = safe_str(row_vals[1])   # C: cad_layer (index 1)
        layer, _ref = _parse_refine_bracket(layer_raw, row_label=_row_label)
        name = safe_str(row_vals[0])    # B: spoly_name (index 0)
        if not name:
            if _ref is not None:
                raise ValueError(
                    f"sc5: [refine {_ref:g}] at {_row_label} has no spoly_name — "
                    f"tag a row belonging to a polygon")
            continue
        x1 = row_vals[2]                # D: x1 (index 2)
        y1 = row_vals[3]                # E: y1 (index 3)
        x2 = row_vals[4]                # F: x2 (index 4)
        y2 = row_vals[5]                # G: y2 (index 5)
        spoly_type = safe_str(row_vals[6])  # H: spoly_type (index 6)
        p_no = safe_num(row_vals[7], 0)     # I: polygon_no (index 7)
        if _ref is not None and int(p_no) == 0:
            raise ValueError(
                f"sc5: [refine {_ref:g}] at {_row_label} has no polygon_no — "
                f"tag a row belonging to a polygon")
        material = safe_str(row_vals[8])    # J: material_before (index 8)
        material_after = safe_str(row_vals[9])  # K: material_after (index 9)

        if x1 is None or y1 is None:
            if _ref is not None:
                raise ValueError(
                    f"sc5: [refine {_ref:g}] at {_row_label} has no geometry — "
                    f"tag a row with x1/y1")
            continue
        raw.append({
            'name': name, 'layer': layer, 'type': spoly_type,
            'polygon_no': int(p_no), 'material': material,
            'material_after': material_after,
            'x1': safe_num(x1), 'y1': safe_num(y1),
            'x2': safe_num(x2), 'y2': safe_num(y2),
            'refine_factor': _ref, '_row': _excel_row,
        })

    # Group by spoly_name (consecutive rows with same name = one polygon)
    polygons = OrderedDict()
    for r in raw:
        name = r['name']
        if name not in polygons:
            polygons[name] = {
                'name': name, 'layer': r['layer'], 'type': r['type'],
                'vertices': [], 'segments': [], 'polygon_no': r['polygon_no'],
                'material': r['material'], 'material_after': r['material_after'],
                'refine_rows': [],
            }
        polygons[name]['segments'].append({
            'x1': r['x1'], 'y1': r['y1'], 'x2': r['x2'], 'y2': r['y2'],
            'type': r['type'],
        })
        polygons[name]['vertices'].append((r['x1'], r['y1']))
        # Add end point of last segment as final vertex
        polygons[name]['_last_end'] = (r['x2'], r['y2'])
        if r.get('refine_factor') is not None:
            polygons[name]['refine_rows'].append((r['_row'], r['refine_factor']))
        # Update materials (last non-empty wins)
        if r['material']:
            polygons[name]['material'] = r['material']
        if r['material_after']:
            polygons[name]['material_after'] = r['material_after']

    # ── [refine f] promotion (column-C gesture → whole polygon_no group) ──
    # Any single tagged row promotes its polygon_no group.  Identical
    # values on sibling rows agree silently; differing values are a
    # fail-fast error (never guess which one the user meant).
    refine_by_pno = {}
    for name, poly in polygons.items():
        p_no = poly.get('polygon_no', 0)
        for _row, _f in poly.get('refine_rows', []):
            if p_no in refine_by_pno:
                if abs(refine_by_pno[p_no][0] - _f) > 1e-9:
                    raise ValueError(
                        f"sc5: conflicting [refine] on Polygon {p_no}: "
                        f"{refine_by_pno[p_no][0]:g} (str_2D!C{refine_by_pno[p_no][1]}) "
                        f"vs {_f:g} (str_2D!C{_row}) — keep one value")
            else:
                refine_by_pno[p_no] = (_f, _row)
    for name, poly in polygons.items():
        p_no = poly.get('polygon_no', 0)
        poly['refine_factor'] = refine_by_pno[p_no][0] if p_no in refine_by_pno else None

    # Finalize: add end vertex and compute Y-centroid
    result = []
    for name, poly in polygons.items():
        if poly['_last_end'] != poly['vertices'][0]:
            poly['vertices'].append(poly['_last_end'])
        del poly['_last_end']
        verts = poly['vertices']
        poly['y_centroid'] = sum(v[1] for v in verts) / len(verts) if verts else 0
        result.append(poly)

    return result


def gen_soil_polygon_code(polygons):
    """Generate PLAXIS code cells for soil polygons.
    Sorted by polygon_no (from sc4) — single source of truth.
    soil (cut/fill) first, str_volume appended after.
    Returns list of notebook cells."""
    cells = []

    # Split into soil (cut/fill/replacement) and all volume-profile modes.
    soil = [p for p in polygons if not _is_volume_profile_type(p.get('type'))]
    volumes = [p for p in polygons if _is_volume_profile_type(p.get('type'))]
    soil_sorted = sorted(soil, key=lambda p: p.get('polygon_no', 0))
    vol_sorted = sorted(volumes, key=lambda p: p.get('polygon_no', 0))
    polygons_by_no = soil_sorted + vol_sorted

    # ── Polygon geometry ──
    cells.append(make_cell('# ── Soil Polygons ──'))
    for i, poly in enumerate(polygons_by_no, 1):
        verts = poly['vertices']
        name = poly['name']
        p_no = poly.get('polygon_no', 0)
        vert_str = ', '.join(f'{round_num(v[0])}, {round_num(v[1])}' for v in verts)
        cells.append(make_cell(
            f'# Polygon {p_no}: {name} ({len(verts)} vertices, Y-centroid={poly.get("y_centroid", 0):.3f})\n'
            f'polygon_{p_no} = g_i.polygon({vert_str})'
        ))

    # ── Material assignments (using Polygons[polygon_no - 1]) ──
    assign_lines = []
    for poly in polygons_by_no:
        material = poly.get('material', '')
        p_no = poly.get('polygon_no', 0)
        if material and p_no:
            formatted = "".join(e for e in material.title() if e.isalnum())
            assign_lines.append(f'g_i.Polygons[{p_no - 1}].Soil.Material = {formatted}')

    if assign_lines:
        cells.append(make_cell('# ── Soil Material Assignments ──'))
        cells.append(make_cell('\n'.join(assign_lines)))

    # v0.6.18: Warn if material_after (column K) is blank for types that
    # need it during staged construction (soil_replacement, str_volume).
    warn_lines = []
    for poly in polygons_by_no:
        ptype = poly.get('type', '')
        if ptype in ('soil_replacement',) or _is_volume_profile_type(ptype):
            mat_after = poly.get('material_after', '')
            if not mat_after:
                pname = poly.get('name', f'Polygon_{poly.get("polygon_no", "?")}')
                warn_lines.append(
                    f'print("WARNING {pname}: {ptype} has no material_after '
                    f'(column K) — setmaterial() will be skipped in staged '
                    f'construction, PLAXIS will use background soil profile")'
                )
    if warn_lines:
        cells.append(make_cell('# ── Material-after warnings ──'))
        cells.append(make_cell('\n'.join(warn_lines)))

    return cells


# ── Drain (PVD) functions ──

def read_drains(sheet, start_row=235):
    """Read PVD drains from the current str_2D table (header row 234,
    data starting at row 235).
    Returns list of dicts, one per drain. Max rows from R233, default 125.

    Optimized: single block read for all drain data."""
    # Read max rows from R233; default 125
    _raw_max = sheet.range(233, 18).value  # R233
    try:
        max_rows = max(1, int(_raw_max))
    except (TypeError, ValueError):
        max_rows = 100

    # Block read: all columns M-S for all data rows (one COM call)
    block = sheet.range((start_row, 13), (start_row + max_rows - 1, 19)).value
    if block is None:
        return []
    if not isinstance(block, list):
        block = [block]

    # Process block data in-memory
    drains = []
    for i, row_vals in enumerate(block):
        if not isinstance(row_vals, list):
            row_vals = [row_vals] * 7  # M..S = 7 columns
        name = safe_str(row_vals[0])   # M: drains_name (index 0)
        if not name:
            break
        drains.append({
            'name': name,
            'cad_layer': safe_str(row_vals[1]),      # N: cad_layer (index 1)
            'x1': safe_num(row_vals[2]),              # O: x1 (index 2)
            'y1': safe_num(row_vals[3]),              # P: y1 (index 3)
            'x2': safe_num(row_vals[4]),              # Q: x2 (index 4)
            'y2': safe_num(row_vals[5]),              # R: y2 (index 5)
            'drain_no': int(row_vals[6] or i + 1),   # S: drain_no (index 6)
        })
    return drains


def gen_drain_code(drains):
    """Generate PLAXIS code cells for PVD drains.
    Returns list of notebook cells."""
    cells = []

    if not drains:
        return cells

    cells.append(make_cell('# ── PVD Drains ──'))
    for i, d in enumerate(drains, 1):
        x1 = round_num(d['x1'])
        y1 = round_num(d['y1'])
        x2 = round_num(d['x2'])
        y2 = round_num(d['y2'])
        cells.append(make_cell(
            f'drain_{i} = g_i.drain(({x1}, {y1}), ({x2}, {y2}))[-1]'
        ))

    return cells


# ── Water level functions ──

WATERLINE_DATA_START = 235

# ── Mesh and Output curve-point inputs (v0.7.10) ──
# str_2D!K2 stores the already-divided relative mesh factor (e.g. 0.05).
MESH_DENSITY_CELL = 'K2'
MESH_GENERATION_CELL = 'K3'
MESH_DENSITY_DEFAULT = 0.05
MESH_SECTION_TAG = '# ── Mesh generation (from sc5, v0.7.10) ──'
CURVE_SECTION_TAG = '# ── Output curve points (from sc5, v0.7.10) ──'
CURVE_OUTPUT_SHEET = 'output'
CURVE_OUTPUT_DATA_START = 9
CURVE_OUTPUT_DATA_END = 108
CURVE_OUTPUT_COL_START = 2  # B
CURVE_OUTPUT_COL_END = 6    # F
WATERLINE_MAX_CELL_Z233 = 'Z233'

# ── Output forces table (v0.7.11) ──
# The table is intentionally separate from the SC4-owned curve-point block.
OUTPUT_FORCES_SHEET = 'output'
OUTPUT_FORCES_HEADER_ROW = 8
OUTPUT_FORCES_DATA_START = 9
OUTPUT_FORCES_COL_START = 11  # K: source element, L: PLAXIS label
OUTPUT_FORCES_PHASE_START = 13  # M
OUTPUT_FORCES_SCAN_ROWS = 1000
OUTPUT_FORCES_SCAN_COLS = 100


def _normalize_output_header(value):
    """Normalize phase names for matching while preserving displayed text."""
    return re.sub(r'[^a-z0-9]+', ' ', safe_str(value).lower()).strip()


def _as_row(value):
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


def _phase_headers_from_assign_sequence(seq_sheet):
    """Read contiguous phase names from assign_sequence row 6, starting at L."""
    values = _as_row(seq_sheet.range((6, 12)).expand('right').value)
    phases = []
    for value in values:
        name = safe_str(value)
        if not name:
            break
        phases.append(name)
    return phases


def _is_output_structural_label(label):
    """Return True for the initial SC8 structural scope, never soil/interfaces."""
    label = safe_str(label).lower()
    return label.startswith(('plate_', 'embeddedbeam_', 'nodetonodeanchor_', 'fixedendanchor_',
                             'geogrid_'))


def _is_manual_output_element(element):
    """Return True for user-owned manual centerline rows.

    Column K carries the mode (case-insensitive): 'manual' (MONO overview
    + per-part sheets) or 'manual_mono' (MONO sheet only). Both spellings
    are preserved verbatim so SC8 can tell them apart. Column L carries
    the bare soil material name (verbatim copy of a str_2D soil_layers
    name, e.g. 'stem-Concrete_E25GPa_nu0.2'). Manual rows are never
    generated from assign_sequence — only preserved by the writer.
    """
    return safe_str(element).lower() in ('manual', 'manual_mono')


def _extract_manual_output_rows(pairs):
    """Dedupe (element, label) pairs preserving table order; skip blank labels.

    Pure helper (no Excel) so the preservation rule is unit-testable.
    Matching is case-insensitive via _normalize_output_header; the first
    spelling wins.
    """
    result = []
    seen = set()
    for element, label in pairs:
        if not _is_manual_output_element(element):
            continue
        if not safe_str(label):
            continue
        key = _normalize_output_header(label)
        if key in seen:
            continue
        seen.add(key)
        result.append((element, label))
    return result


def _missing_manual_materials(manual_rows, soil_names):
    """Return manual labels with no match in the soil profile names.

    Both sides normalize to bare alphanumeric lowercase (same convention
    SC8 uses to match Excel Identification against Output material Name),
    so casing/separator differences never false-positive. Pure helper.
    """
    known = {re.sub(r'[^a-z0-9]+', '', safe_str(n).lower()) for n in soil_names}
    return [label for _, label in manual_rows
            if re.sub(r'[^a-z0-9]+', '', safe_str(label).lower()) not in known]


def _snapshot_output_forces_arguments(output_sheet):
    """Snapshot existing argument cells keyed by (PLAXIS label, phase name).

    Reading is deliberately bounded to the existing table footprint.  Blank
    argument cells remain blank, and duplicate labels keep the first value.
    """
    block = output_sheet.range(
        (OUTPUT_FORCES_HEADER_ROW, OUTPUT_FORCES_COL_START),
        (OUTPUT_FORCES_SCAN_ROWS,
         OUTPUT_FORCES_COL_START + OUTPUT_FORCES_SCAN_COLS - 1),
    ).value
    if block is None:
        return {}, 0, 0
    if not isinstance(block, list):
        block = [_as_row(block)]
    rows = [row if isinstance(row, list) else [row] for row in block]
    header = rows[0] if rows else []
    phase_names = [safe_str(v) for v in header[2:]]
    last_phase = 0
    for index, name in enumerate(phase_names):
        if name:
            last_phase = index + 1
    phase_names = phase_names[:last_phase]
    last_row = 0
    snapshot = {}
    for row_offset, row in enumerate(rows[1:], 1):
        element = safe_str(row[0] if len(row) > 0 else None)
        label = safe_str(row[1] if len(row) > 1 else None)
        if not element and not label:
            continue
        last_row = row_offset
        if not label:
            continue
        for phase_index, phase in enumerate(phase_names):
            value = row[2 + phase_index] if 2 + phase_index < len(row) else None
            if safe_str(value):
                snapshot.setdefault(
                    (_normalize_output_header(label),
                     _normalize_output_header(phase)), value
                )
    return snapshot, last_row, len(phase_names)


def _snapshot_manual_output_rows(output_sheet):
    """Snapshot user-owned manual centerline rows as (element, label) pairs.

    Reads the same output!K8 footprint as _snapshot_output_forces_arguments.
    Only rows with a manual mode in column K ('manual' | 'manual_mono',
    case-insensitive) and a non-blank column L are returned, in table
    order. Phase arguments for these rows are already covered by the
    (label, phase) snapshot — this only preserves row identity so the
    writer can re-append them after the auto rows.
    """
    block = output_sheet.range(
        (OUTPUT_FORCES_HEADER_ROW, OUTPUT_FORCES_COL_START),
        (OUTPUT_FORCES_SCAN_ROWS,
         OUTPUT_FORCES_COL_START + OUTPUT_FORCES_SCAN_COLS - 1),
    ).value
    if block is None:
        return []
    if not isinstance(block, list):
        block = [_as_row(block)]
    rows = [row if isinstance(row, list) else [row] for row in block]
    pairs = []
    for row in rows[1:]:
        element = safe_str(row[0] if len(row) > 0 else None)
        label = safe_str(row[1] if len(row) > 1 else None)
        if not element and not label:
            continue
        pairs.append((element, label))
    return _extract_manual_output_rows(pairs)


def _output_structural_rows(seq_sheet):
    """Read structural rows from assign_sequence without soil polygons.

    v0.7.20: tunnel_volume liner rows (col C etype) are included in
    assign_sequence order so each liner sits exactly between its own
    tunnel's plate rows. The Plate_* label gate is untouched; SC8 skips
    the liner proxy labels until the centerline extraction lands.
    """
    block = seq_sheet.range((7, 2), (OUTPUT_FORCES_SCAN_ROWS, 4)).value
    if block is None:
        return []
    if not isinstance(block, list):
        block = [_as_row(block)]
    result = []
    for row in block:
        row = row if isinstance(row, list) else [row]
        name = safe_str(row[0] if len(row) > 0 else None)
        etype = safe_str(row[1] if len(row) > 1 else None).lower()
        label = safe_str(row[2] if len(row) > 2 else None)
        if not name and not label:
            continue
        if _is_output_structural_label(label) or etype == 'tunnel_volume':
            result.append((name, label))
    return result


def _write_output_forces_table(wb, soil_names=None):
    """Refresh output!K8 onward, preserving arguments by label and phase.

    Phase columns come from assign_sequence row 6; no maximum phase count is
    assumed.  Only K:L and the generated phase columns are owned here.

    Manual centerline rows (column K 'manual' | 'manual_mono', column L
    = bare soil material name) are user-owned: they are snapshotted before
    the clear and re-appended after the auto rows with their element
    spelling and phase arguments intact.
    Tunnel liner rows (etype 'tunnel_volume') and the Plate_* label gate
    are untouched by this path.
    """
    seq_sheet = wb.sheets['assign_sequence']
    output_sheet = wb.sheets[OUTPUT_FORCES_SHEET]
    previous, old_rows, old_phase_count = _snapshot_output_forces_arguments(output_sheet)
    manual_rows = _snapshot_manual_output_rows(output_sheet)
    phases = _phase_headers_from_assign_sequence(seq_sheet)
    structural_rows = _output_structural_rows(seq_sheet)

    if soil_names is None:
        try:
            soil_names = [safe_str(d.get('soil_layers'))
                          for d in read_soil_profile(
                              wb.sheets['str_2D'], header_row=6, max_rows=10)[0]]
        except Exception:
            soil_names = []
    missing = _missing_manual_materials(manual_rows, soil_names)
    for label in missing:
        print(f'sc5: WARNING manual output row "{label}" matches no '
              f'str_2D soil_layers name — kept, but SC8 will skip it')

    # Clear exactly the old table footprint, including removed rows/phase cols.
    old_width = max(2 + old_phase_count, 2)
    old_height = max(old_rows + 1, 1)
    output_sheet.range(
        (OUTPUT_FORCES_HEADER_ROW, OUTPUT_FORCES_COL_START),
        (OUTPUT_FORCES_HEADER_ROW + old_height - 1,
         OUTPUT_FORCES_COL_START + old_width - 1),
    ).clear_contents()

    headers = ['element', 'PLAXIS Label'] + phases
    output_sheet.range(
        (OUTPUT_FORCES_HEADER_ROW, OUTPUT_FORCES_COL_START),
        (OUTPUT_FORCES_HEADER_ROW,
         OUTPUT_FORCES_COL_START + len(headers) - 1),
    ).value = [headers]
    all_rows = list(structural_rows) + [
        (element, label) for element, label in manual_rows
        if (element, label) not in structural_rows
    ]
    if all_rows:
        data = []
        for element, label in all_rows:
            values = [element, label]
            for phase in phases:
                values.append(previous.get(
                    (_normalize_output_header(label),
                     _normalize_output_header(phase)),
                    None,
                ))
            data.append(values)
        output_sheet.range(
            (OUTPUT_FORCES_DATA_START, OUTPUT_FORCES_COL_START),
            (OUTPUT_FORCES_DATA_START + len(data) - 1,
             OUTPUT_FORCES_COL_START + len(headers) - 1),
        ).value = data
    print(
        f"sc5: output forces table updated ({len(structural_rows)} structural "
        f"+ {len(all_rows) - len(structural_rows)} manual elements, "
        f"{len(phases)} phases)"
    )
    return all_rows, phases


def read_waterlines(sheet, start_row=WATERLINE_DATA_START):
    """Read waterline segments from str_2D U:AA (header row 234, data 235+).
    Returns list of dicts grouped by waterline_no, each with ordered vertices.

    Optimized: single block read for all waterline data columns U-AA."""
    _raw_max = sheet.range(WATERLINE_MAX_CELL_Z233).value
    try:
        max_rows = max(1, int(_raw_max))
    except (TypeError, ValueError):
        max_rows = 125

    block = sheet.range((start_row, 21), (start_row + max_rows - 1, 27)).value
    if block is None:
        return []
    if not isinstance(block, list):
        block = [block]

    segments = []  # list of dicts
    for row_vals in block:
        if not isinstance(row_vals, list):
            row_vals = [row_vals] * 7
        name = safe_str(row_vals[0])
        if not name:
            break
        cad_layer = safe_str(row_vals[1])
        # Skip groundwater-flow boundary rows (parsed separately by read_gwflowbcs)
        if cad_layer.startswith('waterboundary'):
            continue
        segments.append({
            'waterline_name': name,
            'cad_layer': cad_layer,
            'x1': safe_num(row_vals[2]),
            'y1': safe_num(row_vals[3]),
            'x2': safe_num(row_vals[4]),
            'y2': safe_num(row_vals[5]),
            'waterline_no': int(row_vals[6] or 0),
        })

    # Group by waterline_no, preserving segment order from SC4
    wl_groups = {}
    for seg in segments:
        wl_no = seg['waterline_no']
        if wl_no not in wl_groups:
            wl_groups[wl_no] = {
                'waterline_name': seg['waterline_name'],
                'cad_layer': seg['cad_layer'],
                'vertices': [],
            }
        wl_groups[wl_no]['vertices'].append((seg['x1'], seg['y1']))
    # Add final endpoint for each chain
    for wl_no, grp in wl_groups.items():
        last_seg = [s for s in segments if s['waterline_no'] == wl_no][-1]
        grp['vertices'].append((last_seg['x2'], last_seg['y2']))

    waterlines = []
    for wl_no in sorted(wl_groups.keys()):
        grp = wl_groups[wl_no]
        waterlines.append({
            'name': grp['waterline_name'],
            'cad_layer': grp['cad_layer'],
            'vertices': grp['vertices'],
            'waterline_no': wl_no,
        })
    return waterlines


def gen_waterlevel_code(waterlines):
    """Generate PLAXIS code cells for water level creation.
    Returns list of notebook cells.

    Creates water levels in gotoflow(), then returns to gotostages().
    Variable names match the PLAXIS name: UserWaterLevel_N."""
    cells = []
    if not waterlines:
        return cells

    cells.append(make_cell(WATERLEVEL_TAG, cell_type='markdown'))
    cells.append(make_cell('g_i.gotoflow()'))
    for wl in waterlines:
        verts = wl['vertices']
        coords = ', '.join(
            f'({round_num(x)}, {round_num(y)})' for x, y in verts
        )
        var_name = f"UserWaterLevel_{wl['waterline_no']}"
        cells.append(make_cell([
            f'{var_name} = g_i.waterlevel(\n',
            f'    {coords},\n',
            f')\n',
        ]))
    cells.append(make_cell('g_i.gotostages()'))
    return cells


# ── GWFlowBC functions ──

def _parse_gwflowbc_behaviour(cad_layer):
    """Derive PLAXIS Behaviour from the DXF cad_layer name."""
    key = safe_str(cad_layer).lower()
    if 'head' in key:
        return 'Head'
    return 'Closed'


def _parse_gwflowbc_href(behaviour, x1, y1, x2, y2):
    """Derive Href from boundary elevation for Head; None for Closed."""
    if behaviour != 'Head':
        return None
    return float(y1)


def read_gwflowbcs(sheet, start_row=WATERLINE_DATA_START):
    """Read GWFlowBC rows from str_2D AB column (header row 234, data 235+).

    A row is a GWFlowBC if AB (column 28) is non-blank.  Behaviour is derived
    from cad_layer; Href equals the boundary Y elevation (y1).
    Returns list of dicts."""
    _raw_max = sheet.range(WATERLINE_MAX_CELL_Z233).value
    try:
        max_rows = max(1, int(_raw_max))
    except (TypeError, ValueError):
        max_rows = 125

    block = sheet.range((start_row, 21), (start_row + max_rows - 1, 28)).value
    if block is None:
        return []
    if not isinstance(block, list):
        block = [block]

    gwflowbcs = []
    for row_vals in block:
        if not isinstance(row_vals, list):
            row_vals = [row_vals] * 8
        gw_no_raw = row_vals[7] if len(row_vals) > 7 else None
        if gw_no_raw is None:
            continue
        try:
            gw_no = int(gw_no_raw)
        except (TypeError, ValueError):
            continue
        if gw_no < 1:
            continue
        cad_layer = safe_str(row_vals[1])
        if not cad_layer.startswith('waterboundary'):
            continue
        x1 = safe_num(row_vals[2])
        y1 = safe_num(row_vals[3])
        x2 = safe_num(row_vals[4])
        y2 = safe_num(row_vals[5])
        behaviour = _parse_gwflowbc_behaviour(cad_layer)
        href = _parse_gwflowbc_href(behaviour, x1, y1, x2, y2)
        gwflowbcs.append({
            'name': safe_str(row_vals[0]),
            'cad_layer': cad_layer,
            'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
            'gwflowbc_no': gw_no,
            'behaviour': behaviour,
            'href': href,
        })

    return gwflowbcs


def gen_gwflowbc_code(gwflowbcs):
    """Generate PLAXIS code cells for groundwater-flow boundary conditions.

    Creates GWFlowBC objects in structural mode (not gotoflow).
    Behaviour is set inline; Href is set separately for Head boundaries."""
    cells = []
    if not gwflowbcs:
        return cells

    cells.append(make_cell(GWFLOW_BC_TAG, cell_type='markdown'))
    lines = []
    for gw in gwflowbcs:
        x1, y1, x2, y2 = gw['x1'], gw['y1'], gw['x2'], gw['y2']
        behaviour = gw['behaviour']
        var_name = f"GWFlowBC_{gw['gwflowbc_no']}"
        lines.append(
            f'{var_name} = g_i.gwfbc(({round_num(x1)}, {round_num(y1)}), '
            f'({round_num(x2)}, {round_num(y2)}), "Behaviour", "{behaviour}")[-1]\n'
        )
        if behaviour == 'Head' and gw.get('href') is not None:
            lines.append(f'{var_name}.Href = {gw["href"]}\n')
    cells.append(make_cell(lines))
    return cells


def read_mesh_generation(sheet):
    """Return whether mesh and curve-point generation is included.

    ``str_2D!K3`` uses ``included``/``excluded`` wording.
    Blank or unrecognised values preserve the legacy behaviour and
    include generation.
    """
    try:
        raw = sheet.range(MESH_GENERATION_CELL).value
    except Exception:
        return True
    value = str(raw).strip().lower() if raw is not None else ''
    if value in ('skipped', 'excluded'):
        return False
    return True


def read_mesh_density(sheet):
    """Read the relative mesh factor from str_2D!K2.

    The user stores the value already divided by 100 (e.g. 0.05).
    Falls back to MESH_DENSITY_DEFAULT when blank or non-numeric."""
    try:
        raw = sheet.range(MESH_DENSITY_CELL).value
    except Exception:
        return MESH_DENSITY_DEFAULT
    try:
        value = float(raw)
    except (TypeError, ValueError):
        return MESH_DENSITY_DEFAULT
    return value if value > 0 else MESH_DENSITY_DEFAULT


def gen_mesh_code(mesh_density):
    """Generate the mesh-mode notebook cells.

    Kept separate from curve points so each step can be probed
    independently in the PLAXIS console."""
    return [
        MESH_SECTION_TAG,
        'g_i.gotomesh()',
        f'g_i.mesh({mesh_density})',
    ]


REFINE_SECTION_TAG = '# ── Local mesh refinement (from sc5: str_2D col-C [refine f]) ──'


def gen_refine_code(refinements):
    """Generate per-polygon refinement cells for the mesh block.

    Emitted BETWEEN gotomesh() and mesh() so the factors bite at the next
    Generate.  Probe-proven route (dead_plan/probe_polygon8_refine.ipynb):
    tag census over mesh-mode g_i.Polygons, single-level
    setproperties('CoarsenessFactor', f) — the _set X.CoarsenessFactor
    equivalent.  Nested attribute writes are blind to the mesher.

    Dual detection per polygon (one pass, startswith-or-substring): buried
    fragments match the substring arm (``BoreholePolygon_1_Polygon_9_…``)
    while soil-less-surface standalones match the startswith arm
    (``Polygon_9_2`` — no leading underscore).  The leading-underscore
    substring alone misses the standalones; dropping it outright would let
    polygon 1 match the ``BoreholePolygon_1_…`` prefix (``e`` before
    ``Polygon``, not ``_``).  Trailing underscore keeps 9 clear of 19/90.
    ``refinements``: list of (polygon_no, factor), sorted by polygon_no.
    Returns notebook lines (no mesh call here)."""
    lines = [REFINE_SECTION_TAG]
    for p_no, factor in sorted(refinements, key=lambda t: t[0]):
        lines.append(f'# Polygon {p_no}: CoarsenessFactor {factor:g}')
        lines.append(f"_refhead_{p_no} = f'Polygon_{p_no}_'")
        lines.append(f"_reftag_{p_no} = f'_Polygon_{p_no}_'")
        lines.append(f'_refn_{p_no} = 0')
        lines.append('for _refc in list(g_i.Polygons):')
        lines.append('    try:')
        lines.append('        _refnm = _refc.Name.value')
        lines.append('    except Exception:')
        lines.append('        _refnm = str(_refc.Name)')
        lines.append(f'    if _refnm.startswith(_refhead_{p_no}) '
                     f'or _reftag_{p_no} in _refnm:')
        lines.append(f"        _refc.setproperties('CoarsenessFactor', {factor:g})")
        lines.append(f'        _refn_{p_no} += 1')
        lines.append(f"print('Polygon_{p_no}: CoarsenessFactor {factor:g} "
                     f"on', _refn_{p_no}, 'slice(s)')")
    return lines


def collect_refinements(polygons):
    """Return sorted [(polygon_no, factor)] for polygons carrying a refine tag.

    Deduped by polygon_no: several spoly_name groups may share one
    polygon_no (promotion is keyed on it), but the mesh tag — and hence
    the emitted block — is per polygon_no."""
    seen = {}
    for poly in polygons or []:
        factor = poly.get('refine_factor')
        p_no = poly.get('polygon_no', 0)
        if factor is not None and p_no:
            seen.setdefault(int(p_no), float(factor))
    return sorted(seen.items(), key=lambda t: t[0])


def _write_refine_echo(sheet, polygons, refinements):
    """Echo [refine f] suffix to column C for the refined polygon_no groups.

    Idempotent: existing refine brackets are stripped first, so a re-press
    converges instead of stacking.  Only rows whose polygon_no is in
    ``refinements`` are touched — never the whole table."""
    wanted = {p_no: factor for p_no, factor in refinements}
    if not wanted:
        return
    _start = 235
    _raw_max = sheet.range(233, 9).value  # I233
    try:
        _max_rows = max(50, int(_raw_max))
    except (TypeError, ValueError):
        _max_rows = 100
    _block = sheet.range((_start, 2), (_start + _max_rows - 1, 9)).value
    if _block is None:
        return
    if not isinstance(_block, list):
        _block = [_block]
    for _idx, row_vals in enumerate(_block):
        if not isinstance(row_vals, list):
            continue
        if len(row_vals) < 8:
            continue
        try:
            p_no = int(float(row_vals[7]))  # I: polygon_no (index 7)
        except (TypeError, ValueError):
            continue
        if p_no not in wanted:
            continue
        current = safe_str(row_vals[1])  # C: cad_layer (index 1)
        if not current:
            continue
        base, _ = _parse_refine_bracket(current)
        if not base:
            # Garbage-only C cell on a tagged polygon (e.g. '[refine x]'
            # typed by mistake): leave it — the reader fails LOUD on the
            # next press instead of silently papering over the typo.
            continue
        sheet.range(_start + _idx, 3).value = f'{base} [refine {wanted[p_no]:g}]'


def read_curve_points(sheet_name_reader, wb=None):
    """Read curve-point rows from the output tab (B:F, data from row 9).

    Accepts either an xlwings Sheet or (workbook, sheet-name) pair so it
    stays testable without a live workbook."""
    if wb is not None:
        sheet = wb.sheets[CURVE_OUTPUT_SHEET]
    else:
        sheet = sheet_name_reader

    n_rows = CURVE_OUTPUT_DATA_END - CURVE_OUTPUT_DATA_START + 1
    block = sheet.range(
        (CURVE_OUTPUT_DATA_START, CURVE_OUTPUT_COL_START),
        (CURVE_OUTPUT_DATA_END, CURVE_OUTPUT_COL_END),
    ).value
    if block is None:
        return []
    if not isinstance(block, list):
        return []

    points = []
    for row_vals in block:
        if not isinstance(row_vals, list):
            row_vals = [row_vals] * 5
        name = safe_str(row_vals[0]) if len(row_vals) > 0 else ''
        cad_layer = safe_str(row_vals[1]) if len(row_vals) > 1 else ''
        x = safe_num(row_vals[2]) if len(row_vals) > 2 else None
        y = safe_num(row_vals[3]) if len(row_vals) > 3 else None
        point_type = safe_str(row_vals[4]) if len(row_vals) > 4 else ''
        if not name and not cad_layer:
            continue
        points.append({
            'name': name,
            'cad_layer': cad_layer,
            'x': x,
            'y': y,
            'type': point_type.lower(),
        })
    return points


def gen_curvepoint_code(points):
    """Generate Output curve-point notebook cells.

    Order: viewmesh → new_server → gotostages → selectmeshpoints →
    one addcurvepoint per row → update. No calculation is emitted."""
    if not points:
        return []

    coord_lines = []
    for p in points:
        ptype = p['type']
        if ptype not in ('node', 'stresspoint'):
            continue
        coord_lines.append(
            f'g_o.addcurvepoint("{ptype}", ({round_num(p["x"])}, {round_num(p["y"])}))'
        )
    if not coord_lines:
        return []

    lines = [
        CURVE_SECTION_TAG,
        '# Open Output on the generated mesh for curve-point selection.',
        'output_port = g_i.viewmesh()',
        "s_o, g_o = new_server('localhost', port=output_port, "
        'password=s_i.connection._password)',
        'g_i.gotostages()',
        'g_i.selectmeshpoints()',
    ]
    lines.extend(coord_lines)
    lines.append('g_o.update()')
    return lines


def read_loads(sheet, header_row=18):
    """Read line load data from str_2D rows 19-38 (header at row 18).

    Returns list of dicts, one per load. Max 20 rows.

    Optimized: block reads for header and data."""
    LOAD_DATA_START = 19
    LOAD_DATA_END = 38

    # Block read: header row (one COM call for B..K)
    header_block = sheet.range((header_row, 2), (header_row, 11)).value
    if header_block is None:
        return []
    if not isinstance(header_block, list):
        header_block = [header_block]

    # Build column map from block header
    col_map = {}
    for idx, val in enumerate(header_block):
        if val is not None:
            col_map[str(val).strip()] = 2 + idx  # B=2, C=3, ...

    # Block read: all data rows (one COM call for B12:K31)
    data_block = sheet.range((LOAD_DATA_START, 2), (LOAD_DATA_END, 11)).value
    if data_block is None:
        return []
    if not isinstance(data_block, list):
        data_block = [data_block]

    # Process block data in-memory
    loads = []
    for row_idx, row_vals in enumerate(data_block):
        if not isinstance(row_vals, list):
            row_vals = [row_vals] * 10  # B..K = 10 columns
        row = LOAD_DATA_START + row_idx
        # Column offsets: B=0, C=1, D=2, E=3, F=4, G=5, H=6, I=7, J=8, K=9
        cad_layer = safe_str(row_vals[1])  # C: cad_layer (offset 1)
        if not cad_layer:
            continue

        loads.append({
            'name': safe_str(row_vals[0]),             # B: Loads (offset 0)
            'cad_layer': cad_layer,
            'x1': safe_num(row_vals[2]),               # D: x1 (offset 2)
            'y1': safe_num(row_vals[3]),               # E: y1 (offset 3)
            'x2': safe_num(row_vals[4]),               # F: x2 (offset 4)
            'y2': safe_num(row_vals[5]),               # G: y2 (offset 5)
            'load_no': safe_str(row_vals[6]),          # H: load_no (offset 6)
            'load_model': safe_str(row_vals[7]),       # I: load_model (offset 7)
            'qx_start': safe_num(row_vals[8]),         # J: qx_start (offset 8)
            'qy_start': safe_num(row_vals[9]),         # K: qy_start (offset 9)
            '_row': row,
        })
    return loads


def gen_load_code(loads):
    """Generate PLAXIS code cells for line loads.
    Creates g_i.lineload() geometry and sets qx_start, qy_start properties.
    Returns list of notebook cells."""
    cells = []

    if not loads:
        return cells

    cells.append(make_cell(LOAD_TAG))
    for i, ld in enumerate(loads, 1):
        x1 = round_num(ld['x1'])
        y1 = round_num(ld['y1'])
        x2 = round_num(ld['x2'])
        y2 = round_num(ld['y2'])
        qx = round_num(ld.get('qx_start', 0))
        qy = round_num(ld.get('qy_start', 0))
        lines = [
            f'g_i.lineload({x1}, {y1}, {x2}, {y2})',
            f'g_i.LineLoad_{i}.setproperties("qx_start", {qx}, "qy_start", {qy})',
        ]
        cells.append(make_cell(lines))

    return cells


# ── Tunnel identification code ──

def _gen_tunnel_id_cells(tunnel_data):
    """Generate BBox identification code cells for SC6 to relocate.

    The code itself is generated by SC5 from the tunnel geometry. It is
    relocated by SC6 to execute after gotostages(), where SoilPolygons exist.
    """
    if not tunnel_data:
        return []

    groups = {}
    for row in tunnel_data:
        key = safe_str(row.get('cad_layer')).lower()
        if key and key not in groups:
            groups[key] = row

    lines = [
        TUNNEL_ID_TAG,
        'import math',
        '',
        'def _sc5_bbox_value(proxy):',
        '    try:',
        '        return float(proxy.value)',
        '    except AttributeError:',
        '        return float(proxy)',
        '',
    ]
    for tun_no, row in enumerate(groups.values(), 1):
        x = round_num(safe_num(row.get('x1')))
        y = round_num(safe_num(row.get('y1')))
        r = round_num(safe_num(row.get('y2')))
        lines.extend([
            f'# Tunnel {tun_no}: center=({x}, {y}), R_DXF={r} (outer primary surface)',
            f'tunnel_{tun_no}_candidates = []',
            'for _sp in g_i.SoilPolygons[:]:',
            '    _bb = _sp.BoundingBox',
            '    _xmin = _sc5_bbox_value(_bb.xMin)',
            '    _xmax = _sc5_bbox_value(_bb.xMax)',
            '    _ymin = _sc5_bbox_value(_bb.yMin)',
            '    _ymax = _sc5_bbox_value(_bb.yMax)',
            '    _xmid = (_xmin + _xmax) / 2.0',
            '    _ymid = (_ymin + _ymax) / 2.0',
            '    _area = (_xmax - _xmin) * (_ymax - _ymin)',
            f'    _dist = math.sqrt((_xmid - {x}) ** 2 + (_ymid - ({y})) ** 2)',
            f'    if _dist < {r}:',
            f'        tunnel_{tun_no}_candidates.append((_sp, _area))',
            f'tunnel_{tun_no}_candidates.sort(key=lambda item: item[1])',
            f'tunnel_{tun_no}_opening = tunnel_{tun_no}_candidates[0][0] if tunnel_{tun_no}_candidates else None',
            f'tunnel_{tun_no}_liners = [item[0] for item in tunnel_{tun_no}_candidates[1:]]',
            '',
        ])
    lines.append('print("Tunnel identification complete.")')
    return [make_cell(lines)]


# ── Point Load functions ──

def gen_pointload_code(loads):
    """Generate PLAXIS code cells for point loads.
    Creates g_i.pointload() geometry and sets Fx, Fy properties.
    Returns list of notebook cells."""
    cells = []

    if not loads:
        return cells

    cells.append(make_cell('# ── Point Loads ──'))
    for i, ld in enumerate(loads, 1):
        x1 = round_num(ld['x1'])
        y1 = round_num(ld['y1'])
        qx = round_num(ld.get('qx_start', 0))
        qy = round_num(ld.get('qy_start', 0))
        lines = [
            f'g_i.pointload({x1}, {y1})',
            f'g_i.PointLoad_{i}.setproperties("Fx", {qx}, "Fy", {qy})',
        ]
        cells.append(make_cell(lines))

    return cells


def _gwflowbc_sequence_name(gwc, closed_counter):
    """Return the descriptive assign_sequence name for one GWFlowBC."""
    behaviour = gwc.get('behaviour', 'Closed')
    href = gwc.get('href')
    cad_layer = gwc.get('cad_layer', 'waterboundary')
    if behaviour == 'Head' and href is not None:
        sign = '+' if float(href) >= 0 else ''
        return f'{cad_layer}_{sign}{float(href):g}', closed_counter
    return f'{cad_layer}_{closed_counter + 1}', closed_counter + 1


def _write_assign_sequence(sheet, plate_data, n2n_data, anc_data, strut_data, pile_data,
                           soil_polys=None, drains=None, loads=None, fea_data=None,
                           tunnel_data=None, geogrid_data=None, waterlines=None,
                           gwflowbcs=None):
    """Write the assign_sequence tab showing which elements will be created in PLAXIS.
    Uses xlwings to write directly into the workbook (preserves VBA macros)."""
    # ── Build sequence list ──
    sequence = []  # (name, type, plaxis_label, geo, parent_plate)

    # Helper: extract short type from name (e.g. "diaphragm_wall_THK=1.0m" → "diaphragm_wall")
    def _short_type(name):
        n = safe_str(name)
        if not n:
            return '?'
        prefix = n.split('_THK=')[0].split('_D=')[0].split('_L=')[0]
        return prefix

    # Helper: extract geometry dict from data row
    def _geo(d):
        return {
            'cad_layer': safe_str(d.get('cad_layer')),
            'x1': d.get('x1'),
            'y1': d.get('y1'),
            'x2': d.get('x2'),
            'y2': d.get('y2'),
        }
    _empty_geo = {'cad_layer': '', 'x1': None, 'y1': None, 'x2': None, 'y2': None}

    # Plates (Plate_1, Plate_2, ...)
    pl_counter = 0
    contraction_counter = 0
    for d in plate_data:
        pl_counter += 1
        name = safe_str(d.get('plate_name'), 'plate')
        ptype = _short_type(d.get('plate_name'))
        sequence.append((name, ptype, f'Plate_{pl_counter}', _geo(d), None, None))

    # N2N anchors (NodeToNodeAnchor_1 through _10)
    nn_counter = 0
    for d in n2n_data:
        nn_counter += 1
        name = safe_str(d.get('nn_name'), 'nn')
        ntype = safe_str(d.get('nn_type'), 'nn')
        sequence.append((name, ntype, f'NodeToNodeAnchor_{nn_counter}', _geo(d), None, None))

    # Embedded beams (EmbeddedBeam_1 through _10)
    eb_counter = 0
    for d in anc_data:
        eb_counter += 1
        name = safe_str(d.get('anc_name'), 'anc')
        atype = safe_str(d.get('anc_type'), 'grout_body')
        sequence.append((name, atype, f'EmbeddedBeam_{eb_counter}', _geo(d), None, None))

    # Strut (continues NodeToNodeAnchor numbering)
    for d in strut_data:
        nn_counter += 1
        name = safe_str(d.get('nn_name'), 'strut')
        sequence.append((name, 'strut_beam', f'NodeToNodeAnchor_{nn_counter}', _geo(d), None, None))

    # Pile (continues EmbeddedBeam numbering)
    for d in pile_data:
        eb_counter += 1
        name = safe_str(d.get('pile_name'), 'pile')
        ptype = safe_str(d.get('anc_type'), 'pile')
        sequence.append((name, ptype, f'EmbeddedBeam_{eb_counter}', _geo(d), None, None))

    # Fixed end anchors (FixedEndAnchor_1, FixedEndAnchor_2, ...)
    fea_counter = 0
    if fea_data:
        for d in fea_data:
            fea_counter += 1
            name = safe_str(d.get('fe_name'), 'fea')
            sequence.append((name, 'fixedendanchor', f'FixedEndAnchor_{fea_counter}', _geo(d), None, None))

    # ── Tunnels ──
    # v0.6.15: Tunnel designer objects + their generated plate/volume sub-elements.
    # Plate segments share the global Plate_N counter with wall/slab plates.
    # Tunnel interface counters share the global PositiveInterface/NegativeInterface counters.
    _tunnel_seq = []  # (cad_key, {'index': int, 'rows': [...]})
    _poly_counter = 0
    if soil_polys:
        _poly_counter = len(soil_polys)
    _tun_idx = 0

    # Interfaces — ordinary plate interfaces FIRST (matches notebook creation order)
    # PLAXIS assigns PositiveInterface_N / NegativeInterface_N by call order,
    # so assign_sequence must enumerate in the same order as the notebook.
    pos_idx = 0
    neg_idx = 0
    pl_idx = 0
    for d in plate_data:
        pl_idx += 1
        iface = safe_str(d.get('interface'), 'none').lower()
        ptype = _short_type(d.get('plate_name'))
        plate_label = f'Plate_{pl_idx}'
        if iface in ('positive', 'both'):
            pos_idx += 1
            sequence.append((f'{ptype} (positive)', 'interface', f'PositiveInterface_{pos_idx}', _geo(d), plate_label, None))
        if iface in ('negative', 'both'):
            neg_idx += 1
            sequence.append((f'{ptype} (negative)', 'interface', f'NegativeInterface_{neg_idx}', _geo(d), plate_label, None))

    if tunnel_data:
        for d in tunnel_data:
            position = safe_str(d.get('liner_position')).lower()
            # v0.7.1: 'grout' accepted (TBM mode: primary + grout + secondary)
            if position not in ('primary', 'secondary', 'grout'):
                continue
            _cad = safe_str(d.get('cad_layer')).lower()
            _existing = None
            for tinfo in _tunnel_seq:
                if tinfo[0] == _cad:
                    _existing = tinfo
                    break
            if _existing is None:
                _tun_idx += 1
                _tinfo = (_cad, {'index': _tun_idx, 'rows': []})
                _tunnel_seq.append(_tinfo)
                _tinfo[1]['rows'].append(d)
            else:
                _existing[1]['rows'].append(d)

        # Write Tunnel_N designer object entries
        for _cad, _tinfo in _tunnel_seq:
            tun_num = _tinfo['index']
            tun_rows = _tinfo['rows']
            _d0 = tun_rows[0]
            sequence.append((
                f'tunnel_{tun_num}_full',
                'tunnel',
                f'Tunnel_{tun_num}',
                _geo(_d0),
                None,
                None
            ))
            # v0.6.16: Add tunnel_opening row (BBox-discovered proxy variable)
            sequence.append((
                f'tunnel_{tun_num}_opening',
                'tunnel_opening',
                f'tunnel_{tun_num}_opening',
                _geo(_d0),
                f'Tunnel_{tun_num}',
                None
            ))
            # Process rows in primary-first order, matching geometry generation.
            # Structural-volume proxy indices are assigned separately below in
            # BBox order (ascending area: opening -> secondary -> primary).
            tun_rows = sorted(
                tun_rows,
                key=lambda _row: 0 if safe_str(_row.get('liner_position')).lower() == 'primary' else 1,
            )
            _volume_rows = [
                _row for _row in tun_rows
                if 'str_volume' in safe_str(_row.get('liner_model')).lower()
            ]
            _volume_rows_bbox_order = sorted(
                _volume_rows,
                key=lambda _row: 0 if safe_str(_row.get('liner_position')).lower() == 'secondary' else 1,
            )
            _volume_proxy_index = {id(_row): _idx for _idx, _row in enumerate(_volume_rows_bbox_order)}
            _tbm_group = any(
                safe_str(_r.get('liner_position')).lower() == 'grout'
                for _r in tun_rows
            )
            _tbm_block_start = None
            _tbm_config_b = False
            if _tbm_group:
                _tbm_secondary = next(
                    (_row for _row in tun_rows
                     if safe_str(_row.get('liner_position')).lower() == 'secondary'),
                    None,
                )
                _tbm_config_b = (
                    _tbm_secondary is not None
                    and 'str_volume' in safe_str(_tbm_secondary.get('liner_model')).lower()
                )
                # Reserve this tunnel's complete actual Plate_N block before
                # processing its rows. This keeps later tunnel labels synced
                # even when a TBM tunnel occurs in the middle of the sequence.
                _tbm_block_start = pl_counter + 1
                pl_counter += 2 if _tbm_config_b else 4
            for _d in tun_rows:
                position = safe_str(_d.get('liner_position')).lower()
                liner_model = safe_str(_d.get('liner_model')).lower()
                thk = safe_num(_d.get('plate_THK(m)'))
                if 'eplastic' in liner_model:
                    model_key = 'plate_eplastic'
                elif 'plate' in liner_model:
                    model_key = 'plate_elastic'
                elif 'str_volume' in liner_model:
                    model_key = 'str_volume'
                else:
                    model_key = liner_model
                _d['_model_key'] = model_key

                if model_key in ('plate_elastic', 'plate_eplastic'):
                    if _tbm_group:
                        # v0.7.1: PLAXIS assigns tunnel plate labels in complete
                        # SliceSegments order (s0..s[n-1]), even though SC5
                        # creates the rings shield-first, main-second. Map each
                        # ring to its palindrome positions within the block
                        # reserved above (group level, keeps later tunnels in
                        # sync when TBM is in the middle of the sequence).
                        #   ring k (ring 0 = shield) -> Plate_{start+k},
                        #   Plate_{start + n-1-k};  n = 4 (A) or 6 (B)
                        # Config A: shield ring 0, main ring 1.
                        # Config B: only shield plated (ring 0).
                        if position == 'primary':
                            _ring_index = 0
                        elif position == 'secondary' and not _tbm_config_b:
                            _ring_index = 1
                        else:
                            continue  # Config B secondary is str_volume (handled above)
                        # Config A: both shield and main rings are plated,
                        # so all four SliceSegments occupy the block and the
                        # palindrome labels are P/P+3 and P+1/P+2.
                        # Config B: only the two shield slices receive plates;
                        # there are no plate objects at s1..s4, so PLAXIS
                        # gives the shield the two consecutive labels P/P+1.
                        if _tbm_config_b:
                            arc_labels = [
                                f'Plate_{_tbm_block_start}',
                                f'Plate_{_tbm_block_start + 1}',
                            ]
                        else:
                            arc_labels = [
                                f'Plate_{_tbm_block_start + _ring_index}',
                                f'Plate_{_tbm_block_start + 4 - 1 - _ring_index}',
                            ]
                    else:
                        # Existing single/double path: two consecutive labels.
                        arc_labels = [f'Plate_{pl_counter + 1}', f'Plate_{pl_counter + 2}']
                        pl_counter += 2
                    # v0.7.20: component display name always carries the
                    # tunnel number so assign matches str_2D tunnel_name
                    # (TBM: tunnel1_TBM_..; ordinary: tunnel2_primary_..).
                    tname = _tunnel_component_name(
                        tun_num, position, thk, _tbm_group)
                    for arc_label in arc_labels:
                        sequence.append((tname, 'tunnel_plate', arc_label, _geo(_d), f'Tunnel_{tun_num}', None))
                    _d['_plate_labels'] = arc_labels
                    iface = safe_str(_d.get('interface'), 'none').lower()
                    if iface in ('positive', 'both'):
                        for arc_label in arc_labels:
                            pos_idx += 1
                            sequence.append((f'{tname} (positive)', 'interface', f'PositiveInterface_{pos_idx}', _geo(_d), arc_label, None))
                    if iface in ('negative', 'both'):
                        for arc_label in arc_labels:
                            neg_idx += 1
                            sequence.append((f'{tname} (negative)', 'interface', f'NegativeInterface_{neg_idx}', _geo(_d), arc_label, None))
                elif model_key == 'str_volume':
                    _poly_counter += 1
                    # v0.7.2: str_volume keeps the tunnel number (mat var
                    # name); TBM groups additionally use TBM display names.
                    tname = _tunnel_component_name(tun_num, position, thk, _tbm_group)
                    # v0.6.16: Use BBox proxy variable for liner identification.
                    # BBox area ascending: opening → secondary → primary,
                    # so liners[0] = secondary, liners[1] = primary.
                    _liner_var = f'tunnel_{tun_num}_liners[{_volume_proxy_index[id(_d)]}]'
                    _mat_var = _clean_mat_var_name(tname)
                    sequence.append((_mat_var, 'tunnel_volume', _liner_var, _geo(_d), f'Tunnel_{tun_num}', None))
                    _d['_poly_label'] = _liner_var

            # Line contractions are created in SC5 geometry immediately after
            # this tunnel's plate/interface setup and before thick lining.
            # Circular tunnels have two SliceSegments. TBM uses the two shield
            # slices only, but still creates two LineContraction objects.
            # Keep these rows in the same tunnel-group order as the notebook;
            # SC6 can then activate/deactivate them by normal markers.
            _c_ref_seq = safe_num(tun_rows[0].get('C_ref'))
            if _c_ref_seq <= 0:
                _c_ref_seq = safe_num(next(
                    (_r.get('C_ref') for _r in tun_rows
                     if safe_num(_r.get('C_ref')) > 0),
                    0,
                ))
            if _c_ref_seq > 0:
                for _ in range(2):
                    contraction_counter += 1
                    sequence.append((
                        f'tunnel_{tun_num}_contraction',
                        'line_contraction',
                        f'LineContraction_{contraction_counter}',
                        _geo(_d0),
                        f'Tunnel_{tun_num}',
                        None,
                    ))

    # Geogrids are created after the tunnel/plate/interface sequence. Their
    # interface labels therefore continue the same global PLAXIS counters.
    if geogrid_data:
        for i, d in enumerate(geogrid_data, 1):
            name = geogrid_display_name(d, i)
            sequence.append((name, 'geogrid', f'Geogrid_{i}', _geo(d), None, None))

        for i, d in enumerate(geogrid_data, 1):
            name = geogrid_display_name(d, i)
            positive_label, negative_label = geogrid_interface_labels(
                i, d, pos_idx + 1, neg_idx + 1,
            )
            if positive_label:
                pos_idx += 1
                sequence.append((f'{name} (positive)', 'interface', positive_label,
                                 _geo(d), f'Geogrid_{i}', None))
            if negative_label:
                neg_idx += 1
                sequence.append((f'{name} (negative)', 'interface', negative_label,
                                 _geo(d), f'Geogrid_{i}', None))

    # Volume-profile interfaces follow geogrid interfaces in global order.
    # Counters are derived from the actual preceding interface calls so the
    # notebook labels and assign_sequence labels remain synchronized.
    if soil_polys:
        _volume_records, _, _ = _plan_volume_profile_interfaces(
            soil_polys, pos_idx, neg_idx,
        )
        for record in _volume_records:
            parent_no = next(
                (p.get('polygon_no') for p in soil_polys
                 if p.get('name') == record['volume_name']), '')
            sequence.append((
                f"{record['volume_name']} ({record['side']})",
                'interface', record['label'],
                {'cad_layer': 'volume_profile', 'x1': record['x1'],
                 'y1': record['y1'], 'x2': record['x2'], 'y2': record['y2']},
                f'Polygon_{parent_no}', None,
            ))

    # Soil Polygons — sorted by polygon_no (from sc4), interface modes are
    # normalized to str_volume for staged material assignment.

    if soil_polys:
        _soil = [p for p in soil_polys if not _is_volume_profile_type(p.get('type'))]
        _vols = [p for p in soil_polys if _is_volume_profile_type(p.get('type'))]
        _soil_sorted = sorted(_soil, key=lambda p: p.get('polygon_no', 0))
        _vol_sorted = sorted(_vols, key=lambda p: p.get('polygon_no', 0))
        _all_polys = _soil_sorted + _vol_sorted
        for poly in _all_polys:
            verts = poly.get('vertices', [])
            p_no = poly.get('polygon_no', 0)
            geo = {
                'cad_layer': poly.get('type', ''),
                'x1': verts[0][0] if verts else None,
                'y1': verts[0][1] if verts else None,
                'x2': verts[-1][0] if verts else None,
                'y2': verts[-1][1] if verts else None,
            }
            if poly.get('type') == 'soil_replacement':
                # v0.6.18: soil_replacement uses material_after (column K)
                # for SC6 setmaterial(). Column J (material) is the existing
                # soil, assigned at creation. Column K is the replacement,
                # assigned on reactivation during staged construction.
                # v0.6.20: blank K writes '' (not None) so SC6 never
                # emits setmaterial(..., None).
                material_after = poly.get('material_after', '')
                mat_name = ("".join(e for e in material_after.title() if e.isalnum())
                            if material_after else '')
                sequence.append((mat_name, 'soil_replacement',
                                 f'Polygon_{p_no}', geo, None, None))
            elif _is_volume_profile_type(poly.get('type')):
                # v0.6.18: str_volume (volume_profile) also uses material_after
                # (column K) for SC6 setmaterial() on reactivation — same
                # hybrid pattern as soil_replacement.
                # v0.6.20: blank K writes '' (not None) so SC6 never
                # emits setmaterial(..., None).
                material_after = poly.get('material_after', '')
                mat_name = ("".join(e for e in material_after.title() if e.isalnum())
                            if material_after else '')
                sequence.append((mat_name, 'str_volume',
                                 f'Polygon_{p_no}', geo, None, None))
            else:
                sequence.append((poly['name'], 'soil_polygon',
                                 f'Polygon_{p_no}', geo, None, None))

    # Drains (PVD) — Drain_1, Drain_2, ...
    if drains:
        drain_counter = 0
        for d in drains:
            drain_counter += 1
            name = safe_str(d.get('name'), 'drain')
            sequence.append((name, 'drain', f'Drain_{drain_counter}', _geo(d), None, None))

    # Loads — LineLoad_N for line loads, PointLoad_N for point loads
    if loads:
        line_counter = 0
        point_counter = 0
        for d in loads:
            load_model = safe_str(d.get('load_model'), 'line_load')
            if load_model == 'line_load':
                line_counter += 1
                name = safe_str(d.get('name'), 'load')
                sequence.append((name, 'line_load', f'LineLoad_{line_counter}', _geo(d), None, None))
            elif load_model == 'point_load':
                point_counter += 1
                name = safe_str(d.get('name'), 'load')
                sequence.append((name, 'point_load', f'PointLoad_{point_counter}', _geo(d), None, None))

    # ── Water levels (global_waterlevel) ──
    # BoreholeWaterLevel_1 is created automatically by PLAXIS when the
    # single SC3 borehole is created.  UserWaterLevel_N objects are created
    # in gotoflow() from DXF waterlines.  SC6 interprets A markers on these
    # rows as setglobalwaterlevel(), not activate().
    sequence.append((
        'BoreholeWaterLevel_1',
        'global_waterlevel',
        'BoreholeWaterLevel_1',
        _empty_geo,
        None,
        None,
    ))
    if waterlines:
        for wl in waterlines:
            wl_no = wl['waterline_no']
            var_name = f"UserWaterLevel_{wl_no}"
            sequence.append((
                var_name,
                'global_waterlevel',
                var_name,
                _empty_geo,
                None,
                None,
            ))

    # ── Groundwater flow boundaries ──
    # Display name: Head → cad_layer_+href, Closed → cad_layer_N (sequential counter)
    if gwflowbcs:
        closed_counter = 0
        for gwc in gwflowbcs:
            gwc_no = gwc['gwflowbc_no']
            display_name, closed_counter = _gwflowbc_sequence_name(gwc, closed_counter)
            sequence.append((
                display_name,
                'gwflowbc',
                f'GWFlowBC_{gwc_no}',
                _geo(gwc),
                None,
                None,
            ))

    # ── Write to Excel ──
    # Create or clear the assign_sequence sheet
    try:
        seq_sheet = sheet.book.sheets.add('assign_sequence')
    except Exception:
        seq_sheet = sheet.book.sheets('assign_sequence')
        # Clear existing data (B6:J200)
        seq_sheet.range('B6:J200').clear_contents()

    # Block write: headers (one COM call)
    headers_row = ['name', 'type', 'PLAXIS Label', 'cad_layer', 'x1', 'y1', 'x2', 'y2', 'parent_plate']
    seq_sheet.range((6, 2), (6, 10)).value = [headers_row]

    # Block write: all data rows (one COM call for the entire block)
    if sequence:
        block_data = []
        for name, etype, label, geo, parent, mat_var in sequence:
            block_data.append([
                name,                    # B
                etype,                   # C
                label,                   # D
                geo['cad_layer'],        # E
                geo['x1'],               # F
                geo['y1'],               # G
                geo['x2'],               # H
                geo['y2'],               # I
                parent,                  # J
            ])
        seq_sheet.range((7, 2), (7 + len(sequence) - 1, 10)).value = block_data

    n = len(sequence)
    print(f"sc5: assign_sequence tab updated ({n} elements)")


def main():
    """
    Read str_2D tab -> append structural PLAXIS API cells to .ipynb.
    Called via: RunPython "import scripts.sc5_structural_to_ipynb as l; l.main()"
    """
    # ── Get workbook ──
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

    # ── Disable screen updating during I/O ──
    app = wb.app
    previous_screen_updating = app.screen_updating
    app.screen_updating = False
    try:
        _main_io(wb)
    finally:
        app.screen_updating = previous_screen_updating


def _main_io(wb):
    """SC5 entry — progress wrapper (mirrors SC6/SC8 pattern).

    Sets the WORKING banner on str_2D!D4; crash → FAILED,
    success path sets the done text. Never raises.
    """
    _set_progress(wb, f"sc5: {_spin()} WORKING — reading str_2D…")
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
                wb, "sc5: FAILED — see console for detail")


def _main_io_inner(wb):
    """Core logic for SC5 — separated from main() for screen_updating wrapper."""
    sheet = wb.sheets['str_2D']

    _set_progress(wb, f"sc5: {_spin()} WORKING — reading sections {_bar(0.05)}")
    # ── Read soil profile (rows 7-16, header at row 6) ──
    soil_data, soil_cols = read_soil_profile(sheet, header_row=6, max_rows=10)

    # ── Determine model type from main tab (for is_3d check) ──
    is_3d = False
    try:
        main_sheet = wb.sheets['main']
        plaxis_type = safe_str(main_sheet['B2'].value)
        is_3d = plaxis_type == '3D'
    except Exception:
        pass

    # ── Read sections (fixed row positions) — returns (data, col_map) ──
    # v0.6.19: soil profile expanded to 10 rows; all lower sections shift by +7
    plate_data, pl_cols  = read_section(sheet, 42, max_rows=20)   # header 42, data 43-62
    tunnel_data, tun_cols = read_section(sheet, 66, max_rows=20)   # header 66, data 67-86
    n2n_data, nn_cols    = read_section(sheet, 90, max_rows=20)   # header 90, data 91-110
    anc_data, anc_cols   = read_section(sheet, 114, max_rows=20)  # header 114, data 115-134
    strut_data, str_cols = read_section(sheet, 138, max_rows=20)  # header 138, data 139-158
    fea_data, fea_cols   = read_section(sheet, 162, max_rows=20)  # header 162, data 163-182
    pile_data, pil_cols  = read_section(sheet, 186, max_rows=20)  # header 186, data 187-206
    geogrid_data = read_geogrids(sheet)

    # ── Read soil polygons (for assign_sequence + volume liner counter) ──
    soil_polys = read_soil_polygons(sheet)
    _set_progress(wb, f"sc5: {_spin()} WORKING — inheritance {_bar(0.15)}")

    # ── Apply parent/child inheritance (group by DXF layer name) ──
    # All elements on same DXF layer share one material
    tunnel_inherit = ['liner_model', 'plate_THK(m)', 'plate_E(Mpa)', 'plate_A1(m2)',
                      'v_nu', 'plate_w(kN/m/m)', 'interface', 'prevent_punching',
                      'isotropic', 'plate_A2(m2)', 'Mp(kNm/m)', 'Np1tens(kN/m)',
                      'Np2tens(kN/m)']
    tunnel_groups = {}
    tunnel_parent_rows = []
    for i, d in enumerate(tunnel_data):
        key = (safe_str(d.get('cad_layer')).lower(),
               safe_str(d.get('liner_position')).lower())
        if not key[0]:
            continue
        if key not in tunnel_groups:
            tunnel_groups[key] = d
            tunnel_parent_rows.append(i)
        else:
            parent = tunnel_groups[key]
            for header in tunnel_inherit:
                if _is_inheritable_blank(d.get(header)):
                    d[header] = parent.get(header)

    nn_inherit = ['nn_type', 'nn_model', 'L_spacing(m)', 'nn_E(Mpa)', 'nn_A(m2)',
                  'Fmax_tens', 'Fmax_comp']
    nn_parents = apply_inheritance(n2n_data, 'cad_layer', nn_inherit)
    str_parents = apply_inheritance(strut_data, 'cad_layer', nn_inherit)

    anc_inherit = ['anc_type', 'anc_model', 'L_spacing(m)', 'anc_E(Mpa)', 'anc_w(kN/m/m)',
                   'diameter(m)', 'resistance_model', 'Tskin_start', 'Tskin_end',
                   'Mp(kNm/m)', 'Nptens(kN/m)']
    eb_parents = apply_inheritance(anc_data, 'cad_layer', anc_inherit)

    fea_inherit = ['fe_type', 'fe_model', 'L_spacing(m)', 'nn_E(Mpa)', 'nn_A(m2)',
                   'Fmax_tens', 'Fmax_comp', 'Direction_x']
    fea_parents = apply_inheritance(fea_data, 'cad_layer', fea_inherit)

    pil_inherit = ['anc_type', 'anc_model', 'L_spacing(m)', 'anc_E(Mpa)', 'anc_w(kN/m/m)',
                   'diameter/width(m)', 'pile_THK(m)', 'resistance_model', 'Tskin_start', 'Tskin_end',
                   'Fmax', 'Mp(kNm/m)', 'Nptens(kN/m)']
    pil_parents = apply_inheritance(pile_data, 'cad_layer', pil_inherit)

    pl_inherit = ['plate_type', 'plate_model', 'plate_THK(m)', 'plate_E(Mpa)', 'plate_A1(m2)',
                  'v_nu', 'plate_w(kN/m/m)', 'interface', 'prevent_punching', 'isotropic',
                  'plate_A2(m2)', 'Mp(kNm/m)', 'Np1tens(kN/m)', 'Np2tens(kN/m)']
    pl_parents = apply_inheritance(plate_data, 'cad_layer', pl_inherit)

    geogrid_inherit = ['gg_type', 'gg_model', 'gg_E(Mpa)', 'gg_A1(m2)',
                       'interface', 'isotropic', 'gg_A2(m2)',
                       'Np1(kN/m)', 'Np2(kN/m)']
    geogrid_parents = apply_inheritance(geogrid_data, 'cad_layer', geogrid_inherit)

    # Convert to sets for O(1) lookup in orange writes
    pl_parent_set = set(pl_parents)
    tunnel_parent_set = set(tunnel_parent_rows)
    nn_parent_set = set(nn_parents)
    eb_parent_set = set(eb_parents)
    str_parent_set = set(str_parents)
    fea_parent_set = set(fea_parents)
    pil_parent_set = set(pil_parents)
    geogrid_parent_set = set(geogrid_parents)

    # ── Calculate & write orange columns + auto-fill blanks + N/A ──
    # All writes use write_orange(ws, row, col_map, header_name, value)
    # — no hardcoded column positions, fully header-name-driven.

    # --- Plate ---
    for i, d in enumerate(plate_data, 1):
        row = 42 + i
        is_child = (i - 1) not in pl_parent_set  # enumerate starts at 1, parent indices are 0-based
        plate_type = safe_str(d.get('plate_type'), 'plate').lower()
        _is_custom = plate_type == 'custom_wall'
        ea1, ei, ea2 = resolve_plate_stiffness(d)

        # Calculate inertia: I = h³/12 per unit width
        a2 = safe_num(d.get('plate_A2(m2)'))
        thk = safe_num(d.get('plate_THK(m)'))
        h = a2 if a2 > 0 else thk
        inertia = round(h ** 3 / 12, 6) if h > 0 else 0

        # Write calculated orange cells
        write_orange(sheet, row, pl_cols, 'plate_I(m4)', inertia)
        write_orange(sheet, row, pl_cols, 'plate_EI', ei)
        write_orange(sheet, row, pl_cols, 'plate_EA1', ea1)

        # custom_wall: flag missing manual stiffness in cells (values still flow as 0)
        if _is_custom:
            if ei == 0:
                write_orange(sheet, row, pl_cols, 'plate_EI', 'MISSING plate_EI')
            if ea1 == 0:
                write_orange(sheet, row, pl_cols, 'plate_EA1', 'MISSING plate_EA1')

        # EA2: calculated value if anisotropic, N/A if isotropic
        iso = _normalize_isotropic(d.get('isotropic'))
        if iso == 'no':
            write_orange(sheet, row, pl_cols, 'plate_EA2', ea2)
            if _is_custom and safe_num(d.get('plate_EA2')) == 0:
                write_orange(sheet, row, pl_cols, 'plate_EA2', 'MISSING plate_EA2')
        else:
            write_orange(sheet, row, pl_cols, 'plate_EA2', na_label(is_child))

        # Auto-generate plate_name
        # custom_wall: bracket content from cad_layer; otherwise: type_THK=thicknessm
        if _is_custom:
            pname = custom_wall_plate_name(d.get('cad_layer'))
            if pname is None:
                pname = 'MISSING brackets in cad_layer'
            write_orange(sheet, row, pl_cols, 'plate_name', pname)
        else:
            pname = f"{plate_type}_THK={thk}m"
            write_orange(sheet, row, pl_cols, 'plate_name', pname)
        d['plate_name'] = pname

        # Mp/Np1/Np2: N/A if elastic, MISSING if elastoplastic with empty values
        model = safe_str(d.get('plate_model'), 'elastic').lower()
        if model == 'elastic':
            write_orange(sheet, row, pl_cols, 'Mp(kNm/m)', na_label(is_child))
            write_orange(sheet, row, pl_cols, 'Np1tens(kN/m)', na_label(is_child))
            write_orange(sheet, row, pl_cols, 'Np2tens(kN/m)', na_label(is_child))
        elif model == 'elastoplastic':
            mp_v = safe_str(d.get('Mp(kNm/m)'))
            np1_v = safe_str(_first_present(d, 'Np1(kN/m)', 'Np1tens(kN/m)'))
            np2_v = safe_str(_first_present(d, 'Np2(kN/m)', 'Np2tens(kN/m)'))
            if mp_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, pl_cols, 'Mp(kNm/m)', 'MISSING Mp')
            if np1_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, pl_cols, 'Np1tens(kN/m)', 'MISSING Np1Tens')
            # Np2Tens only required when anisotropic (isotropic=no)
            if iso == 'no' and np2_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, pl_cols, 'Np2tens(kN/m)', 'MISSING Np2Tens')
            elif iso != 'no':
                write_orange(sheet, row, pl_cols, 'Np2tens(kN/m)', na_label(is_child))

        # A2: N/A if isotropic, MISSING if anisotropic with empty A2
        if iso == 'yes':
            write_orange(sheet, row, pl_cols, 'plate_A2(m2)', na_label(is_child))
        elif iso == 'no':
            a2_v = safe_str(d.get('plate_A2(m2)'))
            if a2_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, pl_cols, 'plate_A2(m2)', 'MISSING A2')

        d['plate_EA1'] = ea1
        d['plate_EI']  = ei
        d['plate_EA2'] = ea2

    # --- Geogrid ---
    # SC5 currently validates and records geogrid inputs only. PLAXIS
    # geogrid material and staging commands await live V2025 API confirmation.
    geogrid_cols = {header: col for header, col in zip(
        [safe_str(sheet.range(GEOGRID_HEADER_ROW, c).value) for c in range(2, 19)],
        range(2, 19),
    ) if header}
    geogrid_input_headers = (
        'gg_type', 'gg_model', 'gg_E(Mpa)', 'gg_A1(m2)', 'interface',
        'isotropic', 'gg_A2(m2)', 'Np1(kN/m)', 'Np2(kN/m)',
    )
    geogrid_warning_headers = (
        'gg_E(Mpa)', 'gg_A1(m2)', 'interface', 'gg_A2(m2)',
        'Np1(kN/m)', 'Np2(kN/m)',
    )
    for i, d in enumerate(geogrid_data, 1):
        row = d.get('_row', GEOGRID_DATA_START)
        # Remove warning text left in the workbook by an earlier SC5 run.
        # Valid inherited values are already present in d at this point.
        for header in geogrid_warning_headers:
            if safe_str(d.get(header)).startswith((
                'MISSING ', 'INVALID ', 'MISMATCH ',
            )):
                write_orange(sheet, row, geogrid_cols, header, None)
        name = geogrid_display_name(d, i)
        write_orange(sheet, row, geogrid_cols, 'geogrid_name', name)
        d['geogrid_name'] = name
        ea1, ea2 = calculate_geogrid_ea(d)
        write_orange(sheet, row, geogrid_cols, 'gg_EA1', ea1)
        write_orange(sheet, row, geogrid_cols, 'gg_EA2', ea2)
        is_child = (i - 1) not in geogrid_parent_set
        if safe_str(d.get('gg_model'), 'elastic').lower() == 'elastic':
            na_value = na_label(is_child)
            write_orange(sheet, row, geogrid_cols, 'Np1(kN/m)', na_value)
            write_orange(sheet, row, geogrid_cols, 'Np2(kN/m)', na_value)
            d['Np1(kN/m)'] = 'N/A'
            d['Np2(kN/m)'] = 'N/A'
        elif _is_yes(d.get('isotropic')):
            write_orange(sheet, row, geogrid_cols, 'Np2(kN/m)', na_label(is_child))
            d['Np2(kN/m)'] = 'N/A'
        for warning in validate_geogrid_row(d):
            if warning.startswith('MISSING gg_E(Mpa)'):
                write_orange(sheet, row, geogrid_cols, 'gg_E(Mpa)', warning)
            elif warning.startswith('MISSING gg_A1'):
                write_orange(sheet, row, geogrid_cols, 'gg_A1(m2)', warning)
            elif warning.startswith('MISSING gg_A2') or warning.startswith('MISMATCH isotropic'):
                write_orange(sheet, row, geogrid_cols, 'gg_A2(m2)', warning)
            elif warning.startswith('MISSING Np1'):
                write_orange(sheet, row, geogrid_cols, 'Np1(kN/m)', warning)
            elif warning.startswith('MISSING Np2'):
                write_orange(sheet, row, geogrid_cols, 'Np2(kN/m)', warning)
            elif warning.startswith('MISSING or INVALID interface'):
                write_orange(sheet, row, geogrid_cols, 'interface', warning)

        d['gg_EA1'] = ea1
        d['gg_EA2'] = ea2

    # --- Tunnel ---
    # v0.6.15: Tunnel section inserted before N2N; same column layout as Plate.
    # liner_model: plate_elastic / plate_ePlastic / str_volume
    # Double plate+plate not allowed — warning in C57.
    tunnel_valid_models = {'plate_elastic', 'plate_eplastic', 'str_volume'}
    _tunnel_double_plate_warning = []

    # Pass 1: double liner validation (group by cad_layer, check primary+secondary)
    _tunnel_groups = {}  # cad_layer_lower → {position: data_dict}
    for i, d in enumerate(tunnel_data, 1):
        key = safe_str(d.get('cad_layer')).lower()
        pos = safe_str(d.get('liner_position')).lower()
        if key and pos:
            _tunnel_groups.setdefault(key, {})[pos] = d
    # Assign tunnel numbers based on group order
    _tunnel_num_map = {}  # cad_key → tun_num
    for _idx, cad_key in enumerate(_tunnel_groups.keys(), 1):
        _tunnel_num_map[cad_key] = _idx
    for cad_key, positions in _tunnel_groups.items():
        if 'primary' in positions and 'secondary' in positions:
            # v0.7.1: TBM mode has primary+grout+secondary — skip double-plate check
            if 'grout' in positions:
                continue
            pri_model = safe_str(positions['primary'].get('liner_model')).lower()
            sec_model = safe_str(positions['secondary'].get('liner_model')).lower()
            if pri_model.startswith('plate') and sec_model.startswith('plate'):
                _tunnel_double_plate_warning.append(cad_key)

    # Write double-plate warning to C57
    if _tunnel_double_plate_warning:
        warn_text = f"WARNING: {'; '.join(_tunnel_double_plate_warning)} double liner uses plate+plate. PLAXIS 2D cannot overlap two plates. Tunnel skipped."
        sheet.range(64, 3).value = warn_text

    # Pass 2: orange writes, calculations, validation
    for i, d in enumerate(tunnel_data, 1):
        row = 66 + i
        liner_model = safe_str(d.get('liner_model')).lower().replace('eplastic', 'ePlastic').replace('elastic', 'elastic')
        # Normalise for comparison (internal key)
        if 'eplastic' in liner_model or 'ePlastic' in liner_model:
            model_key = 'plate_eplastic'
        elif 'plate' in liner_model:
            model_key = 'plate_elastic'
        elif 'str_volume' in liner_model or 'str_volume' in liner_model:
            model_key = 'str_volume'
        else:
            model_key = liner_model

        position = safe_str(d.get('liner_position'), 'primary').lower()
        thk = safe_num(d.get('plate_THK(m)'))
        cad_key = safe_str(d.get('cad_layer')).lower()

        is_double_plate_skip = cad_key in _tunnel_double_plate_warning
        is_child = (i - 1) not in tunnel_parent_set

        # Auto-generate component display name. TBM names describe the
        # shield, grout, and main liner; ordinary tunnel names are unchanged.
        _tnum = _tunnel_num_map.get(cad_key, i)
        _is_tbm = 'grout' in _tunnel_groups.get(cad_key, {})
        tname = _tunnel_component_name(_tnum, position, thk, _is_tbm)
        write_orange(sheet, row, tun_cols, 'tunnel_name', tname)
        d['tunnel_name'] = tname

        # Validate liner_model
        if model_key not in tunnel_valid_models and not is_double_plate_skip:
            write_orange(sheet, row, tun_cols, 'liner_model', f'INVALID: {safe_str(d.get("liner_model"))}')

        if is_double_plate_skip:
            continue  # skip further calculations for invalid double plate tunnels

        # ── Plate liner: same calculations as plate table ──
        if model_key in ('plate_elastic', 'plate_eplastic'):
            ea1, ei, ea2 = calc_plate(d)

            a2 = safe_num(d.get('plate_A2(m2)'))
            h = a2 if a2 > 0 else thk
            inertia = round(h ** 3 / 12, 6) if h > 0 else 0

            write_orange(sheet, row, tun_cols, 'plate_I(m4)', inertia)
            write_orange(sheet, row, tun_cols, 'plate_EI', ei)
            write_orange(sheet, row, tun_cols, 'plate_EA1', ea1)

            iso = _normalize_isotropic(d.get('isotropic'))
            if iso == 'no':
                write_orange(sheet, row, tun_cols, 'plate_EA2', ea2)
            else:
                write_orange(sheet, row, tun_cols, 'plate_EA2', na_label(is_child))

            # Mp/Np1/Np2: N/A if elastic, MISSING if elastoplastic
            if model_key == 'plate_elastic':
                write_orange(sheet, row, tun_cols, 'Mp(kNm/m)', na_label(is_child))
                write_orange(sheet, row, tun_cols, 'Np1tens(kN/m)', na_label(is_child))
                write_orange(sheet, row, tun_cols, 'Np2tens(kN/m)', na_label(is_child))
            elif model_key == 'plate_eplastic':
                mp_v = safe_str(d.get('Mp(kNm/m)'))
                np1_v = safe_str(_first_present(d, 'Np1(kN/m)', 'Np1tens(kN/m)'))
                np2_v = safe_str(_first_present(d, 'Np2(kN/m)', 'Np2tens(kN/m)'))
                if mp_v in ('', 'N/A', '0', 'None'):
                    write_orange(sheet, row, tun_cols, 'Mp(kNm/m)', 'MISSING Mp')
                if np1_v in ('', 'N/A', '0', 'None'):
                    write_orange(sheet, row, tun_cols, 'Np1tens(kN/m)', 'MISSING Np1Tens')
                # Np2Tens only required when anisotropic (isotropic=no)
                if iso == 'no' and np2_v in ('', 'N/A', '0', 'None'):
                    write_orange(sheet, row, tun_cols, 'Np2tens(kN/m)', 'MISSING Np2Tens')
                elif iso != 'no':
                    write_orange(sheet, row, tun_cols, 'Np2tens(kN/m)', na_label(is_child))

            # A2: N/A if isotropic, MISSING if anisotropic with empty A2
            if iso == 'yes':
                write_orange(sheet, row, tun_cols, 'plate_A2(m2)', na_label(is_child))
            elif iso == 'no':
                a2_v = safe_str(d.get('plate_A2(m2)'))
                if a2_v in ('', 'N/A', '0', 'None'):
                    write_orange(sheet, row, tun_cols, 'plate_A2(m2)', 'MISSING A2')

            d['plate_EA1'] = ea1
            d['plate_EI']  = ei
            d['plate_EA2'] = ea2

        # ── Str volume: validate E and nu; N/A for plate-only columns ──
        elif model_key == 'str_volume':
            e_val = safe_num(d.get('plate_E(Mpa)'))
            nu_val = safe_num(d.get('v_nu'))

            if e_val <= 0:
                write_orange(sheet, row, tun_cols, 'plate_E(Mpa)', 'MISSING plate_E')
            if nu_val <= 0:
                write_orange(sheet, row, tun_cols, 'v_nu', 'MISSING v_nu')
            if thk <= 0:
                write_orange(sheet, row, tun_cols, 'plate_THK(m)', 'MISSING THK')
            # v0.7.20: plate_w on str_volume rows is the soil unit weight
            # (kN/m3) passed straight to gammaUnsat (Non-porous:
            # gammaSat is read-only). Gate it.
            if safe_num(d.get('plate_w(kN/m/m)')) <= 0:
                write_orange(sheet, row, tun_cols, 'plate_w(kN/m/m)', 'MISSING plate_w')

            # N/A for plate-only columns (plate_w kept: it carries gamma)
            write_orange(sheet, row, tun_cols, 'interface', 'N/A')
            write_orange(sheet, row, tun_cols, 'prevent_punching', 'N/A')
            write_orange(sheet, row, tun_cols, 'isotropic', 'N/A')
            write_orange(sheet, row, tun_cols, 'plate_A2(m2)', 'N/A')
            write_orange(sheet, row, tun_cols, 'Mp(kNm/m)', 'N/A')
            write_orange(sheet, row, tun_cols, 'Np1tens(kN/m)', 'N/A')
            write_orange(sheet, row, tun_cols, 'Np2tens(kN/m)', 'N/A')
            write_orange(sheet, row, tun_cols, 'plate_I(m4)', 'N/A')
            write_orange(sheet, row, tun_cols, 'plate_EI', 'N/A')
            write_orange(sheet, row, tun_cols, 'plate_EA1', 'N/A')
            write_orange(sheet, row, tun_cols, 'plate_EA2', 'N/A')

    # --- N2N ---
    for i, d in enumerate(n2n_data, 1):
        row = 90 + i
        is_child = (i - 1) not in nn_parent_set
        ea = calc_n2n_ea(d)
        write_orange(sheet, row, nn_cols, 'nn_EA', ea)

        # Auto-generate nn_name: type_L=spacingm_E(MPa)
        nn_type = safe_str(d.get('nn_type'), 'anchor')
        spacing = safe_num(d.get('L_spacing(m)'), 1)
        e_mpa = safe_num(d.get('nn_E(Mpa)'))
        aname = f"{nn_type}_L={spacing}m_{int(e_mpa)}MPa"
        write_orange(sheet, row, nn_cols, 'nn_name', aname)
        d['nn_name'] = aname

        # Fmax: N/A if elastic, MISSING if elastoplastic with empty values
        model = safe_str(d.get('nn_model'), 'elastic').lower()
        if model == 'elastic':
            write_orange(sheet, row, nn_cols, 'Fmax_tens', na_label(is_child))
            write_orange(sheet, row, nn_cols, 'Fmax_comp', na_label(is_child))
        elif model == 'elastoplastic':
            ft_v = safe_str(d.get('Fmax_tens'))
            fc_v = safe_str(d.get('Fmax_comp'))
            if ft_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, nn_cols, 'Fmax_tens', 'MISSING FmaxTens')
            if fc_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, nn_cols, 'Fmax_comp', 'MISSING FmaxComp')

        d['nn_EA'] = ea

    # --- Embedded Beam ---
    for i, d in enumerate(anc_data, 1):
        row = 114 + i
        is_child = (i - 1) not in eb_parent_set

        # Auto-generate anc_name: type_D=diameterm_L=spacingm
        anc_type = safe_str(d.get('anc_type'), 'anchor')
        dia = safe_num(d.get('diameter(m)'))
        spacing = safe_num(d.get('L_spacing(m)'), 1)
        aname = f"{anc_type}_D={dia}m_L={spacing}m"
        write_orange(sheet, row, anc_cols, 'anc_name', aname)
        d['anc_name'] = aname

        # Mp/Nptens: N/A if elastic, MISSING if elastoplastic with empty values
        model = safe_str(d.get('anc_model'), 'elastic').lower()
        if model == 'elastic':
            write_orange(sheet, row, anc_cols, 'Mp(kNm/m)', na_label(is_child))
            write_orange(sheet, row, anc_cols, 'Nptens(kN/m)', na_label(is_child))
        elif model == 'elastoplastic':
            mp_v = safe_str(d.get('Mp(kNm/m)'))
            np_v = safe_str(d.get('Nptens(kN/m)'))
            if mp_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, anc_cols, 'Mp(kNm/m)', 'MISSING Mp')
            if np_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, anc_cols, 'Nptens(kN/m)', 'MISSING Nptens')

    # --- Strut (same column layout as N2N) ---
    for i, d in enumerate(strut_data, 1):
        row = 138 + i
        is_child = (i - 1) not in str_parent_set
        ea = calc_n2n_ea(d)
        write_orange(sheet, row, str_cols, 'nn_EA', ea)

        strut_type = safe_str(d.get('nn_type'), 'strut')
        spacing = safe_num(d.get('L_spacing(m)'), 1)
        e_mpa = safe_num(d.get('nn_E(Mpa)'))
        sname = f"{strut_type}_L={spacing}m_{int(e_mpa)}MPa"
        write_orange(sheet, row, str_cols, 'nn_name', sname)
        d['nn_name'] = sname

        model = safe_str(d.get('nn_model'), 'elastic').lower()
        if model == 'elastic':
            write_orange(sheet, row, str_cols, 'Fmax_tens', na_label(is_child))
            write_orange(sheet, row, str_cols, 'Fmax_comp', na_label(is_child))
        elif model == 'elastoplastic':
            ft_v = safe_str(d.get('Fmax_tens'))
            fc_v = safe_str(d.get('Fmax_comp'))
            if ft_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, str_cols, 'Fmax_tens', 'MISSING FmaxTens')
            if fc_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, str_cols, 'Fmax_comp', 'MISSING FmaxComp')

        d['nn_EA'] = ea

    # --- FEA (fixed end anchor) ---
    for i, d in enumerate(fea_data, 1):
        row = 162 + i
        is_child = (i - 1) not in fea_parent_set
        ea = calc_n2n_ea(d)
        write_orange(sheet, row, fea_cols, 'nn_EA', ea)

        # Auto-generate fe_name: FEA_L=spacingm_E(MPa)_EqL{Direction_x}m
        spacing = safe_num(d.get('L_spacing(m)'), 1)
        e_mpa = safe_num(d.get('nn_E(Mpa)'))
        direction = safe_num(d.get('Direction_x'))
        feaname = f"FEA_L={spacing}m_{int(e_mpa)}MPa_EqL{direction}m"
        write_orange(sheet, row, fea_cols, 'fe_name', feaname)
        d['fe_name'] = feaname

        # Fmax: N/A if elastic, MISSING if elastoplastic with empty values
        model = safe_str(d.get('nn_model'), 'elastic').lower()
        if model == 'elastic':
            write_orange(sheet, row, fea_cols, 'Fmax_tens', na_label(is_child))
            write_orange(sheet, row, fea_cols, 'Fmax_comp', na_label(is_child))
        elif model == 'elastoplastic':
            ft_v = safe_str(d.get('Fmax_tens'))
            fc_v = safe_str(d.get('Fmax_comp'))
            if ft_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, fea_cols, 'Fmax_tens', 'MISSING FmaxTens')
            if fc_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, fea_cols, 'Fmax_comp', 'MISSING FmaxComp')

        d['nn_EA'] = ea

    # --- Pile ---
    for i, d in enumerate(pile_data, 1):
        row = 186 + i
        is_child = (i - 1) not in pil_parent_set
        pile_type = safe_str(d.get('anc_type'), 'solid_circular').lower()
        dia = safe_num(d.get('diameter/width(m)'))
        spacing = safe_num(d.get('L_spacing(m)'), 1)
        pname = f"Pile_{pile_type}_D={dia}m_L={spacing}m"
        write_orange(sheet, row, pil_cols, 'pile_name', pname)
        d['pile_name'] = pname

        # pile_THK: blank for children, N/A for non-circular_tube parent,
        # MISSING if circular_tube parent with empty THK
        thk_val = safe_str(d.get('pile_THK(m)'))
        if pile_type == 'circular_tube':
            if not is_child and thk_val in ('', 'N/A', '0', 'None', 'only for circular tube'):
                write_orange(sheet, row, pil_cols, 'pile_THK(m)', 'MISSING pile_THK')
        elif not is_child:
            write_orange(sheet, row, pil_cols, 'pile_THK(m)', 'N/A')

        # Mp/Nptens: N/A if elastic, MISSING if elastoplastic with empty values
        model = safe_str(d.get('anc_model'), 'elastic').lower()
        if model == 'elastic':
            write_orange(sheet, row, pil_cols, 'Mp(kNm/m)', na_label(is_child))
            write_orange(sheet, row, pil_cols, 'Nptens(kN/m)', na_label(is_child))
        elif model == 'elastoplastic':
            mp_v = safe_str(d.get('Mp(kNm/m)'))
            np_v = safe_str(d.get('Nptens(kN/m)'))
            if mp_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, pil_cols, 'Mp(kNm/m)', 'MISSING Mp')
            if np_v in ('', 'N/A', '0', 'None'):
                write_orange(sheet, row, pil_cols, 'Nptens(kN/m)', 'MISSING Nptens')

    _set_progress(wb, f"sc5: {_spin()} WORKING — orange columns {_bar(0.30)}")
    # ── Hide blank rows per structural section ──
    # Scan column C (cad_layer) directly — same reference read_section uses
    _sections = [
        (19, 38),    # Loads
        (43, 62),   # Plate
        (67, 86),   # Tunnel
        (91, 110),  # N2N
        (115, 134), # Anchor
        (139, 158), # Strut
        (163, 182), # FEA
        (187, 206), # Pile
        (211, 230),  # Geogrid
    ]
    for start, end in _sections:
        for r in range(start, end + 1):
            cad_val = sheet.range(r, 3).value  # column C
            has_data = cad_val is not None and str(cad_val).strip() != ''
            sheet.api.Rows(r).Hidden = not has_data

    # ── Find notebook ──
    nb_path = notebook_path_for_workbook(wb)
    if nb_path.exists():
        with open(nb_path, 'r', encoding='utf-8') as f:
            notebook = json.load(f)
        # Remove old structural, soil material, and drain sections if re-running
        cells = notebook.get('cells', [])
        new_cells = []
        skip = False
        for cell in cells:
            src = ''.join(cell.get('source', []))
            if SECTION_TAG in src or TUNNEL_TAG in src or TUNNEL_ID_TAG in src or SOIL_MAT_TAG in src or DRAIN_TAG in src or LOAD_TAG in src or POINT_LOAD_TAG in src or WATERLEVEL_TAG in src or GWFLOW_BC_TAG in src or MERGE_TAG in src:
                skip = True
                continue
            if skip and cell.get('cell_type') == 'code':
                # Skip all code cells in generated section: g_i calls,
                # section comments (# ── ...), and assignment lines
                if 'g_i.' in src or '# ──' in src.strip()[:4]:
                    continue
            if skip:
                skip = False
            new_cells.append(cell)
        notebook['cells'] = new_cells
    else:
        raise FileNotFoundError(
            f"sc5: Notebook '{nb_path.name}' not found for "
            f"workbook '{Path(wb.fullname).name}'. Run SC3 first."
        )

    _set_progress(wb, f"sc5: {_spin()} WORKING — cleaning notebook {_bar(0.45)}")
    # ── Generate cells ──
    # v0.5: one material per DXF layer; children with custom values get own material
    new_cells = [make_cell(SECTION_TAG, cell_type='markdown')]

    # ── Soil materials (before structural elements) ──
    _set_progress(wb, f"sc5: {_spin()} WORKING — soil materials {_bar(0.55)}")
    soil_cells = gen_soil_profile_code(soil_data, is_3d=is_3d)
    if soil_cells:
        new_cells.extend(soil_cells)

    def gen_shared_section(data, gen_mat_fn, gen_geom_fn, mat_key='cad_layer',
                           inherit=None, embedded_beam_counter=None,
                           plate_material_counter=None):
        """One material per unique DXF layer. Children with custom values
        get their own material. Embedded-beam material colors use one shared
        counter across all callers (nails/anchors and piles)."""
        seen_layers = {}  # layer_name → (material_index, parent_data)
        mat_counter = 0
        geom_cells = []

        for i, d in enumerate(data, 1):
            layer = safe_str(d.get(mat_key), '').lower()
            if not layer:
                continue

            if layer not in seen_layers:
                # First element on this layer → create material
                mat_counter += 1
                seen_layers[layer] = (mat_counter, d)
                mat_idx = mat_counter
                if embedded_beam_counter is not None:
                    mat_lines = gen_mat_fn(
                        mat_idx, d,
                        embedded_beam_index=embedded_beam_counter[0],
                    )
                    embedded_beam_counter[0] += 1
                elif plate_material_counter is not None:
                    mat_lines = gen_mat_fn(
                        mat_idx, d,
                        plate_material_index=plate_material_counter[0],
                    )
                    plate_material_counter[0] += 1
                else:
                    mat_lines = gen_mat_fn(mat_idx, d)
                new_cells.append(make_cell(mat_lines))
            else:
                # Child: check if inheritable values match parent
                parent_mat_idx, parent_d = seen_layers[layer]
                if inherit:
                    def _norm(v):
                        """Normalize value for comparison: treat N/A variants as equivalent."""
                        return safe_str(v)
                    values_match = all(
                        _norm(d.get(h)) == _norm(parent_d.get(h))
                        for h in inherit
                    )
                else:
                    values_match = True
                if values_match:
                    mat_idx = parent_mat_idx
                else:
                    # Custom values → own material (but keep original parent as reference)
                    mat_counter += 1
                    mat_idx = mat_counter
                    if embedded_beam_counter is not None:
                        mat_lines = gen_mat_fn(
                            mat_idx, d,
                            embedded_beam_index=embedded_beam_counter[0],
                        )
                        embedded_beam_counter[0] += 1
                    else:
                        mat_lines = gen_mat_fn(mat_idx, d)
                    new_cells.append(make_cell(mat_lines))

            # Geometry for each element, referencing its material
            # Pass mat_idx to geometry function so it generates correct material reference
            geom_lines = gen_geom_fn(i, d, mat_idx=mat_idx)
            geom_cells.append(make_cell(geom_lines))

        new_cells.extend(geom_cells)

    _set_progress(wb, f"sc5: {_spin()} WORKING — plates {_bar(0.60)}")
    # Plate materials use their own deterministic palette counter.
    plate_material_counter = [0]
    gen_shared_section(plate_data, gen_plate_mat, gen_plate_geom,
                       inherit=pl_inherit,
                       plate_material_counter=plate_material_counter)

    # N2N
    gen_shared_section(n2n_data, gen_n2n_mat, gen_n2n_geom,
                       inherit=nn_inherit)

    _set_progress(wb, f"sc5: {_spin()} WORKING — anchors/struts {_bar(0.70)}")
    # Embedded Beam and piles share one palette counter.
    embedded_beam_counter = [0]
    gen_shared_section(anc_data, gen_anc_mat, gen_anc_geom,
                       inherit=anc_inherit,
                       embedded_beam_counter=embedded_beam_counter)

    # Strut
    gen_shared_section(strut_data, gen_strut_mat, gen_strut_geom,
                       inherit=nn_inherit)

    # FEA (fixed end anchor)
    gen_shared_section(fea_data, gen_fea_mat, gen_fea_geom,
                       inherit=fea_inherit)

    # Pile
    gen_shared_section(pile_data, gen_pile_mat, gen_pile_geom,
                       inherit=pil_inherit,
                       embedded_beam_counter=embedded_beam_counter)

    # Geogrid: one material per cad_layer, one geometry per DXF LINE.
    gen_shared_section(geogrid_data, gen_geogrid_mat, gen_geogrid_geom,
                       inherit=geogrid_inherit)

    # ── Interfaces (AFTER all geometry, adjacent only) ──
    has_interfaces = False
    for i, d in enumerate(plate_data, 1):
        iface = gen_plate_interface(i, d)
        if iface:
            if not has_interfaces:
                new_cells.append(make_cell('# ── Plate Interfaces (after all geometry) ──'))
                has_interfaces = True
            new_cells.append(make_cell(iface))

    _set_progress(wb, f"sc5: {_spin()} WORKING — tunnels {_bar(0.78)}")
    # ── Tunnel generation (circular) ──
    # Tunnel designer labels are Tunnel_N. Plate liners use the same global
    # Plate_N numbering as wall/slab plates; volume liners use the next
    # Polygon_N label after the soil polygons already generated by SC4.
    # The API order is: plate/interface setup → thick lining → generatetunnel.
    if tunnel_data:
        new_cells.append(make_cell('# ── Tunnel elements (from sc5) ──'))

        # Group rows by full DXF layer; rows are primary then secondary.
        _tun_groups = {}
        for _d in tunnel_data:
            _key = safe_str(_d.get('cad_layer')).lower()
            if _key:
                _tun_groups.setdefault(_key, []).append(_d)

        _tun_no = 0
        _tunnel_plate_counter = len(plate_data)
        _tunnel_poly_counter = len(soil_polys) if soil_polys else 0

        for _cad_key, _rows in _tun_groups.items():
            _tun_no += 1
            _d0 = _rows[0]
            _d_primary = next(
                (_row for _row in _rows
                 if safe_str(_row.get('liner_position')).lower() == 'primary'),
                _d0,
            )
            _x = safe_num(_d0.get('x1'))
            _y = safe_num(_d0.get('y1'))
            _r_dxf = safe_num(_d0.get('y2'))
            _c_ref = safe_num(_d_primary.get('C_ref'))

            # Classify model types for this tunnel
            _positions = {}  # position → model_type
            for _row in _rows:
                _p = safe_str(_row.get('liner_position')).lower()
                _m = safe_str(_row.get('liner_model')).lower()
                if 'plate' in _m or 'eplastic' in _m:
                    _positions[_p] = 'plate'
                elif 'str_volume' in _m:
                    _positions[_p] = 'str_volume'
            _primary_type = _positions.get('primary', 'str_volume')
            _secondary_type = _positions.get('secondary')
            _primary_thk = safe_num(_d0.get('plate_THK(m)'))
            for _row in _rows:
                if safe_str(_row.get('liner_position')).lower() == 'primary':
                    _primary_thk = safe_num(_row.get('plate_THK(m)'))
            _secondary_thk = 0.0
            for _row in _rows:
                if safe_str(_row.get('liner_position')).lower() == 'secondary':
                    _secondary_thk = safe_num(_row.get('plate_THK(m)'))

            # ── TBM mode: detect grout position in group ──
            _is_tbm = any(safe_str(_r.get('liner_position')).lower() == 'grout'
                          for _r in _rows)

            if _is_tbm:
                # ── TBM geometry (v0.7.1) ──
                # Three-liner: plate (shield) + str_volume (grout) + plate/volume (main).
                # R_base = R_DXF - t_shield/2 (DXF circle = shield centerline).
                # Offsets: cumulative negative inward. Ring capture: palindrome sandwich.
                _d_grout = next((_r for _r in _rows
                                 if safe_str(_r.get('liner_position')).lower() == 'grout'), None)
                _d_secondary = next((_r for _r in _rows
                                     if safe_str(_r.get('liner_position')).lower() == 'secondary'), None)

                _t_shield = safe_num(_d_primary.get('plate_THK(m)'))
                _t_grout = safe_num(_d_grout.get('plate_THK(m)')) if _d_grout else 0.0
                _t_main = safe_num(_d_secondary.get('plate_THK(m)')) if _d_secondary else 0.0

                # Config B: main=str_volume → 2 offsets; Config A: main=plate → 1 offset
                _sec_model = safe_str(_d_secondary.get('liner_model')).lower() if _d_secondary else ''
                _is_config_b = 'str_volume' in _sec_model

                _r_base = _r_dxf - _t_shield / 2.0
                _y_origin = _y - _r_base
                _tunnel_var = f'tunnel_{_tun_no}'

                new_cells.append(make_cell([
                    f'{_tunnel_var} = g_i.tunnel({round_num(_x)}, {round_num(_y_origin)})',
                    f'{_tunnel_var}.CrossSection.setproperties("ShapeType", "Circular")',
                    f'{_tunnel_var}.CrossSection.Segments[0].ArcProperties.Radius = {round_num(_r_base)}',
                ]))

                # Offset 1: grout annulus (negative inward)
                new_cells.append(make_cell([
                    f'g_i.generatethicklining({_tunnel_var}.CrossSection, {round_num(-_t_grout)})',
                ]))

                # Offset 2: main liner (Config B only, cumulative)
                if _is_config_b:
                    new_cells.append(make_cell([
                        f'g_i.generatethicklining({_tunnel_var}.CrossSection, {round_num(-(_t_grout + _t_main))})',
                    ]))

                # Palindrome ring capture
                _s_var = f'_s_{_tun_no}'
                _n_var = f'_n_{_tun_no}'
                new_cells.append(make_cell([
                    f'{_s_var} = {_tunnel_var}.SliceSegments[:]',
                    f'{_n_var} = len({_s_var})',
                ]))

                _shield_var = f'_shield_{_tun_no}'
                _grout_var = f'_grout_{_tun_no}'
                new_cells.append(make_cell([
                    f'{_shield_var} = [{_s_var}[0], {_s_var}[{_n_var} - 1]]',
                    f'{_grout_var} = [{_s_var}[1], {_s_var}[{_n_var} - 2]]',
                ]))

                if _is_config_b:
                    _main_var = f'_main_{_tun_no}'
                    new_cells.append(make_cell([
                        f'{_main_var} = [{_s_var}[2], {_s_var}[{_n_var} - 3]]',
                    ]))

                # ── Shield plate material ──
                _shield_mat_var = f'tunnel_mat_{_tun_no}_primary'
                _shield_ea1 = safe_num(_d_primary.get('plate_EA1'))
                _shield_ei = safe_num(_d_primary.get('plate_EI'))
                _shield_w = safe_num(_d_primary.get('plate_w(kN/m/m)'))
                _shield_nu = safe_num(_d_primary.get('v_nu'), 0.15)
                _shield_punch = 'True' if safe_str(_d_primary.get('prevent_punching')).lower() == 'yes' else 'False'
                _shield_mat_type = 2 if 'eplastic' in safe_str(_d_primary.get('liner_model')).lower() else 1
                _shield_props = [
                    '"Identification"', repr(safe_str(_d_primary.get('tunnel_name'), f'tunnel_{_tun_no}_shield')),
                    '"MaterialType"', str(_shield_mat_type),
                    '"EA1"', round_num(_shield_ea1),
                    '"EI"', round_num(_shield_ei),
                    '"w"', round_num(_shield_w),
                    '"PreventPunching"', _shield_punch,
                ]
                _shield_iso = _normalize_isotropic(_d_primary.get('isotropic'))
                if _shield_iso == 'yes':
                    _shield_props.extend(['"StructNu"', round_num(_shield_nu)])
                else:
                    _se_mpa = safe_num(_d_primary.get('plate_E(Mpa)'))
                    _sa2 = safe_num(_d_primary.get('plate_A2(m2)'))
                    _sea2 = _se_mpa * 1000 * _sa2 if _se_mpa > 0 and _sa2 > 0 else 0
                    _shield_props.extend(['"isotropic"', 'False', '"EA2"', round_num(_sea2)])
                if _shield_mat_type == 2:
                    _smp = safe_str(_d_primary.get('Mp(kNm/m)'))
                    _snp1 = safe_str(_first_present(_d_primary, 'Np1(kN/m)', 'Np1tens(kN/m)'))
                    _snp2 = safe_str(_first_present(_d_primary, 'Np2(kN/m)', 'Np2tens(kN/m)'))
                    if safe_str(_smp) not in ('', 'N/A'):
                        _shield_props.extend(['"Mp"', round_num(safe_num(_smp))])
                    if safe_str(_snp1) not in ('', 'N/A'):
                        _shield_props.extend(['"Np1Tens"', round_num(safe_num(_snp1))])
                    if safe_str(_snp2) not in ('', 'N/A'):
                        _shield_props.extend(['"Np2Tens"', round_num(safe_num(_snp2))])
                new_cells.append(make_cell([
                    f'{_shield_mat_var} = g_i.platemat()',
                    f'{_shield_mat_var}.setproperties({", ".join(_shield_props)})',
                ]))

                # Plate on shield ring (base circle)
                _shield_plate_var = f'tunnel_plate_{_tun_no}_primary'
                new_cells.append(make_cell([
                    f'{_shield_plate_var} = []',
                    f'for _slice in {_shield_var}:',
                    f'    g_i.plate(_slice)',
                    f'    _slice.Plate.Material = {_shield_mat_var}',
                    f'    {_shield_plate_var}.append(_slice)',
                ]))
                _siface = safe_str(_d_primary.get('interface'), 'none').lower()
                if _siface in ('positive', 'both'):
                    new_cells.append(make_cell([
                        f'for _slice in {_shield_plate_var}:',
                        f'    g_i.posinterface(_slice)',
                    ]))
                if _siface in ('negative', 'both'):
                    new_cells.append(make_cell([
                        f'for _slice in {_shield_plate_var}:',
                        f'    g_i.neginterface(_slice)',
                    ]))

                # ── Main plate (Config A only) ──
                # Row is 'secondary' → vars follow the _secondary convention
                if not _is_config_b and _d_secondary:
                    _mmat_var = f'tunnel_mat_{_tun_no}_secondary'
                    _mmat_type = 2 if 'eplastic' in safe_str(_d_secondary.get('liner_model')).lower() else 1
                    _mea1 = safe_num(_d_secondary.get('plate_EA1'))
                    _mei = safe_num(_d_secondary.get('plate_EI'))
                    _mw = safe_num(_d_secondary.get('plate_w(kN/m/m)'))
                    _mnu = safe_num(_d_secondary.get('v_nu'), 0.15)
                    _mpunch = 'True' if safe_str(_d_secondary.get('prevent_punching')).lower() == 'yes' else 'False'
                    _mprops = [
                        '"Identification"', repr(safe_str(_d_secondary.get('tunnel_name'), f'tunnel_{_tun_no}_main')),
                        '"MaterialType"', str(_mmat_type),
                        '"EA1"', round_num(_mea1),
                        '"EI"', round_num(_mei),
                        '"w"', round_num(_mw),
                        '"PreventPunching"', _mpunch,
                    ]
                    _miso = _normalize_isotropic(_d_secondary.get('isotropic'))
                    if _miso == 'yes':
                        _mprops.extend(['"StructNu"', round_num(_mnu)])
                    else:
                        _emme = safe_num(_d_secondary.get('plate_E(Mpa)'))
                        _ma2 = safe_num(_d_secondary.get('plate_A2(m2)'))
                        _mea2 = _emme * 1000 * _ma2 if _emme > 0 and _ma2 > 0 else 0
                        _mprops.extend(['"isotropic"', 'False', '"EA2"', round_num(_mea2)])
                    if _mmat_type == 2:
                        _mmp = safe_str(_d_secondary.get('Mp(kNm/m)'))
                        _mnp1 = safe_str(_first_present(_d_secondary, 'Np1(kN/m)', 'Np1tens(kN/m)'))
                        _mnp2 = safe_str(_first_present(_d_secondary, 'Np2(kN/m)', 'Np2tens(kN/m)'))
                        if safe_str(_mmp) not in ('', 'N/A'):
                            _mprops.extend(['"Mp"', round_num(safe_num(_mmp))])
                        if safe_str(_mnp1) not in ('', 'N/A'):
                            _mprops.extend(['"Np1Tens"', round_num(safe_num(_mnp1))])
                        if safe_str(_mnp2) not in ('', 'N/A'):
                            _mprops.extend(['"Np2Tens"', round_num(safe_num(_mnp2))])
                    new_cells.append(make_cell([
                        f'{_mmat_var} = g_i.platemat()',
                        f'{_mmat_var}.setproperties({", ".join(_mprops)})',
                    ]))
                    # Plate on grout ring (Config A main liner)
                    _mplate_var = f'tunnel_plate_{_tun_no}_main'
                    new_cells.append(make_cell([
                        f'{_mplate_var} = []',
                        f'for _slice in {_grout_var}:',
                        f'    g_i.plate(_slice)',
                        f'    _slice.Plate.Material = {_mmat_var}',
                        f'    {_mplate_var}.append(_slice)',
                    ]))
                    _miface = safe_str(_d_secondary.get('interface'), 'none').lower()
                    if _miface in ('positive', 'both'):
                        new_cells.append(make_cell([
                            f'for _slice in {_mplate_var}:',
                            f'    g_i.posinterface(_slice)',
                        ]))
                    if _miface in ('negative', 'both'):
                        new_cells.append(make_cell([
                            f'for _slice in {_mplate_var}:',
                            f'    g_i.neginterface(_slice)',
                        ]))

                # ── Grout soil material (always) ──
                # Variable name must match assign_sequence (always cleaned)
                _gmat_var = _clean_mat_var_name(
                    f"tunnel{_tun_no}_grout_THK={_t_grout}m" if _t_grout > 0
                    else f"tunnel{_tun_no}_grout_THK=?"
                )
                _ge = safe_num(_d_grout.get('plate_E(Mpa)')) if _d_grout else 0
                _gnu = safe_num(_d_grout.get('v_nu')) if _d_grout else 0
                # v0.7.20: grout unit weight (kN/m3) from plate_w, straight through.
                _gw = safe_num(_d_grout.get('plate_w(kN/m/m)')) if _d_grout else 0
                _gname = safe_str(_d_grout.get('tunnel_name'), f'tunnel_{_tun_no}_grout') if _d_grout else f'tunnel_{_tun_no}_grout'
                new_cells.append(make_cell([
                    f'{_gmat_var} = g_i.soilmat()',
                    f'{_gmat_var}.setproperties("Identification", {repr(_gname)}, "SoilModel", "Linear elastic", "DrainageType", "Non-porous", "gammaUnsat", {round_num(_gw)}, "ERef", {round_num(_ge * 1000)}, "nu", {round_num(_gnu)})',
                ]))

                # ── Main soil material (Config B only) ──
                # Variable name must match assign_sequence: position is
                # 'secondary', so the material var is tunnel{N}_secondary_THK=…
                if _is_config_b and _d_secondary:
                    _mmat_soil = _clean_mat_var_name(
                        f"tunnel{_tun_no}_secondary_THK={_t_main}m" if _t_main > 0
                        else f"tunnel{_tun_no}_secondary_THK=?"
                    )
                    _mme = safe_num(_d_secondary.get('plate_E(Mpa)'))
                    _mmnu = safe_num(_d_secondary.get('v_nu'))
                    # v0.7.20: Config B main unit weight (kN/m3) from plate_w.
                    _mmw = safe_num(_d_secondary.get('plate_w(kN/m/m)'))
                    _mmname = safe_str(_d_secondary.get('tunnel_name'), f'tunnel_{_tun_no}_main')
                    new_cells.append(make_cell([
                        f'{_mmat_soil} = g_i.soilmat()',
                        f'{_mmat_soil}.setproperties("Identification", {repr(_mmname)}, "SoilModel", "Linear elastic", "DrainageType", "Non-porous", "gammaUnsat", {round_num(_mmw)}, "ERef", {round_num(_mme * 1000)}, "nu", {round_num(_mmnu)})',
                    ]))

                # ── Contraction on shield only ──
                if _c_ref > 0:
                    new_cells.append(make_cell([
                        f'for _slice in {_shield_var}:',
                        f'    g_i.contraction(_slice, "C", {round_num(_c_ref)})',
                    ]))

                new_cells.append(make_cell(f'g_i.generatetunnel({_tunnel_var})'))
                continue  # skip existing single/double liner code

            # DXF radius = outer surface of primary liner.
            # Base circle depends on liner combination:
            #   primary=plate (incl. double plate) → R_DXF - t_primary/2
            #   primary=str_volume + secondary=plate → R_DXF - t_primary
            #   double str_volume  → R_DXF (outer boundary)
            # Double plate = Method 1 (primary plate + secondary str_volume)
            # plus an extra plate on the thick-lining output.
            _double_plate = (_primary_type == 'plate'
                             and _secondary_type == 'plate')
            _mixed_str_plate = (_primary_type == 'str_volume'
                                and _secondary_type == 'plate')
            if _primary_type == 'plate':
                _r_base = _r_dxf - _primary_thk / 2.0
            elif _mixed_str_plate:
                _r_base = _r_dxf - _primary_thk
            else:
                _r_base = _r_dxf
            _tunnel_var = f'tunnel_{_tun_no}'
            _y_origin = _y - _r_base
            new_cells.append(make_cell([
                f'{_tunnel_var} = g_i.tunnel({round_num(_x)}, {round_num(_y_origin)})',
                f'{_tunnel_var}.CrossSection.setproperties("ShapeType", "Circular")',
                f'{_tunnel_var}.CrossSection.Segments[0].ArcProperties.Radius = {round_num(_r_base)}',
            ]))

            # Sort: plate always first (must be on SliceSegments before
            # generatethicklining). Among str_volume: primary first for
            # cumulative offset.
            def _liner_sort_key(_row):
                _p = safe_str(_row.get('liner_position')).lower()
                _m = safe_str(_row.get('liner_model')).lower()
                _is_plate = 'plate' in _m or 'eplastic' in _m
                if _is_plate:
                    return -1
                return 0 if _p == 'primary' else 1

            _rows = sorted(_rows, key=_liner_sort_key)

            _cumulative_thk = 0.0
            _thicklining_calls = []
            for _d in _rows:
                _pos = safe_str(_d.get('liner_position'), 'primary').lower()
                _model = safe_str(_d.get('liner_model')).lower()
                if 'eplastic' in _model:
                    _model_key = 'plate_eplastic'
                elif 'plate' in _model:
                    _model_key = 'plate_elastic'
                elif 'str_volume' in _model:
                    _model_key = 'str_volume'
                else:
                    _model_key = _model

                if _model_key in ('plate_elastic', 'plate_eplastic'):
                    _tunnel_plate_counter += 1
                    _idx = _tunnel_plate_counter
                    _mat_var = f'tunnel_mat_{_tun_no}_{_pos}'
                    _plate_var = f'tunnel_plate_{_tun_no}_{_pos}'
                    _material_type = 2 if _model_key == 'plate_eplastic' else 1
                    _ea1 = safe_num(_d.get('plate_EA1'))
                    _ei = safe_num(_d.get('plate_EI'))
                    _w = safe_num(_d.get('plate_w(kN/m/m)'))
                    _nu = safe_num(_d.get('v_nu'), 0.15)
                    _punch = 'True' if safe_str(_d.get('prevent_punching')).lower() == 'yes' else 'False'
                    _props = [
                        '"Identification"', repr(safe_str(_d.get('tunnel_name'), f'tunnel_{_pos}')),
                        '"MaterialType"', str(_material_type),
                        '"EA1"', round_num(_ea1),
                        '"EI"', round_num(_ei),
                        '"w"', round_num(_w),
                        '"PreventPunching"', _punch,
                    ]
                    # Isotropic → add StructNu; Anisotropic → add isotropic=False + EA2
                    _iso = _normalize_isotropic(_d.get('isotropic'))
                    if _iso == 'yes':
                        _props.extend(['"StructNu"', round_num(_nu)])
                    else:
                        _e_mpa = safe_num(_d.get('plate_E(Mpa)'))
                        _a2 = safe_num(_d.get('plate_A2(m2)'))
                        _ea2 = _e_mpa * 1000 * _a2 if _e_mpa > 0 and _a2 > 0 else 0
                        _props.extend(['"isotropic"', 'False', '"EA2"', round_num(_ea2)])
                    if _model_key == 'plate_eplastic':
                        # Only add Mp/Np1/Np2 if values are present (not blank/N/A)
                        _mp_v = safe_str(_d.get('Mp(kNm/m)'))
                        _np1_v = safe_str(_first_present(_d, 'Np1(kN/m)', 'Np1tens(kN/m)'))
                        _np2_v = safe_str(_first_present(_d, 'Np2(kN/m)', 'Np2tens(kN/m)'))
                        if safe_str(_mp_v) not in ('', 'N/A'):
                            _props.extend(['"Mp"', round_num(safe_num(_mp_v))])
                        if safe_str(_np1_v) not in ('', 'N/A'):
                            _props.extend(['"Np1Tens"', round_num(safe_num(_np1_v))])
                        if safe_str(_np2_v) not in ('', 'N/A'):
                            _props.extend(['"Np2Tens"', round_num(safe_num(_np2_v))])
                    new_cells.append(make_cell([
                        f'{_mat_var} = g_i.platemat()',
                        f'{_mat_var}.setproperties({", ".join(_props)})',
                    ]))
                    # Geometry emission:
                    #   double-plate secondary → deferred to Phase B
                    #   (after thick lining, on the grown slices)
                    #   everything else → base circle SliceSegments[:]
                    if _double_plate and _pos == 'secondary':
                        pass  # geometry emitted in the double-plate block below
                    else:
                        # Use each tunnel slice as one generated arc plate.
                        # Material and interfaces are assigned on the slice,
                        # NOT the plate proxy returned by g_i.plate().
                        new_cells.append(make_cell([
                            f'{_plate_var} = []',
                            f'for _slice in {_tunnel_var}.SliceSegments[:]:',
                            f'    g_i.plate(_slice)',
                            f'    _slice.Plate.Material = {_mat_var}',
                            f'    {_plate_var}.append(_slice)',
                        ]))
                        _iface = safe_str(_d.get('interface'), 'none').lower()
                        if _iface in ('positive', 'both'):
                            new_cells.append(make_cell([
                                f'for _slice in {_plate_var}:',
                                f'    g_i.posinterface(_slice)',
                            ]))
                        if _iface in ('negative', 'both'):
                            new_cells.append(make_cell([
                                f'for _slice in {_plate_var}:',
                                f'    g_i.neginterface(_slice)',
                            ]))

                elif _model_key == 'str_volume':
                    _tunnel_poly_counter += 1
                    _mat_name = f"tunnel{_tun_no}_{_pos}_THK={safe_num(_d.get('plate_THK(m)'))}m" if safe_num(_d.get('plate_THK(m)')) > 0 else f"tunnel{_tun_no}_{_pos}_THK=?"
                    _mat_var = _clean_mat_var_name(_mat_name)
                    _thk = safe_num(_d.get('plate_THK(m)'))
                    _e = safe_num(_d.get('plate_E(Mpa)'))
                    _nu = safe_num(_d.get('v_nu'))
                    # v0.7.20: plate_w on str_volume rows is the soil unit
                    # weight (kN/m3), passed straight to gammaUnsat only
                    # (Non-porous: gammaSat is read-only in PLAXIS).
                    _w = safe_num(_d.get('plate_w(kN/m/m)'))
                    # Offset direction depends on combination:
                    #   mixed (str_volume + plate): positive outward from base
                    #   double str_volume: negative cumulative inward from base
                    if _mixed_str_plate:
                        _cumulative_thk += _thk
                        _offset = _cumulative_thk
                    else:
                        _cumulative_thk += _thk
                        _offset = -_cumulative_thk
                    new_cells.append(make_cell([
                        f'{_mat_var} = g_i.soilmat()',
                        f'{_mat_var}.setproperties("Identification", {repr(safe_str(_d.get("tunnel_name"), f"tunnel_{_pos}"))}, "SoilModel", "Linear elastic", "DrainageType", "Non-porous", "gammaUnsat", {round_num(_w)}, "ERef", {round_num(_e * 1000)}, "nu", {round_num(_nu)})',
                    ]))
                    _thicklining_calls.append(
                        f'g_i.generatethicklining({_tunnel_var}.CrossSection, {round_num(_offset)})'
                    )

            # Apply contraction after all plate/interface setup and before
            # thick lining or tunnel generation.
            if _c_ref > 0:
                new_cells.append(make_cell(gen_tunnel_contraction_code(
                    _tunnel_var, _c_ref)))

            # For double-plate: each generatethicklining call appends new
            # SliceSegments (base circle = [0],[1]; each ring adds its own
            # segments).  Capture the pre-call count, run thick lining
            # inward, then plate the newly appended slices — all BEFORE
            # generatetunnel.
            if _double_plate and _secondary_thk > 0:
                _tl_var = f'thicklining_{_tun_no}'
                _n0_var = f'_n0_{_tun_no}'
                _sec_plate_var = f'tunnel_plate_{_tun_no}_secondary'
                _sec_mat_var = f'tunnel_mat_{_tun_no}_secondary'
                _sec_iface = 'none'
                for _d in _rows:
                    if safe_str(_d.get('liner_position')).lower() == 'secondary':
                        _sec_iface = safe_str(_d.get('interface'), 'none').lower()
                new_cells.append(make_cell([
                    f'{_n0_var} = len({_tunnel_var}.SliceSegments[:])',
                    f'{_tl_var} = g_i.generatethicklining('
                    f'{_tunnel_var}.CrossSection, {round_num(-_secondary_thk)})',
                    f'{_sec_plate_var} = []',
                    f'for _slice in {_tunnel_var}.SliceSegments[{_n0_var}:]:',
                    f'    g_i.plate(_slice)',
                    f'    _slice.Plate.Material = {_sec_mat_var}',
                    f'    {_sec_plate_var}.append(_slice)',
                ]))
                if _sec_iface in ('positive', 'both'):
                    new_cells.append(make_cell([
                        f'for _slice in {_sec_plate_var}:',
                        f'    g_i.posinterface(_slice)',
                    ]))
                if _sec_iface in ('negative', 'both'):
                    new_cells.append(make_cell([
                        f'for _slice in {_sec_plate_var}:',
                        f'    g_i.neginterface(_slice)',
                    ]))
            elif _thicklining_calls:
                new_cells.append(make_cell(_thicklining_calls))

            if _double_plate and _secondary_thk <= 0:
                new_cells.append(make_cell([
                    '# WARNING: double-plate liner but secondary plate_THK is '
                    'blank/0 — secondary plate not created. '
                    'Fill plate_THK(m) on the secondary row.',
                ]))

            # Generate the tunnel only after all liner setup.
            new_cells.append(make_cell(f'g_i.generatetunnel({_tunnel_var})'))

    # Geogrid interfaces are created after ordinary plates and tunnel liners,
    # matching the global interface counters used by assign_sequence.
    if geogrid_data:
        new_cells.append(make_cell('# ── Geogrid Interfaces (after tunnel geometry) ──'))
        for i, d in enumerate(geogrid_data, 1):
            interface_lines = gen_geogrid_interface(i, d)
            if interface_lines:
                new_cells.append(make_cell(interface_lines))
    # ── Volume-profile interfaces (after geogrid interfaces) ──
    if soil_polys:
        volume_interface_records, _, _ = _plan_volume_profile_interfaces(
            soil_polys, 0, 0,
        )
        if volume_interface_records:
            new_cells.append(make_cell(
                '# ── Volume-profile Interfaces (after geogrid interfaces) ──'))
            new_cells.append(make_cell(gen_vol_interface_code(
                volume_interface_records)))

    if soil_polys:
        soil_cells = gen_soil_polygon_code(soil_polys)
        new_cells.extend(soil_cells)

    _set_progress(wb, f"sc5: {_spin()} WORKING — drains/loads {_bar(0.85)}")
    # ── PVD Drains (v0.6.1) ──
    drains = read_drains(sheet)
    if drains:
        drain_cells = gen_drain_code(drains)
        new_cells.extend(drain_cells)

    # ── Line Loads (v0.6.9) ──
    loads = read_loads(sheet)

    # Auto-generate load names and write back to column B + load_no to column H
    LOAD_DATA_END = 38
    LOAD_DATA_START = 19
    _load_rows_written = set()
    _line_counter = 0
    _point_counter = 0
    for i, ld in enumerate(loads, 1):
        load_model = safe_str(ld.get('load_model'), 'line_load')
        if load_model == 'point_load':
            _point_counter += 1
            auto_name = f"point_load{_point_counter}"
            label = f'PointLoad_{_point_counter}'
        else:
            _line_counter += 1
            magnitude = abs(int(ld['qy_start'])) if ld['qy_start'] != 0 else 0
            auto_name = f"line_load{_line_counter}_{magnitude}kPa"
            label = f'LineLoad_{_line_counter}'
        write_orange(sheet, ld['_row'], {'Loads': 2}, 'Loads', auto_name)
        write_orange(sheet, ld['_row'], {'load_no': 8}, 'load_no', label)
        ld['name'] = auto_name
        _load_rows_written.add(ld['_row'])

    # Clear leftover load rows (B + H) that are no longer used
    for r in range(LOAD_DATA_START, LOAD_DATA_END + 1):
        if r not in _load_rows_written:
            write_orange(sheet, r, {'Loads': 2}, 'Loads', None)
            write_orange(sheet, r, {'load_no': 8}, 'load_no', None)

    if loads:
        line_loads = [ld for ld in loads if safe_str(ld.get('load_model'), 'line_load') == 'line_load']
        point_loads = [ld for ld in loads if safe_str(ld.get('load_model'), 'line_load') == 'point_load']
        if line_loads:
            load_cells = gen_load_code(line_loads)
            new_cells.extend(load_cells)
        if point_loads:
            point_cells = gen_pointload_code(point_loads)
            new_cells.extend(point_cells)

    # ── Groundwater Flow Boundaries (v0.7.9) ──
    # Behaviour is derived from cad_layer; Head Href follows boundary elevation.
    # Creation remains in structural mode; water levels switch to flow mode below.
    gwflowbcs = read_gwflowbcs(sheet)
    if gwflowbcs:
        new_cells.extend(gen_gwflowbc_code(gwflowbcs))

    # ── Merge equivalents (v0.7.15) ──
    # Collapse near-duplicate geometry points/lines from DXF imports
    # (command line: _mergeequivalents Geometry, tolerance 0.001 → Python
    #  g_i.mergeequivalents(g_i.Geometry) per KB0107777 translation rule:
    #  commands take the g_i. prefix, object args take g_i. object refs).
    # MUST run in structures mode BEFORE the first gotostages(): in staged
    # mode the server does not expose geometry commands (AttributeError:
    # 'mergeequivalents' is not present). Hence placed before the
    # water-level gotoflow/gotostages sandwich. SC6's cleanup skips this
    # tag so the merge cell survives staged-construction reruns.
    new_cells.append(make_cell(MERGE_TAG, cell_type='markdown'))
    new_cells.append(make_cell([
        '# Merge duplicate geometry (command: _mergeequivalents Geometry)',
        'g_i.gotostructures()',
        'print("sc5: merging equivalent geometry…")',
        'try:',
        '    g_i.mergeequivalents(g_i.Geometry)',
        'except AttributeError:',
        '    try:',
        '        g_i.mergeequivalents("Geometry")',
        '    except AttributeError:',
        '        try:',
        '            g_i.mereq(g_i.Geometry)',
        '        except AttributeError:',
        '            g_i.mereq("Geometry")',
        'print("sc5: merge equivalents done")',
    ]))

    _set_progress(wb, f"sc5: {_spin()} WORKING — merge + water {_bar(0.90)}")
    # ── Water Levels (v0.7.6) — LAST in gotoflow/gotostages sandwich ──
    # Water levels must be created after all other structural elements
    # (plates, anchors, drains, loads) because gotostages() switches mode
    # and subsequent structural creation commands would fail.
    waterlines = read_waterlines(sheet)
    if waterlines:
        wl_cells = gen_waterlevel_code(waterlines)
        new_cells.extend(wl_cells)

    _set_progress(wb, f"sc5: {_spin()} WORKING — writing tables {_bar(0.97)}")
    # ── Write assign_sequence tab (xlwings — preserves VBA) ──
    _write_assign_sequence(sheet, plate_data, n2n_data, anc_data, strut_data, pile_data,
                           soil_polys=soil_polys, drains=drains, loads=loads, fea_data=fea_data,
                           tunnel_data=tunnel_data, geogrid_data=geogrid_data,
                           waterlines=waterlines, gwflowbcs=gwflowbcs)

    # ── Output forces table (v0.7.11) ──
    # Built AFTER assign_sequence so it reads the freshly generated rows and
    # phase headers.  It never touches the SC4-owned B9:F108 curve block.
    try:
        _write_output_forces_table(wb)
    except Exception as exc:
        print(f"sc5: WARNING: output forces table not updated: {exc}")

    # ── Tunnel identification (BBox scan after gotostages) ──
    # v0.6.16: Generate the identification code cell. SC6 will relocate
    # it after gotostages() since SoilPolygons only exist in staged mode.
    if tunnel_data:
        tun_id_cells = _gen_tunnel_id_cells(tunnel_data)
        new_cells.extend(tun_id_cells)

    # ── Mesh + Output curve points (v0.7.10) ──
    # This block deliberately comes after every structural element and water
    # creation cell.  SC9 executes it before SC6's staged-construction cells;
    # no calculation is emitted here — the user calculates manually in PLAXIS.
    mesh_generation = read_mesh_generation(sheet)
    if mesh_generation:
        mesh_density = read_mesh_density(sheet)
        # ── Local mesh refinement (str_2D col-C [refine f], v0.8.5) ──
        # Factors ride between gotomesh() and mesh() so they bite at the
        # next Generate.  Probe-proven route: mesh-mode tag census over
        # g_i.Polygons + single-level setproperties('CoarsenessFactor', f);
        # nested attribute writes are blind to the mesher.  Gated by the
        # same K3 switch as mesh generation — excluded means no mesh, so
        # no refinement cells either.  Echo-back writes [refine f] to
        # column C for the tagged polygon_no groups only (idempotent).
        refine_list = collect_refinements(soil_polys)
        mesh_lines = gen_mesh_code(mesh_density)
        if refine_list:
            mesh_lines = (
                [mesh_lines[0], mesh_lines[1]]
                + gen_refine_code(refine_list)
                + [mesh_lines[2]]
            )
            _write_refine_echo(sheet, soil_polys, refine_list)
            for _p_no, _f in refine_list:
                print(f"sc5: Polygon_{_p_no}: CoarsenessFactor {_f:g} "
                      f"(from str_2D col C [refine {_f:g}])")
        new_cells.append(make_cell(mesh_lines))

        curve_points = []
        try:
            curve_sheet = wb.sheets[CURVE_OUTPUT_SHEET]
            curve_points = read_curve_points(curve_sheet)
        except Exception as exc:
            print(f"sc5: WARNING: cannot read '{CURVE_OUTPUT_SHEET}' curve table: {exc}")

        curve_lines = gen_curvepoint_code(curve_points)
        if curve_lines:
            new_cells.append(make_cell(curve_lines))
    else:
        print(
            f"sc5: mesh + curve generation skipped by "
            f"str_2D!{MESH_GENERATION_CELL}"
        )
        _dropped = collect_refinements(soil_polys)
        if _dropped:
            print(
                f"sc5: WARNING: str_2D!{MESH_GENERATION_CELL} is excluded — "
                f"{len(_dropped)} [refine] tag(s) "
                f"({', '.join(f'Polygon_{p} {_f:g}' for p, _f in _dropped)}) "
                f"silently dropped with mesh generation; set included to apply them"
            )

    notebook['cells'].extend(new_cells)

    # ── Write notebook ──
    with open(nb_path, 'w', encoding='utf-8') as f:
        json.dump(notebook, f, indent=1, ensure_ascii=False)

    n_plate = len(plate_data)
    n_tunnel = len(tunnel_data)
    n_n2n = len(n2n_data)
    n_anc = len(anc_data)
    n_strut = len(strut_data)
    n_fea = len(fea_data)
    n_pile = len(pile_data)
    n_soil = len(soil_polys) if soil_polys else 0
    n_drains = len(drains)
    n_loads = len(loads)
    n_soil_mat = len(soil_data)
    n_geogrid = len(geogrid_data)
    total = n_plate + n_tunnel + n_n2n + n_anc + n_strut + n_fea + n_pile + n_soil + n_drains + n_loads + n_soil_mat + n_geogrid
    print(f"sc5: {total} elements → {nb_path.name}")
    print(f"    Soil materials: {n_soil_mat}, Plate: {n_plate}, Tunnel: {n_tunnel}, N2N: {n_n2n}, Embedded Beam: {n_anc}, Strut: {n_strut}, FEA: {n_fea}, Pile: {n_pile}, Geogrid: {n_geogrid}, Soil polygons: {n_soil}, Drains: {n_drains}, Loads: {n_loads}")
    print("sc5: Done. Open .ipynb in PLAXIS Jupyter console.")
    _set_progress(wb, f"sc5: done — {total} elements {_bar(1.0)}")


if __name__ == '__main__':
    main()
