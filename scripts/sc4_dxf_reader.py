"""
sc4_dxf_reader.py — DXF → Structural Tab (Button 1) [v0.7.9]
============================================================
Reads a .dxf file and writes ONLY the DXF-extracted coordinates
(cad_layer, x1, y1, x2, y2) into the blue cells of the 'str_2D' tab.
Also extracts soil polygons (cut/fill LWPOLYLINE) into the soil
polygon section.

SC4 clears ALL blue data columns (B-G) for each section before writing
new DXF data. This prevents ghost data from removed layers.

Layout (v0.7.4):
  Rows 1-16: Soil profile section (header row 6, data rows 7-16, max 10)
  Rows 18-38: Line loads + blank rows (header 18, data 19-38, max 20)
  Structural sections (title + blank + header + 20 data rows each):
    Plate:     title 40, header 42, data 43-62
    Tunnel:    title 64, header 66, data 67-86
    N2N:       title 88, header 90, data 91-110
    Anchor:    title 112, header 114, data 115-134
    Strut:     title 136, header 138, data 139-158
    FEA:       title 160, header 162, data 163-182
    Pile:      title 184, header 186, data 187-206
  Geogrid:   header row 210, data rows 211-230 (max 20)
  Soil polygon + drain section: header row 234, data rows 235+
    Soil polygon columns: B=spoly_name, C=cad_layer, D=x1, E=y1, F=x2, G=y2, H=spoly_type, I=polygon_no
    Drain columns: M=drains_name, N=cad_layer, O=x1, P=y1, Q=x2, R=y2, S=drain_no
"""

from pathlib import Path
import re
import ezdxf
import xlwings as xw


# ── Layer prefix → PLAXIS element type ──
# Keys are checked case-insensitively against the DXF layer name
# (after bracket stripping and _number suffix removal).
LAYER_MAP = {
    'plate':  {'type': 'plate'},
    'nn':     {'type': 'n2nanchor'},
    'anc':    {'type': 'embeddedbeam'},
    'pile':   {'type': 'pile'},
    'strut':  {'type': 'strut'},
    'fea':    {'type': 'fixedendanchor'},
    'fe_anchor': {'type': 'fixedendanchor'},  # DXF alias (e.g. FE_Anchor[H350X350])
    'pvd':    {'type': 'drain'},
    'geo':    {'type': 'geogrid'},
    'geogrid': {'type': 'geogrid'},
    'lload':  {'type': 'lineload'},
    'load_line': {'type': 'lineload'},  # DXF alias (e.g. load_line[100kPa])
    'pload':  {'type': 'pointload'},
    'load_point': {'type': 'pointload'},  # DXF alias (e.g. load_point[1000kN])
    'tunnel': {'type': 'tunnel'},
    'waterlevel': {'type': 'waterline'},
}


# ── Fixed blue data ranges per element type (B:F, rows) ──
# Keys match LAYER_MAP['type'] values
BLUE_RANGES = {
    'plate':        {'start': 43, 'end': 62},  # header 42, 20 rows
    'tunnel':       {'start': 67, 'end': 86},  # header 66, 20 rows
    'n2nanchor':    {'start': 91, 'end': 110}, # header 90, 20 rows
    'embeddedbeam': {'start': 115, 'end': 134}, # header 114, 20 rows
    'strut':        {'start': 139, 'end': 158},# header 138, 20 rows
    'fixedendanchor': {'start': 163, 'end': 182},# header 162, 20 rows — FEA
    'pile':         {'start': 187, 'end': 206},# header 186, 20 rows
    'geogrid':      {'start': 211, 'end': 230},# header 210, 20 rows
}

# Geogrid geometry is the only part SC4 owns in this table.
GEOGRID_DATA_START = 211
GEOGRID_DATA_END = 230
GEOGRID_WRITE_COLUMNS = (3, 4, 5, 6, 7)  # C:G only

# Current str_2D layout controls. Geogrid commands are intentionally not
# generated yet; this pass only keeps SC4 writes inside the correct sections.
SPOLY_HEADER_ROW = 234
SPOLY_DATA_START = 235
DRAIN_HEADER_ROW = 234
DRAIN_DATA_START = 235
WATERLINE_HEADER_ROW = 234
WATERLINE_DATA_START = 235
WATERLINE_MAX_CELL = 'Z233'
WATERLINE_WRITE_COLUMNS = (21, 22, 23, 24, 25, 26, 27)

# ── GWFlowBC constants ──
# Same rows and geometry columns as waterlines; only AB is GWFlowBC-owned.
GWFLOW_BC_BEHAVIOUR_MAP = {
    'waterboundary_head': 'Head',
    'waterboundary_closed': 'Closed',
}
GWFLOW_BC_WRITE_COLUMNS = (21, 22, 23, 24, 25, 26, 28)
GEOGRID_HEADER_ROW = 210
GEOGRID_DATA_START = 211
GEOGRID_DATA_END = 230

# ── Output curve-point table ──
# SC4-owned output tab block: headers at row 8, data from row 9.
# B:F are cleared before every DXF refresh; other output columns are preserved.
CURVE_NODE_LAYERS = {'curve_node'}
CURVE_STRESS_LAYERS = {'curve_stresspoint'}
CURVE_OUTPUT_HEADER_ROW = 8
CURVE_OUTPUT_DATA_START = 9
CURVE_OUTPUT_DATA_END = 108  # 100 curve points per refresh
CURVE_OUTPUT_CLEAR_RANGE = {
    'col_start': 2, 'col_end': 6,
    'data_start': CURVE_OUTPUT_DATA_START,
    'data_end': CURVE_OUTPUT_DATA_END,
}

# ── Volume profile spoly_type overrides (column H) ──
# SC4 writes 'str_volume' as default for volume_profile DXF layers.
# User may override column H to request interfaces in SC5.
# 'none' is an explicit per-edge no-interface marker: still a structural
# volume everywhere (SC5/SC6 treat it like 'str_volume'), but the planner
# never builds an interface on that edge. First in the dropdown for
# visibility; the auto-written default stays 'str_volume'.
VOLUME_PROFILE_DEFAULT = 'str_volume'
VOLUME_PROFILE_NONE = 'none'
VOLUME_PROFILE_INTERFACES = (
    'str_vol_posinterface',
    'str_vol_neginterface',
    'str_vol_bothinterface',
    'str_vol_neg_EXboth',
    'str_vol_neg_EXtop',
    'str_vol_neg_EXbot',
    'str_vol_neg_EXleft',
    'str_vol_neg_EXright',
    'str_vol_pos_EXboth',
    'str_vol_pos_EXtop',
    'str_vol_pos_EXbot',
    'str_vol_pos_EXleft',
    'str_vol_pos_EXright',
)
VOLUME_PROFILE_VALID = (
    VOLUME_PROFILE_NONE, VOLUME_PROFILE_DEFAULT,
) + VOLUME_PROFILE_INTERFACES


# ── Progress cell (SC4, mirrors SC5/SC6/SC9 pattern) ──
# str_2D!B4 is the SC4 banner cell (own cell — D4 is SC5's banner on the
# same sheet). Written live during DXF import. main() disables
# screen_updating for speed, so each write briefly re-enables it so the
# cell repaints, then restores the previous state. Never raises —
# progress must not break runs.
PROGRESS_SHEET = 'str_2D'
PROGRESS_CELL = 'B4'
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
    """Write *msg* to str_2D!B4. False/None clears. Never raises."""
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


# ── Coordinate helpers ──

def parse_layer_name(layer_name):
    """Parse layer name with optional brackets.
    'anc_1[grout 300mm]' → ('anc', 1, 'grout 300mm')
    'FE_Anchor[H350X350]' → ('FE_Anchor', 0, 'H350X350')
    'load_point(750kN]' → ('load_point', 0, '750kN')  (mismatched bracket)
    Non-numeric suffix → (name, 0, '')."""
    # Strip bracket content for prefix lookup, but keep full name for Excel
    full_name = layer_name
    bracket_content = ''

    # Extract bracket content — handle mismatched brackets too
    # e.g. load_point(750kN] → extract '750kN'
    open_idx = -1
    close_idx = -1
    for i, ch in enumerate(layer_name):
        if ch in ('[', '('):
            open_idx = i
        if ch in (']', ')'):
            close_idx = i
            break
    if open_idx >= 0 and close_idx > open_idx:
        bracket_content = layer_name[open_idx + 1:close_idx]
        layer_name = layer_name[:open_idx]  # strip brackets for parsing

    # Parse prefix and number
    parts = layer_name.rsplit('_', 1)
    if len(parts) == 2 and parts[1].isdigit():
        return parts[0], int(parts[1]), bracket_content
    return layer_name, 0, bracket_content


def _snap_shared_value(a, b):
    """Choose a deterministic clean value for nearly equal coordinates."""
    # A common two-decimal value removes millimetre-level AutoCAD drift:
    # 4.001 and 4.000 both become 4.00, while 4.001 and 4.002 become 4.00.
    if round(a, 2) == round(b, 2):
        return round(a, 2)
    return round((a + b) / 2, 3)


def normalize_polyline_vertices(vertices, axis_tolerance=0.01):
    """Snap near-horizontal / near-vertical segments to exact axis-aligned values.

    AutoCAD sometimes exports lines intended as horizontal or vertical with
    tiny numerical drift (e.g. dy = 0.001 m).  SC5 uses exact coordinate
    comparison (tolerance ~1e-6) to classify segments and apply extended
    interface offsets.  This function detects nearly axis-aligned segments
    and snaps both endpoints to one common axis value.

    The shared-axis value is propagated forward through consecutive segments,
    so adjacent horizontal segments on the same vertex stay consistent.

    Parameters
    ----------
    vertices : list of (float, float)
        Polyline coordinates already rounded to 3 decimals.
    axis_tolerance : float
        Maximum difference on the varying axis to consider a segment
        nearly axis-aligned.  Default 0.01 m (10 mm).
    """
    if len(vertices) < 2:
        return list(vertices)

    normalized = list(vertices)

    for i in range(len(normalized) - 1):
        x1, y1 = normalized[i]
        x2, y2 = normalized[i + 1]

        # Near-horizontal: Y barely changes → snap both Y to same value
        if y1 != y2 and abs(y2 - y1) <= axis_tolerance:
            y_snap = _snap_shared_value(y1, y2)
            normalized[i] = (x1, y_snap)
            normalized[i + 1] = (x2, y_snap)

        # Near-vertical: X barely changes → snap both X to same value
        elif x1 != x2 and abs(x2 - x1) <= axis_tolerance:
            x_snap = _snap_shared_value(x1, x2)
            normalized[i] = (x_snap, y1)
            normalized[i + 1] = (x_snap, y2)

    return normalized


def normalize_polygon_vertices(vertices, axis_tolerance=0.01):
    """Normalize a closed polygon so every shared vertex is exactly equal.

    This extends the segment-level normalization by forcing the closing
    vertex (the last tuple) to be identical to the first tuple.  This
    guarantees that when SC4 writes polygon segments to str_2D, every
    segment's end point is the same Python object as the next segment's
    start point.
    """
    if len(vertices) < 2:
        return list(vertices)

    normalized = normalize_polyline_vertices(vertices,
                                              axis_tolerance=axis_tolerance)

    if len(normalized) >= 3:
        normalized[-1] = normalized[0]

    return normalized


def ensure_closed_polygon_vertices(vertices, closed_flag=False):
    """Return polygon vertices with one explicit closing vertex.

    Some DXF exporters repeat the first vertex but leave the LWPOLYLINE
    closed bit unset.  Treat that representation as closed as well.
    """
    verts = list(vertices)
    if len(verts) < 3:
        return verts
    if verts[-1] == verts[0]:
        return verts
    if closed_flag:
        verts.append(verts[0])
    return verts


def _volume_edge_key(start, end):
    """Return an orientation-independent key for one polygon edge."""
    start = tuple(start)
    end = tuple(end)
    return min((start, end), (end, start))


def _reconstruct_vertex_chain(segments):
    """Reconstruct an ordered polygon vertex chain from Excel segments."""
    if not segments:
        return None
    remaining = [tuple(segment) for segment in segments]
    first = remaining.pop(0)
    chain = [(first[0], first[1]), (first[2], first[3])]
    while remaining:
        end = chain[-1]
        match_index = None
        match = None
        for index, segment in enumerate(remaining):
            start = (segment[0], segment[1])
            finish = (segment[2], segment[3])
            if start == end:
                match_index, match = index, (start, finish)
                break
            if finish == end:
                match_index, match = index, (finish, start)
                break
        if match is None:
            return None
        remaining.pop(match_index)
        chain.append(match[1])
    if chain[-1] != chain[0]:
        return None
    return chain


def _parse_load_magnitude(layer_name):
    """Parse signed integer or decimal magnitude from a load layer."""
    layer = str(layer_name)
    opening = '[' if '[' in layer else '(' if '(' in layer else None
    if opening is None:
        return 0
    closing = ']' if opening == '[' else ')'
    start = layer.index(opening) + 1
    end = layer.find(closing, start)
    content = layer[start:end if end >= 0 else len(layer)]
    match = re.search(r'[-+]?\d+(?:\.\d+)?', content)
    return float(match.group()) if match else 0


def _load_geometry_key(x1, y1, x2, y2):
    """Return a canonical geometry key for a line or point load.

    Line endpoints are unordered so reversing the CAD line does not break
    preservation. Point loads use only their first coordinate pair.
    """
    p1 = (round(float(x1), 3), round(float(y1), 3))
    if x2 is None or y2 is None:
        return ('point', p1)
    p2 = (round(float(x2), 3), round(float(y2), 3))
    return ('line',) + tuple(sorted((p1, p2)))


def _preserved_load_inputs(existing_loads, geo_key, cad_layer):
    """Return preserved inputs only when both geometry and layer match."""
    entry = existing_loads.get((geo_key, cad_layer))
    if entry is None:
        return None
    return (entry['load_no'], entry['load_model'],
            entry['qx_start'], entry['qy_start'])


def _volume_polygon_key(vertices):
    """Return a canonical, dimensionless key for a volume-profile polygon.

    The closing duplicate is ignored. Cyclic vertex shifts and reversed
    traversal are equivalent, so harmless DXF order changes preserve inputs.
    """
    verts = [tuple(v) for v in vertices]
    if len(verts) > 1 and verts[-1] == verts[0]:
        verts.pop()
    if len(verts) < 3:
        return tuple(verts)

    candidates = []
    for sequence in (verts, list(reversed(verts))):
        candidates.extend(
            tuple(sequence[i:] + sequence[:i])
            for i in range(len(sequence))
        )
    return min(candidates)


def extract_line_coords(entity):
    """LINE entity → (x1, y1, x2, y2), with axis-alignment cleanup."""
    s = entity.dxf.start
    e = entity.dxf.end
    vertices = normalize_polyline_vertices([
        (round(s.x, 3), round(s.y, 3)),
        (round(e.x, 3), round(e.y, 3)),
    ])
    return (vertices[0][0], vertices[0][1], vertices[1][0], vertices[1][1])


def extract_polyline_coords(entity):
    """LWPOLYLINE → [(x, y), ...], rounded but otherwise unchanged."""
    points = list(entity.get_points(format='xy'))
    return [(round(x, 3), round(y, 3)) for x, y in points]


def extract_point_coords(entity):
    """POINT entity → (x, y, x, y)."""
    p = entity.dxf.location
    x, y = round(p.x, 3), round(p.y, 3)
    return (x, y, x, y)


def extract_tunnel_circle_coords(entity):
    """CIRCLE → center x/y, display radius label, numeric radius.

    Geometry is always read from the drawn CIRCLE.  The layer text is metadata.
    """
    center = entity.dxf.center
    radius = round(float(entity.dxf.radius), 3)
    radius_label = f"R={radius:g}m"
    return (round(float(center.x), 3), round(float(center.y), 3), radius_label, radius)


def tunnel_row_count(bracket_content):
    """Return row count: 3 for TBM mode, 2 for double liner, 1 otherwise."""
    s = str(bracket_content).lower()
    if 'tbm' in s:
        return 3
    if 'double' in s:
        return 2
    return 1


def expand_tunnel_circle(entity):
    """Expand one circular tunnel entity into liner input rows.

    TBM mode (bracket contains 'tbm'): primary + grout + secondary
    Double liner (bracket contains 'double'): primary + secondary
    Single liner: primary only
    """
    layer = entity.dxf.layer
    _, _, bracket_content = parse_layer_name(layer)
    x1, y1, radius_label, radius = extract_tunnel_circle_coords(entity)
    n_rows = tunnel_row_count(bracket_content)
    rows = [(layer, 'primary', x1, y1, radius_label, radius)]
    if n_rows == 3:
        rows.append((layer, 'grout', x1, y1, radius_label, radius))
        rows.append((layer, 'secondary', x1, y1, radius_label, radius))
    elif n_rows == 2:
        rows.append((layer, 'secondary', x1, y1, radius_label, radius))
    return rows


def collect_geogrid_rows(entities):
    """Return one geometry row for each geogrid LINE entity.

    The original DXF layer, including bracket metadata, is preserved.
    SC4 later writes these values only to columns C:G.
    """
    rows = []
    for entity in entities:
        layer = entity.dxf.layer
        prefix, _, _ = parse_layer_name(layer)
        if prefix.lower() not in ('geo', 'geogrid'):
            continue
        if entity.dxftype() != 'LINE':
            continue
        x1, y1, x2, y2 = extract_line_coords(entity)
        rows.append((layer, x1, y1, x2, y2))
    return rows


def collect_waterline_rows(msp):
    """Return waterline rows from DXF entities.

    Primary path: LWPOLYLINE entities on layers starting with 'waterlevel'.
    Fallback path: LINE entities on the same layers, chained by shared endpoints
    into ordered polylines (each disconnected chain = one waterline).

    Returns list of (waterline_name, cad_layer, x1, y1, x2, y2, waterline_no)
    — one row per segment, matching the soil polygon write pattern.
    """
    # Separate LWPOLYLINE and LINE water entities
    lwpoly_entities = []
    line_entities = []
    for entity in msp:
        layer_lower = entity.dxf.layer.lower()
        if not layer_lower.startswith('waterlevel'):
            continue
        if entity.dxftype() == 'LWPOLYLINE':
            lwpoly_entities.append(entity)
        elif entity.dxftype() == 'LINE':
            line_entities.append(entity)

    rows = []
    wl_no = 0

    # --- LWPOLYLINE path (one entity = one polyline = one waterline) ---
    for entity in lwpoly_entities:
        wl_no += 1
        layer = entity.dxf.layer
        verts = extract_polyline_coords(entity)
        if len(verts) < 2:
            continue
        for i in range(len(verts) - 1):
            rows.append((
                f'waterline_{wl_no}', layer,
                verts[i][0], verts[i][1],
                verts[i + 1][0], verts[i + 1][1],
                wl_no,
            ))

    # --- LINE fallback path (chain connected segments into polylines) ---
    if line_entities:
        # Collect all segments on each layer
        segments_by_layer = {}  # layer → [(start_pt, end_pt), ...]
        for entity in line_entities:
            layer = entity.dxf.layer
            x1, y1, x2, y2 = extract_line_coords(entity)
            segments_by_layer.setdefault(layer, []).append(
                ((x1, y1), (x2, y2))
            )

        for layer, segments in segments_by_layer.items():
            # Build endpoint adjacency
            adj = {}  # point → [(connected_point, segment_index)]
            for idx, (start, end) in enumerate(segments):
                adj.setdefault(start, []).append((end, idx))
                adj.setdefault(end, []).append((start, idx))

            # Walk disconnected chains from endpoints (degree != 2)
            visited_segs = set()
            # Find true endpoints: degree != 2
            true_endpoints = [pt for pt, nbrs in adj.items()
                              if len(nbrs) != 2]
            # If no true endpoints (closed loop), start from any unvisited seg
            if not true_endpoints:
                true_endpoints = [segments[0][0]]

            for start_pt in true_endpoints:
                # Start a chain from this endpoint
                current = start_pt
                chain = [current]
                while True:
                    found_next = False
                    for nbr, seg_idx in adj.get(current, []):
                        if seg_idx in visited_segs:
                            continue
                        visited_segs.add(seg_idx)
                        chain.append(nbr)
                        current = nbr
                        found_next = True
                        break
                    if not found_next:
                        break

                if len(chain) < 2:
                    continue
                wl_no += 1
                for i in range(len(chain) - 1):
                    rows.append((
                        f'waterline_{wl_no}', layer,
                        chain[i][0], chain[i][1],
                        chain[i + 1][0], chain[i + 1][1],
                        wl_no,
                    ))

            # Pick up any unvisited segments (orphans)
            for idx, (start, end) in enumerate(segments):
                if idx not in visited_segs:
                    wl_no += 1
                    rows.append((
                        f'waterline_{wl_no}', layer,
                        start[0], start[1],
                        end[0], end[1],
                        wl_no,
                    ))

    # Sort waterlines by highest Y (top-most first), same as soil polygon sort
    def _max_y(row):
        return max(row[3], row[5])  # max(y1, y2)

    # Group by waterline_no, find max Y per group
    groups = {}
    for row in rows:
        wl_no_key = row[6]
        groups.setdefault(wl_no_key, []).append(row)
    sorted_groups = sorted(groups.items(),
                           key=lambda item: max(_max_y(r) for r in item[1]),
                           reverse=True)

    # Re-number and flatten
    result = []
    new_no = 0
    for _, group_rows in sorted_groups:
        new_no += 1
        for row in group_rows:
            result.append((
                f'waterline_{new_no}', row[1],
                row[2], row[3], row[4], row[5],
                new_no,
            ))

    return result


def collect_gwflowbc_rows(msp):
    """Return groundwater-flow boundary rows from DXF entities.

    Every LINE entity on a layer matching waterboundary_<type> becomes one
    GWFlowBC row.  No chaining — connected segments are broken into
    individual boundaries.  LWPOLYLINE entities on these layers are also
    accepted; each segment becomes its own GWFlowBC.

    Returns list of (gwflowbc_name, cad_layer, x1, y1, x2, y2, gwflowbc_no)
    tuples — same shape as waterline rows.  Behaviour (Head/Closed) is
    derived downstream from cad_layer; href (Head reference) equals the
    boundary elevation Y and is also derived downstream.  SC4 stores no
    behaviour or href columns.
    """
    rows = []
    for entity in msp:
        layer = entity.dxf.layer
        layer_key = layer.lower().strip()

        if GWFLOW_BC_BEHAVIOUR_MAP.get(layer_key) is None:
            continue

        if entity.dxftype() == 'LINE':
            x1, y1, x2, y2 = extract_line_coords(entity)
            segments = [((x1, y1), (x2, y2))]
        elif entity.dxftype() == 'LWPOLYLINE':
            verts = extract_polyline_coords(entity)
            segments = [
                (verts[i], verts[i + 1]) for i in range(len(verts) - 1)
            ]
        else:
            continue

        for (x1, y1), (x2, y2) in segments:
            rows.append((layer, x1, y1, x2, y2))

    # Sort: Head layers before Closed layers; within each group top-to-
    # bottom by midpoint Y (descending).  Layer text decides the group so
    # the parse stays a pure cad_layer concern downstream.
    def _group(layer_name):
        key = layer_name.lower().strip()
        if 'head' in key:
            return 0
        if 'closed' in key:
            return 1
        return 2

    def _sort_key(row):
        layer, x1, y1, x2, y2 = row
        return (_group(layer), -(y1 + y2) / 2.0)

    rows.sort(key=_sort_key)

    return [
        (f'gwflowbc_{no}', layer, x1, y1, x2, y2, no)
        for no, (layer, x1, y1, x2, y2) in enumerate(rows, start=1)
    ]


def _curve_point_type(layer_name):
    """Return PLAXIS curve-point selection type for a DXF layer.

    Valid PLAXIS Output addcurvepoint type strings:
      ``"node"``        - mesh node
      ``"stresspoint"`` - stress integration point
    """
    layer_lower = str(layer_name).strip().lower()
    if layer_lower in {x.lower() for x in CURVE_NODE_LAYERS}:
        return 'node'
    if layer_lower in {x.lower() for x in CURVE_STRESS_LAYERS}:
        return 'stresspoint'
    return None


def collect_curve_points(msp):
    """Collect DXF POINT entities used for Output curve-point selection.

    Returns ``(name, cad_layer, x, y, point_type)`` tuples.  Points are
    grouped by selection type and sorted from high to low Y within each
    group, giving deterministic names and a useful top-to-bottom order.
    """
    grouped = {'node': [], 'stresspoint': []}
    for entity in msp:
        if entity.dxftype() != 'POINT':
            continue
        layer = entity.dxf.layer
        point_type = _curve_point_type(layer)
        if point_type is None:
            continue
        p = entity.dxf.location
        grouped[point_type].append((layer, round(float(p.x), 3),
                                    round(float(p.y), 3)))

    rows = []
    layer_for_type = {'node': 'curve_node', 'stresspoint': 'curve_stresspoint'}
    for point_type in ('node', 'stresspoint'):
        grouped[point_type].sort(key=lambda item: (-item[2], item[1]))
        prefix = layer_for_type[point_type]
        for number, (layer, x, y) in enumerate(grouped[point_type], 1):
            rows.append((f'{prefix}_{number}', layer, x, y, point_type))
    return rows


# ── Main ──
def main(dxf_path=None):
    """Read DXF and write extracted data into the workbook.

    Called from Excel via:
    ``RunPython "import scripts.sc4_dxf_reader as l; l.main()"``
    """
    # ── Get workbook ──
    try:
        wb = xw.Book.caller()
    except Exception:
        wb = xw.books.active

    # ── Find DXF file ──
    # Excel COM round-trips dominate this script's runtime. Disable screen
    # repainting for the duration and restore it even if a later write fails.
    app = wb.app
    previous_screen_updating = app.screen_updating
    app.screen_updating = False
    try:
        return _main_io(wb, dxf_path)
    finally:
        app.screen_updating = previous_screen_updating


def _main_io(wb, dxf_path=None):
    """Run SC4 with Excel screen updating already suspended.

    Progress wrapper (mirrors SC5/SC6 pattern): WORKING banner on
    str_2D!B4; crash → FAILED, success path sets the done text.
    Never raises from progress itself.
    """
    _set_progress(wb, f"sc4: {_spin()} WORKING — opening DXF…")
    try:
        return _main_io_inner(wb, dxf_path)
    finally:
        # Never leave a stale WORKING banner: crash → FAILED, else done
        # text is already set by the success path.
        try:
            cur = wb.sheets[PROGRESS_SHEET].range(PROGRESS_CELL).value
        except Exception:
            cur = None
        if isinstance(cur, str) and 'WORKING' in cur.upper():
            _set_progress(
                wb, "sc4: FAILED — see console for detail")


def _main_io_inner(wb, dxf_path=None):
    """Core logic for SC4 — separated for the progress wrapper."""
    if dxf_path is None:
        wb_dir = Path(wb.fullname).parent
        # Use Excel's native file dialog (always visible over Excel window)
        # 3 = msoFileDialogOpen
        fd = wb.app.api.FileDialog(3)
        fd.Title = 'Select DXF File'
        fd.InitialFileName = str(wb_dir) + '\\'
        fd.Filters.Clear()
        fd.Filters.Add('DXF files', '*.dxf')
        fd.Filters.Add('All files', '*.*')
        if fd.Show() == -1:  # -1 = OK clicked
            dxf_path = Path(fd.SelectedItems(1))
        if not dxf_path:
            raise FileNotFoundError('No DXF file selected')
    else:
        dxf_path = Path(dxf_path)

    print(f"sc4: Reading DXF: {dxf_path.name}")
    _set_progress(wb, f"sc4: {_spin()} WORKING — reading DXF {_bar(0.05)}")

    # ── Parse DXF ──
    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()

    # ── Collect entities by element type ──
    elements = {}  # type → list of (layer_name, x1, y1, x2, y2)
    _debug_layers = []

    for entity in msp:
        layer = entity.dxf.layer
        etype = entity.dxftype()
        prefix, number, bracket_note = parse_layer_name(layer)

        if prefix.lower() not in LAYER_MAP:
            continue

        plaxis_type = LAYER_MAP[prefix.lower()]['type']

        # Extract tunnel circles separately: CIRCLE geometry is valid only here.
        if plaxis_type == 'tunnel':
            if etype == 'CIRCLE':
                elements.setdefault(plaxis_type, []).extend(
                    expand_tunnel_circle(entity)
                )
            continue

        # Only process element types that have blue ranges defined
        if plaxis_type not in BLUE_RANGES:
            continue

        # Geogrid C:G is the only DXF-owned part of the geogrid table.
        # Preserve the original layer, including bracket metadata, in C.
        if plaxis_type == 'geogrid':
            display_name = layer
        elif bracket_note:
            display_name = f"{prefix}_{number}[{bracket_note}]"
        else:
            display_name = f"{prefix}_{number}" if number else prefix

        # Extract coordinates based on entity type
        if etype == 'LINE':
            coords = extract_line_coords(entity)
        elif etype == 'POINT':
            p = entity.dxf.location
            x, y = round(p.x, 3), round(p.y, 3)
            coords = (x, y, None, None)  # POINT: x1,y1 only
        elif etype == 'LWPOLYLINE':
            # Polyline → multiple segments from one normalized vertex list.
            verts = extract_polyline_coords(entity)
            # Ensure the closing edge is emitted for closed polylines
            # (consistent with soil polygon handling below).
            if entity.closed and len(verts) >= 2 and verts[-1] != verts[0]:
                verts = verts + [verts[0]]
            for i in range(len(verts) - 1):
                if plaxis_type not in elements:
                    elements[plaxis_type] = []
                elements[plaxis_type].append((
                    display_name,
                    verts[i][0], verts[i][1],
                    verts[i+1][0], verts[i+1][1],
                ))
            continue
        else:
            continue

        if plaxis_type not in elements:
            elements[plaxis_type] = []
        elements[plaxis_type].append((display_name, coords[0], coords[1], coords[2], coords[3]))

    # ── Sort each element list by number (plate_1 before plate_2) ──
    def _sort_key(row):
        name = row[0]  # e.g. "plate_1" or "anc_1[grout 300mm]"
        # Strip bracket content before parsing number
        base = name.split('[')[0] if '[' in name else name
        parts = base.rsplit('_', 1)
        num = int(parts[1]) if len(parts) == 2 and parts[1].isdigit() else 0
        return (parts[0], num)

    for plaxis_type in elements:
        elements[plaxis_type].sort(key=_sort_key)

    # ── Re-number FEA (fixedendanchor) by unique bracket (layer-based) ──
    # Bracket format: [SECTION_DIRECTION] e.g. [H350X350_-12.5]
    # Parse Direction_x from bracket (last _ separator).
    # cad_layer keeps full bracket for unique grouping: fe_anchor1[H350X350_-12.5]
    _fea_dir_x = {}  # cad_layer_name → Direction_x value
    if 'fixedendanchor' in elements and elements['fixedendanchor']:
        _layer_map = {}   # bracket → sequential number
        _counter = 0
        for i, (name, x1, y1, x2, y2) in enumerate(elements['fixedendanchor']):
            _brk_idx = name.find('[')
            _bracket = name[_brk_idx:] if _brk_idx >= 0 else ''
            if _bracket not in _layer_map:
                _counter += 1
                _layer_map[_bracket] = _counter
            _num = _layer_map[_bracket]
            _prefix = 'fe_anchor'
            _new_name = f'{_prefix}{_num}{_bracket}'

            # Parse Direction_x from bracket: [SECTION_DIRECTION]
            _inner = _bracket[1:-1] if _bracket.startswith('[') and _bracket.endswith(']') else ''
            _uscore = _inner.rfind('_')
            if _uscore > 0:
                _dir_str = _inner[_uscore+1:]
                try:
                    _fea_dir_x[_new_name] = float(_dir_str)
                except ValueError:
                    pass  # no valid direction — leave empty for user input

            elements['fixedendanchor'][i] = (
                _new_name, x1, y1, x2, y2,
            )

        # v0.6.17: Sort FEA by bracket group, then descending Y (shallowest first).
        # Groups all anchors on one wall together, with the topmost as green parent.
        def _fea_sort_key(row):
            _name = row[0]
            _brk_idx = _name.find('[')
            _bracket = _name[_brk_idx:] if _brk_idx >= 0 else ''
            return (_bracket, -row[2])  # (bracket, -y1)

        elements['fixedendanchor'].sort(key=_fea_sort_key)

    # ── Write blue columns into fixed ranges ──
    # IMPORTANT: Never use .clear() — it destroys data validations (dropdowns).
    # Instead, write values directly. .value = None only clears the cell value.
    # v0.5: Read header row to find column positions by name (column-insensitive).
    sheet = wb.sheets['str_2D']
    _set_progress(wb, f"sc4: {_spin()} WORKING — scanning entities {_bar(0.15)}")

    _sec_names = list(BLUE_RANGES)
    for _si, (plaxis_type, rng) in enumerate(BLUE_RANGES.items()):
        data = elements.get(plaxis_type, [])
        start = rng['start']
        end = rng['end']
        header_row = start - 1  # header is one row above first data row

        # Unhide all rows in this section (sc5 may have hidden them).
        # Row visibility is a per-row COM property and cannot be block-written.
        for r in range(start, end + 1):
            sheet.api.Rows(r).Hidden = False

        # Clear all DXF-owned blue columns (B:G).
        # One 2D write replaces 120 individual COM calls per section.
        n_rows = end - start + 1
        sheet.range((start, 2), (end, 7)).value = [[None] * 6 for _ in range(n_rows)]

        # Direction_x (column P) is only SC4-owned in the FEA section.
        # Do NOT clear column P here — it contains user-maintained yellow/orange
        # inputs (e.g. Tskin_end) in other sections (Anchor, Pile, etc.).
        if plaxis_type == 'fixedendanchor':
            sheet.range((start, 16), (end, 16)).value = [[None] for _ in range(n_rows)]

        # Build column header → column number mapping from header row
        col_map = {}
        for col in range(2, 20):  # scan B to S
            val = sheet.range(header_row, col).value
            if val is not None:
                col_map[str(val).strip()] = col

        # Find columns by header name
        c_cad = col_map.get('cad_layer', 3)
        c_x1 = col_map.get('x1', 4)
        c_y1 = col_map.get('y1', 5)
        c_x2 = col_map.get('x2', 6)
        c_y2 = col_map.get('y2', 7)
        c_position = col_map.get('liner_position')

        # ── Batch-write DXF data as one 2D array ──
        # Determine which columns are contiguous and write them in one COM call.
        target_cols = [c_cad, c_x1, c_y1, c_x2, c_y2]
        if plaxis_type == 'tunnel' and c_position is not None:
            target_cols.append(c_position)
        min_col = min(target_cols)
        max_col = max(target_cols)
        col_span = max_col - min_col + 1

        # Build the block array
        block = []
        for i, item in enumerate(data):
            if i >= (end - start + 1):
                break
            row_vals = [None] * col_span
            if plaxis_type == 'tunnel':
                layer_name, liner_position, x1, y1, radius_label, radius = item
                row_vals[c_cad - min_col] = layer_name
                row_vals[c_x1 - min_col] = x1
                row_vals[c_y1 - min_col] = y1
                row_vals[c_x2 - min_col] = radius_label
                row_vals[c_y2 - min_col] = radius
                if c_position is not None:
                    row_vals[c_position - min_col] = liner_position
            else:
                layer_name, x1, y1, x2, y2 = item
                row_vals[c_cad - min_col] = layer_name
                row_vals[c_x1 - min_col] = x1
                row_vals[c_y1 - min_col] = y1
                row_vals[c_x2 - min_col] = x2
                row_vals[c_y2 - min_col] = y2
                # FEA: Direction_x column
                if plaxis_type == 'fixedendanchor':
                    c_dir = col_map.get('Direction_x')
                    dir_val = _fea_dir_x.get(layer_name)
                    if c_dir is not None and dir_val is not None:
                        if min_col <= c_dir <= max_col:
                            row_vals[c_dir - min_col] = dir_val
                        else:
                            # Direction_x is outside the contiguous block — write separately
                            row_vals_out_of_block = True
            block.append(row_vals)

        # Write the entire block in one COM call (replaces 5–6 calls per row)
        if block:
            n_rows = len(block)
            sheet.range((start, min_col), (start + n_rows - 1, max_col)).value = block

        # Handle Direction_x if it fell outside the contiguous block
        if plaxis_type == 'fixedendanchor':
            c_dir = col_map.get('Direction_x')
            if c_dir is not None and (c_dir < min_col or c_dir > max_col):
                dir_block = [[None] for _ in range(len(data))]
                for i, item in enumerate(data[:len(dir_block)]):
                    layer_name = item[0]
                    dir_val = _fea_dir_x.get(layer_name)
                    dir_block[i][0] = dir_val
                if dir_block:
                    sheet.range((start, c_dir), (start + len(dir_block) - 1, c_dir)).value = dir_block

        # ── Color parent rows green ──
        # This must remain per-row because .color is a per-cell COM property.
        GREEN = (200, 230, 200)
        WHITE = (255, 255, 255)
        seen_layers = set()
        cad_vals = [row[target_cols[0] - min_col] for row in block] if block else []
        for i, cad_val in enumerate(cad_vals):
            row = start + i
            if cad_val is None:
                sheet.range(row, c_cad).color = WHITE
                continue
            layer_key = str(cad_val).strip().lower()
            if layer_key and layer_key not in seen_layers:
                sheet.range(row, c_cad).color = GREEN
                seen_layers.add(layer_key)
            else:
                sheet.range(row, c_cad).color = WHITE

        # Per-section tick: 20% → 55% across the 8 BLUE_RANGES sections.
        _frac = 0.20 + 0.35 * (_si + 1) / max(1, len(_sec_names))
        _set_progress(
            wb,
            f"sc4: {_spin()} WORKING — {plaxis_type} {_bar(_frac)}")

    # ── Curve points → output tab (v0.7.10) ──
    _set_progress(wb, f"sc4: {_spin()} WORKING — curve points {_bar(0.60)}")
    # POINT entities on curve_node / curve_stresspoint are SC4-owned.
    # Clear only B:F data rows so other output-tab content is preserved.
    curve_sheet = wb.sheets['output']
    curve_points = collect_curve_points(msp)
    curve_count = CURVE_OUTPUT_DATA_END - CURVE_OUTPUT_DATA_START + 1
    curve_sheet.range(
        (CURVE_OUTPUT_DATA_START, CURVE_OUTPUT_CLEAR_RANGE['col_start']),
        (CURVE_OUTPUT_DATA_END, CURVE_OUTPUT_CLEAR_RANGE['col_end']),
    ).value = [[None] * 5 for _ in range(curve_count)]

    curve_block = [
        [name, layer, x, y, point_type]
        for name, layer, x, y, point_type in curve_points[:curve_count]
    ]
    if curve_block:
        curve_sheet.range(
            (CURVE_OUTPUT_DATA_START, CURVE_OUTPUT_CLEAR_RANGE['col_start']),
            (CURVE_OUTPUT_DATA_START + len(curve_block) - 1,
             CURVE_OUTPUT_CLEAR_RANGE['col_end']),
        ).value = curve_block

    if len(curve_points) > curve_count:
        print(
            f"  WARNING: {len(curve_points)} curve points in DXF but limit is "
            f"{curve_count} — truncated"
        )
    print(
        f"  curve points: {min(len(curve_points), curve_count)} "
        f"({sum(p[4] == 'node' for p in curve_points[:curve_count])} node, "
        f"{sum(p[4] == 'stresspoint' for p in curve_points[:curve_count])} stresspoint) "
        f"→ output!B{CURVE_OUTPUT_DATA_START}:F{CURVE_OUTPUT_DATA_END}"
    )

    # ── Summary (structural) ──
    print(f"sc4: {dxf_path.name} → str_2D tab")
    for plaxis_type in BLUE_RANGES:
        count = len(elements.get(plaxis_type, []))
        if count > 0:
            rng = BLUE_RANGES[plaxis_type]
            print(f"  {plaxis_type}: {count} elements → C{rng['start']}:G{rng['end']}")

    # ── Geogrid LINE entities ──
    # Only C:G is SC4-owned; B and H:R remain untouched.
    geogrid_rows = collect_geogrid_rows(msp)
    geogrid_start = GEOGRID_DATA_START
    geogrid_end = GEOGRID_DATA_END
    geogrid_count = geogrid_end - geogrid_start + 1
    sheet.range((geogrid_start, 3), (geogrid_end, 7)).value = [
        [None] * len(GEOGRID_WRITE_COLUMNS) for _ in range(geogrid_count)
    ]

    _set_progress(wb, f"sc4: {_spin()} WORKING — geogrids {_bar(0.65)}")
    geogrid_block = [list(row) for row in geogrid_rows[:geogrid_count]]
    if geogrid_block:
        sheet.range(
            (geogrid_start, 3),
            (geogrid_start + len(geogrid_block) - 1, 7),
        ).value = geogrid_block

    if len(geogrid_rows) > geogrid_count:
        print(
            f"  WARNING: {len(geogrid_rows)} geogrids in DXF but limit is "
            f"{geogrid_count} — truncated"
        )
    print(
        f"  geogrids: {min(len(geogrid_rows), geogrid_count)} lines "
        f"→ C{geogrid_start}:G{geogrid_end}"
    )

    # ── Soil polygons (cut/fill LWPOLYLINE) ──
    # Header at row 234, data from row 235, max rows from cell I233 (default 100, min 50)
    SOIL_ROW = SPOLY_DATA_START
    # Read max rows from I233; default 100, min 50
    _raw_max = sheet.range(SPOLY_HEADER_ROW - 1, 9).value  # I233
    try:
        SOIL_MAX = max(50, int(_raw_max))
    except (TypeError, ValueError):
        SOIL_MAX = 100

    # Collect soil polygons: raw list first, then sort and number
    _set_progress(wb, f"sc4: {_spin()} WORKING — soil polygons {_bar(0.70)}")
    _soil_raw = []  # [(layer, y_centroid, [verts]), ...]
    for entity in msp:
        layer = entity.dxf.layer.lower()
        if layer not in ('cut', 'fill', 'volume_profile', 'soil_replacement'):
            continue
        if entity.dxftype() != 'LWPOLYLINE':
            continue
        # Normalize the complete polygon once, then expand it into rows.
        raw_verts = extract_polyline_coords(entity)
        had_explicit_closure = len(raw_verts) > 2 and raw_verts[-1] == raw_verts[0]
        verts = ensure_closed_polygon_vertices(
            raw_verts, closed_flag=bool(entity.closed),
        )
        if entity.closed or had_explicit_closure:
            verts = normalize_polygon_vertices(verts)
        if len(verts) < 3:
            continue  # skip degenerate polygons
        # The centroid excludes the duplicate closing tuple.
        centroid_verts = verts[:-1] if verts[-1] == verts[0] else verts
        y_centroid = sum(v[1] for v in centroid_verts) / len(centroid_verts)
        _soil_raw.append((layer, y_centroid, verts))

    # Sort by Y-centroid descending (top-most first) per layer type
    _cut = sorted([p for p in _soil_raw if p[0] == 'cut'], key=lambda x: -x[1])
    _fill = sorted([p for p in _soil_raw if p[0] == 'fill'], key=lambda x: -x[1])
    _soil_repl = sorted([p for p in _soil_raw if p[0] == 'soil_replacement'],
                        key=lambda x: -x[1])

    # Number sequentially: cut_1=top-most cut, fill_1=top-most fill,
    # soil_replacement_1=top-most replacement
    sorted_polys = []
    for i, (layer, yc, verts) in enumerate(_cut, 1):
        sorted_polys.append((f'cut_{i}', {'verts': verts, 'y_centroid': yc, 'layer': layer}))
    for i, (layer, yc, verts) in enumerate(_fill, 1):
        sorted_polys.append((f'fill_{i}', {'verts': verts, 'y_centroid': yc, 'layer': layer}))
    for i, (layer, yc, verts) in enumerate(_soil_repl, 1):
        sorted_polys.append(
            (f'soil_replacement_{i}',
             {'verts': verts, 'y_centroid': yc, 'layer': layer}))

    # Assign polygon_no by Y-centroid across ALL types (matches PLAXIS creation order)
    _all_sorted = sorted(sorted_polys, key=lambda x: -x[1]['y_centroid'])
    _polygon_no_map = {}  # spoly_name → polygon_no
    for i, (name, _) in enumerate(_all_sorted, 1):
        _polygon_no_map[name] = i

    # ── str_volume (volume_profile DXF layer) — appended after soil ──
    _vol_raw = [p for p in _soil_raw if p[0] == 'volume_profile']
    _vol_sorted = sorted(_vol_raw, key=lambda x: -x[1])
    _vol_start_no = len(_polygon_no_map) + 1
    for i, (layer, yc, verts) in enumerate(_vol_sorted):
        spoly_name = f'str_volume_{i + 1}'
        sorted_polys.append((spoly_name, {'verts': verts, 'y_centroid': yc, 'layer': layer}))
        _polygon_no_map[spoly_name] = _vol_start_no + i

    # Build column map from header row 234 (soil polygon columns B-I;
    # J/K are user-maintained before/after material selections).
    col_map = {}
    for col in range(2, 12):  # B=2 to K=11
        val = sheet.range(SPOLY_HEADER_ROW, col).value
        if val is not None:
            col_map[str(val).strip()] = col

    c_name  = col_map.get('spoly_name', 2)
    c_cad   = col_map.get('cad_layer', 3)
    c_x1    = col_map.get('x1', 4)
    c_y1    = col_map.get('y1', 5)
    c_x2    = col_map.get('x2', 6)
    c_y2    = col_map.get('y2', 7)
    c_type  = col_map.get('spoly_type', 8)
    c_pno   = col_map.get('polygon_no', 9)

    # J/K are intentionally not included in target_cols_spoly: they contain
    # user selections and must survive an SC4 refresh. Column H (spoly_type)
    # is also preserved for existing volume_profile rows so the user's
    # interface override survives a DXF refresh.
    target_cols_spoly = [c_name, c_cad, c_x1, c_y1, c_x2, c_y2, c_type, c_pno]
    min_col_spoly = min(target_cols_spoly)
    max_col_spoly = max(target_cols_spoly)
    col_span_spoly = max_col_spoly - min_col_spoly + 1

    # Snapshot existing volume_profile spoly_type values before overwriting.
    _set_progress(wb, f"sc4: {_spin()} WORKING — soil snapshot {_bar(0.78)}")
    # Column H is per segment row, so preservation is per EDGE within a
    # geometry-matched polygon: first match the complete polygon (see
    # _volume_polygon_key), then match each segment by its unordered
    # endpoint pair (_volume_edge_key).  A geometry-matched polygon keeps
    # every per-segment override; an unmatched (edited/new) polygon resets
    # to VOLUME_PROFILE_DEFAULT.
    existing_volume_types = {}  # {geometry_key: {edge_key: spoly_type}}
    _vol_segments = {}          # {spoly_name: [(x1, y1, x2, y2), ...]}
    _vol_types = {}             # {spoly_name: [(row, spoly_type), ...]}
    for r in range(SOIL_ROW, SOIL_ROW + SOIL_MAX):
        _ex_name = sheet.range(r, c_name).value
        if _ex_name is None:
            continue
        _name = str(_ex_name).strip()
        _ex_type = sheet.range(r, c_type).value
        if _ex_type is None or str(_ex_type).strip() not in VOLUME_PROFILE_VALID:
            continue
        _ex_x1 = sheet.range(r, c_x1).value
        _ex_y1 = sheet.range(r, c_y1).value
        _ex_x2 = sheet.range(r, c_x2).value
        _ex_y2 = sheet.range(r, c_y2).value
        if None in (_ex_x1, _ex_y1, _ex_x2, _ex_y2):
            continue
        _seg = (round(float(_ex_x1), 3), round(float(_ex_y1), 3),
                round(float(_ex_x2), 3), round(float(_ex_y2), 3))
        _vol_segments.setdefault(_name, []).append(_seg)
        _vol_types.setdefault(_name, []).append(
            (r, str(_ex_type).strip()))
    for _name, _segs in _vol_segments.items():
        _chain = _reconstruct_vertex_chain(_segs)
        if _chain is None:
            continue
        _chain = normalize_polygon_vertices(_chain)
        # Build edge map from the NORMALIZED chain so keys match the
        # DXF-normalized vertices used in the write phase.  Each chain
        # edge at index i corresponds to the i-th segment in _segs.
        _edge_types = {}
        for _idx in range(len(_chain) - 1):
            _ek = _volume_edge_key(_chain[_idx], _chain[_idx + 1])
            _edge_types[_ek] = _vol_types[_name][_idx][1]
        existing_volume_types[_volume_polygon_key(_chain)] = _edge_types

    # Remove stale validation from the complete spoly_type data range first.
    # Rows may have changed from volume_profile to another type, or may now be
    # blank; both cases must lose the old dropdown on an SC4 reload.
    for r in range(SOIL_ROW, SOIL_ROW + SOIL_MAX):
        try:
            sheet.range(r, c_type).api.Validation.Delete()
        except Exception:
            # Excel raises when a cell has no validation; that is harmless.
            pass

    spoly_block = []
    volume_rows = []  # Excel rows to receive dropdown validation
    n_written = 0
    for name, poly in sorted_polys:
        verts = poly['verts']
        layer = poly['layer']
        p_no = _polygon_no_map[name]
        for i in range(len(verts) - 1):
            if len(spoly_block) >= SOIL_MAX:
                break
            row_vals = [None] * col_span_spoly
            row_vals[c_name - min_col_spoly]  = name
            row_vals[c_cad - min_col_spoly]   = layer
            row_vals[c_x1 - min_col_spoly]    = verts[i][0]
            row_vals[c_y1 - min_col_spoly]    = verts[i][1]
            row_vals[c_x2 - min_col_spoly]    = verts[i+1][0]
            row_vals[c_y2 - min_col_spoly]    = verts[i+1][1]
            if layer == 'volume_profile':
                # Preserve user's interface override per-edge when polygon
                # geometry matches; unmatched edges or polygons reset to default.
                _geo_key = _volume_polygon_key(verts)
                _edge_map = existing_volume_types.get(_geo_key)
                if _edge_map is not None:
                    _ek = _volume_edge_key(
                        (verts[i][0], verts[i][1]),
                        (verts[i + 1][0], verts[i + 1][1]))
                    spoly_type = _edge_map.get(_ek, VOLUME_PROFILE_DEFAULT)
                else:
                    spoly_type = VOLUME_PROFILE_DEFAULT
                volume_rows.append(SOIL_ROW + n_written)
            else:
                spoly_type = layer
            row_vals[c_type - min_col_spoly]  = spoly_type
            row_vals[c_pno - min_col_spoly]   = p_no
            spoly_block.append(row_vals)
            n_written += 1

    if spoly_block:
        n_rows_spoly = len(spoly_block)
        sheet.range((SOIL_ROW, min_col_spoly),
                     (SOIL_ROW + n_rows_spoly - 1, max_col_spoly)).value = spoly_block
    _set_progress(wb, f"sc4: {_spin()} WORKING — soil written {_bar(0.85)}")

    # Apply a dropdown data-validation list on volume_profile rows so the
    # user can select the interface mode without leaving the cell.
    if volume_rows:
        _validation_formula = ','.join(VOLUME_PROFILE_VALID)
        for vr in volume_rows:
            cell = sheet.range(vr, c_type)
            try:
                cell.api.Validation.Delete()
            except Exception:
                pass
            cell.api.Validation.Add(
                Type=3, AlertStyle=1, Operator=1,
                Formula1=_validation_formula,
            )
            cell.api.Validation.IgnoreBlank = True
            cell.api.Validation.InCellDropdown = True

    # Clear leftover rows (block of Nones for columns B-I)
    row = SOIL_ROW + n_written
    clear_end = SOIL_ROW + SOIL_MAX - 1
    if row <= clear_end:
        sheet.range((row, min_col_spoly), (clear_end, max_col_spoly)).value = \
            [[None] * col_span_spoly for _ in range(clear_end - row + 1)]

    print(f"  soil polygons: {len(sorted_polys)} polys ({n_written} segments) → row {SOIL_ROW}+")
    if volume_rows:
        print(f"  volume_profile dropdown applied to {len(volume_rows)} rows")

    # ── PVD Drains (columns M-S, header row 234, data row 235+) ──
    DRAIN_ROW = DRAIN_DATA_START
    # Read max rows from R233; default 125
    _raw_drain_max = sheet.range(DRAIN_HEADER_ROW - 1, 18).value  # R233
    try:
        DRAIN_MAX = max(1, int(_raw_drain_max))
    except (TypeError, ValueError):
        DRAIN_MAX = 100

    # Collect PVD drains from DXF (layer prefix 'pvd')
    _set_progress(wb, f"sc4: {_spin()} WORKING — drains {_bar(0.88)}")
    drain_data = []
    for entity in msp:
        layer = entity.dxf.layer.lower()
        if layer != 'pvd':
            continue
        if entity.dxftype() != 'LINE':
            continue
        s = entity.dxf.start
        e = entity.dxf.end
        x1, y1, x2, y2 = extract_line_coords(entity)
        drain_data.append({
            'cad_layer': 'pvd',
            'x1': x1, 'y1': y1,
            'x2': x2, 'y2': y2,
        })

    # Sort drains by x coordinate (left to right)
    drain_data.sort(key=lambda d: d['x1'])

    # Build drain name: pvd_1, pvd_2, ...
    for i, d in enumerate(drain_data, 1):
        d['name'] = f'pvd_{i}'

    # Build column map from header row 234 (columns M-S)
    drain_col_map = {}
    for col in range(13, 20):  # M=13 to S=19
        val = sheet.range(SPOLY_HEADER_ROW, col).value
        if val is not None:
            drain_col_map[str(val).strip()] = col

    dc_name = drain_col_map.get('drains_name', 13)
    dc_cad = drain_col_map.get('cad_layer', 14)
    dc_x1 = drain_col_map.get('x1', 15)
    dc_y1 = drain_col_map.get('y1', 16)
    dc_x2 = drain_col_map.get('x2', 17)
    dc_y2 = drain_col_map.get('y2', 18)
    dc_no = drain_col_map.get('drain_no', 19)

    # ── Batch-write PVD drain data ──
    target_cols_drain = [dc_name, dc_cad, dc_x1, dc_y1, dc_x2, dc_y2, dc_no]
    min_col_drain = min(target_cols_drain)
    max_col_drain = max(target_cols_drain)
    col_span_drain = max_col_drain - min_col_drain + 1

    drain_block = []
    n_drains_written = 0
    for i, d in enumerate(drain_data):
        if n_drains_written >= DRAIN_MAX:
            break
        row_vals = [None] * col_span_drain
        row_vals[dc_name - min_col_drain]  = d['name']
        row_vals[dc_cad - min_col_drain]   = d['cad_layer']
        row_vals[dc_x1 - min_col_drain]    = d['x1']
        row_vals[dc_y1 - min_col_drain]    = d['y1']
        row_vals[dc_x2 - min_col_drain]    = d['x2']
        row_vals[dc_y2 - min_col_drain]    = d['y2']
        row_vals[dc_no - min_col_drain]    = i + 1
        drain_block.append(row_vals)
        n_drains_written += 1

    if drain_block:
        n_rows_drain = len(drain_block)
        sheet.range((DRAIN_ROW, min_col_drain), (DRAIN_ROW + n_rows_drain - 1, max_col_drain)).value = drain_block

    # Clear leftover drain rows
    clear_start_drain = DRAIN_ROW + n_drains_written
    clear_end_drain = DRAIN_ROW + DRAIN_MAX - 1
    if clear_start_drain <= clear_end_drain:
        sheet.range((clear_start_drain, min_col_drain), (clear_end_drain, max_col_drain)).value = \
            [[None] * col_span_drain for _ in range(clear_end_drain - clear_start_drain + 1)]

    if len(drain_data) > DRAIN_MAX:
        print(f"  WARNING: {len(drain_data)} drains in DXF but limit is {DRAIN_MAX} — truncated")
    print(f"  drains: {n_drains_written} PVD elements (limit={DRAIN_MAX}) → M{DRAIN_ROW}:S{DRAIN_ROW + n_drains_written - 1}")

    # ── Waterlines (columns U-AA, header row 234, data row 235+) ──
    WATERLINE_ROW = WATERLINE_DATA_START
    # Read max rows from Z233; default 125
    _raw_waterline_max = sheet.range(WATERLINE_MAX_CELL).value
    try:
        WATERLINE_MAX = max(1, int(_raw_waterline_max))
    except (TypeError, ValueError):
        WATERLINE_MAX = 125

    # Unhide the complete configured data range before writing.  SC5 may have
    # hidden rows in this shared section.
    for r in range(WATERLINE_ROW, WATERLINE_ROW + WATERLINE_MAX):
        sheet.api.Rows(r).Hidden = False

    _set_progress(wb, f"sc4: {_spin()} WORKING — waterlines {_bar(0.92)}")
    waterline_data = collect_waterline_rows(msp)
    gwflowbc_data = collect_gwflowbc_rows(msp)

    # Build the column map only from the waterline section.  This avoids
    # collisions with the soil-polygon and drain headers on the same row.
    waterline_col_map = {}
    for col in range(21, 29):  # U=21 to AB=28
        val = sheet.range(WATERLINE_HEADER_ROW, col).value
        if val is not None:
            waterline_col_map[str(val).strip()] = col

    wc_name = waterline_col_map.get('waterline_name', 21)
    wc_cad = waterline_col_map.get('cad_layer', 22)
    wc_x1 = waterline_col_map.get('x1', 23)
    wc_y1 = waterline_col_map.get('y1', 24)
    wc_x2 = waterline_col_map.get('x2', 25)
    wc_y2 = waterline_col_map.get('y2', 26)
    wc_no = waterline_col_map.get('waterline_no', 27)
    gwc_no = waterline_col_map.get('GWFlowBC_no', 28)
    all_cols = [wc_name, wc_cad, wc_x1, wc_y1, wc_x2, wc_y2, wc_no, gwc_no]
    min_col = min(all_cols)
    max_col = max(all_cols)
    col_span = max_col - min_col + 1

    # Clear all SC4-owned waterline + GWFlowBC cells before writing.
    sheet.range(
        (WATERLINE_ROW, min_col),
        (WATERLINE_ROW + WATERLINE_MAX - 1, max_col),
    ).value = [[None] * col_span for _ in range(WATERLINE_MAX)]

    # Build stacked block: waterlines first, then GWFlowBCs.
    combined_block = []
    for row_data in waterline_data[:WATERLINE_MAX]:
        name, cad_layer, x1, y1, x2, y2, waterline_no = row_data
        rv = [None] * col_span
        rv[wc_name - min_col] = name
        rv[wc_cad - min_col] = cad_layer
        rv[wc_x1 - min_col] = x1
        rv[wc_y1 - min_col] = y1
        rv[wc_x2 - min_col] = x2
        rv[wc_y2 - min_col] = y2
        rv[wc_no - min_col] = waterline_no
        combined_block.append(rv)

    remaining = WATERLINE_MAX - len(combined_block)
    for row_data in gwflowbc_data[:remaining]:
        name, cad_layer, x1, y1, x2, y2, g_no = row_data
        rv = [None] * col_span
        rv[wc_name - min_col] = name
        rv[wc_cad - min_col] = cad_layer
        rv[wc_x1 - min_col] = x1
        rv[wc_y1 - min_col] = y1
        rv[wc_x2 - min_col] = x2
        rv[wc_y2 - min_col] = y2
        rv[gwc_no - min_col] = g_no
        combined_block.append(rv)

    n_written = len(combined_block)
    if combined_block:
        sheet.range(
            (WATERLINE_ROW, min_col),
            (WATERLINE_ROW + n_written - 1, max_col),
        ).value = combined_block

    n_wl = min(len(waterline_data), WATERLINE_MAX)
    n_gw = min(len(gwflowbc_data), max(0, WATERLINE_MAX - len(waterline_data)))
    if n_written < len(waterline_data) + len(gwflowbc_data):
        print(
            f"  WARNING: {len(waterline_data) + len(gwflowbc_data)} waterline+"
            f"GWFlowBC rows but limit is {WATERLINE_MAX} — truncated"
        )
    print(
        f"  waterlines: {n_wl} segments "
        f"({len(set(r[-1] for r in waterline_data))} polylines) + "
        f"GWFlowBC: {n_gw} "
        f"→ U{WATERLINE_ROW}:AD{WATERLINE_ROW + n_written - 1}"
    )


# ── Line Loads & Point Loads (header row 18, data rows 19-38) ──
    LOAD_HDR_ROW = 18
    LOAD_DATA_START = 19
    LOAD_DATA_END = 38
    LOAD_MAX = LOAD_DATA_END - LOAD_DATA_START + 1  # 20 rows

    # Build column map from header row 18 (columns B-K only)
    load_col_map = {}
    for col in range(2, 12):  # B=2 to K=11
        val = sheet.range(LOAD_HDR_ROW, col).value
        if val is not None:
            load_col_map[str(val).strip()] = col

    lc_name = load_col_map.get('Loads', 2)
    lc_cad  = load_col_map.get('cad_layer', 3)
    lc_x1   = load_col_map.get('x1', 4)
    lc_y1   = load_col_map.get('y1', 5)
    lc_x2   = load_col_map.get('x2', 6)
    lc_y2   = load_col_map.get('y2', 7)
    lc_lno  = load_col_map.get('load_no', 8)
    lc_lmod = load_col_map.get('load_model', 9)
    lc_qx   = load_col_map.get('qx_start', 10)
    lc_qy   = load_col_map.get('qy_start', 11)

    # Snapshot existing loads BEFORE clearing — used for geometry+layer
    # preservation (same mechanism as str_volume column-H preservation).
    _set_progress(wb, f"sc4: {_spin()} WORKING — loads {_bar(0.96)}")
    existing_loads = {}  # {(geo_key, cad_layer): {load_no, load_model, qx, qy}}
    for r in range(LOAD_DATA_START, LOAD_DATA_END + 1):
        _ex_x1   = sheet.range(r, lc_x1).value
        _ex_y1   = sheet.range(r, lc_y1).value
        if None in (_ex_x1, _ex_y1):
            continue
        _ex_cad  = sheet.range(r, lc_cad).value
        _ex_x2   = sheet.range(r, lc_x2).value
        _ex_y2   = sheet.range(r, lc_y2).value
        _lno  = sheet.range(r, lc_lno).value
        _lmod = sheet.range(r, lc_lmod).value
        _qx   = sheet.range(r, lc_qx).value
        _qy   = sheet.range(r, lc_qy).value
        try:
            _gk = _load_geometry_key(_ex_x1, _ex_y1, _ex_x2, _ex_y2)
        except (TypeError, ValueError):
            continue  # junk coordinate text — row cannot match, treat as new
        _cad = str(_ex_cad).strip() if _ex_cad else ''
        existing_loads[(_gk, _cad)] = {
            'load_no':  _lno,
            'load_model': str(_lmod).strip() if _lmod else '',
            'qx_start': _qx if _qx is not None else 0,
            'qy_start': _qy if _qy is not None else 0,
        }

    # Unhide and fully clear the load table before writing new DXF data.
    for r in range(LOAD_DATA_START, LOAD_DATA_END + 1):
        sheet.api.Rows(r).Hidden = False
        for col in range(2, 12):  # B through K
            sheet.range(r, col).value = None

    # Collect load entities from DXF
    _load_raw = []  # [(layer_name, magnitude, x1, y1, x2, y2, load_model)]
    for entity in msp:
        layer = entity.dxf.layer
        layer_lower = layer.lower()
        etype = entity.dxftype()

        if layer_lower.startswith('load_line'):
            mag = _parse_load_magnitude(layer)

            if etype == 'LINE':
                _load_raw.append((layer, mag,
                    *extract_line_coords(entity),
                    'line_load'))
            elif etype == 'LWPOLYLINE':
                verts = extract_polyline_coords(entity)
                for i in range(len(verts) - 1):
                    _load_raw.append((layer, mag,
                        verts[i][0], verts[i][1],
                        verts[i+1][0], verts[i+1][1],
                        'line_load'))

        elif layer_lower.startswith('load_point'):
            mag = _parse_load_magnitude(layer)

            if etype == 'POINT':
                x, y, _, _ = extract_point_coords(entity)
                _load_raw.append((layer, mag,
                    x, y, None, None,
                    'point_load'))

    # Sort: line loads first (by y descending), then point loads (by y descending)
    _line_loads = sorted([l for l in _load_raw if l[6] == 'line_load'],
                         key=lambda x: -x[3])  # sort by y1 desc
    _point_loads = sorted([l for l in _load_raw if l[6] == 'point_load'],
                          key=lambda x: -x[3])  # sort by y1 desc
    _loads_sorted = _line_loads + _point_loads

    # ── Batch-write load data ──
    target_cols_load = [lc_name, lc_cad, lc_x1, lc_y1, lc_x2, lc_y2, lc_lno, lc_lmod, lc_qx, lc_qy]
    min_col_load = min(target_cols_load)
    max_col_load = max(target_cols_load)
    col_span_load = max_col_load - min_col_load + 1

    load_block = []
    n_loads_written = 0
    for layer_name, mag, x1, y1, x2, y2, load_model in _loads_sorted:
        if n_loads_written >= LOAD_MAX:
            break
        row_vals = [None] * col_span_load
        row_vals[lc_cad - min_col_load]  = layer_name
        row_vals[lc_x1 - min_col_load]   = x1
        row_vals[lc_y1 - min_col_load]   = y1
        if load_model == 'line_load':
            row_vals[lc_x2 - min_col_load] = x2
            row_vals[lc_y2 - min_col_load] = y2
        else:
            row_vals[lc_x2 - min_col_load] = None
            row_vals[lc_y2 - min_col_load] = None
        row_vals[lc_lmod - min_col_load] = load_model
        preserved = _preserved_load_inputs(
            existing_loads,
            _load_geometry_key(x1, y1, x2, y2),
            layer_name,
        )
        if preserved is not None:
            row_vals[lc_lno - min_col_load], \
            row_vals[lc_lmod - min_col_load], \
            row_vals[lc_qx - min_col_load], \
            row_vals[lc_qy - min_col_load] = preserved
        else:
            row_vals[lc_qx - min_col_load] = 0
            row_vals[lc_qy - min_col_load] = -mag  # negative = gravity direction
        load_block.append(row_vals)
        n_loads_written += 1

    if load_block:
        n_rows_load = len(load_block)
        sheet.range((LOAD_DATA_START, min_col_load), (LOAD_DATA_START + n_rows_load - 1, max_col_load)).value = load_block

    # Clear leftover load rows
    clear_start_load = LOAD_DATA_START + n_loads_written
    clear_end_load = LOAD_DATA_END
    if clear_start_load <= clear_end_load:
        sheet.range((clear_start_load, min_col_load), (clear_end_load, max_col_load)).value = \
            [[None] * col_span_load for _ in range(clear_end_load - clear_start_load + 1)]

    if len(_loads_sorted) > LOAD_MAX:
        print(f"  WARNING: {len(_loads_sorted)} loads in DXF but limit is {LOAD_MAX} — truncated")
    print(f"  loads: {n_loads_written} elements ({len(_line_loads)} line, {len(_point_loads)} point) → row {LOAD_DATA_START}+")

    print("sc4: Done. Fill yellow columns, then press Button 2.")
    _n_sc4 = sum(len(v) for v in elements.values())
    _set_progress(wb, f"sc4: done — {_n_sc4} elements {_bar(1.0)}")


if __name__ == '__main__':
    main()
