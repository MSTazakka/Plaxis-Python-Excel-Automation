"""
sc3_PythonPlaxis.py — PLAXIS Python API Script Generator
=========================================================
Reads Excel input (same source as sc1/sc2) and generates a Jupyter notebook
(.ipynb) for use inside PLAXIS's Jupyter console.

Scope: soil geometry + material inputs only. No meshing or structural elements.

Output: .ipynb file saved next to the Excel workbook.
"""

import pandas as pd
import numpy as np
from pathlib import Path
import xlwings as xw
import json
import re
from math import isnan

try:
    from .notebook_paths import notebook_path_for_workbook
except ImportError:  # pragma: no cover - supports direct script execution
    from notebook_paths import notebook_path_for_workbook


# ── Soil type mapping ──
SOIL_TYPE_MAP = {
    'clay-HS':   ('Hardening Soil',  'Undrained A',  False, False, 'Clay'),
    'sand-HS':   ('Hardening Soil',  'Drained',      False, False, 'Sand'),
    'silt-HS':   ('Hardening Soil',  'Undrained A',  False, False, 'Silt'),
    'clay-HSS':  ('HS Small',        'Undrained A',  True,  False, 'Clay'),
    'sand-HSS':  ('HS Small',        'Drained',      True,  False, 'Sand'),
    'silt-HSS':  ('HS Small',        'Undrained A',  True,  False, 'Silt'),
    'clay-SSC':  ('Soft Soil Creep', 'Undrained A',  False, True,  'Clay'),
    'sand-SSC':  ('Soft Soil Creep', 'Drained',      False, True,  'Sand'),
    'silt-SSC':  ('Soft Soil Creep', 'Undrained A',  False, True,  'Silt'),
    'linear-elastic': ('Linear Elastic', 'Non-porous', False, False, None),
}


def safe_zero(val):
    if val is None:
        return 0
    try:
        if isnan(val):
            return 0
    except (TypeError, ValueError):
        pass
    return val


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
        if pd.isna(mode):
            return 'automatic'
    except (TypeError, ValueError):
        pass
    m = str(mode).strip()
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


def _fmt_param(p):
    """Format one soilmat element: quoted strings pass through, numerics go
    through round_num (kills binary residue like 1.5000000000000004 so the
    notebook never carries 17-digit floats). Anything unformattable falls
    back to str, matching the old output."""
    if isinstance(p, str) and p.startswith('"'):
        return p
    try:
        return round_num(p)
    except (TypeError, ValueError, OverflowError):
        return str(p)


def make_cell(source, cell_type='code'):
    """Create a Jupyter notebook cell."""
    return {
        'cell_type': cell_type,
        'metadata': {},
        'source': source if isinstance(source, list) else [source],
        'outputs': [],
        'execution_count': None,
    }


def make_md_cell(source):
    return make_cell(source, cell_type='markdown')


def main():
    """Read Excel and generate PLAXIS Python API Jupyter notebook."""

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

    sheet_main = wb.sheets['main']

    # ── Read input data ──
    initial1 = sheet_main['B2:D7'].options(pd.DataFrame, index=True).value.reset_index().dropna(subset=['description'])
    initial2 = sheet_main['F2:H8'].options(pd.DataFrame, index=True).value.reset_index().dropna(subset=['description'])
    df_initial = pd.concat([initial1, initial2])
    df_soilprofile = sheet_main['B12:AH33'].options(pd.DataFrame, index=True).value.reset_index().dropna(subset=['soil_layers'])

    df_initial[''] = 'value'
    df_initial = df_initial.pivot(index='', columns='description', values='value')

    plaxis_type = df_initial['plaxis_type'].iloc[0]
    surface_level = df_initial['surface_level'].iloc[0]
    x_min = df_initial['x_min'].iloc[0]
    y_min = df_initial['y_min'].iloc[0]
    x_max = df_initial['x_max'].iloc[0]
    y_max = df_initial['y_max'].iloc[0]
    water_elevation = df_initial['water_level'].iloc[0]
    n_soil = len(df_soilprofile['soil_layers'])

    # Server connection params
    plx_port = int(sheet_main['V5'].value)
    plx_password = str(sheet_main['V6'].value)

    # Model type: map dropdown (2D_PS / 2D_axis / 3D) to PLAXIS API type
    MODEL_TYPE_MAP = {
        '2D_PS':   'PlaneStrain',
        '2D_axis': 'Axisymmetric',
        '3D':      '3D',
    }
    plx_model_type = MODEL_TYPE_MAP.get(plaxis_type, 'PlaneStrain')
    is_3d = plaxis_type == '3D'

    df_soilprofile['thickness'] = round(df_soilprofile['depth_bot'] - df_soilprofile['depth_top'], 3)

    # NOTE: clay_frac/silt_frac are deliberately NOT zeroed here — a genuinely
    # blank AE/AF cell means "omit the fraction line so PLAXIS falls back to
    # the USDA class default" (see _gw_frac_is_filled in the groundwater block).
    # NOTE: POP is deliberately NOT zeroed here either — HS/HSS/SSC omit
    # blank OCR/POP/einit (see the branches); SSC lambda/kappa/mu still
    # emit raw as before.

    # ── Build notebook cells ──
    cells = []

    # Title
    cells.append(make_md_cell([
        '# PLAXIS Python API — Soil Geometry & Materials\n',
        '> Auto-generated by `sc3_PythonPlaxis.py`  \n',
        '> Scope: soil geometry + material inputs (no mesh / structures)'
    ]))

    # Connection cell
    cells.append(make_md_cell(['## 1. Connection to PLAXIS']))
    cells.append(make_cell([
        'from plxscripting.easy import *\n',
        '\n',
        's_i, g_i = new_server("localhost", port=' + str(plx_port) + ', password="' + plx_password + '")\n',
        'print("s_i and g_i ready.")'
    ]))

    # New model
    cells.append(make_md_cell(['## 2. New Model']))
    # ModelType enum: 0 = Plane strain, 1 = Axisymmetric, 2 = 3D
    MODEL_TYPE_VAL = {'PlaneStrain': 0, 'Axisymmetric': 1, '3D': 2}
    model_type_val = MODEL_TYPE_VAL.get(plx_model_type, 0)
    cells.append(make_cell([
        's_i.new()\n',
        f'g_i.Project.ModelType.set({model_type_val})'
    ]))

    # Soil contour
    cells.append(make_md_cell(['## 3. Soil Contour']))
    cells.append(make_cell([
        f'g_i.SoilContour.initializerectangular({x_min}, {y_min}, {x_max}, {y_max})'
    ]))

    # Borehole
    cells.append(make_md_cell(['## 4. Borehole']))
    if not is_3d:
        bh_code = 'borehole_g = g_i.borehole(0)'
    else:
        bh_code = 'borehole_g = g_i.borehole(0, 0)'
    cells.append(make_cell([
        bh_code + '\n',
        f'borehole_g.Head.set({water_elevation})'
    ]))

    # Soil layers
    cells.append(make_md_cell(['## 5. Soil Layers']))
    slines = []
    for i in range(n_soil):
        slines.append(f'g_i.soillayer({df_soilprofile["thickness"].iloc[i]})')
    cells.append(make_cell('\n'.join(slines)))

    # Material definitions
    cells.append(make_md_cell(['## 6. Material Definitions']))

    for i in range(n_soil):
        soil_name = df_soilprofile['soil_layers'].iloc[i]
        name_formatted = "".join(e for e in soil_name.title() if e.isalnum())
        soil_type = df_soilprofile["type"].iloc[i]
        model_name, drainage, is_hss, is_ssc, usda_class = SOIL_TYPE_MAP.get(
            soil_type, ('Hardening Soil', 'Drained', False, False, 'Sand')
        )

        gammaUnsat = df_soilprofile["gunsat"].iloc[i]
        gammaSat = df_soilprofile["gsat"].iloc[i]
        E50Ref = df_soilprofile["E50_ref"].iloc[i]
        PowerM = df_soilprofile["m"].iloc[i]
        cRef = df_soilprofile["c"].iloc[i]
        phi = df_soilprofile["phi"].iloc[i]
        psi = df_soilprofile["psi"].iloc[i]
        pref = df_soilprofile["pref"].iloc[i]
        Rinter = df_soilprofile["Rinter"].iloc[i]
        ClayFraction = df_soilprofile["clay_frac"].iloc[i]
        SiltFraction = df_soilprofile['silt_frac'].iloc[i]
        mode = _normalise_gw_mode(df_soilprofile['mode'].iloc[i])
        k_ver = df_soilprofile['k_ver'].iloc[i]
        k_ratio = df_soilprofile['k_hor_ratio'].iloc[i]
        if pd.isna(k_ratio) or str(k_ratio).strip() == '':
            k_ratio = 1.5

        # ── Linear Elastic: only ERef and nu ──
        if soil_type == 'linear-elastic':
            # Parse names such as Concrete_E20GPa_nu0.2.
            # PLAXIS stiffness input is kN/m² = kPa: 1 GPa = 1,000,000 kPa.
            e_match = re.search(r'E([\d.]+)GPa', str(soil_name), re.IGNORECASE)
            nu_match = re.search(r'(?:nu|v)([\d.]+)', str(soil_name), re.IGNORECASE)
            name_e_ref = (float(e_match.group(1)) * 1_000_000
                          if e_match else 0)
            name_nu = float(nu_match.group(1)) if nu_match else 0
            raw_e_ref = df_soilprofile["E50_ref"].iloc[i]
            raw_nu = df_soilprofile["psi"].iloc[i]
            ERef = name_e_ref or (0 if pd.isna(raw_e_ref) else raw_e_ref) or 30_000_000
            nu = name_nu or (0 if pd.isna(raw_nu) else raw_nu) or 0.15
            mat_params = [
                '"Identification"', fmt(soil_name),
                '"SoilModel"', fmt(model_name),
                '"DrainageType"', fmt(drainage),
                '"gammaUnsat"', gammaUnsat,
                '"ERef"', ERef,
                '"nu"', nu,
            ]
            cells.append(make_cell([
                f'# Layer {i + 1}: {soil_name} ({soil_type})\n',
                f'{name_formatted} = g_i.soilmat({", ".join(_fmt_param(p) for p in mat_params)})'
            ]))
            continue

        mat_params = [
            '"Identification"', fmt(soil_name),
            '"SoilModel"', fmt(model_name),
            '"DrainageType"', fmt(drainage),
            '"gammaUnsat"', gammaUnsat,
            '"gammaSat"', gammaSat,
        ]

        EoedRef = df_soilprofile["Eoed_ref"].iloc[i]
        EurRef = df_soilprofile["Eur_ref"].iloc[i]

        # HS/HSS overconsolidation: each of OCR/POP/einit emits independently
        # iff its own cell is filled (same blank gate as the groundwater
        # fractions). Blank omits that pair; filled emits the input value.
        # SSC below gates eInit/OCR/POP the same way (blank omits, so no
        # raw-nan NameError); lambda/kappa/mu always emit; POP stays last
        # so a both-filled row keeps POP trailing.
        ocr_pep = []
        if _gw_frac_is_filled(df_soilprofile["OCR"].iloc[i]):
            ocr_pep.append(('"OCR"', df_soilprofile['OCR'].iloc[i]))
        if _gw_frac_is_filled(df_soilprofile["POP"].iloc[i]):
            ocr_pep.append(('"POP"', df_soilprofile['POP'].iloc[i]))
        if _gw_frac_is_filled(df_soilprofile["einit"].iloc[i]):
            ocr_pep.append(('"eInit"', df_soilprofile["einit"].iloc[i]))

        if is_hss:
            mat_params.extend([
                '"E50Ref"', E50Ref, '"EoedRef"', EoedRef, '"EurRef"', EurRef,
                '"PowerM"', PowerM, '"pRef"', pref,
                '"G0Ref"', df_soilprofile["G0_ref"].iloc[i],
                '"gamma07"', df_soilprofile["g07"].iloc[i],
            ])
            for _k, _v in ocr_pep:
                mat_params.extend([_k, _v])
        elif is_ssc:
            if _gw_frac_is_filled(df_soilprofile["einit"].iloc[i]):
                mat_params.extend(['"eInit"', df_soilprofile["einit"].iloc[i]])
            if _gw_frac_is_filled(df_soilprofile["OCR"].iloc[i]):
                mat_params.extend(['"OCR"', df_soilprofile['OCR'].iloc[i]])
            mat_params.extend([
                '"lambdaModified"', df_soilprofile["lambda"].iloc[i],
                '"kappaModified"', df_soilprofile["kappa"].iloc[i],
                '"muModified"', df_soilprofile["mu"].iloc[i],
            ])
            if _gw_frac_is_filled(df_soilprofile["POP"].iloc[i]):
                mat_params.extend(['"POP"', df_soilprofile['POP'].iloc[i]])
        else:
            mat_params.extend([
                '"E50Ref"', E50Ref, '"EoedRef"', EoedRef, '"EurRef"', EurRef,
                '"PowerM"', PowerM, '"pRef"', pref,
            ])
            for _k, _v in ocr_pep:
                mat_params.extend([_k, _v])

        mat_params.extend([
            '"cRef"', cRef, '"phi"', phi, '"psi"', psi,
            '"InterfaceStrengthDetermination"', '"Manual"', '"Rinter"', Rinter,
        ])

        params_str = ", ".join(_fmt_param(p) for p in mat_params)

        lines = [
            f'# Layer {i + 1}: {soil_name} ({soil_type})\n',
            f'{name_formatted} = g_i.soilmat({params_str})\n',
        ]

        # Groundwater — all modes are USDA. The class string comes from the
        # SOIL_TYPE_MAP 5th element only (capitalized Clay/Sand/Silt — never
        # sliced from the soil-type prefix; PLAXIS .set() is case-sensitive).
        # A genuinely blank fraction cell omits that fraction line so PLAXIS
        # falls back to the class default; an explicit 0 still emits.
        if mode == 'automatic':
            lines.append(f'{name_formatted}.GroundwaterClassificationType.set("USDA")\n')
            lines.append(f'{name_formatted}.GroundwaterSoilClassUSDA.set("{usda_class}")\n')
            lines.append(f'{name_formatted}.GwUseDefaults.set(True)\n')
            lines.append(f'{name_formatted}.GwDefaultsMethod.set("From grain size distribution")')
        elif mode == 'by_grain_size':
            lines.append(f'{name_formatted}.GroundwaterClassificationType.set("USDA")\n')
            lines.append(f'{name_formatted}.GroundwaterSoilClassUSDA.set("{usda_class}")\n')
            if _gw_frac_is_filled(ClayFraction):
                lines.append(f'{name_formatted}.ClayFraction.set({ClayFraction})\n')
            if _gw_frac_is_filled(SiltFraction):
                lines.append(f'{name_formatted}.SiltFraction.set({SiltFraction})\n')
            lines.append(f'{name_formatted}.GwUseDefaults.set(True)\n')
            lines.append(f'{name_formatted}.GwDefaultsMethod.set("From grain size distribution")')
        elif mode == 'manual':
            lines.append(f'{name_formatted}.GroundwaterClassificationType.set("USDA")\n')
            lines.append(f'{name_formatted}.GroundwaterSoilClassUSDA.set("{usda_class}")\n')
            if _gw_frac_is_filled(ClayFraction):
                lines.append(f'{name_formatted}.ClayFraction.set({ClayFraction})\n')
            if _gw_frac_is_filled(SiltFraction):
                lines.append(f'{name_formatted}.SiltFraction.set({SiltFraction})\n')
            lines.append(f'{name_formatted}.GwUseDefaults.set(False)\n')
            lines.append(f'{name_formatted}.PermHorizontalPrimary.set({k_ver * k_ratio})\n')
            if is_3d:
                lines.append(f'{name_formatted}.PermHorizontalSecondary.set({k_ver * k_ratio})\n')
            lines.append(f'{name_formatted}.PermVertical.set({k_ver})')

        cells.append(make_cell(lines))

    # Material assignment
    cells.append(make_md_cell(['## 7. Assign Materials to Soil Layers']))
    assign_lines = []
    for i in range(n_soil):
        soil_name = df_soilprofile['soil_layers'].iloc[i]
        name_formatted = "".join(e for e in soil_name.title() if e.isalnum())
        assign_lines.append(f'g_i.Soils[{i}].Material = {name_formatted}')
    cells.append(make_cell('\n'.join(assign_lines)))

    # Soil layer levels
    cells.append(make_md_cell(['## 8. Soil Layer Elevations']))
    elev_lines = []
    top_elev = round(surface_level + df_soilprofile["depth_top"].iloc[0], 3)
    bot_elev_0 = round(surface_level - df_soilprofile["depth_bot"].iloc[0], 3)
    elev_lines.append(f'g_i.setsoillayerlevel(borehole_g, 0, {top_elev})')
    elev_lines.append(f'g_i.setsoillayerlevel(borehole_g, 1, {bot_elev_0})')
    for i in range(1, n_soil):
        bot_elev = round(surface_level - df_soilprofile["depth_bot"].iloc[i], 3)
        elev_lines.append(f'g_i.setsoillayerlevel(borehole_g, {i + 1}, {bot_elev})')
    cells.append(make_cell('\n'.join(elev_lines)))

    # Done
    cells.append(make_md_cell(['## Done']))
    cells.append(make_cell([
        'print("PLAXIS model setup complete — soil geometry + materials.")'
    ]))

    # ── Assemble notebook ──
    notebook = {
        'nbformat': 4,
        'nbformat_minor': 5,
        'metadata': {
            'kernelspec': {
                'display_name': 'Python 3',
                'language': 'python',
                'name': 'python3',
            },
            'language_info': {
                'name': 'python',
                'version': '3.11.0',
            },
        },
        'cells': cells,
    }

    # ── Write .ipynb ──
    out_path = notebook_path_for_workbook(wb)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(notebook, f, indent=1, ensure_ascii=False)

    print(f"sc3: Wrote notebook to {out_path}")


if __name__ == '__main__':
    main()
