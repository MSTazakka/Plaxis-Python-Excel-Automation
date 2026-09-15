# 🏗️ Plaxis Python-Excel Automation (v0.8.5-PUBLIC)

Excel-driven automation for building PLAXIS 2D models: soil notebooks, DXF import,
structures, staged construction, and remesh support.

> ⚠️ **DISCLAIMER — READ FIRST.** Shared for **educational purposes only**, as-is with
> **no warranty**. The authors accept **no responsibility for bugs, wrong results,
> or any damages**. You are solely responsible for verifying every model against
> the PLAXIS GUI and exercising independent engineering judgment. If you do not
> accept this, do not use these files. See `MANUAL.md` (top disclaimer + Sections
> F, H, K–L, P) for the full terms, including the three modelling simplifications
> (stacked-polygon DXF, illustrative material values, wish-in-place concrete).

## 📦 What's inside

| Path | What it is |
|---|---|
| 📗 `EXAMPLE 1.xlsm` … `EXAMPLE 4.xlsm` | Excel front ends (soil profile, geometry, structures, phases) |
| 📐 `Example * - *.dxf` | DXF geometry per example (the `.bak` files are local backups — **do not upload**) |
| 🐍 `scripts/` | `sc3_PythonPlaxis.py`, `sc4_dxf_reader.py`, `sc5_structural_to_ipynb.py`, `sc6_staged_construction.py`, `sc7_curve_points.py`, `sc9_IPYNB_LAUNCHER.py`, `notebook_paths.py`, `tests/` |
| 📖 `MANUAL.md` | Full public manual (lettered A–Y + appendices S–Y) |
| 🛠️ `INSTALL.md` | Python-from-zero install guide (PATH, pip, xlwings add-in, PLAXIS server) |
| 📄 `LICENSE` | MIT licence (educational, as-is) |

Deliberately **excluded**: soil-parameter fitting (SC1), command-script generator
(SC2), post-calculation extraction (SC8). The chain is SC3 → SC4 → SC5 → SC6 → SC9
plus SC7 remesh support. Post-calc review is manual in PLAXIS Output.

## ✅ Requirements

- PLAXIS 2D 2025 with the Python scripting server
- Desktop Excel with macros enabled + xlwings VBA add-in
- Python with `xlwings`, `pandas`, `numpy`, `ezdxf` (see `INSTALL.md`)

## 📐 DXF drawing rules — read before you draw

Your DXF decides whether the model builds cleanly. Two rules, no exceptions:

1. **Correct layer names (`MANUAL.md` Section T).** Format is `<type>[_<number>][bracket]` — e.g. `plate_1[D-Wall 800mm]`, `anc_1[grout 300mm]`, `load_line[100kPa]`, `cut`, `fill`, `volume_profile`. Matching is case-insensitive prefix, except three exact-match layers: `pvd` (`pvd_1` is ignored), `waterboundary_head`, `waterboundary_closed`. Wrong names = silently skipped geometry.

2. **Well snapped, no micro-gaps.** Draw closed polylines, snap every endpoint, no duplicate lines, no hairline gaps or slivers along shared edges. The chain only helps so far: 1 mm rounding, 10 mm axis-snap on structural lines, and a merge-equivalents cell that collapses near-duplicates — but stacked soil polygons are still separate polylines sharing an edge twice, not true shared topology. Leftover gaps change stress flow, interfaces, and mesh quality. After SC4/SC9, inspect every shared edge and the mesh around it (`MANUAL.md` H.8 checklist).

## 🚀 Quick start

1. Read `INSTALL.md` once (install + smoke test).
2. Start the PLAXIS remote scripting server (once per PLAXIS session): Input → **Expert → Configure remote scripting server…** → port `10000` + password → **Start server** (screenshot in `MANUAL.md` C.3); mirror the same port/password into `main!V5`/`main!V6` and keep Input open.
3. Open an `EXAMPLE *.xlsm`, fill `main` (`B2:D4`, `F2:H8`, `B12:AH33`).
4. Run SC3 (Plaxis Python) → SC4 (DXF import) → SC5 (structures) → SC6 (staging) → SC9 (launcher), then mesh + calculate in PLAXIS.
5. Full procedure, range map, and audit checklist: `MANUAL.md` Sections D, S, P.

## 🔍 Full audit trail

Every press appends plain-Python cells to a workbook-scoped `.ipynb` beside the
workbook (SC3 creates it, SC5/SC6 append, SC9 runs it inside PLAXIS), and SC7
writes its own timestamped `curve_<stem>.ipynb` transcript. The entire model
build is inspectable, diffable, and replayable — hand the notebook to any
checker.

## 🚫 Do NOT upload

- `*.bak` (local DXF backups), `EXAMPLE 1.ipynb` (generated notebook with your local PLAXIS password), `.pytest_cache/`, `__pycache__/` — all covered by `.gitignore`, but browser upload takes what you drag, so leave them out.
