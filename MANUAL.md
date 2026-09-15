# PLAXIS Python Automation — Public Manual (v0.8.5-PUBLIC)

> **DISCLAIMER — READ FIRST.** This bundle is shared for **educational purposes only**.
> The scripts are provided **as-is, with no warranty of any kind**. The authors accept
> **no responsibility for any bugs, errors, wrong results, model failures, or any direct
> or indirect damages** caused by using these scripts — including but not limited to
> engineering decisions made from their output. **You are solely responsible for
> verifying every model against the PLAXIS GUI, checking inputs, meshes, phases, and
> results, and exercising independent engineering judgment before any real-world use.**
> If you do not accept this, do not use these files.
>
> This bundle additionally simplifies reality in three ways you must accept before
> use: **(1)** soil bodies are drawn as **stacked closed polygons sharing edges** —
> a drawing convenience that can leave duplicate lines, small gaps, or slivers and
> can affect result accuracy (Section H); **(2)** every material value in the
> example workbooks is **illustrative, not a real-world or recommended value**
> (Section F); **(3)** construction staging is **simplified — concrete appears
> wish-in-place at full stiffness; pour sequence and early-age behaviour are not
> modelled** (Sections K–L). All three are detailed where they apply and must be
> closed out in the Section P audit before any delivery.

Software project: `02_Plaxis_Python` — public pipeline version **v0.8.5-PUBLIC**
(public fork of the internal v0.8.5 chain; soil-parameter fitting and command-script
utilities are **not** included).

---

## A. Document control

### A.1 Purpose and scope

This manual explains how to build a PLAXIS 2D model end to end using the Excel front
end and the Python automation chain in this folder. It covers scripts **SC3–SC7 + SC9 only** (SC8 excluded),
every Excel input the user must maintain, every generated artifact, and the
verification steps required before a model may be trusted.

Two readers:

1. **The modeller**, who drives the workflow entirely from Excel and DXF files.
   Sections B–F, J, K, M, P, Q, and R are sufficient.
2. **The reviewer or maintainer**, who must audit what each script does, what it may
   overwrite, and where every value came from. Sections G–I, L, N, O, P, the
   appendices, and the audit block in Section P serve this reader.

What is **not** in this public bundle (deliberately excluded):

- Soil-parameter derivation from SPT or any curve-fitting workflow. Fill all soil
  and stiffness parameters **directly** in the `main` sheet (Section F).
- The optional PLAXIS command-script generator. The notebook chain (SC3 → SC9) is
  the only supported path.
- Tunnel-liner, TBM, contraction, CenterLine/volume-liner, and radial-picture
  documentation. The scripts are unchanged from the internal chain and may contain
  code paths beyond what this manual describes; anything not documented here is
  **unsupported in this public release** — use it only at your own risk, if you can
  work it out yourself.

### A.2 Target software versions

| Component | Version tested and supported |
|---|---|
| PLAXIS 2D | 2025 |
| Excel | Desktop Excel with macros enabled; xlwings VBA add-in installed |
| Python | The xlwings-managed interpreter; pandas optional but recommended |

### A.3 Coverage of scripts SC3–SC7 + SC9 (SC8 excluded)

| Script | File | Role | In v0.8.5-PUBLIC |
|---|---|---|---|
| SC1 | (not shipped — excluded from this public bundle)  | Soil parameters determination module | No |
| SC2 | (not shipped — excluded from this public bundle)  | Command control / Multiple boreholes support | No |
| SC3 | `scripts/sc3_PythonPlaxis.py` | Soil/material notebook generator | Yes |
| SC4 | `scripts/sc4_dxf_reader.py` | DXF import into `str_2D` + curve points to `output` tab + `str_2D!B4` progress bar | Yes |
| SC5 | `scripts/sc5_structural_to_ipynb.py` | Structural elements, interfaces, `assign_sequence`, merge-equivalents cell, mesh + output forces + `str_2D!D4` progress bar | Yes |
| SC6 | `scripts/sc6_staged_construction.py` | Staged construction + water conditions + `assign_sequence!B5` progress, warnings in `assign_sequence!B4` | Yes |
| SC7 | `scripts/sc7_curve_points.py` | Standalone curve-point (re)assignment after a remesh + `curve_<stem>.ipynb` transcript + `output!L2` progress | Yes |
| SC8 | (not shipped — excluded from this public bundle) | Post-calculation extraction is not part of v0.8.5-PUBLIC; review results manually in PLAXIS Output | No — script absent |
| SC9 | `scripts/sc9_IPYNB_LAUNCHER.py` | Executes the generated notebook inside PLAXIS + `assign_sequence!B5` per-cell progress | Yes |

The shared helper `scripts/notebook_paths.py` is used by SC3, SC5, SC6, and SC9 and
defines the single source of truth for the generated notebook location.

> **Retired buttons.** The Example workbooks were copied from the internal chain. The
> buttons for the excluded utilities (`Field Curve Fitting`, `Plaxis Command`) and the
> `Plaxis Output Generator` (Run_sc8) have been removed from all Example workbooks —
> if you see them in an older copy, do not use them; their scripts are not in this
> folder and they will fail. `Plaxis Python` is the SC3 soil-notebook button and is
> fully supported. The supported
> buttons are: SC3 (soil notebook), SC4/Button 1 (DXF import), SC5/Button 2
> (structures), SC6 (staging), SC9/Button 3 (launcher), SC7 Curve Points.

### A.4 Information-status convention

| Status | Meaning |
|---|---|
| *(documented behavior)* | Verified from the current script source and tests |
| *(engineering convention)* | Engineering assumption or sign/unit convention; not enforced by code |
| `[Verification Required]` | Behavior not yet confirmed in a live PLAXIS run; use with care |
| *(warning)* | A warning string the scripts write into Excel or print to console |

### A.5 Change control (public fork)

This public folder is a frozen fork. There is no version-copy workflow here and no
internal history files. If you modify anything, record what you changed and
re-verify the full chain SC3 → SC4 → SC5 → SC6 → SC9 on one example before
touching others.

---

## B. What this automation does

### B.1 The pipeline at a glance

```text
Excel workbook (main sheet: soil profile + materials, filled directly by you)
        │
        ▼
  SC3  soil notebook             ──► creates <workbook>.ipynb with soil, borehole, materials
        │
DXF file ──► SC4  DXF import    ──► writes geometry into `str_2D` blue columns
        │                         writes curve points into `output` tab B:F
        │                         live progress in `str_2D!B4`
        │
(fill yellow user inputs in str_2D)
        │
        ▼
  SC5  structural notebook       ──► appends structural/loads cells to .ipynb,
        │                           writes orange values + warnings, builds `assign_sequence`
        │                           emits merge-equivalents + mesh + output forces (optional)
        │                           live progress in `str_2D!D4`
        │
(define phases + markers in assign_sequence)
        │
        ▼
  SC6  staged construction       ──► appends phase/water code to .ipynb, warnings to B4,
        │                           live progress in `assign_sequence!B5`
        ▼
  SC9  launcher                  ──► executes all notebook code cells inside PLAXIS,
                                    per-cell progress in `assign_sequence!B5`
        │
        ▼
  PLAXIS 2D V2025 model ready for meshing and calculation
        │
(manual: mesh, calculate, save .p2dx)
        │
        ▼
  SC7  curve-point reassign      ──► re-plants output!B9:F108 rows as fresh Output curve
                                    points after a (re)mesh; additive; `curve_<stem>.ipynb`
                                    transcript; output!L2 progress
        │
(manual: SC7 curve points, calculate, save .p2dx)

```

The complete chain is therefore: **Excel workbook → SC3 soil/material input →
SC4 DXF extraction → `str_2D` → SC5 structural/material notebook + `assign_sequence`
→ SC6 staged-construction notebook section → SC9 launcher → PLAXIS 2D**.

### B.2 Division of responsibility

| Actor | Owns |
|---|---|
| User | All yellow cells, DXF drawing conventions, phase design, soil profile and every soil parameter, engineering judgment |
| SC3 | Creation of the workbook-scoped `.ipynb`, soil/material/geometry cells |
| SC4 | All DXF-derived geometry in `str_2D` (blue columns), polygon numbering, `str_2D!B4` progress |
| SC5 | Orange calculated columns, element names, warnings, all notebook sections after SC3's, the `assign_sequence` tab, `str_2D!D4` progress |
| SC6 | Staged-construction notebook cells, phase handling, water conditions, `assign_sequence!B4` warnings, `assign_sequence!B5` progress |
| SC7 | Curve-point (re)assignment only — re-plants `output!B9:F108` rows as fresh Output curve points; never touches structures, staging, or extraction |
| SC9 | Nothing structural — it only executes; its only Excel write is `assign_sequence!B5` per-cell progress |

The user never writes Python. Everything the user maintains is in the workbook;
everything the scripts produce is reproducible from the workbook and the DXF file.

### B.3 Generated files

| File | Created by | Location |
|---|---|---|
| `<workbook name>.ipynb` | SC3 (created), SC5 and SC6 (appended) | Same folder, same base name as the workbook |
| `assign_sequence` tab | SC5 (written), SC6 (read) | Inside the workbook |

**Critical rule** *(documented behavior)*: the notebook is always named after the
workbook (`Path(wb.fullname).with_suffix('.ipynb')`). If you rename the `.xlsm`,
rename any existing `.ipynb` the same way (or simply re-run SC3/SC5/SC6 to
regenerate). Each workbook in a folder has its own notebook; there is no shared
notebook.

### B.4 Execution order and dependencies

| Step | Run order | Depends on |
|---|---|---|
| SC3 | First notebook step | `main` inputs complete (soil parameters filled directly) |
| SC4 | After SC3, whenever DXF changes | A DXF file following the layer conventions (Section T) |
| SC5 | After SC3 + SC4 + filling yellow cells | Notebook exists; `cad_layer` filled by SC4 |
| SC6 | After SC5, whenever staging changes | Notebook exists; `assign_sequence` filled by SC5 |
| SC9 | With PLAXIS Input session running | Complete notebook from SC3 + SC5 + SC6 |
| SC7 | After any (re)mesh, before calculating | Meshed model; `output!B9:F108` curve table present |

Re-run rules:

- SC3 **recreates** the notebook soil section; SC5 and SC6 **clean and replace**
  their own tagged sections, so the normal working loop is: edit inputs → re-run
  the affected button → SC9.
- SC4 must run again after any DXF change; SC5 must run again after any
  yellow-cell or DXF change; SC6 must run again after any phase/marker change.
- The full chain SC3 → SC4 → SC5 → SC6 → SC9 must be re-run after any
  cross-module change.

---

## C. Installation and prerequisites

> **Never installed Python?** Start with `INSTALL.md` in this folder — it walks through installing Python, xlwings, and dependencies from zero. Then return here.

### C.1 Folder layout

```text
PLAXIS_PYTHON_V0.8.5-PUBLIC/
├── MANUAL.md                    ← this manual (standalone for this folder)
├── v0.8.5 EXAMPLE 1.xlsm … 4.xlsm   ← workbooks (one model per workbook)
├── Example N.dxf                 ← DXF geometry files
├── example_N.p2dx                ← reference PLAXIS models
└── scripts/                      ← SC3–SC7 + SC9 (no SC8), notebook_paths.py, tests/
```

Workbooks and their DXF files live in the same folder; the generated notebook
appears beside the workbook automatically.

### C.2 Excel and xlwings requirements

- Enable macros. Each workbook carries VBA buttons that call, e.g.,
  `RunPython "import scripts.sc4_dxf_reader as l; l.main()"`.
- The xlwings add-in must be installed and able to find the project Python
  environment.
- Scripts locate the workbook via `xw.Book.caller()` when run from buttons. When
  run from a terminal, SC3/SC5/SC6 auto-detect the first `.xlsm` in the current
  directory (then the parent directory). SC4 opens a native Excel *File Open*
  dialog titled "Select DXF File" instead.

### C.3 PLAXIS connection

SC3 reads the PLAXIS server settings from `main!V5` (port, e.g. `10000`) and
`main!V6` (API password). The generated notebook's first cells call
`new_server("localhost", port=…, password="…")`. Keep these cells consistent with
your running PLAXIS Input session. The API password is **not** the PLAXIS licence
— it is the remote-scripting password shown in PLAXIS *expert options* when the
API server is started.

### C.4 Before starting a model

- Confirm you are in this public folder.
- Keep one model = one workbook = one DXF.
- Never edit generated orange cells by hand — they are overwritten by SC5 on
  every run.
- Never save the workbook with openpyxl or any non-Excel tool; only Excel/xlwings
  may write the `.xlsm` (openpyxl corrupts the VBA project).

---

## D. First model: quick-start procedure

### D.1 Step sequence

| # | Action | Tool/Cell | Result |
|---|---|---|---|
| 1 | Fill `main`: geometry, soil profile, material types — **all parameters typed directly** (Section F) | `main!B2:D4`, `F2:H8`, `B12:AH33` | Model definition |
| 2 | Run **SC3** | Button | `<workbook>.ipynb` created with soil + materials |
| 3 | Run **SC4** and pick the DXF | Button 1, "Select DXF File" dialog | `str_2D` blue columns filled: structural geometry, soil polygons, drains, loads; watch `str_2D!B4` progress |
| 4 | Fill yellow inputs in `str_2D` per element type | `str_2D` | Types, models, E, A1/A2, capacities, materials (J/K) |
| 5 | Resolve all `MISSING`/warning marks | orange cells | Clean input |
| 6 | Run **SC5** | Button 2 | Orange values filled, rows hidden for blank rows, `assign_sequence` built, structural cells appended to notebook; watch `str_2D!D4` progress |
| 7 | Design phases in `assign_sequence` | rows 1–6 headers + markers from row 7 | Staging plan |
| 8 | Run **SC6** | Button | Staged-construction cells appended; watch `assign_sequence!B5` progress, check `assign_sequence!B4` for ⚠ warnings |
| 9 | Open the workbook's `.ipynb` in PLAXIS and run **SC9** | Button 3 | Model built inside PLAXIS Input; watch `assign_sequence!B5` per-cell progress |
| 10 | Inspect: geometry, materials, water, phases | PLAXIS | Ready to mesh/calculate |
| 11 | Mesh and calculate in PLAXIS | PLAXIS | Saved `.p2dx` model |
| 12 | After a remesh: run **SC7 Curve Points** | Button | Re-plants `output!B9:F108` rows as fresh Output curve points; `output!L2` progress |
| 13 | Run the audit block | Section P | Record results |

### D.2 Minimal viable model (soil only)

For a soil-only model: step 1, skip SC4/SC5 (or run them with an empty DXF), skip
steps 4–7, run SC6 with a single `Initial` phase column, then SC9. The notebook
then contains only contour, borehole, layers, materials, and the initial phase.

### D.3 If something fails

- SC3/SC6 error "notebook not found" → run SC3 first.
- SC9 stops mid-way → read the console traceback; the model state is partially
  built; fix the input, re-run SC9 (it replays every cell) — but see Section O.4
  on partial state and re-runs first.
- A yellow cell was forgotten → SC5 writes `MISSING …` into the relevant orange
  cell; fix and re-run SC5.
- `assign_sequence!B4` shows `⚠ …` → see Section L.7.

---

## E. Workbook architecture

### E.1 Sheets

| Sheet | Owner | Purpose |
|---|---|---|
| `main` | User (inputs) | Geometry, soil profile, PLAXIS connection |
| `str_2D` | SC4 (blue), user (yellow), SC5 (orange) | Central hub: loads, structural elements, geogrids, soil polygons, drains |
| `assign_sequence` | SC5 (element rows), SC6 (read + warnings at B4 + progress at B5), SC9 (progress at B5), user (phase headers + markers) | Element inventory + staged construction plan |

### E.2 Cell color convention (`str_2D`)

| Color | Meaning | Written by | User may edit? |
|---|---|---|---|
| Blue | DXF-derived geometry → owned by SC4, cleared and rewritten on every SC4 run | SC4 | No |
| Yellow | User inputs: types, models, properties, material names | User | Yes — this is your entire input surface |
| Orange | Calculated values and warning marks | SC5 | No — regenerated every SC5 run |
| Green | First row of each `cad_layer` group = **parent** row (SC4) / material-sharing parent (SC5) | SC4/SC5 | No |

Rules that follow from ownership *(documented behavior)*:

- SC4 clears its blue ranges **before writing** (clear-with-`None`, never `.clear()`,
  which would destroy data validation and formatting). Anything you type into a blue
  column is lost on the next SC4 run.
- SC5 reads the green parent row of each `cad_layer` group and lets children
  inherit blank properties; one PLAXIS material object is generated per unique layer
  unless a child differs.
- User yellow columns outside the blue ranges survive SC4 refreshes. Column P is
  SC4-owned only in the FEA section (`Direction_x`); in every other structural
  section P is user-maintained input and survives SC4 refreshes.

### E.3 Row hiding

After SC5 runs, rows whose `cad_layer` (column C) is blank are hidden in every
structural section. SC4 un-hides all data rows of each section before writing. If
rows look missing, re-run SC4 (it unhides) or SC5 (to re-collapse empties).

### E.4 Warning locations

| Location | Written by | Contains |
|---|---|---|
| Orange cells inside `str_2D` sections | SC5 | `MISSING …`, `MISMATCH …` per affected field |
| `assign_sequence!B4` | SC6 | All staging warnings joined by ` \| `, red font, `⚠ ` prefix |

`MISSING` warnings are written **into cells**, never printed only — if a console says
something but no cell shows it, the cell warning system for that condition simply
does not exist (or the condition is a console diagnostic only, e.g. SC4 truncations).

---

## F. `main` sheet — soil and model inputs (fill directly)

There is no parameter-fitting step in this public bundle. Every value SC3 needs
must be typed into `main` by you, and you are responsible for its correctness.

### F.1 Parameter attribution map

| Physical Variable | Excel location (`main`) | Numerical Tool Parameter | Control Mechanism |
|---|---|---|---|
| Ground surface level | `C3` (`surface_level`) | Soil contour top; layer elevation baseline | Sets all layer top/bottom elevations |
| Water table level | `G3` / `F3` (`water_level`) | `Borehole_1.Head`, pore-pressure distribution | Depth where hydrostatic p = 0 |
| Model extents | `F4:F7` (`x_min, y_min, x_max, y_max`) | `initializerectangular SoilContour …` | Overall geometry size |
| Model type | `F8` (`plaxis_type`) | `2D_PS` → Plane Strain, `2D_axis` → Axisymmetric, `3D` → 3D | Constitutive/mesh dimension context |
| Soil type per layer | `B13:B33` (`type`) | `SOIL_TYPE_MAP` → SoilModel + DrainageType | Governs constitutive family |
| Unit weights | `D13:E33` (`gunsat, gsat`) | `gammaUnsat/gammaSat` (kN/m³) | Stress integration |
| Stiffness | `E50_ref`, `Eoed_ref`, `Eur_ref` columns | `E50Ref/EoedRef/EurRef` | Stiffness magnitude |
| Strength | `J13:L33` (`c, phi, psi`) | cohesion, friction, dilatancy | Shear strength |
| Interface factor | `M13` (`Rinter`) | `Rinter` (Manual mode) | Interface strength reduction |
| Stress dependency | `m`, `pref` columns | `PowerM`, `pRef` | E50 pressure dependence |
| HSS extras | `G0_ref`, `g07` columns (+ `RD, PI, OCR` where applicable) | `G0Ref`, `gamma07` | Small-strain stiffness |
| PLAXIS API port / password | `V5` / `V6` | `new_server(...)` in notebook | Connection only |

### F.2 General parameters (`B2:D4`)

| Row | Field | Unit | Notes |
|---|---|---|---|
| 3 | `surface_level` | m+ref | Datum for all elevations |
| 4 | `water_depth` | m | Formula `=MAX(C3-G3, 0)`; do not overwrite the formula |

Rows 5–7 of block `B2:D8` (`extend_bottom`, `Eoed;ref`, `Eur;ref`) are legacy front
end from the internal chain. No script in this public bundle reads them, so they are
left blank in all Example workbooks and carry no meaning here.

### F.3 Model geometry (`F2:H8`)

`water_level` (F3, m+ref), `x_min` (F4), `y_min` (F5), `x_max` (F6), `y_max` (F7) in
m; `plaxis_type` (F8) ∈ {`2D_PS`, `2D_axis`, `3D`}.

### F.4 Soil profile (rows 13–33, max 21 layers)

SC3 reads the full parameter width including `E50_ref`, `Eoed_ref`, `Eur_ref`,
`G0_ref`, `g07`, `clay_frac`, `silt_frac`, `mode`, `k_ver`, `k_hor_ratio`, and SSC
parameters (`einit, OCR, lambda, kappa, mu, POP`) — all typed directly by you.

HS/HSS/SSC overconsolidation *(documented behavior)*: each of `OCR`, `POP`,
`einit` emits independently iff its own cell is filled — blank omits that
pair, filled emits the input value as-is (e.g. `OCR = 1` with blank `POP`
emits only `"OCR" 1`, no `POP` line). A both-filled row emits both with
`POP` trailing. SSC `lambda`/`kappa`/`mu` always emit. A blank SSC `OCR`
previously emitted raw `nan` and crashed the notebook with
`NameError: name 'nan' is not defined` — fixed by omission. SC3 formats all
`soilmat` values through `round_num` (mirrored from SC5, 3-decimal), so
binary residue like `1.5000000000000004` prints as `1.5`.

Layer `type` must be one of *(documented behavior)*:

| `type` | PLAXIS SoilModel | DrainageType |
|---|---|---|
| `clay-HS` / `silt-HS` | Hardening Soil | Undrained (A) |
| `sand-HS` | Hardening Soil | Drained |
| `clay-HSS` / `silt-HSS` | HS small strain | Undrained (A) |
| `sand-HSS` | HS small strain | Drained |
| `clay-SSC` / `silt-SSC` | Soft Soil Creep | Undrained (A) |
| `sand-SSC` | Soft Soil Creep | Drained |
| `linear-elastic` | Linear Elastic | Non-porous |

### F.5 Linear-elastic naming convention

*(documented behavior)* For `linear-elastic` layers, `E` and ν are **parsed from
the layer name**:

```text
Concrete_E20GPa_nu0.2   →   ERef = 20,000,000 kPa ; nu = 0.2
```

- `E<number>GPa` → GPa is converted ×1,000,000 to kPa (PLAXIS unit kN/m²).
- `nu<number>` or `v<number>` → Poisson's ratio.
- If parsing fails the script falls back to the row's `E50_ref` and **`psi`
  column** misinterpretation must be avoided — do **not** place Poisson's ratio in
  the dilatancy column. Only 5 parameters are set: `Identification, SoilModel,
  DrainageType, gammaUnsat, ERef, nu` — no strength, no groundwater.

### F.6 Groundwater / permeability modes

All groundwater modes are **USDA — never Standard**. The class string comes
from the soil `type` prefix verbatim (`clay-*` → `Clay`, `sand-*` → `Sand`,
`silt-*` → `Silt`; PLAXIS `.set()` is case-sensitive).

Column `mode` per layer *(documented behavior)*: `automatic` (default if the
cell is blank) → USDA class defaults only, no fraction lines. `by_grain_size`
→ USDA class, plus each `ClayFraction`/`SiltFraction` line **only when that
AE/AF cell is genuinely filled**; a blank AE/AF cell omits that line so
PLAXIS falls back to the class default (an explicit `0` still emits — PLAXIS
distinguishes `0` from default). `manual` → USDA class, plus the filled
fractions as above, plus permeabilities from `k_ver` (m/day in PLAXIS
convention) and `k_hor_ratio` (horizontal factor;
`PermHorizontalPrimary = k_ver × k_hor_ratio`, default ratio 1.5 if blank).
K0 is always Jaky `K0 = 1 − sin φ` inside the parameter derivation
*(documented behavior)*. A cell reading `by_grain_s` (legacy truncated
dropdown value) is treated as `by_grain_size`.

### F.7 Pre-run validation checklist

- Every layer row: `type` filled, `depth_top < depth_bot`, unit weights positive.
- HSS layers: stiffness extras present, else small-strain response is unreliable.
- `linear-elastic` names contain `E…GPa` and `nu…`.
- `V5`/`V6` match the running PLAXIS session.
- No manual overwrite of orange/calculated columns.
- Every material value replaced with project-specific values (caution below).

> **Caution — example material values are illustrative only.** No number in any
> example workbook is a recommended or real-world value for your site. Every soil
> and structural parameter — stiffness, strength, small-strain extras,
> permeability, unit weights, and the name-parsed linear-elastic `E`/`nu` — must
> be replaced with project-specific, lab- or field-derived values, checked for
> units and for the constitutive model and drainage type you actually need.
> Results computed with example numbers demonstrate that the chain runs; they say
> nothing about your ground. Close this out in the Section P audit.

---

## G. SC3 — soil/material notebook generator

**File:** `scripts/sc3_PythonPlaxis.py` · **Trigger:** xlwings button;
**prerequisite:** `main` complete. SC3 writes no values back to Excel.

### G.1 What it creates

The workbook-scoped notebook `<workbook name>.ipynb` beside the workbook,
containing, in order:

1. **Connection** — `new_server("localhost", port=<V5>, password="<V6>")`
2. **New Model** — `s_i.new()` + model type: `PlaneStrain = 0`, `Axisymmetric = 1`,
   `3D = 2` (from `F8`)
3. **Soil Contour** — `initializerectangular(x_min, y_min, x_max, y_max)`
4. **Borehole** — `borehole(0)` (2D) / `borehole(0, 0)` (3D); `Head` = water level
5. **Soil Layers** — `soillayer(thickness)` per layer
   (thickness = `depth_bot − depth_top`, 3 decimals)
6. **Material Definitions** — `soilmat(...)` per layer with all HS/HSS/SSC/
   linear-elastic parameters and groundwater settings
7. **Material Assignment** — `g_i.Soils[i].Material = <name>`
8. **Layer Elevations** — `setsoillayerlevel(...)` (top of layer 1 =
   surface + depth_top; bottoms = surface − depth_bot)
9. **Done** marker

### G.2 Naming and units

- Material variable names: title-cased, non-alphanumerics stripped (`Layer3Clay`).
- Units: kN/m³, kPa (= kN/m²), m. The linear-elastic `E…GPa` name converts at
  1 GPa = 1,000,000 kPa.

### G.3 Behavior notes

- SC3 **writes nothing to Excel**; on failure nothing is half-written.
- Re-running SC3 **replaces the notebook file** (then SC5/SC6 re-append their
  sections; after re-running SC3 you must re-run SC5 and SC6).
- First cell failures (port/password mismatch) are the most common startup
  error — see Section Q.3.

---

## H. SC4 — DXF import into `str_2D`

**File:** `scripts/sc4_dxf_reader.py` · **Trigger:** workbook button ("Button 1") →
`import scripts.sc4_dxf_reader as l; l.main()`. A native Excel *File Open* dialog
("Select DXF File") appears. Screen updating is disabled during I/O.

Live progress appears in `str_2D!B4`: Tier-1 text bar
`sc4: ⏳ WORKING — {stage} ██░░░░░░░░ 22%` (10-char bar + `%` + spinner).
A stale WORKING banner flips to `sc4: FAILED — see console` only on crash — the
real cause is in the console. `str_2D!B4` must be blank in the template.

### H.1 What SC4 writes

SC4 writes **only DXF-derived geometry** into `str_2D`: `cad_layer` + coordinates
(blue), FEA `Direction_x` (P, FEA section only), soil polygon names/types/numbers
(B,C,D:G,H,I), drains (M:S), and loads (C:G, I, J, K). It never writes material
properties, calculated values, or element names. Console closing line:
`sc4: Done. Fill yellow columns, then press Button 2.`

### H.2 Supported DXF layers and entity types (public subset)

Layer matching is **case-insensitive prefix** matching after bracket stripping and
`_number` suffix removal *(documented behavior)*:

| DXF layer prefix | Element table | Required DXF entity |
|---|---|---|
| `plate` | Plate | LINE / LWPOLYLINE |
| `nn` | N2N anchor | LINE / LWPOLYLINE |
| `anc` | Embedded-beam anchor | LINE / LWPOLYLINE |
| `strut` | Strut | LINE / LWPOLYLINE |
| `fea` or `fe_anchor` | Fixed-end anchor | **POINT** |
| `pile` | Pile | LINE / LWPOLYLINE |
| `pvd` | Drain | LINE (layer must be **exactly** `pvd`) |
| `geo` or `geogrid` | Geogrid | LINE (dedicated pass; LINEs only) |
| `load_line` | Line load | LINE / LWPOLYLINE (per-segment rows) |
| `load_point` | Point load | **POINT** |
| `cut`, `fill`, `soil_replacement`, `volume_profile` | Soil polygons | LWPOLYLINE |
| `waterlevel` | Waterline | LWPOLYLINE or LINE (chained by shared endpoints) |
| `waterboundary_head` | GWFlowBC Head | LINE / LWPOLYLINE |
| `waterboundary_closed` | GWFlowBC Closed | LINE / LWPOLYLINE |
| `curve_node` | Output curve point (node) | **POINT** |
| `curve_stresspoint` | Output curve point (stress) | **POINT** |

Notes *(documented behavior)*:

- A layer named `pvd_1` is **not** a drain (drain matching is exact `'pvd'`);
  conversely `load_line[100kPa]` works because loads use prefix matching.

### H.3 Bracket metadata

Layer names may carry `[bracket]` metadata, parsed tolerantly:

| Example layer | Meaning |
|---|---|
| `anc_1[grout 300mm]` | Anchor group 1, bracket round-trips into `cad_layer` |
| `load_line[100kPa]` | Line load magnitude 100 → `qy_start = −100` (kPa, gravity direction) |
| `load_point[1000kN]` | Point load magnitude 1000 → `qy_start = −1000` (kN) |
| `fe_anchor[H350X350_-12.5]` | Section `H350X350`; text after last `_` = `Direction_x = −12.5` |

The `cad_layer` written to Excel preserves number and bracket:
`anc_1[grout 300mm]` stays intact. FEA rows are re-sorted so all anchors of one
bracket group are together, topmost first; the first row of each group is the
green parent.

### H.4 Coordinates, normalization, closure

- **Universal 3-decimal rounding** (1 mm) on every coordinate written to Excel.
- LINE entities and all structural polylines: axis-snap normalization with a
  **0.01 m tolerance** — near-horizontal/vertical segments (drift ≤ 10 mm) snap to
  a clean shared value, removing AutoCAD drift.
- Soil polygons (cut/fill/volume_profile/soil_replacement LWPOLYLINEs) get full
  polygon normalization: segment snapping, then the **closing vertex is set exactly
  equal to the first vertex**.
- Geogrid and load polylines are **not** axis-snapped (only 3-decimal rounding) to
  protect their geometry.

### H.5 Numbering and section ranges

| Section | Header row | Data rows | SC4 writes | SC4 clears |
|---|---|---:|---|---|
| Loads | 18 | 19–38 (max 20) | C:G, I, J=0, K=−mag | all B:K |
| Plate | 42 | 43–62 (max 20) | C:G | B:G |
| N2N | 90 | 91–110 | C:G | B:G |
| Anchor | 114 | 115–134 | C:G | B:G |
| Strut | 138 | 139–158 | C:G | B:G |
| FEA | 162 | 163–182 | C:G + P (`Direction_x`) | B:G + P (`Direction_x` is refreshed by SC4) |
| Pile | 186 | 187–206 | C:G | B:G |
| Geogrid | 210 | 211–230 (max 20) | C:G (dedicated pass; LINEs only) | B:G (generic pass) then C:G |
| Soil polygons | 234 | 235+ (cap `I233`, ≥50, default 100) | B:I (per polygon **segment** row) | B:I beyond data |
| Drains | 234 | 235+ (cap `R233`, default 100) | M:S | all M:S leftovers |

Soil polygon numbering *(documented behavior)*: `cut_*`, `fill_*`,
`soil_replacement_*` each sorted by Y-centroid **descending** (top first);
`polygon_no` 1..N assigned across all three by that order (matching PLAXIS
creation order); then `volume_profile` polygons appended, continuing the numbering.
One Excel row = one polygon edge; rows of a polygon share `spoly_name`.

Config cells: `I233` = soil polygon maximum limit (clamped ≥ 50), `R233` = drain
maximum limit. More entities than the cap are **truncated with a console warning
only** (no Excel warning) — structural sections beyond 20 rows are silently dropped.

### H.6 Write-safety mechanics *(documented behavior)*

- Clearing is done with `None` block writes; `.clear()` is never used (destroys
  validation/formatting).
- All section data rows are un-hidden row by row before writing (SC5's row collapse
  must not swallow SC4's data).
- Header-driven column maps: writes locate columns by header name, not fixed
  positions — but headers are case-sensitive; a silently no-op write is possible on
  header mismatch.
- First occurrence of each unique `cad_layer` in a section is painted **green**
  (parent row); duplicates are white.

### H.7 SC4 quirks you must know

1. **Column-P ownership is section-specific.** SC4 clears column P only for FEA
   rows, where P stores the DXF-derived `Direction_x`. In other structural sections
   column P is user-maintained and is preserved on SC4 refreshes.
2. Load brackets lacking a closing `]` crash the load parser (the generic parser
   tolerates them; the load parser does not).
3. >20 geogrids/drains/loads: console truncation warning; structural sections:
   silent drop.

### H.8 DXF import verification checklist

- Counts in console (`plate: n elements → C43:G62` style) match your DXF intent.
- Soil polygon `polygon_no` order = top-down.
- No unexpected silently dropped rows beyond the 20-row structural caps.
- Shared polygon edges inspected in PLAXIS after SC9 (caution below).

> **Caution — stacked polygons are a drawing convenience, not shared topology.**
> Each soil body (`cut`, `fill`, `soil_replacement`, `volume_profile`) is drawn
> as a separate closed polyline, so every shared boundary exists twice in the
> DXF. The chain rounds to 1 mm, normalises each polygon individually, and the
> merge-equivalents cell collapses near-duplicates at build time — but none of
> this guarantees a clean single edge in PLAXIS. Duplicate lines, hairline gaps,
> or slivers along shared edges can change stress flow, interface behaviour, flow
> conditions, and mesh quality, and therefore result accuracy. Before trusting any
> result: inspect the built geometry for duplicate lines and slivers along every
> shared edge, confirm the `polygon_no` top-down order matches your intent, check
> the merge console output, and inspect the mesh around the boundaries. If exact
> shared topology matters for your problem, redraw or repair the geometry in
> CAD/PLAXIS rather than accepting the stacked import as-is. Close this out in
> the Section P audit.

### H.9 Curve points → `output` tab

SC4 reads POINT entities on `curve_node` and `curve_stresspoint` layers from the
DXF and writes them into the `output` tab range `B9:F108` *(documented behavior)*.
This range is cleared and rewritten on every SC4 refresh; other `output`-tab
content is preserved.

| Column | Header | Content |
|---|---|---|
| B | `point_name` | Auto-generated: `node_N` or `stresspoint_N` |
| C | `cad_layer` | `curve_node` or `curve_stresspoint` |
| D | `x` | m (3-decimal rounding) |
| E | `y` | m (3-decimal rounding) |
| F | `point_type` | `node` or `stresspoint` |

Points are grouped node-first, then sorted by descending Y within each group. Up to
100 curve points are supported per refresh (cap enforced at row 108).

SC5 reads this table and emits typed curve-point commands in the mesh/curve-point
notebook block:

```python
g_i.curvepoint("node", (x, y), "point_name")
g_i.curvepoint("stresspoint", (x, y), "point_name")
```

These commands run after `g_i.openmesh()` in the mesh section.

---

## I. `str_2D` data dictionary

Header names are case-sensitive everywhere *(documented behavior)*. "Owner":
B = blue/DXF (SC4), Y = yellow/user, O = orange/SC5-calculated.

### I.1 Loads (header 18, rows 19–38, B:K)

| Col | Header | Owner | Unit | Notes |
|---|---|---|---|---|
| B | `Load_name` | O (name) | — | Auto: `line_load1_100kPa`, `point_load1` |
| C | `cad_layer` | B | — | Full DXF layer text |
| D–G | `x1,y1,x2,y2` | B | m | Point loads: F,G blank |
| H | `load_no` | O | — | Sequential |
| I | `load_model` | B | — | `line_load` / `point_load` |
| J | `qx_start` | B | kPa / kN | Written 0 by SC4 |
| K | `qy_start` | B | kPa / kN | **Negative** = gravity direction (−bracket magnitude) |

PLAXIS: line → `g_i.lineload(...) + setproperties("qx_start",…, "qy_start",…)`;
point → `g_i.pointload(...) + setproperties("Fx",…, "Fy",…)`.

### I.2 Plate (header 42, rows 43–62, B:Y)

| Col | Header | Owner | Unit | Effect |
|---|---|---|---|---|
| B | `plate_name` | O | — | `{type}_THK={t}m` |
| C | `cad_layer` | B | — | Inheritance group key |
| D–G | `x1,y1,x2,y2` | B | m | Line geometry |
| H | `plate_type` | Y | — | Descriptive (goes into name) |
| I | `plate_model` | Y | — | `elastic` / `elastoplastic` |
| J | `plate_THK(m)` | Y | m | EI via h³/12 (h=THK if A2 blank) |
| K | `plate_E(Mpa)` | Y | MPa | ×1000 → kN/m² basis |
| L | `plate_A1(m2)` | Y | m²/m | EA1 = E·A1·1000 |
| M | `plate_A2(m2)` | Y | m²/m | EA2 = E·A2·1000 (anisotropic) |
| N | `v_nu` | Y | – | Poisson (default 0.15) |
| O | `plate_w(kN/m/m)` | Y | kN/m/m | weight `w` |
| P | `prevent_punching` | Y | — | yes/no — preserved on SC4 refresh |
| Q | `interface` | Y | — | none/positive/negative/both |
| R | `isotropic` | Y | — | yes/no (TRUE/FALSE/1/0/y/n also accepted) |
| S | `Mp(kNm/m)` | Y | kNm/m | elastoplastic only |
| T/U | `Np1tens(kN/m)`, `Np2tens(kN/m)` | Y | kN/m | elastoplastic only |
| V–Y | `plate_I(m4)`, `plate_EI`, `plate_EA1`, `plate_EA2` | O | m⁴, kNm²/m, kN/m | I=h³/12; EI=E·1000·h³/12; EA=E·1000·A |

### I.3 N2N anchor (header 90) and Strut (header 138) — identical layout, rows +20/each

| Col | Header | Owner | Unit | Effect |
|---|---|---|---|---|
| B | `nn_name` | O | — | `{type}_L={s}m_{E}MPa` |
| H | `nn_type` | Y | — | Descriptive |
| I | `nn_model` | Y | — | `elastic` / `elastoplastic` |
| J | `L_spacing(m)` | Y | m | → `LSpacing` |
| K | `nn_E(Mpa)` | Y | MPa | EA = E·A·1000 |
| L | `nn_A(m2)` | Y | m² | → `EA` |
| M/N | `Fmax_tens`, `Fmax_comp` | Y | kN | elastoplastic only (`FmaxTens`/`FmaxComp`) |
| O | `nn_EA` | O | kN/m | calculated |

PLAXIS: `g_i.anchormat(...)` + `g_i.n2nanchor(x1,y1,x2,y2)`. Labels:
`NodeToNodeAnchor_N` — struts **continue** the N2N counter.

### I.4 Embedded-beam anchor (header 114) and Pile (header 186) — related layouts

Anchor (B:`anc_name`, H:`anc_type`, …):

| Key cols | Header | Owner | Notes |
|---|---|---|---|
| J | `L_spacing(m)` | Y | m |
| K/L | `anc_E(Mpa)`, (E→kN/m²) | Y | PLAXIS computes EA/EI from E + diameter |
| M | `anc_w(kN/m/m)` | Y | Gamma |
| N | `diameter(m)` | Y | m |
| O | `resistance_model` | Y | `linear` / `layer_dependent` → `AxialSkinResistance` |
| P | `Tskin_end` | Y | kN/m — preserved on SC4 refresh |
| R | `Nptens(kN/m)` | Y | elastoplastic only |
| Q | `Tskin_start` | Y | kN/m |

Pile (B:`pile_name`, H:`anc_type` ∈ `solid_circular`/`circular_tube`/`square`, plus
`pile_THK(m)` for tubes — SC4 preserves user columns, including P — and `Fmax`):

PLAXIS: `g_i.embeddedbeammat(...)` (`PredefinedCrossSectionType`:
"Solid circular beam"/"Circular tube"/"Solid square beam") +
`g_i.embeddedbeam(...)`. Labels: `EmbeddedBeam_N` — piles **continue** the anchor
counter.

### I.5 FEA — fixed-end anchor (header 162)

N2N-like columns plus `Direction_x` (P, written by SC4 from the DXF bracket after
the last `_`). Name: `FEA_L={s}m_{E}MPa_EqL{dir}m`. PLAXIS:
`g_i.fixedendanchor(x, y)` **point** element + `a.Direction_x = value` + anchormat
material. Label: `FixedEndAnchor_N`.

### I.6 Geogrid (header 210, rows 211–230)

| Col | Header | Owner | Notes |
|---|---|---|---|
| B | `geogrid_name` | O | `geogrid{N}_{bracket}` |
| H | `gg_type` | Y | Descriptive |
| I | `gg_model` | Y | `elastic` / `elastoplastic` |
| J | `gg_E(Mpa)` | Y | MPa |
| K/L | `gg_A1(m2)`, `gg_A2(m2)` | Y | EA1 = E·A1·1000; EA2 = E·A2·1000 |
| M | `interface` | Y | positive/negative/both/none |
| N | `isotropic` | Y | anisotropic requires A2 = A1 (`MISMATCH` warning otherwise) |
| O/P | `Np1(kN/m)`, `Np2(kN/m)` | Y | elastoplastic only — **note `Np1`, not `Np1tens`**; P preserved on SC4 refresh |
| Q | `gg_EA1` / R: `gg_EA2` | O | calculated |

PLAXIS: `g_i.geogridmat(...)` (`EA1`, `isotropic`, `EA2`, `Np1`, `Np2`) +
`g_i.line((x1,y1),(x2,y2))[-1]` + `g_i.geogrid(line, "Material", mat)` +
pos/neg interfaces. **Geogrids have no EI/Mp.**

**Critical header distinction** *(documented behavior)*: plates use
`Np1tens(kN/m)`/`Np2tens(kN/m)` → PLAXIS `Np1Tens`/`Np2Tens`; geogrids use
`Np1(kN/m)`/`Np2(kN/m)` → PLAXIS `Np1`/`Np2`. Never interchange.

### I.7 Soil polygons (header 234, rows 235+, B:K)

| Col | Header | Owner | Notes |
|---|---|---|---|
| B | `spoly_name` | B | `cut_1`, `fill_1`, `soil_replacement_1`, `str_volume_N` |
| C | `cad_layer` | B | Raw DXF layer |
| D–G | segment x1,y1,x2,y2 | B | One row per polygon edge |
| H | `spoly_type` | B+Y | cut/fill/soil_replacement (SC4) or volume dropdown (user) |
| I | `polygon_no` | B | Creation order (top-down Y-centroid) |
| J | `material` (existing) | Y | Material at creation |
| K | `material` (replacement) | Y | Material on reactivation (blank → warning; `setmaterial` skipped) |

Rules *(documented behavior)*: J = existing soil, K = replacement soil; a blank K
on a reactivatable polygon emits a notebook `WARNING … setmaterial() will be
skipped`. PLAXIS: `g_i.polygon(...)` then
`g_i.Polygons[polygon_no−1].Soil.Material = …`.

### I.8 Drains (header 234, rows 235+, M:S)

M:`drains_name`, N:`cad_layer` (`pvd`), O–R: coordinates, S:`drain_no`. SC4 sorts
left→right by x1. PLAXIS: `g_i.drain(...)`, label `Drain_N`. Cap `R233`.

### I.9 Groundwater-flow boundaries (GWFlowBC)

SC4 imports LINE entities from DXF layers `waterboundary_head` and
`waterboundary_closed` into the shared waterline block at rows 235+. SC5 creates
all `GWFlowBC_N` objects before the ordinary water-level block:

```python
GWFlowBC_1 = g_i.gwfbc((x1, y1), (x2, y2), "Behaviour", "Head")[-1]
GWFlowBC_1.Href = z_w
```

`Href` is prescribed hydraulic head, in metres. A `Closed` boundary represents
zero Darcy flux and receives no `Href`.

---

## J. SC5 — structural/material notebook generator

**File:** `scripts/sc5_structural_to_ipynb.py` · **Trigger:** Button 2.
**Prerequisites:** notebook exists (SC3), `cad_layer` filled (SC4), yellow cells
filled.

Live progress appears in `str_2D!D4`: same Tier-1 bar/spinner pattern as SC4.
Stale WORKING flips to FAILED only on crash. `str_2D!D4` must be blank in the
template.

### J.1 What SC5 does

Reads every `str_2D` section, computes orange values, writes them back (with
warnings), generates all structural/loads/material notebook cells, emits the
merge-equivalents cell, and builds the `assign_sequence` tab. It also hides blank
rows (column C scan) and cleans its own previous notebook sections before
appending (tag-based cleanup preserves user-authored cells).

### J.2 Material inheritance *(documented behavior)*

Rows sharing the same `cad_layer` form a group: first row = **parent** (green),
rest = children. Children with blank properties inherit the parent's; each unique
effective property set becomes one PLAXIS material object; children identical to
the parent reuse it. If a child **differs** from the parent, it silently gets its
own material — check green/white grouping if materials multiply unexpectedly.

### J.3 Calculated formulas (all *documented behavior*)

| Quantity | Formula | Units |
|---|---|---|
| Plate EA1 | `E(MPa) × A1(m²) × 1000` | kN/m |
| Plate EA2 | `E × A2 × 1000` (anisotropic only) | kN/m |
| Plate EI | `E × 1000 × h³ / 12`, `h = A2 if A2 > 0 else THK` | kNm²/m |
| Plate I | `h³ / 12` | m⁴ |
| N2N/strut/FEA EA | `E × A × 1000` | kN/m |
| Geogrid EA1/EA2 | `E × A1(2) × 1000` | kN/m |
| Anchor E | `E(MPa) × 1000` → kN/m² (EA/EI derived by PLAXIS from E + diameter) | — |

### J.4 Plate / N2N / anchor / strut / FEA / pile / geogrid generation

Per type, SC5 generates material object(s) then geometry then material assignment,
with these PLAXIS calls and labels:

| Element | Material call | Geometry call | Label |
|---|---|---|---|
| Plate | `g_i.platemat()` | `g_i.plate(x1,y1,x2,y2)` → `[Pt, Pt, Line]` — interface target is index **[2]** | `Plate_N` |
| N2N / strut | `g_i.anchormat()` | `g_i.n2nanchor(...)` | `NodeToNodeAnchor_N` (shared counter) |
| Anchor | `g_i.embeddedbeammat()` | `g_i.embeddedbeam(...)` | `EmbeddedBeam_N` |
| Pile | `g_i.embeddedbeammat()` with `PredefinedCrossSectionType` | `g_i.embeddedbeam(...)` | `EmbeddedBeam_N` (continues) |
| FEA | `g_i.anchormat()` | `g_i.fixedendanchor(x,y)` + `Direction_x` | `FixedEndAnchor_N` |
| Geogrid | `g_i.geogridmat()` (`EA1, isotropic, EA2, Np1, Np2`) | `g_i.line(...)[-1]` + `g_i.geogrid(line, "Material", mat)` | `Geogrid_N` |
| Loads | — | `g_i.lineload` / `g_i.pointload` | `LineLoad_N` / `PointLoad_N` |

Interfaces (`g_i.posinterface(pl_N[2])` / `g_i.neginterface(...)`) are created
**after all geometry** and use the global `PositiveInterface_N`/
`NegativeInterface_N` counters.

### J.5 Global creation-order rule *(engineering convention — must be preserved)*

`Plate_N`, `PositiveInterface_N`/`NegativeInterface_N`, `NodeToNodeAnchor_N`,
`EmbeddedBeam_N` are **creation-order labels**. The generation order is plates →
N2N → embedded beams → struts → piles → FEA → geogrids → (interfaces after all
geometry) → soil polygons → drains → loads. Any change to element types must
append to `assign_sequence` in exact notebook creation order with a shared global
counter.

### J.6 Interface order and extension *(documented behavior)*

Global counter order: plate interfaces → geogrid interfaces → volume-profile
interfaces. Geogrid endpoint extension into soil is a **known limitation** —
AutoCAD snap does not guarantee PLAXIS topology.

### J.7 SC5 warnings and resolution

| Warning (cell) | Cause | Fix |
|---|---|---|
| `MISSING plate_E` / `MISSING v_nu` / `MISSING THK` | plate row lacks E, ν, or THK | Fill yellow plate columns |
| `MISSING Mp` / `MISSING Np1Tens` / `MISSING Np2Tens` | elastoplastic model without capacities | Fill S/T/U |
| `MISSING A2` | anisotropic without A2 | Fill plate_A2 |
| `MISMATCH isotropic geogrid requires gg_A2 = gg_A1` | isotropic geogrid with A2 ≠ A1 | Make A2 = A1 or set anisotropic |
| `MISSING FmaxTens` / `MISSING FmaxComp` | elastoplastic N2N/strut/FEA | Fill M/N |
| `MISSING Nptens` | elastoplastic anchor/pile | Fill `Nptens(kN/m)` |
| `MISSING pile_THK` | circular_tube pile without THK | Fill pile_THK |
| Notebook print: `WARNING {name}: … has no material_after` | soil polygon K blank | Fill K |

`isotropic` normalization accepts `yes/no, TRUE/FALSE, 1/0, y/n`
*(documented behavior)*.

### J.8 Notebook sections and re-run behavior

Section tags (cleaned and replaced on every SC5 run): structural elements, soil
polygons, soil materials, drains, loads, interface sections, mesh generation,
output forces. Cleanup removes only tag-matched generated cells; user cells
survive. After any SC5 change, re-run SC5 then SC6 (SC6 must re-append after SC5's
cleanup of shared notebook content).

### J.9 Mesh generation and output forces table

SC5 optionally emits a mesh generation block and an output-forces table at the end
of the notebook *(documented behavior)*.

**Mesh toggle:** `str_2D!K3` (`mesh_generation`) controls whether the mesh block
is generated: `included` generates mesh + curve-point commands; `excluded` /
`skipped` skips the block entirely; blank/unrecognized preserves legacy inclusion
behavior (mesh included).

**Mesh factor:** `str_2D!K2` holds the relative mesh factor (numeric). SC5 emits:

```python
g_i.gotomesh()
g_i.mesh(<factor>)
g_i.openmesh()
# ... curve-point commands ...
```

After meshing, SC5 opens the mesh in Output and emits typed curve-point commands
for each point in the `output` tab `B9:F108` range.

**Output forces table**: SC5 writes the `output`-sheet table at K8 onward
(K8 = `element`, L8 = `PLAXIS Label`, M8 onward = phase names from
`assign_sequence` row 6) covering `Plate_*`, `EmbeddedBeam_*`,
`NodeToNodeAnchor_*`, `FixedEndAnchor_*`, and `Geogrid_*` rows. In this public
bundle there is no post-calculation extractor, so SC5 writes and preserves this
table on rewrite but nothing consumes it here. Arguments are typed per
element-row × phase cell (comma-separated); SC5 preserves them by label + phase
on rewrite.

**Local mesh refinement**: type `[refine f]` in any `str_2D` column-C cell of the
target polygon (e.g. `volume_profile[refine 0.2]`), one tagged row per polygon is
enough — SC5 promotes the tag to the whole `polygon_no` group. Factor range
0.05–1.0 (lower = finer); non-numeric, zero, negative, or out-of-range values
abort LOUD (never a silent skip). Gated by the same K3 switch: `excluded` drops
the tags with a console warning instead of emitting. Parser-only: an SC4 reload
wipes column C, so re-add the tags after any reload.

SC5 deliberately does **not** emit a `calculate` command, so the user can inspect
the model before running calculation manually.

### J.10 Merge-equivalents cell *(documented behavior)*

SC5 emits a merge cell that collapses near-duplicate DXF geometry points/lines
(`g_i.mergeequivalents(g_i.Geometry)` with `mereq` fallbacks).

Placement rule — **merge must run in structures mode BEFORE the first
`gotostages()`** *(documented behavior)*: in staged mode the server does not
expose geometry commands, so a merge cell after `gotostages()` crashes. The merge
cell therefore sits before the water-level block: `gotostructures()` → merge →
water levels → mesh. SC6's tag cleanup skips the merge tag so the cell survives
staged-construction reruns.

---

## K. `assign_sequence` and staged construction inputs

### K.1 Tab layout

Element inventory is written by SC5 into columns B–J starting at row 7:

| Col | Header | Content |
|---|---|---|
| B | name | Generated element name |
| C | type | `plate`, `interface`, `soil_polygon`, `soil_replacement`, `str_volume`, `drain`, `line_load`, `point_load`, `fixedendanchor`, `embeddedbeam`, … |
| D | PLAXIS Label | `Plate_N`, `PositiveInterface_N`, … |
| E | cad_layer | DXF group |
| F–I | x1,y1,x2,y2 | geometry |
| J | parent_plate | e.g. interface's parent plate |

### K.2 Phase header rows

Phase columns start at **column L** and extend right (read block L:BI). A column
is a phase **iff row 1 is non-blank** *(documented behavior — row 1 is the phase
boundary)*.

| Row | Field | Emitted to |
|---|---:|---|
| 1 | **Calculation type** (phase boundary; blank = not a phase) | `DeformCalcType` |
| 2 | Time interval | `TimeInterval` (only if non-blank) |
| 3 | Updated mesh/pressures | `yes/true/1/y` → `Deform.UseUpdatedMesh = True` and `Deform.UseUpdatedWaterPressures = True` |
| 4 | Identification text | `.Identification` (repr string, only if non-blank) |
| 5 | *(unused)* | — |
| 6 | Phase name | `.Identify` (only if user typed one; blank → PLAXIS default `Phase_{n}`) |

Allowed calculation types *(documented behavior, normalized; unrecognized →
silently `plastic)*: `k0procedure` (`k0`), `fieldstress` (`field`),
`gravityloading` (`gravity`), `flowonly` (`flow`), `plastic`, `consolidation`
(`consolid`), `safety`, `dynamic`.

**The first phase column must be `Initial`** (row 6 = "Initial"): only that column
branches to `g_i.InitialPhase`; every later phase chains from the previous one via
`g_i.phase(prev)`.

### K.3 Marker syntax (rows 7+, phase columns; comma-separated tokens)

| Token | Meaning | Restrictions |
|---|---|---|
| `A` / `ACTIVATE` | Activate the element in that phase | — |
| `D` / `DEACTIVATE` | Deactivate | — |
| `Dry` | Direct Dry on the marked polygon only (no stack propagation) | Soil rows |
| `Dewatering` | Stack-propagation dewatering (deepest-target rule) | Soil rows only |
| `Interp` | Direct Interpolate on the marked polygon only (no stack propagation) | Soil rows |
| `SS` | Steady-state groundwater flow (`Phase_N.PorePresCalcType`) | Any row; emits once per phase |
| `P200` | Prestress 200 kN | N2N anchors and FEA only; positive digits only; may stand alone or combine `A,P200` |

*(documented behavior)* Tokens are processed independently; unknown tokens are
silently ignored. `P` values are rejected with a `B4` warning if
negative/malformed/non-positive.

Staged material replacement: for `soil_replacement` / `str_volume` rows,
**column B of that row** holds the new material name; activation with a material
name emits a `setmaterial(phase, name)` call *(documented behavior)*.

### K.4 Naming and labels

The marker engine works on **column D labels** (scan stops at the first blank D).
Generated calls use `g_i.{label}` for normal objects *(documented behavior)*.

### K.5 Typical phase patterns

```text
Phase columns:   L: Initial   M: Excavate-1   N: Install-anchors   O: Consolidation
Row 1:           k0procedure  plastic         plastic              consolidation
Row 2:                        —               —                    30
Row 3:                                        yes
Row 6:           Initial      Excavation 1    Anchors L1           Consol 30d
Anchor row:                                   A, P200              (activate + prestress — N2N/FEA only)
Soil polygon:                 A, Dewatering    —                    (activate cut + stack-propagation dewatering)
BoreholeWaterLevel:                           A, SS                (set borehole water + steady-state flow)
UserWaterLevel:                                 A                    (assign user water level to this phase)
```

Design rules of thumb *(engineering convention)*: the Initial phase sets
`k0procedure` for straight geometries and `gravityloading` for non-straight
surfaces; reset-displacement stays on by default; consolidation phases need
`TimeInterval` (row 2) and typically `UseUpdatedWaterPressures`.

> **Caution — staging is simplified; concrete pour is wish-in-place.**
> Structural elements activate at full final stiffness in their activation phase.
> The chain does not model pour lifts, formwork support, curing time, early-age
> stiffness or strength gain, or fresh-concrete pressures. A wall or slab that in
> reality carries load progressively over days appears here as instantly
> structural — which generally understates early deformation and redistributes
> load. If pour sequence or early-age behaviour governs your problem (staged wall
> lifts, slabs loaded before curing, any claim about early deformation or
> cracking), this simplification is unconservative: model the sequence separately
> or accept and record the simplification explicitly in your verification record.
> Close this out in the Section P audit.

---

## L. SC6 — staged construction and water conditions

**File:** `scripts/sc6_staged_construction.py` · **Trigger:** Button.
**Prerequisite:** notebook exists with SC5 sections; `assign_sequence` filled.

Live progress appears in `assign_sequence!B5`: same Tier-1 bar/spinner pattern.
SC9 later overwrites B5 with its own per-cell run (each script runs its own 0–100
fresh). Stale WORKING flips to FAILED only on crash. `assign_sequence!B5` must be
blank in the template.

### L.1 What SC6 appends

One tagged notebook section `# ── Staged Construction (from sc6) ──`:
`g_i.gotostages()`, the `SOURCE_BOUNDS` helper (only if soil water markers exist),
one code cell per phase, and a summary line. Re-running SC6 cleans its own old
section first (idempotent).

### L.2 Phase code emitted

Initial phase: `g_i.setcurrentphase(g_i.InitialPhase)` +
`g_i.InitialPhase.DeformCalcType = "…"` (+ TimeInterval). Later phases:
`ph_N = g_i.phase(prev)`, `setcurrentphase`, `.Identify` (only if named),
`.Identification` (row 4), `.DeformCalcType`, `.TimeInterval`,
`.Deform.UseUpdatedMesh/UseUpdatedWaterPressures` (row 3 = yes), and always
`.MaxStepsStored = 50`.

### L.3 Actions, prestress, materials *(documented behavior)*

- Activation/deactivation: `g_i.activate(g_i.{label}, phase)` /
  `g_i.deactivate(...)`.
- Prestress applies only to N2N anchors and FEA (`FixedEndAnchor`) rows. The
  target label is `{label}_1`. SC6 emits phase-dependent direct property
  assignment:

```python
g_i.NodeToNodeAnchor_1_1.AdjustPrestress[Phase_3] = True
g_i.NodeToNodeAnchor_1_1.PrestressForce[Phase_3] = -100
```

The marker stores positive magnitude (`P100`); PLAXIS receives the negative force.
`P` may be standalone for an already active anchor or combined with `A,P100`.
Plates and embedded beams do not receive prestress commands.

- Material replacement on activation: `g_i.{label}.setmaterial(...)` for
  soil-replacement rows.

### L.4 Additional water levels and the SS marker

SC4 writes DXF waterline segments to `str_2D` U:AA. SC5 reconstructs the ordered
lines and creates `UserWaterLevel_N` in a final `gotoflow()` block after all
structural creation commands, then returns to `gotostages()`.
`BoreholeWaterLevel_1` is always added to `assign_sequence`; PLAXIS creates it
automatically when the chained borehole is created. On a `global_waterlevel` row,
`A` emits `g_i.setglobalwaterlevel(g_i.{label}, phase)`.

`SS` can be placed in any phase cell and is independent of the row action. SC6
emits `phase.PorePresCalcType = "Steady state groundwater flow"` once for that
phase.

### L.5 Soil water conditions *(documented behavior)*

`SOURCE_BOUNDS` (per-polygon bounding boxes from `str_2D`) + union-find on
**X-range overlap** (> 1e-6) form vertical excavation stacks (handles zigzag
profiles). Two distinct water-conditioning mechanisms exist:

**Stack-propagation dewatering** (`Dewatering` marker): Each `Dewatering` target
locates its stack; polygons at/above the target Y-centroid get `"Dry"`, below get
`"Interpolate"`. Rules:

- Multiple `Dewatering` targets in one stack → the **deepest** wins, with a `B4`
  warning.
- A `Dewatering` target found in no stack → `B4` warning.

**Direct water conditions** (`Dry` / `Interp` markers): Applied to the marked
polygon only, no stack propagation.

### L.6 Updated-mesh chaining *(documented behavior)*

If the **first** phase with upd_mesh=yes sets it, a `B4` warning reminds that all
subsequent phases chain updated mesh in PLAXIS.

### L.7 SC6 warnings (`assign_sequence!B4`, red font, `⚠ ` prefix, joined by ` | `)

| Warning | Meaning |
|---|---|
| `upd_mesh=yes on {phase} -> all subsequent phases also have updated mesh (PLAXIS chaining)` | informational |
| `Malformed prestress marker "…"` / `Negative prestress rejected: "…"` | fix the P token |
| `Polygon_{pno} not in any excavation stack (phase {pname})` | geometry/X-overlap issue — check polygon extents |
| `Multiple Dry targets in same stack (phase {pname}): […]. Using deepest Polygon_{deepest}` | usually intended; verify |

---

## M. SC9 — notebook launcher

**File:** `scripts/sc9_IPYNB_LAUNCHER.py` · **Trigger:** Button 3.
**Prerequisite:** complete notebook (SC3 + SC5 + SC6).

### M.1 What it does

Reads `<workbook>.ipynb` (exact path via `notebook_path_for_workbook`; raises
`FileNotFoundError("sc9: No notebook found at … Run sc3/sc5 first.")` if missing),
gathers every `code` cell in order, and `exec`s each in the **shared
xlwings-hosted namespace** — the same namespace where earlier scripts' variables
live. No Jupyter kernel, no subprocess.

### M.2 Execution semantics *(documented behavior)*

- Cells execute strictly in order; variables from earlier cells remain available
  to later cells — SC6's cells depend on this.
- Progress: `sc9: [i/N] {first 80 chars…}` per cell on console plus live Tier-1
  bar in `assign_sequence!B5`: `sc9: ⏳ WORKING — cell i/n ██░░░░░░░░ 22%`;
  PLAXIS echoes and helper prints go to the console.
- **Stop on first error**: an exception prints the traceback and halts — later
  cells are not run. B5 shows `sc9: FAILED — cell {i}/{n}: {first line}` naming
  the failing cell. Banner: `sc9: Done. {success} cells executed, {failed}
  failed.` + B5 `sc9: done — {success}/{n} cells ██████████ 100%`.

### M.3 Why the connection cell matters

The notebook's first cell (`new_server(...)`) must connect to the running PLAXIS
Input session. If you restarted PLAXIS after generating the notebook, re-check
port/password (`main!V5/V6`) before launching.

### M.4 Recovery after a failure

Because cells share state, a mid-way failure leaves a **partially built model**.
The safe recovery sequence *(engineering convention)*: fix the input that caused
the failure, then in PLAXIS undo to a clean state or simply close without saving,
then re-run SC9 on the fixed notebook. If the failure came from stale inputs,
re-run the offending SC script first (SC5/SC6 regenerate their sections), then SC9
again. Never hand-patch a half-built PLAXIS model — always regenerate from the
workbook.

---

## N. SC7 — curve-point (re)assignment after a remesh

> **Post-calculation extraction (SC8) is not part of this public bundle.** The
> internal chain's `scripts/sc8_post_calc_extractor.py` — plate / embedded-beam /
> geogrid / anchor `Plate_*` / `EmbBeam_*` / `Geogrid_*` / `Anchor_*` sheets,
> native charts, NMQ/FU/NU keywords, `Curve_*` sheets, and the `output!N4–N7`
> notes — is deliberately excluded. Review forces, displacements, and curve
> results manually in PLAXIS Output. The `output`-sheet table at K8 onward is
> still written and preserved by SC5 (Section J.9) but has no consumer here.

### N.1 What SC7 does

**File:** `scripts/sc7_curve_points.py` · **Trigger:** Curve Points button.
**Prerequisite:** a meshed model (mesh Output must open); the `output!B9:F108`
table as SC4 wrote it — extra hand rows below are fine.

**What SC7 does** *(documented behavior)* — after a PLAXIS remesh, mesh nodes and
stress points get new internal identities, so previously planted curve points are
stale. SC7 walks the existing curve-point block and re-plants every valid row as
a fresh Output curve point, top to bottom, exactly as the rows sit. SC7 is
standalone — SC4/SC5/SC6/SC9 untouched.

**When to press** — remesh (or mesh from scratch) → SC7 Curve Points → calculate
in PLAXIS. SC7 never calculates; it only plants selection points on mesh geometry, so
a mesh must exist first. No mesh → short `sc7: FAILED` hint in `output!L2` (full
cause in console + MsgBox).

**Additive only** *(documented behavior)* — existing CurvePoints are never cleared.

**Progress + transcript** — progress in `output!L2` with the `sc7:` prefix.
Every press writes `curve_<workbook-stem>.ipynb` beside the workbook (run record
+ replay cells; never fatal). Banner `v0.8.3-sc7` proves fresh code.

### N.2 Current scope and limitations (public)

- SC7 re-plants curve points after a remesh — additive, mesh-first, never
  calculates.
- SC4 writes geogrid polylines as LINE entities only.
- Post-calculation result extraction is not included — review forces,
  displacements, and curves manually in PLAXIS Output.

---

## O. PLAXIS modelling conventions

### O.1 Parameter attribution — pipeline cross-section

| Physical Variable | Numerical Tool Parameter | Control Mechanism / Governing Behavior |
| :--- | :--- | :--- |
| Soil stiffness (E50/Eoed/Eur) | `main` profile → `soilmat` matrix | Settlement shape and magnitude |
| Structural stiffness | plate `E, A1, A2, THK` → `EA1, EA2, EI` (kN/m, kNm²/m) | Structural share of deformation |
| Water table | `Borehole_1.Head`, `BoreholeWaterLevel_1`, `UserWaterLevel_N`, `Conditions[phase]` | Pore-pressure distribution per phase |
| Excavation | cut polygon `A` + `Dewatering` or `Dry` marker + routing | Which polygons are void and their water state |
| Support activation | `A` / `P…` markers, `setmaterial` | When structure carries load; N2N/FEA prestress magnitude |
| Interface strength | `Rinter` (Manual) | Soil–structure interaction stiffness/strength |

### O.2 Units and signs *(engineering convention, enforced where documented)*

| Quantity | Unit | Sign convention |
|---|---|---|
| Coordinates, elevations | m | Y-up; gravity acts in **−Y** |
| Stress, stiffness | kPa (= kN/m²) | Compression positive in PLAXIS convention for pore pressure |
| Plate EA/EI | kN/m, kNm²/m | — |
| Line load | kPa | `qy_start` **negative** = gravity direction |
| Point load | kN | `Fy` negative = downward |
| Prestress | kN | Marker `P200` = +200; N2N/FEA `PrestressForce = −200` |
| Unit weights | kN/m³ | — |
| GPa conversion | 1 GPa = 1,000,000 kPa | linear-elastic name parsing |

### O.3 Naming formats *(documented behavior)*

| Element | Format | Example |
|---|---|---|
| Plate | `{type}_THK={t}m` | `wall_THK=0.3m` |
| N2N / strut | `{type}_L={s}m_{E}MPa` | `anchor_L=1.0m_30MPa` |
| Anchor | `{type}_D={d}m_L={s}m` | `grout_body_D=0.8m_L=1.5m` |
| FEA | `FEA_L={s}m_{E}MPa_EqL{dir}m` | `FEA_L=1.0m_200MPa_EqL-12.5m` |
| Pile | `Pile_{type}_D={d}m_L={s}m` | `Pile_solid_circular_D=0.9m_L=3.5m` |
| Geogrid | `geogrid{N}_{bracket}` | `geogrid1_SG1` |
| Loads | `line_load{N}_{q}kPa`, `point_load{N}` | `line_load1_100kPa` |
| Soil (linear-elastic) | name-parsed `E…GPa`, `nu…` | `Concrete_E20GPa_nu0.2` |
| Soil materials | title-cased, strip non-alphanumerics | `Layer3Clay` |

### O.4 Counters and labels *(documented behavior)*

- `Plate_N`: creation-order labels.
- `NodeToNodeAnchor_N`: shared by N2N and struts. `EmbeddedBeam_N`: shared by
  anchors and piles. `FixedEndAnchor_N`, `Geogrid_N`, `Drain_N`, `LineLoad_N`,
  `PointLoad_N`: per-type counters.
- `PositiveInterface_N`/`NegativeInterface_N`: one global pair of counters across
  plates, geogrids, and volume profiles.
- `Polygon_N`: soil polygons in cut/fill/replacement top-down order, then
  `str_volume` appended.
- The first phase column must be named `Initial`.

---

## P. End-to-end verification and audit

### P.1 The audit block (mandatory for deliverables)

| # | Check | Method in this pipeline |
|---|---|---|
| 1 | **Dimensional check** | All inputs unit-tagged (Section O.2). Verify E conversions (GPa→kPa ×10⁶; MPa→kPa ×10³), load signs, prestress sign |
| 2 | **Static/kinematic equilibrium** | Post-run PLAXIS check: reaction sums vs applied loads; initial phase K0 equilibrium (`k0procedure` for straight geometry, `gravityloading` otherwise) |
| 3 | **Mesh/discretization sensitivity** | PLAXIS mesh refinement study at least once per geometry family; watch geogrid/interface endpoints |
| 4 | **Closed-form analytical cross-check** | Compare against hand calculations (e.g. earth-pressure theory, settlement correlations); record in the verification record |
| 5 | **Modelling-limitation close-out** | (a) Stacked-polygon shared edges inspected — no duplicate lines/slivers (Section H caution); (b) every material value replaced with project-specific values (Section F caution); (c) pour/staging simplification accepted and recorded, or modelled separately (Sections K–L caution). None of the three may remain open at delivery. |

### P.2 Audit chain — each stage

| Stage | What to verify | How |
|---|---|---|
| Excel input | Types/models valid, no `MISSING` orange warnings, yellow complete | SC5 run + visual |
| DXF → Excel | Counts, coordinates, polygon closure | SC4 console + H.8 checklist |
| Excel → notebook | Orange values = intended formulas; `assign_sequence` rows present and ordered | Re-derive one EA/EI by hand (J.3) |
| Notebook → PLAXIS | Object counts, labels, positions | PLAXIS object browser after SC9 |
| Phases | Calc types, active sets, water conditions per phase | PLAXIS phase tree + `B4` warnings reviewed |
| Water | Dry above / Interp below, deepest-target rule | Section L.5 rules |
| Reproducibility | Workbook + DXF + scripts + PLAXIS version + notebook archived together | Folder discipline (C.1) |

### P.3 Reproducibility record (minimum)

Record in every delivery: workbook file + version, DXF file, script folder version,
PLAXIS 2D version (e.g. 2025), generated notebook, test-suite status
(`scripts/tests`), verification date, and the audit-block results (including the
modelling-limitation close-out, P.1 row 5).

---

## Q. Troubleshooting

### Q.1 Workbook not detected

SC6/SC9 (or terminal runs) glob `*.xlsm` in cwd then parent. If nothing found:
`FileNotFoundError("No .xlsm file found in …")`. Run from the button, or set cwd
to this folder.

### Q.2 Notebook not found

`sc9: No notebook found at … Run sc3/sc5 first.` / `sc6: Notebook '…' not found …
Run SC3 first.` → Run SC3, then SC5, then SC6. Remember the notebook name follows
the workbook name exactly.

### Q.3 PLAXIS connection failure

First notebook cell fails (`new_server`). Check: PLAXIS Input open, API/RPS
server started, `main!V5` port and `main!V6` password match the session, no
firewall blocking localhost. SC9 will stop at this cell with the PLAXIS traceback
in console.

### Q.4 Empty `str_2D` sections after SC4

DXF layer names don't match Section H.2 conventions (case-insensitive prefix rule
still requires the right prefix, e.g. `pvd_1` is not a drain). Check the console
per-section counts line.

### Q.5 Stale values after DXF refresh

User-typed content in blue columns is intentionally wiped. Column P is preserved
in every structural section except FEA (where SC4 rewrites `Direction_x`).
Re-fill any other yellow inputs, then SC5.

### Q.6 Header mismatch / silent no-op

Column maps are header-name based and case-sensitive; a renamed or re-spelled
header silently skips mapping. Verify exact header spelling/case before editing
`str_2D` layout.

### Q.7 Missing orange inputs / `MISSING` warnings

Table in 9.7. Commonest: elastoplastic models without capacities; anisotropic
without A2.

### Q.8 `assign_sequence!B4` warnings

Table in 11.7. All are fix-in-Excel issues; none require script edits.

### Q.9 Geogrid mesh error (`Deformation not compatible in STRESBL [Error code: 10]`)

Trigger: geogrid interfaces with an endpoint terminating **at or too close to** a
soil-boundary intersection (≈ within PLAXIS' 0.01 m tolerance). Meshing succeeds
when the geogrid ends before the intersection. **Do not** "fix" by inflating
geogrid stiffness. Proper remedy: end the geogrid before the intersection.

### Q.10 Wrong water-condition assignment

Check: X-range stack membership (11.5), deepest-target rule, `polygon_no`
untouched by hand. `Dry` above / `Interpolate` below the target centroid is by
design.

### Q.11 Invalid marker errors

P syntax table in K.3 with `B4` warnings in L.7. Remember: negative prestress
rejected.

### Q.12 SC9 partial run

Stop-on-first-error leaves a partial model (M.4). Fix input → regenerate →
re-run SC9 cleanly.

### Q.13 Stale `__pycache__`

Symptom: your `.py` edit "doesn't take effect". Delete all `__pycache__` in this
folder's `scripts/` tree. If the edit still doesn't take effect after clearing
`__pycache__`, restart Excel — the Python interpreter may hold a stale in-memory
copy of the module from a prior `RunPython` invocation.

### Q.14 SC7 curve-point failures

`sc7: FAILED — mesh not open (meshed? SC5 K3), see console` → the model has no
generated mesh: mesh in PLAXIS (SC5 mesh block, `str_2D!K2/K3`, Section J.9),
save, press Curve Points again. `sc7: FAILED — connect failed, see console` →
PLAXIS/Input unreachable or model-stem mismatch (`output!L4` stem); the console +
MsgBox carry the full cause.

---

## R. Recommended modelling workflow patterns (public subset)

### R.1 Basic soil-only model

`main` inputs (typed directly) → SC3 → SC6 (single Initial phase, `k0procedure`)
→ SC9. Verify layer elevations and K0 stresses.

### R.2 Excavation with struts/anchors + dewatering

DXF: walls (`plate`), struts/anchors (`strut`/`anc`, brackets for capacities),
excavation outline (`cut`), PVDs (`pvd`). Yellow: plate models, anchor spacings,
J/K materials. Phases: Initial (k0) → Exc-1 (cut `A,Dry`; walls already placed in
Initial) → Strut L1 (`A, P200`) → Exc-2 … Consol (`TimeInterval`). Watch stack
warnings for zigzag digs.

### R.3 Road/embankment fill (material replacement)

`fill` polygons with J = existing, K = new material; activate fills in phases;
`setmaterial` handles the replacement per K.3/L.3.

### R.4 Consolidation

Phase col with `consolidation`, `TimeInterval` (row 2), typically
`UseUpdatedWaterPressures` (row 3 = yes). Drains (`pvd` layer) must exist in DXF;
drains activate/deactivate like any element.

### R.5 Anchored wall with staged prestress

N2N/FEA rows with `A,P200` in install phase → prestress applied. Plates and
embedded beams do not receive prestress commands.

### R.6 Soil improvement / replacement

`soil_replacement` polygons + material in K; activation with `setmaterial` swaps
the soil.

---

## S. Complete Excel range map (public subset)

### S.1 `main`

| Block | Range | Content |
|---|---|---|
| General parameters | `B2:D4` | surface_level, water_depth (formula); rows 5–7 legacy, unread, left blank |
| Model geometry | `F2:H8` | water_level, x_min, y_min, x_max, y_max, plaxis_type |
| Soil profile | rows 13–33 | full parameter matrix, typed directly (Section F) |
| PLAXIS connection | `V5`, `V6` | port, password |

### S.2 `str_2D`

| Section | Header | Data | SC4 writes |
|---|---|---:|---|
| Loads | 18 | 19–38 | C:G, I, J, K |
| Plate | 42 | 43–62 | C:G |
| N2N | 90 | 91–110 | C:G |
| Anchor | 114 | 115–134 | C:G |
| Strut | 138 | 139–158 | C:G |
| FEA | 162 | 163–182 | C:G + P |
| Pile | 186 | 187–206 | C:G |
| Geogrid | 208 | 211–230 | C:G (LINEs only) |
| Soil polygon | 234 | 235+ (cap I233) | B:I |
| Drains | — | 234 | 235+ (cap R233) | M:S |

Config: `I233` soil cap, `R233` drain cap.

### S.3 `assign_sequence`

| Block | Rows | Columns |
|---|---|---|
| Element inventory | header 6, data 7+ | B name, C type, D PLAXIS label, E cad_layer, F–I coords, J parent |
| Phase block | rows 1–6 | columns L–BI |
| Markers | rows 7+ | same phase columns |
| Warnings | `B4` | SC6 |
| Progress | `B5` (shared SC6→SC9 handoff) | SC6, SC9 |

---

## T. DXF layer naming standard (public subset)

```text
<type>[_<number>][<bracket>]

type     : plate | nn | anc | strut | fea | fe_anchor | pile | pvd | geo | geogrid
           | load_line | load_point | cut | fill | volume_profile | soil_replacement
           | waterlevel | waterboundary_head | waterboundary_closed
           | curve_node | curve_stresspoint
number   : integer group number (optional for some types)
bracket  : [metadata] or (metadata) — mismatched brackets tolerated by the generic parser
```

| Purpose | Example |
|---|---|
| Plate group | `plate_1[D-Wall 800mm]` |
| Anchor group with metadata | `anc_1[grout 300mm]` |
| FEA with section+direction | `fe_anchor[H350X350_-12.5]` |
| Line load with magnitude | `load_line[100kPa]` |
| Point load | `load_point[750kN]` |
| Drain | `pvd` (exact; `pvd_1` is ignored) |
| Geogrid | `geogrid_1[SG1]` |
| Soil polygon | `cut`, `fill`, `volume_profile`, `soil_replacement` |
| Water level (polyline) | `waterlevel_N` (any suffix; matched by `waterlevel` prefix) |
| GWFlowBC head boundary | `waterboundary_head` (exact match; not prefix) |
| GWFlowBC closed boundary | `waterboundary_closed` (exact match; not prefix) |
| Output curve node | `curve_node` (POINT entities) |
| Output curve stress point | `curve_stresspoint` (POINT entities) |

Entity types required: LINE/LWPOLYLINE (most), **POINT** (`fea`, `load_point`,
`curve_node`, `curve_stresspoint`). Drain layer must be exactly `pvd`.
Water-boundary layers use exact match, not prefix matching.

---

## U. PLAXIS API property mapping (public subset)

| Excel header | Formula/translation | PLAXIS object property |
|---|---|---|
| `plate_E(Mpa)`, `plate_A1(m2)` | ×1000 → kN/m | `EA1` |
| `plate_E(Mpa)`, `plate_A2(m2)` | ×1000 (anisotropic) | `EA2` (with `isotropic=False`) |
| `plate_E(Mpa)`, `THK/A2` | E·1000·h³/12 | `EI` |
| `v_nu` | direct | `StructNu` |
| `plate_w(kN/m/m)` | direct | `w` |
| `prevent_punching` | yes→True | `PreventPunching` |
| `Mp(kNm/m)` | direct | `Mp` |
| `Np1tens(kN/m)` / `Np2tens(kN/m)` | direct | `Np1Tens` / `Np2Tens` |
| geogrid `Np1(kN/m)` / `Np2(kN/m)` | direct | `Np1` / `Np2` |
| `isotropic` | yes/no normalized | `isotropic` (False + EA2 for aniso) |
| plate model | elastic→1, elastoplastic→2 | `MaterialType` |
| `nn_E(Mpa)`, `nn_A(m2)` | EA = E·A·1000 | `EA` (anchormat) |
| `Fmax_tens` / `Fmax_comp` | direct | `FmaxTens` / `FmaxComp` |
| anchor/pile `anc_E(Mpa)` | ×1000 | `E` (kN/m²) |
| `diameter(m)`, `pile_THK(m)` | direct | `Diameter` / `Thickness` |
| `resistance_model` | linear→0, layer_dependent→2 | `AxialSkinResistance` |
| `L_spacing(m)` | direct | `LSpacing` |
| `Tskin_start` / `Tskin_end` | direct | `TSkinStartMax` / `TSkinEndMax` |
| `Direction_x` | direct | `Direction_x` (FixedEndAnchor) |
| row-1 calc type | normalized (K.2) | `DeformCalcType` |
| row-2 time | float | `TimeInterval` |
| row-3 yes | → True | `Deform.UseUpdatedMesh`, `Deform.UseUpdatedWaterPressures` |
| prestress `P200` | AdjustPrestress True + `PrestressForce = −200` on `{label}_1` | `AdjustPrestress`, `PrestressForce` |

---

## V. Auto-naming examples (quick reference, public subset)

```text
wall_THK=0.3m                          → Plate_1
slab_THK=0.25m                         → Plate_2
anchor_L=1.0m_30MPa                    → NodeToNodeAnchor_1
strut_beam_L=3.0m_200MPa               → NodeToNodeAnchor_5   (continues N2N counter)
grout_body_D=0.8m_L=1.5m               → EmbeddedBeam_1
Pile_solid_circular_D=0.9m_L=3.5m      → EmbeddedBeam_7       (continues anchor counter)
FEA_L=3.0m_210000MPa_EqL-12.5m         → FixedEndAnchor_1
geogrid1_SG1                           → Geogrid_1
line_load1_100kPa                      → LineLoad_1
point_load1                            → PointLoad_1
cut_1, fill_1, str_volume_2            → Polygon_no per column I
Layer3Clay                             → soil material variable
```

---

## W. Marker syntax quick reference (public subset)

| Cell content | Effect |
|---|---|
| `A` | Activate element in this phase |
| `D` | Deactivate |
| `A,Dry` | Activate + set Dry water (soil rows) |
| `A,Interp` | Activate + Interp |
| `P200` | Prestress 200 kN on already-active element (N2N/FEA only) |
| `A,P200` | Activate + prestress (N2N/FEA only) |

Multiple tokens comma-separated, case-insensitive; unknown tokens ignored;
malformed P → B4 warning.

---

## X. Warning message reference (public subset)

| Where | Text (pattern) | Producer | Meaning |
|---|---|---|---|
| str_2D orange | `MISSING plate_E`, `MISSING v_nu`, `MISSING THK`, `MISSING Mp`, `MISSING Np1Tens`, `MISSING Np2Tens`, `MISSING A2`, `MISSING FmaxTens`, `MISSING FmaxComp`, `MISSING Nptens`, `MISSING pile_THK` | SC5 | Yellow input missing (J.7) |
| str_2D orange | `MISMATCH isotropic geogrid requires gg_A2 = gg_A1` | SC5 | Geogrid isotropic/A2 conflict |
| assign B4 | `⚠ Malformed prestress marker "…"` / `⚠ Negative prestress rejected: "…"` | SC6 | P token syntax |
| assign B4 | `⚠ Polygon_{N} not in any excavation stack (phase …)` | SC6 | Stack membership failure |
| assign B4 | `⚠ Multiple Dry targets in same stack (…): Using deepest Polygon_{N}` | SC6 | Info; verify intent |
| assign B4 | `⚠ upd_mesh=yes on {phase} -> all subsequent phases also have updated mesh (PLAXIS chaining)` | SC6 | Info |
| Notebook print | `WARNING {name} … setmaterial() will be skipped` | SC5 | Soil polygon K blank |
| Console only | geogrid/drain/load truncation warnings | SC4 | Cap exceeded (I233/R233/20) |
| PLAXIS mesh | `Deformation not compatible in STRESBL [Error code: 10]` | PLAXIS | Geogrid endpoint/interface topology (Q.9) |

---

## Y. Known limitations

| # | Item | Status |
|---|---|---|
| G1 | Numerical validity of any model | Never established by pure-Python tests alone; requires full chain + PLAXIS results review |
| G2 | Geogrid endpoint extension into soil | Not implemented; end geogrids before intersections |
| G3 | Anything not documented in this manual (including code paths the scripts may contain beyond this text) | Unsupported in this public release — educational use only, verify everything |

---

*End of public manual. Verify every model against the PLAXIS GUI. The authors bear
no responsibility for bugs, wrong results, or any damages from using these scripts.*
