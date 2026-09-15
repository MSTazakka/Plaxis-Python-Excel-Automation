# INSTALL — Setting up Python + Excel + PLAXIS (from zero)

> For first-time users. No Python experience assumed. Follow A → G in order;
> each section ends with a check so you know it worked before moving on.
> When everything here is green, open `MANUAL.md` Section D and build your
> first model.

---

## A. What you need before starting

| Item | Why | Where to get it |
|---|---|---|
| Windows 10/11 PC | PLAXIS 2D and the Excel link only run on Windows | — |
| PLAXIS 2D (V2025 or newer) with a licence | The automation drives PLAXIS, it does not replace it | Your Bentley/Seequent licence administrator |
| Desktop Microsoft Excel (not Excel Online) | The workbooks use VBA macro buttons | Microsoft 365 or Office 2021+ |
| Administrator rights on your PC (for steps B–D) | Installing Python and the Excel add-in needs it | Your IT department if the PC is managed |
| This folder (`PLAXIS_PYTHON_V0.8.5-PUBLIC`) on a local drive | The workbooks, scripts, and manuals live here | Copy it somewhere simple, e.g. `C:\PlaxisPublic\` |

Check: you can open one of the `v0.8.5 EXAMPLE N.xlsm` files in desktop Excel.
(If Excel warns about macros, click **Enable Content** — the buttons need macros.)

---

## B. Install Python

1. Go to [python.org/downloads](https://www.python.org/downloads/) and download
   the latest **Python 3.12 or 3.13** Windows installer (64-bit).
2. Run the installer. On the very first screen, tick the box
   **"Add python.exe to PATH"** — this is the step everyone forgets, and
   everything after it fails without it. Then click **Install Now**.
3. When it finishes, open a terminal: press `Win + R`, type `cmd`, press Enter.
4. Type the following and press Enter after each line:

```text
python --version
pip --version
```

Check: both commands print a version number (e.g. `Python 3.13.1`,
`pip 25.0.1`). If Windows says "not recognized", re-run the installer and make
sure the PATH box is ticked.

---

## C. Install the Python packages

The scripts need four packages: `xlwings` (Excel link), `pandas` + `numpy`
(data tables), and `ezdxf` (reading DXF files). (`plxscripting`, the PLAXIS
module, ships with PLAXIS itself — see Section E.)

1. In the same `cmd` terminal, type:

```text
pip install xlwings pandas numpy ezdxf
```

2. Wait for the downloads to finish (a few minutes on first install).

Check: type the following and confirm each prints a version with no errors:

```text
python -c "import xlwings, pandas, numpy, ezdxf; print('all packages OK')"
```

Check: it prints `all packages OK`. If a package is missing, re-run the
`pip install` line for that name.

---

## D. Connect Excel to Python (xlwings add-in)

1. In the terminal, type:

```text
xlwings addin install
```

2. Open desktop Excel. Go to **File → Options → Add-ins**. At the bottom, set
   **Manage: Excel Add-ins**, click **Go…**, and confirm **xlwings** is ticked.
   You should now see an **xlwings** tab in the Excel ribbon.
3. The workbook buttons run Python code, so Excel must allow that. Go to
   **File → Options → Trust Center → Trust Center Settings → Macro Settings**:
   - Select **Enable all macros** (or "Disable all macros with notification"
     and then always click Enable Content — either works).
   - Tick **Trust access to the VBA project object model**. Without this, the
     buttons fail.
4. Open `v0.8.5 EXAMPLE 1.xlsm` from this folder. If a yellow security bar
   appears, click **Enable Content**.

Check: press `Alt + F11` in Excel — the VBA editor opens and you can see the
workbook's modules. Close it again without changing anything. Then, on any
workbook button press later, Python should launch instead of an error about
`RunPython` or a missing interpreter.

> If your PC already had Python from another program, the buttons may find the
> wrong one. Fix: in Excel, go to the **xlwings** ribbon tab → **Interpreter**
> and point it at the Python you installed in Section B (usually
> `C:\Users\<you>\AppData\Local\Programs\Python\Python313\pythonw.exe`).

---

## E. PLAXIS scripting setup

1. Open PLAXIS 2D (Input). Start the remote scripting server: in PLAXIS go to
   **Expert → Configure remote scripting server** (exact menu wording depends
   on version), set a port (default `10000`) and a password, and start it.
2. In the workbook, open the `main` sheet and find cells `V5` (port) and `V6`
   (password). Type the same port and password there.
3. Keep PLAXIS Input open while you press the SC3/SC4/SC5/SC6/SC9 buttons —
   the generated notebook connects to that running session.

Check: the port in `main!V5` matches the PLAXIS server port, and the password
in `main!V6` matches. Mismatches are the single most common first-run failure
(see `MANUAL.md` Section Q.3).

> The API password is **not** your PLAXIS licence — it is the remote-scripting
> password shown in PLAXIS when the API server is started.

---

## F. First-run smoke test

1. PLAXIS Input open with the scripting server running (Section E).
2. Excel open with `v0.8.5 EXAMPLE 1.xlsm`, macros enabled (Section D).
3. Press the **SC3** button. A file named `v0.8.5 EXAMPLE 1.ipynb` should appear
   next to the workbook within a few seconds.
4. Press **Button 1 (SC4)** and select the Example 1 DXF when asked. The
   `str_2D` sheet fills with geometry and `str_2D!B4` shows a progress bar.

Check: both steps complete with no error pop-ups. If so, your installation
works — continue with `MANUAL.md` Section D (quick-start procedure).

---

## G. Troubleshooting install problems

| Symptom | Cause | Fix |
|---|---|---|
| `python` is not recognized | PATH box unticked during install | Re-run installer, tick "Add python.exe to PATH" |
| `pip install` fails with network errors | Firewall/proxy blocks PyPI | Retry on another network, or ask IT to allow `pypi.org` |
| `import xlwings` fails | Packages installed to a different Python | `pip --version` shows which Python pip serves; use `python -m pip install …` to force the right one |
| Button press: error about `RunPython` | xlwings add-in not active | Repeat Section D steps 1–2 |
| Button press: "Trust access" error | VBA object-model access off | Section D step 3, tick the box |
| Buttons open the wrong Python / old packages | Multiple Pythons on PC | Set Interpreter path in xlwings ribbon tab (Section D note) |
| SC3 notebook fails at first cell | Port/password mismatch | Section E; `MANUAL.md` Q.3 |
| `plxscripting` import fails | Running outside PLAXIS's Python context or PLAXIS closed | SC7 and the notebook must run with PLAXIS open; `plxscripting` ships with PLAXIS, never `pip install` it |
| Edits to `.py` files "don't take effect" | Stale `__pycache__` / Excel holds old module | Delete `scripts/__pycache__` folders; if still stale, restart Excel (`MANUAL.md` Q.13) |

Still stuck? Record the exact error text (screenshot or copy from the console),
plus your `python --version` and `pip list` output — that is everything needed
to diagnose it.
