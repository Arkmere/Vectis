STATE.md — Vectis (NM Triage Tool)
1. Project Identity

Project Name:
Vectis

Purpose:
A local, desktop-based aviation OSINT triage tool designed to process large NM (Network Manager) movement datasets and extract government, military, contractor, and non-standard aviation activity.

Naming Note:
“Vectis” (Latin name for the Isle of Wight) aligns with the broader Vectair ecosystem as an internal intelligence-processing tool.

Core Principle:

Remove the normal until only the intelligence-relevant remains.

2. Problem Statement

Daily NM exports contain:

50,000–60,000 rows

The majority are:

generic GAT (scheduled civil airline traffic)

Current workflow:

filter by tricode
remove high-frequency operators (RYR, EZY, etc.)
manually scan remaining rows
search by reg / aircraft type / route
identify “unusual” movements

Limitations:

time-intensive
non-repeatable logic
reliance on memory and intuition
no structured feedback loop
difficult to scale
3. Solution Overview

Vectis is a local Python application with a minimal GUI that:

1. Accepts NM export (CSV/XLSX)
2. Runs structured, ordered “passes”
3. Extracts rows matching binary criteria
4. Outputs a structured Excel workbook
5. Identifies VKB gaps alongside operational hits
6. Preserves auditability of all decisions
4. Core Design Philosophy
4.1 Binary-first detection

Vectis prioritises:

explicit, deterministic indicators

NOT:

statistical anomaly detection (initially)
4.2 Pass-based processing

Vectis operates via ordered extraction passes.

Each pass:

evaluates all non-extracted rows
copies matches to output
marks rows as extracted

Important:

input dataset is never modified (non-destructive)
4.3 VKB-centric architecture

Vectair Knowledge Base (VKB) is used as:

reference dataset
classification layer
data gap detection system
4.4 Human-in-the-loop

Vectis:

does not make final judgements

It:

presents filtered, structured candidates for human review
5. Input Data
Supported formats
.csv
.xlsx
Strict v0.1 input contract
Vectis v0.1 expects the project owner's canonical NM/NOP export schema exactly. Required headers are:
TOT/TA
LS
STA
ARCID
ATYP
RM
ADEP
ADES
ALT1
ALT2
D
T
ARF
IOBT
LV
U
E/CTOT
X
F
S
CL
A/TTOT
AT
TOBT
TSAT
TT
Delay
R
RRP RespBy
Opp
YY
Turn
W
MSG
REGUL+
O
Column1
Impacted
CCAMS
Helper Column
Helper Number

Field usage notes
ARCID   (callsign / aircraft identification)
ATYP    (ICAO aircraft type designator)
RM      (registration / registration-mark field)
ADEP    (departure aerodrome)
ADES    (destination aerodrome)
ALT1    (preserved for now; future location checks)
ALT2    (preserved for now; future location checks)

Timing, flow, regulation, helper, and other NM/NOP fields are preserved but not interpreted in v0.1.
REG is not part of the v0.1 input contract and is not accepted as an alias for RM.

All original fields are preserved unchanged. Derived/audit columns are appended after the original columns.

6. Output Data
Primary output
<filename>_triaged.xlsx
Workbook structure
01_EXTRACTED_ALL
02_LONGFORM_CALLSIGNS
03_KNOWN_ICAO_INTEREST
04_KNOWN_SENSITIVE_REGS
05_INTEREST_AIRCRAFT_TYPES
06_UNKNOWN_OPERATOR_TRICODES
07_VKB_LOCATION_GAPS
08_VKB_MILITARY_LOCATIONS
09_ZZZZ_AERODROMES
10_SENSITIVE_REGIONS
11_REMAINDER_UNEXTRACTED
12_VKB_UPDATE_CANDIDATES
13_SUMMARY
Additional columns added
CALLSIGN_FORM
CALLSIGN_ROOT
CALLSIGN_PREFIX3

EXTRACTED_PASS
FIRST_MATCH_REASON
ALL_MATCH_REASONS

MATCHED_FIELD
MATCHED_VALUE

EXTRACTED_FLAG

VKB_UPDATE_REQUIRED
OPERATIONAL_INTEREST
7. Core Processing Model
7.1 Callsign classification

ARCID must be classified into:

LONGFORM_NONSTANDARD
OFFICIAL_ICAO
STANDARD_CIVIL
REGISTRATION_CALLSIGN
UNKNOWN_SHORT_CODE

Critical invariant:

EAGLE01 ≠ EAG tricode
7.2 Pass System (Version 0.1)
Pass 1 — Longform callsigns
ARCID starts with ≥4 letters
AND not a registration
Pass 2 — Known ICAO/state callsigns

Uses VKB operator list.

Pass 3 — Known sensitive registrations

Uses VKB registration list.

Pass 4 — Aircraft of interest

Includes:

military aircraft
Soviet transport
ISR platforms
Pass 5 — Unknown tricodes
ARCID prefix not in VKB
AND not previously extracted
Pass 6 — VKB location checks
ADEP/ADES not in VKB
OR marked MILITARY / STATE
Pass 7 — ZZZZ aerodromes
ADEP or ADES == ZZZZ
Pass 8 — Sensitive regions
Location in restricted region list
8. VKB Integration
Current usage
Operator tricodes
Location codes
Extended VKB model (target)
Operators
TRICODE → operator → country → category
Locations
ICAO → country → type → sensitivity
Aircraft types
ATYP → category
Registrations
RM → operator → category → notes
Callsign roots
ROOT → category → notes
Dynamic tracking (future)
ENTITY → first_seen → last_seen → count
9. GUI (Version 0.1)
Technology
Python Tkinter
Features
Select input file
Select VKB folder
Select output folder
Process button
Open output folder
Status display
Constraint
GUI must remain minimal
All logic resides in backend
10. File Structure
vectis/
  vectis_gui.py
  triage_engine.py

  config/
    interest_icao_codes.csv
    interest_aircraft_types.csv
    sensitive_registrations.csv
    longform_roots.csv
    sensitive_regions.csv

  vkb/
    operators.csv
    locations.csv

  input/
  output/
11. Execution Model

Run via:

python vectis_gui.py

Future:

PyInstaller packaged executable
Desktop shortcut
12. Key Invariants
Input data is never modified
Each row appears once in EXTRACTED_ALL
Rows may accumulate multiple reasons
Pass order must remain fixed
Longform callsigns must not be reduced to tricodes
VKB mismatches must be surfaced, not ignored
13. Development Phases
v0.1 — Binary extraction tool
pass system
minimal GUI
Excel output
v0.2 — VKB integration
operator classification
location classification
VKB gap detection
v0.3 — usability refinement
config-driven passes
improved parsing
enhanced summaries
v1.0 — OSINT platform
cross-day tracking
pattern detection
VKB growth automation
expanded interface
14. End Goal

Vectis should:

reduce ~60,000 rows → a few hundred high-value movements

while simultaneously:

improving VKB coverage and accuracy
15. Current Status
Concept: complete
Pass system: defined
VKB integration: defined
GUI design: defined
Implementation: not started

16. Implementation Update — 2026-05-12 — Analyst Workbook v2 / GUI Refresh

Ticket: VECTIS-AW2-GUI-001
Branch: feature/analyst-workbook-v2-gui-refresh

Implemented:
- Analyst Workbook v2 workbook ordering with README, analyst-priority, multi-signal, operational-interest, VKB-hygiene, ranked candidate, diagnostics, audit, pass, and summary sheets.
- Deduplication metadata: exact-row hash, exact duplicate group sizing, strict/loose movement keys, movement group sizing, first-row flags, representative movement row flag, representative selection reason, and group variation summary.
- Raw-vs-unique accounting in the summary sheet, including extracted/remainder, VKB hygiene, analyst priority, and multi-signal metrics.
- Deterministic analyst scoring and priority bands, including hygiene-only separation, multi-signal flags, primary reason, short explanation, and suggested next action.
- Ranked unknown operator and unknown location enrichment queues with raw counts, unique strict/loose movement counts where applicable, samples, first/last seen IOBT, operational-interest context, max analyst score, and update priority.
- VKB load audit sheet showing loaded resources, expected paths, row/key counts, duplicate/blank/non-standard keys, and standard callsign sanity checks for BAW, EZY, RYR, UPS, KAL, RAM, SVA, THA, TCA, and MSR.
- VKB match diagnostics sheet for parsed callsign roots, deterministic match status, unmatched reasons, raw/unique counts, and non-operative near-match hints.
- Pass-specific sheets now include sheet-level match reason/field/value plus preserved first-match and all-match audit trail.
- Workbook formatting refresh: frozen headers, autofilters, bold header rows, restrained widths, wrapped explanation/audit columns, and analyst-band styling.
- GUI refresh with script-relative asset paths, assets/vectis.png branding, cleaner internal-tool layout, path sections, action section, readable log, clear button, and bottom state line.
- Logo fallback: if assets/vectis.png is missing or fails to load, the GUI launches with text branding and logs a non-fatal warning.

Validation evidence:
- `python validate_sample.py` passed canonical schema validation.
- `triage_file('input/sample_nm.csv', 'output', '.')` generated `output/sample_nm_triaged.xlsx` successfully with 12 total rows, 11 extracted rows, and 1 remainder row.
- Generated workbook check confirmed all v2 sheets from `00_README` through `20_SUMMARY` are present.
- Generated workbook check confirmed audit columns `EXACT_ROW_HASH`, `MOVEMENT_KEY_STRICT`, `MOVEMENT_KEY_LOOSE`, `ANALYST_SCORE`, `ANALYST_BAND`, and `VKB_HYGIENE_ONLY` exist.
- Generated workbook check confirmed `01_ANALYST_PRIORITY` excludes hygiene-only rows for the sample.

Notes:
- `input/May test Incomplete.csv` was not present in this checkout, so the May-file manual test could not be run.
- Generated output workbooks remain uncommitted and should not be added to git.

17. Implementation QA Update — 2026-05-14 — VECTIS-GUI-001 Reconciliation

Ticket: VECTIS-GUI-001-QA
Branch: main / worktree QA branch
Reference merge: a7ea2c7 Merge pull request #6 from Arkmere/codex/add-progress-bar-and-heartbeat-indicator

Result:
- VECTIS-GUI-001 has already been implemented and merged on main in commit a7ea2c7.
- This QA pass reconciled the existing implementation rather than adding a new feature.
- No ARCID/RM classification, NSCD candidate matching, scoring refinement, or workbook semantic changes were implemented in this ticket.

Verified VECTIS-GUI-001 features:
- Threaded GUI processing is present: `process_file()` starts a background `threading.Thread` named `VectisTriageWorker` and does not run `triage_file()` on the Tkinter main thread.
- Queue-based progress updates are present: the worker posts progress, completion, and error events to `queue.Queue`; Tkinter consumes those events via `.after()` polling.
- Tkinter widget mutation is confined to the Tkinter main-thread path: progress, status, log, messagebox, and control-state updates occur from GUI callbacks, queue polling, heartbeat, completion handling, or error handling.
- Heartbeat is present: `_heartbeat()` updates elapsed active-processing status on a recurring `.after()` timer while processing is active.
- Progress bar and current-stage text are present and driven by the progress queue.
- Vectair styling is present through the Vectair colour constants, themed frames, header, accent button, progress style, and branded layout.
- GUI controls are disabled during active processing and restored on completion or error; duplicate process requests are ignored while processing is active.
- Missing-logo fallback is present: failed logo loading falls back to text branding and logs a non-fatal warning.
- GUI errors are reported to both the log and `messagebox.showerror()`.
- Output-folder opening remains present through `open_output_folder()` and backend `open_folder()`.
- `triage_file(input_path, output_dir, reference_dir)` remains supported for existing non-GUI usage.
- `triage_file(input_path, output_dir, reference_dir, progress_callback=callback)` is supported for GUI progress usage.
- Local input/output `.gitignore` protection is present for `input/*.csv`, `input/*.xlsx`, `output/*.xlsx`, and `output/*.xlsm`.
- `VECTIS_IMPLEMENTATION_PLAN.md` is present at the repository root and records Phase 1 GUI requirements.

Verified progress stages emitted by `triage_file()`:
- 5 — Reading input file
- 10 — Validating canonical schema
- 18 — Loading config and VKB references
- 28 — Adding duplicate / movement identity metadata
- 36 — Classifying callsigns
- 52 — Running deterministic passes
- 65 — Applying analyst scoring and representative movement logic
- 76 — Building ranked VKB candidate sheets
- 86 — Building diagnostics and summary sheets
- 95 — Writing Excel workbook
- 100 — Complete

Validation evidence:
- `python -m py_compile triage_engine.py vectis_gui.py` passed.
- `python validate_sample.py` passed canonical schema validation.
- Callback validation using `triage_file('input/sample_nm.csv', 'output', '.', progress_callback=callback)` generated `output/sample_nm_triaged.xlsx` successfully and emitted the expected ordered progress stages.
- Workbook generation with the Phase 5-adjacent statistics/reporting additions succeeded on the sample input.
- Generated workbook check confirmed the Phase 5-adjacent sheets `07_VKB_IMPACT_PRIORITY`, `21_KNOWN_INTEREST_GAPS`, `22_SIGNAL_DISTRIBUTION`, and `20_SUMMARY` are present.

Environment / manual-test notes:
- `python vectis_gui.py` could not be interactively validated in this headless container because Tkinter failed with `TclError: no display name and no $DISPLAY environment variable`.
- The May test input file was not present in this checkout; only `input/sample_nm.csv` was available locally.
- Because of the headless environment, manual confirmations that require visual interaction with the GUI window remain to be performed on a workstation with a display: launch, visual logo/fallback confirmation, file selection, live responsiveness, live heartbeat/progress observation, completion dialog, duplicate-run click test, and Open output folder click test.

Scope note — early Phase 5 work already landed with the GUI merge:
- The same merge also introduced early Phase 5 workbook-statistics/reporting features.
- These additions are already partially implemented and should be reviewed later, not blindly reimplemented.
- Observed Phase 5-adjacent outputs include:
  - `07_VKB_IMPACT_PRIORITY`
  - `21_KNOWN_INTEREST_GAPS`
  - `22_SIGNAL_DISTRIBUTION`
  - expanded `20_SUMMARY` blocks
  - VKB coverage metrics

18. Implementation Update — 2026-05-14 — VECTIS-ARCID-001 Contextual ARCID/RM Classification

Ticket: VECTIS-ARCID-001
Branch: current feature branch

Implemented:
- Added contextual ARCID/RM/airframe identifier classification.
- Added strict tricode rule requiring three letters followed by a digit.
- Added RM status classification.
- Added ARCID/RM equality handling.
- Added Spanish medical ME + EC registration-derived callsign handling.
- Added config seeds for civil registration patterns, military serial patterns, word-like registrations, and aircraft type context.
- Added ARCID classification audit fields.
- Added ARCID classification diagnostics workbook sheet.
- Suppressed false unknown-operator candidates from pure-letter ARCID values and registration-derived identifiers.

Validation:
- `python -m py_compile triage_engine.py vectis_gui.py` passed.
- `python validate_sample.py` passed.
- `python tests/validate_arcid_context.py` passed focused classifier acceptance examples and a synthetic workbook unknown-operator suppression check.
- Workbook generation against `input/sample_nm.csv` passed.
- Generated workbook check confirmed `ARCID_CLASSIFICATION_DIAGNOSTICS` exists.
- Unknown-operator queue check confirmed no MEE/HBJ/FEV/MOO false positives where sample data permits.

Notes:
- This ticket does not implement full NSCD contextual candidate scoring.
- This ticket does not complete Phase 4 scoring refinement.
- `input/May test Incomplete.csv` was not present in this checkout, so before/after May-file unknown-operator counts could not be reported.
