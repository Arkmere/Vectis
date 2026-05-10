# Vectis

Vectis is a local desktop utility for deterministic triage of daily NM movement exports. It accepts a CSV or XLSX file, preserves the original rows and columns, runs ordered extraction passes, and writes a triaged Excel workbook for human review.

Vectis v0.1 is intentionally small and auditable. It does **not** include statistical anomaly detection, cross-day tracking, database storage, a web UI, or complex scoring.

## Requirements

- Python 3.12+
- tkinter (normally included with Python desktop installs)
- pandas
- openpyxl

Install Python dependencies:

```bash
pip install -r requirements.txt
```

## Run the GUI

From the repository folder:

```bash
python vectis_gui.py
```

The GUI lets you choose:

- an input `.csv` or `.xlsx` NM export
- an output folder
- a VKB/reference folder containing `config/` and `vkb/`

Click **Process file** to create the triaged workbook.

## Expected input columns

Vectis v0.1 expects the project owner's canonical NM/NOP export schema exactly. The input CSV/XLSX must contain all of these headers:

- `TOT/TA`
- `LS`
- `STA`
- `ARCID`
- `ATYP`
- `RM`
- `ADEP`
- `ADES`
- `ALT1`
- `ALT2`
- `D`
- `T`
- `ARF`
- `IOBT`
- `LV`
- `U`
- `E/CTOT`
- `X`
- `F`
- `S`
- `CL`
- `A/TTOT`
- `AT`
- `TOBT`
- `TSAT`
- `TT`
- `Delay`
- `R`
- `RRP RespBy`
- `Opp`
- `YY`
- `Turn`
- `W`
- `MSG`
- `REGUL+`
- `O`
- `Column1`
- `Impacted`
- `CCAMS`
- `Helper Column`
- `Helper Number`

Field usage in v0.1:

- `ARCID` — callsign / aircraft identification
- `ATYP` — ICAO aircraft type designator
- `RM` — canonical registration / registration-mark field
- `ADEP` — departure aerodrome
- `ADES` — destination aerodrome
- `ALT1` and `ALT2` — preserved for future location checks

Timing, flow, regulation, helper, and other NM/NOP fields are preserved unchanged but not interpreted in v0.1. Vectis does not accept a `REG` alias in v0.1; use `RM` from the canonical export.

All original columns are preserved unchanged in the output workbook. Derived audit columns are appended after the original columns.

## Output

Vectis writes this file to the selected output folder:

```text
<input_stem>_triaged.xlsx
```

The workbook contains extraction tabs, a remainder tab, VKB update candidates, and a summary tab.

## Reference CSV files

Vectis v0.1 uses small `config/` files for explicit interest lists and the real VKB CSV exports in `vkb/` for known-entity lookups:

- `config/interest_icao_codes.csv` — known ICAO/state/military callsign codes of interest for Pass 2
- `config/interest_aircraft_types.csv` — aircraft types of operational interest for Pass 4
- `config/sensitive_registrations.csv` — sensitive registrations
- `config/sensitive_regions.csv` — sensitive ICAO location prefixes
- `config/longform_roots.csv` — additional known longform callsign roots
- `vkb/FDMS_CALLSIGNS_STANDARD.csv` — known standard operators; `TRICODE` is the operator key and the callsign/name/country fields are available for enrichment
- `vkb/FDMS_CALLSIGNS_NONSTANDARD CALLSIGNS.csv` — known non-standard callsign roots/names used alongside `config/longform_roots.csv`
- `vkb/FDMS_LOCATIONS_B_E_L.csv` — known aerodromes; `ICAO CODE` is the location key and `USER` values `MILITARY`, `STATE`, and `DUAL` drive the military/state/dual-use pass while `TYPE` remains descriptive only
- `vkb/FDMS_AIRCRAFT_TYPES.csv` — aircraft type reference keyed by `ICAO Type Designator`; this is loaded for enrichment only and does not make an aircraft operationally interesting by itself

If any reference file is missing, Vectis creates an empty placeholder with the expected headers and continues with a warning. Local checkouts that have not yet received the FDMS VKB exports can still use the legacy sample `vkb/operators.csv` and `vkb/locations.csv` files as a fallback for the bundled sample data.

## Sample data

A small demonstration file is included at:

```text
input/sample_nm.csv
```

Run it through the engine with:

```bash
python - <<'PY'
from triage_engine import triage_file
result = triage_file('input/sample_nm.csv', 'output', '.')
print(result.output_path)
PY
```

## Current limitations

- v0.1 is a deterministic, pass-based triage tool only.
- Rows are extracted by the first matching pass but may accumulate reasons from later checks.
- Reference CSV quality directly affects output quality.
- No database, web server, background processing, anomaly detection, or cross-day tracking is included.
