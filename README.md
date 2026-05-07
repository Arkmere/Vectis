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

The input file must contain these columns:

- `ARCID` — callsign
- `REG` — registration
- `ATYP` — aircraft type
- `ADEP` — departure aerodrome
- `ADES` — destination aerodrome

Additional input columns are preserved in the output workbook.

## Output

Vectis writes this file to the selected output folder:

```text
<input_stem>_triaged.xlsx
```

The workbook contains extraction tabs, a remainder tab, VKB update candidates, and a summary tab.

## Reference CSV files

Vectis loads reference data from:

- `config/interest_icao_codes.csv` — known ICAO/state/military callsign codes of interest
- `config/interest_aircraft_types.csv` — aircraft types of interest
- `config/sensitive_registrations.csv` — sensitive registrations
- `config/sensitive_regions.csv` — sensitive ICAO location prefixes
- `config/longform_roots.csv` — known longform callsign roots
- `vkb/operators.csv` — known operator tricodes
- `vkb/locations.csv` — known locations and military/state/dual-use flags

If any reference file is missing, Vectis creates an empty placeholder with the expected headers and continues with a warning.

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
