"""Deterministic pass-based triage engine for Vectis v0.1."""

from __future__ import annotations

import os
import platform
import re
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl.utils import get_column_letter

REQUIRED_COLUMNS = [
    "TOT/TA",
    "LS",
    "STA",
    "ARCID",
    "ATYP",
    "RM",
    "ADEP",
    "ADES",
    "ALT1",
    "ALT2",
    "D",
    "T",
    "ARF",
    "IOBT",
    "LV",
    "U",
    "E/CTOT",
    "X",
    "F",
    "S",
    "CL",
    "A/TTOT",
    "AT",
    "TOBT",
    "TSAT",
    "TT",
    "Delay",
    "R",
    "RRP RespBy",
    "Opp",
    "YY",
    "Turn",
    "W",
    "MSG",
    "REGUL+",
    "O",
    "Column1",
    "Impacted",
    "CCAMS",
    "Helper Column",
    "Helper Number",
]

CANONICAL_NM_HEADERS = REQUIRED_COLUMNS

ADDED_COLUMNS = [
    "CALLSIGN_FORM",
    "CALLSIGN_ROOT",
    "CALLSIGN_PREFIX3",
    "EXTRACTED_PASS",
    "FIRST_MATCH_REASON",
    "ALL_MATCH_REASONS",
    "MATCHED_FIELD",
    "MATCHED_VALUE",
    "EXTRACTED_FLAG",
    "VKB_UPDATE_REQUIRED",
    "OPERATIONAL_INTEREST",
]

CONFIG_REFERENCE_SPECS = {
    "config/interest_icao_codes.csv": ["CODE", "NAME", "CATEGORY", "NOTES"],
    "config/interest_aircraft_types.csv": ["ATYP", "CATEGORY", "NOTES"],
    "config/sensitive_registrations.csv": ["REG", "OPERATOR", "CATEGORY", "NOTES"],
    "config/sensitive_regions.csv": [
        "COUNTRY",
        "ICAO_PREFIX",
        "REGION_NAME",
        "CATEGORY",
        "NOTES",
    ],
    "config/longform_roots.csv": ["ROOT", "CATEGORY", "NOTES"],
}

FDMS_REFERENCE_SPECS = {
    "vkb/FDMS_CALLSIGNS_STANDARD.csv": [
        "TRICODE",
        "CALLSIGN",
        "COMMON NAME",
        "COMPANY/CORPORATE NAME",
        "COUNTRY",
    ],
    "vkb/FDMS_CALLSIGNS_NONSTANDARD CALLSIGNS.csv": [
        "CALLSIGN",
        "COMMON NAME",
        "COMPANY/CORPORATE NAME",
        "COUNTRY",
    ],
    "vkb/FDMS_LOCATIONS_B_E_L.csv": ["ICAO CODE", "USER", "TYPE"],
    "vkb/FDMS_AIRCRAFT_TYPES.csv": ["ICAO Type Designator"],
}

LEGACY_VKB_REFERENCE_SPECS = {
    "vkb/operators.csv": ["CODE", "NAME", "CATEGORY", "COUNTRY", "NOTES"],
    "vkb/locations.csv": ["ICAO", "NAME", "COUNTRY", "TYPE", "SENSITIVITY", "NOTES"],
}

REFERENCE_SPECS = CONFIG_REFERENCE_SPECS | FDMS_REFERENCE_SPECS

PASS_NAMES = {
    1: "LONGFORM_CALLSIGNS",
    2: "KNOWN_ICAO_INTEREST",
    3: "KNOWN_SENSITIVE_REGS",
    4: "INTEREST_AIRCRAFT_TYPES",
    5: "UNKNOWN_OPERATOR_TRICODES",
    6: "VKB_LOCATION_GAPS",
    7: "VKB_MILITARY_LOCATIONS",
    8: "ZZZZ_AERODROMES",
    9: "SENSITIVE_REGIONS",
}


@dataclass(frozen=True)
class TriageResult:
    output_path: Path
    warnings: list[str]
    total_rows: int
    extracted_rows: int
    remainder_rows: int


def _clean(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_code(value: object) -> str:
    return re.sub(r"[\s-]+", "", _clean(value).upper())


def is_registration_callsign(value: object) -> bool:
    """Conservative registration-style callsign detector for v0.1."""
    code = normalize_code(value)
    if not code:
        return False
    patterns = [
        r"^N[1-9][0-9]{0,4}[A-Z]{0,2}$",  # N650RX, N556PM, N60125
        r"^[GDF][A-Z]{4}$",  # GABCD / DABCD / FXXXX, optionally hyphenated before normalization
        r"^OE[A-Z]{3}$",  # OEABC / OE-ABC
    ]
    return any(re.match(pattern, code) for pattern in patterns)


def classify_callsign(arcid: str) -> dict[str, str]:
    """Classify an ARCID into v0.1 callsign form, root, and three-letter prefix."""
    code = normalize_code(arcid)
    if not code:
        return {"CALLSIGN_FORM": "UNKNOWN", "CALLSIGN_ROOT": "", "CALLSIGN_PREFIX3": ""}

    if is_registration_callsign(code):
        return {
            "CALLSIGN_FORM": "REGISTRATION_CALLSIGN",
            "CALLSIGN_ROOT": code,
            "CALLSIGN_PREFIX3": code[:3],
        }

    leading_letters = re.match(r"^([A-Z]+)", code)
    root = leading_letters.group(1) if leading_letters else ""

    if len(root) >= 4 and re.match(r"^[A-Z]{4,}\d", code):
        return {
            "CALLSIGN_FORM": "LONGFORM_NONSTANDARD",
            "CALLSIGN_ROOT": root,
            "CALLSIGN_PREFIX3": root[:3],
        }

    if re.match(r"^[A-Z]{3}[A-Z0-9]+$", code):
        return {
            "CALLSIGN_FORM": "TRICODE_STYLE",
            "CALLSIGN_ROOT": code[:3],
            "CALLSIGN_PREFIX3": code[:3],
        }

    return {
        "CALLSIGN_FORM": "UNKNOWN",
        "CALLSIGN_ROOT": root,
        "CALLSIGN_PREFIX3": root[:3] if root else "",
    }


def ensure_reference_files(base_dir: Path) -> list[str]:
    warnings: list[str] = []
    for relative_path, headers in REFERENCE_SPECS.items():
        path = base_dir / relative_path
        if not path.exists():
            path.parent.mkdir(parents=True, exist_ok=True)
            pd.DataFrame(columns=headers).to_csv(path, index=False)
            warnings.append(f"Created missing reference placeholder: {path}")
    return warnings


def _read_reference(
    base_dir: Path, relative_path: str, headers: list[str]
) -> pd.DataFrame:
    path = base_dir / relative_path
    try:
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except pd.errors.EmptyDataError:
        df = pd.DataFrame(columns=headers)
    for header in headers:
        if header not in df.columns:
            df[header] = ""
    return df


def _load_references(base_dir: Path) -> dict[str, pd.DataFrame]:
    return {
        relative_path: _read_reference(base_dir, relative_path, headers)
        for relative_path, headers in REFERENCE_SPECS.items()
    }


def _read_optional_reference(
    base_dir: Path, relative_path: str, headers: list[str]
) -> pd.DataFrame:
    path = base_dir / relative_path
    if not path.exists():
        return pd.DataFrame(columns=headers)
    return _read_reference(base_dir, relative_path, headers)


def _load_standard_operator_codes(
    refs: dict[str, pd.DataFrame], base_dir: Path
) -> set[str]:
    """Load known standard operator tricodes from the real FDMS callsign VKB."""
    fdms_standard = refs["vkb/FDMS_CALLSIGNS_STANDARD.csv"]
    codes = set(fdms_standard["TRICODE"].map(normalize_code)) - {""}
    if codes:
        return codes

    # Backward-compatible fallback keeps the bundled sample usable if the real
    # FDMS export has not yet been placed under vkb/ in a local checkout.
    legacy = _read_optional_reference(
        base_dir, "vkb/operators.csv", LEGACY_VKB_REFERENCE_SPECS["vkb/operators.csv"]
    )
    return set(legacy["CODE"].map(normalize_code)) - {""}


def _load_known_longform_roots(refs: dict[str, pd.DataFrame]) -> set[str]:
    config_roots = set(
        refs["config/longform_roots.csv"]["ROOT"].map(normalize_code)
    ) - {""}
    fdms_nonstandard = refs["vkb/FDMS_CALLSIGNS_NONSTANDARD CALLSIGNS.csv"]
    fdms_roots: set[str] = set()
    preferred_columns = {"ROOT", "CALLSIGN", "CALLSIGN ROOT", "LONGFORM ROOT"}
    candidate_columns = [
        column
        for column in fdms_nonstandard.columns
        if normalize_code(column)
        in {normalize_code(name) for name in preferred_columns}
        or "CALLSIGN" in normalize_code(column)
        or "ROOT" in normalize_code(column)
    ]
    for column in candidate_columns:
        fdms_roots.update(set(fdms_nonstandard[column].map(normalize_code)) - {""})
    return config_roots | fdms_roots


def _load_location_reference(
    refs: dict[str, pd.DataFrame], base_dir: Path
) -> tuple[set[str], set[str]]:
    """Return known location ICAOs and operationally-interesting USER codes.

    The real FDMS locations file uses ICAO CODE as the key. USER drives
    military/state/dual-use logic; TYPE is descriptive and intentionally not
    used for extraction logic.
    """
    fdms_locations = refs["vkb/FDMS_LOCATIONS_B_E_L.csv"].copy()
    fdms_locations["ICAO_NORM"] = fdms_locations["ICAO CODE"].map(normalize_code)
    known_locations = set(fdms_locations["ICAO_NORM"]) - {""}
    interesting_locations = set(
        fdms_locations.loc[
            fdms_locations["USER"].map(
                lambda value: normalize_code(value) in {"MILITARY", "STATE", "DUAL"}
            ),
            "ICAO_NORM",
        ]
    ) - {""}
    if known_locations:
        return known_locations, interesting_locations

    legacy = _read_optional_reference(
        base_dir, "vkb/locations.csv", LEGACY_VKB_REFERENCE_SPECS["vkb/locations.csv"]
    )
    legacy["ICAO_NORM"] = legacy["ICAO"].map(normalize_code)
    known_locations = set(legacy["ICAO_NORM"]) - {""}
    interesting_locations = set(
        legacy.loc[
            legacy["TYPE"].map(
                lambda value: normalize_code(value)
                in {"MILITARY", "STATE", "DUAL_USE", "DUAL"}
            ),
            "ICAO_NORM",
        ]
    ) - {""}
    return known_locations, interesting_locations


def _load_aircraft_type_reference(refs: dict[str, pd.DataFrame]) -> set[str]:
    """Load FDMS aircraft type keys for enrichment-only lookups.

    Vectis v0.1 deliberately does not use this set to decide Pass 4 interest;
    Pass 4 remains driven only by config/interest_aircraft_types.csv.
    """
    fdms_aircraft = refs["vkb/FDMS_AIRCRAFT_TYPES.csv"]
    return set(fdms_aircraft["ICAO Type Designator"].map(normalize_code)) - {""}


def read_input_file(input_path: str | Path) -> pd.DataFrame:
    path = Path(input_path)
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, dtype=str, keep_default_na=False)
    raise ValueError("Input must be a .csv or .xlsx file.")


def _validate_required_columns(df: pd.DataFrame) -> None:
    missing = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(
            f"Input file is missing canonical NM/NOP headers: {', '.join(missing)}"
        )


def _append_reason(
    reasons: list[list[str]],
    idx: int,
    pass_no: int,
    reason: str,
    first_pass: list[int | None],
    first_reason: list[str],
    matched_field: list[str],
    matched_value: list[str],
    field: str,
    value: str,
) -> None:
    reasons[idx].append(reason)
    if first_pass[idx] is None:
        first_pass[idx] = pass_no
        first_reason[idx] = reason
        matched_field[idx] = field
        matched_value[idx] = value


def _starts_with_any_prefix(code: str, prefixes: Iterable[str]) -> str:
    for prefix in sorted(prefixes, key=len, reverse=True):
        if prefix and code.startswith(prefix):
            return prefix
    return ""


def _reason_contains(value: object, tokens: tuple[str, ...]) -> bool:
    text = _clean(value)
    return any(token in text for token in tokens)


def triage_file(
    input_path: str | Path,
    output_dir: str | Path,
    reference_dir: str | Path | None = None,
) -> TriageResult:
    """Run Vectis v0.1 deterministic triage and write the Excel workbook."""
    input_path = Path(input_path)
    output_dir = Path(output_dir)
    base_dir = Path(reference_dir) if reference_dir else Path.cwd()

    warnings = ensure_reference_files(base_dir)
    refs = _load_references(base_dir)

    df = read_input_file(input_path)
    _validate_required_columns(df)
    original_columns = list(df.columns)
    work = df.copy(deep=True)

    classifications = work["ARCID"].apply(classify_callsign).apply(pd.Series)
    for column in ["CALLSIGN_FORM", "CALLSIGN_ROOT", "CALLSIGN_PREFIX3"]:
        work[column] = classifications[column]

    interest_codes = set(
        refs["config/interest_icao_codes.csv"]["CODE"].map(normalize_code)
    ) - {""}
    interest_types = set(
        refs["config/interest_aircraft_types.csv"]["ATYP"].map(normalize_code)
    ) - {""}
    sensitive_regs = set(
        refs["config/sensitive_registrations.csv"]["REG"].map(normalize_code)
    ) - {""}
    longform_roots = _load_known_longform_roots(refs)
    operator_codes = _load_standard_operator_codes(refs, base_dir)
    location_codes, military_location_codes = _load_location_reference(refs, base_dir)
    _aircraft_type_keys = _load_aircraft_type_reference(refs)

    regions_df = refs["config/sensitive_regions.csv"].copy()
    regions_df["PREFIX_NORM"] = regions_df["ICAO_PREFIX"].map(normalize_code)
    region_prefix_to_name = {
        row["PREFIX_NORM"]: _clean(row["REGION_NAME"])
        or _clean(row["COUNTRY"])
        or row["PREFIX_NORM"]
        for _, row in regions_df.iterrows()
        if row["PREFIX_NORM"]
    }

    row_count = len(work)
    reasons: list[list[str]] = [[] for _ in range(row_count)]
    first_pass: list[int | None] = [None] * row_count
    first_reason: list[str] = [""] * row_count
    matched_field: list[str] = [""] * row_count
    matched_value: list[str] = [""] * row_count
    vkb_update_required = [False] * row_count
    operational_interest = [False] * row_count

    # Pass 1: Longform/non-standard ARCID callsigns.
    for idx, row in work.iterrows():
        root = normalize_code(row["CALLSIGN_ROOT"])
        if row["CALLSIGN_FORM"] == "LONGFORM_NONSTANDARD":
            _append_reason(
                reasons,
                idx,
                1,
                f"LONGFORM_CALLSIGN:{root}",
                first_pass,
                first_reason,
                matched_field,
                matched_value,
                "ARCID",
                _clean(row["ARCID"]),
            )
            operational_interest[idx] = True
            if root in longform_roots:
                reasons[idx].append(f"KNOWN_LONGFORM_ROOT:{root}")

    # Pass 2: Known ICAO/state/military callsigns of interest.
    for idx, row in work.iterrows():
        root = normalize_code(row["CALLSIGN_ROOT"])
        if row["CALLSIGN_FORM"] == "TRICODE_STYLE" and root in interest_codes:
            _append_reason(
                reasons,
                idx,
                2,
                f"KNOWN_ICAO_INTEREST:{root}",
                first_pass,
                first_reason,
                matched_field,
                matched_value,
                "ARCID",
                _clean(row["ARCID"]),
            )
            operational_interest[idx] = True

    # Pass 3: Known sensitive registrations.
    for idx, row in work.iterrows():
        reg = normalize_code(row["RM"])
        arcid = normalize_code(row["ARCID"])
        if reg and reg in sensitive_regs:
            _append_reason(
                reasons,
                idx,
                3,
                f"KNOWN_SENSITIVE_REG:{reg}",
                first_pass,
                first_reason,
                matched_field,
                matched_value,
                "RM",
                _clean(row["RM"]),
            )
            operational_interest[idx] = True
        elif (
            row["CALLSIGN_FORM"] == "REGISTRATION_CALLSIGN" and arcid in sensitive_regs
        ):
            _append_reason(
                reasons,
                idx,
                3,
                f"KNOWN_SENSITIVE_REG:{arcid}",
                first_pass,
                first_reason,
                matched_field,
                matched_value,
                "ARCID",
                _clean(row["ARCID"]),
            )
            operational_interest[idx] = True

    # Pass 4: Aircraft of interest.
    for idx, row in work.iterrows():
        atyp = normalize_code(row["ATYP"])
        if atyp in interest_types:
            _append_reason(
                reasons,
                idx,
                4,
                f"INTEREST_AIRCRAFT_TYPE:{atyp}",
                first_pass,
                first_reason,
                matched_field,
                matched_value,
                "ATYP",
                _clean(row["ATYP"]),
            )
            operational_interest[idx] = True

    # Pass 5: Unknown operator tricodes.
    for idx, row in work.iterrows():
        root = normalize_code(row["CALLSIGN_ROOT"])
        if (
            row["CALLSIGN_FORM"] == "TRICODE_STYLE"
            and root
            and root not in operator_codes
            and first_pass[idx] is None
        ):
            _append_reason(
                reasons,
                idx,
                5,
                f"UNKNOWN_OPERATOR_TRICODE:{root}",
                first_pass,
                first_reason,
                matched_field,
                matched_value,
                "ARCID",
                _clean(row["ARCID"]),
            )
            vkb_update_required[idx] = True
            operational_interest[idx] = True

    # Pass 6: VKB location check.
    for idx, row in work.iterrows():
        for field in ("ADEP", "ADES"):
            code = normalize_code(row[field])
            if code and code not in location_codes:
                _append_reason(
                    reasons,
                    idx,
                    6,
                    f"{field}_NOT_IN_VKB:{code}",
                    first_pass,
                    first_reason,
                    matched_field,
                    matched_value,
                    field,
                    _clean(row[field]),
                )
                vkb_update_required[idx] = True

    # Pass 7: VKB military/state locations.
    for idx, row in work.iterrows():
        for field in ("ADEP", "ADES"):
            code = normalize_code(row[field])
            if code in military_location_codes:
                _append_reason(
                    reasons,
                    idx,
                    7,
                    f"{field}_VKB_MILITARY:{code}",
                    first_pass,
                    first_reason,
                    matched_field,
                    matched_value,
                    field,
                    _clean(row[field]),
                )
                operational_interest[idx] = True

    # Pass 8: ZZZZ aerodromes.
    for idx, row in work.iterrows():
        for field in ("ADEP", "ADES"):
            if normalize_code(row[field]) == "ZZZZ":
                _append_reason(
                    reasons,
                    idx,
                    8,
                    f"{field}_ZZZZ",
                    first_pass,
                    first_reason,
                    matched_field,
                    matched_value,
                    field,
                    _clean(row[field]),
                )
                operational_interest[idx] = True

    # Pass 9: Sensitive regions.
    for idx, row in work.iterrows():
        for field in ("ADEP", "ADES"):
            code = normalize_code(row[field])
            prefix = _starts_with_any_prefix(code, region_prefix_to_name.keys())
            if prefix:
                region_name = region_prefix_to_name[prefix]
                _append_reason(
                    reasons,
                    idx,
                    9,
                    f"{field}_SENSITIVE_REGION:{region_name}",
                    first_pass,
                    first_reason,
                    matched_field,
                    matched_value,
                    field,
                    _clean(row[field]),
                )
                operational_interest[idx] = True

    work["EXTRACTED_PASS"] = [
        PASS_NAMES.get(pass_no, "") if pass_no else "" for pass_no in first_pass
    ]
    work["FIRST_MATCH_REASON"] = first_reason
    work["ALL_MATCH_REASONS"] = ["; ".join(row_reasons) for row_reasons in reasons]
    work["MATCHED_FIELD"] = matched_field
    work["MATCHED_VALUE"] = matched_value
    work["EXTRACTED_FLAG"] = [bool(pass_no) for pass_no in first_pass]
    work["VKB_UPDATE_REQUIRED"] = vkb_update_required
    work["OPERATIONAL_INTEREST"] = operational_interest

    output_columns = original_columns + [
        column for column in ADDED_COLUMNS if column not in original_columns
    ]
    work = work[output_columns]

    extracted = work[work["EXTRACTED_FLAG"]].copy()
    remainder = work[~work["EXTRACTED_FLAG"]].copy()

    missing_operator_counts = (
        work["ALL_MATCH_REASONS"]
        .str.extractall(r"UNKNOWN_OPERATOR_TRICODE:([^;]+)")[0]
        .value_counts()
        .sort_index()
        if row_count
        else pd.Series(dtype=int)
    )
    missing_location_counts = (
        work["ALL_MATCH_REASONS"]
        .str.extractall(r"(?:ADEP|ADES)_NOT_IN_VKB:([^;]+)")[0]
        .value_counts()
        .sort_index()
        if row_count
        else pd.Series(dtype=int)
    )
    update_rows = [
        {"CANDIDATE_TYPE": "OPERATOR_TRICODE", "CODE": code, "COUNT": int(count)}
        for code, count in missing_operator_counts.items()
    ] + [
        {"CANDIDATE_TYPE": "LOCATION", "CODE": code, "COUNT": int(count)}
        for code, count in missing_location_counts.items()
    ]
    update_candidates = pd.DataFrame(
        update_rows, columns=["CANDIDATE_TYPE", "CODE", "COUNT"]
    )

    summary_rows = [
        {"METRIC": "total input rows", "VALUE": row_count},
        {"METRIC": "extracted rows", "VALUE": len(extracted)},
        {"METRIC": "remainder rows", "VALUE": len(remainder)},
        {
            "METRIC": "missing operator tricodes",
            "VALUE": int(len(missing_operator_counts)),
        },
        {"METRIC": "missing locations", "VALUE": int(len(missing_location_counts))},
        {
            "METRIC": "output timestamp",
            "VALUE": datetime.now().isoformat(timespec="seconds"),
        },
    ]
    for pass_name, count in (
        work.loc[work["EXTRACTED_FLAG"], "EXTRACTED_PASS"]
        .value_counts()
        .sort_index()
        .items()
    ):
        summary_rows.append(
            {"METRIC": f"count by first pass: {pass_name}", "VALUE": int(count)}
        )
    summary = pd.DataFrame(summary_rows, columns=["METRIC", "VALUE"])

    sheets: dict[str, pd.DataFrame] = {
        "01_EXTRACTED_ALL": extracted,
        "02_LONGFORM_CALLSIGNS": extracted[
            extracted["EXTRACTED_PASS"] == PASS_NAMES[1]
        ],
        "03_KNOWN_ICAO_INTEREST": extracted[
            extracted["EXTRACTED_PASS"] == PASS_NAMES[2]
        ],
        "04_KNOWN_SENSITIVE_REGS": extracted[
            extracted["EXTRACTED_PASS"] == PASS_NAMES[3]
        ],
        "05_INTEREST_AIRCRAFT_TYPES": extracted[
            extracted["EXTRACTED_PASS"] == PASS_NAMES[4]
        ],
        "06_UNKNOWN_OPERATOR_TRICODES": extracted[
            extracted["EXTRACTED_PASS"] == PASS_NAMES[5]
        ],
        "07_VKB_LOCATION_GAPS": work[
            work["ALL_MATCH_REASONS"].map(
                lambda value: _reason_contains(
                    value, ("ADEP_NOT_IN_VKB", "ADES_NOT_IN_VKB")
                )
            )
        ],
        "08_VKB_MILITARY_LOCATIONS": work[
            work["ALL_MATCH_REASONS"].map(
                lambda value: _reason_contains(
                    value, ("ADEP_VKB_MILITARY", "ADES_VKB_MILITARY")
                )
            )
        ],
        "09_ZZZZ_AERODROMES": work[
            work["ALL_MATCH_REASONS"].map(
                lambda value: _reason_contains(value, ("ADEP_ZZZZ", "ADES_ZZZZ"))
            )
        ],
        "10_SENSITIVE_REGIONS": work[
            work["ALL_MATCH_REASONS"].map(
                lambda value: _reason_contains(
                    value, ("ADEP_SENSITIVE_REGION", "ADES_SENSITIVE_REGION")
                )
            )
        ],
        "11_REMAINDER_UNEXTRACTED": remainder,
        "12_VKB_UPDATE_CANDIDATES": update_candidates,
        "13_SUMMARY": summary,
    }

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{input_path.stem}_triaged.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        _format_workbook(writer.book)

    return TriageResult(
        output_path=output_path,
        warnings=warnings,
        total_rows=row_count,
        extracted_rows=len(extracted),
        remainder_rows=len(remainder),
    )


def _format_workbook(workbook) -> None:
    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        if worksheet.max_row >= 1 and worksheet.max_column >= 1:
            worksheet.auto_filter.ref = worksheet.dimensions
        for column_cells in worksheet.columns:
            column_letter = get_column_letter(column_cells[0].column)
            max_length = 0
            for cell in column_cells:
                max_length = max(
                    max_length, len(str(cell.value)) if cell.value is not None else 0
                )
            worksheet.column_dimensions[column_letter].width = min(
                max(max_length + 2, 10), 45
            )


def open_folder(path: str | Path) -> None:
    folder = Path(path)
    if platform.system() == "Windows":
        os.startfile(folder)  # type: ignore[attr-defined]
    elif platform.system() == "Darwin":
        subprocess.run(["open", str(folder)], check=False)
    else:
        subprocess.run(["xdg-open", str(folder)], check=False)
