"""Deterministic pass-based triage engine for Vectis v0.1."""

from __future__ import annotations

import hashlib
import os
import platform
import re
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
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
    "EXACT_ROW_HASH",
    "EXACT_DUPLICATE_GROUP_SIZE",
    "IS_FIRST_EXACT_DUPLICATE",
    "EXACT_DUPLICATE_EXCESS",
    "MOVEMENT_KEY_STRICT",
    "STRICT_MOVEMENT_GROUP_SIZE",
    "IS_FIRST_STRICT_MOVEMENT",
    "MOVEMENT_KEY_LOOSE",
    "LOOSE_MOVEMENT_GROUP_SIZE",
    "IS_FIRST_LOOSE_MOVEMENT",
    "IS_REPRESENTATIVE_MOVEMENT_ROW",
    "REPRESENTATIVE_SELECTION_REASON",
    "MOVEMENT_GROUP_VARIATION_SUMMARY",
    "EXTRACTED_PASS",
    "FIRST_MATCH_REASON",
    "ALL_MATCH_REASONS",
    "MATCHED_FIELD",
    "MATCHED_VALUE",
    "EXTRACTED_FLAG",
    "VKB_UPDATE_REQUIRED",
    "OPERATIONAL_INTEREST",
    "ANALYST_SCORE",
    "ANALYST_BAND",
    "VKB_HYGIENE_ONLY",
    "MULTI_SIGNAL",
    "PRIMARY_INTELLIGENCE_REASON",
    "INTELLIGENCE_REASONS",
    "ANALYST_EXPLANATION_SHORT",
    "WHAT_TO_DO_NEXT",
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
    "vkb/FDMS_LOCATIONS.csv": [
        "ICAO CODE",
        "IATA CODE",
        "ALTERNATIVE CODE",
        "LOCATION SERVED",
        "AIRPORT",
        "ALTERNATIVE NAME",
        "HISTORICAL NAME",
        "COUNTRY",
        "TYPE",
        "USER",
        "NOTES",
        "ICAO REGION",
    ],
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
    return re.sub(r"[\s-]+", "", _clean(value).replace("\ufeff", "").upper())


def _row_value(value: object) -> str:
    return _clean(value)


def _hash_values(values: Iterable[object]) -> str:
    normalised = [_row_value(value) for value in values]
    return hashlib.sha256("\x1f".join(normalised).encode("utf-8")).hexdigest()


def _movement_key(row: pd.Series, fields: Iterable[str]) -> str:
    return "|".join(normalize_code(row.get(field, "")) for field in fields)


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

    Prefer the full FDMS locations dataset if present:
        vkb/FDMS_LOCATIONS.csv

    Fall back to the older partial B_E_L extract only if the full file is absent
    or produces no usable ICAO keys.

    The FDMS locations file uses ICAO CODE as the key. USER drives
    military/state/dual-use logic; TYPE is descriptive and intentionally not
    used for extraction logic.
    """
    location_sources = [
        "vkb/FDMS_LOCATIONS.csv",
        "vkb/FDMS_LOCATIONS_B_E_L.csv",
    ]

    for source in location_sources:
        if source in refs:
            fdms_locations = refs[source].copy()
        else:
            fdms_locations = _read_optional_reference(
                base_dir,
                source,
                ["ICAO CODE", "USER", "TYPE"],
            )

        if "ICAO CODE" not in fdms_locations.columns:
            continue

        if "USER" not in fdms_locations.columns:
            fdms_locations["USER"] = ""

        fdms_locations["ICAO_NORM"] = fdms_locations["ICAO CODE"].map(normalize_code)

        known_locations = set(fdms_locations["ICAO_NORM"]) - {""}
        interesting_locations = set(
            fdms_locations.loc[
                fdms_locations["USER"].map(
                    lambda value: normalize_code(value)
                    in {"MILITARY", "STATE", "DUAL", "DUAL_USE"}
                ),
                "ICAO_NORM",
            ]
        ) - {""}

        if known_locations:
            return known_locations, interesting_locations

    legacy = _read_optional_reference(
        base_dir,
        "vkb/locations.csv",
        LEGACY_VKB_REFERENCE_SPECS["vkb/locations.csv"],
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


SCORE_WEIGHTS = {
    "KNOWN_ICAO_INTEREST": 80,
    "LONGFORM_CALLSIGN": 70,
    "VKB_MILITARY": 65,
    "SENSITIVE_REGION": 60,
    "INTEREST_AIRCRAFT_TYPE": 50,
    "ZZZZ": 40,
    "UNKNOWN_OPERATOR_TRICODE": 25,
    "LOCATION_GAP": 5,
}

REASON_LABELS = {
    "KNOWN_ICAO_INTEREST": "Known interest callsign",
    "LONGFORM_CALLSIGN": "Longform/non-standard callsign",
    "VKB_MILITARY": "VKB military/state location",
    "SENSITIVE_REGION": "Sensitive region route",
    "INTEREST_AIRCRAFT_TYPE": "Aircraft type of interest",
    "ZZZZ": "ZZZZ aerodrome",
    "UNKNOWN_OPERATOR_TRICODE": "Unknown operator tricode",
    "LOCATION_GAP": "Unknown location in VKB",
    "KNOWN_SENSITIVE_REG": "Known sensitive registration",
}


def _signal_key(reason: str) -> str:
    if "KNOWN_ICAO_INTEREST" in reason:
        return "KNOWN_ICAO_INTEREST"
    if "LONGFORM_CALLSIGN" in reason:
        return "LONGFORM_CALLSIGN"
    if "VKB_MILITARY" in reason:
        return "VKB_MILITARY"
    if "SENSITIVE_REGION" in reason:
        return "SENSITIVE_REGION"
    if "INTEREST_AIRCRAFT_TYPE" in reason:
        return "INTEREST_AIRCRAFT_TYPE"
    if "ZZZZ" in reason:
        return "ZZZZ"
    if "UNKNOWN_OPERATOR_TRICODE" in reason:
        return "UNKNOWN_OPERATOR_TRICODE"
    if "NOT_IN_VKB" in reason:
        return "LOCATION_GAP"
    if "KNOWN_SENSITIVE_REG" in reason:
        return "KNOWN_ICAO_INTEREST"
    return ""


def _score_reasons(row_reasons: list[str]) -> tuple[int, str, bool, bool, str, str, str]:
    keys = {_signal_key(reason) for reason in row_reasons}
    keys.discard("")
    non_hygiene_keys = keys - {"LOCATION_GAP"}
    hygiene_only = bool(keys) and not non_hygiene_keys
    score = 0 if hygiene_only else sum(SCORE_WEIGHTS.get(key, 0) for key in keys)
    if hygiene_only:
        score = 5
    if score >= 90:
        band = "PRIORITY_INTELLIGENCE_REVIEW"
    elif score >= 60:
        band = "OPERATIONALLY_INTERESTING"
    elif score >= 30:
        band = "REVIEW_IF_TIME_PERMITS"
    elif score >= 1:
        band = "VKB_ENRICHMENT_INTEREST"
    else:
        band = "VKB_HYGIENE_ONLY"
    primary_key = max(keys, key=lambda key: SCORE_WEIGHTS.get(key, 0), default="")
    primary = REASON_LABELS.get(primary_key, "")
    intel = "; ".join(REASON_LABELS.get(key, key) for key in sorted(keys, key=lambda key: SCORE_WEIGHTS.get(key, 0), reverse=True))
    multi = len(non_hygiene_keys) >= 2
    if hygiene_only:
        explanation = "Unknown location only; no other intelligence indicators. Treat as VKB enrichment candidate."
        action = "vkb_hygiene_only"
    elif primary_key == "UNKNOWN_OPERATOR_TRICODE" and not (non_hygiene_keys - {"UNKNOWN_OPERATOR_TRICODE"}):
        explanation = "Unknown operator tricode with no higher-value intelligence indicators. Enrich operator if relevant."
        action = "enrich_operator"
    elif "LOCATION_GAP" in keys and not non_hygiene_keys:
        explanation = "Unknown location only; enrich VKB location data."
        action = "enrich_location"
    else:
        explanation = intel or "No intelligence indicators."
        action = "review_movement"
    return score, band, hygiene_only, multi, primary, intel, explanation, action


def _variation_summary(group: pd.DataFrame) -> str:
    fields = ["ARCID", "ATYP", "RM", "ADEP", "ADES", "ALT1", "ALT2", "Delay", "MSG", "E/CTOT"]
    differing = []
    for field in fields:
        if field in group.columns and group[field].map(_row_value).nunique(dropna=False) > 1:
            differing.append(field)
    return "Differing fields: " + ", ".join(differing) if differing else "No material field differences"


def _add_identity_metadata(work: pd.DataFrame, original_columns: list[str]) -> None:
    work["EXACT_ROW_HASH"] = work[original_columns].apply(lambda row: _hash_values(row.values), axis=1)
    work["EXACT_DUPLICATE_GROUP_SIZE"] = work.groupby("EXACT_ROW_HASH")["EXACT_ROW_HASH"].transform("size")
    work["IS_FIRST_EXACT_DUPLICATE"] = ~work.duplicated("EXACT_ROW_HASH")
    work["EXACT_DUPLICATE_EXCESS"] = work["EXACT_DUPLICATE_GROUP_SIZE"] - 1
    work["MOVEMENT_KEY_STRICT"] = work.apply(lambda row: _movement_key(row, ["ARCID", "ATYP", "RM", "ADEP", "ADES", "IOBT"]), axis=1)
    work["STRICT_MOVEMENT_GROUP_SIZE"] = work.groupby("MOVEMENT_KEY_STRICT")["MOVEMENT_KEY_STRICT"].transform("size")
    work["IS_FIRST_STRICT_MOVEMENT"] = ~work.duplicated("MOVEMENT_KEY_STRICT")
    work["MOVEMENT_KEY_LOOSE"] = work.apply(lambda row: _movement_key(row, ["ARCID", "ATYP", "ADEP", "ADES", "IOBT"]), axis=1)
    work["LOOSE_MOVEMENT_GROUP_SIZE"] = work.groupby("MOVEMENT_KEY_LOOSE")["MOVEMENT_KEY_LOOSE"].transform("size")
    work["IS_FIRST_LOOSE_MOVEMENT"] = ~work.duplicated("MOVEMENT_KEY_LOOSE")


def _add_representatives(work: pd.DataFrame) -> None:
    work["IS_REPRESENTATIVE_MOVEMENT_ROW"] = False
    work["REPRESENTATIVE_SELECTION_REASON"] = ""
    work["MOVEMENT_GROUP_VARIATION_SUMMARY"] = ""
    for _, group in work.groupby("MOVEMENT_KEY_STRICT", sort=False):
        scoring = group.copy()
        scoring["_NON_HYGIENE_REASON_COUNT"] = scoring["ALL_MATCH_REASONS"].map(lambda value: len({_signal_key(reason.strip()) for reason in str(value).split(";") if _signal_key(reason.strip()) and _signal_key(reason.strip()) != "LOCATION_GAP"}))
        scoring["_HAS_RM"] = scoring["RM"].map(lambda value: bool(normalize_code(value)) and normalize_code(value) != "ONFILE")
        route_fields = [field for field in ["ADEP", "ADES", "ALT1", "ALT2", "ATYP", "IOBT"] if field in scoring.columns]
        scoring["_ROUTE_COMPLETENESS"] = scoring[route_fields].apply(lambda row: sum(bool(_row_value(value)) for value in row.values), axis=1)
        ordered = scoring.sort_values(["ANALYST_SCORE", "_NON_HYGIENE_REASON_COUNT", "_HAS_RM", "_ROUTE_COMPLETENESS"], ascending=[False, False, False, False], kind="mergesort")
        rep_idx = ordered.index[0]
        reason = "highest_score"
        if len(group) > 1:
            top = ordered.iloc[0]
            if top["ANALYST_SCORE"] == 0 and top["_NON_HYGIENE_REASON_COUNT"] == 0 and not top["_HAS_RM"]:
                reason = "first_row_fallback"
            elif top["_HAS_RM"]:
                reason = "most_complete_rm" if top["ANALYST_SCORE"] == ordered["ANALYST_SCORE"].max() else "highest_score"
        work.loc[rep_idx, "IS_REPRESENTATIVE_MOVEMENT_ROW"] = True
        work.loc[rep_idx, "REPRESENTATIVE_SELECTION_REASON"] = reason
        work.loc[group.index, "MOVEMENT_GROUP_VARIATION_SUMMARY"] = _variation_summary(group)


def _pass_sheet(work: pd.DataFrame, token: str, reason_label: str) -> pd.DataFrame:
    subset = work[work["ALL_MATCH_REASONS"].map(lambda value: token in str(value))].copy()
    def matching_reason(value: object) -> str:
        matches = [reason.strip() for reason in str(value).split(";") if token in reason]
        return "; ".join(matches)
    subset["SHEET_MATCH_REASON"] = subset["ALL_MATCH_REASONS"].map(matching_reason)
    subset["SHEET_MATCH_FIELD"] = subset["SHEET_MATCH_REASON"].map(lambda value: "ARCID" if any(t in value for t in ["CALLSIGN", "ICAO", "OPERATOR"]) else ("ATYP" if "AIRCRAFT" in value else ("RM" if "REG" in value else "ADEP/ADES/ALT1/ALT2")))
    subset["SHEET_MATCH_VALUE"] = subset["SHEET_MATCH_REASON"].map(lambda value: "; ".join(reason.split(":", 1)[-1] for reason in str(value).split("; ") if reason))
    subset["FIRST_MATCH_FIELD"] = subset["MATCHED_FIELD"]
    subset["FIRST_MATCH_VALUE"] = subset["MATCHED_VALUE"]
    return subset

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
    _add_identity_metadata(work, original_columns)

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
        for field in ("ADEP", "ADES", "ALT1", "ALT2"):
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

    scoring = [_score_reasons(row_reasons) for row_reasons in reasons]
    work["ANALYST_SCORE"] = [item[0] for item in scoring]
    work["ANALYST_BAND"] = [item[1] for item in scoring]
    work["VKB_HYGIENE_ONLY"] = [item[2] for item in scoring]
    work["MULTI_SIGNAL"] = [item[3] for item in scoring]
    work["PRIMARY_INTELLIGENCE_REASON"] = [item[4] for item in scoring]
    work["INTELLIGENCE_REASONS"] = [item[5] for item in scoring]
    work["ANALYST_EXPLANATION_SHORT"] = [item[6] for item in scoring]
    work["WHAT_TO_DO_NEXT"] = [item[7] for item in scoring]
    _add_representatives(work)

    output_columns = original_columns + [
        column for column in ADDED_COLUMNS if column not in original_columns
    ]
    work = work[output_columns]

    extracted = work[work["EXTRACTED_FLAG"]].copy()
    remainder = work[~work["EXTRACTED_FLAG"]].copy()
    representative = work[work["IS_REPRESENTATIVE_MOVEMENT_ROW"]].copy()
    analyst_priority = representative[
        (~representative["VKB_HYGIENE_ONLY"])
        & (
            representative["ANALYST_BAND"].isin(["PRIORITY_INTELLIGENCE_REVIEW", "OPERATIONALLY_INTERESTING"])
            | representative["MULTI_SIGNAL"]
        )
    ].copy()
    analyst_priority = analyst_priority.sort_values(
        ["ANALYST_SCORE", "MULTI_SIGNAL", "IOBT", "CALLSIGN_ROOT", "ARCID"],
        ascending=[False, False, True, True, True],
        kind="mergesort",
    )
    multi_signal = representative[representative["MULTI_SIGNAL"]].sort_values(
        "ANALYST_SCORE", ascending=False, kind="mergesort"
    )
    operational_all = extracted[extracted["OPERATIONAL_INTEREST"]].copy()
    hygiene_only = extracted[extracted["VKB_HYGIENE_ONLY"]].copy()

    unknown_operator_ranked = _build_unknown_operator_ranked(work)
    unknown_location_ranked = _build_unknown_location_ranked(work, region_prefix_to_name)
    vkb_load_audit = _build_vkb_load_audit(refs, base_dir)
    vkb_match_diagnostics = _build_vkb_match_diagnostics(work, operator_codes)

    summary = _build_summary(
        work,
        extracted,
        remainder,
        analyst_priority,
        multi_signal,
        unknown_operator_ranked,
        unknown_location_ranked,
    )
    readme = pd.DataFrame(
        [
            {"SECTION": "Purpose", "DETAIL": "Analyst Workbook v2 separates operational intelligence from VKB hygiene and preserves raw audit rows."},
            {"SECTION": "Raw vs unique", "DETAIL": "Raw rows count source records. Exact unique rows use EXACT_ROW_HASH. Strict movements use ARCID|ATYP|RM|ADEP|ADES|IOBT."},
            {"SECTION": "Score bands", "DETAIL": "90+ priority review; 60-89 operational; 30-59 review if time; 1-29 VKB enrichment; 0 hygiene only."},
            {"SECTION": "Representative rows", "DETAIL": "Analyst-facing priority sheets use one selected row per strict movement group; full raw rows remain in 09_EXTRACTED_ALL_AUDIT."},
        ]
    )

    sheets: dict[str, pd.DataFrame] = {
        "00_README": readme,
        "01_ANALYST_PRIORITY": analyst_priority,
        "02_MULTI_SIGNAL_REVIEW": multi_signal,
        "03_OPERATIONAL_INTEREST_ALL": operational_all,
        "04_VKB_HYGIENE_ONLY": hygiene_only,
        "05_UNKNOWN_OPERATOR_RANKED": unknown_operator_ranked,
        "06_UNKNOWN_LOCATION_RANKED": unknown_location_ranked,
        "07_VKB_MATCH_DIAGNOSTICS": vkb_match_diagnostics,
        "08_VKB_LOAD_AUDIT": vkb_load_audit,
        "09_EXTRACTED_ALL_AUDIT": extracted,
        "10_REMAINDER_UNEXTRACTED": remainder,
        "11_PASS_LONGFORM_CALLSIGNS": _pass_sheet(work, "LONGFORM_CALLSIGN", "LONGFORM_CALLSIGNS"),
        "12_PASS_KNOWN_INTEREST": _pass_sheet(work, "KNOWN_ICAO_INTEREST", "KNOWN_INTEREST"),
        "13_PASS_SENSITIVE_REGS": _pass_sheet(work, "KNOWN_SENSITIVE_REG", "SENSITIVE_REGS"),
        "14_PASS_INTEREST_TYPES": _pass_sheet(work, "INTEREST_AIRCRAFT_TYPE", "INTEREST_TYPES"),
        "15_PASS_UNKNOWN_OPERATORS": _pass_sheet(work, "UNKNOWN_OPERATOR_TRICODE", "UNKNOWN_OPERATORS"),
        "16_PASS_LOCATION_GAPS": _pass_sheet(work, "NOT_IN_VKB", "LOCATION_GAPS"),
        "17_PASS_MILITARY_LOCATIONS": _pass_sheet(work, "VKB_MILITARY", "MILITARY_LOCATIONS"),
        "18_PASS_ZZZZ": _pass_sheet(work, "ZZZZ", "ZZZZ"),
        "19_PASS_SENSITIVE_REGIONS": _pass_sheet(work, "SENSITIVE_REGION", "SENSITIVE_REGIONS"),
        "20_SUMMARY": summary,
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


def _sample_values(series: pd.Series, limit: int = 10, exclude_onfile: bool = False) -> str:
    values: list[str] = []
    for value in series:
        cleaned = _row_value(value)
        if exclude_onfile and normalize_code(cleaned) in {"", "ONFILE"}:
            continue
        if cleaned and cleaned not in values:
            values.append(cleaned)
        if len(values) >= limit:
            break
    return "; ".join(values)


def _build_unknown_operator_ranked(work: pd.DataFrame) -> pd.DataFrame:
    rows = []
    mask = work["ALL_MATCH_REASONS"].map(lambda value: "UNKNOWN_OPERATOR_TRICODE" in str(value))
    for tricode, group in work[mask].groupby("CALLSIGN_ROOT"):
        code = normalize_code(tricode)
        if not code:
            continue
        unique_strict = group["MOVEMENT_KEY_STRICT"].nunique()
        op_interest = int(group["OPERATIONAL_INTEREST"].sum())
        if unique_strict >= 20 or (op_interest > 0 and unique_strict >= 5):
            priority = "HIGH"
        elif unique_strict >= 5:
            priority = "MEDIUM"
        else:
            priority = "LOW"
        routes = group.apply(lambda row: f"{_row_value(row['ADEP'])}→{_row_value(row['ADES'])}", axis=1)
        rows.append({
            "TRICODE": code,
            "COUNT_RAW_ROWS": len(group),
            "COUNT_UNIQUE_STRICT_MOVEMENTS": unique_strict,
            "COUNT_UNIQUE_LOOSE_MOVEMENTS": group["MOVEMENT_KEY_LOOSE"].nunique(),
            "COUNT_EXACT_DUPLICATE_GROUPS": group["EXACT_ROW_HASH"].nunique(),
            "FIRST_SEEN_IOBT": _sample_values(pd.Series([group["IOBT"].map(_row_value).replace("", pd.NA).dropna().min() if "IOBT" in group else ""]), 1),
            "LAST_SEEN_IOBT": _sample_values(pd.Series([group["IOBT"].map(_row_value).replace("", pd.NA).dropna().max() if "IOBT" in group else ""]), 1),
            "SAMPLE_ARCIDS": _sample_values(group["ARCID"]),
            "SAMPLE_ROUTES": _sample_values(routes),
            "SAMPLE_ATYP": _sample_values(group["ATYP"]),
            "SAMPLE_RM": _sample_values(group["RM"], exclude_onfile=True),
            "OPERATIONAL_INTEREST_ROWS": op_interest,
            "MAX_ANALYST_SCORE": int(group["ANALYST_SCORE"].max()),
            "VKB_UPDATE_PRIORITY": priority,
            "NOTES": "",
        })
    columns = ["TRICODE", "COUNT_RAW_ROWS", "COUNT_UNIQUE_STRICT_MOVEMENTS", "COUNT_UNIQUE_LOOSE_MOVEMENTS", "COUNT_EXACT_DUPLICATE_GROUPS", "FIRST_SEEN_IOBT", "LAST_SEEN_IOBT", "SAMPLE_ARCIDS", "SAMPLE_ROUTES", "SAMPLE_ATYP", "SAMPLE_RM", "OPERATIONAL_INTEREST_ROWS", "MAX_ANALYST_SCORE", "VKB_UPDATE_PRIORITY", "NOTES"]
    df = pd.DataFrame(rows, columns=columns)
    return df.sort_values(["COUNT_UNIQUE_STRICT_MOVEMENTS", "COUNT_RAW_ROWS", "MAX_ANALYST_SCORE", "TRICODE"], ascending=[False, False, False, True], kind="mergesort") if not df.empty else df


def _location_gap_codes(row: pd.Series) -> list[tuple[str, str]]:
    pairs = []
    reasons = str(row.get("ALL_MATCH_REASONS", ""))
    for field in ("ADEP", "ADES", "ALT1", "ALT2"):
        prefix = f"{field}_NOT_IN_VKB:"
        for reason in reasons.split("; "):
            if reason.startswith(prefix):
                pairs.append((field, reason.split(":", 1)[1]))
    return pairs


def _build_unknown_location_ranked(work: pd.DataFrame, region_prefix_to_name: dict[str, str]) -> pd.DataFrame:
    rows = []
    exploded = []
    for idx, row in work.iterrows():
        for field, code in _location_gap_codes(row):
            exploded.append({"_IDX": idx, "FIELD": field, "ICAO": code})
    if not exploded:
        columns = ["ICAO", "COUNT_RAW_ROWS", "COUNT_UNIQUE_STRICT_MOVEMENTS", "COUNT_AS_ADEP", "COUNT_AS_ADES", "COUNT_AS_ALT1", "COUNT_AS_ALT2", "FIRST_SEEN_IOBT", "LAST_SEEN_IOBT", "SAMPLE_ROUTES", "SAMPLE_OPERATORS", "SAMPLE_TYPES", "SENSITIVE_REGION_HINT", "ZZZZ_RELATED", "MAX_ANALYST_SCORE", "VKB_UPDATE_PRIORITY", "NOTES"]
        return pd.DataFrame(columns=columns)
    exploded_df = pd.DataFrame(exploded)
    for code, code_rows in exploded_df.groupby("ICAO"):
        group = work.loc[code_rows["_IDX"].unique()]
        unique_strict = group["MOVEMENT_KEY_STRICT"].nunique()
        priority = "HIGH" if unique_strict >= 20 else ("MEDIUM" if unique_strict >= 5 else "LOW")
        routes = group.apply(lambda row: f"{_row_value(row['ADEP'])}→{_row_value(row['ADES'])}", axis=1)
        prefix = _starts_with_any_prefix(code, region_prefix_to_name.keys())
        rows.append({
            "ICAO": code,
            "COUNT_RAW_ROWS": len(group),
            "COUNT_UNIQUE_STRICT_MOVEMENTS": unique_strict,
            "COUNT_AS_ADEP": int((code_rows["FIELD"] == "ADEP").sum()),
            "COUNT_AS_ADES": int((code_rows["FIELD"] == "ADES").sum()),
            "COUNT_AS_ALT1": int((code_rows["FIELD"] == "ALT1").sum()),
            "COUNT_AS_ALT2": int((code_rows["FIELD"] == "ALT2").sum()),
            "FIRST_SEEN_IOBT": _sample_values(pd.Series([group["IOBT"].map(_row_value).replace("", pd.NA).dropna().min() if "IOBT" in group else ""]), 1),
            "LAST_SEEN_IOBT": _sample_values(pd.Series([group["IOBT"].map(_row_value).replace("", pd.NA).dropna().max() if "IOBT" in group else ""]), 1),
            "SAMPLE_ROUTES": _sample_values(routes),
            "SAMPLE_OPERATORS": _sample_values(group["CALLSIGN_ROOT"]),
            "SAMPLE_TYPES": _sample_values(group["ATYP"]),
            "SENSITIVE_REGION_HINT": region_prefix_to_name.get(prefix, "") if prefix else "",
            "ZZZZ_RELATED": bool((group["ADEP"].map(normalize_code) == "ZZZZ").any() or (group["ADES"].map(normalize_code) == "ZZZZ").any()),
            "MAX_ANALYST_SCORE": int(group["ANALYST_SCORE"].max()),
            "VKB_UPDATE_PRIORITY": priority,
            "NOTES": "",
        })
    columns = ["ICAO", "COUNT_RAW_ROWS", "COUNT_UNIQUE_STRICT_MOVEMENTS", "COUNT_AS_ADEP", "COUNT_AS_ADES", "COUNT_AS_ALT1", "COUNT_AS_ALT2", "FIRST_SEEN_IOBT", "LAST_SEEN_IOBT", "SAMPLE_ROUTES", "SAMPLE_OPERATORS", "SAMPLE_TYPES", "SENSITIVE_REGION_HINT", "ZZZZ_RELATED", "MAX_ANALYST_SCORE", "VKB_UPDATE_PRIORITY", "NOTES"]
    df = pd.DataFrame(rows, columns=columns)
    return df.sort_values(["COUNT_UNIQUE_STRICT_MOVEMENTS", "COUNT_RAW_ROWS", "MAX_ANALYST_SCORE", "ICAO"], ascending=[False, False, False, True], kind="mergesort")


def _build_vkb_load_audit(refs: dict[str, pd.DataFrame], base_dir: Path) -> pd.DataFrame:
    rows = []
    key_columns = {
        "vkb/FDMS_CALLSIGNS_STANDARD.csv": "TRICODE",
        "vkb/FDMS_CALLSIGNS_NONSTANDARD CALLSIGNS.csv": "CALLSIGN",
        "vkb/FDMS_LOCATIONS_B_E_L.csv": "ICAO CODE",
        "vkb/FDMS_AIRCRAFT_TYPES.csv": "ICAO Type Designator",
        "config/interest_icao_codes.csv": "CODE",
        "config/interest_aircraft_types.csv": "ATYP",
        "config/sensitive_registrations.csv": "REG",
        "config/sensitive_regions.csv": "ICAO_PREFIX",
        "config/longform_roots.csv": "ROOT",
    }
    for resource, df in refs.items():
        key_col = key_columns.get(resource, df.columns[0] if len(df.columns) else "")
        key_found = key_col in df.columns
        keys = df[key_col].map(normalize_code) if key_found else pd.Series(dtype=str)
        rows.append({
            "RESOURCE": resource,
            "EXPECTED_PATH": str(base_dir / resource),
            "LOADED": (base_dir / resource).exists(),
            "ROW_COUNT": len(df),
            "USABLE_KEY_COUNT": int((keys != "").sum()) if key_found else 0,
            "EXPECTED_KEY_COLUMN": key_col,
            "KEY_COLUMN_FOUND": key_found,
            "DUPLICATE_KEY_COUNT": int(keys[keys != ""].duplicated().sum()) if key_found else 0,
            "BLANK_KEY_COUNT": int((keys == "").sum()) if key_found else 0,
            "NON_STANDARD_KEY_COUNT": int(keys[(keys != "") & (keys.str.len() != 3)].count()) if key_found and key_col in {"TRICODE", "CODE", "ROOT"} else 0,
            "NOTES": "",
            "SANITY_CODE": "",
            "PRESENT_IN_STANDARD_CALLSIGNS": "",
            "RAW_MATCH_VALUE": "",
            "NORMALISED_MATCH_VALUE": "",
        })
    std = refs.get("vkb/FDMS_CALLSIGNS_STANDARD.csv", pd.DataFrame())
    std_codes = std["TRICODE"].map(normalize_code) if "TRICODE" in std else pd.Series(dtype=str)
    for code in ["BAW", "EZY", "RYR", "UPS", "KAL", "RAM", "SVA", "THA", "TCA", "MSR"]:
        matches = std.loc[std_codes == code, "TRICODE"] if "TRICODE" in std else pd.Series(dtype=str)
        raw = _row_value(matches.iloc[0]) if not matches.empty else ""
        rows.append({
            "RESOURCE": "SANITY_STANDARD_CALLSIGNS",
            "EXPECTED_PATH": str(base_dir / "vkb/FDMS_CALLSIGNS_STANDARD.csv"),
            "LOADED": (base_dir / "vkb/FDMS_CALLSIGNS_STANDARD.csv").exists(),
            "ROW_COUNT": len(std),
            "USABLE_KEY_COUNT": int((std_codes != "").sum()) if len(std_codes) else 0,
            "EXPECTED_KEY_COLUMN": "TRICODE",
            "KEY_COLUMN_FOUND": "TRICODE" in std,
            "DUPLICATE_KEY_COUNT": "",
            "BLANK_KEY_COUNT": "",
            "NON_STANDARD_KEY_COUNT": "",
            "NOTES": "common callsign sanity check",
            "SANITY_CODE": code,
            "PRESENT_IN_STANDARD_CALLSIGNS": not matches.empty,
            "RAW_MATCH_VALUE": raw,
            "NORMALISED_MATCH_VALUE": normalize_code(raw),
        })
    return pd.DataFrame(rows)


def _build_vkb_match_diagnostics(work: pd.DataFrame, operator_codes: set[str]) -> pd.DataFrame:
    rows = []
    for code, group in work.groupby("CALLSIGN_ROOT"):
        norm = normalize_code(code)
        if not norm:
            status = "malformed"
            reason = "blank_code"
        elif not norm.isalpha():
            status = "malformed"
            reason = "non_alpha"
        elif norm in operator_codes:
            status = "matched"
            reason = ""
        else:
            status = "unmatched"
            reason = "not_in_loaded_tricode_set"
        near = []
        if status == "unmatched":
            near = [known for known in sorted(operator_codes) if known.strip() == norm or known.replace("\ufeff", "") == norm][:5]
        rows.append({
            "CODE_FROM_EXPORT": code,
            "NORMALISED_CODE": norm,
            "MATCH_STATUS": status,
            "MATCHED_VKB_FILE": "vkb/FDMS_CALLSIGNS_STANDARD.csv" if status == "matched" else "",
            "MATCHED_VKB_COLUMN": "TRICODE" if status == "matched" else "",
            "REASON_UNMATCHED": reason,
            "COUNT_RAW_ROWS": len(group),
            "COUNT_UNIQUE_STRICT_MOVEMENTS": group["MOVEMENT_KEY_STRICT"].nunique(),
            "POSSIBLE_NEAR_MATCHES": "; ".join(near),
        })
    df = pd.DataFrame(rows)
    return df.sort_values(["MATCH_STATUS", "COUNT_RAW_ROWS", "NORMALISED_CODE"], ascending=[True, False, True], kind="mergesort") if not df.empty else df


def _build_summary(work: pd.DataFrame, extracted: pd.DataFrame, remainder: pd.DataFrame, analyst_priority: pd.DataFrame, multi_signal: pd.DataFrame, unknown_operator_ranked: pd.DataFrame, unknown_location_ranked: pd.DataFrame) -> pd.DataFrame:
    rows = [
        {"BLOCK": "METRIC", "METRIC": "total input raw rows", "VALUE": len(work)},
        {"BLOCK": "METRIC", "METRIC": "exact unique input rows", "VALUE": work["EXACT_ROW_HASH"].nunique()},
        {"BLOCK": "METRIC", "METRIC": "exact duplicate excess rows", "VALUE": len(work) - work["EXACT_ROW_HASH"].nunique()},
        {"BLOCK": "METRIC", "METRIC": "unique strict movement count", "VALUE": work["MOVEMENT_KEY_STRICT"].nunique()},
        {"BLOCK": "METRIC", "METRIC": "unique loose movement count", "VALUE": work["MOVEMENT_KEY_LOOSE"].nunique()},
        {"BLOCK": "METRIC", "METRIC": "raw extracted rows", "VALUE": len(extracted)},
        {"BLOCK": "METRIC", "METRIC": "unique strict extracted movements", "VALUE": extracted["MOVEMENT_KEY_STRICT"].nunique()},
        {"BLOCK": "METRIC", "METRIC": "raw remainder rows", "VALUE": len(remainder)},
        {"BLOCK": "METRIC", "METRIC": "unique strict remainder movements", "VALUE": remainder["MOVEMENT_KEY_STRICT"].nunique()},
        {"BLOCK": "METRIC", "METRIC": "VKB hygiene only raw rows", "VALUE": int(work["VKB_HYGIENE_ONLY"].sum())},
        {"BLOCK": "METRIC", "METRIC": "VKB hygiene only unique strict movements", "VALUE": work.loc[work["VKB_HYGIENE_ONLY"], "MOVEMENT_KEY_STRICT"].nunique()},
        {"BLOCK": "METRIC", "METRIC": "analyst priority raw rows", "VALUE": len(analyst_priority)},
        {"BLOCK": "METRIC", "METRIC": "analyst priority unique strict movements", "VALUE": analyst_priority["MOVEMENT_KEY_STRICT"].nunique()},
        {"BLOCK": "METRIC", "METRIC": "multi-signal raw rows", "VALUE": len(multi_signal)},
        {"BLOCK": "METRIC", "METRIC": "multi-signal unique strict movements", "VALUE": multi_signal["MOVEMENT_KEY_STRICT"].nunique()},
        {"BLOCK": "METRIC", "METRIC": "unknown operator candidate count", "VALUE": len(unknown_operator_ranked)},
        {"BLOCK": "METRIC", "METRIC": "unknown location candidate count", "VALUE": len(unknown_location_ranked)},
        {"BLOCK": "METRIC", "METRIC": "output timestamp", "VALUE": datetime.now().isoformat(timespec="seconds")},
    ]
    def add_top(block: str, df: pd.DataFrame, key: str, value: str) -> None:
        if df.empty or key not in df or value not in df:
            return
        for _, row in df.head(10).iterrows():
            rows.append({"BLOCK": block, "METRIC": row[key], "VALUE": row[value]})
    add_top("Top unknown operators", unknown_operator_ranked, "TRICODE", "COUNT_UNIQUE_STRICT_MOVEMENTS")
    add_top("Top missing locations", unknown_location_ranked, "ICAO", "COUNT_UNIQUE_STRICT_MOVEMENTS")
    callsigns = work[work["OPERATIONAL_INTEREST"]].groupby("CALLSIGN_ROOT").agg(MAX_SCORE=("ANALYST_SCORE", "max"), COUNT=("CALLSIGN_ROOT", "size")).reset_index().sort_values(["MAX_SCORE", "COUNT"], ascending=[False, False])
    for _, row in callsigns.head(10).iterrows():
        rows.append({"BLOCK": "Top operational callsign roots", "METRIC": row["CALLSIGN_ROOT"], "VALUE": f"max_score={row['MAX_SCORE']}; count={row['COUNT']}"})
    routes = work[work["OPERATIONAL_INTEREST"]].assign(ROUTE=lambda df: df["ADEP"].map(_row_value) + "→" + df["ADES"].map(_row_value)).groupby("ROUTE").size().sort_values(ascending=False)
    for route, count in routes.head(10).items():
        rows.append({"BLOCK": "Top routes by operational interest", "METRIC": route, "VALUE": int(count)})
    types = work[work["ALL_MATCH_REASONS"].map(lambda value: "INTEREST_AIRCRAFT_TYPE" in str(value))].groupby("ATYP").size().sort_values(ascending=False)
    for atyp, count in types.head(10).items():
        rows.append({"BLOCK": "Top aircraft types of interest", "METRIC": atyp, "VALUE": int(count)})
    combos = work[work["MULTI_SIGNAL"]].groupby("INTELLIGENCE_REASONS").size().sort_values(ascending=False)
    for combo, count in combos.head(10).items():
        rows.append({"BLOCK": "Top multi-signal reason combinations", "METRIC": combo, "VALUE": int(count)})
    return pd.DataFrame(rows, columns=["BLOCK", "METRIC", "VALUE"])

def _format_workbook(workbook) -> None:
    band_fills = {
        "PRIORITY_INTELLIGENCE_REVIEW": PatternFill("solid", fgColor="F4CCCC"),
        "OPERATIONALLY_INTERESTING": PatternFill("solid", fgColor="FCE5CD"),
        "REVIEW_IF_TIME_PERMITS": PatternFill("solid", fgColor="FFF2CC"),
        "VKB_ENRICHMENT_INTEREST": PatternFill("solid", fgColor="D9EAD3"),
        "VKB_HYGIENE_ONLY": PatternFill("solid", fgColor="D9E2F3"),
    }
    wrap_headers = {
        "ANALYST_EXPLANATION_SHORT",
        "WHAT_TO_DO_NEXT",
        "INTELLIGENCE_REASONS",
        "ALL_MATCH_REASONS",
        "MOVEMENT_GROUP_VARIATION_SUMMARY",
    }
    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        if worksheet.max_row >= 1 and worksheet.max_column >= 1:
            worksheet.auto_filter.ref = worksheet.dimensions
        header_to_column = {}
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            header_to_column[str(cell.value)] = cell.column
        band_column = header_to_column.get("ANALYST_BAND")
        if band_column:
            for row in range(2, worksheet.max_row + 1):
                cell = worksheet.cell(row=row, column=band_column)
                if cell.value in band_fills:
                    cell.fill = band_fills[cell.value]
        wrap_columns = [column for header, column in header_to_column.items() if header in wrap_headers or "EXPLANATION" in header or "NOTES" in header]
        for column in wrap_columns:
            for row in range(1, worksheet.max_row + 1):
                worksheet.cell(row=row, column=column).alignment = Alignment(wrap_text=True, vertical="top")
        for column_cells in worksheet.columns:
            column_letter = get_column_letter(column_cells[0].column)
            max_length = 0
            for cell in column_cells:
                max_length = max(
                    max_length, len(str(cell.value)) if cell.value is not None else 0
                )
            worksheet.column_dimensions[column_letter].width = min(
                max(max_length + 2, 10), 60
            )


def open_folder(path: str | Path) -> None:
    folder = Path(path)
    if platform.system() == "Windows":
        os.startfile(folder)  # type: ignore[attr-defined]
    elif platform.system() == "Darwin":
        subprocess.run(["open", str(folder)], check=False)
    else:
        subprocess.run(["xdg-open", str(folder)], check=False)
