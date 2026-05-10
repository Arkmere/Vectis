"""Validate Vectis v0.1 canonical NM/NOP sample input contract."""

from __future__ import annotations

import ast
import csv
from pathlib import Path

EXPECTED_CANONICAL_HEADERS = [
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


def load_required_columns() -> list[str]:
    """Read REQUIRED_COLUMNS without importing pandas-dependent triage_engine."""
    module = ast.parse(Path("triage_engine.py").read_text())
    for node in module.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "REQUIRED_COLUMNS":
                    value = ast.literal_eval(node.value)
                    if not isinstance(value, list) or not all(
                        isinstance(item, str) for item in value
                    ):
                        raise TypeError(
                            "triage_engine.REQUIRED_COLUMNS must be a list[str]"
                        )
                    return value
    raise AssertionError("triage_engine.py does not define REQUIRED_COLUMNS")


def load_sample_header() -> list[str]:
    with Path("input/sample_nm.csv").open(newline="") as sample_file:
        return next(csv.reader(sample_file))


def main() -> None:
    required_columns = load_required_columns()
    assert "REG" not in required_columns, "REG must not be accepted or required in v0.1"
    assert "RM" in required_columns, "RM must be the v0.1 registration-mark field"
    assert (
        required_columns == EXPECTED_CANONICAL_HEADERS
    ), "REQUIRED_COLUMNS must match canonical NM/NOP headers"
    assert (
        load_sample_header() == EXPECTED_CANONICAL_HEADERS
    ), "sample_nm.csv header must match canonical NM/NOP headers"
    print("canonical_schema_validation_passed")


if __name__ == "__main__":
    main()
