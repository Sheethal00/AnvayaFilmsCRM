#!/usr/bin/env python3
"""Lint schema/*.yaml against docs/SPEC.md rules.

Checks (see tests/test_schema_integrity.py for the pytest-wrapped version
of the same checks, run in CI):

  1. Every `foreign_key.sheet` referenced by a column resolves to a real
     schema file, and the referenced column exists in that sheet's
     `columns` list.
  2. Every `list_ref` resolves to a list defined in schema/19_settings.yaml.
  3. Every `id_generation.counter_key` resolves to a counter defined in
     schema/19_settings.yaml id_counters.
  4. Table-type sheets declare exactly one `primary_key` column
     (view/config/static sheets are exempt).

Exits non-zero with a list of problems if any check fails.
"""
from __future__ import annotations

import sys
from pathlib import Path

import yaml

SCHEMA_DIR = Path(__file__).resolve().parent.parent / "schema"


def load_all_schemas() -> dict[str, dict]:
    schemas = {}
    for path in sorted(SCHEMA_DIR.glob("*.yaml")):
        with path.open() as f:
            data = yaml.safe_load(f)
        if not data or "sheet" not in data:
            continue
        schemas[data["sheet"]] = data
    return schemas


def check_foreign_keys(schemas: dict[str, dict], errors: list[str]) -> None:
    for sheet_name, schema in schemas.items():
        for col in schema.get("columns", []):
            fk = col.get("foreign_key")
            if not fk:
                continue
            target_sheet = fk["sheet"]
            target_col = fk["column"]
            if target_sheet not in schemas:
                errors.append(
                    f"{sheet_name}.{col['name']}: foreign_key references "
                    f"unknown sheet '{target_sheet}'"
                )
                continue
            if "table" in fk:
                # Points at a named sub-table inside a config sheet (e.g.
                # 19_Settings.team_list) rather than a top-level `columns`
                # list — those sub-tables store a flat list of column
                # names, not column dicts.
                target_columns = set(
                    schemas[target_sheet].get(fk["table"], {}).get("columns", [])
                )
            else:
                target_columns = {
                    c["name"] for c in schemas[target_sheet].get("columns", [])
                }
            if target_col not in target_columns:
                errors.append(
                    f"{sheet_name}.{col['name']}: foreign_key references "
                    f"'{target_sheet}.{target_col}', but that column doesn't exist"
                )


def check_list_refs(schemas: dict[str, dict], errors: list[str]) -> None:
    settings = schemas.get("19_Settings", {})
    known_lists = set(settings.get("lists", {}).keys())
    for sheet_name, schema in schemas.items():
        for col in schema.get("columns", []):
            list_ref = col.get("list_ref")
            if list_ref and list_ref not in known_lists:
                errors.append(
                    f"{sheet_name}.{col['name']}: list_ref '{list_ref}' not "
                    f"found in schema/19_settings.yaml lists"
                )


def check_id_counters(schemas: dict[str, dict], errors: list[str]) -> None:
    settings = schemas.get("19_Settings", {})
    known_counters = {
        c["key"] for c in settings.get("id_counters", [])
    }
    for sheet_name, schema in schemas.items():
        for col in schema.get("columns", []):
            gen = col.get("id_generation")
            if gen and gen.get("counter_key") not in known_counters:
                errors.append(
                    f"{sheet_name}.{col['name']}: id_generation.counter_key "
                    f"'{gen.get('counter_key')}' not found in "
                    f"schema/19_settings.yaml id_counters"
                )


def check_primary_keys(schemas: dict[str, dict], errors: list[str]) -> None:
    for sheet_name, schema in schemas.items():
        if schema.get("type") != "table":
            continue
        declared_pk = schema.get("primary_key")
        sub_table = schema.get("parent_sheet") is not None
        if sub_table or declared_pk is None:
            # Sub-tables key off their parent's FK. Row-log tables (e.g.
            # 06_Wedding_Calendar, 09_Shoot_Checklist) intentionally have
            # no single-column PK — they don't declare `primary_key` at
            # the sheet level, which is the opt-out signal.
            continue
        pk_cols = [c for c in schema.get("columns", []) if c.get("primary_key")]
        if len(pk_cols) != 1:
            errors.append(
                f"{sheet_name}: expected exactly one primary_key column, "
                f"found {len(pk_cols)}"
            )
        elif declared_pk and pk_cols[0]["name"] != declared_pk:
            errors.append(
                f"{sheet_name}: top-level primary_key '{declared_pk}' doesn't "
                f"match the column marked primary_key: true "
                f"('{pk_cols[0]['name']}')"
            )


def main() -> int:
    schemas = load_all_schemas()
    errors: list[str] = []

    check_foreign_keys(schemas, errors)
    check_list_refs(schemas, errors)
    check_id_counters(schemas, errors)
    check_primary_keys(schemas, errors)

    if errors:
        print(f"schema validation failed ({len(errors)} problem(s)):\n")
        for e in errors:
            print(f"  - {e}")
        return 1

    print(f"schema validation passed ({len(schemas)} sheets checked)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
