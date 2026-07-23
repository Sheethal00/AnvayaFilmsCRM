"""Every FK column in schema/*.yaml must resolve to a real PK in another
schema file (catches typos like `BookngID`) — CLAUDE.md Testing Philosophy.

Reuses the same check functions scripts/validate_schema.py runs at the
command line, so there's exactly one implementation of each rule.
"""
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))

import validate_schema  # noqa: E402


@pytest.fixture(scope="module")
def schemas():
    return validate_schema.load_all_schemas()


def test_schema_files_exist(schemas):
    assert len(schemas) == 22, (
        f"expected 22 schema files (20 sheets + 2 sub-tables per SPEC.md §3), "
        f"found {len(schemas)}"
    )


def test_foreign_keys_resolve(schemas):
    errors: list[str] = []
    validate_schema.check_foreign_keys(schemas, errors)
    assert not errors, "\n".join(errors)


def test_list_refs_resolve(schemas):
    errors: list[str] = []
    validate_schema.check_list_refs(schemas, errors)
    assert not errors, "\n".join(errors)


def test_id_counters_resolve(schemas):
    errors: list[str] = []
    validate_schema.check_id_counters(schemas, errors)
    assert not errors, "\n".join(errors)


def test_primary_keys_well_formed(schemas):
    errors: list[str] = []
    validate_schema.check_primary_keys(schemas, errors)
    assert not errors, "\n".join(errors)


def test_every_table_sheet_has_columns(schemas):
    for name, schema in schemas.items():
        if schema.get("type") == "table":
            assert schema.get("columns"), f"{name}: table sheet has no columns"
