"""Runs the ID generator logic in isolation — a Python port of
vba/modules/IDGenerator.bas's GenerateID, exercised against known inputs
per CLAUDE.md Testing Philosophy ("extracted VBA logic reimplemented/
tested in Python"). Not used by the build pipeline; real ID generation
happens in Excel via the VBA macro. Keep this port in sync with the .bas
file by hand when the macro logic changes.
"""
from __future__ import annotations


def format_id(pattern: str, num: int, year: int) -> str:
    hash_count = pattern.count("#")
    padded = str(num).zfill(hash_count)
    result = pattern.replace("YYYY", str(year))
    result = result.replace("#" * hash_count, padded)
    return result


def generate_id(pattern: str, last_value: str | None, current_year: int) -> tuple[str, str]:
    """Mirrors IDGenerator.bas GenerateID. Returns (new_id, new_counter_value)."""
    resets_yearly = "YYYY" in pattern

    if not last_value:
        last_num, last_year = 0, 0
    elif ":" in last_value:
        year_str, num_str = last_value.split(":")
        last_year, last_num = int(year_str), int(num_str)
    else:
        last_num, last_year = int(last_value), 0

    if resets_yearly and last_year != current_year:
        next_num = 1
    else:
        next_num = last_num + 1

    new_id = format_id(pattern, next_num, current_year)
    new_counter = f"{current_year}:{next_num}" if resets_yearly else str(next_num)
    return new_id, new_counter


def test_simple_sequential_pattern():
    new_id, counter = generate_id("LD-####", None, 2026)
    assert new_id == "LD-0001"
    assert counter == "1"

    new_id, counter = generate_id("LD-####", counter, 2026)
    assert new_id == "LD-0002"
    assert counter == "2"


def test_zero_padding_matches_hash_count():
    new_id, _ = generate_id("BK-####", "999", 2026)
    assert new_id == "BK-1000"


def test_yearly_reset_pattern_first_use():
    new_id, counter = generate_id("RCT-YYYY-####", None, 2026)
    assert new_id == "RCT-2026-0001"
    assert counter == "2026:1"


def test_yearly_reset_continues_within_same_year():
    new_id, counter = generate_id("RCT-YYYY-####", "2026:1", 2026)
    assert new_id == "RCT-2026-0002"
    assert counter == "2026:2"


def test_yearly_reset_resets_on_new_year():
    new_id, counter = generate_id("RCT-YYYY-####", "2026:47", 2027)
    assert new_id == "RCT-2027-0001"
    assert counter == "2027:1"


def test_all_settings_patterns_produce_valid_ids():
    import yaml
    from pathlib import Path

    settings_path = Path(__file__).resolve().parent.parent / "schema" / "19_settings.yaml"
    settings = yaml.safe_load(settings_path.read_text())

    for counter in settings["id_counters"]:
        new_id, _ = generate_id(counter["pattern"], None, 2026)
        # every generated ID should be non-empty and contain no leftover '#'
        assert new_id and "#" not in new_id, counter["key"]
