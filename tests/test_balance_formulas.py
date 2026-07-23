"""Checks the formula strings recorded in schema/*.yaml match the expected
formula for known columns (e.g. BalanceAmount references the right range)
— CLAUDE.md Testing Philosophy. schema/*.yaml is the intermediate
representation build_workbook.py reads; this guards against a formula
column being silently redefined out of sync with SPEC.md.
"""
from pathlib import Path

import yaml

SCHEMA_DIR = Path(__file__).resolve().parent.parent / "schema"


def load(sheet_file: str) -> dict:
    return yaml.safe_load((SCHEMA_DIR / sheet_file).read_text())


def column(schema: dict, name: str) -> dict:
    for col in schema["columns"]:
        if col["name"] == name:
            return col
    raise AssertionError(f"{schema['sheet']}: no column named {name!r}")


def test_balance_amount_references_package_value_and_payments():
    schema = load("05_bookings.yaml")
    col = column(schema, "BalanceAmount")
    assert col["type"] == "formula"
    assert "PackageValue" in col["formula"]
    assert "12_Payments" in col["formula"]
    assert "Cleared" in col["formula"], "must exclude Pending/Bounced payments from the balance"


def test_delivery_deadline_references_wedding_date_and_settings():
    schema = load("05_bookings.yaml")
    col = column(schema, "DeliveryDeadline")
    assert col["type"] == "formula"
    assert "WeddingDate" in col["formula"]
    assert "DefaultDeliveryDays" in col["formula"]


def test_gst_amount_references_settings_rate_not_hardcoded():
    schema = load("03_quotation_manager.yaml")
    col = column(schema, "GSTAmount")
    assert col["type"] == "formula"
    assert "0.18" not in col["formula"], "GST rate must come from Settings, never hardcoded"
    assert "Subtotal" in col["formula"]


def test_quote_total_is_subtotal_plus_gst():
    schema = load("03_quotation_manager.yaml")
    col = column(schema, "Total")
    assert col["formula"] == "Subtotal + GSTAmount"


def test_line_total_is_qty_times_unit_price():
    schema = load("03a_quotation_lineitems.yaml")
    col = column(schema, "LineTotal")
    assert col["formula"] == "Qty * UnitPrice"


def test_countdown_days_is_event_date_minus_today():
    schema = load("06_wedding_calendar.yaml")
    col = column(schema, "CountdownDays")
    assert col["formula"] == "EventDate - TODAY()"


def test_client_lifetime_value_sums_bookings():
    schema = load("04_client_master.yaml")
    col = column(schema, "TotalLifetimeValue")
    assert "05_Bookings" in col["formula"]
    assert "PackageValue" in col["formula"]
