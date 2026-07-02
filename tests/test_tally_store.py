"""Tests for the Neon (Postgres) Tally reference store.

The store talks to a real Neon database, so the CRUD paths are *integration*
tests: they run only when a throwaway test database URL is provided via the
``BOE_TEST_DATABASE_URL`` environment variable (they create/drop their own rows
by unique name and never touch production data). The pure helpers and error
handling are always tested without a database.
"""

from __future__ import annotations

import os
import uuid

import pytest

from boe_converter.tally_store import (
    BuyerRecord,
    SellerRecord,
    StoreError,
    TallyStore,
    _as_lines,
    _dsn,
    _row_to_buyer,
    _row_to_seller,
)


# ---------------------------------------------------------------------------
# DB-free unit tests (helpers, records, error handling)
# ---------------------------------------------------------------------------
def test_record_defaults():
    assert BuyerRecord().address_lines == []
    assert SellerRecord().address_lines == []


def test_as_lines_handles_list_str_and_garbage():
    assert _as_lines(["a", 1]) == ["a", "1"]
    assert _as_lines('["x", "y"]') == ["x", "y"]
    assert _as_lines("not json") == []
    assert _as_lines(None) == []


def test_row_to_buyer_coerces_nulls():
    b = _row_to_buyer(
        {"name": "ACME", "gstin": None, "state": None, "pincode": None, "address_lines": ["A/43"]}
    )
    assert b.name == "ACME"
    assert b.gstin == "" and b.state == "" and b.pincode == ""
    assert b.address_lines == ["A/43"]


def test_row_to_seller_coerces_nulls():
    s = _row_to_seller({"name": "Prayan", "country": None, "address_lines": None})
    assert s.name == "Prayan" and s.country == "" and s.address_lines == []


def test_dsn_prefers_explicit_then_env(monkeypatch):
    monkeypatch.delenv("DATABASE_URL", raising=False)
    monkeypatch.delenv("NEON_DATABASE_URL", raising=False)
    assert _dsn("postgresql://explicit") == "postgresql://explicit"
    monkeypatch.setenv("DATABASE_URL", "postgresql://from-env")
    assert _dsn(None) == "postgresql://from-env"


def test_dsn_missing_raises(monkeypatch):
    monkeypatch.delenv("DATABASE_URL", raising=False)
    monkeypatch.delenv("NEON_DATABASE_URL", raising=False)
    with pytest.raises(StoreError):
        _dsn(None)


def test_constructor_requires_dsn(monkeypatch):
    monkeypatch.delenv("DATABASE_URL", raising=False)
    monkeypatch.delenv("NEON_DATABASE_URL", raising=False)
    with pytest.raises(StoreError):
        TallyStore()


# ---------------------------------------------------------------------------
# Integration tests (real Neon DB) — opt in via BOE_TEST_DATABASE_URL
# ---------------------------------------------------------------------------
_TEST_DSN = os.environ.get("BOE_TEST_DATABASE_URL")
integration = pytest.mark.skipif(
    not _TEST_DSN, reason="set BOE_TEST_DATABASE_URL to run Neon integration tests"
)


@pytest.fixture()
def store():
    s = TallyStore(dsn=_TEST_DSN)
    s.init_schema()
    yield s


@integration
def test_ping_ok(store):
    store.ping()  # must not raise


@integration
def test_stock_item_add_list_delete(store):
    name = f"TEST STOCK {uuid.uuid4().hex[:8]}"
    try:
        assert store.add_stock_item(name)
        assert not store.add_stock_item(name.lower())  # case-insensitive dedup
        assert name in store.list_stock_items()
    finally:
        assert store.delete_stock_item(name)
        assert name not in store.list_stock_items()


@integration
def test_buyer_crud(store):
    name = f"TEST BUYER {uuid.uuid4().hex[:8]}"
    try:
        store.add_buyer(
            BuyerRecord(name=name, gstin="27AAYFG7003K1ZW", state="Maharashtra",
                        pincode="400080", address_lines=["A/43", "Mulund"])
        )
        found = store.find_buyer(name.lower())
        assert found and found.gstin == "27AAYFG7003K1ZW"
        assert found.address_lines == ["A/43", "Mulund"]
        # Upsert on same name (case-insensitive).
        store.add_buyer(BuyerRecord(name=name.lower(), gstin="NEW"))
        assert store.find_buyer(name).gstin == "NEW"
    finally:
        assert store.delete_buyer(name)


@integration
def test_seller_crud(store):
    name = f"TEST SELLER {uuid.uuid4().hex[:8]}"
    try:
        store.add_seller(SellerRecord(name=name, country="China", address_lines=["Room 203"]))
        found = store.find_seller(name)
        assert found and found.country == "China"
    finally:
        assert store.delete_seller(name)
