"""Phase 0 tests: the interlock, the redactor, and the fill-walk arithmetic.

These run offline. Anything needing the network belongs in a Phase 0 smoke
test that is run by hand against production.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from decimal import Decimal

import pytest

from polypaper.logging_setup import RedactingFilter
from polypaper.safety import LiveTradingDisabled, engage_read_only_interlock
from polypaper.scripts_phase0 import walk_book


@dataclass(frozen=True)
class FakeLevel:
    """Structural stand-in for polymarket OrderBookLevel."""

    price: Decimal
    size: Decimal


# --------------------------------------------------------------------------
# Interlock
# --------------------------------------------------------------------------


def test_interlock_makes_secure_clients_unconstructable() -> None:
    import polymarket

    engage_read_only_interlock()

    with pytest.raises(LiveTradingDisabled):
        polymarket.SecureClient()  # pyright: ignore[reportCallIssue]
    with pytest.raises(LiveTradingDisabled):
        polymarket.AsyncSecureClient()  # pyright: ignore[reportCallIssue]


def test_interlock_refuses_to_start_with_a_signing_key_present(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("POLYMARKET_PRIVATE_KEY", "0x" + "a" * 64)
    with pytest.raises(LiveTradingDisabled, match="signing credentials"):
        engage_read_only_interlock()


def test_public_client_still_works_after_interlock() -> None:
    import polymarket

    engage_read_only_interlock()
    # Construction must not raise; read-only work is the point.
    client = polymarket.AsyncPublicClient()
    assert client is not None


# --------------------------------------------------------------------------
# Redaction
# --------------------------------------------------------------------------


def _apply(msg: str, *args: object) -> str:
    record = logging.LogRecord("t", logging.INFO, __file__, 1, msg, args or None, None)
    RedactingFilter().filter(record)
    return record.getMessage()


def test_redacts_hex_private_key() -> None:
    assert "deadbeef" not in _apply("key is 0x" + "deadbeef" * 8)


def test_redacts_bare_64_hex() -> None:
    assert "[REDACTED]" in _apply("abc " + "f" * 64)


def test_redacts_pem_block() -> None:
    pem = "-----BEGIN EC PRIVATE KEY-----\nMHcCAQE\n-----END EC PRIVATE KEY-----"
    assert "MHcCAQE" not in _apply(pem)


def test_redacts_mnemonic() -> None:
    words = " ".join(["abandon"] * 12)
    assert "[REDACTED]" in _apply(f"mnemonic {words}")


def test_redacts_key_value_but_keeps_key_name() -> None:
    out = _apply("passphrase=hunter2 rest")
    assert "hunter2" not in out
    assert "passphrase" in out


def test_redacts_inside_lazy_args() -> None:
    assert "hunter2" not in _apply("creds %s", "secret=hunter2")


def test_leaves_innocent_text_alone() -> None:
    assert _apply("best bid 0.61 size 120") == "best bid 0.61 size 120"


# --------------------------------------------------------------------------
# Walking the book — the arithmetic the simulator will depend on
# --------------------------------------------------------------------------


def test_walk_consumes_best_price_first() -> None:
    asks = [
        FakeLevel(Decimal("0.50"), Decimal("10")),  # $5 capacity
        FakeLevel(Decimal("0.60"), Decimal("10")),  # $6 capacity
    ]
    result = walk_book(asks, Decimal("5"))
    assert result.vwap == Decimal("0.50")
    assert result.shares == Decimal("10")
    assert result.levels_consumed == 1
    assert result.fully_filled


def test_walk_vwap_is_worse_than_top_of_book_when_crossing_levels() -> None:
    asks = [
        FakeLevel(Decimal("0.50"), Decimal("10")),  # $5
        FakeLevel(Decimal("0.60"), Decimal("10")),  # $6
    ]
    result = walk_book(asks, Decimal("11"))
    assert result.vwap is not None
    assert result.shares == Decimal("20")
    assert result.vwap == Decimal("0.55")
    # This is the whole point: the midpoint would have lied.
    assert result.vwap > Decimal("0.50")
    assert result.levels_consumed == 2


def test_walk_reports_partial_fill_rather_than_inventing_depth() -> None:
    asks = [FakeLevel(Decimal("0.50"), Decimal("10"))]  # only $5 available
    result = walk_book(asks, Decimal("25"))
    assert result.exhausted_book
    assert not result.fully_filled
    assert result.filled_notional == Decimal("5")
    assert result.shares == Decimal("10")


def test_walk_on_empty_book_fills_nothing() -> None:
    result = walk_book([], Decimal("10"))
    assert result.vwap is None
    assert result.shares == Decimal(0)
    assert result.exhausted_book


def test_walk_never_reports_more_notional_than_requested() -> None:
    asks = [FakeLevel(Decimal("0.01"), Decimal("100000"))]
    result = walk_book(asks, Decimal("10"))
    assert result.filled_notional == Decimal("10")
