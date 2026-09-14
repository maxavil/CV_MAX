"""Phase 0 deliverable: print a live Polymarket order book. Read-only.

What this proves, and nothing more:

  * ``AsyncPublicClient`` connects and speaks CLOB V2.
  * Markets can be discovered and filtered (Gamma).
  * A real order book, with real depth, can be read and re-read.
  * The realtime websocket subscription delivers book updates.
  * Tick size, minimum order size and fee configuration are read *from the
    API at runtime*, never hardcoded.

It cannot trade. See :mod:`polypaper.safety`.

Usage
-----
    python -m polypaper.scripts_phase0 --slug <market-slug>
    python -m polypaper.scripts_phase0 --discover        # most liquid market
    python -m polypaper.scripts_phase0 --discover --follow
"""

from __future__ import annotations

import argparse
import asyncio
import contextlib
import logging
import signal
from collections.abc import Sequence
from dataclasses import dataclass
from decimal import Decimal
from typing import TYPE_CHECKING, Protocol

from polypaper.logging_setup import configure_logging
from polypaper.safety import engage_read_only_interlock

if TYPE_CHECKING:
    from polymarket import AsyncPublicClient
    from polymarket.models.clob.order_book import OrderBook
    from polymarket.models.gamma.market import Market


class PriceLevel(Protocol):
    """Structural view of one order book level.

    A Protocol rather than the SDK class so the fill arithmetic can be tested
    without fabricating SDK objects, and so Phase 2 can replay stored snapshots
    through the exact same code path.
    """

    @property
    def price(self) -> Decimal: ...

    @property
    def size(self) -> Decimal: ...

LOG = logging.getLogger("polypaper.phase0")

# Depth rows shown per side. The interesting part of a thin market is the top.
_DEPTH_ROWS = 10

# Notional sizes used for the walk-the-book preview, in pUSD. Chosen around the
# real constraint: $25 of capital and a 10 pUSD platform minimum means the only
# order sizes that will ever matter are small ones.
_WALK_NOTIONALS: tuple[Decimal, ...] = (Decimal("10"), Decimal("12.5"), Decimal("25"))


@dataclass(frozen=True)
class WalkResult:
    """Outcome of consuming an order book from the top down.

    This is a preview of the Phase 2 fill model, deliberately kept here so the
    difference between the midpoint and what you would actually pay is visible
    from day one.
    """

    requested_notional: Decimal
    filled_notional: Decimal
    shares: Decimal
    vwap: Decimal | None
    levels_consumed: int
    exhausted_book: bool

    @property
    def fully_filled(self) -> bool:
        return not self.exhausted_book


def walk_book(levels: Sequence[PriceLevel], notional: Decimal) -> WalkResult:
    """Spend ``notional`` against ``levels``, top of book first.

    ``levels`` must already be sorted best-price-first for the side being
    consumed. Returns partial fills honestly rather than pretending the book
    is deeper than it is.
    """
    remaining = notional
    shares = Decimal(0)
    spent = Decimal(0)
    consumed = 0

    for level in levels:
        if remaining <= 0:
            break
        level_capacity = level.price * level.size
        if level_capacity <= remaining:
            shares += level.size
            spent += level_capacity
            remaining -= level_capacity
        else:
            shares += remaining / level.price
            spent += remaining
            remaining = Decimal(0)
        consumed += 1

    vwap = (spent / shares) if shares > 0 else None
    return WalkResult(
        requested_notional=notional,
        filled_notional=spent,
        shares=shares,
        vwap=vwap,
        levels_consumed=consumed,
        exhausted_book=remaining > 0,
    )


def _sorted_sides(book: OrderBook) -> tuple[Sequence[PriceLevel], Sequence[PriceLevel]]:
    """Return (bids desc, asks asc).

    Sort order is asserted locally rather than assumed from the API, so a
    change in server-side ordering cannot silently invert the fill model.
    """
    bids = sorted(book.bids, key=lambda level: level.price, reverse=True)
    asks = sorted(book.asks, key=lambda level: level.price)
    return bids, asks


def _fmt(value: Decimal | None, places: str = "0.0001") -> str:
    return "—" if value is None else str(value.quantize(Decimal(places)))


def render_book(book: OrderBook, market: Market | None) -> str:
    """Format a book snapshot for a terminal."""
    bids, asks = _sorted_sides(book)
    best_bid = bids[0].price if bids else None
    best_ask = asks[0].price if asks else None
    spread = (best_ask - best_bid) if (best_bid is not None and best_ask is not None) else None
    if best_bid is not None and best_ask is not None:
        midpoint: Decimal | None = (best_ask + best_bid) / 2
    else:
        midpoint = None

    out: list[str] = []
    out.append("=" * 78)
    if market is not None and market.question:
        out.append(f"  {market.question}")
    out.append(f"  condition_id : {book.condition_id}")
    out.append(f"  asset_id     : {book.asset_id}")
    out.append(f"  book hash    : {book.hash}")
    out.append(f"  timestamp    : {book.timestamp.isoformat() if book.timestamp else '—'}")
    out.append("=" * 78)

    # Everything below comes from the API response, not from a constant in
    # this repo. Restriction #4 of the brief.
    out.append("  VENUE PARAMETERS (read at runtime, never hardcoded)")
    out.append(f"    tick_size       : {book.tick_size}")
    out.append(f"    min_order_size  : {book.min_order_size}")
    out.append(f"    neg_risk        : {book.neg_risk}")
    out.append(f"    last_trade_price: {_fmt(book.last_trade_price)}")
    if market is not None:
        trading = market.trading
        out.append(f"    fees_enabled    : {trading.fees_enabled}")
        out.append(f"    fee_type        : {trading.fee_type}")
        out.append(f"    fee_schedule    : {trading.fee_schedule}")
        out.append(f"    seconds_delay   : {trading.seconds_delay}")
    out.append("")

    out.append(f"  {'BIDS (buy)':>34}   |   {'ASKS (sell)':<34}")
    out.append(f"  {'price':>12} {'size':>12} {'cum$':>8}   |   "
               f"{'price':<12} {'size':<12} {'cum$':<8}")
    out.append("  " + "-" * 74)

    cum_bid = Decimal(0)
    cum_ask = Decimal(0)
    for row in range(_DEPTH_ROWS):
        bid = bids[row] if row < len(bids) else None
        ask = asks[row] if row < len(asks) else None
        if bid is None and ask is None:
            break
        if bid is not None:
            cum_bid += bid.price * bid.size
            left = f"{_fmt(bid.price):>12} {_fmt(bid.size, '0.01'):>12} {_fmt(cum_bid, '0.01'):>8}"
        else:
            left = " " * 34
        if ask is not None:
            cum_ask += ask.price * ask.size
            right = f"{_fmt(ask.price):<12} {_fmt(ask.size, '0.01'):<12} {_fmt(cum_ask, '0.01'):<8}"
        else:
            right = ""
        out.append(f"  {left}   |   {right}")

    out.append("")
    out.append(f"  best bid {_fmt(best_bid)}   best ask {_fmt(best_ask)}   "
               f"spread {_fmt(spread)}   midpoint {_fmt(midpoint)}")
    out.append("")

    # The number that matters. Paper trading at the midpoint is the single
    # most common way a backtest lies; show the gap explicitly, up front.
    out.append("  WALKING THE BOOK  (what a taker BUY actually costs vs. the midpoint)")
    if not asks:
        out.append("    ask side empty — nothing to buy")
    else:
        for notional in _WALK_NOTIONALS:
            result = walk_book(asks, notional)
            if result.vwap is None:
                out.append(f"    ${notional:>6}  no fill")
                continue
            slip = (result.vwap - midpoint) if midpoint is not None else None
            slip_bps = (
                (slip / midpoint * Decimal(10_000)) if (slip is not None and midpoint) else None
            )
            status = "FULL" if result.fully_filled else "PARTIAL"
            out.append(
                f"    ${notional:>6}  vwap {_fmt(result.vwap)}  "
                f"shares {_fmt(result.shares, '0.01'):>10}  "
                f"filled ${_fmt(result.filled_notional, '0.01')}  "
                f"levels {result.levels_consumed:>2}  "
                f"slip {_fmt(slip_bps, '0.1') if slip_bps is not None else '—':>7} bps  "
                f"[{status}]"
            )
    out.append("=" * 78)
    return "\n".join(out)


async def discover_market(client: AsyncPublicClient) -> Market | None:
    """Pick the most liquid open, order-book-enabled market as a demo target."""
    best: Market | None = None
    best_liquidity = Decimal(0)
    seen = 0

    paginator = client.list_markets(closed=False, order="liquidity", ascending=False, page_size=50)
    # iter_items() flattens pages. Iterating the paginator directly yields
    # Page[Market], which silently type-errors into nonsense at runtime.
    async for market in paginator.iter_items():
        seen += 1
        state = market.state
        if not (state.active and state.enable_order_book and state.accepting_orders):
            continue
        metrics = market.metrics
        liquidity = metrics.liquidity_num or metrics.liquidity or Decimal(0)
        if liquidity > best_liquidity:
            best, best_liquidity = market, liquidity
        if seen >= 200:
            break

    LOG.info("scanned %d markets; picked liquidity=%s", seen, best_liquidity)
    return best


def _first_token_id(market: Market) -> str | None:
    for outcome in (market.outcomes.yes, market.outcomes.no):
        if outcome.token_id:
            return str(outcome.token_id)
    return None


async def follow_book(client: AsyncPublicClient, token_id: str, stop: asyncio.Event) -> None:
    """Stream live book updates over the websocket until ``stop`` is set."""
    from polymarket.streams import MarketSpec

    LOG.info("subscribing to realtime updates for asset_id=%s", token_id)

    subscription = await client.subscribe(MarketSpec(asset_ids=[token_id]))
    async with subscription:
        async for event in subscription:
            if stop.is_set():
                break
            event_type = getattr(event, "type", type(event).__name__)
            payload = getattr(event, "payload", None)
            timestamp = getattr(payload, "timestamp", None)
            stamp = timestamp.isoformat() if timestamp is not None else "—"

            if event_type == "book":
                bids = getattr(payload, "bids", ())
                asks = getattr(payload, "asks", ())
                print(f"[{stamp}] BOOK   depth {len(bids)}x{len(asks)}")
            elif event_type == "price_change":
                changes = getattr(payload, "price_changes", ())
                print(f"[{stamp}] PRICE  {len(changes)} level change(s)")
            elif event_type == "best_bid_ask":
                bid = getattr(payload, "best_bid", None)
                ask = getattr(payload, "best_ask", None)
                print(f"[{stamp}] BBO    bid {_fmt(bid)}  ask {_fmt(ask)}")
            elif event_type == "tick_size_change":
                print(f"[{stamp}] TICK   {payload!r}")
            else:
                print(f"[{stamp}] {event_type.upper()}")


async def run(args: argparse.Namespace) -> int:
    # Before importing or constructing anything that touches the network.
    engage_read_only_interlock()

    from polymarket import AsyncPublicClient

    client = AsyncPublicClient()
    try:
        market: Market | None = None
        if args.slug:
            market = await client.get_market(slug=args.slug)
        elif args.discover:
            market = await discover_market(client)
            if market is None:
                LOG.error("discovery found no tradable market")
                return 2

        if args.token_id:
            token_id = args.token_id
        else:
            assert market is not None
            found = _first_token_id(market)
            if found is None:
                LOG.error("market %r exposes no CLOB token id", args.slug or "discovered")
                return 2
            token_id = found

        book = await client.get_order_book(token_id=token_id)
        print(render_book(book, market))

        if args.follow:
            stop = asyncio.Event()
            loop = asyncio.get_running_loop()
            for sig in (signal.SIGINT, signal.SIGTERM):
                with contextlib.suppress(NotImplementedError):
                    loop.add_signal_handler(sig, stop.set)
            print("\nstreaming live updates — Ctrl+C to stop\n")
            with contextlib.suppress(asyncio.CancelledError, KeyboardInterrupt):
                await follow_book(client, token_id, stop)
        return 0
    finally:
        await client.close()


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="polypaper-book",
        description="Phase 0: print a live Polymarket order book. Read-only; cannot trade.",
    )
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--slug", help="market slug, e.g. from the market's URL")
    source.add_argument("--discover", action="store_true", help="pick the most liquid open market")
    source.add_argument("--token-id", help="CLOB token id (asset id) to read directly")
    parser.add_argument("--follow", action="store_true", help="stream live websocket updates")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    configure_logging(logging.DEBUG if args.verbose else logging.INFO)

    from polymarket.errors import TransportError

    try:
        return asyncio.run(run(args))
    except KeyboardInterrupt:
        return 130
    except TransportError as error:
        # Distinguish "the venue said no" from "this machine has no route",
        # because on a locked-down network these look identical in a traceback.
        LOG.error("transport error talking to Polymarket: %s", error)
        LOG.error(
            "If this is 403/000 for every endpoint, check egress: this host may "
            "not be allowed to reach *.polymarket.com."
        )
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
