# API verification — findings before building on the brief

Date: 2026-09-14. Method: PyPI metadata, direct introspection of the installed
SDK, and web sources. **`docs.polymarket.com` was unreachable from the build
environment** (HTTP 403 at the egress proxy), so primary-source items below are
marked as such and flagged where they remain open.

## Summary

| # | Brief says | Verdict |
|---|---|---|
| 1 | `py-clob-client` is dead; CLOB V2 landed 2026-04-28, no backward compat | **Confirmed** |
| 2 | The SDK is `polymarket-client` with those four client classes | **Confirmed** |
| 3 | Realtime subscriptions are async-only | **Confirmed** |
| 4 | TWAP window for 5-minute markets is disputed, 30s vs 60s | **Resolved — not a contradiction** |
| 5 | — | **New finding: a second official V2 SDK exists** |

---

## 1. CLOB V2 and `py-clob-client` — confirmed

`py-clob-client` still resolves on PyPI at **0.34.6**, last uploaded
**2026-02-19**, i.e. it stopped before the April cutover. Sources agree V2 went
live 2026-04-28, open orders were wiped, order-struct fields changed, and
collateral moved from USDC.e to pUSD with no backward compatibility. A V1
client is not "degraded", it is rejected.

The brief's instruction stands. `py-clob-client` is not a dependency here and
must not become one.

## 2. `polymarket-client` — confirmed, with all four classes

```
name            polymarket-client
version         0.10.0  (uploaded 2026-09-10, four days ago)
summary         Official Python client for Polymarket
author-email    engineering@polymarket.com
repository      github.com/Polymarket/py-sdk
requires-python >=3.11
```

Introspection of the installed package confirms all four names the brief
specified: `PublicClient`, `SecureClient`, `AsyncPublicClient`,
`AsyncSecureClient`.

## 3. Realtime is async-only — confirmed

`subscribe` exists **only** on `AsyncPublicClient` and `AsyncSecureClient`. The
synchronous classes do not expose it. The brief was right, and it is a
structural fact of the SDK rather than a style preference.

`subscribe()` is a **coroutine function** returning a `SubscriptionHandle`, so
it must be awaited *before* entering the async context manager. Writing the
natural-looking `async with client.subscribe(...)` is a type error and fails at
runtime. Pyright in strict mode caught this; it is fixed in the Phase 0 script.

## 4. TWAP 30s vs 60s — the sources are not in conflict, they are a timeline

This was the item flagged as contradictory. It is not a contradiction, it is
two points in time, and both numbers were correct when published:

1. Crypto up/down markets moved from a single-price snapshot to a Chainlink
   TWAP, after research documented settlement manipulation — the "five-second
   trick". At that rollout (**2026-08-07**) 5-minute markets used a **30-second**
   lookback, while 15-minute and 4-hour markets used **60 seconds**.
2. A later change moved **5-minute markets to a 60-second TWAP**, effective
   00:00 UTC, replacing the 30-second window.

**Current value to build on: 60 seconds for 5-minute markets.**

> **Open — needs your confirmation.** Both search results attribute this to the
> Polymarket changelog, but I could not open `docs.polymarket.com` from here to
> read the primary text and its effective date. Given the brief's rule about not
> building on unverified constants, treat 60s as *strongly indicated, not yet
> primary-source confirmed*. The system should not hardcode it either way: the
> SDK exposes a `CryptoPricesChainlinkTwapSpec` realtime topic, so the TWAP feed
> can be consumed directly rather than reconstructed from an assumed window.

## 5. New finding: there are *two* official V2 Python SDKs

The brief names one. PyPI has two, both from Polymarket:

| Package | Version | Repo | Last upload |
|---|---|---|---|
| `polymarket-client` | 0.10.0 | `Polymarket/py-sdk` | 2026-09-10 |
| `py-clob-client-v2` | 1.1.0 | `Polymarket/py-clob-client-v2` | 2026-07-17 |

`py-clob-client-v2` is authored by "Polymarket Engineering", first published
2026-04-08 — three weeks before the cutover — so it looks like the direct V2
port of the old client, and several of the migration write-ups point at it.

I went with **`polymarket-client`**, for three reasons: it self-describes as the
*official* client, it is the broader workflow-oriented SDK rather than a CLOB-only
port, and it is actively maintained (released four days ago versus two months).
This matches the brief. Flagging it only so the choice is explicit and not an
accident — if you have a reason to prefer the CLOB-only client, now is the
cheap time to switch.

## 6. Values that must be read at runtime — and are

The brief forbids inventing venue constants. Good news: the SDK returns them,
so nothing needs to be hardcoded or scraped.

`OrderBook` carries, per book:

- `tick_size: Decimal`
- `min_order_size: Decimal`
- `neg_risk: bool`
- `last_trade_price: Decimal | None`, `timestamp`, `hash`

`Market.trading` carries:

- `minimum_order_size`, `minimum_tick_size`
- `fees_enabled: bool | None`, `fee_type: str | None`, `fee_schedule`
- `seconds_delay: int | None`

The Phase 0 script prints all of these from the live response. **No venue
constant appears anywhere in this repository.**

> **Open.** The brief states a 10 pUSD order minimum. That is per-market
> `min_order_size` in the API, so it will be read, not assumed — but I could not
> confirm the platform-wide floor against primary docs. The $25-of-capital,
> two-positions reasoning is unaffected unless the true minimum is *higher*.

> **Open — matters for Phase 2.** The **taker fee schedule** is the one number
> the simulator cannot work without, and I could not verify it. `fee_type` and
> `fee_schedule` are per-market API fields, so the simulator will read them per
> market rather than assume a flat rate. The coin-flip test is what will prove
> the fee is actually being charged: if fees are wired up wrong, that test comes
> out at zero instead of negative, which is exactly the failure it exists to catch.

## 7. Rate limits — partially verified

Trading-side limits are tiered per signer address via token buckets (order and
cancel buckets, refilled continuously; tier follows 30-day maker volume,
reassigned roughly every three hours). Those do not apply to Phase 0, which
places no orders.

**Read-side limits I could not confirm**, again because docs were unreachable.
Mitigation: the SDK ships its own retry-with-rate-limit-backoff in its transport
layer (`async_run_with_rate_limit_retry`), which the script inherits for free.

## 8. Environment constraint affecting the demo

This build environment cannot reach `*.polymarket.com` — the egress proxy
answers `403 Forbidden` to `CONNECT` for `clob.`, `gamma-api.`, `data-api.` and
`ws-subscriptions-clob.`. **So I could not run the Phase 0 script against live
data from here.** It gets as far as the HTTP layer and returns a clean
`TransportError: 403 Forbidden`, which tells us the code path is right and only
the network is missing.

It should run correctly on your Windows machine, which has normal egress. That
live run is the actual Phase 0 acceptance test, and it is yours to perform.

## Sources

- PyPI JSON metadata for `polymarket-client`, `py-clob-client`, `py-clob-client-v2`
- Direct introspection of `polymarket-client` 0.10.0
- [Migrating to CLOB V2](https://docs.polymarket.com/v2-migration) (via search; page not directly reachable)
- [Predictions changelog](https://docs.polymarket.com/changelog/predictions) (via search)
- [Chainlink TWAP prices](https://docs.polymarket.com/market-data/chainlink-twap) (via search)
- [Rate limits](https://docs.polymarket.com/api-reference/rate-limits) (via search)
- [How a five-second trick let traders drain millions from Polymarket](https://www.coindesk.com/business/2026/08/07/how-a-five-second-trick-let-traders-drain-millions-from-polymarket) — CoinDesk
- [Chainlink TWAP Data Streams settle Polymarket 5- and 15-minute markets](https://genfinity.io/2026/08/12/chainlink-twap-data-streams-polymarket-crypto-markets/) — Genfinity
