# polypaper

Paper-trading and edge-measurement harness for Polymarket.

**This build cannot trade with real money, by construction.** The authenticated
SDK clients are made unconstructable at startup and the process refuses to run
if a signing key is present in the environment. See `src/polypaper/safety.py`.

## Why paper

Real capital is $25. The platform minimum is 10 pUSD per order. That funds two
positions, which makes running five hypotheses in parallel arithmetically
impossible. Paper trading is not caution here, it is the only way to measure
several hypotheses at once at this size.

## The metric

Per hypothesis, over its resolved decisions:

```
BSS = 1 − (Brier_hypothesis / Brier_market_price)
```

`BSS > 0` beats the market price. Simulated PnL is not a promotion criterion —
on small samples it is noise, and a lucky run on an edgeless strategy is the
most expensive mistake available.

## Status

| Phase | What | State |
|---|---|---|
| 0 | Read-only connectivity, live order book | **delivered, awaiting review** |
| 1 | SQLite schema (WAL), `Thesis` / `Hypothesis`, journal | not started |
| 2 | Fill simulator + coin-flip sanity test | not started |
| 3 | Risk: fractional Kelly, caps, kill switch | not started |
| 4 | Research (the only place an LLM lives) | not started |
| 5 | Shadow run, parameters frozen | not started |
| 6 | Scoreboard: BSS with confidence intervals | not started |

## Phase 0 — running it

Requires Python 3.11+ and outbound network access to Polymarket.

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -e ".[dev]"

# most liquid open market
python scripts\phase0_book.py --discover

# a specific market, streaming live
python scripts\phase0_book.py --slug <market-slug> --follow
```

Offline checks:

```powershell
pytest -q
pyright
```

## Reading the output

The script prints venue parameters (`tick_size`, `min_order_size`, fee
configuration) **as read from the API at runtime**. No such value is hardcoded
anywhere in this repository; that is a standing rule, not a Phase 0 detail.

It also prints a "walking the book" block: what a $10 / $12.50 / $25 taker buy
would actually pay versus the midpoint. That gap, plus taker fees and latency,
is what separates a paper strategy that looks profitable from one that is.

See `docs/api-verification.md` for what was verified against primary sources
and what remains open.
