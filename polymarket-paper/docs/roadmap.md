# Roadmap

Phase gates are review gates. Nothing advances without sign-off.

- **Phase 0 — read-only.** `AsyncPublicClient`, market discovery, live book,
  websocket subscription. *Delivered; awaiting review.*
- **Phase 1 — data.** SQLite (WAL), `Thesis` / `Hypothesis`, immutable journal
  rows carrying `hypothesis_id`, `model_version`, `prompt_version`, and empty
  `resolved_outcome` / `brier_score` to be filled in later.
- **Phase 2 — simulator.** Walk the stored book, taker fees, measured latency
  δ, partial fills, exit cost, no look-ahead. **Gate: the coin-flip test must
  lose money at roughly the fee rate.** Zero or positive means the simulator is
  broken and every later number is worthless.
- **Phase 3 — risk.** Fractional Kelly (¼ default), absolute per-position cap
  independent of edge, daily loss limit, kill switch. Pure functions, unit
  tests per rejection branch. No LLM.
- **Phase 4 — research.** The only module where an LLM lives. Emits `Thesis`
  and nothing else: no size, no price, no action. The five hypotheses land here.
- **Phase 5 — run and do not touch.** Shadow loop, parameters frozen at start.
- **Phase 6 — scoreboard.** BSS per hypothesis with sample size and confidence
  intervals, calibration curves, breakdown by category and horizon.

## Directory map

```
ingest/     order books (websocket), market metadata (Gamma), external sources
research/   the only place the LLM lives: evidence -> estimated probability
risk/       deterministic, no LLM: fractional Kelly, caps, kill switch
simulate/   fill engine — the critical piece
journal/    SQLite, immutable log of every decision
evaluate/   calibration against real resolutions
notify/     notify(event) behind an abstraction; markdown + console first
```
