---
name: a-share-paper-quant
description: Run A-share paper-trading decision workflow. Use when user wants simulated quant trading with automatic regime selection (trend/range/risk-off), rule-based entries/exits, strict risk limits, and executable daily trade plans.
---

# A-Share Paper Quant

## Overview
Generate and execute (simulated) A股交易计划 with hard risk controls.

## Rules
- Simulation only unless user explicitly re-confirms real-money switch.
- Hard limits:
  - Single-trade risk <= 1%
  - Daily max drawdown <= 2% then stop
  - 3 consecutive losses => pause one day
- If market regime is unclear, default to reduced exposure.

## Regime Selection
Classify each day into one regime:
1. Trend: index above MA20 and MA20 slope positive.
2. Range: index near MA20, low trend strength.
3. Risk-off: broad weakness / high volatility / gap-down sentiment.

## Execution by Regime
- Trend: allow swing/overnight setups with breakout + volume confirmation.
- Range: mean-reversion or no-trade if edge is weak.
- Risk-off: mostly cash, only A+ setups or fully stand aside.

## Daily Workflow
1. Pre-market: generate watchlist, regime, invalidation levels.
2. Intraday: only trigger on rule match, no discretionary chasing.
3. Post-close: compute PnL, win-rate, payoff ratio, drawdown, mistakes.
4. Write outputs under `quant/`.

## Output Files
- `quant/daily-plan-YYYY-MM-DD.md`
- `quant/trades-YYYY-MM-DD.csv`
- `quant/review-YYYY-MM-DD.md`

## Resources
Use `scripts/regime_engine.py` for rule-based regime and signal scaffolding.
