#!/usr/bin/env python3
"""Minimal rule engine scaffold for A-share paper trading decisions."""
from dataclasses import dataclass

@dataclass
class MarketSnapshot:
    index_close: float
    ma20: float
    ma20_slope: float
    volatility: float  # normalized 0-1
    breadth: float     # normalized 0-1


def detect_regime(s: MarketSnapshot) -> str:
    if s.index_close > s.ma20 and s.ma20_slope > 0 and s.breadth > 0.55 and s.volatility < 0.6:
        return "trend"
    if s.volatility > 0.75 or s.breadth < 0.4:
        return "risk_off"
    return "range"


def risk_position_multiplier(regime: str) -> float:
    return {
        "trend": 1.0,
        "range": 0.5,
        "risk_off": 0.1,
    }[regime]


if __name__ == "__main__":
    # example
    snap = MarketSnapshot(index_close=3200, ma20=3170, ma20_slope=1.2, volatility=0.45, breadth=0.61)
    r = detect_regime(snap)
    print({"regime": r, "pos_mult": risk_position_multiplier(r)})
