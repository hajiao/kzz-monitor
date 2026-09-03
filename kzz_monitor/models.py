from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from enum import Enum


class Trend(str, Enum):
    UP = "上升"
    DOWN = "下降"
    FLAT = "震荡"
    UNKNOWN = "未知"


@dataclass(slots=True)
class BondConfig:
    code: str
    enabled: bool = True
    name: str = ""
    sell_trigger_price: float = 130.0
    sell_drawdown_pct: float = 5.0
    trend_window: int = 3
    trend_epsilon: float = 0.10
    build_line: float | None = None
    add_line: float | None = None
    heavy_line: float | None = None
    rating: str = ""


@dataclass(slots=True)
class Quote:
    code: str
    name: str
    price: float
    timestamp: datetime
    stock_code: str = ""
    stock_name: str = ""
    stock_price: float | None = None
    conversion_price: float | None = None
    redeem_status: str = ""


@dataclass(slots=True)
class Evaluation:
    trend: Trend
    one_year_high: float
    monitored_peak: float
    drawdown_pct: float
    zone: str
    sell_alert: bool
    alert_messages: list[str]
