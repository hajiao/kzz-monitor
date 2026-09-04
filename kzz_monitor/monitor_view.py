from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True, slots=True)
class VisualState:
    key: str
    label: str
    is_alert: bool


def match_bond_rows(query: str, rows: list[dict[str, Any]], limit: int = 20) -> list[dict[str, Any]]:
    """按精确、前缀、包含顺序模糊匹配转债代码或名称。"""
    needle = query.strip().casefold().replace(" ", "")
    if not needle:
        return rows[:limit]
    ranked: list[tuple[int, int, dict[str, Any]]] = []
    for index, row in enumerate(rows):
        code = str(row.get("转债代码") or "").strip().casefold().replace(" ", "")
        name = str(row.get("名称") or "").strip().casefold().replace(" ", "")
        fields = (code, name)
        if needle in fields:
            rank = 0
        elif any(value.startswith(needle) for value in fields):
            rank = 1
        elif any(needle in value for value in fields):
            rank = 2
        else:
            continue
        ranked.append((rank, index, row))
    ranked.sort(key=lambda item: (item[0], item[1]))
    return [row for _, _, row in ranked[:limit]]


def _number(value: Any) -> float | None:
    try:
        if value in (None, "", "-"):
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def classify_bond_row(row: dict[str, Any], sell_latched: bool = False) -> VisualState:
    """按最高优先级返回行状态：无行情 > 卖出 > 重仓 > 加仓 > 建仓 > 强赎。"""
    trend = str(row.get("趋势") or "").strip()
    if trend == "无实时行情":
        return VisualState("unavailable", "无实时行情", True)

    peak = _number(row.get("监控峰值"))
    current = _number(row.get("当前价"))
    trigger = _number(row.get("卖出观察价"))
    drawdown = _number(row.get("回撤%"))
    threshold = _number(row.get("回撤提醒%"))
    derived_sell = bool(
        peak is not None and current is not None and trigger is not None
        and drawdown is not None and threshold is not None
        and peak >= trigger and trend == "下降" and drawdown >= threshold
    )
    if sell_latched or derived_sell:
        return VisualState("sell", "卖出观察", True)

    zone = str(row.get("仓位区域") or "").strip()
    if zone == "重仓":
        return VisualState("heavy", "重仓区", True)
    if zone == "加仓":
        return VisualState("add", "加仓区", True)
    if zone == "建仓":
        return VisualState("build", "建仓区", True)

    redeem = str(row.get("强赎状态") or "").strip()
    harmless = ("", "未开始", "未触发", "不强赎", "暂无")
    if redeem not in harmless:
        return VisualState("redeem", "强赎关注", True)
    return VisualState("normal", "正常观察", False)
