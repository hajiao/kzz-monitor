from __future__ import annotations

import logging
import math
from datetime import date, datetime, timedelta
from typing import Any

from .clock import china_now
from .config import normalize_code
from .models import Quote

logger = logging.getLogger(__name__)


def _find_column(columns: Any, *aliases: str) -> str | None:
    normalized = {str(column).strip(): str(column) for column in columns}
    for alias in aliases:
        if alias in normalized:
            return normalized[alias]
    return None


def _number(value: Any) -> float | None:
    try:
        if value is None or str(value).strip() in {"", "-", "--", "nan"}:
            return None
        result = float(str(value).replace("%", "").replace(",", ""))
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


def exchange_symbol(code: str) -> str:
    code = normalize_code(code)
    if code.startswith(("110", "111", "113", "118")):
        return f"sh{code}"
    if code.startswith(("123", "127", "128")):
        return f"sz{code}"
    raise ValueError(f"无法判断转债 {code} 所属交易所")


class AkShareProvider:
    """所有数据源调用集中在这里，便于以后替换接口或调整频率。"""

    def __init__(self) -> None:
        import akshare as ak

        self.ak = ak
        self._spot_cache: dict[str, Quote] = {}
        self._spot_cache_at: datetime | None = None
        self._trade_dates: set[date] = set()
        self._calendar_loaded_at: date | None = None

    def refresh_spot(self, now: datetime | None = None) -> dict[str, Quote]:
        now = now or china_now()
        errors: list[str] = []
        for method_name in ("bond_zh_hs_cov_spot", "bond_cb_redeem_jsl"):
            method = getattr(self.ak, method_name, None)
            if method is None:
                continue
            try:
                frame = method()
                parsed = self._parse_spot(frame, now)
                if parsed:
                    self._spot_cache = parsed
                    self._spot_cache_at = now
                    return parsed
            except Exception as exc:  # 网络/API 错误需要交给重试循环处理
                errors.append(f"{method_name}: {exc}")
        raise RuntimeError("实时行情接口均不可用；" + "；".join(errors))

    def _parse_spot(self, frame: Any, now: datetime) -> dict[str, Quote]:
        code_col = _find_column(frame.columns, "代码", "债券代码", "转债代码")
        price_col = _find_column(frame.columns, "现价", "最新价", "最新", "成交价")
        if not code_col or not price_col:
            raise ValueError(f"行情字段不兼容: {list(frame.columns)}")
        name_col = _find_column(frame.columns, "名称", "债券名称", "转债名称")
        stock_code_col = _find_column(frame.columns, "正股代码")
        stock_name_col = _find_column(frame.columns, "正股名称")
        stock_price_col = _find_column(frame.columns, "正股价", "正股最新价")
        conversion_col = _find_column(frame.columns, "转股价")
        redeem_col = _find_column(frame.columns, "强赎状态")
        result: dict[str, Quote] = {}
        for _, row in frame.iterrows():
            code = normalize_code(row[code_col])
            price = _number(row[price_col])
            if not code or price is None or price <= 0:
                continue
            quote = Quote(
                code=code,
                name=str(row[name_col]) if name_col and row[name_col] is not None else "",
                price=price,
                timestamp=now,
                stock_code=normalize_code(row[stock_code_col]) if stock_code_col else "",
                stock_name=str(row[stock_name_col]) if stock_name_col else "",
                stock_price=_number(row[stock_price_col]) if stock_price_col else None,
                conversion_price=_number(row[conversion_col]) if conversion_col else None,
                redeem_status=str(row[redeem_col] or "未开始") if redeem_col else "",
            )
            existing = result.get(code)
            if existing is None or quote.timestamp >= existing.timestamp:
                result[code] = quote
            else:
                logger.warning("实时行情包含重复代码 %s，保留较新的记录", code)
        return result

    def one_year_high(self, code: str, today: date | None = None) -> float:
        today = today or china_now().date()
        frame = self.history(code)
        date_col = _find_column(frame.columns, "date", "日期")
        high_col = _find_column(frame.columns, "high", "最高")
        if not date_col or not high_col:
            raise ValueError(f"历史行情字段不兼容: {list(frame.columns)}")
        frame = frame.copy()
        frame[date_col] = frame[date_col].apply(
            lambda value: value.date() if hasattr(value, "date") else date.fromisoformat(str(value)[:10])
        )
        frame.sort_values(date_col, inplace=True)
        dates = frame[date_col]
        cutoff = today - timedelta(days=365)
        mask = dates >= cutoff
        values = frame.loc[mask, high_col].apply(_number).dropna()
        if values.empty:
            raise ValueError(f"{code} 最近一年没有历史行情")
        return float(values.max())

    def history(self, code: str) -> Any:
        return self.ak.bond_zh_hs_cov_daily(symbol=exchange_symbol(code))

    def last_trade(self, code: str) -> tuple[date, float] | None:
        """返回最后交易日和收盘价，用于识别已退出实时行情列表的转债。"""
        frame = self.history(code)
        if frame.empty:
            return None
        date_col = _find_column(frame.columns, "date", "日期")
        close_col = _find_column(frame.columns, "close", "收盘")
        if not date_col or not close_col:
            return None
        frame = frame.copy()
        frame[date_col] = frame[date_col].apply(
            lambda value: value.date() if hasattr(value, "date") else date.fromisoformat(str(value)[:10])
        )
        frame.sort_values(date_col, inplace=True)
        row = frame.iloc[-1]
        value = row[date_col]
        trade_date = value.date() if hasattr(value, "date") else date.fromisoformat(str(value)[:10])
        close = _number(row[close_col])
        return (trade_date, close) if close is not None else None

    def is_trade_date(self, target: date) -> bool:
        if self._calendar_loaded_at != china_now().date() or not self._trade_dates:
            frame = self.ak.tool_trade_date_hist_sina()
            column = _find_column(frame.columns, "trade_date", "交易日期", "日期")
            if not column:
                raise ValueError(f"交易日历字段不兼容: {list(frame.columns)}")
            self._trade_dates = {
                value.date() if hasattr(value, "date") else date.fromisoformat(str(value)[:10])
                for value in frame[column]
            }
            self._calendar_loaded_at = china_now().date()
        return target in self._trade_dates
