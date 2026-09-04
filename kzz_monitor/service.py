from __future__ import annotations

import logging
import json
import signal
import threading
from time import monotonic
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Callable

from .clock import china_now
from .config import AppSettings, BondConfig, load_configuration
from .excel_store import ExcelStore
from .models import Evaluation, Quote, Trend
from .notifier import Notifier
from .provider import AkShareProvider
from .storage import StateStore

logger = logging.getLogger(__name__)


class AlertEngine:
    def __init__(self, state: StateStore) -> None:
        self.state = state

    def evaluate(self, config: BondConfig, quote: Quote, historical_high: float) -> Evaluation:
        self.state.add_price(config.code, quote.timestamp, quote.price)
        prices = self.state.recent_prices(config.code, config.trend_window)
        trend = self._trend(prices, config.trend_epsilon)
        peak = max(self.state.monitored_peak(config.code), quote.price)
        one_year_high = max(historical_high, quote.price)
        drawdown = max(0.0, (peak - quote.price) / peak * 100) if peak else 0.0
        zone = self._zone(config, quote.price)
        messages: list[str] = []

        zone_key = f"zone:{config.code}"
        previous_zone = self.state.get_value(zone_key, "观察")
        if zone != previous_zone and self._zone_rank(zone) > self._zone_rank(previous_zone):
            messages.append(f"进入{zone}：现价 {quote.price:.2f}")
        self.state.set_value(zone_key, zone)

        latch_key = f"sell_latched:{config.code}"
        latched = self.state.get_value(latch_key, "0") == "1"
        armed = peak >= config.sell_trigger_price
        sell_alert = armed and trend == Trend.DOWN and drawdown >= config.sell_drawdown_pct
        if latched and drawdown <= config.sell_drawdown_pct * 0.4:
            latched = False
            self.state.set_value(latch_key, "0")
        if sell_alert and not latched:
            messages.append(
                f"卖出观察：峰值 {peak:.2f} 已超过观察价 {config.sell_trigger_price:.2f}，"
                f"当前 {quote.price:.2f}，趋势下降，回撤 {drawdown:.2f}%"
            )
            self.state.set_value(latch_key, "1")
        return Evaluation(trend, one_year_high, peak, drawdown, zone, sell_alert, messages)

    @staticmethod
    def _trend(prices: list[float], epsilon: float) -> Trend:
        if len(prices) < 2:
            return Trend.UNKNOWN
        change = prices[-1] - prices[0]
        deltas = [right - left for left, right in zip(prices, prices[1:])]
        negative = sum(delta < 0 for delta in deltas)
        positive = sum(delta > 0 for delta in deltas)
        if change <= -abs(epsilon) and negative >= positive:
            return Trend.DOWN
        if change >= abs(epsilon) and positive >= negative:
            return Trend.UP
        return Trend.FLAT

    @staticmethod
    def _zone(config: BondConfig, price: float) -> str:
        if config.heavy_line and price <= config.heavy_line:
            return "重仓"
        if config.add_line and price <= config.add_line:
            return "加仓"
        if config.build_line and price <= config.build_line:
            return "建仓"
        return "观察"

    @staticmethod
    def _zone_rank(zone: str) -> int:
        return {"观察": 0, "建仓": 1, "加仓": 2, "重仓": 3}.get(zone, 0)


class MonitorService:
    def __init__(
        self,
        workbook: Path,
        state_path: Path,
        provider: AkShareProvider | None = None,
        now: Callable[[], datetime] = china_now,
        status_callback: Callable[[str, str], None] | None = None,
    ) -> None:
        self.workbook = workbook
        self.state_path = state_path
        self.excel = ExcelStore(workbook)
        self.state = StateStore(state_path)
        self.engine = AlertEngine(self.state)
        self.provider = provider or AkShareProvider()
        self.now = now
        self.status_callback = status_callback
        self.stop_event = threading.Event()
        self.wake_event = threading.Event()
        self.force_cycle_event = threading.Event()
        self.refresh_adq_event = threading.Event()

    def _status(self, state: str, message: str) -> None:
        if self.status_callback:
            self.status_callback(state, message)

    def stop(self, *_: object) -> None:
        self.stop_event.set()
        self.wake_event.set()

    def request_force_cycle(self) -> None:
        self.force_cycle_event.set()
        self.wake_event.set()

    def request_adq_refresh(self) -> None:
        self.refresh_adq_event.set()
        self.wake_event.set()

    def _wait(self, seconds: float) -> bool:
        """等待或被控制台操作唤醒；返回是否提前唤醒。"""
        interrupted = self.wake_event.wait(seconds)
        self.wake_event.clear()
        return interrupted

    def run(self, install_signal_handlers: bool = True) -> None:
        if install_signal_handlers and threading.current_thread() is threading.main_thread():
            signal.signal(signal.SIGINT, self.stop)
            signal.signal(signal.SIGTERM, self.stop)
        logger.info("可转债监控服务已启动，工作簿：%s", self.workbook)
        self._status("checking", "正在检查交易日与市场状态")
        while not self.stop_event.is_set():
            try:
                settings, bonds = self._prepare_configuration()
                if self.refresh_adq_event.is_set():
                    self.refresh_adq_event.clear()
                    self.wake_event.clear()
                    if not settings.adq_file:
                        raise ValueError("尚未设置安道全文件")
                    count = self.excel.import_adq(settings.adq_file, settings.adq_sheet)
                    self.state.set_value("adq_import_date", self.now().date().isoformat())
                    logger.info("手动刷新安道全完成，共更新 %d 只转债", count)
                    self._status("checking", f"安道全刷新完成：{count} 只")
                    settings, bonds = load_configuration(self.workbook)
                if self.force_cycle_event.is_set():
                    self.force_cycle_event.clear()
                    self.wake_event.clear()
                    self._status("running", "正在执行强制完整轮询")
                    logger.info("收到控制台指令：立即开始新一轮完整轮询")
                    cycle_started = monotonic()
                    if self._run_complete_cycle(settings, bonds):
                        self._mark_cycle_completed()
                    self._wait_for_next_cycle(settings, cycle_started)
                    continue
                if not bonds:
                    logger.warning("监控列表为空，60 秒后重试")
                    self._wait(60)
                    continue
                today = self.now().date()
                if not self.provider.is_trade_date(today):
                    logger.info("%s 不是交易日，等待下一天", today)
                    self._status("waiting", "今天不是交易日，等待下一交易日")
                    self._wait_until_next_day()
                    continue
                phase = self._market_phase(self.now().time(), settings)
                if self._needs_lunch_catchup(phase, today):
                    logger.info("今天尚未完成过轮询；午间休市期间先补跑一轮")
                    self._status("running", "今日尚未轮询，正在午休期间补跑一轮")
                    if self._run_complete_cycle(settings, bonds):
                        self._mark_cycle_completed()
                    continue
                if phase == "open":
                    self._status("running", "交易时段，监控正常")
                    cycle_started = monotonic()
                    if self._run_open_cycle(settings, bonds):
                        self._mark_cycle_completed()
                    self._wait_for_next_cycle(settings, cycle_started)
                elif phase == "closed" and settings.final_cycle_after_close:
                    key = f"final_cycle:{today.isoformat()}"
                    if self.state.get_value(key) != "done":
                        logger.info("收盘后执行最后一轮完整更新（支持断点续跑）")
                        if self._run_resumable_final_cycle(settings, bonds, today):
                            self._mark_cycle_completed()
                            self.state.set_value(key, "done")
                    else:
                        self._status("waiting", "已收盘并完成最终更新，等待下一交易日")
                        self._wait_until_next_day()
                else:
                    self._status("waiting", self._phase_wait_message(phase, settings))
                    self._wait(30)
            except Exception:
                logger.exception("主循环异常，60 秒后重试")
                self._status("error", "数据源异常，60 秒后重试；请查看日志")
                self._wait(60)
        self.state.close()
        self._status("stopped", "监控已停止")
        logger.info("可转债监控服务已停止")

    def run_once(self) -> None:
        settings, bonds = self._prepare_configuration()
        if self._run_complete_cycle(settings, bonds, wait_between=False):
            self._mark_cycle_completed()
        self.state.close()

    def _prepare_configuration(self) -> tuple[AppSettings, list[BondConfig]]:
        settings, _ = load_configuration(self.workbook)
        today = self.now().date().isoformat()
        if settings.adq_file and self.state.get_value("adq_import_date") != today:
            count = self.excel.import_adq(settings.adq_file, settings.adq_sheet)
            self.state.set_value("adq_import_date", today)
            logger.info("已从安道全文件更新 %d 只转债的评级和三段线", count)
        return load_configuration(self.workbook)

    def _run_open_cycle(self, settings: AppSettings, bonds: list[BondConfig]) -> bool:
        quotes = self.provider.refresh_spot(self.now())
        for index, config in enumerate(bonds):
            if self.stop_event.is_set() or self._market_phase(self.now().time(), settings) != "open":
                return False
            self._process(config, settings, quotes)
            if index < len(bonds) - 1 and self._wait(settings.poll_interval_seconds):
                return False
        return True

    def _run_complete_cycle(self, settings: AppSettings, bonds: list[BondConfig], wait_between: bool = True) -> bool:
        quotes = self.provider.refresh_spot(self.now())
        for index, config in enumerate(bonds):
            if self.stop_event.is_set():
                return False
            self._process(config, settings, quotes)
            if wait_between and index < len(bonds) - 1:
                if self._wait(settings.poll_interval_seconds):
                    return False
        return True

    def _run_resumable_final_cycle(
        self, settings: AppSettings, bonds: list[BondConfig], trade_date: date
    ) -> bool:
        progress_key = f"final_cycle_progress:{trade_date.isoformat()}"
        raw_progress = self.state.get_value(progress_key, "[]")
        try:
            completed = {str(code) for code in json.loads(raw_progress)}
        except (TypeError, ValueError, json.JSONDecodeError):
            logger.warning("收盘轮询断点记录损坏，已从空进度恢复")
            completed = set()
        current_codes = {bond.code for bond in bonds}
        completed.intersection_update(current_codes)
        if completed:
            logger.info("恢复收盘轮询断点：已完成 %d/%d，只处理剩余转债", len(completed), len(bonds))
            self._status("running", f"恢复收盘轮询：已完成 {len(completed)}/{len(bonds)}")
        remaining = [bond for bond in bonds if bond.code not in completed]
        if not remaining:
            self.state.delete_value(progress_key)
            return True
        quotes = self.provider.refresh_spot(self.now())
        for index, config in enumerate(remaining):
            if self.stop_event.is_set():
                return False
            self._process(config, settings, quotes)
            completed.add(config.code)
            self.state.set_value(progress_key, json.dumps(sorted(completed), ensure_ascii=False))
            logger.info("收盘轮询进度：%d/%d（刚完成 %s）", len(completed), len(bonds), config.code)
            self._status("running", f"收盘轮询进度 {len(completed)}/{len(bonds)}")
            if index < len(remaining) - 1 and self._wait(settings.poll_interval_seconds):
                return False
        self.state.delete_value(progress_key)
        return True

    def _mark_cycle_completed(self) -> None:
        completed_at = self.now().isoformat(timespec="seconds")
        self.state.set_value("last_cycle_completed_at", completed_at)
        logger.info("完整轮询完成：%s", completed_at)

    def _last_cycle_is_today(self, today: date) -> bool:
        value = self.state.get_value("last_cycle_completed_at")
        if not value:
            return False
        try:
            return datetime.fromisoformat(value).date() == today
        except ValueError:
            return False

    def _needs_lunch_catchup(self, phase: str, today: date) -> bool:
        return phase == "lunch" and not self._last_cycle_is_today(today)

    @staticmethod
    def _cycle_delay_seconds(cycle_interval_minutes: int, elapsed_seconds: float) -> float:
        """按两轮开始时间计算等待；本轮超时则返回 0。"""
        return max(0.0, cycle_interval_minutes * 60 - elapsed_seconds)

    def _wait_for_next_cycle(self, settings: AppSettings, cycle_started: float) -> None:
        if self.stop_event.is_set() or self.force_cycle_event.is_set() or self.refresh_adq_event.is_set():
            return
        remaining = self._cycle_delay_seconds(
            settings.cycle_interval_minutes, monotonic() - cycle_started
        )
        if remaining <= 0:
            logger.info("本轮耗时已达到整轮间隔，不额外等待")
            return
        logger.info(
            "本轮完成；整轮间隔 %d 分钟，约 %.1f 分钟后可开始下一轮",
            settings.cycle_interval_minutes,
            remaining / 60,
        )
        self._status("waiting", f"本轮已完成，约 {remaining / 60:.1f} 分钟后开始下一轮")
        deadline = monotonic() + remaining
        while not self.stop_event.is_set():
            if self.force_cycle_event.is_set() or self.refresh_adq_event.is_set():
                return
            if self._market_phase(self.now().time(), settings) != "open":
                return
            remaining = deadline - monotonic()
            if remaining <= 0:
                return
            if self._wait(min(30, remaining)):
                return

    def _process(
        self,
        config: BondConfig,
        settings: AppSettings,
        quotes: dict[str, Quote] | None = None,
    ) -> None:
        try:
            self._flush_pending_excel_writes()
            if quotes is None:
                quotes = self.provider.refresh_spot(self.now())
            quote = quotes.get(config.code)
            if quote is None:
                cache_key = f"unavailable:{config.code}"
                explanation = self.state.get_value(cache_key)
                if not explanation:
                    last_trade = self.provider.last_trade(config.code)
                    if last_trade:
                        trade_date, close = last_trade
                        explanation = f"实时行情未返回；最后交易日 {trade_date}，收盘价 {close:.3f}，可能已退市或停止交易"
                    else:
                        explanation = "实时行情未返回，且未查到历史行情；请核对转债代码或稍后重试"
                    self.state.set_value(cache_key, explanation)
                logger.warning("%s | %s", config.code, explanation)
                try:
                    self.excel.mark_unavailable(config.code, self.now(), explanation)
                except PermissionError:
                    logger.warning("Excel 正在打开，暂时无法写入无行情状态")
                return
            high = self.state.get_daily_high(config.code, quote.timestamp.date())
            if high is None:
                high = self.provider.one_year_high(config.code, quote.timestamp.date())
                self.state.set_daily_high(config.code, quote.timestamp.date(), high)
            evaluation = self.engine.evaluate(config, quote, high)
            try:
                self.excel.update_result(quote, evaluation)
            except PermissionError:
                self._queue_result_write(quote, evaluation)
                logger.warning(
                    "Excel 面板暂时无法写入；已存入待写队列，关闭 Excel 后会自动补写"
                )
            notifier = Notifier(settings, self.state_path.parent / "smtp_secret.bin")
            for message in evaluation.alert_messages:
                alert_type = "卖出" if message.startswith("卖出") else evaluation.zone
                if self.state.claim_alert(f"{config.code}:{alert_type}:{quote.timestamp.isoformat(timespec='minutes')}", quote.timestamp):
                    notifier.send(f"{quote.name or config.name or config.code} {alert_type}提醒", message)
                    try:
                        self.excel.append_alert(quote, alert_type, message)
                    except PermissionError:
                        self._queue_alert_write(quote, alert_type, message)
                        logger.warning("提醒已发送；Excel 提醒记录已存入待写队列")
            logger.info("%s %s：%.2f，%s，回撤 %.2f%%", config.code, quote.name, quote.price, evaluation.trend, evaluation.drawdown_pct)
        except Exception as exc:
            logger.exception("更新 %s 失败", config.code)
            try:
                placeholder = Quote(config.code, config.name, 0, self.now())
                evaluation = Evaluation(Trend.UNKNOWN, 0, 0, 0, "未知", False, [])
                self.excel.update_result(placeholder, evaluation, f"更新失败：{exc}")
            except Exception:
                logger.exception("更新失败状态也无法写入 Excel")

    @staticmethod
    def _quote_payload(quote: Quote) -> dict[str, object]:
        return {
            "code": quote.code, "name": quote.name, "price": quote.price,
            "timestamp": quote.timestamp.isoformat(), "stock_code": quote.stock_code,
            "stock_name": quote.stock_name, "stock_price": quote.stock_price,
            "conversion_price": quote.conversion_price, "redeem_status": quote.redeem_status,
        }

    @staticmethod
    def _quote_from_payload(data: dict[str, object]) -> Quote:
        return Quote(
            code=str(data["code"]), name=str(data["name"]), price=float(data["price"]),
            timestamp=datetime.fromisoformat(str(data["timestamp"])),
            stock_code=str(data.get("stock_code") or ""),
            stock_name=str(data.get("stock_name") or ""),
            stock_price=float(data["stock_price"]) if data.get("stock_price") is not None else None,
            conversion_price=float(data["conversion_price"]) if data.get("conversion_price") is not None else None,
            redeem_status=str(data.get("redeem_status") or ""),
        )

    def _queue_result_write(self, quote: Quote, evaluation: Evaluation) -> None:
        payload = {
            "quote": self._quote_payload(quote),
            "evaluation": {
                "trend": evaluation.trend.value,
                "one_year_high": evaluation.one_year_high,
                "monitored_peak": evaluation.monitored_peak,
                "drawdown_pct": evaluation.drawdown_pct,
                "zone": evaluation.zone,
                "sell_alert": evaluation.sell_alert,
                "alert_messages": evaluation.alert_messages,
            },
        }
        self.state.queue_excel_write(
            f"result:{quote.code}", "result", json.dumps(payload, ensure_ascii=False), quote.timestamp
        )

    def _queue_alert_write(self, quote: Quote, alert_type: str, message: str) -> None:
        payload = {"quote": self._quote_payload(quote), "alert_type": alert_type, "message": message}
        key = f"alert:{quote.code}:{alert_type}:{quote.timestamp.isoformat(timespec='seconds')}"
        self.state.queue_excel_write(key, "alert", json.dumps(payload, ensure_ascii=False), quote.timestamp)

    def _flush_pending_excel_writes(self) -> None:
        pending = self.state.pending_excel_writes()
        if not pending:
            return
        completed = 0
        for item_key, item_type, raw_payload in pending:
            try:
                payload = json.loads(raw_payload)
                quote = self._quote_from_payload(payload["quote"])
                if item_type == "result":
                    data = payload["evaluation"]
                    evaluation = Evaluation(
                        trend=Trend(str(data["trend"])),
                        one_year_high=float(data["one_year_high"]),
                        monitored_peak=float(data["monitored_peak"]),
                        drawdown_pct=float(data["drawdown_pct"]),
                        zone=str(data["zone"]), sell_alert=bool(data["sell_alert"]),
                        alert_messages=list(data.get("alert_messages", [])),
                    )
                    self.excel.update_result(quote, evaluation)
                elif item_type == "alert":
                    self.excel.append_alert(quote, str(payload["alert_type"]), str(payload["message"]))
                else:
                    logger.error("丢弃未知 Excel 待写类型：%s", item_type)
                self.state.delete_pending_excel_write(item_key)
                completed += 1
            except PermissionError:
                logger.info("Excel 仍在打开，保留 %d 条待写内容", len(pending) - completed)
                return
            except Exception:
                logger.exception("补写 Excel 队列项目失败：%s", item_key)
                self.state.fail_pending_excel_write(item_key, "载荷损坏或无法写入目标行", self.now())
                logger.error("该项目已移入失败队列，继续补写后续项目：%s", item_key)
                continue
        logger.info("Excel 待写队列补写完成：%d 条", completed)

    @staticmethod
    def _market_phase(current: time, settings: AppSettings) -> str:
        if current < settings.open_time:
            return "before"
        if settings.open_time <= current < settings.lunch_start:
            return "open"
        if settings.lunch_start <= current < settings.lunch_end:
            return "lunch"
        if settings.lunch_end <= current < settings.close_time:
            return "open"
        return "closed"

    @staticmethod
    def _phase_wait_message(phase: str, settings: AppSettings) -> str:
        if phase == "before":
            return f"开盘前，等待 {settings.open_time.strftime('%H:%M')} 开市"
        if phase == "lunch":
            return f"午间休市，等待 {settings.lunch_end.strftime('%H:%M')} 恢复交易"
        if phase == "closed":
            return "今日已收盘，等待下一交易日"
        return "等待进入交易时段"

    def _wait_until_next_day(self) -> None:
        now = self.now()
        tomorrow = now.replace(hour=0, minute=1, second=0, microsecond=0) + timedelta(days=1)
        self._wait(max(30, min(3600, (tomorrow - now).total_seconds())))
