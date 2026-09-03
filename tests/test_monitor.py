from datetime import datetime, time
import hashlib
import json
from pathlib import Path
import threading
from zoneinfo import ZoneInfo

import pytest

from kzz_monitor.config import AppSettings, BondConfig, normalize_code
from kzz_monitor.config import create_workbook, load_configuration, migrate_workbook
from kzz_monitor.excel_store import ExcelStore
from kzz_monitor.models import Evaluation, Quote, Trend
from kzz_monitor.provider import exchange_symbol
from kzz_monitor.service import AlertEngine, MonitorService
from kzz_monitor.storage import StateStore
from kzz_monitor.updater import check_for_update


def quote(price: float, minute: int) -> Quote:
    return Quote("127015", "希望转债", price, datetime(2026, 9, 3, 10, minute))


def test_sell_alert_requires_arming_downtrend_and_drawdown(tmp_path):
    state = StateStore(tmp_path / "state.db")
    engine = AlertEngine(state)
    config = BondConfig("127015", sell_trigger_price=130, sell_drawdown_pct=5, trend_window=2, trend_epsilon=0.1)

    assert not engine.evaluate(config, quote(140, 0), 145).sell_alert
    result = engine.evaluate(config, quote(132, 1), 145)
    assert result.trend == Trend.DOWN
    assert result.sell_alert
    assert result.drawdown_pct == pytest.approx(5.714, rel=1e-3)
    assert "卖出观察" in result.alert_messages[0]

    result = engine.evaluate(config, quote(130, 2), 145)
    assert result.sell_alert
    assert not result.alert_messages  # 同一轮下跌只提醒一次
    state.close()


def test_position_zone_only_alerts_when_moving_deeper(tmp_path):
    state = StateStore(tmp_path / "state.db")
    engine = AlertEngine(state)
    config = BondConfig("127015", build_line=120, add_line=115, heavy_line=110)

    assert engine.evaluate(config, quote(119, 0), 130).alert_messages == ["进入建仓：现价 119.00"]
    assert engine.evaluate(config, quote(114, 1), 130).alert_messages == ["进入加仓：现价 114.00"]
    assert engine.evaluate(config, quote(116, 2), 130).alert_messages == []
    state.close()


def test_market_phases():
    settings = AppSettings()
    assert MonitorService._market_phase(time(9, 0), settings) == "before"
    assert MonitorService._market_phase(time(9, 30), settings) == "open"
    assert MonitorService._market_phase(time(12, 0), settings) == "lunch"
    assert MonitorService._market_phase(time(14, 59), settings) == "open"
    assert MonitorService._market_phase(time(15, 0), settings) == "closed"


def test_cycle_delay_is_start_to_start_and_never_negative():
    assert MonitorService._cycle_delay_seconds(60, 15 * 60) == 45 * 60
    assert MonitorService._cycle_delay_seconds(60, 60 * 60) == 0
    assert MonitorService._cycle_delay_seconds(60, 70 * 60) == 0


def test_wait_until_next_day_preserves_timezone(tmp_path):
    class ProviderStub:
        pass

    workbook = tmp_path / "monitor.xlsx"
    create_workbook(workbook, ["113043"])
    current = datetime(2026, 9, 3, 19, 9, 58, tzinfo=ZoneInfo("Asia/Shanghai"))
    service = MonitorService(
        workbook,
        tmp_path / "monitor.db",
        provider=ProviderStub(),  # type: ignore[arg-type]
        now=lambda: current,
    )
    waits: list[float] = []
    service._wait = lambda seconds: waits.append(seconds) or False  # type: ignore[method-assign]
    service._wait_until_next_day()
    assert waits == [3600]
    service.state.close()


def test_code_and_exchange_normalization():
    assert normalize_code("SZ127015") == "127015"
    assert normalize_code(110059.0) == "110059"
    assert exchange_symbol("113056") == "sh113056"
    assert exchange_symbol("127015") == "sz127015"


def test_excel_accepts_timezone_aware_quote(tmp_path):
    workbook = tmp_path / "monitor.xlsx"
    create_workbook(workbook, ["127015"])
    current = Quote("127015", "希望转债", 123.45, datetime(2026, 9, 3, 10, 0, tzinfo=ZoneInfo("Asia/Shanghai")))
    evaluation = Evaluation(Trend.FLAT, 140, 130, 5, "观察", False, [])
    ExcelStore(workbook).update_result(current, evaluation)


def test_excel_marks_unavailable_bond_without_error(tmp_path):
    workbook = tmp_path / "monitor.xlsx"
    create_workbook(workbook, ["127015"])
    ExcelStore(workbook).mark_unavailable(
        "127015",
        datetime(2026, 9, 3, 10, 0, tzinfo=ZoneInfo("Asia/Shanghai")),
        "最后交易日 2025-12-26，可能已退市",
    )


def test_state_store_can_move_from_gui_thread_to_worker(tmp_path):
    state = StateStore(tmp_path / "threaded.db")
    errors: list[BaseException] = []

    def worker() -> None:
        try:
            state.set_daily_high("113043", datetime(2026, 9, 3).date(), 145.5)
            assert state.get_daily_high("113043", datetime(2026, 9, 3).date()) == 145.5
            state.add_price("113043", datetime(2026, 9, 3, 10, 0), 120.5)
        except BaseException as exc:
            errors.append(exc)

    thread = threading.Thread(target=worker)
    thread.start()
    thread.join()
    assert errors == []
    assert state.recent_prices("113043", 1) == [120.5]
    state.close()


def test_monitor_service_retains_state_path_for_notifier(tmp_path):
    class ProviderStub:
        pass

    workbook = tmp_path / "monitor.xlsx"
    database = tmp_path / "data" / "monitor.db"
    create_workbook(workbook, ["113043"])
    service = MonitorService(workbook, database, provider=ProviderStub())  # type: ignore[arg-type]
    assert service.state_path == database
    assert service.state_path.parent / "smtp_secret.bin" == tmp_path / "data" / "smtp_secret.bin"
    service.state.close()


def test_workbook_migration_preserves_existing_settings(tmp_path):
    workbook = tmp_path / "monitor.xlsx"
    create_workbook(workbook, ["113043"])
    migrate_workbook(workbook)
    settings, bonds = load_configuration(workbook)
    assert settings.cycle_interval_minutes == 60
    assert bonds[0].code == "113043"


def test_local_update_manifest_resolves_relative_package(tmp_path):
    package = tmp_path / "KzzMonitor-update.zip"
    package.write_bytes(b"example-update")
    manifest = tmp_path / "update-manifest.json"
    manifest.write_text(json.dumps({
        "version": "9.0.0",
        "notes": "test",
        "artifacts": {key: {
                "url": package.name,
                "sha256": hashlib.sha256(package.read_bytes()).hexdigest(),
            } for key in ("windows-x64", "macos-arm64", "macos-x64")},
    }), encoding="utf-8")
    info = check_for_update(str(manifest))
    assert info is not None
    assert Path(info.url) == package.resolve()


def test_excel_pending_queue_persists_and_flushes(tmp_path):
    class ProviderStub:
        pass

    workbook = tmp_path / "monitor.xlsx"
    database = tmp_path / "data" / "monitor.db"
    create_workbook(workbook, ["113043"])
    service = MonitorService(workbook, database, provider=ProviderStub())  # type: ignore[arg-type]
    current = Quote("113043", "财通转债", 121.5, datetime(2026, 9, 3, 10, 0))
    evaluation = Evaluation(Trend.DOWN, 140, 130, 6.5, "观察", True, ["卖出观察"])
    service._queue_result_write(current, evaluation)
    service._queue_alert_write(current, "卖出", "卖出观察")
    assert service.state.pending_excel_write_count() == 2
    service.state.close()

    restarted = MonitorService(workbook, database, provider=ProviderStub())  # type: ignore[arg-type]
    assert restarted.state.pending_excel_write_count() == 2
    restarted._flush_pending_excel_writes()
    assert restarted.state.pending_excel_write_count() == 0
    restarted.state.close()
