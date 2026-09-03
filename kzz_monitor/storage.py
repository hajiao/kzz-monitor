from __future__ import annotations

import sqlite3
import threading
from datetime import date, datetime
from pathlib import Path


class StateStore:
    def __init__(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        self._closed = False
        self.connection = sqlite3.connect(path, check_same_thread=False, timeout=30)
        with self._lock:
            self.connection.execute("PRAGMA journal_mode=WAL")
            self.connection.execute("PRAGMA busy_timeout=30000")
            self.connection.executescript(
                """
            CREATE TABLE IF NOT EXISTS prices (
                code TEXT NOT NULL,
                timestamp TEXT NOT NULL,
                price REAL NOT NULL,
                PRIMARY KEY (code, timestamp)
            );
            CREATE INDEX IF NOT EXISTS idx_prices_code_time ON prices(code, timestamp DESC);
            CREATE TABLE IF NOT EXISTS runtime_state (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS alerts (
                fingerprint TEXT PRIMARY KEY,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS daily_highs (
                code TEXT NOT NULL,
                trade_date TEXT NOT NULL,
                high REAL NOT NULL,
                PRIMARY KEY (code, trade_date)
            );
            CREATE TABLE IF NOT EXISTS pending_excel_writes (
                item_key TEXT PRIMARY KEY,
                item_type TEXT NOT NULL,
                payload TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            """
            )
            self.connection.commit()

    def close(self) -> None:
        with self._lock:
            if not self._closed:
                self.connection.close()
                self._closed = True

    def add_price(self, code: str, timestamp: datetime, price: float) -> None:
        with self._lock:
            self.connection.execute(
                "INSERT OR REPLACE INTO prices(code, timestamp, price) VALUES (?, ?, ?)",
                (code, timestamp.isoformat(timespec="seconds"), price),
            )
            self.connection.commit()

    def recent_prices(self, code: str, count: int) -> list[float]:
        with self._lock:
            rows = self.connection.execute(
                "SELECT price FROM prices WHERE code=? ORDER BY timestamp DESC LIMIT ?", (code, count)
            ).fetchall()
        return [float(row[0]) for row in reversed(rows)]

    def monitored_peak(self, code: str) -> float:
        with self._lock:
            row = self.connection.execute("SELECT MAX(price) FROM prices WHERE code=?", (code,)).fetchone()
        return float(row[0]) if row and row[0] is not None else 0.0

    def set_value(self, key: str, value: str) -> None:
        with self._lock:
            self.connection.execute(
                "INSERT INTO runtime_state(key, value) VALUES (?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, value),
            )
            self.connection.commit()

    def get_value(self, key: str, default: str = "") -> str:
        with self._lock:
            row = self.connection.execute("SELECT value FROM runtime_state WHERE key=?", (key,)).fetchone()
        return str(row[0]) if row else default

    def claim_alert(self, fingerprint: str, now: datetime) -> bool:
        with self._lock:
            cursor = self.connection.execute(
                "INSERT OR IGNORE INTO alerts(fingerprint, created_at) VALUES (?, ?)",
                (fingerprint, now.isoformat(timespec="seconds")),
            )
            self.connection.commit()
            return cursor.rowcount == 1

    def get_daily_high(self, code: str, trade_date: date) -> float | None:
        with self._lock:
            row = self.connection.execute(
                "SELECT high FROM daily_highs WHERE code=? AND trade_date=?", (code, trade_date.isoformat())
            ).fetchone()
        return float(row[0]) if row else None

    def set_daily_high(self, code: str, trade_date: date, high: float) -> None:
        with self._lock:
            self.connection.execute(
                "INSERT OR REPLACE INTO daily_highs(code, trade_date, high) VALUES (?, ?, ?)",
                (code, trade_date.isoformat(), high),
            )
            self.connection.commit()

    def purge_prices(self, before: datetime) -> None:
        with self._lock:
            self.connection.execute("DELETE FROM prices WHERE timestamp < ?", (before.isoformat(),))
            self.connection.commit()

    def queue_excel_write(self, item_key: str, item_type: str, payload: str, now: datetime) -> None:
        """持久化待写内容；相同 key 的行情结果只保留最新值。"""
        with self._lock:
            self.connection.execute(
                "INSERT INTO pending_excel_writes(item_key, item_type, payload, created_at) VALUES (?, ?, ?, ?) "
                "ON CONFLICT(item_key) DO UPDATE SET item_type=excluded.item_type, "
                "payload=excluded.payload, created_at=excluded.created_at",
                (item_key, item_type, payload, now.isoformat(timespec="seconds")),
            )
            self.connection.commit()

    def pending_excel_writes(self) -> list[tuple[str, str, str]]:
        with self._lock:
            rows = self.connection.execute(
                "SELECT item_key, item_type, payload FROM pending_excel_writes ORDER BY created_at, item_key"
            ).fetchall()
        return [(str(key), str(item_type), str(payload)) for key, item_type, payload in rows]

    def delete_pending_excel_write(self, item_key: str) -> None:
        with self._lock:
            self.connection.execute("DELETE FROM pending_excel_writes WHERE item_key=?", (item_key,))
            self.connection.commit()

    def pending_excel_write_count(self) -> int:
        with self._lock:
            row = self.connection.execute("SELECT COUNT(*) FROM pending_excel_writes").fetchone()
        return int(row[0]) if row else 0
