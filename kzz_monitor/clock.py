from datetime import datetime
from zoneinfo import ZoneInfo

CHINA_TZ = ZoneInfo("Asia/Shanghai")


def china_now() -> datetime:
    """返回中国证券市场使用的北京时间，不依赖 Windows 系统时区。"""
    return datetime.now(CHINA_TZ)
