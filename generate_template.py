from pathlib import Path

from kzz_monitor.config import create_workbook

# 仅用于首次安装包的无个人信息模板。正式使用者自行维护监控列表。
create_workbook(Path("可转债监控.xlsx"), ["113043", "113056"])
print("Generated 可转债监控.xlsx")
