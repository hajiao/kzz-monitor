from __future__ import annotations

import logging
import queue
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from PIL import Image, ImageDraw
import pystray

from . import __version__
from .clock import china_now
from .config import load_configuration, update_settings
from .notifier import Notifier
from .platform_utils import open_path
from .secrets import load_secret, save_secret
from .service import MonitorService
from .updater import UpdateInfo, check_for_update, stage_and_launch_update

logger = logging.getLogger(__name__)


class QueueLogHandler(logging.Handler):
    def __init__(self, messages: queue.Queue[str]) -> None:
        super().__init__()
        self.messages = messages
        self.setFormatter(logging.Formatter("%(asctime)s %(levelname)s | %(message)s", "%H:%M:%S"))

    def emit(self, record: logging.LogRecord) -> None:
        self.messages.put(self.format(record))


def tray_image() -> Image.Image:
    image = Image.new("RGBA", (64, 64), (30, 90, 165, 255))
    draw = ImageDraw.Draw(image)
    draw.ellipse((5, 5, 59, 59), fill=(45, 125, 210, 255), outline="white", width=3)
    draw.text((13, 18), "KZZ", fill="white", stroke_width=1)
    return image


class MonitorApp:
    def __init__(self, root: tk.Tk, workbook: Path, state_path: Path) -> None:
        self.root = root
        self.workbook = workbook
        self.state_path = state_path
        self.service: MonitorService | None = None
        self.worker: threading.Thread | None = None
        self.quitting = False
        self.messages: queue.Queue[str] = queue.Queue()
        self.log_handler = QueueLogHandler(self.messages)
        logging.getLogger().addHandler(self.log_handler)
        self.tray: pystray.Icon | None = None
        self.vars: dict[str, tk.Variable] = {}

        self.root.title("可转债监控控制台")
        self.root.geometry("980x720")
        self.root.minsize(860, 620)
        self.root.protocol("WM_DELETE_WINDOW", self.confirm_close)
        self._build_ui()
        self._load_settings()
        self._start_tray()
        self.root.after(200, self._drain_logs)
        self.root.after(500, self.start_monitor)
        self.root.after(1000, self._refresh_status)
        self.root.after(2500, self._startup_update_check)

    def _build_ui(self) -> None:
        style = ttk.Style()
        ui_font = "PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei UI"
        style.configure("Title.TLabel", font=(ui_font, 15, "bold"))
        style.configure("Status.TLabel", font=(ui_font, 11, "bold"))
        container = ttk.Frame(self.root, padding=14)
        container.pack(fill="both", expand=True)

        top = ttk.Frame(container)
        top.pack(fill="x")
        ttk.Label(top, text=f"可转债监控控制台  v{__version__}", style="Title.TLabel").pack(side="left")
        self.status_label = ttk.Label(top, text="● 正在准备", foreground="#c27c00", style="Status.TLabel")
        self.status_label.pack(side="right")

        self.summary_label = ttk.Label(container, text="", padding=(0, 8))
        self.summary_label.pack(fill="x")

        settings_box = ttk.LabelFrame(container, text="常用设置（保存后下一轮生效）", padding=10)
        settings_box.pack(fill="x", pady=(4, 10))
        definitions = [
            ("轮询间隔秒", "单只查询间隔（秒）", "entry"),
            ("整轮间隔分钟", "整轮启动间隔（分钟）", "entry"),
            ("开盘时间", "开盘时间", "entry"),
            ("午间休市开始", "午休开始", "entry"),
            ("午间休市结束", "午休结束", "entry"),
            ("收盘时间", "收盘时间", "entry"),
            ("安道全文件", "安道全文件", "file"),
            ("桌面通知", "桌面通知", "check"),
            ("邮件通知", "邮件通知", "check"),
            ("邮件收件人", "邮件收件人", "entry"),
            ("SMTP服务器", "SMTP 服务器", "entry"),
            ("SMTP端口", "SMTP 端口", "entry"),
            ("SMTP用户名", "SMTP 用户名", "entry"),
            ("SMTP发件人", "SMTP 发件人", "entry"),
            ("SMTP使用SSL", "SMTP 使用 SSL", "check"),
            ("SMTP授权码", "SMTP 授权码", "password"),
            ("更新清单地址", "更新清单地址", "entry"),
            ("启动时检查更新", "启动时检查更新", "check"),
        ]
        for index, (key, label, kind) in enumerate(definitions):
            row, pair = divmod(index, 2)
            base = pair * 3
            ttk.Label(settings_box, text=label).grid(row=row, column=base, sticky="w", padx=(0, 6), pady=4)
            if kind == "check":
                variable: tk.Variable = tk.BooleanVar()
                ttk.Checkbutton(settings_box, variable=variable).grid(row=row, column=base + 1, sticky="w", pady=4)
            else:
                variable = tk.StringVar()
                entry = ttk.Entry(settings_box, textvariable=variable, width=28 if kind != "file" else 42)
                if kind == "password":
                    entry.configure(show="●")
                entry.grid(row=row, column=base + 1, sticky="ew", pady=4)
                if kind == "file":
                    ttk.Button(settings_box, text="浏览", command=self._choose_adq).grid(row=row, column=base + 2, padx=(5, 8))
            self.vars[key] = variable
        settings_box.columnconfigure(1, weight=1)
        settings_box.columnconfigure(4, weight=1)

        actions = ttk.Frame(container)
        actions.pack(fill="x", pady=(0, 10))
        primary_actions = ttk.Frame(actions)
        primary_actions.pack(fill="x", pady=(0, 5))
        secondary_actions = ttk.Frame(actions)
        secondary_actions.pack(fill="x")
        self.start_button = ttk.Button(primary_actions, text="开启轮询", command=self.start_monitor)
        self.start_button.pack(side="left")
        self.stop_button = ttk.Button(primary_actions, text="结束轮询", command=self.stop_monitor)
        self.stop_button.pack(side="left", padx=6)
        ttk.Button(primary_actions, text="立即强制新一轮", command=self.force_cycle).pack(side="left", padx=6)
        ttk.Button(primary_actions, text="刷新安道全", command=self.refresh_adq).pack(side="left", padx=6)
        ttk.Button(primary_actions, text="保存设置", command=self.save_settings).pack(side="left", padx=6)
        ttk.Button(primary_actions, text="测试邮件", command=self.test_email).pack(side="left", padx=6)
        ttk.Button(secondary_actions, text="打开监控表", command=self.open_workbook).pack(side="left")
        ttk.Button(secondary_actions, text="操作手册", command=self.open_manual).pack(side="left", padx=6)
        ttk.Button(secondary_actions, text="检查更新", command=self.check_updates).pack(side="left", padx=6)
        ttk.Button(secondary_actions, text="打开日志目录", command=self.open_logs).pack(side="left", padx=6)
        ttk.Button(secondary_actions, text="隐藏到托盘", command=self.hide_window).pack(side="right")

        ttk.Label(container, text="运行日志").pack(anchor="w")
        log_frame = ttk.Frame(container)
        log_frame.pack(fill="both", expand=True, pady=(4, 0))
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled", font=("Consolas", 9), bg="#101820", fg="#d7e3ef")
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    def _load_settings(self) -> None:
        settings, bonds = load_configuration(self.workbook)
        values: dict[str, Any] = {
            "轮询间隔秒": settings.poll_interval_seconds,
            "整轮间隔分钟": settings.cycle_interval_minutes,
            "开盘时间": settings.open_time.strftime("%H:%M"),
            "午间休市开始": settings.lunch_start.strftime("%H:%M"),
            "午间休市结束": settings.lunch_end.strftime("%H:%M"),
            "收盘时间": settings.close_time.strftime("%H:%M"),
            "安道全文件": str(settings.adq_file or ""),
            "桌面通知": settings.desktop_notification,
            "邮件通知": settings.email_notification,
            "邮件收件人": ",".join(settings.email_recipients),
            "SMTP服务器": settings.smtp_host,
            "SMTP端口": settings.smtp_port,
            "SMTP用户名": settings.smtp_user,
            "SMTP发件人": settings.smtp_from,
            "SMTP使用SSL": settings.smtp_ssl,
            "SMTP授权码": load_secret(self.state_path.parent / "smtp_secret.bin"),
            "更新清单地址": settings.update_manifest_url,
            "启动时检查更新": settings.check_updates_on_startup,
        }
        for key, value in values.items():
            self.vars[key].set(value)
        self.summary_label.configure(text=f"配置文件：{self.workbook}    已启用转债：{len(bonds)} 只")

    def save_settings(self) -> bool:
        try:
            interval = max(10, int(str(self.vars["轮询间隔秒"].get()).strip()))
            cycle_interval = max(1, int(str(self.vars["整轮间隔分钟"].get()).strip()))
            for key in ("开盘时间", "午间休市开始", "午间休市结束", "收盘时间"):
                datetime.strptime(str(self.vars[key].get()).strip(), "%H:%M")
            values = {key: variable.get() for key, variable in self.vars.items()}
            values["轮询间隔秒"] = interval
            values["整轮间隔分钟"] = cycle_interval
            password = str(values.pop("SMTP授权码", ""))
            values["SMTP端口"] = int(str(values["SMTP端口"]).strip())
            update_settings(self.workbook, values)
            save_secret(self.state_path.parent / "smtp_secret.bin", password)
            self._load_settings()
            self._append_log("设置已保存；后台将在下一轮重新读取。")
            return True
        except PermissionError:
            messagebox.showerror("无法保存", "监控工作簿正在 Excel 中打开。请先保存并关闭 Excel。")
            return False
        except Exception as exc:
            messagebox.showerror("设置有误", str(exc))
            return False

    def _choose_adq(self) -> None:
        selected = filedialog.askopenfilename(title="选择安道全 Excel", filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if selected:
            self.vars["安道全文件"].set(selected)

    def start_monitor(self) -> None:
        if self.worker and self.worker.is_alive():
            self.show_window()
            return
        self.service = MonitorService(self.workbook, self.state_path, status_callback=self._service_status)
        self.worker = threading.Thread(target=self._run_service, name="KzzMonitorWorker", daemon=True)
        self.worker.start()
        self.status_label.configure(text="● 监控运行中", foreground="#14833b")
        self._update_tray_menu()

    def _service_status(self, state: str, message: str) -> None:
        if not self.quitting:
            self.root.after(0, self._apply_service_status, state, message)

    def _apply_service_status(self, state: str, message: str) -> None:
        labels = {
            "checking": ("● 正在检查", "#c27c00"),
            "running": ("● 监控正常", "#14833b"),
            "waiting": ("● 非交易时段", "#2563a8"),
            "error": ("● 异常重试中", "#b42318"),
            "stopped": ("● 已停止", "#b42318"),
        }
        text, color = labels.get(state, ("● 运行中", "#14833b"))
        self.status_label.configure(text=text, foreground=color)
        self._append_log(f"状态：{message}")
        if self.tray:
            self.tray.title = f"可转债监控 - {message}"

    def _run_service(self) -> None:
        assert self.service is not None
        try:
            self.service.run(install_signal_handlers=False)
        except Exception:
            logger.exception("监控线程异常退出")
        finally:
            if not self.quitting:
                self.root.after(0, self._mark_stopped)

    def stop_monitor(self) -> None:
        if self.service:
            self.service.stop()
        self.status_label.configure(text="● 正在停止", foreground="#c27c00")

    def force_cycle(self) -> None:
        if not self.worker or not self.worker.is_alive():
            self.start_monitor()
        if self.service:
            self.service.request_force_cycle()
            self._append_log("已请求立即开始新一轮完整轮询（仍按设置的单只间隔执行）。")

    def refresh_adq(self) -> None:
        if not self.worker or not self.worker.is_alive():
            self.start_monitor()
        if self.service:
            self.service.request_adq_refresh()
            self._append_log("已请求立即重新导入安道全评级和三段线。")

    def test_email(self) -> None:
        if not self.save_settings():
            return
        try:
            settings, _ = load_configuration(self.workbook)
            success = Notifier(settings, self.state_path.parent / "smtp_secret.bin").send_email_test()
            if success:
                messagebox.showinfo("测试成功", "测试邮件已经发出，请检查收件箱和垃圾邮件目录。")
            else:
                messagebox.showerror("测试失败", "邮件未发送。请检查 SMTP 设置，并查看运行日志。")
        except Exception as exc:
            logger.exception("测试邮件失败")
            messagebox.showerror("测试失败", str(exc))

    def _mark_stopped(self) -> None:
        self.service = None
        self.worker = None
        self.status_label.configure(text="● 已停止", foreground="#b42318")
        self._update_tray_menu()

    def _refresh_status(self) -> None:
        running = bool(self.worker and self.worker.is_alive())
        now = china_now().strftime("%Y-%m-%d %H:%M:%S")
        current = "运行中" if running else "已停止"
        pending = 0
        if self.service and running:
            try:
                pending = self.service.state.pending_excel_write_count()
            except Exception:
                pending = 0
        self.root.title(f"可转债监控控制台 - {current}")
        base_text = self.summary_label.cget("text").split("    北京时间")[0]
        self.summary_label.configure(
            text=f"{base_text}    北京时间：{now}    Excel待写：{pending} 条"
        )
        if not self.quitting:
            self.root.after(1000, self._refresh_status)

    def _drain_logs(self) -> None:
        while True:
            try:
                self._append_log(self.messages.get_nowait())
            except queue.Empty:
                break
        if not self.quitting:
            self.root.after(200, self._drain_logs)

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def open_workbook(self) -> None:
        open_path(self.workbook)

    def open_logs(self) -> None:
        folder = self.state_path.parent.parent / "logs"
        folder.mkdir(parents=True, exist_ok=True)
        open_path(folder)

    def open_manual(self) -> None:
        base = self.state_path.parent.parent
        for name in ("KzzMonitor详细操作手册.html", "KzzMonitor详细操作手册.pdf", "README.md"):
            manual = base / name
            if manual.exists():
                open_path(manual)
                return
        messagebox.showinfo("操作手册", "当前目录没有操作手册，请重新解压完整便携包。")

    def _startup_update_check(self) -> None:
        try:
            settings, _ = load_configuration(self.workbook)
            if settings.check_updates_on_startup and settings.update_manifest_url:
                self.check_updates(silent_when_current=True)
        except Exception:
            logger.exception("启动时检查更新失败")

    def check_updates(self, silent_when_current: bool = False) -> None:
        try:
            location = str(self.vars["更新清单地址"].get()).strip()
            if not location:
                if not silent_when_current:
                    messagebox.showinfo("检查更新", "请先填写“更新清单地址”并保存设置。")
                return
            self._append_log("正在检查程序更新……")
            threading.Thread(
                target=self._check_updates_worker,
                args=(location, silent_when_current),
                daemon=True,
                name="KzzUpdateChecker",
            ).start()
        except Exception as exc:
            messagebox.showerror("检查更新失败", str(exc))

    def _check_updates_worker(self, location: str, silent_when_current: bool) -> None:
        try:
            info = check_for_update(location)
            self.root.after(0, self._show_update_result, info, silent_when_current)
        except Exception as exc:
            logger.exception("检查更新失败")
            self.root.after(0, messagebox.showerror, "检查更新失败", str(exc))

    def _show_update_result(self, info: UpdateInfo | None, silent_when_current: bool) -> None:
        if info is None:
            self._append_log(f"当前已是最新版本 v{__version__}。")
            if not silent_when_current:
                messagebox.showinfo("检查更新", f"当前已是最新版本 v{__version__}。")
            return
        notes = f"\n\n更新说明：\n{info.notes}" if info.notes else ""
        install = messagebox.askyesno(
            "发现新版本",
            f"当前版本：v{__version__}\n新版本：v{info.version}{notes}\n\n"
            "是否立即下载并安装？程序会自动退出、替换并重启；Excel、data 和 logs 不会被覆盖。",
        )
        if not install:
            return
        self._append_log(f"正在下载并校验 v{info.version} 更新包……")
        threading.Thread(
            target=self._install_update_worker,
            args=(info,),
            daemon=True,
            name="KzzUpdateInstaller",
        ).start()

    def _install_update_worker(self, info: UpdateInfo) -> None:
        try:
            stage_and_launch_update(info, self.state_path.parent.parent)
            self.root.after(0, self._exit_main)
        except Exception as exc:
            logger.exception("安装更新失败")
            self.root.after(0, messagebox.showerror, "安装更新失败", str(exc))

    def hide_window(self) -> None:
        self.root.withdraw()
        if self.tray:
            try:
                self.tray.notify("程序仍在后台运行。双击托盘图标可重新打开控制台。", "可转债监控")
            except Exception:
                pass

    def confirm_close(self) -> None:
        choice = messagebox.askyesnocancel(
            "关闭 KzzMonitor",
            "是否彻底退出 KzzMonitor？\n\n"
            "“是”：停止轮询并退出程序\n"
            "“否”：隐藏到后台，继续运行\n"
            "“取消”：返回控制台",
            icon="question",
        )
        if choice is True:
            self._exit_main()
        elif choice is False:
            self.hide_window()

    def show_window(self, *_: object) -> None:
        self.root.after(0, self._show_window_main)

    def _show_window_main(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def exit_app(self, *_: object) -> None:
        self.root.after(0, self._exit_main)

    def _exit_main(self) -> None:
        self.quitting = True
        if self.service:
            self.service.stop()
        if self.tray:
            self.tray.stop()
        logging.getLogger().removeHandler(self.log_handler)
        self.root.destroy()

    def _start_tray(self) -> None:
        menu = pystray.Menu(
            pystray.MenuItem("显示控制台", self.show_window, default=True),
            pystray.MenuItem("启动监控", lambda *_: self.root.after(0, self.start_monitor)),
            pystray.MenuItem("停止监控", lambda *_: self.root.after(0, self.stop_monitor)),
            pystray.MenuItem("立即强制新一轮", lambda *_: self.root.after(0, self.force_cycle)),
            pystray.MenuItem("刷新安道全", lambda *_: self.root.after(0, self.refresh_adq)),
            pystray.MenuItem("检查更新", lambda *_: self.root.after(0, self.check_updates)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("退出 KzzMonitor", self.exit_app),
        )
        tray_options: dict[str, Any] = {}
        if sys.platform == "darwin":
            from AppKit import NSApplication

            tray_options["darwin_nsapplication"] = NSApplication.sharedApplication()
        self.tray = pystray.Icon(
            "KzzMonitor", tray_image(), "可转债监控 - 正在启动", menu, **tray_options
        )
        self.tray.run_detached()

    def _update_tray_menu(self) -> None:
        if self.tray:
            running = bool(self.worker and self.worker.is_alive())
            self.tray.title = f"可转债监控 - {'运行中' if running else '已停止'}"
            self.tray.update_menu()


def run_gui(workbook: Path, state_path: Path) -> None:
    root = tk.Tk()
    MonitorApp(root, workbook, state_path)
    root.mainloop()
