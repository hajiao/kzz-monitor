from __future__ import annotations

import json
import logging
import os
import smtplib
import ssl
import subprocess
import sys
from email.message import EmailMessage
from pathlib import Path

from .config import AppSettings
from .secrets import load_secret

logger = logging.getLogger(__name__)


class Notifier:
    def __init__(self, settings: AppSettings, secret_path: Path) -> None:
        self.settings = settings
        self.secret_path = secret_path

    def send(self, title: str, message: str) -> None:
        logger.warning("%s | %s", title, message)
        if self.settings.desktop_notification:
            self._desktop(title, message)
        if self.settings.email_notification and self.settings.email_recipients:
            self._email(title, message)

    def send_email_test(self) -> bool:
        return self._email("邮件配置测试", "KzzMonitor 邮件发送设置有效。")

    @staticmethod
    def _desktop(title: str, message: str) -> None:
        try:
            if sys.platform == "darwin":
                script = f"display notification {json.dumps(message)} with title {json.dumps(title)}"
                subprocess.run(["osascript", "-e", script], check=True, capture_output=True)
            else:
                from winotify import Notification

                Notification(app_id="可转债监控", title=title, msg=message).show()
        except Exception:
            logger.exception("桌面通知发送失败")

    def _email(self, title: str, message: str) -> bool:
        host = self.settings.smtp_host or os.getenv("KZZ_SMTP_HOST", "").strip()
        user = self.settings.smtp_user or os.getenv("KZZ_SMTP_USER", "").strip()
        password = load_secret(self.secret_path) or os.getenv("KZZ_SMTP_PASSWORD", "")
        sender = self.settings.smtp_from or os.getenv("KZZ_SMTP_FROM", user).strip()
        port = self.settings.smtp_port or int(os.getenv("KZZ_SMTP_PORT", "465"))
        use_ssl = self.settings.smtp_ssl
        if not host or not sender or not password or not self.settings.email_recipients:
            logger.error("SMTP 服务器、发件人、授权码或收件人尚未配置")
            return False
        mail = EmailMessage()
        mail["Subject"] = f"[可转债监控] {title}"
        mail["From"] = sender
        mail["To"] = ", ".join(self.settings.email_recipients)
        mail.set_content(message)
        try:
            if use_ssl:
                with smtplib.SMTP_SSL(host, port, timeout=20, context=ssl.create_default_context()) as smtp:
                    if user:
                        smtp.login(user, password)
                    smtp.send_message(mail)
            else:
                with smtplib.SMTP(host, port, timeout=20) as smtp:
                    smtp.starttls(context=ssl.create_default_context())
                    if user:
                        smtp.login(user, password)
                    smtp.send_message(mail)
            logger.info("测试邮件发送成功：%s", ", ".join(self.settings.email_recipients))
            return True
        except Exception:
            logger.exception("邮件发送失败")
            return False
