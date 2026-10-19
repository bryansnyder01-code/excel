"""Outgoing email (stdlib smtplib). Without SMTP settings, mail is logged instead so local dev works."""

from __future__ import annotations

import logging
import smtplib
import ssl
from email.message import EmailMessage

from app.config import Settings

log = logging.getLogger(__name__)


def smtp_configured(settings: Settings) -> bool:
    return bool(settings.smtp_host and settings.smtp_from)


def send_mail(settings: Settings, to: str, subject: str, body: str,
              attachments: list[tuple[str, bytes, str]] | None = None) -> None:
    """Send a plain-text email, optionally with (filename, data, mime type) attachments.
    Errors are logged, never raised: callers run this after responding."""
    msg = EmailMessage()
    msg["From"] = settings.smtp_from
    msg["To"] = to
    msg["Subject"] = subject
    msg.set_content(body)
    for filename, data, mime in attachments or []:
        maintype, _, subtype = mime.partition("/")
        msg.add_attachment(data, maintype=maintype, subtype=subtype or "octet-stream", filename=filename)
    try:
        if settings.smtp_port == 465:
            server = smtplib.SMTP_SSL(settings.smtp_host, settings.smtp_port, timeout=20,
                                      context=ssl.create_default_context())
        else:
            server = smtplib.SMTP(settings.smtp_host, settings.smtp_port, timeout=20)
        with server:
            if settings.smtp_port == 587:
                server.starttls(context=ssl.create_default_context())
            if settings.smtp_user:
                server.login(settings.smtp_user, settings.smtp_password)
            server.send_message(msg)
    except (OSError, smtplib.SMTPException):
        log.exception("Couldn't send email to %s", to)


def send_or_log(settings: Settings, to: str, subject: str, body: str) -> None:
    """send_mail when SMTP is set up; otherwise log the message (local dev and tests)."""
    if not smtp_configured(settings):
        log.warning("Email to %s (SMTP not set up): %s\n%s", to, subject, body)
        return
    send_mail(settings, to, subject, body)
