"""Lightweight SMTP email sender. Currently used for the monthly SPOC error
report. Gmail SMTP is the default backend — needs an App Password (not the
regular Gmail password) generated at https://myaccount.google.com/apppasswords.

Env vars required:
    SMTP_HOST   default: smtp.gmail.com
    SMTP_PORT   default: 587
    SMTP_USER   sender Gmail address
    SMTP_PASS   16-char Gmail App Password
    SMTP_FROM   optional display name, e.g. 'BCML TMS <noreply@bcml.in>'
"""
import os
import smtplib
import logging
from email.message import EmailMessage


def send_email(to_addrs, subject, body_html, body_text=None, attachments=None):
    """Send an email. attachments = list of (filename, bytes, mime_type) tuples.

    Returns (ok: bool, detail: str). Never raises — logs errors instead so the
    caller (cron endpoint) can return a clean JSON response.
    """
    host = os.environ.get('SMTP_HOST', 'smtp.gmail.com')
    try:
        port = int(os.environ.get('SMTP_PORT', '587'))
    except ValueError:
        port = 587
    user = os.environ.get('SMTP_USER', '').strip()
    pwd  = os.environ.get('SMTP_PASS', '').strip()
    if not user or not pwd:
        return False, 'SMTP_USER or SMTP_PASS env var not set'

    sender_display = os.environ.get('SMTP_FROM', '').strip() or user

    if isinstance(to_addrs, str):
        to_addrs = [a.strip() for a in to_addrs.split(',') if a.strip()]
    if not to_addrs:
        return False, 'No recipients'

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From']    = sender_display
    msg['To']      = ', '.join(to_addrs)
    msg.set_content(body_text or 'This email requires an HTML-capable client.')
    msg.add_alternative(body_html, subtype='html')

    for fname, data, mime in (attachments or []):
        try:
            maintype, _, subtype = mime.partition('/')
            msg.add_attachment(data, maintype=maintype or 'application',
                               subtype=subtype or 'octet-stream', filename=fname)
        except Exception as e:
            logging.warning(f'email attach failed for {fname}: {e}')

    try:
        with smtplib.SMTP(host, port, timeout=30) as srv:
            srv.starttls()
            srv.login(user, pwd)
            srv.send_message(msg)
        return True, f'Sent to {", ".join(to_addrs)}'
    except Exception as e:
        logging.error(f'SMTP send failed: {e}', exc_info=True)
        return False, str(e)
