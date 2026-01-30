#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import email
import html
import imaplib
import os
import re
import secrets
import socket
import sqlite3
import string
import sys
import threading
import time
import webbrowser
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime

import requests
from flask import (
    Flask,
    abort,
    flash,
    jsonify,
    redirect,
    render_template_string,
    request,
    url_for,
)

APP = Flask(__name__)
APP.secret_key = os.environ.get("MAILADMIN_SECRET", secrets.token_hex(16))

#DB_PATH = os.path.join(os.path.dirname(__file__), "mailadmin.db")
DB_PATH = "mailadmin.db"
IMAP_HOST = os.environ.get("MAILADMIN_IMAP_HOST", "outlook.live.com")
DEFAULT_LIMIT = 10
MAX_LIMIT = 50
SHARE_CODE_LEN = 8
JUNK_FOLDERS = [
    "junk",
    "Junk",
    "Junk Email",
    "Junk E-mail",
]

BASE_TEMPLATE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ title }}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Fraunces:wght@500;700&family=Space+Grotesk:wght@400;500;600&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg: #f5f1e8;
      --bg-accent: #e9f1ff;
      --ink: #1c1b19;
      --muted: #6b6a66;
      --card: #ffffff;
      --accent: #e24a3b;
      --accent-2: #2e5f9b;
      --ring: rgba(226, 74, 59, 0.25);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background: radial-gradient(circle at 10% 20%, #f4e2c8, transparent 45%),
                  radial-gradient(circle at 90% 10%, #d7e6ff, transparent 40%),
                  linear-gradient(120deg, var(--bg), var(--bg-accent));
      color: var(--ink);
      font-family: "Space Grotesk", sans-serif;
      display: flex;
      justify-content: center;
    }
    .page {
      width: min(1100px, 92vw);
      padding: 32px 0 60px;
      position: relative;
      z-index: 2;
    }
    .orb {
      position: fixed;
      border-radius: 50%;
      filter: blur(0.5px);
      opacity: 0.5;
      z-index: 1;
      animation: float 16s ease-in-out infinite;
    }
    .orb.one {
      width: 220px;
      height: 220px;
      background: rgba(226, 74, 59, 0.2);
      top: 10%;
      left: 6%;
      animation-delay: -3s;
    }
    .orb.two {
      width: 260px;
      height: 260px;
      background: rgba(46, 95, 155, 0.18);
      bottom: 8%;
      right: 4%;
      animation-delay: -6s;
    }
    @keyframes float {
      0%, 100% { transform: translateY(0px); }
      50% { transform: translateY(-18px); }
    }
    .topbar {
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      gap: 16px;
      margin-bottom: 24px;
    }
    .topbar-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      width: 100%;
      gap: 16px;
    }
    .brand {
      font-family: "Fraunces", serif;
      font-size: 28px;
      letter-spacing: 0.5px;
    }
    .brand span {
      color: var(--accent);
    }
    .tagline {
      color: var(--muted);
      font-size: 14px;
    }
    .nav {
      display: inline-flex;
      gap: 10px;
      background: rgba(255, 255, 255, 0.7);
      border-radius: 999px;
      padding: 6px;
      border: 1px solid rgba(28, 27, 25, 0.1);
      flex-wrap: wrap;
    }
    .nav a {
      text-decoration: none;
      padding: 6px 14px;
      border-radius: 999px;
      color: var(--ink);
      font-weight: 600;
      font-size: 14px;
    }
    .nav a.active {
      background: var(--accent);
      color: white;
      box-shadow: 0 10px 20px var(--ring);
    }
    .card {
      background: var(--card);
      border-radius: 18px;
      padding: 20px;
      box-shadow: 0 18px 40px rgba(15, 20, 35, 0.08);
      border: 1px solid rgba(28, 27, 25, 0.06);
      animation: rise 0.6s ease forwards;
    }
    @keyframes rise {
      from { opacity: 0; transform: translateY(12px); }
      to { opacity: 1; transform: translateY(0); }
    }
    .grid {
      display: grid;
      gap: 20px;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    }
    h1, h2, h3 {
      font-family: "Fraunces", serif;
      margin: 0 0 10px;
    }
    h1 { font-size: 34px; }
    h2 { font-size: 20px; }
    p { margin: 0 0 12px; color: var(--muted); }
    .btn {
      background: var(--accent);
      color: white;
      border: none;
      padding: 10px 16px;
      border-radius: 999px;
      font-weight: 600;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      box-shadow: 0 10px 20px var(--ring);
      transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .btn:hover { transform: translateY(-1px); box-shadow: 0 12px 26px var(--ring); }
    .btn.secondary {
      background: var(--accent-2);
      box-shadow: 0 10px 20px rgba(46, 95, 155, 0.25);
    }
    .btn.ghost {
      background: transparent;
      color: var(--ink);
      border: 1px solid rgba(28, 27, 25, 0.2);
      box-shadow: none;
    }
    .btn.small { padding: 6px 12px; font-size: 13px; }
    textarea, input, select {
      width: 100%;
      padding: 10px 12px;
      border-radius: 10px;
      border: 1px solid rgba(28, 27, 25, 0.2);
      font-family: inherit;
      font-size: 14px;
    }
    textarea { min-height: 130px; resize: vertical; }
    .flash-stack { margin-bottom: 16px; display: grid; gap: 8px; }
    .flash {
      padding: 10px 12px;
      border-radius: 12px;
      font-size: 14px;
      background: rgba(226, 74, 59, 0.1);
      color: #a13227;
      border: 1px solid rgba(226, 74, 59, 0.3);
    }
    .flash.success {
      background: rgba(46, 95, 155, 0.12);
      color: #244a76;
      border-color: rgba(46, 95, 155, 0.3);
    }
    .mailbox-item, .message-item {
      border-radius: 14px;
      border: 1px solid rgba(28, 27, 25, 0.08);
      padding: 14px;
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: center;
    }
    .mailbox-item + .mailbox-item, .message-item + .message-item {
      margin-top: 12px;
    }
    .mailbox-meta, .message-meta {
      font-size: 13px;
      color: var(--muted);
    }
    .section-title {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 12px;
    }
    .toolbar {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      align-items: center;
    }
    .message-list { display: grid; gap: 12px; }
    .message-item {
      text-decoration: none;
      color: inherit;
      transition: background 0.2s ease;
    }
    .message-item:hover { background: rgba(46, 95, 155, 0.08); }
    .message-subject { font-weight: 600; }
    .iframe-wrap {
      border-radius: 16px;
      overflow: hidden;
      border: 1px solid rgba(28, 27, 25, 0.1);
      min-height: 380px;
      background: #fff;
    }
    iframe {
      width: 100%;
      height: 540px;
      border: none;
    }
    pre.plain {
      white-space: pre-wrap;
      word-break: break-word;
      margin: 0;
      font-family: "Space Grotesk", sans-serif;
      background: #f7f7f7;
      border-radius: 14px;
      padding: 18px;
    }
    .footer {
      margin-top: 24px;
      text-align: center;
      font-size: 12px;
      color: var(--muted);
    }
    @media (max-width: 720px) {
      .topbar { align-items: stretch; }
      .topbar-row { flex-direction: column; align-items: flex-start; }
      .mailbox-item, .message-item { flex-direction: column; align-items: flex-start; }
      iframe { height: 420px; }
    }
  </style>
</head>
<body>
  <div class="orb one"></div>
  <div class="orb two"></div>
  <div class="page">
    <div class="topbar">
      <div class="topbar-row">
        <div>
          <div class="brand">Mail<span>Admin</span></div>
          <div class="tagline">Multi-mailbox control with HTML previews and share links.</div>
        </div>
        <div>
          <a class="btn ghost" href="{{ url_for('index') }}">Home</a>
        </div>
      </div>
      <div class="nav">
        <a class="{{ 'active' if active == 'mailboxes' else '' }}" href="{{ url_for('index') }}">Mailboxes</a>
        <a class="{{ 'active' if active == 'import' else '' }}" href="{{ url_for('import_page') }}">Import</a>
        <a class="{{ 'active' if active == 'shares' else '' }}" href="{{ url_for('shares_page') }}">Recent shares</a>
      </div>
    </div>
    {% with messages = get_flashed_messages(with_categories=True) %}
      {% if messages %}
      <div class="flash-stack">
        {% for category, message in messages %}
          <div class="flash {{ category }}">{{ message }}</div>
        {% endfor %}
      </div>
      {% endif %}
    {% endwith %}
    {{ content|safe }}
    <div class="footer">MailAdmin local web console</div>
  </div>
</body>
</html>
"""

INDEX_TEMPLATE = """
<div class="card">
  <div class="section-title">
    <h2>Mailboxes</h2>
    <span class="mailbox-meta">{{ mailboxes|length }} total</span>
  </div>
  {% if not mailboxes %}
    <p>No mailboxes yet.</p>
  {% else %}
    {% for box in mailboxes %}
    <div class="mailbox-item">
      <div>
        <div><strong>{{ box['address'] }}</strong></div>
        <div class="mailbox-meta">
          Auth: {{ auth_label(box) }}
          | Added: {{ format_ts(box['created_at']) }}
        </div>
      </div>
      <div style="display:flex; gap:8px; flex-wrap: wrap;">
        <a class="btn secondary small" href="{{ url_for('view_mailbox', address=box['address']) }}">Inbox</a>
        <form method="post" action="{{ url_for('delete_mailbox', address=box['address']) }}">
          <button class="btn ghost small" type="submit">Delete</button>
        </form>
      </div>
    </div>
    {% endfor %}
  {% endif %}
</div>
"""

IMPORT_TEMPLATE = """
<div class="card">
  <h2>Import mailboxes</h2>
  <p>Format per line: Address----Password----ClientID----OAuth2Token</p>
  <p>If ClientID is empty, OAuth2Token is treated as an access token.</p>
  <form method="post" action="{{ url_for('import_mailboxes') }}">
    <textarea name="payload" placeholder="user@example.com----password----client-id----refresh-token"></textarea>
    <div style="margin-top: 12px;">
      <button class="btn" type="submit">Import</button>
    </div>
  </form>
</div>
"""

SHARES_TEMPLATE = """
<div class="card">
  <div class="section-title">
    <h2>Recent shares</h2>
    <span class="mailbox-meta">{{ shares|length }} shown</span>
  </div>
  {% if not shares %}
    <p>No shares yet.</p>
  {% else %}
    {% for share in shares %}
      <div class="mailbox-item">
        <div>
          <div><strong>{{ share['subject'] or '(No subject)' }}</strong></div>
          <div class="mailbox-meta">
            {{ share['code'] }} | {{ share['mail_from'] }} | {{ format_ts(share['created_at']) }}
          </div>
        </div>
        <div style="display:flex; gap:8px; flex-wrap: wrap;">
          <a class="btn small" href="{{ url_for('view_share', code=share['code']) }}">Open</a>
          <form method="post" action="{{ url_for('delete_share', code=share['code']) }}">
            <button class="btn ghost small" type="submit">Revoke</button>
          </form>
        </div>
      </div>
    {% endfor %}
  {% endif %}
</div>
"""

MAILBOX_TEMPLATE = """
<div class="card">
  <div class="section-title">
    <h2>{{ mailbox['address'] }}</h2>
    <span class="mailbox-meta">Showing {{ messages|length }} message(s) | Inbox + Junk</span>
  </div>
  <form method="get" class="toolbar">
    <div style="display:flex; align-items:center; gap:8px;">
      <label for="limit">Limit</label>
      <input id="limit" type="number" name="limit" min="1" max="{{ max_limit }}" value="{{ limit }}">
    </div>
    <button class="btn secondary small" type="submit">Refresh</button>
    <a class="btn ghost small" href="{{ url_for('index') }}">Back</a>
  </form>
</div>
<div class="card" style="margin-top: 16px;">
  {% if error %}
    <p>{{ error }}</p>
  {% elif not messages %}
    <p>No messages found.</p>
  {% else %}
    <div class="message-list">
      {% for msg in messages %}
      <a class="message-item" href="{{ url_for('view_message', address=mailbox['address'], uid=msg['uid'], folder=msg['folder']) }}">
        <div>
          <div class="message-subject">{{ msg['subject'] or '(No subject)' }}</div>
          <div class="message-meta">{{ msg['mail_from'] }} | {{ msg['mail_dt'] }} | {{ msg['folder_label'] }}</div>
        </div>
        <div class="message-meta">View</div>
      </a>
      {% endfor %}
    </div>
  {% endif %}
</div>
"""

MESSAGE_TEMPLATE = """
<div class="card">
  <div class="section-title">
    <h2>{{ message['subject'] or '(No subject)' }}</h2>
    <span class="mailbox-meta">{{ mailbox['address'] }}</span>
  </div>
  <div class="mailbox-meta">
    From: {{ message['mail_from'] or '-' }}<br>
    To: {{ message['mail_to'] or '-' }}<br>
    Date: {{ message['mail_dt'] or '-' }}<br>
    Folder: {{ message['folder_label'] or '-' }}
  </div>
  <div style="display:flex; gap:8px; flex-wrap: wrap; margin-top: 12px;">
    <form method="post" action="{{ url_for('share_message', address=mailbox['address'], uid=message['uid'], folder=message['folder']) }}">
      <button class="btn secondary small" type="submit">Create share link</button>
    </form>
    <a class="btn ghost small" href="{{ url_for('view_mailbox', address=mailbox['address']) }}">Back</a>
  </div>
</div>
<div class="card" style="margin-top: 16px;">
  {% if message['safe_body_html'] %}
    <div class="iframe-wrap">
      <iframe sandbox="allow-same-origin" referrerpolicy="no-referrer" srcdoc="{{ message['safe_body_html'] | e }}"></iframe>
    </div>
  {% else %}
    <pre class="plain">{{ message['body_text'] }}</pre>
  {% endif %}
</div>
"""

SHARE_TEMPLATE = """
<div class="card">
  <div class="section-title">
    <h2>{{ share['subject'] or '(No subject)' }}</h2>
    <span class="mailbox-meta">Shared mail</span>
  </div>
  <div class="mailbox-meta">
    From: {{ share['mail_from'] or '-' }}<br>
    To: {{ share['mail_to'] or '-' }}<br>
    Date: {{ share['mail_dt'] or '-' }}<br>
    Code: {{ share['code'] }}
  </div>
</div>
<div class="card" style="margin-top: 16px;">
  <div class="iframe-wrap">
    <iframe sandbox="allow-same-origin" referrerpolicy="no-referrer" srcdoc="{{ share['body_html'] | e }}"></iframe>
  </div>
</div>
"""


class MailError(Exception):
    pass


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS mailboxes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                address TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL DEFAULT '',
                client_id TEXT NOT NULL DEFAULT '',
                refresh_token TEXT NOT NULL DEFAULT '',
                created_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS shares (
                code TEXT PRIMARY KEY,
                subject TEXT,
                mail_from TEXT,
                mail_to TEXT,
                mail_dt TEXT,
                body_html TEXT,
                created_at INTEGER NOT NULL
            )
            """
        )


def render_page(title, template, **context):
    content = render_template_string(template, **context)
    active = context.pop("active", "")
    return render_template_string(
        BASE_TEMPLATE, title=title, content=content, active=active
    )


def api_ok(data=None, status=200):
    payload = {"ok": True, "data": data}
    return jsonify(payload), status


def api_error(message, status=400):
    payload = {"ok": False, "error": message}
    return jsonify(payload), status


def row_to_dict(row):
    if row is None:
        return None
    return dict(row)


def is_port_open(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.5)
        return sock.connect_ex(("127.0.0.1", port)) == 0


def is_mailadmin_running(port):
    url = f"http://127.0.0.1:{port}/api/health"
    try:
        resp = requests.get(url, timeout=1)
        payload = resp.json()
    except Exception:
        return False
    return resp.status_code == 200 and payload.get("ok") and payload.get("data", {}).get("status") == "ok"


def maybe_open_browser(port):
    if os.environ.get("MAILADMIN_NO_BROWSER") == "1":
        return
    url = f"http://127.0.0.1:{port}/"
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()


def format_ts(ts):
    if not ts:
        return "-"
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ts))


def normalize_folder(name):
    if not name:
        return "INBOX"
    return name


def folder_label(name):
    if name and name.upper() == "INBOX":
        return "Inbox"
    return "Junk"


def auth_label(box):
    if box["refresh_token"]:
        if box["client_id"]:
            return "OAuth (refresh)"
        return "OAuth (token)"
    if box["password"]:
        return "Password"
    return "None"


def parse_import_payload(payload):
    results = []
    errors = []
    for idx, line in enumerate(payload.splitlines(), start=1):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split("----")
        if len(parts) != 4:
            errors.append(f"Line {idx}: expected 4 fields.")
            continue
        address, password, client_id, refresh_token = [p.strip() for p in parts]
        if not address:
            errors.append(f"Line {idx}: address missing.")
            continue
        results.append((address, password, client_id, refresh_token))
    return results, errors


def get_access_token(client_id, refresh_token):
    url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
    }
    try:
        response = requests.post(url, data=data, timeout=10)
        response.raise_for_status()
        payload = response.json()
    except requests.RequestException as exc:
        raise MailError(f"Token request failed: {exc}") from exc
    except ValueError as exc:
        raise MailError("Token response invalid JSON.") from exc

    error_description = payload.get("error_description")
    error = payload.get("error")
    if error_description or error:
        detail = error_description or error
        raise MailError(f"Token error: {detail}")

    access_token = payload.get("access_token")
    if not access_token:
        raise MailError("Access token missing.")
    return access_token


def generate_auth_string(email_name, access_token):
    auth_string = f"user={email_name}\x01auth=Bearer {access_token}\x01\x01"
    return auth_string.encode("utf-8")


def _decode_payload(payload, charset):
    if isinstance(payload, bytes):
        codec = charset or "utf-8"
        try:
            return payload.decode(codec, errors="ignore")
        except LookupError:
            return payload.decode("utf-8", errors="ignore")
    if isinstance(payload, str):
        return payload
    return ""


def decode_header_value(value):
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def parse_date(value):
    if not value:
        return "", 0
    try:
        dt = parsedate_to_datetime(value)
        if not dt:
            return "", 0
        if dt.tzinfo:
            local_dt = dt.astimezone()
        else:
            local_dt = dt
        return local_dt.strftime("%Y-%m-%d %H:%M:%S"), int(local_dt.timestamp())
    except Exception:
        return "", 0


def connect_mailbox(mailbox):
    mail = imaplib.IMAP4_SSL(IMAP_HOST)
    if mailbox["refresh_token"]:
        if mailbox["client_id"]:
            access_token = get_access_token(mailbox["client_id"], mailbox["refresh_token"])
        else:
            access_token = mailbox["refresh_token"]
        mail.authenticate("XOAUTH2", lambda _: generate_auth_string(mailbox["address"], access_token))
    elif mailbox["password"]:
        mail.login(mailbox["address"], mailbox["password"])
    else:
        raise MailError("No authentication data configured.")
    return mail


def extract_message(raw_email):
    email_message = email.message_from_bytes(raw_email)
    subject = decode_header_value(email_message.get("Subject"))
    mail_from = decode_header_value(email_message.get("From"))
    mail_to = decode_header_value(email_message.get("To"))
    mail_dt, mail_ts = parse_date(email_message.get("Date"))

    html_parts = []
    text_parts = []
    if email_message.is_multipart():
        for part in email_message.walk():
            if part.get_content_maintype() == "multipart":
                continue
            content_type = part.get_content_type()
            payload = part.get_payload(decode=True)
            if content_type == "text/html":
                html_parts.append(_decode_payload(payload, part.get_content_charset()))
            elif content_type == "text/plain":
                text_parts.append(_decode_payload(payload, part.get_content_charset()))
    else:
        payload = email_message.get_payload(decode=True)
        html_parts.append(_decode_payload(payload, email_message.get_content_charset()))

    body_html = "".join(html_parts).strip()
    body_text = "\n".join(text_parts).strip()
    return {
        "subject": subject,
        "mail_from": mail_from,
        "mail_to": mail_to,
        "mail_dt": mail_dt,
        "mail_ts": mail_ts,
        "body_html": body_html,
        "body_text": body_text,
    }


def sanitize_html(value):
    if not value:
        return ""
    cleaned = re.sub(r"(?is)<script[^>]*>.*?</script>", "", value)
    cleaned = re.sub(r"(?is)<base[^>]*>", "", cleaned)
    cleaned = re.sub(r"(?is)on\w+\s*=\s*(['\"]).*?\1", "", cleaned)
    cleaned = re.sub(r"(?is)javascript:", "", cleaned)
    return cleaned


def list_messages(mailbox, limit):
    mail = None
    try:
        mail = connect_mailbox(mailbox)
        messages = []
        folders = ["INBOX"] + JUNK_FOLDERS
        for folder in folders:
            try:
                status, _ = mail.select(folder, readonly=True)
            except imaplib.IMAP4.error:
                continue
            if status != "OK":
                continue

            status, data = mail.uid("search", None, "ALL")
            if status != "OK":
                continue

            uids = data[0].split()
            if not uids:
                continue
            uids = uids[-limit:]
            uids.reverse()

            for uid in uids:
                uid_str = uid.decode() if isinstance(uid, bytes) else str(uid)
                status, msg_data = mail.uid("fetch", uid_str, "(RFC822)")
                if status != "OK" or not msg_data:
                    continue
                raw_email = None
                for part in msg_data:
                    if isinstance(part, tuple) and len(part) > 1:
                        raw_email = part[1]
                        break
                if not raw_email:
                    continue
                parsed = extract_message(raw_email)
                parsed["uid"] = uid_str
                parsed["folder"] = folder
                parsed["folder_label"] = folder_label(folder)
                messages.append(parsed)

        messages.sort(key=lambda item: item.get("mail_ts", 0), reverse=True)
        return messages[:limit]
    finally:
        if mail is not None:
            try:
                mail.logout()
            except Exception:
                pass


def fetch_message(mailbox, uid, folder=None):
    mail = None
    try:
        mail = connect_mailbox(mailbox)
        folder_name = normalize_folder(folder)
        try:
            status, _ = mail.select(folder_name, readonly=True)
        except imaplib.IMAP4.error as exc:
            raise MailError(f"Mailbox select failed: {exc}") from exc
        if status != "OK":
            raise MailError("Mailbox select failed.")
        status, msg_data = mail.uid("fetch", uid, "(RFC822)")
        if status != "OK" or not msg_data:
            raise MailError("Message fetch failed.")
        raw_email = None
        for part in msg_data:
            if isinstance(part, tuple) and len(part) > 1:
                raw_email = part[1]
                break
        if not raw_email:
            raise MailError("Message empty.")
        parsed = extract_message(raw_email)
        parsed["uid"] = uid
        parsed["folder"] = folder_name
        parsed["folder_label"] = folder_label(folder_name)
        return parsed
    finally:
        if mail is not None:
            try:
                mail.logout()
            except Exception:
                pass


def build_share_body(message):
    if message["body_html"]:
        return sanitize_html(message["body_html"])
    if message["body_text"]:
        return "<pre>" + html.escape(message["body_text"]) + "</pre>"
    return "<p>(empty)</p>"


def generate_share_code(conn):
    alphabet = string.ascii_letters + string.digits
    for _ in range(20):
        code = "".join(secrets.choice(alphabet) for _ in range(SHARE_CODE_LEN))
        if not (any(c.isalpha() for c in code) and any(c.isdigit() for c in code)):
            continue
        exists = conn.execute("SELECT 1 FROM shares WHERE code = ?", (code,)).fetchone()
        if not exists:
            return code
    raise MailError("Unable to generate share code.")


@APP.after_request
def set_security_headers(response):
    response.headers["Content-Security-Policy"] = (
        "default-src 'none'; "
        "img-src https: data:; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
        "font-src https://fonts.gstatic.com; "
        "form-action 'self'; "
        "frame-ancestors 'none'; "
        "base-uri 'none'"
    )
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    return response


@APP.route("/")
def index():
    with get_db() as conn:
        mailboxes = conn.execute(
            "SELECT * FROM mailboxes ORDER BY created_at DESC"
        ).fetchall()
    return render_page(
        "Mailboxes - MailAdmin",
        INDEX_TEMPLATE,
        mailboxes=mailboxes,
        format_ts=format_ts,
        auth_label=auth_label,
        active="mailboxes",
    )


@APP.route("/import")
def import_page():
    return render_page(
        "Import - MailAdmin",
        IMPORT_TEMPLATE,
        active="import",
    )


@APP.route("/shares")
def shares_page():
    with get_db() as conn:
        shares = conn.execute(
            "SELECT * FROM shares ORDER BY created_at DESC"
        ).fetchall()
    return render_page(
        "Shares - MailAdmin",
        SHARES_TEMPLATE,
        shares=shares,
        format_ts=format_ts,
        active="shares",
    )


@APP.get("/api/health")
def api_health():
    return api_ok({"status": "ok", "time": int(time.time())})


@APP.get("/api/mailboxes")
def api_list_mailboxes():
    with get_db() as conn:
        mailboxes = conn.execute(
            "SELECT * FROM mailboxes ORDER BY created_at DESC"
        ).fetchall()
    payload = [row_to_dict(row) for row in mailboxes]
    return api_ok(payload)


@APP.post("/api/mailboxes")
def api_import_mailboxes():
    data = request.get_json(silent=True) or {}
    items = data.get("items")
    entries = []
    errors = []
    if isinstance(items, list):
        for idx, item in enumerate(items, start=1):
            if not isinstance(item, dict):
                errors.append(f"Item {idx}: expected object.")
                continue
            address = (item.get("address") or "").strip()
            password = (item.get("password") or "").strip()
            client_id = (item.get("client_id") or "").strip()
            refresh_token = (item.get("refresh_token") or item.get("token") or "").strip()
            if not address:
                errors.append(f"Item {idx}: address missing.")
                continue
            entries.append((address, password, client_id, refresh_token))
    else:
        payload = request.form.get("payload")
        if payload is None:
            payload = request.get_data(as_text=True) or ""
        entries, errors = parse_import_payload(payload)

    if errors and not entries:
        return api_error({"errors": errors}, status=400)

    now = int(time.time())
    imported = 0
    with get_db() as conn:
        for address, password, client_id, refresh_token in entries:
            existing = conn.execute(
                "SELECT id FROM mailboxes WHERE address = ?", (address,)
            ).fetchone()
            if existing:
                conn.execute(
                    """
                    UPDATE mailboxes
                    SET password = ?, client_id = ?, refresh_token = ?, updated_at = ?
                    WHERE address = ?
                    """,
                    (password, client_id, refresh_token, now, address),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO mailboxes (address, password, client_id, refresh_token, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (address, password, client_id, refresh_token, now, now),
                )
            imported += 1
    return api_ok({"imported": imported, "errors": errors})


@APP.delete("/api/mailboxes/<path:address>")
def api_delete_mailbox(address):
    with get_db() as conn:
        conn.execute("DELETE FROM mailboxes WHERE address = ?", (address,))
    return api_ok({"deleted": address})


@APP.get("/api/mailboxes/<path:address>/messages")
def api_list_messages(address):
    limit_raw = request.args.get("limit", str(DEFAULT_LIMIT))
    try:
        limit = int(limit_raw)
    except ValueError:
        limit = DEFAULT_LIMIT
    limit = max(1, min(MAX_LIMIT, limit))
    with get_db() as conn:
        mailbox = conn.execute(
            "SELECT * FROM mailboxes WHERE address = ?", (address,)
        ).fetchone()
    if not mailbox:
        return api_error("Mailbox not found.", status=404)
    try:
        messages = list_messages(mailbox, limit)
    except MailError as exc:
        return api_error(str(exc), status=500)
    payload = []
    for msg in messages:
        payload.append(
            {
                "uid": msg.get("uid"),
                "subject": msg.get("subject"),
                "mail_from": msg.get("mail_from"),
                "mail_to": msg.get("mail_to"),
                "mail_dt": msg.get("mail_dt"),
                "mail_ts": msg.get("mail_ts"),
                "folder": msg.get("folder"),
                "folder_label": msg.get("folder_label"),
            }
        )
    return api_ok(payload)


@APP.get("/api/mailboxes/<path:address>/message/<uid>")
def api_get_message(address, uid):
    with get_db() as conn:
        mailbox = conn.execute(
            "SELECT * FROM mailboxes WHERE address = ?", (address,)
        ).fetchone()
    if not mailbox:
        return api_error("Mailbox not found.", status=404)
    folder = request.args.get("folder")
    try:
        message = fetch_message(mailbox, uid, folder=folder)
    except MailError as exc:
        return api_error(str(exc), status=500)
    message["safe_body_html"] = sanitize_html(message["body_html"])
    return api_ok(message)


@APP.post("/api/mailboxes/<path:address>/message/<uid>/share")
def api_share_message(address, uid):
    with get_db() as conn:
        mailbox = conn.execute(
            "SELECT * FROM mailboxes WHERE address = ?", (address,)
        ).fetchone()
    if not mailbox:
        return api_error("Mailbox not found.", status=404)
    folder = request.args.get("folder")
    try:
        message = fetch_message(mailbox, uid, folder=folder)
    except MailError as exc:
        return api_error(str(exc), status=500)

    with get_db() as conn:
        code = generate_share_code(conn)
        conn.execute(
            """
            INSERT INTO shares (code, subject, mail_from, mail_to, mail_dt, body_html, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                code,
                message["subject"],
                message["mail_from"],
                message["mail_to"],
                message["mail_dt"],
                build_share_body(message),
                int(time.time()),
            ),
        )

    share_url = request.host_url.rstrip("/") + url_for("view_share", code=code)
    return api_ok({"code": code, "url": share_url})


@APP.get("/api/shares")
def api_list_shares():
    with get_db() as conn:
        shares = conn.execute(
            "SELECT * FROM shares ORDER BY created_at DESC"
        ).fetchall()
    payload = [row_to_dict(row) for row in shares]
    return api_ok(payload)


@APP.get("/api/shares/<code>")
def api_get_share(code):
    with get_db() as conn:
        share = conn.execute(
            "SELECT * FROM shares WHERE code = ?", (code,)
        ).fetchone()
    if not share:
        return api_error("Share not found.", status=404)
    return api_ok(row_to_dict(share))


@APP.delete("/api/shares/<code>")
def api_delete_share(code):
    with get_db() as conn:
        conn.execute("DELETE FROM shares WHERE code = ?", (code,))
    return api_ok({"deleted": code})


@APP.post("/import")
def import_mailboxes():
    payload = request.form.get("payload", "").strip()
    if not payload:
        flash("Import payload is empty.", "error")
        return redirect(url_for("index"))

    entries, errors = parse_import_payload(payload)
    if errors:
        for err in errors[:3]:
            flash(err, "error")
        if len(errors) > 3:
            flash(f"{len(errors) - 3} more errors hidden.", "error")
    if not entries:
        return redirect(url_for("index"))

    now = int(time.time())
    imported = 0
    with get_db() as conn:
        for address, password, client_id, refresh_token in entries:
            existing = conn.execute(
                "SELECT id FROM mailboxes WHERE address = ?", (address,)
            ).fetchone()
            if existing:
                conn.execute(
                    """
                    UPDATE mailboxes
                    SET password = ?, client_id = ?, refresh_token = ?, updated_at = ?
                    WHERE address = ?
                    """,
                    (password, client_id, refresh_token, now, address),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO mailboxes (address, password, client_id, refresh_token, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (address, password, client_id, refresh_token, now, now),
                )
            imported += 1
    flash(f"Imported {imported} mailbox(es).", "success")
    return redirect(url_for("import_page"))


@APP.post("/mailbox/<path:address>/delete")
def delete_mailbox(address):
    with get_db() as conn:
        conn.execute("DELETE FROM mailboxes WHERE address = ?", (address,))
    flash("Mailbox deleted.", "success")
    return redirect(url_for("index"))


@APP.route("/mailbox/<path:address>")
def view_mailbox(address):
    limit_raw = request.args.get("limit", str(DEFAULT_LIMIT))
    try:
        limit = int(limit_raw)
    except ValueError:
        limit = DEFAULT_LIMIT
    limit = max(1, min(MAX_LIMIT, limit))

    with get_db() as conn:
        mailbox = conn.execute(
            "SELECT * FROM mailboxes WHERE address = ?", (address,)
        ).fetchone()
    if not mailbox:
        abort(404)

    error = None
    messages = []
    try:
        messages = list_messages(mailbox, limit)
    except MailError as exc:
        error = str(exc)

    return render_page(
        f"Inbox - {mailbox['address']}",
        MAILBOX_TEMPLATE,
        mailbox=mailbox,
        messages=messages,
        error=error,
        limit=limit,
        max_limit=MAX_LIMIT,
        active="mailboxes",
    )


@APP.route("/mailbox/<path:address>/message/<uid>")
def view_message(address, uid):
    with get_db() as conn:
        mailbox = conn.execute(
            "SELECT * FROM mailboxes WHERE address = ?", (address,)
        ).fetchone()
    if not mailbox:
        abort(404)

    try:
        folder = request.args.get("folder")
        message = fetch_message(mailbox, uid, folder=folder)
    except MailError as exc:
        flash(str(exc), "error")
        return redirect(url_for("view_mailbox", address=address))

    message["safe_body_html"] = sanitize_html(message["body_html"])
    return render_page(
        f"Message - {mailbox['address']}",
        MESSAGE_TEMPLATE,
        mailbox=mailbox,
        message=message,
        active="mailboxes",
    )


@APP.post("/mailbox/<path:address>/message/<uid>/share")
def share_message(address, uid):
    with get_db() as conn:
        mailbox = conn.execute(
            "SELECT * FROM mailboxes WHERE address = ?", (address,)
        ).fetchone()
    if not mailbox:
        abort(404)

    try:
        folder = request.args.get("folder")
        message = fetch_message(mailbox, uid, folder=folder)
    except MailError as exc:
        flash(str(exc), "error")
        return redirect(url_for("view_mailbox", address=address))

    with get_db() as conn:
        code = generate_share_code(conn)
        conn.execute(
            """
            INSERT INTO shares (code, subject, mail_from, mail_to, mail_dt, body_html, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                code,
                message["subject"],
                message["mail_from"],
                message["mail_to"],
                message["mail_dt"],
                build_share_body(message),
                int(time.time()),
            ),
        )

    flash("Share link created.", "success")
    return redirect(url_for("view_share", code=code))


@APP.route("/share/<code>")
def view_share(code):
    with get_db() as conn:
        share = conn.execute(
            "SELECT * FROM shares WHERE code = ?", (code,)
        ).fetchone()
    if not share:
        abort(404)
    return render_page(
        f"Share - {share['code']}",
        SHARE_TEMPLATE,
        share=share,
        active="shares",
    )


@APP.post("/share/<code>/delete")
def delete_share(code):
    with get_db() as conn:
        conn.execute("DELETE FROM shares WHERE code = ?", (code,))
    flash("Share revoked.", "success")
    return redirect(url_for("index"))


def main():
    init_db()
    host = os.environ.get("MAILADMIN_HOST", "0.0.0.0")
    port = int(os.environ.get("MAILADMIN_PORT", "5000"))
    if is_port_open(port):
        if is_mailadmin_running(port):
            print(f"MailAdmin already running at http://127.0.0.1:{port}/")
            maybe_open_browser(port)
            return
        print(f"Port {port} is already in use. Set MAILADMIN_PORT to another value.")
        sys.exit(1)
    maybe_open_browser(port)
    APP.run(host=host, port=port, debug=False)


if __name__ == "__main__":
    main()
