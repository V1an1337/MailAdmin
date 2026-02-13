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

def load_env_file(path):
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as handle:
            for line in handle:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = value
    except OSError:
        pass


load_env_file(".env")


APP = Flask(__name__)
APP.secret_key = os.environ.get("MAILADMIN_SECRET", secrets.token_hex(16))

#DB_PATH = os.path.join(os.path.dirname(__file__), "mailadmin.db")
DB_PATH = "mailadmin.db"
IMAP_HOST = os.environ.get("MAILADMIN_IMAP_HOST", "outlook.live.com")
DEFAULT_LIMIT = 10
MAX_LIMIT = 50
SHARE_CODE_LEN = 8
MULTI_USER = os.environ.get("MAILADMIN_MULTI_USER", "0") == "1"
API_KEY_COOKIE = "api_key"
PUBLIC_BASE_URL = os.environ.get("MAILADMIN_PUBLIC_BASE_URL", "https://oauth.v1an.xyz").rstrip("/")
OPENAPI_VERSION = "1.0.0"
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
    .modal-backdrop {
      position: fixed;
      inset: 0;
      background: rgba(20, 20, 20, 0.55);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 99;
      padding: 24px;
    }
    .modal-card {
      background: var(--card);
      border-radius: 18px;
      padding: 24px;
      max-width: 520px;
      width: min(520px, 92vw);
      box-shadow: 0 18px 40px rgba(15, 20, 35, 0.2);
      border: 1px solid rgba(28, 27, 25, 0.08);
      text-align: left;
    }
    .key-box {
      font-family: "Space Grotesk", sans-serif;
      font-size: 16px;
      font-weight: 600;
      background: #f7f4ee;
      border: 1px dashed rgba(28, 27, 25, 0.2);
      border-radius: 12px;
      padding: 12px;
      word-break: break-all;
      margin: 12px 0 16px;
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
        <a class="{{ 'active' if active == 'docs' else '' }}" href="{{ url_for('api_docs_page') }}">API Docs</a>
        <a class="{{ 'active' if active == 'account' else '' }}" href="{{ url_for('account_page') }}">Account</a>
        {% if show_logout %}
        <a href="{{ url_for('logout_page') }}">Logout</a>
        {% endif %}
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
  <script>
    (function () {
      var key = window.localStorage ? localStorage.getItem("api_key") : "";
      if (!key) {
        return;
      }
      if (document.cookie.indexOf("{{ api_key_cookie }}=") === -1) {
        document.cookie = "{{ api_key_cookie }}=" + encodeURIComponent(key) + "; path=/; max-age=31536000; SameSite=Lax";
        if (!sessionStorage.getItem("api_key_synced")) {
          sessionStorage.setItem("api_key_synced", "1");
          window.location.reload();
        }
      }
    })();
  </script>
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

LOGIN_TEMPLATE = """
<div class="card">
  <h2>Login or Register</h2>
  <p>{{ message or 'Please login with your API key or register to get a new one.' }}</p>
  <div style="display:grid; gap:12px;">
    <input id="apiKeyInput" placeholder="API Key">
    <div style="display:flex; gap:8px; flex-wrap: wrap;">
      <button class="btn" type="button" id="loginBtn">Login</button>
      <button class="btn secondary" type="button" id="registerBtn">Register</button>
    </div>
    <div id="loginStatus" class="mailbox-meta"></div>
  </div>
</div>
<div id="keyModal" class="modal-backdrop" style="display:none;">
  <div class="modal-card">
    <h2>Save your API key</h2>
    <p>Please save this key safely. You will need it to log in.</p>
    <div class="key-box" id="keyBox"></div>
    <button class="btn" type="button" id="confirmKeyBtn">I have saved it</button>
  </div>
</div>
<script>
  (function () {
    var input = document.getElementById("apiKeyInput");
    var status = document.getElementById("loginStatus");
    var keyModal = document.getElementById("keyModal");
    var keyBox = document.getElementById("keyBox");
    var confirmBtn = document.getElementById("confirmKeyBtn");
    var pendingKey = "";
    var stored = window.localStorage ? localStorage.getItem("api_key") : "";
    if (stored) {
      input.value = stored;
    }
    function saveKey(key) {
      if (window.localStorage) {
        localStorage.setItem("api_key", key);
      }
      document.cookie = "{{ api_key_cookie }}=" + encodeURIComponent(key) + "; path=/; max-age=31536000; SameSite=Lax";
    }
    async function login() {
      var key = input.value.trim();
      if (!key) {
        status.textContent = "API key is required.";
        return;
      }
      status.textContent = "Checking...";
      var resp = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ api_key: key })
      });
      var data = await resp.json().catch(function () { return {}; });
      if (!data.ok) {
        status.textContent = data.error || "Login failed.";
        return;
      }
      saveKey(key);
      status.textContent = "Logged in.";
      window.location.href = "/";
    }
    async function register() {
      status.textContent = "Creating API key...";
      var resp = await fetch("/api/auth/register", { method: "POST" });
      var data = await resp.json().catch(function () { return {}; });
      if (!data.ok) {
        status.textContent = data.error || "Register failed.";
        return;
      }
      pendingKey = data.data.api_key;
      keyBox.textContent = pendingKey;
      keyModal.style.display = "flex";
    }
    confirmBtn.addEventListener("click", function () {
      if (!pendingKey) {
        return;
      }
      saveKey(pendingKey);
      window.location.href = "/";
    });
    document.getElementById("loginBtn").addEventListener("click", login);
    document.getElementById("registerBtn").addEventListener("click", register);
  })();
</script>
"""

ACCOUNT_TEMPLATE = """
<div class="card">
  <h2>Account</h2>
  <p>Rotate your API key to invalidate the old one. Save the new key safely.</p>
  <div style="display:flex; gap:8px; flex-wrap: wrap;">
    <button class="btn secondary" type="button" id="rotateBtn">Rotate API Key</button>
  </div>
  <div id="rotateStatus" class="mailbox-meta" style="margin-top: 12px;"></div>
</div>
<div id="rotateModal" class="modal-backdrop" style="display:none;">
  <div class="modal-card">
    <h2>New API key</h2>
    <p>Save this new key safely. The old key will stop working.</p>
    <div class="key-box" id="rotateKeyBox"></div>
    <button class="btn" type="button" id="confirmRotateBtn">I have saved it</button>
  </div>
</div>
<script>
  (function () {
    var status = document.getElementById("rotateStatus");
    var modal = document.getElementById("rotateModal");
    var keyBox = document.getElementById("rotateKeyBox");
    var confirmBtn = document.getElementById("confirmRotateBtn");
    var pendingKey = "";
    function saveKey(key) {
      if (window.localStorage) {
        localStorage.setItem("api_key", key);
      }
      document.cookie = "{{ api_key_cookie }}=" + encodeURIComponent(key) + "; path=/; max-age=31536000; SameSite=Lax";
    }
    async function rotate() {
      status.textContent = "Rotating...";
      var key = window.localStorage ? localStorage.getItem("api_key") : "";
      var headers = { "Content-Type": "application/json" };
      if (key) {
        headers["X-API-Key"] = key;
      }
      var resp = await fetch("/api/auth/rotate", {
        method: "POST",
        headers: headers
      });
      var data = await resp.json().catch(function () { return {}; });
      if (!data.ok) {
        status.textContent = data.error || "Rotate failed.";
        return;
      }
      pendingKey = data.data.api_key;
      keyBox.textContent = pendingKey;
      modal.style.display = "flex";
    }
    confirmBtn.addEventListener("click", function () {
      if (!pendingKey) {
        return;
      }
      saveKey(pendingKey);
      window.location.href = "/";
    });
    document.getElementById("rotateBtn").addEventListener("click", rotate);
  })();
</script>
"""

LOGOUT_TEMPLATE = """
<div class="card">
  <h2>Logging out...</h2>
  <p>Clearing local credentials.</p>
</div>
<script>
  (function () {
    if (window.localStorage) {
      localStorage.removeItem("api_key");
    }
    document.cookie = "{{ api_key_cookie }}=; path=/; max-age=0; SameSite=Lax";
    window.location.href = "/login";
  })();
</script>
"""


DOCS_TEMPLATE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>MailAdmin API Docs</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/swagger-ui-dist@5/swagger-ui.css">
  <style>
    body { margin: 0; background: #faf7f0; }
    .top {
      padding: 12px 16px;
      font-family: "Space Grotesk", sans-serif;
      background: #1c1b19;
      color: #ffffff;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
    }
    .top a {
      color: #ffffff;
      text-decoration: none;
      border: 1px solid rgba(255, 255, 255, 0.35);
      padding: 6px 10px;
      border-radius: 999px;
      font-size: 13px;
    }
  </style>
</head>
<body>
  <div class="top">
    <div>MailAdmin API Documentation</div>
    <a href="{{ home_url }}">Back to App</a>
  </div>
  <div id="swagger-ui"></div>
  <script src="https://cdn.jsdelivr.net/npm/swagger-ui-dist@5/swagger-ui-bundle.js"></script>
  <script>
    window.ui = SwaggerUIBundle({
      url: "{{ spec_url }}",
      dom_id: "#swagger-ui",
      deepLinking: true,
      displayRequestDuration: true,
      persistAuthorization: true,
      docExpansion: "list"
    });
  </script>
</body>
</html>
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
                owner_key TEXT NOT NULL DEFAULT '',
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
                owner_key TEXT NOT NULL DEFAULT '',
                created_at INTEGER NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                api_key TEXT PRIMARY KEY,
                created_at INTEGER NOT NULL,
                last_rotated_at INTEGER NOT NULL
            )
            """
        )
        ensure_column(conn, "mailboxes", "owner_key", "TEXT NOT NULL DEFAULT ''")
        ensure_column(conn, "shares", "owner_key", "TEXT NOT NULL DEFAULT ''")


def ensure_column(conn, table, column, column_type):
    rows = conn.execute(f"PRAGMA table_info({table})").fetchall()
    names = {row[1] for row in rows}
    if column not in names:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {column_type}")


def render_page(title, template, **context):
    context.setdefault("api_key_cookie", API_KEY_COOKIE)
    context.setdefault("show_logout", is_logged_in())
    content = render_template_string(template, **context)
    active = context.pop("active", "")
    return render_template_string(
        BASE_TEMPLATE,
        title=title,
        content=content,
        active=active,
        api_key_cookie=API_KEY_COOKIE,
        show_logout=context.get("show_logout", False),
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


def get_api_key():
    auth_header = request.headers.get("Authorization", "")
    if auth_header.lower().startswith("bearer "):
        return auth_header.split(" ", 1)[1].strip()
    key = request.headers.get("X-API-Key")
    if key:
        return key.strip()
    key = request.args.get("api_key") or request.cookies.get(API_KEY_COOKIE)
    return key.strip() if key else ""


def is_logged_in():
    if not MULTI_USER:
        return False
    api_key = get_api_key()
    return is_valid_api_key(api_key)


def is_valid_api_key(api_key):
    if not api_key:
        return False
    with get_db() as conn:
        row = conn.execute(
            "SELECT api_key FROM users WHERE api_key = ?", (api_key,)
        ).fetchone()
    return row is not None


def generate_api_key(conn):
    for _ in range(20):
        key = secrets.token_hex(16)
        exists = conn.execute(
            "SELECT 1 FROM users WHERE api_key = ?", (key,)
        ).fetchone()
        if not exists:
            return key
    raise MailError("Failed to generate API key.")


def require_user(api=False):
    if not MULTI_USER:
        return ""
    api_key = get_api_key()
    if not api_key:
        return api_error("Unauthorized", status=401) if api else None
    if not is_valid_api_key(api_key):
        return api_error("Invalid API key.", status=401) if api else None
    return api_key


def render_login_page(message=None):
    return render_page(
        "Login - MailAdmin",
        LOGIN_TEMPLATE,
        message=message,
        active="account",
    )


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


def openapi_servers():
    host_url = ""
    try:
        host_url = request.host_url.rstrip("/")
    except RuntimeError:
        host_url = ""
    servers = []
    seen = set()
    for candidate in [host_url, "http://127.0.0.1:5000", PUBLIC_BASE_URL]:
        if not candidate:
            continue
        url = candidate.rstrip("/")
        if url and url not in seen:
            seen.add(url)
            servers.append({"url": url})
    return servers


def build_openapi_spec():
    auth_security = [{"ApiKeyAuth": []}, {"BearerAuth": []}]
    return {
        "openapi": "3.0.3",
        "info": {
            "title": "MailAdmin API",
            "version": OPENAPI_VERSION,
            "description": (
                "MailAdmin public API. In multi-user mode, authenticated endpoints "
                "require X-API-Key or Bearer token."
            ),
        },
        "servers": openapi_servers(),
        "tags": [
            {"name": "Health"},
            {"name": "Auth"},
            {"name": "Mailboxes"},
            {"name": "Messages"},
            {"name": "Shares"},
        ],
        "paths": {
            "/api/health": {
                "get": {
                    "tags": ["Health"],
                    "summary": "Health check",
                    "operationId": "getHealth",
                    "responses": {"200": {"description": "OK", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}}},
                }
            },
            "/api/auth/login": {
                "post": {
                    "tags": ["Auth"],
                    "summary": "Validate API key",
                    "operationId": "authLogin",
                    "requestBody": {
                        "required": True,
                        "content": {"application/json": {"schema": {"$ref": "#/components/schemas/AuthLoginRequest"}}},
                    },
                    "responses": {
                        "200": {"description": "Login success", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}},
                        "401": {"description": "Invalid key", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiError"}}}},
                    },
                }
            },
            "/api/auth/register": {
                "post": {
                    "tags": ["Auth"],
                    "summary": "Register and return new API key",
                    "operationId": "authRegister",
                    "responses": {
                        "200": {"description": "Register success", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}},
                        "400": {"description": "Multi-user disabled", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiError"}}}},
                    },
                }
            },
            "/api/auth/rotate": {
                "post": {
                    "tags": ["Auth"],
                    "summary": "Rotate API key (24h cooldown)",
                    "operationId": "authRotate",
                    "security": auth_security,
                    "responses": {
                        "200": {"description": "Rotate success", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}},
                        "401": {"description": "Unauthorized", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiError"}}}},
                        "429": {"description": "Cooldown", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiError"}}}},
                    },
                }
            },
            "/api/mailboxes": {
                "get": {
                    "tags": ["Mailboxes"],
                    "summary": "List mailboxes",
                    "operationId": "listMailboxes",
                    "security": auth_security,
                    "responses": {"200": {"description": "Mailbox list", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}}},
                },
                "post": {
                    "tags": ["Mailboxes"],
                    "summary": "Import/update mailboxes",
                    "operationId": "importMailboxes",
                    "security": auth_security,
                    "requestBody": {
                        "required": False,
                        "content": {
                            "application/json": {"schema": {"$ref": "#/components/schemas/MailboxImportRequest"}},
                            "text/plain": {"schema": {"type": "string", "description": "Each line: Address----Password----ClientID----OAuth2Token"}},
                            "application/x-www-form-urlencoded": {
                                "schema": {"type": "object", "properties": {"payload": {"type": "string"}}}
                            },
                        },
                    },
                    "responses": {"200": {"description": "Import result", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}}},
                },
            },
            "/api/mailboxes/{address}": {
                "delete": {
                    "tags": ["Mailboxes"],
                    "summary": "Delete mailbox",
                    "operationId": "deleteMailbox",
                    "security": auth_security,
                    "parameters": [{"$ref": "#/components/parameters/AddressParam"}],
                    "responses": {"200": {"description": "Delete result", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}}},
                }
            },
            "/api/mailboxes/{address}/messages": {
                "get": {
                    "tags": ["Messages"],
                    "summary": "List mailbox messages (Inbox + Junk)",
                    "operationId": "listMailboxMessages",
                    "security": auth_security,
                    "parameters": [
                        {"$ref": "#/components/parameters/AddressParam"},
                        {
                            "name": "limit",
                            "in": "query",
                            "required": False,
                            "schema": {"type": "integer", "minimum": 1, "maximum": 50, "default": 10},
                            "description": "Number of messages to return.",
                        },
                    ],
                    "responses": {
                        "200": {"description": "Message list", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}},
                        "404": {"description": "Mailbox not found", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiError"}}}},
                    },
                }
            },
            "/api/mailboxes/{address}/message/{uid}": {
                "get": {
                    "tags": ["Messages"],
                    "summary": "Get one message detail",
                    "operationId": "getMailboxMessage",
                    "security": auth_security,
                    "parameters": [
                        {"$ref": "#/components/parameters/AddressParam"},
                        {"$ref": "#/components/parameters/UidParam"},
                        {
                            "name": "folder",
                            "in": "query",
                            "required": False,
                            "schema": {"type": "string"},
                            "description": "IMAP folder name from message list response.",
                        },
                    ],
                    "responses": {
                        "200": {"description": "Message detail", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}},
                        "404": {"description": "Mailbox not found", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiError"}}}},
                    },
                }
            },
            "/api/mailboxes/{address}/message/{uid}/share": {
                "post": {
                    "tags": ["Messages"],
                    "summary": "Create share link for message",
                    "operationId": "shareMailboxMessage",
                    "security": auth_security,
                    "parameters": [
                        {"$ref": "#/components/parameters/AddressParam"},
                        {"$ref": "#/components/parameters/UidParam"},
                        {
                            "name": "folder",
                            "in": "query",
                            "required": False,
                            "schema": {"type": "string"},
                            "description": "IMAP folder name from message list response.",
                        },
                    ],
                    "responses": {
                        "200": {"description": "Share created", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}},
                        "404": {"description": "Mailbox not found", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiError"}}}},
                    },
                }
            },
            "/api/shares": {
                "get": {
                    "tags": ["Shares"],
                    "summary": "List shares",
                    "operationId": "listShares",
                    "security": auth_security,
                    "responses": {"200": {"description": "Share list", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}}},
                }
            },
            "/api/shares/{code}": {
                "get": {
                    "tags": ["Shares"],
                    "summary": "Get share detail",
                    "operationId": "getShare",
                    "security": auth_security,
                    "parameters": [{"$ref": "#/components/parameters/CodeParam"}],
                    "responses": {
                        "200": {"description": "Share detail", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}},
                        "404": {"description": "Share not found", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiError"}}}},
                    },
                },
                "delete": {
                    "tags": ["Shares"],
                    "summary": "Delete share",
                    "operationId": "deleteShare",
                    "security": auth_security,
                    "parameters": [{"$ref": "#/components/parameters/CodeParam"}],
                    "responses": {"200": {"description": "Share deleted", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/ApiEnvelope"}}}}},
                },
            },
        },
        "components": {
            "securitySchemes": {
                "ApiKeyAuth": {"type": "apiKey", "in": "header", "name": "X-API-Key"},
                "BearerAuth": {"type": "http", "scheme": "bearer"},
            },
            "parameters": {
                "AddressParam": {
                    "name": "address",
                    "in": "path",
                    "required": True,
                    "schema": {"type": "string"},
                    "description": "URL-encoded mailbox address, e.g. user%40outlook.com",
                },
                "UidParam": {
                    "name": "uid",
                    "in": "path",
                    "required": True,
                    "schema": {"type": "string"},
                    "description": "IMAP message UID.",
                },
                "CodeParam": {
                    "name": "code",
                    "in": "path",
                    "required": True,
                    "schema": {"type": "string"},
                    "description": "Share code.",
                },
            },
            "schemas": {
                "ApiEnvelope": {
                    "type": "object",
                    "properties": {
                        "ok": {"type": "boolean"},
                        "data": {"type": "object", "additionalProperties": True},
                    },
                    "required": ["ok"],
                },
                "ApiError": {
                    "type": "object",
                    "properties": {
                        "ok": {"type": "boolean", "example": False},
                        "error": {"type": "string"},
                    },
                    "required": ["ok", "error"],
                },
                "AuthLoginRequest": {
                    "type": "object",
                    "properties": {"api_key": {"type": "string"}},
                    "required": ["api_key"],
                },
                "MailboxImportItem": {
                    "type": "object",
                    "properties": {
                        "address": {"type": "string"},
                        "password": {"type": "string"},
                        "client_id": {"type": "string"},
                        "refresh_token": {"type": "string"},
                    },
                    "required": ["address"],
                },
                "MailboxImportRequest": {
                    "type": "object",
                    "properties": {
                        "items": {
                            "type": "array",
                            "items": {"$ref": "#/components/schemas/MailboxImportItem"},
                        }
                    },
                },
            },
        },
    }


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
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com https://cdn.jsdelivr.net; "
        "font-src https://fonts.gstatic.com; "
        "script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net; "
        "connect-src 'self'; "
        "form-action 'self'; "
        "frame-ancestors 'none'; "
        "base-uri 'none'"
    )
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    return response


@APP.get("/openapi.json")
def openapi_json():
    return jsonify(build_openapi_spec())


@APP.get("/docs")
def api_docs_page():
    return render_template_string(
        DOCS_TEMPLATE,
        spec_url=url_for("openapi_json"),
        home_url=url_for("index"),
    )


@APP.route("/")
def index():
    user_key = require_user()
    if user_key is None:
        return render_login_page("Please login or register to continue.")
    with get_db() as conn:
        if MULTI_USER:
            mailboxes = conn.execute(
                "SELECT * FROM mailboxes WHERE owner_key = ? ORDER BY created_at DESC",
                (user_key,),
            ).fetchall()
        else:
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
    user_key = require_user()
    if user_key is None:
        return render_login_page("Please login to import mailboxes.")
    return render_page(
        "Import - MailAdmin",
        IMPORT_TEMPLATE,
        active="import",
    )


@APP.route("/shares")
def shares_page():
    user_key = require_user()
    if user_key is None:
        return render_login_page("Please login to view shares.")
    with get_db() as conn:
        if MULTI_USER:
            shares = conn.execute(
                "SELECT * FROM shares WHERE owner_key = ? ORDER BY created_at DESC",
                (user_key,),
            ).fetchall()
        else:
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


@APP.route("/login")
def login_page():
    if not MULTI_USER:
        return redirect(url_for("index"))
    user_key = require_user()
    if user_key is not None:
        return redirect(url_for("index"))
    return render_login_page()


@APP.route("/account")
def account_page():
    user_key = require_user()
    if user_key is None:
        return render_login_page("Please login to manage your account.")
    return render_page(
        "Account - MailAdmin",
        ACCOUNT_TEMPLATE,
        active="account",
    )


@APP.route("/logout")
def logout_page():
    response = render_page("Logout - MailAdmin", LOGOUT_TEMPLATE, active="account")
    resp = APP.make_response(response)
    resp.set_cookie(API_KEY_COOKIE, "", max_age=0, path="/")
    return resp


@APP.get("/api/health")
def api_health():
    return api_ok({"status": "ok", "time": int(time.time()), "multi_user": MULTI_USER})


@APP.post("/api/auth/login")
def api_auth_login():
    data = request.get_json(silent=True) or {}
    api_key = (data.get("api_key") or "").strip()
    if not api_key:
        return api_error("Missing api_key.", status=400)
    if not is_valid_api_key(api_key):
        return api_error("Invalid API key.", status=401)
    return api_ok({"api_key": api_key})


@APP.post("/api/auth/register")
def api_auth_register():
    if not MULTI_USER:
        return api_error("Multi-user mode is disabled.", status=400)
    with get_db() as conn:
        api_key = generate_api_key(conn)
        now = int(time.time())
        conn.execute(
            "INSERT INTO users (api_key, created_at, last_rotated_at) VALUES (?, ?, ?)",
            (api_key, now, now),
        )
    return api_ok({"api_key": api_key})


@APP.post("/api/auth/rotate")
def api_auth_rotate():
    user_key = require_user(api=True)
    if isinstance(user_key, tuple):
        return user_key
    if not MULTI_USER:
        return api_error("Multi-user mode is disabled.", status=400)
    with get_db() as conn:
        row = conn.execute(
            "SELECT last_rotated_at FROM users WHERE api_key = ?",
            (user_key,),
        ).fetchone()
        if not row:
            return api_error("Invalid API key.", status=401)
        now = int(time.time())
        delta = now - int(row["last_rotated_at"] or 0)
        if delta < 86400:
            remaining = 86400 - delta
            hours = remaining // 3600
            minutes = (remaining % 3600) // 60
            return api_error(
                f"Rotate allowed once every 24 hours. Try again in {hours}h {minutes}m.",
                status=429,
            )
        new_key = generate_api_key(conn)
        conn.execute(
            "INSERT INTO users (api_key, created_at, last_rotated_at) VALUES (?, ?, ?)",
            (new_key, now, now),
        )
        conn.execute(
            "UPDATE mailboxes SET owner_key = ? WHERE owner_key = ?",
            (new_key, user_key),
        )
        conn.execute(
            "UPDATE shares SET owner_key = ? WHERE owner_key = ?",
            (new_key, user_key),
        )
        conn.execute("DELETE FROM users WHERE api_key = ?", (user_key,))
    return api_ok({"api_key": new_key})


@APP.get("/api/mailboxes")
def api_list_mailboxes():
    user_key = require_user(api=True)
    if isinstance(user_key, tuple):
        return user_key
    with get_db() as conn:
        if MULTI_USER:
            mailboxes = conn.execute(
                "SELECT * FROM mailboxes WHERE owner_key = ? ORDER BY created_at DESC",
                (user_key,),
            ).fetchall()
        else:
            mailboxes = conn.execute(
                "SELECT * FROM mailboxes ORDER BY created_at DESC"
            ).fetchall()
    payload = [row_to_dict(row) for row in mailboxes]
    return api_ok(payload)


@APP.post("/api/mailboxes")
def api_import_mailboxes():
    user_key = require_user(api=True)
    if isinstance(user_key, tuple):
        return user_key
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
    owner_errors = []
    with get_db() as conn:
        for address, password, client_id, refresh_token in entries:
            existing = conn.execute(
                "SELECT id, owner_key FROM mailboxes WHERE address = ?",
                (address,),
            ).fetchone()
            if existing:
                if MULTI_USER and existing["owner_key"] not in ("", user_key):
                    owner_errors.append(f"Address {address} already owned by another user.")
                    continue
                conn.execute(
                    """
                    UPDATE mailboxes
                    SET password = ?, client_id = ?, refresh_token = ?, owner_key = ?, updated_at = ?
                    WHERE address = ?
                    """,
                    (
                        password,
                        client_id,
                        refresh_token,
                        user_key if MULTI_USER else "",
                        now,
                        address,
                    ),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO mailboxes (address, password, client_id, refresh_token, owner_key, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        address,
                        password,
                        client_id,
                        refresh_token,
                        user_key if MULTI_USER else "",
                        now,
                        now,
                    ),
                )
            imported += 1
    if owner_errors:
        for err in owner_errors[:3]:
            flash(err, "error")
        if len(owner_errors) > 3:
            flash(f"{len(owner_errors) - 3} more errors hidden.", "error")
    return api_ok({"imported": imported, "errors": errors})


@APP.delete("/api/mailboxes/<path:address>")
def api_delete_mailbox(address):
    user_key = require_user(api=True)
    if isinstance(user_key, tuple):
        return user_key
    with get_db() as conn:
        if MULTI_USER:
            conn.execute(
                "DELETE FROM mailboxes WHERE address = ? AND owner_key = ?",
                (address, user_key),
            )
        else:
            conn.execute("DELETE FROM mailboxes WHERE address = ?", (address,))
    return api_ok({"deleted": address})


@APP.get("/api/mailboxes/<path:address>/messages")
def api_list_messages(address):
    user_key = require_user(api=True)
    if isinstance(user_key, tuple):
        return user_key
    limit_raw = request.args.get("limit", str(DEFAULT_LIMIT))
    try:
        limit = int(limit_raw)
    except ValueError:
        limit = DEFAULT_LIMIT
    limit = max(1, min(MAX_LIMIT, limit))
    with get_db() as conn:
        if MULTI_USER:
            mailbox = conn.execute(
                "SELECT * FROM mailboxes WHERE address = ? AND owner_key = ?",
                (address, user_key),
            ).fetchone()
        else:
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
    user_key = require_user(api=True)
    if isinstance(user_key, tuple):
        return user_key
    with get_db() as conn:
        if MULTI_USER:
            mailbox = conn.execute(
                "SELECT * FROM mailboxes WHERE address = ? AND owner_key = ?",
                (address, user_key),
            ).fetchone()
        else:
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
    user_key = require_user(api=True)
    if isinstance(user_key, tuple):
        return user_key
    with get_db() as conn:
        if MULTI_USER:
            mailbox = conn.execute(
                "SELECT * FROM mailboxes WHERE address = ? AND owner_key = ?",
                (address, user_key),
            ).fetchone()
        else:
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
            INSERT INTO shares (code, subject, mail_from, mail_to, mail_dt, body_html, owner_key, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                code,
                message["subject"],
                message["mail_from"],
                message["mail_to"],
                message["mail_dt"],
                build_share_body(message),
                user_key if MULTI_USER else "",
                int(time.time()),
            ),
        )

    share_url = request.host_url.rstrip("/") + url_for("view_share", code=code)
    return api_ok({"code": code, "url": share_url})


@APP.get("/api/shares")
def api_list_shares():
    user_key = require_user(api=True)
    if isinstance(user_key, tuple):
        return user_key
    with get_db() as conn:
        if MULTI_USER:
            shares = conn.execute(
                "SELECT * FROM shares WHERE owner_key = ? ORDER BY created_at DESC",
                (user_key,),
            ).fetchall()
        else:
            shares = conn.execute(
                "SELECT * FROM shares ORDER BY created_at DESC"
            ).fetchall()
    payload = [row_to_dict(row) for row in shares]
    return api_ok(payload)


@APP.get("/api/shares/<code>")
def api_get_share(code):
    user_key = require_user(api=True)
    if isinstance(user_key, tuple):
        return user_key
    with get_db() as conn:
        if MULTI_USER:
            share = conn.execute(
                "SELECT * FROM shares WHERE code = ? AND owner_key = ?",
                (code, user_key),
            ).fetchone()
        else:
            share = conn.execute(
                "SELECT * FROM shares WHERE code = ?", (code,)
            ).fetchone()
    if not share:
        return api_error("Share not found.", status=404)
    return api_ok(row_to_dict(share))


@APP.delete("/api/shares/<code>")
def api_delete_share(code):
    user_key = require_user(api=True)
    if isinstance(user_key, tuple):
        return user_key
    with get_db() as conn:
        if MULTI_USER:
            conn.execute(
                "DELETE FROM shares WHERE code = ? AND owner_key = ?",
                (code, user_key),
            )
        else:
            conn.execute("DELETE FROM shares WHERE code = ?", (code,))
    return api_ok({"deleted": code})


@APP.post("/import")
def import_mailboxes():
    user_key = require_user()
    if user_key is None:
        return render_login_page("Please login to import mailboxes.")
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
                "SELECT id, owner_key FROM mailboxes WHERE address = ?",
                (address,),
            ).fetchone()
            if existing:
                if MULTI_USER and existing["owner_key"] not in ("", user_key):
                    errors.append(f"Address {address} already owned by another user.")
                    continue
                conn.execute(
                    """
                    UPDATE mailboxes
                    SET password = ?, client_id = ?, refresh_token = ?, owner_key = ?, updated_at = ?
                    WHERE address = ?
                    """,
                    (
                        password,
                        client_id,
                        refresh_token,
                        user_key if MULTI_USER else "",
                        now,
                        address,
                    ),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO mailboxes (address, password, client_id, refresh_token, owner_key, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        address,
                        password,
                        client_id,
                        refresh_token,
                        user_key if MULTI_USER else "",
                        now,
                        now,
                    ),
                )
            imported += 1
    flash(f"Imported {imported} mailbox(es).", "success")
    return redirect(url_for("import_page"))


@APP.post("/mailbox/<path:address>/delete")
def delete_mailbox(address):
    user_key = require_user()
    if user_key is None:
        return render_login_page("Please login to delete mailboxes.")
    with get_db() as conn:
        if MULTI_USER:
            conn.execute(
                "DELETE FROM mailboxes WHERE address = ? AND owner_key = ?",
                (address, user_key),
            )
        else:
            conn.execute("DELETE FROM mailboxes WHERE address = ?", (address,))
    flash("Mailbox deleted.", "success")
    return redirect(url_for("index"))


@APP.route("/mailbox/<path:address>")
def view_mailbox(address):
    user_key = require_user()
    if user_key is None:
        return render_login_page("Please login to view mailboxes.")
    limit_raw = request.args.get("limit", str(DEFAULT_LIMIT))
    try:
        limit = int(limit_raw)
    except ValueError:
        limit = DEFAULT_LIMIT
    limit = max(1, min(MAX_LIMIT, limit))

    with get_db() as conn:
        if MULTI_USER:
            mailbox = conn.execute(
                "SELECT * FROM mailboxes WHERE address = ? AND owner_key = ?",
                (address, user_key),
            ).fetchone()
        else:
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
    user_key = require_user()
    if user_key is None:
        return render_login_page("Please login to view messages.")
    with get_db() as conn:
        if MULTI_USER:
            mailbox = conn.execute(
                "SELECT * FROM mailboxes WHERE address = ? AND owner_key = ?",
                (address, user_key),
            ).fetchone()
        else:
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
    user_key = require_user()
    if user_key is None:
        return render_login_page("Please login to share messages.")
    with get_db() as conn:
        if MULTI_USER:
            mailbox = conn.execute(
                "SELECT * FROM mailboxes WHERE address = ? AND owner_key = ?",
                (address, user_key),
            ).fetchone()
        else:
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
            INSERT INTO shares (code, subject, mail_from, mail_to, mail_dt, body_html, owner_key, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                code,
                message["subject"],
                message["mail_from"],
                message["mail_to"],
                message["mail_dt"],
                build_share_body(message),
                user_key if MULTI_USER else "",
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
    user_key = require_user()
    if user_key is None:
        return render_login_page("Please login to manage shares.")
    with get_db() as conn:
        if MULTI_USER:
            conn.execute(
                "DELETE FROM shares WHERE code = ? AND owner_key = ?",
                (code, user_key),
            )
        else:
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
