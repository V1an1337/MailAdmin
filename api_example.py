#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import os
import sys
from urllib.parse import quote

import requests


def die(message):
    print(message, file=sys.stderr)
    sys.exit(1)


def request_json(method, url, **kwargs):
    try:
        resp = requests.request(method, url, timeout=15, **kwargs)
    except requests.RequestException as exc:
        die(f"Request failed: {exc}")
    try:
        payload = resp.json()
    except ValueError:
        die(f"Invalid JSON response ({resp.status_code}).")
    if not payload.get("ok"):
        die(f"API error: {payload.get('error')}")
    return payload.get("data")


def choose_mailbox(base_url, mailbox_address, headers):
    if mailbox_address:
        return mailbox_address
    mailboxes = request_json("GET", f"{base_url}/api/mailboxes", headers=headers)
    if not mailboxes:
        die("No mailboxes found. Import a mailbox first.")
    return mailboxes[0]["address"]


def main():
    parser = argparse.ArgumentParser(description="MailAdmin API inbox test client")
    parser.add_argument("--base-url", default="http://127.0.0.1:5000")
    parser.add_argument("--mailbox-address")
    parser.add_argument("--limit", type=int, default=5)
    parser.add_argument(
        "--api-key",
        default=os.environ.get("MAILADMIN_API_KEY", ""),
        help="API key for multi-user mode (or set MAILADMIN_API_KEY)",
    )
    parser.add_argument(
        "--no-show-message",
        action="store_true",
        help="Do not fetch the first message content",
    )
    parser.add_argument(
        "--max-body",
        type=int,
        default=2000,
        help="Max chars to print for message body",
    )
    args = parser.parse_args()

    headers = {}
    if args.api_key:
        headers["X-API-Key"] = args.api_key
    mailbox_address = choose_mailbox(args.base_url, args.mailbox_address, headers)
    mailbox_path = quote(mailbox_address, safe="")
    messages = request_json(
        "GET",
        f"{args.base_url}/api/mailboxes/{mailbox_path}/messages",
        params={"limit": args.limit},
        headers=headers,
    )

    if not messages:
        print("No messages.")
        return

    print(f"Mailbox {mailbox_address} messages:")
    for msg in messages:
        subject = msg.get("subject") or "(No subject)"
        sender = msg.get("mail_from") or "-"
        when = msg.get("mail_dt") or "-"
        folder = msg.get("folder_label") or msg.get("folder") or "-"
        print(f"- {when} | {folder} | {sender} | {subject}")

    if args.no_show_message:
        return

    first = messages[0]
    uid = first.get("uid")
    if not uid:
        die("First message missing UID.")
    folder = first.get("folder")
    params = {}
    if folder:
        params["folder"] = folder

    detail = request_json(
        "GET",
        f"{args.base_url}/api/mailboxes/{mailbox_path}/message/{uid}",
        params=params,
        headers=headers,
    )

    subject = detail.get("subject") or "(No subject)"
    sender = detail.get("mail_from") or "-"
    recipient = detail.get("mail_to") or "-"
    when = detail.get("mail_dt") or "-"
    folder_label = detail.get("folder_label") or detail.get("folder") or "-"
    body_text = detail.get("body_text") or ""
    body_html = detail.get("safe_body_html") or ""
    body = body_text if body_text.strip() else body_html
    body = body.strip()
    if args.max_body and len(body) > args.max_body:
        body = body[: args.max_body] + "...(truncated)"

    print("\nFirst message detail:")
    print(f"Subject: {subject}")
    print(f"From: {sender}")
    print(f"To: {recipient}")
    print(f"Date: {when}")
    print(f"Folder: {folder_label}")
    print("\nBody:")
    print(body or "(empty)")


if __name__ == "__main__":
    main()
