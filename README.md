# MailAdmin

Multi-mailbox Outlook mail viewer with HTML previews, share links, and a JSON API.

## Features
- Import and manage multiple mailboxes.
- Merge Inbox + Junk into one list, sorted by time.
- View message HTML safely in a sandboxed iframe.
- Share messages via `/share/{8chars}` links.
- Full JSON API for automation.
- Auto-opens the default browser on start.

## Requirements
- Python 3.10+
- Outlook IMAP enabled

Install deps:
```bash
pip install -r requirements.txt
```

## Run
```bash
python getMail.py
```

By default the app listens on `0.0.0.0:5000` and opens
`http://127.0.0.1:5000/` in your browser.

### Environment variables
- `MAILADMIN_HOST` (default `0.0.0.0`)
- `MAILADMIN_PORT` (default `5000`)
- `MAILADMIN_IMAP_HOST` (default `outlook.live.com`)
- `MAILADMIN_SECRET` (Flask session secret; auto-generated if not set)
- `MAILADMIN_NO_BROWSER=1` (disable auto open)

## Import mailboxes
Import format (one per line):
```
Address----Password----ClientID----OAuth2Token
```

Notes:
- If `ClientID` is empty, `OAuth2Token` is treated as an access token.
- If both OAuth fields are empty, password login is used.

## Web pages
- `/` Mailboxes
- `/import` Import
- `/shares` Share history

## API
Base: `http://127.0.0.1:5000`
When calling endpoints with `{address}`, URL-encode the mailbox address.

### Quick API guide
Add a mailbox (JSON):
```bash
curl -X POST http://127.0.0.1:5000/api/mailboxes \
  -H "Content-Type: application/json" \
  -d '{"items":[{"address":"user@outlook.com","password":"","client_id":"","refresh_token":"YOUR_REFRESH_TOKEN"}]}'
```

List mailboxes:
```bash
curl http://127.0.0.1:5000/api/mailboxes
```

List messages for a mailbox (Inbox + Junk):
```bash
curl "http://127.0.0.1:5000/api/mailboxes/user%40outlook.com/messages?limit=5"
```

Get a single message (use `uid` + `folder` from the list response):
```bash
curl "http://127.0.0.1:5000/api/mailboxes/user%40outlook.com/message/12345?folder=Junk"
```

### Health
```
GET /api/health
```

### Mailboxes
```
GET /api/mailboxes
POST /api/mailboxes
DELETE /api/mailboxes/{address}
```

`POST /api/mailboxes` accepts either:
- JSON:
```json
{
  "items": [
    {
      "address": "user@example.com",
      "password": "",
      "client_id": "",
      "refresh_token": ""
    }
  ]
}
```
- Plain text body / form field `payload` with multi-line format.

### Messages
```
GET /api/mailboxes/{address}/messages?limit=10
GET /api/mailboxes/{address}/message/{uid}?folder=Junk
POST /api/mailboxes/{address}/message/{uid}/share?folder=Junk
```

### Shares
```
GET /api/shares
GET /api/shares/{code}
DELETE /api/shares/{code}
```

## Notes on Junk folder
The app queries both Inbox and Junk folders for Outlook. If your Junk
folder name differs, add it to `JUNK_FOLDERS` in `getMail.py`.
