# MailAdmin

Multi-mailbox Outlook mail viewer with HTML previews, share links, and a JSON API.
<img width="1445" height="692" alt="image" src="https://github.com/user-attachments/assets/ed7408db-9033-4a6a-a65b-dc8e49bf9726" />


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
python server.py
```

By default the app listens on `0.0.0.0:5000` and opens
`http://127.0.0.1:5000/` in your browser.

### Environment variables
- `MAILADMIN_HOST` (default `0.0.0.0`)
- `MAILADMIN_PORT` (default `5000`)
- `MAILADMIN_IMAP_HOST` (default `outlook.live.com`)
- `MAILADMIN_SECRET` (Flask session secret; auto-generated if not set)
- `MAILADMIN_NO_BROWSER=1` (disable auto open)
- `MAILADMIN_MULTI_USER=1` (enable multi-user mode with API keys)

The app auto-loads `.env` if present.

## Multi-user mode
When `MAILADMIN_MULTI_USER=1`, every user is identified by an API key:
- Register to get a new API key (shown once). Save it safely.
- Login validates the API key.
- Logout clears local storage and cookies.
- API key rotation is allowed once every 24 hours.
- Mailboxes and shares are isolated per API key.

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
- `/login` Login/Register
- `/account` Rotate API key
- `/logout` Clear local credentials

## API
Base: `http://127.0.0.1:5000`
When calling endpoints with `{address}`, URL-encode the mailbox address.
When multi-user mode is enabled, pass the API key via `X-API-Key` header (or `Authorization: Bearer <key>`).
Full API docs: [api.md](https://github.com/V1an1337/MailAdmin/blob/main/api.md).

### Quick API guide
Register a new API key:
```bash
curl -X POST http://127.0.0.1:5000/api/auth/register
```

Add a mailbox (JSON):
```bash
curl -X POST http://127.0.0.1:5000/api/mailboxes \
  -H "Content-Type: application/json" \
  -H "X-API-Key: YOUR_API_KEY" \
  -d '{"items":[{"address":"user@outlook.com","password":"","client_id":"","refresh_token":"YOUR_REFRESH_TOKEN"}]}'
```

List mailboxes:
```bash
curl -H "X-API-Key: YOUR_API_KEY" http://127.0.0.1:5000/api/mailboxes
```

List messages for a mailbox (Inbox + Junk):
```bash
curl -H "X-API-Key: YOUR_API_KEY" \
  "http://127.0.0.1:5000/api/mailboxes/user%40outlook.com/messages?limit=5"
```

Get a single message (use `uid` + `folder` from the list response):
```bash
curl -H "X-API-Key: YOUR_API_KEY" \
  "http://127.0.0.1:5000/api/mailboxes/user%40outlook.com/message/12345?folder=Junk"
```

### Health
```
GET /api/health
```

### Auth
```
POST /api/auth/login
POST /api/auth/register
POST /api/auth/rotate
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
folder name differs, add it to `JUNK_FOLDERS` in `server.py`.
