# MailAdmin API Documentation

Base URL: `http://127.0.0.1:5000`

Production URL: `https://oauth.v1an.xyz`

Interactive docs:
- Swagger UI: `/docs`
- OpenAPI spec JSON: `/openapi.json`

All responses are JSON in the form:
```json
{ "ok": true, "data": ... }
```
or
```json
{ "ok": false, "error": "message" }
```

## Authentication

If `MAILADMIN_MULTI_USER=1`:
- Pass API key via `X-API-Key` header, or
- `Authorization: Bearer <api_key>`

If multi-user is disabled, auth headers are ignored.

## Token Keepalive

The server runs a background refresh worker for Outlook OAuth refresh tokens.

Config:
- `MAILADMIN_TOKEN_REFRESH_ENABLED` (default `1`)
- `MAILADMIN_TOKEN_REFRESH_INTERVAL_SEC` (default `1.0`)
- `MAILADMIN_TOKEN_REFRESH_DAYS` (default `60`)

Behavior:
- New OAuth mailbox imports are marked `pending_initial` and queued immediately.
- Successful refresh rotates to new refresh token when returned.
- Previous refresh token is retained for rollback.
- On token usage failure, server may auto-fallback to previous token.

## OpenAPI and SDK generation

This project now exposes an OpenAPI spec at:
- Local: `http://127.0.0.1:5000/openapi.json`
- Production: `https://oauth.v1an.xyz/openapi.json`

Generate Python client using `openapi-python-client`:
```bash
pip install openapi-python-client
openapi-python-client generate --url https://oauth.v1an.xyz/openapi.json
```

Generate Python client using `openapi-generator`:
```bash
openapi-generator-cli generate \
  -i https://oauth.v1an.xyz/openapi.json \
  -g python \
  -o ./mailadmin-python-sdk
```

## Health

### GET /api/health
Check server status.

Response:
```json
{
  "ok": true,
  "data": {
    "status": "ok",
    "time": 1700000000,
    "multi_user": true
  }
}
```

Example:
```bash
curl http://127.0.0.1:5000/api/health
```

## Documentation endpoints

### GET /openapi.json
Return machine-readable OpenAPI JSON schema.

Example:
```bash
curl http://127.0.0.1:5000/openapi.json
```

### GET /docs
Swagger UI web docs (human-friendly interactive API explorer).

## Auth

### POST /api/auth/login
Validate API key.

Request JSON:
```json
{ "api_key": "YOUR_API_KEY" }
```

Response:
```json
{ "ok": true, "data": { "api_key": "YOUR_API_KEY" } }
```

Example:
```bash
curl -X POST http://127.0.0.1:5000/api/auth/login \
  -H "Content-Type: application/json" \
  -d '{"api_key":"YOUR_API_KEY"}'
```

### POST /api/auth/register
Create a new API key (multi-user only).

Response:
```json
{ "ok": true, "data": { "api_key": "NEW_API_KEY" } }
```

Example:
```bash
curl -X POST http://127.0.0.1:5000/api/auth/register
```

### POST /api/auth/rotate
Rotate API key (multi-user only, 24h cooldown).

Headers:
```
X-API-Key: YOUR_API_KEY
```

Response:
```json
{ "ok": true, "data": { "api_key": "NEW_API_KEY" } }
```

Cooldown error:
```json
{ "ok": false, "error": "Rotate allowed once every 24 hours. Try again in 2h 15m." }
```

Example:
```bash
curl -X POST http://127.0.0.1:5000/api/auth/rotate \
  -H "X-API-Key: YOUR_API_KEY"
```

## Mailboxes

### GET /api/mailboxes
List mailboxes owned by the user (or all if multi-user disabled).

Headers (multi-user):
```
X-API-Key: YOUR_API_KEY
```

Response item fields:
- `id` (int)
- `address` (string)
- `password` (string)
- `client_id` (string)
- `refresh_token` (string)
- `owner_key` (string)
- `created_at` (unix timestamp)
- `updated_at` (unix timestamp)
- `token_status` (`unknown` | `healthy` | `pending_initial` | `warning` | `rollback_ok` | `degraded`)
- `token_last_warning` (string)
- `token_last_error` (string)
- `token_last_error_at` (unix timestamp)
- `token_next_refresh_at` (unix timestamp)
- `refresh_token_updated_at` (unix timestamp)
- `refresh_token_prev_updated_at` (unix timestamp)
- `refresh_token_expires_at` (unix timestamp)

Example:
```bash
curl http://127.0.0.1:5000/api/mailboxes \
  -H "X-API-Key: YOUR_API_KEY"
```

### POST /api/mailboxes
Import or update mailboxes.

Headers (multi-user):
```
X-API-Key: YOUR_API_KEY
```

Request (JSON):
```json
{
  "items": [
    {
      "address": "user@outlook.com",
      "password": "",
      "client_id": "",
      "refresh_token": "REFRESH_TOKEN_OR_ACCESS_TOKEN"
    }
  ]
}
```

Request (plain text or form field `payload`):
```
Address----Password----ClientID----OAuth2Token
```

Response:
```json
{ "ok": true, "data": { "imported": 1, "errors": [] } }
```

Example (JSON):
```bash
curl -X POST http://127.0.0.1:5000/api/mailboxes \
  -H "Content-Type: application/json" \
  -H "X-API-Key: YOUR_API_KEY" \
  -d '{"items":[{"address":"user@outlook.com","password":"","client_id":"","refresh_token":"TOKEN"}]}'
```

Example (text):
```bash
curl -X POST http://127.0.0.1:5000/api/mailboxes \
  -H "X-API-Key: YOUR_API_KEY" \
  --data-binary $'user@outlook.com----pass----clientid----token\n'
```

### DELETE /api/mailboxes/{address}
Delete a mailbox by address.

Path params:
- `address` (URL-encoded email address)

Headers (multi-user):
```
X-API-Key: YOUR_API_KEY
```

Response:
```json
{ "ok": true, "data": { "deleted": "user@outlook.com" } }
```

Example:
```bash
curl -X DELETE \
  http://127.0.0.1:5000/api/mailboxes/user%40outlook.com \
  -H "X-API-Key: YOUR_API_KEY"
```

## Messages

### GET /api/mailboxes/{address}/messages
List recent messages (merged Inbox + Junk).

Path params:
- `address` (URL-encoded email address)

Query params:
- `limit` (int, 1-50, default 10)

Headers (multi-user):
```
X-API-Key: YOUR_API_KEY
```

Response item fields:
- `uid` (string)
- `subject` (string)
- `mail_from` (string)
- `mail_to` (string)
- `mail_dt` (string, local time)
- `mail_ts` (int, unix timestamp)
- `folder` (string, IMAP folder name)
- `folder_label` ("Inbox" or "Junk")

If token fallback happened during this request, `data` may be:
```json
{
  "messages": [ ... ],
  "warning": "Rollback applied for user@outlook.com: switched to previous refresh token."
}
```

Example:
```bash
curl "http://127.0.0.1:5000/api/mailboxes/user%40outlook.com/messages?limit=5" \
  -H "X-API-Key: YOUR_API_KEY"
```

### GET /api/mailboxes/{address}/message/{uid}
Fetch full message body.

Path params:
- `address` (URL-encoded email address)
- `uid` (string)

Query params:
- `folder` (string, from list response)

Headers (multi-user):
```
X-API-Key: YOUR_API_KEY
```

Response fields include:
- `subject`, `mail_from`, `mail_to`, `mail_dt`, `mail_ts`
- `body_text`, `body_html`, `safe_body_html`
- `folder`, `folder_label`, `uid`
- Optional: `warning` (string) when fallback occurred during this request

Example:
```bash
curl "http://127.0.0.1:5000/api/mailboxes/user%40outlook.com/message/12345?folder=Junk" \
  -H "X-API-Key: YOUR_API_KEY"
```

### POST /api/mailboxes/{address}/message/{uid}/share
Create a share link for a message.

Path params:
- `address` (URL-encoded email address)
- `uid` (string)

Query params:
- `folder` (string, from list response)

Headers (multi-user):
```
X-API-Key: YOUR_API_KEY
```

Response:
```json
{ "ok": true, "data": { "code": "Ab12Cd34", "url": "http://127.0.0.1:5000/share/Ab12Cd34" } }
```

If fallback happened, response may also contain:
```json
{
  "ok": true,
  "data": {
    "code": "Ab12Cd34",
    "url": "http://127.0.0.1:5000/share/Ab12Cd34",
    "warning": "Rollback applied for user@outlook.com: switched to previous refresh token."
  }
}
```

Example:
```bash
curl -X POST \
  "http://127.0.0.1:5000/api/mailboxes/user%40outlook.com/message/12345/share?folder=Junk" \
  -H "X-API-Key: YOUR_API_KEY"
```

## Shares

### GET /api/shares
List share records owned by the user.

Headers (multi-user):
```
X-API-Key: YOUR_API_KEY
```

Response item fields:
- `code`, `subject`, `mail_from`, `mail_to`, `mail_dt`
- `body_html`, `owner_key`, `created_at`

Example:
```bash
curl http://127.0.0.1:5000/api/shares \
  -H "X-API-Key: YOUR_API_KEY"
```

### GET /api/shares/{code}
Get one share record by code.

Headers (multi-user):
```
X-API-Key: YOUR_API_KEY
```

Example:
```bash
curl http://127.0.0.1:5000/api/shares/Ab12Cd34 \
  -H "X-API-Key: YOUR_API_KEY"
```

### DELETE /api/shares/{code}
Delete a share record.

Headers (multi-user):
```
X-API-Key: YOUR_API_KEY
```

Example:
```bash
curl -X DELETE http://127.0.0.1:5000/api/shares/Ab12Cd34 \
  -H "X-API-Key: YOUR_API_KEY"
```
