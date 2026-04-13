# outlook-mcp

MCP server for on-premise Microsoft Exchange via EWS (`exchangelib`).

## What it can do

### System
- `ping_exchange`
- `get_mailbox_info`

### Email
- `list_emails`
- `get_email`
- `search_emails`
- `send_email`
- `reply_email`
- `forward_email`
- `move_email`
- `copy_email`
- `delete_email`
- `mark_email`
- `list_folders`
- `create_folder`
- `create_draft`
- `send_draft`
- `get_attachment`

### Calendar
- `list_events`
- `get_event`
- `create_event`
- `update_event`
- `delete_event`
- `respond_to_invite`
- `find_free_slots`
- `get_my_availability`
- `list_calendars`

### Contacts
- `search_contacts`
- `get_contact`
- `create_contact`
- `update_contact`
- `delete_contact`

## Security notes

What the current code does:
- connects only to the Exchange/EWS endpoint configured in `EXCHANGE_SERVER`
- does **not** contain telemetry, analytics, or third-party data export logic
- keeps secrets in environment variables / `.env`
- ignores local secret files via `.gitignore` (`.env`, `.env.*`, while keeping `.env.example`)
- logs only tool name, status, duration, and error code; it does **not** log message bodies, attachment contents, or passwords

What you should still be careful with:
- `EXCHANGE_VERIFY_SSL=false` disables TLS certificate verification and should be used only for trusted internal/self-signed environments
- `get_attachment` writes files to disk, so choose a safe destination directory
- `outlook-mcp-smoke` is privacy-safe by default and prints only masked mailbox info plus counts; set `OUTLOOK_MCP_SMOKE_INCLUDE_DATA=true` only if you explicitly want real inbox/event data in stdout
- if you enable file logging with `LOG_FILE`, protect that file with OS permissions

## Local run

```bash
uv venv
source .venv/bin/activate
uv pip install -e .[dev]
cp .env.example .env
outlook-mcp
```

By default the server runs in `stdio` mode. Set `MCP_TRANSPORT=sse` to start an HTTP server.

## Configuration

Example `.env`:

```dotenv
EXCHANGE_SERVER=https://mail.company.com/EWS/Exchange.asmx
EXCHANGE_USERNAME=DOMAIN\\username
EXCHANGE_PASSWORD=secret
EXCHANGE_EMAIL_ADDRESS=user@company.com
EXCHANGE_VERIFY_SSL=true
EXCHANGE_AUTH_TYPE=NTLM
EXCHANGE_VERSION=EXCHANGE_2016
EXCHANGE_TIMEOUT=30
EXCHANGE_MAX_RETRIES=3
EXCHANGE_TIMEZONE=Europe/Moscow
EXCHANGE_IMPERSONATE_AS=
ATTACHMENT_MAX_SIZE_MB=10
MCP_TRANSPORT=stdio
MCP_SSE_HOST=127.0.0.1
MCP_SSE_PORT=8080
LOG_LEVEL=INFO
LOG_FILE=
```

Notes:
- set `EXCHANGE_EMAIL_ADDRESS` when `EXCHANGE_USERNAME` is not an SMTP address
- `OAuth2` is reserved in config, but this build currently supports live auth with `NTLM` or `Basic`
- `EXCHANGE_IMPERSONATE_AS` enables mailbox impersonation when Exchange permissions are configured accordingly

## Claude Desktop example

```json
{
  "mcpServers": {
    "outlook": {
      "command": "outlook-mcp",
      "env": {
        "EXCHANGE_SERVER": "https://mail.company.com/EWS/Exchange.asmx",
        "EXCHANGE_USERNAME": "DOMAIN\\username",
        "EXCHANGE_PASSWORD": "secret",
        "EXCHANGE_EMAIL_ADDRESS": "user@company.com",
        "EXCHANGE_AUTH_TYPE": "NTLM"
      }
    }
  }
}
```

## Smoke check

After filling `.env`, run:

```bash
outlook-mcp-smoke
```

Default output is sanitized for safer verification. If you intentionally want sample mailbox/event data in the output:

```bash
OUTLOOK_MCP_SMOKE_INCLUDE_DATA=true outlook-mcp-smoke
```

## Docker

```bash
docker build -t outlook-mcp .
docker run --rm --env-file .env outlook-mcp
```

## Development

```bash
uv run --python 3.12 --with '.[dev]' pytest -q
uv run --python 3.12 --with '.[dev]' ruff check .
```

## Project notes

- The implementation is centered around a single `ExchangeClient` abstraction so auth, transport, retries, and error mapping stay centralized.
- Errors are returned in a structured JSON form suitable for MCP `isError=true` handling.
- `RTK.md` is not present in this repository, so the implementation follows `PLAN.md`.
