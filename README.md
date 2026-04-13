# outlook-mcp

![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![MCP](https://img.shields.io/badge/MCP-server-7c3aed)
![Exchange](https://img.shields.io/badge/Microsoft%20Exchange-EWS-0a7ea4)
![Status](https://img.shields.io/badge/status-beta-orange)

`outlook-mcp` is an MCP server for on-prem Microsoft Exchange via EWS (`exchangelib`).
It gives MCP-compatible clients access to email, calendar, contacts, folders, attachments, and availability data through a single, testable Python service.

## Short description

Secure MCP server for on-prem Microsoft Exchange (EWS) with tools for email, calendar, contacts, folders, attachments, and free/busy availability.

## Suggested repository topics / tags

`mcp`, `model-context-protocol`, `exchange`, `microsoft-exchange`, `ews`, `outlook`, `email`, `calendar`, `contacts`, `python`, `automation`, `exchangelib`

## Highlights

- email operations: list, search, read, send, reply, forward, move, copy, delete, mark
- calendar operations: list, create, update, delete, respond to invites, find free slots
- contacts operations: search, read, create, update, delete
- folder operations and attachment download
- Exchange auth via `NTLM` and `Basic`
- MCP transport via `stdio` and `SSE`
- centralized error mapping and a single `ExchangeClient` abstraction
- privacy-safer smoke check output by default
- Docker support and GitLab CI/CD pipeline included

## Tool catalog

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

## Typical use cases

- connect Claude Desktop or another MCP client to on-prem Exchange
- search inbox messages and fetch full email content
- send or draft emails from AI workflows
- inspect calendars and create meetings
- check free/busy windows for scheduling
- search personal contacts or the GAL
- expose Exchange operations through a controlled MCP boundary instead of direct mailbox scripting

## Security notes

What the current code does:
- connects only to the Exchange/EWS endpoint configured in `EXCHANGE_SERVER`
- does **not** contain telemetry, analytics, or third-party data export logic
- keeps secrets in environment variables / `.env`
- ignores local secret files via `.gitignore` (`.env`, `.env.*`, while keeping `.env.example`)
- logs only tool name, status, duration, and error code; it does **not** log message bodies, attachment contents, or passwords
- excludes `.env`, tests, caches, and VCS metadata from Docker build context via `.dockerignore`

What you should still be careful with:
- `EXCHANGE_VERIFY_SSL=false` disables TLS certificate verification and should be used only for trusted internal/self-signed environments
- `get_attachment` writes files to disk, so choose a safe destination directory
- `outlook-mcp-smoke` is privacy-safe by default and prints only masked mailbox info plus counts; set `OUTLOOK_MCP_SMOKE_INCLUDE_DATA=true` only if you explicitly want real inbox/event data in stdout
- if you enable file logging with `LOG_FILE`, protect that file with OS permissions
- if you publish Docker images from CI, protect GitLab/GitHub project access and registry permissions

## Quick start

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
- `ATTACHMENT_MAX_SIZE_MB` is enforced before sending/replying/forwarding/drafting with local attachments

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

## GitLab CI/CD

The repository includes `.gitlab-ci.yml` with these stages:

- `lint` — runs `ruff`
- `test` — runs `pytest`
- `build` — builds Python package artifacts into `dist/`
- `release` — builds and pushes a Docker image to the GitLab Container Registry on the default branch and on tags

Default image tagging behavior:
- default branch: pushes `:$CI_COMMIT_SHORT_SHA` and `:latest`
- git tag: pushes `:$CI_COMMIT_TAG`

GitLab built-in registry variables are used:
- `CI_REGISTRY`
- `CI_REGISTRY_USER`
- `CI_REGISTRY_PASSWORD`
- `CI_REGISTRY_IMAGE`

## Development

```bash
uv run --python 3.12 --with '.[dev]' ruff check .
uv run --python 3.12 --with '.[dev]' pytest -q
```

## Project notes

- The implementation is centered around a single `ExchangeClient` abstraction so auth, transport, retries, and error mapping stay centralized.
- Errors are returned in a structured JSON form suitable for MCP `isError=true` handling.
- `RTK.md` is not present in this repository, so the implementation follows `PLAN.md`.
