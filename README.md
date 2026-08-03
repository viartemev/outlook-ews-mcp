# outlook-ews-mcp

![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![MCP](https://img.shields.io/badge/MCP-server-7c3aed)
![Exchange](https://img.shields.io/badge/Microsoft%20Exchange-EWS-0a7ea4)
![Status](https://img.shields.io/badge/status-beta-orange)
![License](https://img.shields.io/badge/license-MIT-green)

`outlook-ews-mcp` is an MCP server for on-prem Microsoft Exchange via EWS (`exchangelib`).
It gives MCP-compatible clients access to email, calendar, contacts, folders, attachments, and availability data through a single, testable Python service.

> Note: this project was previously referred to as `outlook-mcp`. It was renamed because that name is already taken on PyPI by an unrelated project — use `outlook-ews-mcp` for `pip`/`uvx` installs and for the CLI commands below.

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
- `EXCHANGE_AUTH_TYPE=Basic` sends credentials in the clear, so the server refuses to start against an `http://` `EXCHANGE_SERVER`; only override with `EXCHANGE_ALLOW_INSECURE_BASIC_AUTH=true` for a local/test server you control
- `get_attachment` writes files to disk, so choose a safe destination directory
- `send_email`/`reply_email`/`forward_email`/`create_draft` read local files (via `attachments`) and attach their contents to outgoing mail — combined with untrusted email content, this is a plausible path for a prompt-injected exfiltration of any file readable by the process; set `EXCHANGE_ATTACHMENT_ROOT` to restrict both which files can be attached and where `get_attachment` may write downloads
- `outlook-ews-mcp-smoke` is privacy-safe by default and prints only masked mailbox info plus counts; set `OUTLOOK_MCP_SMOKE_INCLUDE_DATA=true` only if you explicitly want real inbox/event data in stdout
- if you enable file logging with `LOG_FILE`, protect that file with OS permissions
- if you publish Docker images from CI, protect GitLab/GitHub project access and registry permissions

## Quick start

```bash
uv venv
source .venv/bin/activate
uv pip install -e .[dev]
cp .env.example .env
outlook-ews-mcp
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
EXCHANGE_ALLOW_INSECURE_BASIC_AUTH=false
EXCHANGE_VERSION=EXCHANGE_2016
EXCHANGE_TIMEOUT=30
EXCHANGE_MAX_RETRY_WAIT_SECONDS=90
EXCHANGE_TIMEZONE=Europe/Moscow
EXCHANGE_IMPERSONATE_AS=
ATTACHMENT_MAX_SIZE_MB=10
EXCHANGE_ATTACHMENT_ROOT=
MCP_TRANSPORT=stdio
MCP_SSE_HOST=127.0.0.1
MCP_SSE_PORT=8080
LOG_LEVEL=INFO
LOG_FILE=
```

Notes:
- set `EXCHANGE_EMAIL_ADDRESS` when `EXCHANGE_USERNAME` is not an SMTP address
- `OAuth2` is reserved in config, but this build currently supports live auth with `NTLM` or `Basic`
- `EXCHANGE_ALLOW_INSECURE_BASIC_AUTH` only matters with `EXCHANGE_AUTH_TYPE=Basic` and an `http://` `EXCHANGE_SERVER`; startup fails otherwise unless it's set `true`
- `EXCHANGE_IMPERSONATE_AS` enables mailbox impersonation when Exchange permissions are configured accordingly
- `ATTACHMENT_MAX_SIZE_MB` is enforced both on local files attached to outgoing email and on attachments downloaded via `get_attachment`
- `EXCHANGE_ATTACHMENT_ROOT` is unset (unrestricted) by default; set it to an absolute directory to confine both `attachments` paths (send/reply/forward/create_draft) and `get_attachment`'s `save_path` to that directory tree
- `EXCHANGE_MAX_RETRY_WAIT_SECONDS` is a total backoff time budget (exchangelib retries transient errors with exponential backoff until this many seconds have elapsed), not a retry count; set to `0` to disable retries and fail fast on the first error

## Claude Desktop example

```json
{
  "mcpServers": {
    "outlook": {
      "command": "outlook-ews-mcp",
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
outlook-ews-mcp-smoke
```

Default output is sanitized for safer verification. If you intentionally want sample mailbox/event data in the output:

```bash
OUTLOOK_MCP_SMOKE_INCLUDE_DATA=true outlook-ews-mcp-smoke
```

## Docker

```bash
docker build -t outlook-ews-mcp .
docker run --rm --env-file .env outlook-ews-mcp
```

## CI/CD

The repository includes both a GitHub Actions workflow (`.github/workflows/ci.yml`) and a GitLab CI pipeline (`.gitlab-ci.yml`), each running the same `lint` (ruff) and `test` (pytest) stages on pushes and pull/merge requests.

GitLab CI additionally has:
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

## Contributing

Bug reports and PRs are welcome — see [CONTRIBUTING.md](CONTRIBUTING.md) for how to set up a dev environment and run the test suite without a real Exchange server. For vulnerability reports, see [SECURITY.md](SECURITY.md).

## License

MIT — see [LICENSE](LICENSE).
