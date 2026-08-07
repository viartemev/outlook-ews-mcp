# outlook-ews-mcp

![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![MCP](https://img.shields.io/badge/MCP-server-7c3aed)
![Exchange](https://img.shields.io/badge/Microsoft%20Exchange-EWS-0a7ea4)
![Status](https://img.shields.io/badge/status-beta-orange)
![License](https://img.shields.io/badge/license-MIT-green)

`outlook-ews-mcp` is an MCP server for on-prem Microsoft Exchange via EWS (`exchangelib`).
It gives MCP-compatible clients access to email, calendar, contacts, folders, attachments, and availability data through a single, testable Python service.

> Note: this project was previously referred to as `outlook-mcp`. It was renamed because that name is already taken on PyPI by an unrelated project. The distribution and CLI name is `outlook-ews-mcp`; until the first tagged PyPI release, install it from this repository as shown below.

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
- Docker support plus GitHub and GitLab CI/CD pipelines included

## Tool catalog

### System
- `ping_exchange`
- `get_mailbox_info`

### Email
- `list_emails`
- `get_email`
- `get_email_mime`
- `search_emails`
- `send_email`
- `reply_email`
- `forward_email`
- `move_email`
- `copy_email`
- `delete_email`
- `mark_email`
- `bulk_move_emails`
- `bulk_delete_emails`
- `bulk_mark_emails`
- `bulk_categorize_emails`
- `list_folders`
- `create_folder`
- `rename_folder`
- `delete_folder`
- `create_draft`
- `update_draft`
- `send_draft`
- `get_attachment`
- `list_rules`
- `create_rule`
- `update_rule`
- `delete_rule`

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
- structured MCP error responses do not include raw Exchange exception text, message bodies, attachment contents, or passwords; successful tools return the mailbox data they were asked for
- server logs do not include message bodies or attachment contents either: `LOG_LEVEL` only controls verbosity of the app's own `outlook_mcp.*` loggers, and `exchangelib`'s SOAP XML loggers (which would otherwise dump full request/response XML, including at `ERROR` level on unexpected transport errors) are always force-silenced regardless of `LOG_LEVEL`
- excludes `.env`, tests, caches, and VCS metadata from Docker build context via `.dockerignore`

What you should still be careful with:
- `EXCHANGE_VERIFY_SSL=false` disables TLS certificate verification and should be used only for trusted internal/self-signed environments
- `EXCHANGE_AUTH_TYPE=Basic` sends credentials in the clear, so the server refuses to start against an `http://` `EXCHANGE_SERVER`; only override with `EXCHANGE_ALLOW_INSECURE_BASIC_AUTH=true` for a local/test server you control
- `get_attachment` writes files to disk, and `send_email`/`reply_email`/`forward_email`/`create_draft` read local files (via `attachments`) and attach their contents to outgoing mail — combined with untrusted email content, this is a plausible path for a prompt-injected exfiltration of any file readable by the process; local file access is **refused by default** and only works once `EXCHANGE_ATTACHMENT_ROOT` is set to an absolute directory, which then confines both `attachments` paths and `get_attachment`'s `save_path` to that directory tree (an unset `save_path` still falls back to the system temp directory)
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
EXCHANGE_TIMEZONE_FALLBACK=Europe/Moscow
EXCHANGE_IMPERSONATE_AS=
EXCHANGE_ATTACHMENT_MAX_SIZE_MB=10
EXCHANGE_ATTACHMENT_MAX_COUNT=10
EXCHANGE_ATTACHMENT_MAX_TOTAL_SIZE_MB=25
EXCHANGE_ATTACHMENT_ROOT=
EXCHANGE_EMAIL_BODY_MAX_CHARS=200000
MCP_TRANSPORT=stdio
MCP_SSE_HOST=127.0.0.1
MCP_SSE_PORT=8080
MCP_MAX_CONCURRENCY=1
MCP_MAX_QUEUE_SIZE=20
LOG_LEVEL=INFO
LOG_FILE=
```

Notes:
- `MCP_MAX_CONCURRENCY` is how many tool calls execute at once; more calls than this queue to wait their turn, up to `MCP_MAX_QUEUE_SIZE`. See [Request queue](#request-queue).
- `MCP_MAX_QUEUE_SIZE` caps how many tool calls can be admitted at once (running + waiting); once full, further calls are rejected immediately with a `server_busy` error
- set `EXCHANGE_EMAIL_ADDRESS` when `EXCHANGE_USERNAME` is not an SMTP address
- `EXCHANGE_ALLOW_INSECURE_BASIC_AUTH` only matters with `EXCHANGE_AUTH_TYPE=Basic` and an `http://` `EXCHANGE_SERVER`; startup fails otherwise unless it's set `true`
- `EXCHANGE_IMPERSONATE_AS` enables mailbox impersonation when Exchange permissions are configured accordingly
- `EXCHANGE_TIMEZONE_FALLBACK` is only used when Exchange reports a timezone as an unresolvable GUID id; normal operations (naive datetimes, all-day event math) use the mailbox's own default timezone instead
- `EXCHANGE_ATTACHMENT_MAX_SIZE_MB` is enforced both on local files attached to outgoing email and on attachments downloaded via `get_attachment`
- `EXCHANGE_ATTACHMENT_MAX_COUNT` and `EXCHANGE_ATTACHMENT_MAX_TOTAL_SIZE_MB` cap the number and combined size of attachments on a single `send_email`/`reply_email`/`forward_email`/`create_draft` call
- `EXCHANGE_ATTACHMENT_ROOT` is unset by default, which **refuses** rather than allows local file access: any non-empty `attachments` list (send/reply/forward/create_draft) or explicit `get_attachment` `save_path` is rejected until it's set to an absolute directory, which then confines those paths to that directory tree. `get_attachment` with no `save_path` still works unset, falling back to the system temp directory
- `EXCHANGE_MAX_RETRY_WAIT_SECONDS` is a total wall-clock budget for retrying **read-only** calls when Exchange reports itself busy, not a retry count; set to `0` to disable retries and fail fast on the first error. Writes (`send_email`, `create_event`, `delete_contact`, ...) are never auto-retried, since an ambiguous failure could mean the operation already happened on the server before the error came back. See [Request queue](#request-queue).
- `EXCHANGE_EMAIL_BODY_MAX_CHARS` caps `get_email`'s `body_text`/`body_html`; a message beyond the cap is truncated and `truncated: true` is set on the response instead of returning an unbounded MCP payload
- send operations return `id: null` when EWS does not provide a durable ID for the sent copy (notably replies, forwards, and sent drafts)
- attachment metadata includes `downloadable`; embedded Exchange item attachments have `downloadable: false` and cannot be saved by `get_attachment`

## Request queue

Clients issue several tool calls in parallel. Exchange work is blocking, so the
server runs it in worker threads and admits calls through one shared FIFO queue.

- **`MCP_MAX_CONCURRENCY`** (default `1`) sets how many run at once. One mailbox,
  one conversation with Exchange, predictable load -- raise it deliberately.
  Callers beyond that wait their turn, served in the order they arrived.
- **`MCP_MAX_QUEUE_SIZE`** (default `20`) caps how many calls can be admitted at
  once, running or waiting. Once that many are already in, further calls get an
  immediate `server_busy` error instead of joining an unbounded queue.
- **The transport stays responsive while work is in flight.** Tools are awaited
  rather than run on the event loop thread, so finished responses go out
  immediately and pings are answered while a long call is still running.
- **There is no per-call timeout, deliberately.** A thread blocked on a socket
  read cannot be killed from outside; the runtime can only stop *waiting* for it,
  which abandons the thread along with the EWS session it holds. exchangelib's
  session pool has a hard maximum and hands out sessions in a loop with no
  give-up path, so leaked sessions eventually starve it and every later call
  blocks forever. A slow call is waited out instead, bounded by
  `EXCHANGE_TIMEOUT` plus `EXCHANGE_MAX_RETRY_WAIT_SECONDS`: the account's retry
  policy is fail-fast, so every EWS call raises on its first transient error
  rather than exchangelib retrying it forever internally, and `ExchangeClient`
  retries only read-only calls itself, bounded by that wall-clock budget. Writes
  are never auto-retried. Overruns past the expected budget are logged.

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

GitHub Actions and GitLab CI both run lint, formatting, type checks, tests, dependency audit, and package builds. Both use the uv version pinned in `pyproject.toml`.

GitHub additionally publishes tagged releases (`v*`) to PyPI with OIDC trusted publishing. Before the first release, configure a PyPI pending publisher for repository `viartemev/outlook-ews-mcp`, workflow `ci.yml`, and environment `pypi`; no long-lived PyPI token is stored in GitHub.

GitLab additionally builds and pushes a Docker image to the GitLab Container Registry on the default branch and on tags.

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
