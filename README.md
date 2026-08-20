# outlook-ews-mcp

<p align="right"><b>English</b> · <a href="README.ru.md">Русский</a></p>

![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![MCP](https://img.shields.io/badge/MCP-server-7c3aed)
![Exchange](https://img.shields.io/badge/Microsoft%20Exchange-EWS-0a7ea4)
![Status](https://img.shields.io/badge/status-beta-orange)
![License](https://img.shields.io/badge/license-MIT-green)

`outlook-ews-mcp` is an [MCP](https://modelcontextprotocol.io) server for on-prem Microsoft
Exchange via EWS ([`exchangelib`](https://github.com/ecederstrand/exchangelib)). It gives
MCP-compatible clients (Claude Desktop, Claude Code, and any other MCP client) access to
email, calendar, contacts, folders, attachments, and availability data through a single,
testable Python service — no direct mailbox scripting required.

> **Renamed from `outlook-mcp`.** That name was already taken on PyPI by an unrelated
> project, so the distribution and CLI name are now `outlook-ews-mcp`. The Python import
> path is unchanged. Until the first tagged PyPI release, install from this repository as
> shown below.

## Contents

- [Highlights](#highlights)
- [Tool catalog](#tool-catalog)
- [Typical use cases](#typical-use-cases)
- [Security notes](#security-notes)
- [Quick start](#quick-start)
- [Configuration](#configuration)
- [Request queue](#request-queue)
- [Claude Desktop example](#claude-desktop-example)
- [Smoke check](#smoke-check)
- [Docker](#docker)
- [CI/CD](#cicd)
- [Development](#development)
- [Project notes](#project-notes)
- [Contributing](#contributing)
- [License](#license)

## Highlights

- **Email** — list, search (substring or Advanced Query Syntax), read, send, reply,
  forward, move, copy, delete, mark, categorize, bulk actions, raw MIME export,
  attachment add/delete
- **System** — Inbox Rules, Out-of-Office (automatic replies), read-only delegate listing
- **Calendar** — list, create, update, delete, respond to invites, find free slots, view a
  shared/delegate mailbox's calendar, Room Finder, bulk actions
- **Contacts** — search, read, create, update, delete
- **Folders & attachments** — folder CRUD and attachment download
- **Auth** — `NTLM` and `Basic` against on-prem Exchange
- **Transport** — `stdio` and `SSE`
- **Architecture** — centralized error mapping through a single `ExchangeClient`
  abstraction (see [Project notes](#project-notes))
- **Safety** — a privacy-safer smoke check by default (see [Smoke check](#smoke-check))
- **Ops** — Docker image plus GitHub and GitLab CI/CD pipelines included

## Tool catalog

Every tool below is registered in [`tool_specs.py`](src/outlook_mcp/tool_specs.py), the
single source of truth for its name, description, and schema. **Read-only** marks tools
that never modify the mailbox — they get more concurrency (see
[Request queue](#request-queue)) and are safe to call speculatively.

### System

| Tool | Description | Read-only |
| --- | --- | --- |
| `ping_exchange` | Check connectivity to Exchange | ✅ |
| `get_mailbox_info` | Get mailbox metadata | ✅ |
| `list_delegates` | List mailbox delegates and their folder permission levels — read-only because `exchangelib` has no delegate write support | ✅ |
| `list_inbox_rules` | List server-side inbox rules | ✅ |
| `create_inbox_rule` | Create a server-side inbox rule, e.g. "from this sender → move to folder" | |
| `update_inbox_rule` | Enable/disable a rule or change its priority (other fields aren't updatable here) | |
| `delete_inbox_rule` | Delete a server-side inbox rule by id | |
| `get_out_of_office` | Get the out-of-office (automatic reply) settings | ✅ |
| `set_out_of_office` | Turn automatic replies off, on, or schedule a start/end window | |

> ⚠️ `create_inbox_rule` / `update_inbox_rule` / `delete_inbox_rule` manage rules over
> EWS, which removes the client-side rule blob desktop Outlook keeps — this can wipe
> rules a user created in Outlook itself. This is documented EWS behavior, not a bug here.

### Email

| Tool | Description | Read-only |
| --- | --- | --- |
| `list_emails` | List emails in a folder | ✅ |
| `get_email` | Get a full email by id | ✅ |
| `get_email_mime` | Export a message's raw RFC 822 MIME content, base64-encoded | ✅ |
| `get_thread` | Get every message of a conversation in order, bodies included | ✅ |
| `search_emails` | Search by substring (subject/body/sender) or server-side Advanced Query Syntax | ✅ |
| `send_email` | Send a new email | |
| `reply_email` | Reply to an email | |
| `forward_email` | Forward an email | |
| `move_email` | Move an email to another folder | |
| `copy_email` | Copy an email to another folder | |
| `move_emails` | Bulk move, with per-item results — one bad id doesn't fail the rest | |
| `copy_emails` | Bulk copy, with per-item results | |
| `delete_emails` | Bulk delete, with per-item results (soft-deletes unless `hard_delete`) | |
| `delete_email` | Delete an email | |
| `mark_email` | Update read state, importance, or the follow-up flag | |
| `categorize_email` | Set, add, or remove Outlook categories (the coloured labels) | |
| `mark_emails` | Bulk version of `mark_email`, with per-item results | |
| `categorize_emails` | Bulk version of `categorize_email`, with per-item results | |
| `list_categories` | List categories in use with counts, sampled from recent messages (not the mailbox master category list) | ✅ |
| `list_folders` | List mailbox folders | ✅ |
| `create_folder` | Create a mailbox folder | |
| `rename_folder` | Rename a folder — refuses built-in folders (Inbox, Sent Items, Calendar, ...) | |
| `delete_folder` | Delete a folder and everything in it — refuses built-in folders | |
| `create_draft` | Create an email draft | |
| `update_draft` | Update a draft; omitted fields are left unchanged, `attachments` (if given) replaces the whole set | |
| `send_draft` | Send an existing draft | |
| `add_attachment` | Attach a local file to a message, typically a draft — the file must live under `EXCHANGE_ATTACHMENT_ROOT` | |
| `delete_attachment` | Remove one attachment from a message by id | |
| `get_attachment` | Save an attachment to disk | ✅ |

### Calendar

| Tool | Description | Read-only |
| --- | --- | --- |
| `list_events` | List calendar events in a time range; pass `mailbox` for a colleague's default calendar (needs delegate/impersonation access, not combinable with `calendar_id`) | ✅ |
| `get_event` | Get a calendar event by id; pass `mailbox` for a colleague's calendar | ✅ |
| `create_event` | Create a calendar event | |
| `update_event` | Update a calendar event | |
| `delete_event` | Delete a calendar event | |
| `respond_to_invite` | Accept, decline, or tentatively respond to an invite | |
| `find_free_slots` | Find open meeting time slots | ✅ |
| `delete_events` | Bulk delete events, with per-item results | |
| `respond_to_invites` | Bulk respond to invites, with per-item results | |
| `get_my_availability` | Get free/busy slots; pass `mailbox` for a colleague's calendar | ✅ |
| `list_calendars` | List calendars | ✅ |
| `list_room_lists` | List Room Finder room lists (groups of meeting rooms) | ✅ |
| `list_rooms` | List the meeting rooms in a Room Finder room list | ✅ |

### Contacts

| Tool | Description | Read-only |
| --- | --- | --- |
| `search_contacts` | Search contacts | ✅ |
| `get_contact` | Get a contact by id | ✅ |
| `create_contact` | Create a personal contact | |
| `update_contact` | Update a personal contact | |
| `delete_contact` | Delete a personal contact | |

## Typical use cases

- Connect Claude Desktop or another MCP client to on-prem Exchange
- Search inbox messages and fetch full email content
- Send or draft emails from AI workflows
- Inspect calendars and create meetings
- Check free/busy windows for scheduling
- Search personal contacts or the Global Address List
- Expose Exchange operations through a controlled MCP boundary instead of direct mailbox scripting

## Security notes

**What the current code does:**

| | |
| --- | --- |
| Scoped connectivity | Connects only to the Exchange/EWS endpoint configured in `EXCHANGE_SERVER` |
| No telemetry | Contains no telemetry, analytics, or third-party data export logic |
| Secrets stay local | Keeps secrets in environment variables / `.env`, ignored by `.gitignore` (`.env`, `.env.*`, while keeping `.env.example`) |
| Clean error payloads | Structured MCP error responses never include raw Exchange exception text, message bodies, attachment contents, or passwords; successful tools return only the mailbox data they were asked for |
| Clean logs | `LOG_LEVEL` only controls the app's own `outlook_mcp.*` loggers; `exchangelib`'s SOAP XML loggers — which would otherwise dump full request/response XML, even at `ERROR` level on transport errors — are always force-silenced |
| Clean Docker builds | `.dockerignore` excludes `.env`, tests, caches, and VCS metadata from the build context |

**What you should still be careful with:**

- `EXCHANGE_VERIFY_SSL=false` disables TLS certificate verification — trusted internal/self-signed environments only.
- `EXCHANGE_AUTH_TYPE=Basic` sends credentials in the clear, so the server refuses to
  start against an `http://` `EXCHANGE_SERVER`; only override with
  `EXCHANGE_ALLOW_INSECURE_BASIC_AUTH=true` for a local/test server you control.
- `get_attachment` writes files to disk, and `send_email`/`reply_email`/`forward_email`/
  `create_draft` read local files (via `attachments`) and attach their contents to
  outgoing mail. Combined with untrusted email content, this is a plausible path for
  prompt-injected exfiltration of any file readable by the process. Local file access is
  **refused by default** and only works once `EXCHANGE_ATTACHMENT_ROOT` is set to an
  absolute directory, which then confines both `attachments` paths and `get_attachment`'s
  `save_path` to that directory tree (an unset `save_path` still falls back to the system
  temp directory).
- `outlook-ews-mcp-smoke` is privacy-safe by default and prints only masked mailbox info
  plus counts; set `OUTLOOK_MCP_SMOKE_INCLUDE_DATA=true` only if you explicitly want real
  inbox/event data in stdout.
- If you enable file logging with `LOG_FILE`, protect that file with OS permissions.
- If you publish Docker images from CI, protect GitLab/GitHub project access and
  registry permissions.

## Quick start

```bash
uv venv
source .venv/bin/activate
uv pip install -e .[dev]
cp .env.example .env
outlook-ews-mcp
```

By default the server runs in `stdio` mode. Set `MCP_TRANSPORT=sse` to start an HTTP
server instead.

## Configuration

Minimal `.env` to get started — everything else below has a working default:

```dotenv
EXCHANGE_SERVER=https://mail.company.com/EWS/Exchange.asmx
EXCHANGE_USERNAME=DOMAIN\username
EXCHANGE_PASSWORD=secret
EXCHANGE_EMAIL_ADDRESS=user@company.com
EXCHANGE_AUTH_TYPE=NTLM
```

A fully commented copy of every variable lives in [`.env.example`](.env.example).

| Variable | Default | Description |
| --- | --- | --- |
| `EXCHANGE_SERVER` | *(required)* | EWS endpoint URL, e.g. `https://mail.company.com/EWS/Exchange.asmx` |
| `EXCHANGE_USERNAME` | *(required)* | `DOMAIN\username` or a UPN. Exactly one backslash — dotenv does not process escape sequences |
| `EXCHANGE_PASSWORD` | *(required)* | Account password |
| `EXCHANGE_EMAIL_ADDRESS` | unset | SMTP address; set when `EXCHANGE_USERNAME` isn't one |
| `EXCHANGE_AUTH_TYPE` | `NTLM` | `NTLM` or `Basic` |
| `EXCHANGE_ALLOW_INSECURE_BASIC_AUTH` | `false` | Allow `Basic` auth over `http://` — local/test servers only |
| `EXCHANGE_VERIFY_SSL` | `true` | Verify the server's TLS certificate; `false` only for trusted internal/self-signed setups |
| `EXCHANGE_VERSION` | unset (auto-detected) | Exchange server version, e.g. `EXCHANGE_2016` |
| `EXCHANGE_TIMEZONE_FALLBACK` | `Europe/Moscow` | Used only when Exchange reports an unresolvable GUID timezone id; normal operations use the mailbox's own default timezone |
| `EXCHANGE_TIMEOUT` | `30` | Per-request timeout in seconds (1–300) |
| `EXCHANGE_MAX_RETRY_WAIT_SECONDS` | `90` | Wall-clock retry budget for **read-only** calls when Exchange reports itself busy, not a retry count; `0` disables retries. Writes are never auto-retried |
| `EXCHANGE_IMPERSONATE_AS` | unset | Mailbox to impersonate (requires Exchange impersonation permissions) |
| `EXCHANGE_ATTACHMENT_MAX_SIZE_MB` | `10` | Max size per attachment, enforced on both upload and `get_attachment` download (1–100) |
| `EXCHANGE_ATTACHMENT_MAX_COUNT` | `10` | Max attachments on a single send/reply/forward/create_draft call (1–100) |
| `EXCHANGE_ATTACHMENT_MAX_TOTAL_SIZE_MB` | `25` | Max combined attachment size on a single call (1–500) |
| `EXCHANGE_ATTACHMENT_ROOT` | unset (disabled) | Directory that confines attachment paths. Unset **refuses** all local file access for `attachments`/`save_path`; set to an absolute directory to allow paths inside it |
| `EXCHANGE_EMAIL_BODY_MAX_CHARS` | `200000` | Cap on `get_email`'s `body_text`/`body_html` (1,000–5,000,000); longer bodies are truncated with `truncated: true` |
| `EXCHANGE_EMAIL_MIME_MAX_SIZE_MB` | `25` | Cap on raw MIME export size before base64 expansion (1–100) |
| `EXCHANGE_SIGNATURE_TEXT` | unset | Appended to outgoing text bodies and replies/forwards. No EWS signature API exists, so this is configuration, not the mailbox's Outlook signature |
| `EXCHANGE_SIGNATURE_HTML` | unset | Appended to outgoing HTML bodies. Same caveat as above; no cross-conversion between the two. Either can be skipped per call with `include_signature: false` |
| `MCP_TRANSPORT` | `stdio` | `stdio` or `sse` |
| `MCP_SSE_HOST` | `127.0.0.1` | Bind host when `MCP_TRANSPORT=sse` |
| `MCP_SSE_PORT` | `8080` | Bind port when `MCP_TRANSPORT=sse` |
| `MCP_MAX_CONCURRENCY` | `4` | Concurrent read-only tool calls (1–8); mutating calls always run exclusively. See [Request queue](#request-queue) |
| `MCP_MAX_QUEUE_SIZE` | `20` | Max calls admitted at once, running + waiting (1–1000); beyond that, calls get an immediate `server_busy` error |
| `LOG_LEVEL` | `INFO` | `DEBUG`, `INFO`, `WARNING`, or `ERROR` |
| `LOG_FILE` | unset (stderr) | Log file path; protect it with OS permissions if set |

**Behavior notes that aren't tied to a single variable:**

- `list_events` and `find_free_slots` accept a bounded `limit` (default 200, maximum
  1000); event ranges are capped at 366 days and free-slot ranges at 31 days, so broad
  queries can't produce unbounded EWS or MCP responses.
- Listings stay lean by design: email summaries carry the sender but not recipient lists
  (`get_email` has them), `list_events` returns events without bodies (`get_event` has
  them), and `get_email` returns RFC-822 headers only with `include_headers: true`.
- Send operations return `id: null` when EWS doesn't provide a durable id for the sent
  copy (notably replies, forwards, and sent drafts).
- Attachment metadata includes `downloadable`; embedded Exchange item attachments have
  `downloadable: false` and can't be saved by `get_attachment`.

## Request queue

Clients issue several tool calls in parallel. Exchange work is blocking, so the server
runs it in worker threads and admits calls through one shared FIFO queue.

- **`MCP_MAX_CONCURRENCY`** (default `4`) sets how many *read-only* calls run at once, so
  an agent asking for an email, the folder list, and the calendar pays the slowest round
  trip instead of the sum. Mutating calls always run exclusively — one at a time, never
  overlapping a read — so read/write races on shared account state can't happen. Callers
  beyond the limit wait their turn, served in arrival order; a waiting mutation blocks
  later reads from overtaking it.
- **`MCP_MAX_QUEUE_SIZE`** (default `20`) caps how many calls can be admitted at once,
  running or waiting. Once that many are already in, further calls get an immediate
  `server_busy` error instead of joining an unbounded queue.
- **The transport stays responsive while work is in flight.** Tools are awaited rather
  than run on the event loop thread, so finished responses go out immediately and pings
  are answered while a long call is still running.
- **There is no per-call timeout, deliberately.** A thread blocked on a socket read
  can't be killed from outside; the runtime can only stop *waiting* for it, which
  abandons the thread along with the EWS session it holds. `exchangelib`'s session pool
  has a hard maximum and hands out sessions in a loop with no give-up path, so leaked
  sessions eventually starve it and every later call blocks forever. A slow call is
  waited out instead, bounded by `EXCHANGE_TIMEOUT` plus
  `EXCHANGE_MAX_RETRY_WAIT_SECONDS`: the account's retry policy is fail-fast, so every
  EWS call raises on its first transient error rather than `exchangelib` retrying it
  forever internally, and `ExchangeClient` retries only read-only calls itself, bounded
  by that wall-clock budget. Writes are never auto-retried. Overruns past the expected
  budget are logged.

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

After filling in `.env`, run:

```bash
outlook-ews-mcp-smoke
```

Default output is sanitized for safer verification. If you intentionally want sample
mailbox/event data in the output:

```bash
OUTLOOK_MCP_SMOKE_INCLUDE_DATA=true outlook-ews-mcp-smoke
```

## Docker

```bash
docker build -t outlook-ews-mcp .
docker run --rm --env-file .env outlook-ews-mcp
```

## CI/CD

GitHub Actions and GitLab CI both run lint, formatting, type checks, tests, dependency
audit, and package builds, using the `uv` version pinned in `pyproject.toml`.

| | |
| --- | --- |
| GitHub | Additionally publishes tagged releases (`v*`) to PyPI via OIDC trusted publishing. Before the first release, configure a PyPI pending publisher for repository `viartemev/outlook-ews-mcp`, workflow `ci.yml`, and environment `pypi` — no long-lived PyPI token is stored in GitHub. |
| GitLab | Additionally builds and pushes a Docker image to the GitLab Container Registry on the default branch and on tags, using the built-in `CI_REGISTRY` / `CI_REGISTRY_USER` / `CI_REGISTRY_PASSWORD` / `CI_REGISTRY_IMAGE` variables. |

Default image tagging behavior:

| Trigger | Tags pushed |
| --- | --- |
| Default branch | `:$CI_COMMIT_SHORT_SHA` and `:latest` |
| Git tag | `:$CI_COMMIT_TAG` |

## Development

```bash
uv run --python 3.12 --with '.[dev]' ruff check .
uv run --python 3.12 --with '.[dev]' pytest -q
```

## Project notes

- The implementation is centered around a single `ExchangeClient` abstraction so auth,
  transport, retries, and error mapping stay centralized.
- Errors are returned in a structured JSON form suitable for MCP `isError=true` handling.

## Contributing

Bug reports and PRs are welcome — see [CONTRIBUTING.md](CONTRIBUTING.md) for how to set
up a dev environment and run the test suite without a real Exchange server. For
vulnerability reports, see [SECURITY.md](SECURITY.md).

## License

MIT — see [LICENSE](LICENSE).
