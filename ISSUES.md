# Known issues / improvement backlog

Findings from a full codebase review (2026-08-03). Each item is a standalone,
independently-committable task. Ordered by priority within each section.

## Security

### 1. `get_attachment` allows writing to any directory on disk
- **Where:** `src/outlook_mcp/models.py:215` (`GetAttachmentRequest.save_path`),
  `src/outlook_mcp/exchange_client/email.py:333-353` (`get_attachment`)
- **Problem:** `save_path` is caller-supplied (i.e. LLM-controlled) and passed straight
  to `Path()` with no root restriction. `_sanitize_attachment_filename` only cleans the
  *filename*, not the directory. README documents this as the caller's responsibility,
  but nothing enforces it server-side — a prompt-injected email could steer the
  assistant into writing a file to an arbitrary path the process can reach.
- **Fix:**
  - [ ] Add `EXCHANGE_ATTACHMENT_SAVE_ROOT` (or similar) to `Settings` (`config.py`),
        documented per `AGENTS.md` conventions (README env table + `.env.example`).
  - [ ] In `get_attachment`, resolve `target_dir` and reject (`APIError` /
        `validation_error`) if `target_dir.resolve()` is not relative to the
        configured root. Default root: system temp dir (current fallback) if unset.
  - [ ] Add a test covering rejection of a path outside the root and traversal via
        `save_path` (e.g. `../../etc`).

### 2. Outgoing attachments read arbitrary local files (undocumented exfiltration path)
- **Where:** `src/outlook_mcp/models.py:147,155,162,199` (`attachments: list[Path]` on
  `SendEmailRequest`/`ReplyEmailRequest`/`ForwardEmailRequest`/`DraftEmailRequest`),
  `src/outlook_mcp/exchange_client/email.py:107-110` (`_attach_files`)
- **Problem:** Any file readable by the process can be attached and emailed out.
  Combined with prompt injection from untrusted email content, this is a plausible
  secret-exfiltration path. Unlike `get_attachment`, this isn't called out in
  README's "Security notes" at all.
- **Fix:**
  - [ ] Add a README "Security notes" bullet documenting this explicitly.
  - [ ] Reuse the same allow-list root from item 1 (e.g.
        `EXCHANGE_ATTACHMENT_SAVE_ROOT`, or a separate `..._READ_ROOT` if the two
        should differ) to restrict which local paths can be attached.
  - [ ] Add a test covering rejection of an attachment path outside the root.

### 3. `get_attachment` has no size cap, unlike outgoing attachments
- **Where:** `src/outlook_mcp/exchange_client/email.py:333-353`,
  compare `src/outlook_mcp/tools/email.py:153-179` (`_validate_attachments`, only
  wired into send/reply/forward/create_draft)
- **Problem:** `ATTACHMENT_MAX_SIZE_MB` is enforced on outgoing attachments but not
  when downloading incoming ones — an oversized attachment can be written to disk
  without limit.
- **Fix:**
  - [ ] In `get_attachment`, check `attachment.size` (or the fetched content length)
        against `settings.attachment_max_size_mb` before writing; raise
        `validation_error` (or a new dedicated error) if it exceeds the limit.
  - [ ] Add a test with a mock attachment over the configured limit.

## Correctness

### 4. `search_emails` silently swallows all exceptions on the first filter pass
- **Where:** `src/outlook_mcp/exchange_client/email.py:181-196`
- **Problem:** The bare `except Exception: items = []` around the
  `subject__icontains` query treats *any* failure (auth error, bad query, connectivity
  issue) as "no results," then falls through to the `text_body__icontains` pass. The
  second pass, by contrast, correctly maps exceptions via `self._map_exception`. This
  hides real errors as empty results.
- **Fix:**
  - [ ] Narrow the first `except` to the same exception set used elsewhere
        (`RateLimitError, TransportError, TimeoutError, UnauthorizedError`) or to the
        specific "field not filterable" error the fallback is meant to catch, and map
        everything else via `self._map_exception`.
  - [ ] Add a regression test asserting an auth failure during `search_emails` raises
        `AuthFailedError` rather than returning `[]`.

### 5. Global monkeypatch of `BaseProtocol.raw_session` / `EWSTimeZone.from_ms_id` only honors the first backend's settings
- **Where:** `src/outlook_mcp/exchange_client/base.py:118-166`
  (`_configure_ssl_verification`, `_configure_timezone_fallback`,
  module-level `_RAW_SESSION_PATCHED` / `_TIMEZONE_FALLBACK_PATCHED`)
- **Problem:** Both patches close over `self.settings` but only apply once per
  process (guarded by a module-global flag). A second `BaseEWSBackend` built with a
  different `exchange_verify_ssl` or `exchange_timezone` silently gets the *first*
  instance's behavior instead of its own — e.g. `EXCHANGE_VERIFY_SSL=false` could be
  ignored if another backend with `true` was constructed earlier in the same process.
  Currently masked by `get_settings()` being an `lru_cache`d singleton, but it's a
  process-global side effect standing in for what should be per-instance
  configuration, and it has zero test coverage today (no test exercises
  `_build_account`/`_configure_ssl_verification` — confirmed via grep).
- **Fix:**
  - [ ] Replace the monkeypatched classmethods with per-instance configuration — e.g.
        pass a `session_factory`/adapter into `Configuration`, or wrap
        `Account.protocol` post-construction — so behavior is scoped to the backend
        instance, not the process.
  - [ ] Add a test constructing two `EWSExchangeBackend`s with different
        `exchange_verify_ssl` values in the same process and asserting each keeps its
        own setting.

