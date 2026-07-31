# Security Policy

## Reporting a vulnerability

Please do not open a public GitHub issue for security vulnerabilities. Instead, use [GitHub's private vulnerability reporting](https://github.com/viartemev/outlook-mcp/security/advisories/new) for this repository, or email the maintainer directly if that is unavailable.

Include:
- affected version/commit
- a minimal reproduction (config, request/tool call, expected vs. actual behavior)
- impact you believe it has (e.g. credential exposure, path traversal, auth bypass)

We aim to acknowledge reports within a few days.

## Scope and known sensitive areas

This project connects to on-prem Microsoft Exchange via EWS and handles credentials and mailbox data. Areas most likely to matter for a security report:

- `EXCHANGE_USERNAME` / `EXCHANGE_PASSWORD` handling in `auth.py` / `config.py`
- TLS verification (`EXCHANGE_VERIFY_SSL`) in `exchange_client/base.py`
- `get_attachment` writing files to disk (path handling / traversal) in `exchange_client/email.py`
- error mapping in `errors.py` potentially leaking sensitive details into MCP responses or logs

See the "Security notes" section in [README.md](README.md) for what the current code does and does not do (e.g. no telemetry, no message-body logging).

## Supported versions

This project is in beta (pre-1.0) with a single active line of development on `main`. Fixes land on `main`; there is no separate maintenance branch.
