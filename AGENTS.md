# AGENTS.md

## Engineering rules
- Do not preserve backward compatibility unless explicitly asked — this is a pre-1.0 beta project, rename/break freely.
- Choose the simplest implementation that fully meets the current requirements. No speculative abstractions, feature flags, or hooks for hypothetical future needs.
- Prefer established, well-maintained libraries over custom implementations (e.g. `exchangelib`, `pydantic`, `mcp` — don't hand-roll what they already solve).
- Fix the root cause, not the symptom. If a bug points at a deeper design issue, say so instead of patching around it.
- Suggest best practices even if they require refactoring — flag them rather than silently working around bad structure.

## Project shape
MCP server exposing on-prem Exchange (EWS) to LLM clients. Source lives under `src/outlook_mcp/`:
- `config.py` — `Settings` (pydantic-settings, reads `.env`, `EXCHANGE_*` / `MCP_*` env vars)
- `auth.py` — NTLM/Basic credential wiring for `exchangelib`
- `exchange_client/` — the Exchange abstraction, split by domain: `base.py` (session/TLS setup), `backend.py` (`EWSExchangeBackend`, `build_default_backend`), `email.py`, `calendar.py`, `contacts.py`, `protocol.py` (`ExchangeBackend` interface), `unconfigured.py` (fallback when no credentials), `facade.py` (`ExchangeClient`, the single entry point everything else uses)
- `models.py` — pydantic request/response models for every MCP tool
- `errors.py` — `APIError` and subclasses (`ExchangeUnavailableError`, `AuthFailedError`, ...); `.to_dict()` feeds MCP `isError=true` responses
- `tools/` — MCP tool definitions by domain (`email.py`, `calendar.py`, `contacts.py`, `system.py`), registered via `mcp_tools.py`
- `server.py` — `FastMCP` app entrypoint (`stdio` or `sse` transport)
- `smoke.py` — privacy-safe manual smoke check against a real mailbox (`OUTLOOK_MCP_SMOKE_INCLUDE_DATA=true` to see real data)

Package/CLI/repo are all `outlook-ews-mcp` now — `outlook-mcp` was already taken on PyPI, so the PyPI distribution, the two console scripts, and the GitHub repo itself were renamed to match. The Python import path stays `outlook_mcp` (unrenamed, and not worth churning across every file).

## Testing
- Dev deps are an optional extra, not a dependency-group: `uv sync --extra dev` (NOT `uv sync --group dev`, that errors).
- Run tests: `.venv/bin/python -m pytest -q` or `uv run --extra dev pytest -q`. If a bare `pytest` on PATH fails with `ModuleNotFoundError: No module named '_pytest.scope'`, that's a mise-managed global pytest/anyio-plugin mismatch shadowing the project venv — use the venv's pytest directly instead of debugging the global install.
- Tests never hit a real Exchange server: `tests/conftest.py` provides `FakeExchangeBackend` implementing the `ExchangeBackend` protocol. Extend that fake rather than mocking `exchangelib` directly.
- Lint: `.venv/bin/python -m ruff check .` (line length 100, configured in `pyproject.toml`).
- Both checks also run in CI: GitHub Actions (`.github/workflows/ci.yml`) and GitLab CI (`.gitlab-ci.yml`, additionally builds the package and publishes a Docker image on `main`/tags).

## Conventions
- Errors from `exchange_client/` should surface as `APIError` subclasses so `errors.py` can map them to structured MCP error payloads — don't let raw `exchangelib` exceptions escape to `tools/`.
- Keep `ExchangeClient` (`exchange_client/facade.py`) as the only thing `tools/*.py` talks to; domain modules (`email.py`, `calendar.py`, `contacts.py`) stay internal to `exchange_client/`.
- New Settings fields go in `config.py` with an `EXCHANGE_*`/`MCP_*` alias and must be documented in `README.md`'s env var table and `.env.example`.
