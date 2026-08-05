# Contributing

## Dev setup

```bash
uv venv
source .venv/bin/activate
uv pip install -e .[dev]
cp .env.example .env
```

## Running checks

```bash
uv run --python 3.12 --with '.[dev]' ruff check .
uv run --python 3.12 --with '.[dev]' ruff format --check .
uv run --python 3.12 --with '.[dev]' pytest -q
```

All run in CI (GitHub Actions and GitLab CI) on every push and pull/merge request.

## Testing without a real Exchange server

None of the test suite talks to a real Exchange server. Tests fake the EWS layer in one of two ways:

- `tests/conftest.py` defines `FakeExchangeBackend`, a stand-in for `ExchangeBackend` used to exercise the MCP tool layer (`tests/test_mcp_tools.py`, `tests/test_registry.py`, ...) without touching `exchangelib` at all.
- `tests/test_ews_backend.py` and `tests/test_backend_edge_cases.py` exercise `EWSExchangeBackend` directly by assigning a `SimpleNamespace`/fake object to `backend._account`, and by monkeypatching `exchangelib` classes (`ResolveNames`, `FolderCollection`, ...) imported into `outlook_mcp.exchange_client`.

Follow the same pattern for new tests: build the smallest fake that supports the attributes/methods your code path touches, and assert both the returned model and any request captured from the fake.

## Pull requests

- Keep PRs focused — one fix or feature per PR.
- Add or update tests for any behavior change; tests, lint, and format checks must pass.
- Update `README.md` if you change configuration, tool behavior, or the tool catalog.
- Describe what real-server behavior prompted the change when fixing an Exchange/EWS-specific bug (timezone handling, address formats, etc.) — these are easy to regress silently.

## Reporting bugs

Open a GitHub issue with the tool name, the request you sent, the error returned (mask any secrets/PII), and your Exchange version if known.
