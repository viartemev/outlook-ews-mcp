# syntax=docker/dockerfile:1

# ghcr.io/astral-sh/uv:latest
FROM ghcr.io/astral-sh/uv:latest@sha256:cf4eedcaa81655197f625739489effcbe71b61ceb1506f332c3facae5deceded AS uv

# python:3.12-slim
FROM python:3.12-slim@sha256:57cd7c3a7a273101a6485ba99423ee568157882804b1124b4dd04266317710de AS builder
COPY --from=uv /uv /uvx /bin/

WORKDIR /app
COPY pyproject.toml uv.lock README.md ./
COPY src ./src

# --frozen refuses to run if uv.lock is out of sync with pyproject.toml, so the
# image always installs exactly what CI tested and pip-audit scanned.
RUN uv sync --frozen --no-dev --no-editable

FROM python:3.12-slim@sha256:57cd7c3a7a273101a6485ba99423ee568157882804b1124b4dd04266317710de
RUN groupadd --gid 1000 outlook-mcp && useradd --uid 1000 --gid outlook-mcp --create-home outlook-mcp

WORKDIR /app
COPY --from=builder --chown=outlook-mcp:outlook-mcp /app/.venv ./.venv

ENV PATH="/app/.venv/bin:${PATH}"
USER outlook-mcp

CMD ["outlook-ews-mcp"]
