# syntax=docker/dockerfile:1

# ghcr.io/astral-sh/uv:latest
FROM ghcr.io/astral-sh/uv:latest@sha256:4a1f3666816dafadd20ceb71eac9a7d56d7ef30279f7d64048519ab901ada336 AS uv

# python:3.12-slim -- the newest version the CI test matrix covers; keep in sync.
FROM python:3.12-slim@sha256:191083cdf6c7bda8c76158983163ead5d5cf7b77cf3033e805bb0f5296f51b4b AS builder
COPY --from=uv /uv /uvx /bin/

WORKDIR /app
COPY pyproject.toml uv.lock README.md ./
COPY src ./src

# --frozen refuses to run if uv.lock is out of sync with pyproject.toml, so the
# image always installs exactly what CI tested and pip-audit scanned.
RUN uv sync --frozen --no-dev --no-editable

FROM python:3.12-slim@sha256:191083cdf6c7bda8c76158983163ead5d5cf7b77cf3033e805bb0f5296f51b4b
RUN groupadd --gid 1000 outlook-mcp && useradd --uid 1000 --gid outlook-mcp --create-home outlook-mcp

WORKDIR /app
COPY --from=builder --chown=outlook-mcp:outlook-mcp /app/.venv ./.venv

ENV PATH="/app/.venv/bin:${PATH}"
USER outlook-mcp

CMD ["outlook-ews-mcp"]
