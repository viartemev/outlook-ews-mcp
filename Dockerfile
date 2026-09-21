# syntax=docker/dockerfile:1

# ghcr.io/astral-sh/uv:latest
FROM ghcr.io/astral-sh/uv:latest@sha256:10787c682e4184e4f290de1171fd4703dc63de99221f10fe1c99002ce7fa9acc AS uv

# python:3.12-slim
FROM python:3.14-slim@sha256:cad9a2c871761c413caa6fdd6441c783451e740a48aaeba60ae62a8b53525ef6 AS builder
COPY --from=uv /uv /uvx /bin/

WORKDIR /app
COPY pyproject.toml uv.lock README.md ./
COPY src ./src

# --frozen refuses to run if uv.lock is out of sync with pyproject.toml, so the
# image always installs exactly what CI tested and pip-audit scanned.
RUN uv sync --frozen --no-dev --no-editable

FROM python:3.14-slim@sha256:cad9a2c871761c413caa6fdd6441c783451e740a48aaeba60ae62a8b53525ef6
RUN groupadd --gid 1000 outlook-mcp && useradd --uid 1000 --gid outlook-mcp --create-home outlook-mcp

WORKDIR /app
COPY --from=builder --chown=outlook-mcp:outlook-mcp /app/.venv ./.venv

ENV PATH="/app/.venv/bin:${PATH}"
USER outlook-mcp

CMD ["outlook-ews-mcp"]
