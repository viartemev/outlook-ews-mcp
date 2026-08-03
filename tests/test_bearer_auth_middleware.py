from __future__ import annotations

import httpx
import pytest

from outlook_mcp.server import BearerTokenMiddleware

pytestmark = pytest.mark.anyio


async def _dummy_app(scope, receive, send):
    await send({"type": "http.response.start", "status": 200, "headers": []})
    await send({"type": "http.response.body", "body": b"ok"})


def _client(token: str = "s3cret") -> httpx.AsyncClient:
    app = BearerTokenMiddleware(_dummy_app, token=token)
    return httpx.AsyncClient(transport=httpx.ASGITransport(app=app), base_url="http://test")


async def test_rejects_missing_authorization_header():
    async with _client() as client:
        response = await client.get("/sse")
    assert response.status_code == 401
    assert response.headers["www-authenticate"] == "Bearer"


async def test_rejects_wrong_token():
    async with _client() as client:
        response = await client.get("/sse", headers={"Authorization": "Bearer wrong"})
    assert response.status_code == 401


async def test_accepts_correct_bearer_token():
    async with _client() as client:
        response = await client.get("/sse", headers={"Authorization": "Bearer s3cret"})
    assert response.status_code == 200
    assert response.text == "ok"
