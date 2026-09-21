"""Render Outlook/OWA mention anchors, with explicit recipient metadata.

OWAAM anchors and X-Mentions headers are Outlook client conventions, not a
documented EWS mentions API. Both are needed: a link alone does not set the @
indicator. Keep actual Outlook recognition covered by live verification.
"""

from __future__ import annotations

import re
from html import escape
from html.parser import HTMLParser
from typing import TYPE_CHECKING
from uuid import uuid4

if TYPE_CHECKING:
    from .models import Mention

_TOKEN = re.compile(r"\{\{mention:([A-Za-z0-9_-]+)\}\}")


class _MentionText(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=False)
        self.keys: list[str] = []
        self.blocked = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"a", "script", "style"}:
            self.blocked += 1

    def handle_endtag(self, tag: str) -> None:
        if tag in {"a", "script", "style"}:
            self.blocked = max(0, self.blocked - 1)

    def handle_data(self, data: str) -> None:
        if not self.blocked:
            self.keys.extend(_TOKEN.findall(data))


def validate_mentions(body: str, body_type: str, mentions: list[Mention]) -> None:
    keys = [mention.key for mention in mentions]
    tokens = _TOKEN.findall(body)
    if len(keys) != len(set(keys)):
        raise ValueError("mention keys must be unique")
    if set(tokens) != set(keys):
        raise ValueError(
            "every {{mention:key}} must match a mention, and every mention needs a token"
        )
    if body_type == "html" and mentions:
        parser = _MentionText()
        parser.feed(body)
        if sorted(parser.keys) != sorted(tokens):
            raise ValueError(
                "mention tokens must be in HTML text, outside links, scripts and styles"
            )


def render_mentions(body: str, body_type: str, mentions: list[Mention]) -> tuple[str, str]:
    validate_mentions(body, body_type, mentions)
    if not mentions:
        return body, body_type
    if body_type == "text":
        body = escape(body).replace("\n", "<br>")
    by_key = {mention.key: mention for mention in mentions}

    def replace(match: re.Match[str]) -> str:
        mention = by_key[match.group(1)]
        return (
            f'<a id="OWAAM{uuid4().hex.upper()}" href="mailto:{escape(str(mention.email))}">'
            f'<span style="text-decoration:none">@{escape(mention.display_name)}</span></a>'
        )

    return _TOKEN.sub(replace, body), "html"
