from __future__ import annotations

import logging
import re
import warnings
from collections.abc import Iterable
from contextvars import ContextVar
from datetime import UTC, datetime
from typing import Any
from urllib.parse import urlparse

from exchangelib import (
    Account,
    BASIC,
    Configuration,
    Credentials,
    DELEGATE,
    EWSTimeZone,
    Folder,
    HTMLBody,
    IMPERSONATION,
    Mailbox,
    NTLM,
)
from exchangelib.ewsdatetime import EWSDateTime
from exchangelib.errors import (
    ErrorItemSavePropertyError,
    ErrorFolderSavePropertyError,
    RateLimitError,
    ResponseMessageError,
    TransportError,
    UnknownTimeZone,
    UnauthorizedError,
)
from exchangelib.folders import FolderCollection
from exchangelib.properties import ItemId
from exchangelib.protocol import BaseProtocol, FailFast, FaultTolerance
from urllib3.exceptions import InsecureRequestWarning

from ..auth import build_auth_context
from ..config import Settings
from ..errors import (
    APIError,
    AuthFailedError,
    ConflictError,
    ExchangeUnavailableError,
    NotFoundError,
    PermissionDeniedError,
    TimeoutAPIError,
)
from ..models import EmailAddress, MailboxInfo, PingResult

logger = logging.getLogger(__name__)
_TIMEZONE_FALLBACK_PATCHED = False
_GUID_TIMEZONE_RE = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"
)
# EWSTimeZone.from_ms_id is a bare classmethod with no reference back to the account/backend
# that triggered the parse, so the per-instance fallback timezone can't be passed as an argument.
# Track it here instead: whichever backend's `.account` was most recently touched (on this thread/
# async task) is the one whose fallback applies, since EWS response parsing always happens
# synchronously underneath that access.
_active_timezone_fallback: ContextVar[str | None] = ContextVar(
    "_active_timezone_fallback", default=None
)


class BaseEWSBackend:
    """Shared account/auth/error-mapping infra used by every domain mixin."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self._account: Account | None = None

    @property
    def account(self) -> Account:
        if self._account is None:
            self._account = self._build_account()
        _active_timezone_fallback.set(self.settings.exchange_timezone)
        return self._account

    def _build_account(self) -> Account:
        auth = build_auth_context(self.settings)
        if auth.auth_type == "OAuth2":
            raise APIError(
                "validation_error",
                "OAuth2 is not wired in this build yet",
                details=[{"field": "EXCHANGE_AUTH_TYPE", "reason": "supported values for live checks are NTLM or Basic"}],
            )

        self._configure_timezone_fallback()
        retry_policy = (
            FailFast()
            if self.settings.exchange_max_retry_wait == 0
            else FaultTolerance(max_wait=self.settings.exchange_max_retry_wait)
        )
        credentials = Credentials(username=auth.username, password=auth.password)
        auth_type = BASIC if auth.auth_type == "Basic" else NTLM
        service_endpoint = self._normalize_service_endpoint(self.settings.exchange_server)
        config = Configuration(
            service_endpoint=service_endpoint,
            credentials=credentials,
            auth_type=auth_type,
            retry_policy=retry_policy,
        )
        access_type = IMPERSONATION if auth.impersonate_as else DELEGATE
        try:
            account = Account(
                primary_smtp_address=auth.primary_smtp_address,
                config=config,
                autodiscover=False,
                access_type=access_type,
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        self._configure_protocol(account.protocol)
        return account

    def _normalize_service_endpoint(self, value: str) -> str:
        endpoint = value.strip()
        if "://" not in endpoint:
            endpoint = f"https://{endpoint}"
        if not endpoint.lower().endswith("/ews/exchange.asmx"):
            endpoint = endpoint.rstrip("/") + "/EWS/Exchange.asmx"
        return endpoint

    def _configure_protocol(self, protocol: BaseProtocol) -> None:
        """Scope timeout/SSL-verification settings to this backend's own protocol instance.

        These used to be set as mutations on the shared `BaseProtocol` class, which meant the
        first backend built in a process silently decided the behavior for every later one.
        Setting them directly on the instance shadows the class attribute/classmethod (both are
        non-data descriptors, so an instance attribute of the same name wins) without touching
        any other backend's protocol.
        """
        protocol.TIMEOUT = self.settings.exchange_timeout

        verify_ssl = self.settings.exchange_verify_ssl
        original_raw_session = protocol.raw_session

        def raw_session_with_verify(prefix, oauth2_client=None, oauth2_session_params=None, oauth2_token_endpoint=None):
            session = original_raw_session(
                prefix,
                oauth2_client=oauth2_client,
                oauth2_session_params=oauth2_session_params,
                oauth2_token_endpoint=oauth2_token_endpoint,
            )
            session.verify = verify_ssl
            return session

        protocol.raw_session = raw_session_with_verify

        if not verify_ssl:
            warnings.filterwarnings("ignore", category=InsecureRequestWarning)
            logger.warning("SSL certificate verification is disabled for Exchange connections")

    def _configure_timezone_fallback(self) -> None:
        global _TIMEZONE_FALLBACK_PATCHED
        if _TIMEZONE_FALLBACK_PATCHED:
            return

        original = EWSTimeZone.from_ms_id.__func__

        def from_ms_id_with_fallback(cls, ms_id):
            try:
                return original(cls, ms_id)
            except UnknownTimeZone:
                fallback_timezone = _active_timezone_fallback.get()
                if fallback_timezone and isinstance(ms_id, str) and _GUID_TIMEZONE_RE.match(ms_id):
                    logger.info(
                        "Mapping unknown Exchange timezone id %s to configured timezone %s",
                        ms_id,
                        fallback_timezone,
                    )
                    return cls(fallback_timezone)
                raise

        EWSTimeZone.from_ms_id = classmethod(from_ms_id_with_fallback)
        _TIMEZONE_FALLBACK_PATCHED = True

    def _resolve_folder(self, value: str | None) -> Folder:
        account = self.account
        if not value or value == "root":
            return account.root
        normalized = value.strip("/").lower()
        archive_root = account.root
        if normalized == "archive":
            try:
                archive_root = account.archive_root
            except Exception:  # noqa: BLE001
                archive_root = account.root
        builtin = {
            "inbox": account.inbox,
            "sent": account.sent,
            "sentitems": account.sent,
            "drafts": account.drafts,
            "deleted": account.trash,
            "trash": account.trash,
            "junk": account.junk,
            "archive": archive_root,
            "calendar": account.calendar,
            "contacts": account.contacts,
        }
        if normalized in builtin:
            return builtin[normalized]

        by_id = self._get_folder_by_id(value)
        if by_id is not None:
            return by_id

        current = account.root
        for part in [segment for segment in value.strip("/").split("/") if segment]:
            next_folder = next(
                (child for child in current.children if child.name.lower() == part.lower()),
                None,
            )
            if next_folder is None:
                raise NotFoundError(value)
            current = next_folder
        return current

    def _get_folder_by_id(self, folder_id: str) -> Folder | None:
        """Resolve a folder by EWS id with a single targeted GetFolder call.

        Deliberately avoids account.root.walk(), which recurses through every
        folder in the mailbox and does not scale on accounts with large folder
        trees (see issue #5).
        """
        try:
            resolved = list(
                FolderCollection(account=self.account, folders=[Folder(id=folder_id)]).resolve()
            )
        except Exception:  # noqa: BLE001
            return None
        for item in resolved:
            if item is None or isinstance(item, Exception):
                return None
            return item
        return None

    def _fetch_item(self, item_id: str, folder: Folder | None = None) -> Any:
        try:
            item = next(self.account.fetch(ids=[ItemId(id=item_id, changekey=None)], folder=folder))
        except StopIteration as exc:
            raise NotFoundError(item_id) from exc
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=item_id) from exc
        if isinstance(item, Exception):
            raise self._map_exception(item, item_id=item_id)
        return item

    def _mailbox(self, address: str) -> Mailbox:
        return Mailbox(email_address=address)

    def _email_address(self, mailbox: Any) -> EmailAddress:
        if mailbox is None:
            return EmailAddress(email="unknown@example.invalid", name=None)
        email = getattr(mailbox, "email_address", None) or getattr(mailbox, "email", None) or "unknown@example.invalid"
        name = getattr(mailbox, "name", None)
        return EmailAddress(email=email, name=name)

    def _recipients(self, values: Iterable[Any] | None) -> list[EmailAddress]:
        return [self._email_address(value) for value in values or []]

    def _normalize_importance(self, value: Any) -> str:
        normalized = str(value or "normal").lower()
        return normalized if normalized in {"low", "normal", "high"} else "normal"

    def _extract_message_body(self, item: Any) -> tuple[str, str | None]:
        text = ""
        html = None
        if getattr(item, "text_body", None):
            text = str(item.text_body)
        elif getattr(item, "body", None):
            text = str(item.body)

        body = getattr(item, "body", None)
        if isinstance(body, HTMLBody):
            html = str(body)
        elif body is not None and "</" in str(body):
            html = str(body)
        return text, html

    def _headers_to_dict(self, headers: Any) -> dict[str, str]:
        result: dict[str, str] = {}
        for header in headers or []:
            name = getattr(header, "name", None)
            value = getattr(header, "value", None)
            if name and value is not None:
                result[str(name)] = str(value)
        return result

    def _to_ews_datetime(self, value: datetime) -> EWSDateTime:
        if value.tzinfo is None:
            value = value.replace(tzinfo=self.account.default_timezone)
        else:
            value = value.astimezone(self.account.default_timezone)
        if isinstance(value, EWSDateTime):
            return value
        return EWSDateTime.from_datetime(value)

    def _map_exception(self, exc: Exception, item_id: str | None = None) -> APIError:
        message = str(exc)
        if isinstance(exc, APIError):
            return exc
        if isinstance(exc, UnauthorizedError):
            return AuthFailedError()
        if isinstance(exc, RateLimitError):
            return ExchangeUnavailableError("exchange throttling or rate limit encountered")
        if isinstance(exc, (ErrorItemSavePropertyError, ErrorFolderSavePropertyError)):
            return ConflictError(message)
        # ResponseMessageError subclasses TransportError, so it has to be matched first.
        if isinstance(exc, ResponseMessageError):
            lowered = message.lower()
            if "not found" in lowered and item_id:
                return NotFoundError(item_id)
            if "access is denied" in lowered or "permission" in lowered:
                return PermissionDeniedError()
            return APIError("exchange_error", message)
        if isinstance(exc, (TransportError, TimeoutError)):
            if "timed out" in message.lower():
                return TimeoutAPIError(self.settings.exchange_timeout)
            return ExchangeUnavailableError(message)
        return ExchangeUnavailableError(message)

    def ping(self) -> PingResult:
        started = datetime.now(UTC)
        account = self.account
        try:
            inbox = account.inbox
            count = inbox.total_count
            del count
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        latency_ms = round((datetime.now(UTC) - started).total_seconds() * 1000)
        parsed = urlparse(self.settings.exchange_server)
        version = getattr(getattr(account.protocol, "version", None), "api_version", None)
        return PingResult(status="ok", server=parsed.netloc or self.settings.exchange_server, version=version, latency_ms=latency_ms)

    def get_mailbox_info(self) -> MailboxInfo:
        account = self.account
        version = getattr(getattr(account.protocol, "version", None), "api_version", None)
        return MailboxInfo(
            email_address=account.primary_smtp_address,
            display_name=account.fullname or account.primary_smtp_address,
            timezone=str(account.default_timezone),
            exchange_version=version,
        )
