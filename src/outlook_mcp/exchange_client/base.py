from __future__ import annotations

import logging
import re
import threading
from collections.abc import Callable, Iterable
from contextvars import ContextVar
from datetime import UTC, datetime
from typing import Any, Literal
from urllib.parse import urlparse

from exchangelib import version as ews_version
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
    ErrorAccessDenied,
    ErrorExceededConnectionCount,
    ErrorFolderSavePropertyError,
    ErrorInternalServerTransientError,
    ErrorIrresolvableConflict,
    ErrorItemNotFound,
    ErrorItemSavePropertyError,
    ErrorServerBusy,
    ErrorTimeoutExpired,
    ErrorTooManyObjectsOpened,
    RateLimitError,
    ResponseMessageError,
    TransportError,
    UnknownTimeZone,
    UnauthorizedError,
)
from exchangelib.folders import FolderCollection
from exchangelib.items import Item
from exchangelib.properties import ItemId
from exchangelib.protocol import BaseProtocol, FailFast
from exchangelib.version import Build, Version

from .. import telemetry
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
from ..models import (
    BulkItemFailure,
    BulkItemResult,
    BulkResult,
    EmailAddress,
    MailboxInfo,
    PingResult,
)

logger = logging.getLogger(__name__)


def _ews_response_hook(response: Any, **kwargs: Any) -> None:
    """Time and count one EWS HTTP request.

    requests measures `elapsed` itself (send to headers-complete), so no clock
    handling here. Only the URL *path* is logged -- never the query string,
    which is where anything sensitive would live. The body is deliberately not
    touched: reading response.content on a streamed download would pull the
    whole attachment into memory. A request that raises before a response
    exists (connection refused, TLS, timeout) never reaches a response hook;
    those failures are already logged when _map_exception surfaces them.
    """
    elapsed_ms = response.elapsed.total_seconds() * 1000
    telemetry.record_request(elapsed_ms)
    logger.debug(
        "ews request path=%s status=%s elapsed_ms=%.0f",
        urlparse(response.url).path,
        response.status_code,
        elapsed_ms,
    )


_TIMEZONE_FALLBACK_PATCHED = False
#: Guards _TIMEZONE_FALLBACK_PATCHED: the patch itself is idempotent, but without a
#: lock, concurrent first calls (ToolGateway runs tool bodies in a thread pool once
#: MCP_MAX_CONCURRENCY > 1) can both observe it unset and both proceed to patch --
#: harmless here, but the same check-then-act shape as the account lock below, so
#: it gets the same treatment rather than relying on the patch being a no-op twice.
_TIMEZONE_FALLBACK_LOCK = threading.Lock()
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
        #: Guards the check-then-act below. ToolGateway runs tool bodies in a thread
        #: pool (anyio.to_thread), so with MCP_MAX_CONCURRENCY > 1 an unguarded first
        #: burst of concurrent calls could each see self._account as None and race to
        #: build their own Account/Configuration/protocol, wasting an auth round trip
        #: per racer and leaking every loser's protocol object.
        self._account_lock = threading.Lock()
        self._other_accounts: dict[str, Account] = {}

    @property
    def account(self) -> Account:
        if self._account is None:
            with self._account_lock:
                if self._account is None:
                    self._account = self._build_account()
        _active_timezone_fallback.set(self.settings.exchange_timezone_fallback)
        return self._account

    def _build_account(self) -> Account:
        auth = build_auth_context(self.settings)
        self._configure_timezone_fallback()
        # Always FailFast: exchangelib's own FaultTolerance retries a request until a
        # single backoff step exceeds max_wait, not until total elapsed time does, so
        # a server that keeps answering "busy, retry in N seconds" (N under max_wait)
        # never stops retrying -- max_wait bounds one step, not the operation. Bounded,
        # cumulative retrying lives in ExchangeClient._retry_read instead, and only
        # wraps read-only calls (see exchange_client/facade.py); _map_exception below
        # marks which raw exceptions are safe for that loop to retry on.
        retry_policy = FailFast()
        credentials = Credentials(username=auth.username, password=auth.password)
        auth_type = BASIC if auth.auth_type == "Basic" else NTLM
        service_endpoint = self._normalize_service_endpoint(self.settings.exchange_server)
        config = Configuration(
            service_endpoint=service_endpoint,
            credentials=credentials,
            auth_type=auth_type,
            retry_policy=retry_policy,
            version=self._resolve_configured_version(),
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

    def _account_for(self, mailbox: str) -> Account:
        """A secondary Account for viewing another mailbox's calendar.

        Reuses the primary account's already-authenticated Configuration (and,
        via exchangelib's own protocol caching, its underlying Protocol/session
        pool) rather than re-authenticating per mailbox. Requires the service
        account to already have delegate or impersonation rights on `mailbox`
        -- this only changes which mailbox is addressed, not what access the
        credentials actually grant.
        """
        if mailbox not in self._other_accounts:
            primary = self.account
            auth = build_auth_context(self.settings)
            access_type = IMPERSONATION if auth.impersonate_as else DELEGATE
            try:
                other = Account(
                    primary_smtp_address=mailbox,
                    config=primary.protocol.config,
                    autodiscover=False,
                    access_type=access_type,
                )
            except Exception as exc:  # noqa: BLE001
                raise self._map_exception(exc) from exc
            self._configure_protocol(other.protocol)
            self._other_accounts[mailbox] = other
        return self._other_accounts[mailbox]

    def _resolve_configured_version(self) -> Version | None:
        """Turn EXCHANGE_VERSION (e.g. "EXCHANGE_2016") into a Configuration(version=...).

        Skipping this made EXCHANGE_VERSION a documented setting that did
        nothing: exchangelib always fell back to autodetecting the server
        version instead. Leaving it unset here preserves that autodetection.
        """
        name = self.settings.exchange_version
        if not name:
            return None
        build = getattr(ews_version, name, None)
        if not isinstance(build, Build):
            raise APIError(
                "validation_error",
                f"unknown EXCHANGE_VERSION: {name!r}",
                details=[
                    {
                        "field": "EXCHANGE_VERSION",
                        "reason": (
                            "must match an exchangelib.version.EXCHANGE_* constant, "
                            "e.g. EXCHANGE_2016"
                        ),
                    }
                ],
            )
        return Version(build=build)

    def _normalize_service_endpoint(self, value: str) -> str:
        # Strip the trailing slash before the suffix check: a URL copied from a
        # browser ends in "Exchange.asmx/", which used to miss the check and get
        # a second "/EWS/Exchange.asmx" appended.
        endpoint = value.strip().rstrip("/")
        if "://" not in endpoint:
            endpoint = f"https://{endpoint}"
        if not endpoint.lower().endswith("/ews/exchange.asmx"):
            endpoint = f"{endpoint}/EWS/Exchange.asmx"
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

        def raw_session_with_verify(
            prefix, oauth2_client=None, oauth2_session_params=None, oauth2_token_endpoint=None
        ):
            session = original_raw_session(
                prefix,
                oauth2_client=oauth2_client,
                oauth2_session_params=oauth2_session_params,
                oauth2_token_endpoint=oauth2_token_endpoint,
            )
            session.verify = verify_ssl
            # Every session the pool ever creates -- including renewals and
            # credential refreshes -- passes through raw_session, so this one
            # hook covers the protocol's whole lifetime.
            session.hooks["response"].append(_ews_response_hook)
            return session

        protocol.raw_session = raw_session_with_verify

        if not verify_ssl:
            logger.warning("SSL certificate verification is disabled for Exchange connections")

    def _configure_timezone_fallback(self) -> None:
        global _TIMEZONE_FALLBACK_PATCHED
        if _TIMEZONE_FALLBACK_PATCHED:
            return

        with _TIMEZONE_FALLBACK_LOCK:
            if _TIMEZONE_FALLBACK_PATCHED:
                return

            original = EWSTimeZone.from_ms_id.__func__

            def from_ms_id_with_fallback(cls, ms_id):
                try:
                    return original(cls, ms_id)
                except UnknownTimeZone:
                    fallback_timezone = _active_timezone_fallback.get()
                    if (
                        fallback_timezone
                        and isinstance(ms_id, str)
                        and _GUID_TIMEZONE_RE.match(ms_id)
                    ):
                        logger.info("Mapping unknown Exchange timezone id to configured fallback")
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
            except ResponseMessageError:
                # A scoped EWS response fault (e.g. no archive mailbox configured
                # for this account) -- safe to fall back to root.
                archive_root = account.root
            except Exception as exc:  # noqa: BLE001
                # Auth/network/throttling failures must propagate instead of
                # silently pretending the archive folder is root.
                raise self._map_exception(exc) from exc
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

        parts = [segment for segment in value.strip("/").split("/") if segment]
        if not parts:
            return account.root

        # A name or path may start at a well-known folder ("Inbox/Projects") or at
        # the mailbox top level ("Projects"). Neither lives under account.root,
        # whose children are Top of Information Store plus dozens of hidden system
        # folders -- user folders hang off TOIS itself.
        candidates: list[tuple[Folder, list[str]]] = []
        if parts[0].lower() in builtin:
            candidates.append((builtin[parts[0].lower()], parts[1:]))
        try:
            candidates.append((account.root.tois, parts))
        except Exception:  # noqa: BLE001
            # Not every mailbox exposes TOIS, or grants access to it.
            logger.debug("Top of Information Store is unavailable; falling back to root")
        candidates.append((account.root, parts))

        for start, remainder in candidates:
            found = self._walk_child_folders(start, remainder)
            if found is not None:
                return found
        raise NotFoundError(value)

    def _walk_child_folders(self, start: Folder, parts: list[str]) -> Folder | None:
        current = start
        for part in parts:
            try:
                children = current.children
            except Exception:  # noqa: BLE001
                return None
            next_folder = next(
                (child for child in children if child.name.lower() == part.lower()),
                None,
            )
            if next_folder is None:
                return None
            current = next_folder
        return current if current is not start else None

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
        except ResponseMessageError:
            # A scoped EWS response fault about this specific id (malformed, not
            # found, ...) -- safe to treat as "not an id" and fall through to a
            # path-based lookup.
            return None
        except Exception as exc:  # noqa: BLE001
            # A failure of the GetFolder call itself (auth, network, throttling) is not
            # scoped to this one id and must propagate rather than fall through to a
            # path-based lookup that can only ever produce a misleading not_found.
            raise self._map_exception(exc, item_id=folder_id) from exc
        for item in resolved:
            if item is None or isinstance(item, Exception):
                return None
            return item
        return None

    def _fetch_item(
        self,
        item_id: str,
        folder: Folder | None = None,
        expected_type: type[Item] | tuple[type[Item], ...] | None = None,
        account: Account | None = None,
        only_fields: tuple[str, ...] | None = None,
    ) -> Any:
        """Fetch one item by id.

        ``only_fields`` projects the fetch down to those fields (id and changekey
        always come along). Callers that only invoke a method on the item --
        delete, send, accept -- have no use for the full body and attachment
        metadata a bare GetItem downloads; they pass ("parent_folder_id",), the
        one field the folder re-check below actually reads.
        """
        target_account = account if account is not None else self.account
        fetch_kwargs: dict[str, Any] = {
            "ids": [ItemId(id=item_id, changekey=None)],
            "folder": folder,
        }
        if only_fields is not None:
            fetch_kwargs["only_fields"] = list(only_fields)
        try:
            item = next(target_account.fetch(**fetch_kwargs))
        except StopIteration as exc:
            raise NotFoundError(item_id) from exc
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=item_id) from exc
        if isinstance(item, Exception):
            raise self._map_exception(item, item_id=item_id)
        # Account.fetch() only uses `folder` to validate `only_fields` against that
        # folder's allowed fields -- it does NOT restrict which item the id resolves
        # to. Any item id in the mailbox fetches successfully regardless of the
        # `folder` passed here, so callers relying on `folder` as a scoping boundary
        # (contacts, calendar events, drafts) need this explicit re-check after the
        # fetch, or a calendar_id/folder argument silently fails to constrain which
        # item gets acted on.
        if (
            isinstance(item, Item)
            and expected_type is not None
            and not isinstance(item, expected_type)
        ):
            raise NotFoundError(item_id)
        if folder is not None and not self._item_in_folder(item, folder):
            raise NotFoundError(item_id)
        return item

    def _item_in_folder(self, item: Any, folder: Folder) -> bool:
        parent_id = getattr(getattr(item, "parent_folder_id", None), "id", None)
        folder_id = getattr(folder, "id", None)
        return parent_id is not None and folder_id is not None and parent_id == folder_id

    def _mailbox(self, address: str) -> Mailbox:
        return Mailbox(email_address=address)

    def _email_address(self, mailbox: Any) -> EmailAddress:
        if mailbox is None:
            return EmailAddress(email="unknown@example.invalid", name=None)
        email = (
            getattr(mailbox, "email_address", None)
            or getattr(mailbox, "email", None)
            or "unknown@example.invalid"
        )
        name = getattr(mailbox, "name", None)
        return EmailAddress(email=email, name=name)

    def _recipients(self, values: Iterable[Any] | None) -> list[EmailAddress]:
        return [self._email_address(value) for value in values or []]

    def _normalize_importance(self, value: Any) -> Literal["low", "normal", "high"]:
        normalized = str(value or "normal").lower()
        if normalized in ("low", "normal", "high"):
            return normalized  # type: ignore[return-value]
        return "normal"

    def _normalize_free_busy_status(
        self, value: Any
    ) -> Literal["free", "tentative", "busy", "oof", "working_elsewhere", "unknown"]:
        # EWS's own choices are Free/Tentative/Busy/OOF/NoData/WorkingElsewhere.
        # Anything we don't recognize is treated as "busy", the same
        # conservative default exchangelib itself uses for new appointments.
        mapping: dict[str, Literal["free", "tentative", "busy", "oof", "working_elsewhere"]] = {
            "free": "free",
            "tentative": "tentative",
            "busy": "busy",
            "oof": "oof",
            "workingelsewhere": "working_elsewhere",
        }
        normalized = str(value or "").lower()
        return mapping.get(normalized, "busy" if normalized != "nodata" else "unknown")

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

    #: Defensive bound on a single header's value -- internet headers are normally
    #: well under this, but nothing stops a message from arriving with a pathological
    #: one (e.g. a hostile X-* header) that would otherwise inflate the MCP response.
    _MAX_HEADER_VALUE_LENGTH = 4096

    def _headers_to_dict(self, headers: Any) -> dict[str, str]:
        result: dict[str, str] = {}
        for header in headers or []:
            name = getattr(header, "name", None)
            value = getattr(header, "value", None)
            if name and value is not None:
                result[str(name)] = str(value)[: self._MAX_HEADER_VALUE_LENGTH]
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
        if isinstance(exc, APIError):
            return exc
        if isinstance(exc, UnauthorizedError):
            return AuthFailedError()
        if isinstance(exc, ErrorAccessDenied):
            return PermissionDeniedError()
        if isinstance(exc, RateLimitError):
            rate_limit_error = ExchangeUnavailableError(
                "exchange throttling or rate limit encountered"
            )
            rate_limit_error.retryable = True
            return rate_limit_error
        if isinstance(exc, ErrorIrresolvableConflict):
            # The only one of this group that's an actual optimistic-concurrency
            # conflict (item was modified since it was last synced).
            logger.warning("Exchange conflict")
            return ConflictError()
        if isinstance(exc, ErrorTimeoutExpired):
            timeout_error = TimeoutAPIError(self.settings.exchange_timeout)
            timeout_error.retryable = True
            return timeout_error
        if isinstance(
            exc,
            (
                ErrorServerBusy,
                ErrorInternalServerTransientError,
                ErrorTooManyObjectsOpened,
                ErrorExceededConnectionCount,
            ),
        ):
            # FailFast means these surface on the first occurrence instead of being
            # retried internally by exchangelib -- safe for ExchangeClient to retry
            # itself on read-only calls, bounded by a deadline (see facade.py).
            logger.warning("Exchange reported itself busy: %s", type(exc).__name__)
            busy_error = ExchangeUnavailableError("exchange reported itself busy; try again")
            busy_error.retryable = True
            return busy_error
        if isinstance(exc, ErrorItemNotFound):
            if item_id:
                return NotFoundError(item_id)
            logger.warning("Exchange item not found")
            return APIError("not_found", "item was not found")
        if isinstance(exc, (ErrorItemSavePropertyError, ErrorFolderSavePropertyError)):
            # An invalid/read-only property in the save request -- not a concurrent-
            # edit conflict, so this must not be reported as ConflictError.
            logger.warning("Exchange save-property error")
            return APIError(
                "exchange_error", "exchange rejected one or more properties in the request"
            )
        # ResponseMessageError subclasses TransportError, so it has to be matched first.
        if isinstance(exc, ResponseMessageError):
            logger.warning("Unmapped Exchange response error: %s", type(exc).__name__)
            return APIError("exchange_error", "exchange rejected the request")
        if isinstance(exc, TimeoutError):
            return TimeoutAPIError(self.settings.exchange_timeout)
        if isinstance(exc, TransportError):
            logger.warning("Exchange transport error: %s", type(exc).__name__)
            return ExchangeUnavailableError("exchange is unavailable")
        logger.warning("Unmapped Exchange exception: %s", type(exc).__name__)
        return ExchangeUnavailableError("exchange is unavailable")

    def _bulk(self, ids: Iterable[str], action: Callable[[str], Any]) -> BulkResult:
        """Run a single-item action over every id, capturing per-item failures.

        For operations with no native EWS batch equivalent (marking, categorizing,
        responding to invites, ...): Exchange has no transactional multi-item
        guarantee for these anyway, so a failure on one id is reported in `failed`
        instead of aborting ids that would otherwise have succeeded.
        """
        succeeded: list[BulkItemResult] = []
        failed: list[BulkItemFailure] = []
        for item_id in ids:
            try:
                action(item_id)
                succeeded.append(BulkItemResult(id=item_id))
            except APIError as exc:
                failed.append(BulkItemFailure(id=item_id, error=exc.code, message=exc.message))
        return BulkResult(succeeded=succeeded, failed=failed)

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
        return PingResult(
            status="ok",
            server=parsed.netloc or self.settings.exchange_server,
            version=version,
            latency_ms=latency_ms,
        )

    def get_mailbox_info(self) -> MailboxInfo:
        account = self.account
        version = getattr(getattr(account.protocol, "version", None), "api_version", None)
        return MailboxInfo(
            email_address=account.primary_smtp_address,
            display_name=account.fullname or account.primary_smtp_address,
            timezone=str(account.default_timezone),
            exchange_version=version,
        )
