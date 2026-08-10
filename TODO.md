# Missing Exchange methods — implementation plan

Scope note: this list was revised after checking each item against the installed
`exchangelib==5.6.0` API. Two ideas from the original discussion turned out to be
non-issues or infeasible and were dropped/replaced:
- "search across all folders" already exists (`search_emails` with `folder=None`
  walks every mail folder) — nothing to add.
- "edit a single occurrence of a recurring event" and "propose a new meeting time"
  have no clean, verifiable support in this exchangelib version (occurrence
  targeting needs a 1-based `InstanceIndex` with no date→index resolver, and
  there is no `propose_new_time`-style API on `CalendarItem`) — building either
  without a live server to validate against was judged too risky. Replaced with
  Inbox Rules, Out-of-Office, and Room Finder, all confirmed present in the
  installed library.

Each task = one self-contained slice: models → protocol → backend impl → facade →
tools → tool_specs → `FakeExchangeBackend` → tests → README. Implemented and
committed one at a time on `feature/missing-exchange-methods`.

## 1. `update_draft` — edit an existing draft in place
- [x] `UpdateDraftRequest` (partial update via `model_fields_set`, same pattern as
      `UpdateContactRequest`): `id`, `to`, `subject`, `body`, `body_type`, `cc`,
      `bcc`, `attachments` (replaces the full attachment set when provided).
- [x] `ExchangeBackend.update_draft` / `ExchangeClient.update_draft`.
- [x] EWS impl in `email.py`: fetch from `account.drafts`, apply provided fields,
      re-attach files when `attachments` is set, `save()`.
- [x] Tool `update_draft`, reuses the attachment validation hook from `create_draft`.
- [x] Tests + README.

## 2. `get_email_mime` — export a message as raw MIME
- [x] `GetEmailMimeRequest(id)` / `EmailMimeResult(id, filename, content_type,
      size, mime_base64)`.
- [x] EWS impl: `item.mime_content` (bytes) → base64.
- [x] Read-only tool `get_email_mime`.
- [x] Tests + README.

## 3. Bulk email actions
- [x] `BulkMoveEmailsRequest`, `BulkDeleteEmailsRequest`, `BulkMarkEmailsRequest`,
      `BulkCategorizeEmailsRequest` — each `ids: list[str]` plus the same fields
      as its singular counterpart.
- [x] Backend: thin loop over the existing single-item methods; a per-item
      `APIError` becomes `ActionResult(id=.., status="error", warning=message)`
      instead of aborting the whole batch.
- [x] Tools `bulk_move_emails`, `bulk_delete_emails`, `bulk_mark_emails`,
      `bulk_categorize_emails`, all `destructive=True`.
- [x] Tests + README.

## 4. Inbox Rules
- [x] `MailRule` (id, display_name, priority, is_enabled, from_addresses,
      contains_subject_strings, has_attachments, move_to_folder, mark_as_read,
      assign_categories, delete, stop_processing_rules), `CreateRuleRequest`,
      `UpdateRuleRequest` (full replace — matches `SetInboxRule` semantics,
      not a partial patch), `DeleteRuleRequest`.
- [x] EWS impl using `account.rules` / `account.create_rule` / `account.set_rule`
      / `account.delete_rule` (`exchangelib.properties.Rule/Conditions/Actions`).
- [x] Tools `list_rules`, `create_rule`, `update_rule`, `delete_rule`.
- [x] Tests + README.

## 5. Out-of-Office (autoresponder)
- [x] `OofSettingsModel` (state, external_audience, start, end, internal_reply,
      external_reply), `SetOofSettingsRequest` mirrors it with the same
      scheduled/non-disabled validation exchangelib enforces server-side.
- [x] EWS impl via `account.oof_settings` getter/setter.
- [x] Tools `get_oof_settings` (read-only), `set_oof_settings`.
- [x] Tests + README.

## 6. Room Finder
- [x] `RoomListInfo(name, email)`, `RoomInfo(name, email)`, `ListRoomsRequest
      (room_list: EmailStr)`.
- [x] EWS impl via `account.protocol.get_roomlists()` / `.get_rooms(room_list)`.
- [x] Read-only tools `list_room_lists`, `list_rooms`.
- [x] Tests + README.

## 7. View a shared/delegate calendar
- [x] Add optional `mailbox: EmailStr | None` to `ListEventsRequest` and
      `GetEventRequest`; reject combining it with `calendar_id` (keeps
      `_resolve_folder` untouched — mailbox scoping only ever targets that
      mailbox's default calendar).
- [x] `BaseEWSBackend._account_for(mailbox)`: caches one extra `Account` per
      mailbox on the backend instance, reusing the existing protocol config.
- [x] `_fetch_item` gains an optional `account` param; `list_events`/`get_event`/
      `get_my_availability` thread `mailbox` through.
- [x] Documented simplification: the query window still uses the service
      account's timezone, not the target mailbox's.
- [x] Tests + README.

## 8. Bulk calendar actions
- [x] `BulkDeleteEventsRequest`, `BulkRespondToInvitesRequest` — same
      per-item-loop-with-error-capture pattern as task 3.
- [x] Tools `bulk_delete_events`, `bulk_respond_to_invites`.
- [x] Tests + README.

## Wrap-up
- [x] `ruff check .`, `ruff format .`, `mypy src`, `pytest -q --cov=outlook_mcp
      --cov-fail-under=70` all green after every task.
- [x] README tool catalog + Highlights updated.
