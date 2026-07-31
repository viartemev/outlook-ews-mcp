from __future__ import annotations

from typing import Any

from exchangelib.indexed_properties import EmailAddress as IndexedEmailAddress
from exchangelib.indexed_properties import PhoneNumber as IndexedPhoneNumber
from exchangelib.items import Contact
from exchangelib.services import ResolveNames

from ..errors import NotFoundError
from ..models import (
    ActionResult,
    ContactEmailAddress,
    ContactFull,
    ContactSummary,
    CreateContactRequest,
    DeleteContactRequest,
    GetContactRequest,
    SearchContactsRequest,
    UpdateContactRequest,
)


class ContactOperationsMixin:
    def _contact_summary_from_contact(self, contact: Contact, source: str) -> ContactSummary:
        emails = [entry.email for entry in getattr(contact, "email_addresses", None) or [] if getattr(entry, "email", None)]
        phones = [entry.phone_number for entry in getattr(contact, "phone_numbers", None) or [] if getattr(entry, "phone_number", None)]
        return ContactSummary(
            id=contact.id,
            display_name=contact.display_name or contact.file_as or "",
            email_addresses=emails,
            phone_numbers=phones,
            company=getattr(contact, "company_name", None),
            job_title=getattr(contact, "job_title", None),
            department=getattr(contact, "department", None),
            source=source,
        )

    def search_contacts(self, request: SearchContactsRequest) -> list[ContactSummary]:
        results: list[ContactSummary] = []
        if request.source in {"personal", "all"}:
            try:
                qs = self.account.contacts.filter(display_name__icontains=request.query)[: request.limit]
                results.extend(self._contact_summary_from_contact(contact, "personal") for contact in qs)
            except Exception as exc:  # noqa: BLE001
                raise self._map_exception(exc) from exc
        if request.source in {"gal", "all"} and len(results) < request.limit:
            try:
                resolved = ResolveNames(protocol=self.account.protocol).call(
                    unresolved_entries=[request.query],
                    return_full_contact_data=True,
                    search_scope="ActiveDirectory",
                    contact_data_shape="AllProperties",
                )
                for mailbox, contact in resolved:
                    if contact is not None and getattr(contact, "id", None):
                        results.append(self._contact_summary_from_contact(contact, "gal"))
                    elif mailbox is not None:
                        results.append(
                            ContactSummary(
                                id=getattr(mailbox, "email_address", None) or request.query,
                                display_name=getattr(mailbox, "name", None) or getattr(mailbox, "email_address", None) or request.query,
                                email_addresses=[getattr(mailbox, "email_address", None)] if getattr(mailbox, "email_address", None) else [],
                                phone_numbers=[],
                                source="gal",
                            )
                        )
                    if len(results) >= request.limit:
                        break
            except Exception as exc:  # noqa: BLE001
                raise self._map_exception(exc) from exc
        return results[: request.limit]

    def get_contact(self, request: GetContactRequest) -> ContactFull:
        if "@" in request.id:
            return self._get_gal_contact(request.id)
        item = self._fetch_item(request.id, folder=self.account.contacts)
        return self._contact_full_from_item(item, item_id=item.id, source="personal")

    def _get_gal_contact(self, address: str) -> ContactFull:
        try:
            resolved = list(
                ResolveNames(protocol=self.account.protocol).call(
                    unresolved_entries=[address],
                    return_full_contact_data=True,
                    search_scope="ActiveDirectory",
                    contact_data_shape="AllProperties",
                )
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=address) from exc

        for entry in resolved:
            if isinstance(entry, Exception):
                raise self._map_exception(entry, item_id=address)
            mailbox, contact = entry
            primary = self._smtp_address(getattr(mailbox, "email_address", None)) if mailbox is not None else None
            if contact is not None:
                full = self._contact_full_from_item(contact, item_id=address, source="gal")
                if primary and all(str(known.address).lower() != primary.lower() for known in full.email_addresses):
                    full.email_addresses.insert(0, ContactEmailAddress(type="SMTP", address=primary))
                return full
            if mailbox is not None:
                return ContactFull(
                    id=address,
                    display_name=getattr(mailbox, "name", None) or primary or address,
                    email_addresses=[{"type": "SMTP", "address": primary}] if primary else [],
                    source="gal",
                )
        raise NotFoundError(address)

    @staticmethod
    def _smtp_address(value: str | None) -> str | None:
        """Drop non-SMTP proxy addresses (X500, EX, ...) and strip the `SMTP:` prefix the GAL adds."""
        address = (value or "").strip()
        prefix, separator, remainder = address.partition(":")
        if separator:
            if prefix.lower() != "smtp":
                return None
            address = remainder.strip()
        return address or None

    def _contact_full_from_item(self, item: Any, *, item_id: str, source: str) -> ContactFull:
        return ContactFull(
            id=item_id,
            display_name=getattr(item, "display_name", None) or getattr(item, "file_as", None) or "",
            first_name=getattr(item, "given_name", None),
            last_name=getattr(item, "surname", None),
            email_addresses=[
                {"type": entry.label, "address": self._smtp_address(entry.email)}
                for entry in getattr(item, "email_addresses", None) or []
                if self._smtp_address(getattr(entry, "email", None))
            ],
            phone_numbers=[
                {"type": entry.label, "number": entry.phone_number}
                for entry in getattr(item, "phone_numbers", None) or []
                if getattr(entry, "phone_number", None)
            ],
            addresses=[
                {
                    "type": entry.label,
                    "street": entry.street,
                    "city": entry.city,
                    "state": entry.state,
                    "postal_code": entry.zipcode,
                    "country": entry.country,
                }
                for entry in getattr(item, "physical_addresses", None) or []
            ],
            company=getattr(item, "company_name", None),
            job_title=getattr(item, "job_title", None),
            department=getattr(item, "department", None),
            manager=getattr(item, "manager", None),
            notes=getattr(item, "notes", None),
            birthday=getattr(item, "birthday", None),
            source=source,
        )

    def create_contact(self, request: CreateContactRequest) -> ActionResult:
        contact = Contact(
            account=self.account,
            folder=self.account.contacts,
            display_name=request.display_name,
            given_name=request.first_name,
            surname=request.last_name,
            company_name=request.company,
            job_title=request.job_title,
            notes=request.notes,
            email_addresses=[IndexedEmailAddress(label="EmailAddress1", email=str(request.email))] if request.email else [],
            phone_numbers=[IndexedPhoneNumber(label="PrimaryPhone", phone_number=request.phone)] if request.phone else [],
        )
        try:
            contact.save()
            return ActionResult(id=contact.id or "", status="created")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def update_contact(self, request: UpdateContactRequest) -> ActionResult:
        contact = self._fetch_item(request.id, folder=self.account.contacts)
        updated_fields: list[str] = []
        field_map = {
            "display_name": "display_name",
            "first_name": "given_name",
            "last_name": "surname",
            "company": "company_name",
            "job_title": "job_title",
            "notes": "notes",
        }
        for request_field, item_field in field_map.items():
            value = getattr(request, request_field)
            if value is not None:
                setattr(contact, item_field, value)
                updated_fields.append(request_field)
        if request.email is not None:
            contact.email_addresses = [IndexedEmailAddress(label="EmailAddress1", email=str(request.email))]
            updated_fields.append("email")
        if request.phone is not None:
            contact.phone_numbers = [IndexedPhoneNumber(label="PrimaryPhone", phone_number=request.phone)]
            updated_fields.append("phone")
        try:
            contact.save(update_fields=updated_fields or None)
            return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def delete_contact(self, request: DeleteContactRequest) -> ActionResult:
        contact = self._fetch_item(request.id, folder=self.account.contacts)
        try:
            contact.move_to_trash()
            return ActionResult(id=request.id, status="deleted")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc
