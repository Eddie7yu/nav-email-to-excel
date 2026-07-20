from __future__ import annotations

import datetime as dt
import imaplib
import re
import ssl
from email.header import decode_header, make_header
from email.message import Message
from email.parser import BytesParser
from email.policy import default
from email.utils import getaddresses
from typing import Any

from nav_config import active_routes
from runtime_secret import read_password


class MailError(RuntimeError):
    pass


MONTHS = (
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
)
NETEASE_IMAP_SUFFIXES = ("163.com", "126.com", "yeah.net")
IMAP_CLIENT_ID = '("name" "nav-email-to-excel" "version" "1")'


def imap_date(value: dt.date) -> str:
    return f"{value.day:02d}-{MONTHS[value.month - 1]}-{value.year:04d}"


def exact_from_matches(message: Message, sender: str) -> bool:
    addresses = {
        address.casefold()
        for _, address in getaddresses(message.get_all("From", []))
        if address
    }
    return addresses == {sender.casefold()}


def single_from_address(message: Message) -> str | None:
    addresses = {
        address.strip().casefold()
        for _, address in getaddresses(message.get_all("From", []))
        if address.strip()
    }
    return next(iter(addresses)) if len(addresses) == 1 else None


def decoded(value: Any) -> str:
    try:
        return str(make_header(decode_header(str(value or ""))))
    except (LookupError, UnicodeError):
        return str(value or "")


def needs_imap_id(host: str) -> bool:
    normalized = str(host).strip().rstrip(".").casefold()
    return any(
        normalized == suffix or normalized.endswith(f".{suffix}")
        for suffix in NETEASE_IMAP_SUFFIXES
    )


def _send_required_imap_id(client: imaplib.IMAP4_SSL, host: str) -> None:
    if not needs_imap_id(host):
        return
    try:
        status, _ = client.xatom("ID", IMAP_CLIENT_ID)
    except imaplib.IMAP4.error as exc:
        raise MailError(
            "NetEase IMAP ID handshake failed before mailbox selection"
        ) from exc
    if status != "OK":
        raise MailError("NetEase IMAP ID handshake was rejected")


def _logout_safely(client: imaplib.IMAP4_SSL | None) -> None:
    if client is None:
        return
    try:
        client.logout()
    except (imaplib.IMAP4.error, OSError):
        pass


def connect(config: dict[str, Any]) -> imaplib.IMAP4_SSL:
    imap = config["imap"]
    host = str(imap["host"])
    password = read_password(str(config["runtime_id"]))
    if not password:
        raise MailError(
            "No local IMAP secret. Run `navctl.py secret set` or set NAV_EMAIL_PASSWORD."
        )
    client = None
    try:
        client = imaplib.IMAP4_SSL(
            host,
            int(imap.get("port", 993)),
            ssl_context=ssl.create_default_context(),
            timeout=30,
        )
        client.login(str(imap["user"]), password)
        _send_required_imap_id(client, host)
        status, _ = client.select(str(imap.get("mailbox") or "INBOX"), readonly=True)
        if status != "OK":
            raise MailError("Could not open the configured mailbox read-only")
        return client
    except MailError:
        _logout_safely(client)
        raise
    except (imaplib.IMAP4.error, OSError, ssl.SSLError) as exc:
        _logout_safely(client)
        raise MailError("Could not connect to the configured IMAP mailbox") from exc


def _limits(config: dict[str, Any]) -> tuple[int, int, int, int]:
    imap = config.get("imap") or {}
    return (
        int(imap.get("lookback_days", 180)),
        int(imap.get("max_messages", 2000)),
        int(imap.get("max_message_bytes", 25 * 1024 * 1024)),
        int(imap.get("max_total_bytes", 100 * 1024 * 1024)),
    )


def _message_size(client: imaplib.IMAP4_SSL, uid: bytes) -> int:
    status, size_parts = client.uid("fetch", uid, "(RFC822.SIZE)")
    if status != "OK":
        raise MailError("Could not read the size of a mailbox message")
    metadata = b" ".join(part for part in size_parts if isinstance(part, bytes))
    match = re.search(rb"RFC822\.SIZE\s+(\d+)", metadata)
    if not match:
        raise MailError("The IMAP server did not report a mailbox message size")
    return int(match.group(1))


def _message_payload(client: imaplib.IMAP4_SSL, uid: bytes) -> Message:
    status, parts = client.uid("fetch", uid, "(BODY.PEEK[])")
    if status != "OK":
        raise MailError("Could not fetch a mailbox message")
    payload = next(
        (
            item[1]
            for item in parts
            if isinstance(item, tuple) and isinstance(item[1], bytes)
        ),
        None,
    )
    if not payload:
        raise MailError("A mailbox message had no readable payload")
    return BytesParser(policy=default).parsebytes(payload)


def _close(client: imaplib.IMAP4_SSL) -> None:
    try:
        client.close()
    except (imaplib.IMAP4.error, OSError):
        pass
    try:
        client.logout()
    except (imaplib.IMAP4.error, OSError):
        pass


def fetch_candidate_messages(
    config: dict[str, Any],
) -> tuple[list[Message], dict[str, Any]]:
    lookback, max_messages, max_message_bytes, max_total_bytes = _limits(config)
    since = imap_date(dt.date.today() - dt.timedelta(days=lookback))
    client = connect(config)
    messages: list[Message] = []
    total_bytes = 0
    skipped_oversize = 0
    truncated = False
    found = 0
    try:
        status, data = client.uid("search", None, f"(SINCE {since})")
        if status != "OK":
            raise MailError("IMAP search failed while discovering NAV senders")
        all_uids = (data[0] or b"").split()
        found = len(all_uids)
        if len(all_uids) > max_messages:
            truncated = True
        for uid in reversed(all_uids[-max_messages:]):
            size = _message_size(client, uid)
            if size > max_message_bytes:
                skipped_oversize += 1
                continue
            if total_bytes + size > max_total_bytes:
                truncated = True
                break
            messages.append(_message_payload(client, uid))
            total_bytes += size
    finally:
        _close(client)
    return messages, {
        "messages_found": found,
        "messages_fetched": len(messages),
        "bytes_fetched": total_bytes,
        "skipped_oversize": skipped_oversize,
        "truncated": truncated,
    }


def fetch_authorized_messages(config: dict[str, Any]) -> dict[str, list[Message]]:
    senders = sorted(
        {str(route["sender"]).strip().lower() for route in active_routes(config)}
    )
    if not senders:
        return {}
    lookback, max_messages, max_message_bytes, max_total_bytes = _limits(config)
    since = imap_date(dt.date.today() - dt.timedelta(days=lookback))
    client = connect(config)
    output: dict[str, list[Message]] = {sender: [] for sender in senders}
    message_count = 0
    total_bytes = 0
    try:
        for sender in senders:
            query = f'(SINCE {since} FROM "{sender}")'
            status, data = client.uid("search", None, query)
            if status != "OK":
                raise MailError("IMAP search failed for an authorized sender")
            uids = (data[0] or b"").split()
            if message_count + len(uids) > max_messages:
                raise MailError(
                    "Authorized-message count exceeds imap.max_messages; narrow the lookback window"
                )
            for uid in uids:
                size = _message_size(client, uid)
                if size > max_message_bytes:
                    raise MailError(
                        "An authorized message exceeds imap.max_message_bytes"
                    )
                if total_bytes + size > max_total_bytes:
                    raise MailError(
                        "Authorized messages exceed imap.max_total_bytes; narrow the lookback window"
                    )
                message = _message_payload(client, uid)
                if not exact_from_matches(message, sender):
                    raise MailError(
                        "An IMAP sender-search result did not have the exact authorized From address"
                    )
                output[sender].append(message)
                message_count += 1
                total_bytes += size
    finally:
        _close(client)
    return output
