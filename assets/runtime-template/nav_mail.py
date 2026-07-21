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
            "没有可用的本地 IMAP 授权码；请先运行 secret launch 并由 AI 用 secret status 复查"
        )
    client = None
    stage = "建立 TLS 连接"
    try:
        client = imaplib.IMAP4_SSL(
            host,
            int(imap.get("port", 993)),
            ssl_context=ssl.create_default_context(),
            timeout=30,
        )
        stage = "登录"
        client.login(str(imap["user"]), password)
        stage = "完成服务商握手"
        _send_required_imap_id(client, host)
        stage = "只读打开邮箱文件夹"
        status, _ = client.select(str(imap.get("mailbox") or "INBOX"), readonly=True)
        if status != "OK":
            raise MailError("无法以只读方式打开配置的邮箱文件夹")
        return client
    except MailError:
        _logout_safely(client)
        raise
    except imaplib.IMAP4.error as exc:
        _logout_safely(client)
        hint = (
            "请检查授权码、IMAP 开关和服务商的第三方客户端策略"
            if stage == "登录"
            else "请检查服务器地址、网络和服务商状态"
        )
        raise MailError(f"IMAP 在{stage}阶段被服务器拒绝；{hint}") from exc
    except (OSError, ssl.SSLError) as exc:
        _logout_safely(client)
        raise MailError(
            f"IMAP 在{stage}阶段发生网络或 TLS 错误；请检查网络后重试"
        ) from exc


def _limits(config: dict[str, Any]) -> tuple[int, int, int, int]:
    imap = config.get("imap") or {}
    return (
        int(imap.get("lookback_days", 180)),
        int(imap.get("max_messages", 2000)),
        int(imap.get("max_message_bytes", 25 * 1024 * 1024)),
        int(imap.get("max_total_bytes", 100 * 1024 * 1024)),
    )


def _message_sizes(
    client: imaplib.IMAP4_SSL, uids: list[bytes], batch_size: int = 200
) -> dict[bytes, int]:
    sizes: dict[bytes, int] = {}
    for offset in range(0, len(uids), batch_size):
        batch = uids[offset : offset + batch_size]
        status, size_parts = client.uid("fetch", b",".join(batch), "(UID RFC822.SIZE)")
        if status != "OK":
            raise MailError("无法批量读取邮箱邮件大小")
        fallback_sizes: list[int] = []
        for part in size_parts:
            metadata = (
                part
                if isinstance(part, bytes)
                else part[0]
                if isinstance(part, tuple) and part and isinstance(part[0], bytes)
                else b""
            )
            uid_match = re.search(rb"\bUID\s+(\d+)\b", metadata, re.IGNORECASE)
            size_match = re.search(
                rb"\bRFC822\.SIZE\s+(\d+)\b", metadata, re.IGNORECASE
            )
            if not size_match:
                continue
            size = int(size_match.group(1))
            if uid_match:
                sizes[uid_match.group(1)] = size
            else:
                fallback_sizes.append(size)
        if len(batch) == 1 and batch[0] not in sizes and len(fallback_sizes) == 1:
            sizes[batch[0]] = fallback_sizes[0]
        missing = [uid for uid in batch if uid not in sizes]
        if missing:
            raise MailError("IMAP 服务器没有返回完整的邮件大小信息")
    return sizes


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
    stage = "搜索候选邮件"
    try:
        status, data = client.uid("search", None, f"(SINCE {since})")
        if status != "OK":
            raise MailError("IMAP search failed while discovering NAV senders")
        all_uids = (data[0] or b"").split()
        found = len(all_uids)
        if len(all_uids) > max_messages:
            truncated = True
        selected_uids = all_uids[-max_messages:]
        stage = "批量读取候选邮件大小"
        sizes = _message_sizes(client, selected_uids)
        for uid in reversed(selected_uids):
            size = sizes[uid]
            if size > max_message_bytes:
                skipped_oversize += 1
                continue
            if total_bytes + size > max_total_bytes:
                truncated = True
                break
            stage = "读取候选邮件内容"
            messages.append(_message_payload(client, uid))
            total_bytes += size
    except MailError:
        raise
    except (imaplib.IMAP4.error, OSError, ssl.SSLError) as exc:
        raise MailError(
            f"IMAP 会话在{stage}时意外断开；本次没有生成可用扫描结果，请检查网络后重试"
        ) from exc
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
    stage = "搜索已授权发件人邮件"
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
            stage = "批量读取已授权邮件大小"
            sizes = _message_sizes(client, uids)
            for uid in uids:
                size = sizes[uid]
                if size > max_message_bytes:
                    raise MailError(
                        "An authorized message exceeds imap.max_message_bytes"
                    )
                if total_bytes + size > max_total_bytes:
                    raise MailError(
                        "Authorized messages exceed imap.max_total_bytes; narrow the lookback window"
                    )
                stage = "读取已授权邮件内容"
                message = _message_payload(client, uid)
                if not exact_from_matches(message, sender):
                    raise MailError(
                        "An IMAP sender-search result did not have the exact authorized From address"
                    )
                output[sender].append(message)
                message_count += 1
                total_bytes += size
    except MailError:
        raise
    except (imaplib.IMAP4.error, OSError, ssl.SSLError) as exc:
        raise MailError(
            f"IMAP 会话在{stage}时意外断开；本次没有生成可用扫描结果，请检查网络后重试"
        ) from exc
    finally:
        _close(client)
    return output
