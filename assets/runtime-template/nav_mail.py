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


def _mail_read_error(stage: str, reason: str, suggestion: str) -> MailError:
    return MailError(
        f"邮件读取失败；阶段：{stage}；原因：{reason}；处理建议：{suggestion}"
    )


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
            raise _mail_read_error(
                "读取邮件大小",
                "IMAP 服务器拒绝了批量大小查询",
                "稍后重试；若持续失败，请检查服务商 IMAP 兼容性",
            )
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
            raise _mail_read_error(
                "读取邮件大小",
                "IMAP 服务器返回的大小信息不完整",
                "稍后重试；若持续失败，请检查服务商 IMAP 兼容性",
            )
    return sizes


def _message_headers(
    client: imaplib.IMAP4_SSL, uids: list[bytes], batch_size: int = 200
) -> dict[bytes, Message]:
    headers: dict[bytes, Message] = {}
    for offset in range(0, len(uids), batch_size):
        batch = uids[offset : offset + batch_size]
        status, parts = client.uid(
            "fetch",
            b",".join(batch),
            "(UID BODY.PEEK[HEADER.FIELDS (FROM SUBJECT)])",
        )
        if status != "OK":
            raise _mail_read_error(
                "读取最小邮件头",
                "IMAP 服务器拒绝了批量邮件头查询",
                "稍后重试；若持续失败，请检查服务商 IMAP 兼容性",
            )
        for part in parts:
            if not isinstance(part, tuple) or len(part) < 2:
                continue
            metadata, payload = part[0], part[1]
            if not isinstance(metadata, bytes) or not isinstance(payload, bytes):
                continue
            uid_match = re.search(rb"\bUID\s+(\d+)\b", metadata, re.IGNORECASE)
            if not uid_match:
                continue
            headers[uid_match.group(1)] = BytesParser(policy=default).parsebytes(
                payload
            )
        missing = [uid for uid in batch if uid not in headers]
        if missing:
            raise _mail_read_error(
                "读取最小邮件头",
                "IMAP 服务器返回的邮件头信息不完整",
                "稍后重试；若持续失败，请检查服务商 IMAP 兼容性",
            )
    return headers


def _message_payload(client: imaplib.IMAP4_SSL, uid: bytes) -> Message:
    status, parts = client.uid("fetch", uid, "(BODY.PEEK[])")
    if status != "OK":
        raise _mail_read_error(
            "读取完整邮件",
            "IMAP 服务器拒绝了邮件内容查询",
            "稍后重试；若持续失败，请检查服务商 IMAP 兼容性",
        )
    payload = next(
        (
            item[1]
            for item in parts
            if isinstance(item, tuple) and isinstance(item[1], bytes)
        ),
        None,
    )
    if not payload:
        raise _mail_read_error(
            "读取完整邮件",
            "IMAP 服务器未返回可解析的邮件内容",
            "稍后重试；若持续失败，请检查服务商 IMAP 兼容性",
        )
    return BytesParser(policy=default).parsebytes(payload)


def _max_header_messages(config: dict[str, Any]) -> int:
    imap = config.get("imap") or {}
    return int(imap.get("max_header_messages", 20000))


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
    routes = active_routes(config)
    senders = sorted({str(route["sender"]).strip().lower() for route in routes})
    if not senders:
        return {}
    sender_subjects: dict[str, set[str | None]] = {sender: set() for sender in senders}
    for route in routes:
        sender = str(route["sender"]).strip().lower()
        subject_filter = str(route.get("subject_contains") or "").strip()
        sender_subjects[sender].add(subject_filter.casefold() or None)
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
                raise _mail_read_error(
                    "搜索授权发件人邮件",
                    "IMAP 服务器拒绝了搜索请求",
                    "稍后重试；若持续失败，请检查服务商 IMAP 兼容性",
                )
            sender_uids = (data[0] or b"").split()
            max_header_messages = _max_header_messages(config)
            if len(sender_uids) > max_header_messages:
                raise _mail_read_error(
                    "邮件头扫描边界",
                    "授权发件人在回看期内的邮件数量超过邮件头扫描上限",
                    "缩短回看窗口，或在核实邮箱规模后调整 imap.max_header_messages",
                )
            stage = "读取最小邮件头并按主题预筛选"
            headers = _message_headers(client, sender_uids)
            subject_filters = sender_subjects[sender]
            uids: list[bytes] = []
            for uid in sender_uids:
                header = headers[uid]
                if not exact_from_matches(header, sender):
                    raise _mail_read_error(
                        "校验最小邮件头",
                        "搜索结果的 From 与授权发件人不完全一致",
                        "停止本次读取并检查邮箱服务商的搜索行为",
                    )
                subject = decoded(header.get("Subject")).casefold()
                if None in subject_filters or any(
                    item is not None and item in subject for item in subject_filters
                ):
                    uids.append(uid)
            if message_count + len(uids) > max_messages:
                raise _mail_read_error(
                    "主题预筛选后计数",
                    "符合路由的授权邮件数量超过配置上限",
                    "缩短回看窗口或在核实资源容量后调整 imap.max_messages",
                )
            stage = "批量读取已授权邮件大小"
            sizes = _message_sizes(client, uids)
            for uid in uids:
                size = sizes[uid]
                if size > max_message_bytes:
                    raise _mail_read_error(
                        "校验单封邮件大小",
                        "符合路由的授权邮件超过单封大小上限",
                        "核实邮件来源后调整 imap.max_message_bytes 或改用更小的附件",
                    )
                if total_bytes + size > max_total_bytes:
                    raise _mail_read_error(
                        "校验邮件总大小",
                        "符合路由的授权邮件累计大小超过配置上限",
                        "缩短回看窗口或在核实资源容量后调整 imap.max_total_bytes",
                    )
                stage = "读取已授权邮件内容"
                message = _message_payload(client, uid)
                if not exact_from_matches(message, sender):
                    raise _mail_read_error(
                        "复核完整邮件",
                        "完整邮件的 From 与已校验邮件头或授权发件人不完全一致",
                        "停止本次读取并检查邮箱服务商返回内容",
                    )
                output[sender].append(message)
                message_count += 1
                total_bytes += size
    except MailError:
        raise
    except (imaplib.IMAP4.error, OSError, ssl.SSLError) as exc:
        raise _mail_read_error(
            stage,
            "IMAP 会话意外断开，本次没有生成可用扫描结果",
            "检查网络和邮箱服务状态后重试",
        ) from exc
    finally:
        _close(client)
    return output
