from __future__ import annotations

import datetime as dt
import hashlib
import imaplib
import re
import ssl
import time
from email.header import decode_header, make_header
from email.message import Message
from email.parser import BytesParser
from email.policy import default
from email.utils import getaddresses
from typing import Any, Callable

from nav_config import active_routes
from runtime_secret import read_password


class MailError(RuntimeError):
    pass


class AuthorizedMessageMap(dict[str, list[Message]]):
    def __init__(self, senders: list[str]):
        super().__init__((sender, []) for sender in senders)
        self.excluded_non_nav_messages = 0
        self.excluded_non_nav_reasons: dict[str, int] = {}
        self.scope_fingerprint: str | None = None
        self.messages_selected = 0
        self.total_bytes = 0
        self.reconnect_count = 0


def non_nav_subject_category(subject: str) -> str | None:
    compact = re.sub(r"\s+", "", subject).casefold()
    explicit_nav_signal = (
        "净值" in compact
        or bool(re.search(r"\bnav\b", subject.casefold()))
        or "unitnav" in compact
        or "unitvalue" in compact
    )
    if any(
        marker in compact
        for marker in (
            "虚拟估算",
            "模拟估算",
            "净值估算",
            "估算净值",
            "estimatednav",
            "navestimate",
        )
    ):
        return "virtual-estimate"
    if not explicit_nav_signal and any(
        marker in compact
        for marker in (
            "交易确认",
            "成交确认",
            "申购确认",
            "赎回确认",
            "tradeconfirmation",
            "transactionconfirmation",
        )
    ):
        return "transaction-confirmation"
    if not explicit_nav_signal and any(
        marker in compact
        for marker in (
            "月报",
            "季报",
            "monthlyreport",
            "quarterlyreport",
        )
    ):
        return "periodic-report"
    return None


ProgressCallback = Callable[[str, int, int], None]


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
    client: imaplib.IMAP4_SSL,
    uids: list[bytes],
    batch_size: int = 200,
    *,
    deadline: float | None = None,
    progress: ProgressCallback | None = None,
) -> dict[bytes, int]:
    sizes: dict[bytes, int] = {}
    for offset in range(0, len(uids), batch_size):
        if deadline is not None and time.monotonic() >= deadline:
            break
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
        if progress:
            progress("sizes", min(offset + len(batch), len(uids)), len(uids))
    return sizes


def _message_headers(
    client: imaplib.IMAP4_SSL,
    uids: list[bytes],
    batch_size: int = 200,
    *,
    deadline: float | None = None,
    progress: ProgressCallback | None = None,
) -> dict[bytes, Message]:
    headers: dict[bytes, Message] = {}
    for offset in range(0, len(uids), batch_size):
        if deadline is not None and time.monotonic() >= deadline:
            break
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
        if progress:
            progress("headers", min(offset + len(batch), len(uids)), len(uids))
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


def _uidvalidity(client: imaplib.IMAP4_SSL) -> str:
    try:
        status, values = client.response("UIDVALIDITY")
    except (AttributeError, imaplib.IMAP4.error, OSError):
        return "unavailable"
    if status != "UIDVALIDITY":
        return "unavailable"
    value = next(
        (
            item.decode("ascii", errors="replace")
            for item in values or []
            if isinstance(item, bytes) and item
        ),
        None,
    )
    return value or "unavailable"


class _ReconnectableIMAP:
    def __init__(
        self,
        config: dict[str, Any],
        *,
        progress: ProgressCallback | None = None,
        max_reconnects: int = 2,
    ):
        self.config = config
        self.progress = progress
        self.max_reconnects = max_reconnects
        self.reconnect_count = 0
        self.client = connect(config)
        self.uidvalidity = _uidvalidity(self.client)

    def call(
        self,
        stage: str,
        operation: Callable[[imaplib.IMAP4_SSL], Any],
    ) -> Any:
        while True:
            try:
                return operation(self.client)
            except (imaplib.IMAP4.abort, OSError, ssl.SSLError):
                if self.reconnect_count >= self.max_reconnects:
                    raise
                self.reconnect_count += 1
                _close(self.client)
                self.client = connect(self.config)
                current_uidvalidity = _uidvalidity(self.client)
                if current_uidvalidity != self.uidvalidity:
                    _close(self.client)
                    raise _mail_read_error(
                        stage,
                        "IMAP 重新连接后 UIDVALIDITY 发生变化",
                        "停止复用当前 UID；稍后重新开始本次只读扫描",
                    )
                if self.progress:
                    self.progress(
                        "reconnect",
                        self.reconnect_count,
                        self.max_reconnects,
                    )

    def close(self) -> None:
        _close(self.client)


def _max_header_messages(config: dict[str, Any]) -> int:
    imap = config.get("imap") or {}
    return int(imap.get("max_header_messages", 20000))


def _search_text(value: str) -> str:
    if any(character in value for character in ("\x00", "\r", "\n")):
        raise MailError("IMAP 搜索条件包含非法控制字符")
    return value.replace("\\", "\\\\").replace('"', '\\"')


def fetch_candidate_headers(
    config: dict[str, Any], limit: int = 25
) -> tuple[list[Message], dict[str, Any]]:
    """Read only From/Subject headers for a bounded recent mailbox sample."""

    lookback, _, _, _ = _limits(config)
    max_header_messages = _max_header_messages(config)
    if isinstance(limit, bool) or not 1 <= int(limit) <= max_header_messages:
        raise MailError(f"候选邮件头数量必须在 1 到 {max_header_messages} 之间")
    since = imap_date(dt.date.today() - dt.timedelta(days=lookback))
    session = _ReconnectableIMAP(config)
    found = 0
    stage = "搜索候选邮件头"
    try:
        status, data = session.call(
            stage,
            lambda client: client.uid("search", None, f"(SINCE {since})"),
        )
        if status != "OK":
            raise _mail_read_error(
                "搜索候选邮件头",
                "IMAP 服务器拒绝了候选搜索请求",
                "稍后重试；若持续失败，请检查服务商 IMAP 兼容性",
            )
        all_uids = (data[0] or b"").split()
        found = len(all_uids)
        selected_uids = all_uids[-int(limit) :]
        stage = "读取最小候选邮件头"
        headers = session.call(
            stage,
            lambda client: _message_headers(client, selected_uids),
        )
        return [headers[uid] for uid in reversed(selected_uids)], {
            "messages_found": found,
            "headers_fetched": len(selected_uids),
            "messages_fetched": 0,
            "bytes_fetched": 0,
            "truncated": found > len(selected_uids),
            "imap_reconnects": session.reconnect_count,
        }
    except MailError:
        raise
    except (imaplib.IMAP4.error, OSError, ssl.SSLError) as exc:
        raise _mail_read_error(
            stage,
            "IMAP 会话意外断开，本次没有生成可用邮件头报告",
            "检查网络和邮箱服务状态后重试",
        ) from exc
    finally:
        session.close()


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
    *,
    sender: str | None = None,
    subject_contains: str | None = None,
    before_uid: int | None = None,
    batch_messages: int | None = None,
    deadline: float | None = None,
    progress: ProgressCallback | None = None,
) -> tuple[list[Message], dict[str, Any]]:
    lookback, max_messages, max_message_bytes, max_total_bytes = _limits(config)
    since = imap_date(dt.date.today() - dt.timedelta(days=lookback))
    selected_sender = str(sender or "").strip().casefold()
    selected_subject = str(subject_contains or "").strip().casefold()
    if selected_subject and not selected_sender:
        raise MailError("按主题选择候选邮件时必须同时指定精确发件人")
    session = _ReconnectableIMAP(config, progress=progress)
    messages: list[Message] = []
    total_bytes = 0
    skipped_oversize = 0
    truncated = False
    found = 0
    stage = "搜索候选邮件"
    try:
        status, data = session.call(
            stage,
            lambda client: client.uid("search", None, f"(SINCE {since})"),
        )
        if status != "OK":
            raise MailError("IMAP search failed while discovering NAV senders")
        all_uids = (data[0] or b"").split()
        found = len(all_uids)
        if selected_sender:
            stage = "按发件人收窄候选邮件"
            sender_query = f'(SINCE {since} FROM "{_search_text(selected_sender)}"'
            if selected_subject:
                sender_query += f' SUBJECT "{_search_text(selected_subject)}"'
            sender_query += ")"
            status, data = session.call(
                stage,
                lambda client: client.uid("search", None, sender_query),
            )
            if status != "OK":
                raise _mail_read_error(
                    "按发件人收窄候选邮件",
                    "IMAP 服务器拒绝了发件人范围搜索",
                    "稍后重试；若持续失败，请检查服务商 IMAP 搜索兼容性",
                )
            sender_uids = (data[0] or b"").split()
            server_scope_matches = len(sender_uids)
            if progress:
                progress("search", len(sender_uids), len(sender_uids))
            if before_uid is not None:
                sender_uids = [
                    uid for uid in sender_uids if uid.isdigit() and int(uid) < before_uid
                ]
            max_header_messages = _max_header_messages(config)
            if len(sender_uids) > max_header_messages:
                raise _mail_read_error(
                    "候选邮件头扫描边界",
                    "所选发件人在回看期内的邮件数量超过候选邮件头扫描上限",
                    "缩短回看窗口，或使用更精确的 subject_contains",
                )
            header_uids = sender_uids
        else:
            sender_uids = all_uids
            server_scope_matches = len(sender_uids)
            if len(all_uids) > max_messages:
                truncated = True
            header_uids = all_uids[-max_messages:]
        stage = "读取最小候选邮件头"
        headers = session.call(
            stage,
            lambda client: _message_headers(
                client, header_uids, progress=progress
            ),
        )
        if selected_sender:
            selected_uids = []
            for uid in header_uids:
                header = headers[uid]
                if not exact_from_matches(header, selected_sender):
                    continue
                subject = decoded(header.get("Subject")).casefold()
                if selected_subject and selected_subject not in subject:
                    continue
                selected_uids.append(uid)
            if len(selected_uids) > max_messages:
                raise _mail_read_error(
                    "候选邮件选择后计数",
                    "符合发件人和主题范围的候选邮件超过配置上限",
                    "缩短回看窗口或使用更精确的 subject_contains",
                )
        else:
            selected_uids = header_uids
        matching_messages = len(selected_uids)
        batch_limit = max_messages
        if batch_messages is not None:
            if (
                isinstance(batch_messages, bool)
                or not 1 <= int(batch_messages) <= max_messages
            ):
                raise MailError(
                    f"单次候选解析邮件数必须在 1 到 {max_messages} 之间"
                )
            batch_limit = int(batch_messages)
        older_remaining = len(selected_uids) > batch_limit
        selected_uids = selected_uids[-batch_limit:]
        stage = "批量读取候选邮件大小"
        sizes = session.call(
            stage,
            lambda client: _message_sizes(
                client, selected_uids, progress=progress
            ),
        )
        handled_uids: list[bytes] = []
        timed_out = False
        for uid in reversed(selected_uids):
            if deadline is not None and time.monotonic() >= deadline:
                timed_out = True
                break
            size = sizes[uid]
            if size > max_message_bytes:
                skipped_oversize += 1
                handled_uids.append(uid)
                continue
            if total_bytes + size > max_total_bytes:
                truncated = True
                break
            stage = "读取候选邮件内容"
            message = session.call(
                stage,
                lambda client: _message_payload(client, uid),
            )
            header_sender = single_from_address(headers[uid])
            if single_from_address(message) != header_sender:
                raise _mail_read_error(
                    "复核完整候选邮件",
                    "完整邮件的 From 与已读取的最小邮件头不一致",
                    "停止本次读取并检查邮箱服务商返回内容",
                )
            messages.append(message)
            setattr(message, "_nav_source_uid", int(uid))
            handled_uids.append(uid)
            total_bytes += size
            if progress:
                progress("messages", len(handled_uids), len(selected_uids))
    except MailError:
        raise
    except (imaplib.IMAP4.error, OSError, ssl.SSLError) as exc:
        raise MailError(
            f"IMAP 会话在{stage}时意外断开；本次没有生成可用扫描结果，请检查网络后重试"
        ) from exc
    finally:
        session.close()
    range_complete = not (older_remaining or timed_out or truncated)
    resume_before_uid = None
    if not range_complete:
        if handled_uids:
            resume_before_uid = min(int(uid) for uid in handled_uids)
        elif selected_uids:
            resume_before_uid = max(int(uid) for uid in selected_uids) + 1
        elif before_uid is not None:
            resume_before_uid = before_uid
    return messages, {
        "messages_found": found,
        "server_since_matches": found,
        "server_sender_matches": server_scope_matches if selected_sender else None,
        "server_subject_filter_applied": bool(selected_subject),
        "headers_fetched": len(header_uids),
        "messages_selected": len(selected_uids),
        "matching_messages_in_range": matching_messages,
        "messages_fetched": len(messages),
        "bytes_fetched": total_bytes,
        "skipped_oversize": skipped_oversize,
        "truncated": truncated,
        "timed_out": timed_out,
        "range_complete": range_complete,
        "resume_before_uid": resume_before_uid,
        "imap_reconnects": session.reconnect_count,
        "selection_applied": bool(selected_sender),
        "selection": {
            "mode": (
                "sender-subject"
                if selected_subject
                else "sender"
                if selected_sender
                else "all"
            ),
            "sender": selected_sender or None,
            "subject_contains": str(subject_contains or "").strip() or None,
        },
    }


def fetch_authorized_messages(
    config: dict[str, Any], *, load_bodies: bool = True
) -> AuthorizedMessageMap:
    routes = active_routes(config)
    senders = sorted({str(route["sender"]).strip().lower() for route in routes})
    if not senders:
        return AuthorizedMessageMap([])
    sender_subjects: dict[str, set[str | None]] = {sender: set() for sender in senders}
    sender_routes: dict[str, list[dict[str, Any]]] = {sender: [] for sender in senders}
    for route in routes:
        sender = str(route["sender"]).strip().lower()
        sender_routes[sender].append(route)
        subject_filter = str(route.get("subject_contains") or "").strip()
        sender_subjects[sender].add(subject_filter.casefold() or None)
    lookback, max_messages, max_message_bytes, max_total_bytes = _limits(config)
    since = imap_date(dt.date.today() - dt.timedelta(days=lookback))
    session = _ReconnectableIMAP(config)
    output = AuthorizedMessageMap(senders)
    message_count = 0
    total_bytes = 0
    scope_items: list[str] = [f"uidvalidity:{session.uidvalidity}"]
    stage = "搜索已授权发件人邮件"
    try:
        for sender in senders:
            subject_filters = sender_subjects[sender]
            search_subjects = [None] if None in subject_filters else sorted(subject_filters)
            scoped_uids: set[bytes] = set()
            for subject_filter in search_subjects:
                query = f'(SINCE {since} FROM "{_search_text(sender)}"'
                if subject_filter:
                    query += f' SUBJECT "{_search_text(subject_filter)}"'
                query += ")"
                status, data = session.call(
                    stage,
                    lambda client: client.uid("search", None, query),
                )
                if status != "OK":
                    raise _mail_read_error(
                        "搜索授权发件人邮件",
                        "IMAP 服务器拒绝了搜索请求",
                        "稍后重试；若持续失败，请检查服务商 IMAP 兼容性",
                    )
                scoped_uids.update((data[0] or b"").split())
            sender_uids = sorted(
                scoped_uids,
                key=lambda uid: int(uid) if uid.isdigit() else 0,
            )
            max_header_messages = _max_header_messages(config)
            if len(sender_uids) > max_header_messages:
                raise _mail_read_error(
                    "邮件头扫描边界",
                    "授权发件人在回看期内的邮件数量超过邮件头扫描上限",
                    "缩短回看窗口，或在核实邮箱规模后调整 imap.max_header_messages",
                )
            stage = "读取最小邮件头并按主题预筛选"
            headers = session.call(
                stage,
                lambda client: _message_headers(client, sender_uids),
            )
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
                positively_matched = [
                    route
                    for route in sender_routes[sender]
                    if (
                        not str(route.get("subject_contains") or "").strip()
                        or str(route.get("subject_contains")).strip().casefold()
                        in subject
                    )
                ]
                accepted = False
                exclusion_reasons: set[str] = set()
                built_in_category = non_nav_subject_category(subject)
                for route in positively_matched:
                    if any(
                        str(excluded).strip().casefold() in subject
                        for excluded in route.get("subject_excludes") or []
                    ):
                        exclusion_reasons.add("configured-subject-exclude")
                        continue
                    if built_in_category:
                        exclusion_reasons.add(built_in_category)
                        continue
                    accepted = True
                    break
                if accepted:
                    uids.append(uid)
                elif positively_matched and exclusion_reasons:
                    output.excluded_non_nav_messages += 1
                    for reason in exclusion_reasons:
                        output.excluded_non_nav_reasons[reason] = (
                            output.excluded_non_nav_reasons.get(reason, 0) + 1
                        )
                if positively_matched:
                    scope_items.append(
                        "\x1f".join(
                            (
                                sender,
                                uid.decode("ascii", errors="replace"),
                                decoded(header.get("From")),
                                decoded(header.get("Subject")),
                                "accepted"
                                if accepted
                                else ",".join(sorted(exclusion_reasons))
                                or "not-accepted",
                            )
                        )
                    )
            if message_count + len(uids) > max_messages:
                raise _mail_read_error(
                    "主题预筛选后计数",
                    "符合路由的授权邮件数量超过配置上限",
                    "缩短回看窗口或在核实资源容量后调整 imap.max_messages",
                )
            stage = "批量读取已授权邮件大小"
            sizes = session.call(
                stage,
                lambda client: _message_sizes(client, uids),
            )
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
                scope_items.append(
                    "\x1f".join(
                        (
                            sender,
                            uid.decode("ascii", errors="replace"),
                            str(size),
                            "accepted-size",
                        )
                    )
                )
                message_count += 1
                total_bytes += size
                if not load_bodies:
                    continue
                stage = "读取已授权邮件内容"
                message = session.call(
                    stage,
                    lambda client: _message_payload(client, uid),
                )
                if not exact_from_matches(message, sender):
                    raise _mail_read_error(
                        "复核完整邮件",
                        "完整邮件的 From 与已校验邮件头或授权发件人不完全一致",
                        "停止本次读取并检查邮箱服务商返回内容",
                    )
                setattr(message, "_nav_source_uid", int(uid))
                output[sender].append(message)
    except MailError:
        raise
    except (imaplib.IMAP4.error, OSError, ssl.SSLError) as exc:
        raise _mail_read_error(
            stage,
            "IMAP 会话意外断开，本次没有生成可用扫描结果",
            "检查网络和邮箱服务状态后重试",
        ) from exc
    finally:
        session.close()
    output.scope_fingerprint = hashlib.sha256(
        "\x1e".join(sorted(scope_items)).encode("utf-8", errors="replace")
    ).hexdigest()
    output.messages_selected = message_count
    output.total_bytes = total_bytes
    output.reconnect_count = session.reconnect_count
    return output
