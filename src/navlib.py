# -*- coding: utf-8 -*-
"""Shared library: IMAP access, HTML/text helpers, date parsing, and per-format
NAV parsers. Each parser takes (subject, body_text) and returns a dict:
    {code, date(datetime.date), unit(float), cum(float|None), source(str), extra(dict)}
or None if it cannot parse.
"""
import imaplib, email, re, json, os, datetime
from email.header import decode_header, make_header

# ----------------------------------------------------------------------------- config
HERE = os.path.dirname(os.path.abspath(__file__))          # nav_tool/
PROJECT_DIR = os.path.dirname(HERE)                         # folder holding the workbook
CFG = json.load(open(os.path.join(HERE, "config.json"), encoding="utf-8"))

def _resolve_master():
    """Locate the master workbook portably so the tool moves between machines:
    1) configured absolute path if it exists; 2) configured name under PROJECT_DIR;
    3) the single real .xlsx in PROJECT_DIR (ignoring temp/preview/backup files)."""
    import glob
    mp = (CFG.get("master_path") or "").strip()
    if mp and os.path.isabs(mp) and os.path.exists(mp):
        return mp
    if mp:
        cand = os.path.join(PROJECT_DIR, os.path.basename(mp))
        if os.path.exists(cand):
            return cand
    xs = [f for f in glob.glob(os.path.join(PROJECT_DIR, "*.xlsx"))
          if not os.path.basename(f).startswith("~$")
          and "预览" not in f and "_自动更新" not in f and "_TEST" not in f
          and "备份" not in os.path.basename(f) and "backup" not in os.path.basename(f).lower()]
    if len(xs) == 1:
        return xs[0]
    # 多个 .xlsx 时, 优先文件名含「净值」的(公司表名为「各私募净值…」)
    nav = [f for f in xs if "净值" in os.path.basename(f)]
    if len(nav) == 1:
        return nav[0]
    if mp:
        return mp  # let it fail later with a clear path
    raise FileNotFoundError("找不到净值表，请在 nav_tool/config.json 的 master_path 填写文件名或绝对路径")

# Resolve portable paths once, overwriting config values so existing references keep working.
CFG["master_path"] = _resolve_master()
CFG["registry_path"] = os.path.join(HERE, "registry.json")
CFG["workdir"] = HERE

def load_registry():
    return json.load(open(CFG["registry_path"], encoding="utf-8"))

def all_scope():
    """All sheets the tool manages (phase-1 body formats + phase-2 attachments)."""
    return list(CFG["scope_sheets"]) + list(CFG.get("scope_sheets_p2", []))

def base_code(code):
    """Strip share-class suffix like '(B级)' / '(A级)' and whitespace/newlines."""
    if not code:
        return code
    return re.split(r"[（(]", code.strip())[0].strip()

def code_to_sheet():
    """Map base product code -> sheet name (restricted to scope_sheets)."""
    reg = load_registry()
    m = {}
    for sheet in CFG["scope_sheets"]:
        for c in reg[sheet]["codes"]:
            m[base_code(c)] = sheet
    return m

# ----------------------------------------------------------------------------- header
def dh(s):
    if not s:
        return ""
    try:
        return str(make_header(decode_header(s)))
    except Exception:
        return str(s)

# ----------------------------------------------------------------------------- html
def strip_html(h):
    h = re.sub(r"(?is)<(script|style).*?>.*?</\1>", "", h)
    h = re.sub(r"(?is)<br\s*/?>", "\n", h)
    h = re.sub(r"(?is)</(p|div|tr|table|h\d)>", "\n", h)
    h = re.sub(r"(?is)</td>", " | ", h)
    h = re.sub(r"(?is)<[^>]+>", "", h)
    h = h.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    h = re.sub(r"[ \t]+", " ", h)
    h = re.sub(r"\n\s*\n+", "\n", h)
    return h.strip()

def body_text(msg):
    out = []
    for p in (msg.walk() if msg.is_multipart() else [msg]):
        if "attachment" in str(p.get("Content-Disposition") or "").lower():
            continue
        if p.get_content_type() in ("text/plain", "text/html"):
            pl = p.get_payload(decode=True)
            if not pl:
                continue
            txt = pl.decode(p.get_content_charset() or "utf-8", errors="replace")
            if p.get_content_type() == "text/html":
                txt = strip_html(txt)
            out.append(txt)
    return "\n--\n".join(out)

# ----------------------------------------------------------------------------- numbers/dates
def num(s):
    if s is None:
        return None
    s = str(s).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None

def parse_date(s):
    if not s:
        return None
    s = str(s).strip()
    m = re.search(r"(\d{4})\D(\d{1,2})\D(\d{1,2})", s)  # 2026-06-11 / 2026年06月11日
    if m:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r"(\d{4})(\d{2})(\d{2})", s)  # 20260611
    if m:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None

# ----------------------------------------------------------------------------- subject parse (for fast indexing)
def parse_subject(fmt, subj):
    """Return (base_code, date, subtype, unit_hint) from subject without body."""
    code = date = subtype = unit = None
    if fmt == "gtht":
        # 资产净值公告_DEMO04_产品名_..._2026-06-05
        m = re.search(r"资产净值公告_([A-Z0-9]+)_", subj)
        if m: code = m.group(1)
        date = parse_date(subj)
    elif fmt == "citics":
        subtype = "virtual" if ("虚拟" in subj or "估算" in subj) else "real"
        m = re.search(r"】\s*([A-Z0-9]+(?:[（(][^）)]*[）)])?)_", subj)
        if m: code = base_code(m.group(1))
        # date = the YYYY-MM-DD in subject
        date = parse_date(subj)
    elif fmt == "htsc_glrfw":
        m = re.search(r"产品([A-Z0-9]+)", subj)
        if m: code = m.group(1)
        date = parse_date(subj)
    elif fmt == "htsc_incos":
        m = re.search(r"】\s*([A-Z0-9]+)_", subj)
        if m: code = m.group(1)
        date = parse_date(subj)
        mu = re.search(r"单位净值[:：]\s*([\d.]+)", subj)
        if mu: unit = num(mu.group(1))
    elif fmt == "xyzq":
        m = re.match(r"\s*([A-Z0-9]+)", subj)
        if m: code = m.group(1)
        date = parse_date(subj)
    return base_code(code) if code else None, date, subtype, unit

# ----------------------------------------------------------------------------- body parsers
_NUM = r"([\d,]+\.?\d*)"

def parse_body(fmt, subj, body, subj_code=None):
    """Return dict(code,date,unit,cum,virtual,source) or None."""
    if fmt == "gtht":
        m = re.search(r"([A-Z0-9]{5,8})\s*\|\s*[^|]+\|\s*(\d{4}-\d{2}-\d{2})\s*\|\s*"+_NUM+r"\s*\|\s*"+_NUM, body)
        if m:
            return dict(code=base_code(m.group(1)), date=parse_date(m.group(2)),
                        unit=num(m.group(3)), cum=num(m.group(4)), virtual=None, source="gtht")
    elif fmt == "citics":
        if "虚拟" in subj or "估算" in subj:
            # name | acct | date | TA计提 | shares | virtual | actual | actual_cum
            # cols: date | TA计提 | 持仓份额 | 虚拟净值(g3) | 实际净值(g4) | 实际累计净值(g5)
            # Linda records 实际累计净值 (g5) as her 单位净值.
            m = re.search(r"\|\s*(\d{4}-\d{2}-\d{2})\s*\|\s*[^|]*计提\s*\|\s*"+_NUM+r"\s*\|\s*"+_NUM+r"\s*\|\s*"+_NUM+r"\s*\|\s*"+_NUM, body)
            if m:
                return dict(code=subj_code, date=parse_date(m.group(1)),
                            unit=num(m.group(5)), cum=num(m.group(5)),
                            actual=num(m.group(4)), virtual=num(m.group(3)), source="citics_virtual")
        else:
            # CODE(级) | name | date | unit | cum | ...
            m = re.search(r"([A-Z0-9]{5,8}(?:[（(][^）)]*[）)])?)\s*\|\s*[^|]+\|\s*(\d{4}-\d{2}-\d{2})\s*\|\s*"+_NUM+r"\s*\|\s*"+_NUM, body)
            if m:
                return dict(code=base_code(m.group(1)), date=parse_date(m.group(2)),
                            unit=num(m.group(3)), cum=num(m.group(4)), virtual=None, source="citics_real")
    elif fmt == "htsc_glrfw":
        m = re.search(r"\|\s*([A-Z0-9]{5,8})\s*\|\s*[^|]+\|\s*(\d{4}\D\d{1,2}\D\d{1,2}\D?)\s*\|\s*"+_NUM+r"\s*\|\s*"+_NUM, body)
        if m:
            return dict(code=base_code(m.group(1)), date=parse_date(m.group(2)),
                        unit=num(m.group(3)), cum=num(m.group(4)), virtual=None, source="htsc_glrfw")
    elif fmt == "htsc_incos":
        m = re.search(r"(\d{4}-\d{2}-\d{2})\s*\|\s*([A-Z0-9]{5,8})\s*\|\s*[^|]+\|\s*"+_NUM+r"\s*\|\s*"+_NUM, body)
        if m:
            return dict(code=base_code(m.group(2)), date=parse_date(m.group(1)),
                        unit=num(m.group(3)), cum=num(m.group(4)), virtual=None, source="htsc_incos")
    elif fmt == "xyzq":
        c = re.search(r"产品代码[:：]\s*([A-Z0-9]+)", body)
        d = re.search(r"净值日期[:：]\s*(\d{4}-\d{2}-\d{2})", body)
        u = re.search(r"产品单位净值[:：]\s*"+_NUM, body)
        cu = re.search(r"产品累计单位净值[:：]\s*"+_NUM, body)
        if c and d and u:
            return dict(code=base_code(c.group(1)), date=parse_date(d.group(1)),
                        unit=num(u.group(1)), cum=num(cu.group(1)) if cu else None,
                        virtual=None, source="xyzq")
    return None

# ----------------------------------------------------------------------------- IMAP
def get_password():
    """Resolve the QQ IMAP authorization code from a NON-synced location, so it
    never lives in the OneDrive-synced config: env var -> %LOCALAPPDATA% secret
    -> config.json (legacy fallback)."""
    pw = os.environ.get("NAV_QQ_PW")
    if pw:
        return pw
    sp = os.path.join(os.environ.get("LOCALAPPDATA", ""), "nav_tool", "secret.json")
    if os.path.exists(sp):
        try:
            v = json.load(open(sp, encoding="utf-8")).get("password")
            if v:
                return v
        except Exception:
            pass
    return CFG["imap"].get("password") or ""

def connect():
    c = CFG["imap"]
    M = imaplib.IMAP4_SSL(c["host"], c["port"])
    M.login(c["user"], get_password())
    try:
        M._simple_command("ID", '("name" "navtool" "version" "1.0")')
    except Exception:
        pass
    M.select(c["mailbox"], readonly=True)
    return M

# NOTE: all message access uses IMAP UIDs (stable), never sequence numbers,
# because this is a live inbox whose sequence numbers shift as mail arrives.
def search_from(M, sender):
    typ, data = M.uid("SEARCH", None, '(FROM "%s")' % sender)
    if typ != "OK" or not data or not data[0]:
        return []
    return data[0].split()

def fetch_headers(M, uids, fields="(DATE SUBJECT)"):
    """Batch UID-fetch header fields. Returns {uid:int -> msg}, keyed by true UID."""
    out = {}
    B = 200
    for i in range(0, len(uids), B):
        chunk = b",".join(uids[i:i+B])
        typ, data = M.uid("FETCH", chunk, "(UID BODY.PEEK[HEADER.FIELDS %s])" % fields)
        if typ != "OK":
            continue
        for part in data:
            if isinstance(part, tuple):
                m = re.search(rb"UID\s+(\d+)", part[0])
                if not m:
                    continue
                out[int(m.group(1))] = email.message_from_bytes(part[1])
    return out

def fetch_full(M, uid):
    u = uid if isinstance(uid, bytes) else str(uid).encode()
    for _ in range(3):
        typ, md = M.uid("FETCH", u, "(RFC822)")
        if typ == "OK":
            for part in md:
                if isinstance(part, tuple) and part[1]:
                    return email.message_from_bytes(part[1])
    return None
