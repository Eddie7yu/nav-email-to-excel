# -*- coding: utf-8 -*-
"""Benchmark index daily closes from Eastmoney's public API (only needs `requests`,
no token, works well from mainland China). Returns { 'YYYY-MM-DD': close }.

secid format: <market>.<code>   market 1=SH, 0=SZ/BSE, 2=CSI(中证)
"""
import time, datetime, requests

# index name (as written in the sheet header, column G) -> Eastmoney secid
SECID = {
    "中证500":  ["1.000905"],
    "中证1000": ["1.000852"],
    "中证2000": ["2.932000", "1.932000", "0.932000"],
    "北证50":   ["0.899050"],
    "北证等权":  ["0.899050", "2.899050"],
}
HOSTS = ["https://push2his.eastmoney.com", "https://push2his.eastmoney.com",
         "http://push2his.eastmoney.com"]
PATH = "/api/qt/stock/kline/get"
UT = "fa5fd1943c7b386f172d6893dbfba10b"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Referer": "https://quote.eastmoney.com/",
    "Accept": "application/json, text/plain, */*",
}

def _try(secid, start, end):
    params = dict(secid=secid, ut=UT, fields1="f1", fields2="f51,f53",
                  klt="101", fqt="0", beg=start, end=end, lmt="100000")
    last_exc = None
    for attempt in range(4):
        try:
            host = HOSTS[min(attempt, len(HOSTS) - 1)]
            r = requests.get(host + PATH, params=params, headers=HEADERS, timeout=20)
            data = (r.json() or {}).get("data") or {}
            out = {}
            for line in data.get("klines", []):
                parts = line.split(",")
                if len(parts) >= 2:
                    out[parts[0]] = float(parts[1])
            if out:
                return out, data.get("name")
        except Exception as e:
            last_exc = e
        time.sleep(1.5 * (attempt + 1))
    return {}, None

def fetch_closes(index_name, start="20250101", end=None):
    end = end or datetime.date.today().strftime("%Y%m%d")
    for secid in SECID.get(index_name, []):
        closes, name = _try(secid, start, end)
        if closes:
            return closes, secid, name
        time.sleep(1.0)
    return {}, None, None

if __name__ == "__main__":
    for nm in ["中证500", "中证1000", "中证2000", "北证50", "北证等权"]:
        c, secid, name = fetch_closes(nm, start="20260601")
        last = sorted(c)[-1] if c else None
        print(f"{nm:8s} secid={secid} name={name} rows={len(c)} last={last}={c.get(last) if last else None}")
