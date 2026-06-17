# -*- coding: utf-8 -*-
"""Fetch headers for every in-scope sender, parse code+date+subtype from the
subject, and cache a lightweight index to index.json. Bodies are NOT fetched here.
"""
import os, json, datetime
import navlib as L
import phase2 as P2

def main():
    M = L.connect()
    c2s = L.code_to_sheet()
    p2_code2sheet, p2_names = P2.build_routers()
    index = []
    counts = {}
    for sender, fmt in L.CFG["senders"].items():
        uids = L.search_from(M, sender)
        counts[sender] = len(uids)
        print(f"[{sender}] {fmt}: {len(uids)} messages -> fetching headers...")
        hdrs = L.fetch_headers(M, uids)
        for uid, msg in hdrs.items():
            subj = L.dh(msg.get("Subject"))
            if fmt in P2.FMTS:
                sheet, code, date = P2.route(subj, p2_code2sheet, p2_names)
                subtype, unit_hint = None, None
            else:
                code, date, subtype, unit_hint = L.parse_subject(fmt, subj)
                sheet = c2s.get(code) if code else None
            if not sheet:
                continue  # not an in-scope product
            index.append({
                "uid": uid,
                "sender": sender,
                "fmt": fmt,
                "code": code,
                "sheet": sheet,
                "date": date.isoformat() if date else None,
                "subtype": subtype,
                "unit_hint": unit_hint,
                "subject": subj,
            })
    M.logout()
    index.sort(key=lambda r: (r["sheet"], r["date"] or "", r["subtype"] or ""))
    out = os.path.join(L.HERE, "index.json")
    json.dump({"built": datetime.datetime.now().isoformat(), "counts": counts, "index": index},
              open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    # summary per sheet
    per = {}
    for r in index:
        per.setdefault(r["sheet"], []).append(r["date"])
    print("\n=== in-scope messages indexed per sheet ===")
    for sheet in L.all_scope():
        ds = sorted(d for d in per.get(sheet, []) if d)
        print(f"{sheet:10s} n={len(per.get(sheet, [])):4d}  latest={ds[-1] if ds else '-'}  earliest={ds[0] if ds else '-'}")
    print(f"\nTotal in-scope index rows: {len(index)}  -> {out}")

if __name__ == "__main__":
    main()
