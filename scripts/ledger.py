#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ledger.py - 文献台账：查重、写入、列出

存储：CSV（utf-8-sig，Excel 可直接打开），默认
  C:\\Users\\WIT_User\\Desktop\\OH-WorkSpace\\文献总结\\文献台账.csv

子命令：
  check  --doi 10.xxxx/xxx                     是否写过这篇（返回 matched / not_found）
  add    --title ... --doi ... --innovation ... --borrowable ... [--journal] [--year] [--docx]
  list   [--limit 20]
  remove --doi 10.xxxx/xxx                     写错了撤回一条
"""

import argparse
import csv
import datetime as dt
import json
import os
import re
import sys

DEFAULT_LEDGER = os.environ.get(
    "LIT_SUMMARY_LEDGER",
    os.path.join(os.path.expanduser("~"), "Desktop", "OH-WorkSpace", "文献总结", "文献台账.csv"),
)
FIELDS = ["date", "doi", "title", "journal", "year", "innovation", "borrowable", "docx"]


def norm_doi(doi):
    if not doi:
        return ""
    d = doi.strip().lower()
    d = re.sub(r"^(https?://(dx\.)?doi\.org/|doi:\s*)", "", d)
    return d.strip().rstrip(".")


def norm_title(title):
    t = (title or "").lower()
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", t)


def load(path):
    if not os.path.isfile(path):
        return []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def save(path, rows):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in FIELDS})


def find(rows, doi, title):
    d, t = norm_doi(doi), norm_title(title)
    for r in rows:
        if d and norm_doi(r.get("doi")) == d:
            return r
        if not d and t and norm_title(r.get("title")) == t:
            return r
    return None


def cmd_check(args):
    rows = load(args.ledger)
    hit = find(rows, args.doi, args.title)
    if hit:
        print(json.dumps({"status": "matched", "entry": hit}, ensure_ascii=False, indent=2))
    else:
        print(json.dumps({"status": "not_found", "total": len(rows)}, ensure_ascii=False))


def cmd_add(args):
    rows = load(args.ledger)
    hit = find(rows, args.doi, args.title)
    if hit and not args.force:
        print(json.dumps({"status": "duplicate", "entry": hit}, ensure_ascii=False, indent=2))
        return
    entry = {
        "date": args.date or dt.date.today().strftime("%Y-%m-%d"),
        "doi": norm_doi(args.doi),
        "title": args.title or "",
        "journal": args.journal or "",
        "year": str(args.year) if args.year else "",
        "innovation": args.innovation or "",
        "borrowable": args.borrowable or "",
        "docx": args.docx or "",
    }
    if hit:
        rows[rows.index(hit)] = entry
        status = "updated"
    else:
        rows.append(entry)
        status = "added"
    save(args.ledger, rows)
    print(json.dumps({"status": status, "entry": entry, "total": len(rows)}, ensure_ascii=False, indent=2))


def cmd_list(args):
    rows = load(args.ledger)
    rows = list(reversed(rows))[: args.limit]
    print(json.dumps({"total_in_ledger": len(load(args.ledger)), "shown": len(rows), "entries": rows},
                     ensure_ascii=False, indent=2))


def cmd_remove(args):
    rows = load(args.ledger)
    hit = find(rows, args.doi, args.title)
    if not hit:
        print(json.dumps({"status": "not_found"}, ensure_ascii=False))
        return
    rows.remove(hit)
    save(args.ledger, rows)
    print(json.dumps({"status": "removed", "entry": hit, "total": len(rows)}, ensure_ascii=False, indent=2))


def main():
    ap = argparse.ArgumentParser(description="文献台账管理")
    ap.add_argument("--ledger", default=DEFAULT_LEDGER)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("check"); p.add_argument("--doi", default=""); p.add_argument("--title", default="")
    p.set_defaults(func=cmd_check)

    p = sub.add_parser("add")
    for name in ("--doi", "--title", "--journal", "--innovation", "--borrowable", "--docx", "--date"):
        p.add_argument(name, default="")
    p.add_argument("--year", type=int)
    p.add_argument("--force", action="store_true", help="同 DOI 已存在时覆盖")
    p.set_defaults(func=cmd_add)

    p = sub.add_parser("list"); p.add_argument("--limit", type=int, default=20); p.set_defaults(func=cmd_list)

    p = sub.add_parser("remove"); p.add_argument("--doi", default=""); p.add_argument("--title", default="")
    p.set_defaults(func=cmd_remove)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
