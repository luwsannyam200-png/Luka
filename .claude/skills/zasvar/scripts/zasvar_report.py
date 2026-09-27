#!/usr/bin/env python3
"""Repair request report from the Mine2TL export ("засвар ...rqst.xlsx").

Output workbook:
  Тайлан    rows Засварын байршил > Үйлчилгээ, columns = request day (report layout)
  Pivot     rows Засварын төрөл > Засварын байршил > Үйлчилгээ, columns = request day, values = count
  Хүлээлт   same rows, columns = days waiting (0, 1-3, 4-7, 8+)
  + the original export sheet(s), unchanged

Usage:
  zasvar_report.py --input "засвар 9-26rqst.xlsx" --out "Засварын хүсэлт_09.27.xlsx" [--today 2026-09-27]
--today defaults to the day of the latest request in the file.
"""
import argparse
import datetime as dt
import json
import re
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BLANK = "(хоосон)"
BUCKETS = [("Өнөөдөр", 0, 0), ("1–3 хоног", 1, 3), ("4–7 хоног", 4, 7), ("8+ хоног", 8, 10 ** 6)]

HEAD = PatternFill("solid", fgColor="B4E5D2")
SUB1 = PatternFill("solid", fgColor="DDF3EA")
SUB2 = PatternFill("solid", fgColor="F0FAF5")
RED = PatternFill("solid", fgColor="F8D2CE")
THIN = Border(bottom=Side(style="thin", color="C8D3CE"))


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def to_dt(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    if v is None or str(v).strip() == "":
        return None
    s = str(v).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%m/%d/%Y %H:%M", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def read_requests(path):
    # The data may be on any sheet (e.g. after a pivot sheet); use the first
    # sheet that has the Mine2TL header row.
    found = None
    for ws in openpyxl.load_workbook(path, data_only=True).worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for hi, r in enumerate(rows[:30]):
            names = [norm(c) for c in r]
            if "засварын төрөл" in names and "үйлчилгээ" in names and "засварын байршил" in names:
                found = (ws.title, rows, hi, names)
                break
        if found:
            break
    if not found:
        sys.exit('Гарчгийн мөр олдсонгүй ("Засварын төрөл", "Засварын байршил", "Үйлчилгээ")')
    _, rows, hi, names = found

    def find(pred, label):
        for i, n in enumerate(names):
            if pred(n):
                return i
        sys.exit(f'"{label}" багана олдсонгүй')

    ix = {
        "type": find(lambda n: n == "засварын төрөл", "Засварын төрөл"),
        "tech": find(lambda n: n.startswith("техни"), "Техникийн №"),
        "loc": find(lambda n: n == "засварын байршил", "Засварын байршил"),
        "srv": find(lambda n: n == "үйлчилгээ", "Үйлчилгээ"),
        "desc": find(lambda n: n == "тайлбар", "Тайлбар"),
        "plan": find(lambda n: n.startswith("төлөвлөгөөт цаг"), "Төлөвлөгөөт цаг"),
        "load": find(lambda n: n.startswith("ачаатай"), "Ачаатай/ачаагүй"),
        "req": find(lambda n: n.startswith("хүсэлт гаргасан"), "Хүсэлт гаргасан огноо"),
    }
    out, skipped = [], 0
    for r in rows[hi + 1:]:
        if r[ix["tech"]] is None or r[ix["type"]] is None:
            if any(c is not None for c in r):
                skipped += 1  # totals row etc.
            continue
        req = to_dt(r[ix["req"]])
        if req is None:
            skipped += 1
            continue
        out.append({
            "type": str(r[ix["type"]]).strip() or BLANK,
            "tech": str(r[ix["tech"]]).strip(),
            "loc": str(r[ix["loc"]] or "").strip() or BLANK,
            "srv": str(r[ix["srv"]] or "").strip() or BLANK,
            "desc": str(r[ix["desc"]] or "").strip(),
            "plan": to_dt(r[ix["plan"]]),
            "load": str(r[ix["load"]] or "").strip(),
            "req": req,
        })
    return out, skipped


def header(ws, row, values, widths=None):
    for c, v in enumerate(values, 1):
        cell = ws.cell(row, c, v)
        cell.font = Font(bold=True)
        cell.fill = HEAD
        cell.alignment = Alignment(horizontal="center" if c > 1 else "left", vertical="center", wrap_text=True)
    if widths:
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w


def tree_table(ws, reqs, col_keys, col_of, title_cells, first_width=32):
    """Rows: type > location > service, with subtotals; columns from col_keys."""
    header(ws, 1, title_cells + ["Нийт"], [first_width] + [9] * (len(col_keys) + 1))
    ws.freeze_panes = "B2"
    tree = defaultdict(lambda: defaultdict(lambda: defaultdict(Counter)))
    for q in reqs:
        tree[q["type"]][q["loc"]][q["srv"]][col_of(q)] += 1

    def write(r, label, counter, level):
        ws.cell(r, 1, ("    " * level) + label)
        total = 0
        for c, k in enumerate(col_keys, 2):
            v = counter.get(k, 0)
            total += v
            if v:
                ws.cell(r, c, v)
        ws.cell(r, len(col_keys) + 2, total)
        fill = SUB1 if level == 0 else SUB2 if level == 1 else None
        for c in range(1, len(col_keys) + 3):
            cell = ws.cell(r, c)
            cell.border = THIN
            if fill:
                cell.fill = fill
            if level < 2:
                cell.font = Font(bold=True)
            if c > 1:
                cell.alignment = Alignment(horizontal="center")

    r = 2
    for t in sorted(tree):
        t_sum = Counter()
        for loc in tree[t].values():
            for srv in loc.values():
                t_sum.update(srv)
        write(r, t, t_sum, 0)
        r += 1
        for loc in sorted(tree[t]):
            l_sum = Counter()
            for srv in tree[t][loc].values():
                l_sum.update(srv)
            write(r, loc, l_sum, 1)
            r += 1
            # Blank services count in the location row but are not listed.
            for srv in sorted(x for x in tree[t][loc] if x != BLANK):
                write(r, srv, tree[t][loc][srv], 2)
                r += 1
    g = Counter(col_of(q) for q in reqs)
    write(r, "Нийт", g, 0)
    for c in range(1, len(col_keys) + 3):
        ws.cell(r, c).fill = HEAD
    return r


NAVY = PatternFill("solid", fgColor="1F3864")
BLUE = PatternFill("solid", fgColor="2F5597")
LIGHT = PatternFill("solid", fgColor="8EA9DB")
GRID = Border(*(Side(style="thin", color="1F3864"),) * 4)


def report_table(ws, reqs, days, top=4, left=3):
    """Presentation table: location (dark) > service (light) rows, request-day
    columns, "Хүсэлт" total column and "Нийт хүсэлт" row. Blank services are
    counted in the location row but not listed."""
    cnt = defaultdict(lambda: defaultdict(Counter))
    for q in reqs:
        cnt[q["loc"]][q["srv"]][q["req"].date()] += 1
    ncol = len(days) + 2

    def put(r, label, counter, fill, white):
        vals = [label] + [counter.get(d, 0) or None for d in days] + [sum(counter.values())]
        for i, v in enumerate(vals):
            c = ws.cell(r, left + i, v)
            c.fill = fill
            c.border = GRID
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.font = Font(bold=white, color="FFFFFF" if white else "000000")

    put(top, "Засварын төрөл", {}, NAVY, True)
    for i, d in enumerate(days, 1):
        ws.cell(top, left + i, d.strftime("%d-%b"))
    ws.cell(top, left + ncol - 1, "Хүсэлт")
    r = top + 1
    for loc in sorted(cnt):
        total = Counter()
        for s in cnt[loc].values():
            total.update(s)
        services = sorted(s for s in cnt[loc] if s != BLANK)
        # A location with no listed services (e.g. Замд) is styled like a service row.
        put(r, loc, total, BLUE if services else LIGHT, bool(services))
        r += 1
        for srv in services:
            put(r, srv, cnt[loc][srv], LIGHT, False)
            r += 1
    put(r, "Нийт хүсэлт", Counter(q["req"].date() for q in reqs), NAVY, True)
    ws.column_dimensions[get_column_letter(left)].width = 20
    for i in range(1, ncol):
        ws.column_dimensions[get_column_letter(left + i)].width = 10
    return r


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--today")
    a = ap.parse_args()

    reqs, skipped = read_requests(a.input)
    if not reqs:
        sys.exit("Засварын хүсэлт олдсонгүй")
    today = dt.date.fromisoformat(a.today) if a.today else max(q["req"] for q in reqs).date()
    for q in reqs:
        q["days"] = (today - q["req"].date()).days

    # Keep the original export as-is and put the report sheets in front of it.
    wb = openpyxl.load_workbook(a.input)
    src_names = wb.sheetnames

    def free(name):
        n, k = name, 2
        while n in wb.sheetnames:
            n, k = f"{name} ({k})", k + 1
        return n

    days = sorted({q["req"].date() for q in reqs})

    # 1. Тайлан: presentation table (location > service by request day)
    ws0 = wb.create_sheet(free("Тайлан"), 0)
    report_table(ws0, reqs, days)

    # 2. Pivot: same with Засварын төрөл as the top level
    ws = wb.create_sheet(free("Pivot"), 1)
    tree_table(ws, reqs, days, lambda q: q["req"].date(),
               ["Засварын төрөл / байршил / үйлчилгээ"] + [d.strftime("%d-%b") for d in days])

    # 3. Хүлээлт: columns = waiting buckets
    ws2 = wb.create_sheet(free("Хүлээлт"), 2)
    labels = [b[0] for b in BUCKETS]
    bucket = lambda q: next(b[0] for b in BUCKETS if b[1] <= q["days"] <= b[2])
    last = tree_table(ws2, reqs, labels, bucket, [f"Өнөөдөр: {today:%Y-%m-%d}"] + labels)
    for r in range(2, last + 1):
        c = ws2.cell(r, 1 + len(labels))  # 8+ column
        if c.value:
            c.fill = RED

    for w in wb.worksheets:
        w.sheet_view.tabSelected = w is ws0
    wb.active = 0
    wb.save(a.out)
    json.dump({
        "requests": len(reqs), "skipped_rows": skipped, "today": str(today),
        "by_type": Counter(q["type"] for q in reqs), "by_location": Counter(q["loc"] for q in reqs),
        "by_service": Counter(q["srv"] for q in reqs), "by_day": {str(d): sum(1 for q in reqs if q["req"].date() == d) for d in days},
        "waiting": Counter(bucket(q) for q in reqs), "source_sheets": src_names,
    }, sys.stdout, ensure_ascii=False, indent=1)
    print()


if __name__ == "__main__":
    main()
