#!/usr/bin/env python3
"""Add a new weekly sheet to the "Weekly draft" workbook.

The new sheet is a copy of the latest sheet with the stock columns refreshed:
  УХ  = Oracle stock in subinventories starting with TL
  ЦХ  = Oracle stock in subinventories starting with TKH
  AODE / Очлуур од / Parts and oil = supplier stock files (optional)

The workbook is edited at the XML level so every existing sheet, cached
formula value and external link stays byte-identical. Saving through
openpyxl would drop cached values and break the external-link formulas.

Usage:
  build_weekly.py --draft "Weekly draft.xlsx" --stock "09-26 үлдэгдэл.xlsx" \
      --date 09.26 --out "Weekly draft_09.26.xlsx" \
      [--parts "Parts and Oil.xlsx"] [--ochluur "Очлуур-Од.xlsx"] [--aode "AODE.xlsx"] \
      [--manual manual.json]

manual.json supplies supplier values that could not be read from a file:
  {"Очлуур од": {"4016150400048": 167, ...}}
Suppliers with no file and no manual values are left blank and yellow.
"""
import argparse
import json
import re
import sys
import zipfile
from collections import defaultdict

import openpyxl
from openpyxl.utils import get_column_letter

SUPPLIERS = {"aode": "AODE", "ochluur": "Очлуур од", "parts": "Parts and oil"}
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def num(x):
    x = round(float(x), 6)
    return str(int(x)) if x == int(x) else repr(x)


def read_oracle_stock(path):
    """Sum Quantity per item code for TL* and TKH* subinventories.

    The report may sit on any sheet (a TL_TKH summary sheet is often added in
    front of it). Without the detailed report, a summary sheet with
    Item Code / TL / TKH columns is read as is."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    found = summary = None
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for i, r in enumerate(rows[:30]):
            names = [norm(c) for c in r]
            if "item code" not in names:
                continue
            if "subinventory" in names and "quantity" in names:
                found = (rows[i + 1:], {k: names.index(k) for k in ("item code", "subinventory", "quantity")})
            elif summary is None and "tl" in names and "tkh" in names:
                summary = (rows[i + 1:], {k: names.index(k) for k in ("item code", "tl", "tkh")})
            break
        if found:
            break
    tl, tkh = defaultdict(float), defaultdict(float)
    if found:
        rows, ix = found
        for r in rows:
            code = r[ix["item code"]]
            if code is None or not isinstance(r[ix["quantity"]], (int, float)):
                continue  # wrapped description lines, totals
            code = str(code).strip()
            sub = str(r[ix["subinventory"]] or "").strip().upper()
            if sub.startswith("TKH"):
                tkh[code] += r[ix["quantity"]]
            elif sub.startswith("TL"):
                tl[code] += r[ix["quantity"]]
        return tl, tkh
    if summary:
        rows, ix = summary
        for r in rows:
            code = str(r[ix["item code"]] or "").strip()
            if not code or re.search(r"total|нийт|дүн", code, re.I):
                continue  # totals
            if isinstance(r[ix["tl"]], (int, float)):
                tl[code] += r[ix["tl"]]
            if isinstance(r[ix["tkh"]], (int, float)):
                tkh[code] += r[ix["tkh"]]
        return tl, tkh
    sys.exit("Oracle stock file: header row with Item Code/Subinventory/Quantity (or Item Code/TL/TKH) not found")


def read_supplier(path):
    """Supplier files differ in layout: take the 13-digit item code in each row
    and the last numeric cell after it as the quantity."""
    out = defaultdict(float)
    for ws in openpyxl.load_workbook(path, data_only=True).worksheets[:1]:
        for r in ws.iter_rows(values_only=True):
            cells = list(r)
            ci = next((i for i, c in enumerate(cells) if re.fullmatch(r"\d{13}", str(c).strip() if c is not None else "")), None)
            if ci is None:
                continue
            q = None
            for c in cells[ci + 1:]:
                if isinstance(c, (int, float)):
                    q = c
                elif isinstance(c, str) and re.fullmatch(r"[\d,]+(\.\d*)?", c.strip()):
                    q = float(c.strip().replace(",", "").rstrip("."))
            if q is not None:
                out[str(cells[ci]).strip()] += q
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--draft", required=True)
    ap.add_argument("--stock", required=True)
    ap.add_argument("--date", required=True, help="new sheet name, e.g. 09.26")
    ap.add_argument("--out", required=True)
    for k in SUPPLIERS:
        ap.add_argument(f"--{k}")
    ap.add_argument("--manual")
    a = ap.parse_args()

    tl, tkh = read_oracle_stock(a.stock)
    supplier = {}
    for k, label in SUPPLIERS.items():
        if getattr(a, k):
            supplier[label] = read_supplier(getattr(a, k))
    if a.manual:
        for label, vals in json.load(open(a.manual, encoding="utf-8")).items():
            supplier[label] = defaultdict(float, {str(c): v for c, v in vals.items()})

    z = zipfile.ZipFile(a.draft)
    files = {n: z.read(n) for n in z.namelist()}
    order = z.namelist()
    wbx = files["xl/workbook.xml"].decode()
    rels = files["xl/_rels/workbook.xml.rels"].decode()

    sheets = re.findall(r'<sheet [^>]*?name="([^"]*)"[^>]*?r:id="(rId\d+)"', wbx)
    if any(n == a.date for n, _ in sheets):
        sys.exit(f'Sheet "{a.date}" already exists')
    tpl_name, tpl_rid = sheets[-1]
    tpl_target = re.search(rf'Id="{tpl_rid}"[^>]*Target="([^"]+)"', rels) or re.search(rf'Target="([^"]+)"[^>]*Id="{tpl_rid}"', rels)
    tpl_path = "xl/" + tpl_target.group(1).lstrip("/").replace("xl/", "")

    # Header lookup on the template sheet (column positions have shifted over time).
    ws = openpyxl.load_workbook(a.draft, data_only=True)[tpl_name]
    col = {}
    for c in range(1, ws.max_column + 1):
        for r in (1, 2):
            h = norm(ws.cell(r, c).value)
            if h:
                col.setdefault(h, c)
    def need(*names):
        for n in names:
            if norm(n) in col:
                return get_column_letter(col[norm(n)])
        sys.exit(f"Template sheet {tpl_name}: column {names} not found")
    C = {
        "code": need("item code"), "zbn": need("ЗБН"), "ux": need("УХ"), "cx": need("ЦХ"),
        "tlw": need("TL Warehouse"), "total": need("Нийт агуулах"), "all": need("All warehouse"),
        "name": need("Сэлбэгийн мэдээлэл"),
    }
    SUPCOL = {label: need(label) for label in SUPPLIERS.values()}
    last_sup = get_column_letter(max(col[norm(k)] for k in SUPCOL))
    first_row = 3
    last_row = max(r for r in range(first_row, ws.max_row + 1) if ws[f"{C['code']}{r}"].value is not None)

    sheet = files[tpl_path].decode()
    styles = files["xl/styles.xml"].decode()
    fills = re.search(r'<fills count="(\d+)">(.*?)</fills>', styles, re.S)
    nfill = int(fills.group(1))
    styles = styles.replace(fills.group(0), f'<fills count="{nfill + 1}">{fills.group(2)}<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>')
    cx = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles, re.S)
    xfs = re.findall(r"<xf [^>]*?(?:/>|>.*?</xf>)", cx.group(2), re.S)
    assert len(xfs) == int(cx.group(1)), "unexpected styles.xml layout"
    yellow = {}

    def ystyle(s):
        if s not in yellow:
            x = re.sub(r' fillId="\d+"', f' fillId="{nfill}"', xfs[int(s)])
            x = re.sub(r' applyFill="\d"', "", x).replace("<xf ", '<xf applyFill="1" ', 1)
            yellow[s] = str(len(xfs))
            xfs.append(x)
        return yellow[s]

    def setcell(ref, inner, fill_yellow=False):
        nonlocal sheet
        m = re.search(rf'<c r="{ref}"( [^>]*?)?(/>|>.*?</c>)', sheet)
        assert m, f"cell {ref} missing in template"
        s = re.search(r' s="(\d+)"', m.group(1) or "")
        sattr = f' s="{ystyle(s.group(1)) if fill_yellow else s.group(1)}"' if s else ""
        new = f'<c r="{ref}"{sattr}>{inner}</c>' if inner else f'<c r="{ref}"{sattr}/>'
        sheet = sheet[: m.start()] + new + sheet[m.end():]

    report = []
    for r in range(first_row, last_row + 1):
        code = str(ws[f"{C['code']}{r}"].value).strip()
        zbn = ws[f"{C['zbn']}{r}"].value or 0
        ux, cxv = tl.get(code, 0), tkh.get(code, 0)
        setcell(f"{C['ux']}{r}", f"<v>{num(ux)}</v>")
        setcell(f"{C['cx']}{r}", f"<v>{num(cxv)}</v>")
        total = ux + cxv
        row = {"row": r, "name": ws[f"{C['name']}{r}"].value, "code": code, "УХ": ux, "ЦХ": cxv}
        for label, L in SUPCOL.items():
            if label in supplier:
                v = supplier[label].get(code, 0)
                setcell(f"{L}{r}", f"<v>{num(v)}</v>")
                total += v
                row[label] = v
            else:
                setcell(f"{L}{r}", "", True)
                row[label] = None
        rng = f"{C['ux']}{r}:{last_sup}{r}"
        tlw = (ux + cxv) / zbn if zbn else 0
        setcell(f"{C['tlw']}{r}", f"<f>SUM({C['ux']}{r}:{C['cx']}{r})/{C['zbn']}{r}</f><v>{num(tlw)}</v>")
        setcell(f"{C['total']}{r}", f"<f>SUM({rng})</f><v>{num(total)}</v>")
        setcell(f"{C['all']}{r}", f"<f>{C['total']}{r}/{C['zbn']}{r}</f><v>{num(total / zbn if zbn else 0)}</v>")
        row["prev"] = {"УХ": ws[f"{C['ux']}{r}"].value, "ЦХ": ws[f"{C['cx']}{r}"].value,
                       **{k: ws[f"{L}{r}"].value for k, L in SUPCOL.items()}}
        report.append(row)

    sheet = re.sub(r' xr:uid="\{[^}]*\}"', "", sheet, count=1)
    sheet = re.sub(r'(<pageSetup[^>]*?) r:id="rId\d+"', r"\1", sheet)
    sheet = re.sub(r"<selection [^>]*/>", '<selection activeCell="A1" sqref="A1"/>', sheet, count=1)
    files["xl/styles.xml"] = re.sub(r'<cellXfs count="\d+">.*?</cellXfs>',
                                    lambda m: f'<cellXfs count="{len(xfs)}">{"".join(xfs)}</cellXfs>', styles, flags=re.S).encode()

    n = 1
    while f"xl/worksheets/sheet{n}.xml" in files:
        n += 1
    newpath = f"xl/worksheets/sheet{n}.xml"
    files[newpath] = sheet.encode()
    # Only the new sheet is selected.
    for p in list(files):
        if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", p) and p != newpath:
            files[p] = files[p].replace(b'<sheetView tabSelected="1" ', b"<sheetView ", 1)

    sid = max(map(int, re.findall(r'<sheet [^>]*sheetId="(\d+)"', wbx))) + 1
    rid = max(map(int, re.findall(r'Id="rId(\d+)"', rels))) + 1
    wbx = wbx.replace("</sheets>", f'<sheet name="{a.date}" sheetId="{sid}" r:id="rId{rid}"/></sheets>')
    wbx = re.sub(r'activeTab="\d+"', f'activeTab="{len(sheets)}"', wbx)
    files["xl/workbook.xml"] = wbx.encode()
    files["xl/_rels/workbook.xml.rels"] = rels.replace(
        "</Relationships>",
        f'<Relationship Id="rId{rid}" Type="{REL}/worksheet" Target="worksheets/sheet{n}.xml"/></Relationships>').encode()
    files["[Content_Types].xml"] = files["[Content_Types].xml"].decode().replace(
        "</Types>",
        f'<Override PartName="/{newpath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>').encode()
    order.insert(order.index(tpl_path) + 1, newpath)

    with zipfile.ZipFile(a.out, "w", zipfile.ZIP_DEFLATED) as o:
        for name in order:
            o.writestr(z.getinfo(name) if name in z.namelist() else name, files[name])

    json.dump({"template": tpl_name, "new_sheet": a.date, "suppliers_filled": sorted(supplier),
               "suppliers_blank": sorted(set(SUPPLIERS.values()) - set(supplier)), "rows": report},
              sys.stdout, ensure_ascii=False, indent=1, default=str)
    print()


if __name__ == "__main__":
    main()
