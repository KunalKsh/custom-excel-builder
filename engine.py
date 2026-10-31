"""Excel engine that runs inside the browser (Pyodide). Files are read straight from disk, never uploaded."""
import csv
import datetime
import json
import os
import re

import xlsxwriter

MAX_ROWS = 1048575
DATE_SERIAL_MIN = 20000
FILES = {}  # id -> {"name","path","kind"}


def kind_of(name):
    e = os.path.splitext(name)[1].lower()
    return {".xlsb": "xlsb", ".xlsx": "xlsx", ".xlsm": "xlsx", ".xls": "xls", ".csv": "csv", ".tsv": "csv"}.get(e)


def sheet_names(f):
    k, p = f["kind"], f["path"]
    if k == "xlsb":
        from pyxlsb import open_workbook
        with open_workbook(p) as wb:
            return list(wb.sheets)
    if k == "xlsx":
        from openpyxl import load_workbook
        wb = load_workbook(p, read_only=True)
        n = wb.sheetnames
        wb.close()
        return n
    if k == "xls":
        import xlrd
        return xlrd.open_workbook(p, on_demand=True).sheet_names()
    return ["Sheet1"]


def iter_rows(f, sheet, want, start):
    """Yield (row_number0, {col: value}) for rows at/after `start`, only for columns in `want`."""
    k, p = f["kind"], f["path"]
    want = set(want)
    if k == "xlsb":
        from pyxlsb import open_workbook
        with open_workbook(p) as wb, wb.get_sheet(sheet) as sh:
            for row in sh.rows():
                if not row or row[0].r < start:
                    continue
                yield row[0].r, {c.c: c.v for c in row if c.c in want and c.v is not None}
    elif k == "xlsx":
        from openpyxl import load_workbook
        wb = load_workbook(p, read_only=True, data_only=True)
        for n, row in enumerate(wb[sheet].iter_rows(values_only=True)):
            if n >= start:
                yield n, {i: v for i, v in enumerate(row) if i in want and v is not None}
        wb.close()
    elif k == "xls":
        import xlrd
        ws = xlrd.open_workbook(p, on_demand=True).sheet_by_name(sheet)
        for n in range(start, ws.nrows):
            yield n, {i: c.value for i, c in enumerate(ws.row(n)) if i in want and c.ctype not in (0, 5, 6)}
    else:
        delim = "\t" if p.lower().endswith(".tsv") else ","
        with open(p, newline="", encoding="utf-8-sig", errors="replace") as fh:
            for n, row in enumerate(csv.reader(fh, delimiter=delim)):
                if n >= start:
                    yield n, {i: v for i, v in enumerate(row) if i in want and v != ""}


def headers_for(f, sheet, hrow):
    n = hrow - 1
    for r, d in iter_rows(f, sheet, range(2000), n):
        if r == n:
            return [{"idx": i, "label": str(v).strip()} for i, v in sorted(d.items()) if str(v).strip()]
        break
    return []


def file_payload(fid, hrow):
    f = FILES[fid]
    if f["kind"] == "xlsb":
        # opening a big .xlsb is slow (shared strings), so open it once for every sheet
        from pyxlsb import open_workbook
        sheets, n = [], hrow - 1
        with open_workbook(f["path"]) as wb:
            for s in wb.sheets:
                hdrs = []
                with wb.get_sheet(s) as sh:
                    for row in sh.rows():
                        if not row or row[0].r < n:
                            continue
                        if row[0].r == n:
                            hdrs = [{"idx": c.c, "label": str(c.v).strip()} for c in row
                                    if c.v is not None and str(c.v).strip()]
                        break
                sheets.append({"name": s, "headers": hdrs})
    else:
        sheets = [{"name": s, "headers": headers_for(f, s, hrow)} for s in sheet_names(f)]
    return {"id": fid, "name": f["name"], "size": os.path.getsize(f["path"]), "sheets": sheets}


def add_file(fid, name, path, hrow):
    k = kind_of(name)
    if not k:
        raise ValueError("Unsupported file type. Use .xlsx .xlsb .xlsm .xls .csv .tsv")
    FILES[fid] = {"name": name, "path": path, "kind": k}
    return json.dumps(file_payload(fid, int(hrow)))


def headers(ids_json, hrow):
    return json.dumps([file_payload(i, int(hrow)) for i in json.loads(ids_json)])


def remove_file(fid):
    FILES.pop(fid, None)


def clean_name(used, base):
    n = (re.sub(r"[\\/?*\[\]:]", "_", base).strip()[:31]) or "Sheet"
    k = 1
    while n.lower() in used:
        k += 1
        suf = f"_{k}"
        n = n[: 31 - len(suf)] + suf
    used.add(n.lower())
    return n


def num_str(v):
    return re.fullmatch(r"-?[1-9]\d{0,14}(\.\d+)?|0(\.\d+)?", v) is not None


def is_date_label(lbl):
    l = lbl.lower().strip()
    return "date" in l or l == "created on"


def export(req_json, progress):
    """Build the workbook at /tmp/out.xlsx. Returns JSON {path, name, truncated}."""
    req = json.loads(req_json)
    hrow = max(1, int(req.get("hrow", 1)))
    merge = bool(req.get("merge"))
    out_name = re.sub(r'[\\/:*?"<>|]', "_", req.get("name") or "Custom_Output") + ".xlsx"
    dest = "/tmp/out.xlsx"
    wb = xlsxwriter.Workbook(dest, {"constant_memory": True, "nan_inf_to_errors": True})
    hdr = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
    datef = wb.add_format({"num_format": "dd-mmm-yyyy hh:mm"})
    longi = wb.add_format({"num_format": "0"})
    used, truncated = set(), False

    jobs = []
    for fr in req["files"]:
        f = FILES[fr["id"]]
        for sr in fr["sheets"]:
            if sr["cols"]:
                jobs.append((f, sr["name"], sr["cols"]))
    multi = len({j[0]["name"] for j in jobs}) > 1

    def write_cell(ws, r, c, v, date_col):
        if isinstance(v, bool):
            ws.write_boolean(r, c, v)
        elif isinstance(v, (datetime.datetime, datetime.date)):
            ws.write_datetime(r, c, v, datef)
        elif isinstance(v, (int, float)):
            if date_col and v >= DATE_SERIAL_MIN:
                ws.write_number(r, c, v, datef)
            elif isinstance(v, float) and v.is_integer() and abs(v) >= 1e9:
                ws.write_number(r, c, v, longi)
            else:
                ws.write_number(r, c, v)
        else:
            s = str(v)
            if num_str(s):
                x = float(s)
                ws.write_number(r, c, x, longi if x.is_integer() else None)
            else:
                ws.write_string(r, c, s)

    if merge:
        labels = []
        for _, _, cols in jobs:
            for c in cols:
                if c["label"] not in labels:
                    labels.append(c["label"])
        ws = wb.add_worksheet("Merged")
        ws.freeze_panes(1, 0)
        for n, h in enumerate(["Source file", "Source sheet"] + labels):
            ws.write(0, n, h, hdr)
            ws.set_column(n, n, max(12, min(len(h) + 4, 40)))
        out = 1
        for f, sheet, cols in jobs:
            pos = {c["idx"]: labels.index(c["label"]) + 2 for c in cols}
            dcols = {pos[c["idx"]] for c in cols if is_date_label(c["label"])}
            for r, d in iter_rows(f, sheet, pos.keys(), hrow):
                if r == hrow - 1 or not d:
                    continue
                if out > MAX_ROWS:
                    truncated = True
                    break
                ws.write_string(out, 0, f["name"])
                ws.write_string(out, 1, sheet.strip())
                for i, v in d.items():
                    write_cell(ws, out, pos[i], v, pos[i] in dcols)
                out += 1
                if out % 20000 == 0:
                    progress(f"{f['name']} / {sheet.strip()}: {out:,} rows written...")
        ws.autofilter(0, 0, max(out - 1, 1), len(labels) + 1)
    else:
        for f, sheet, cols in jobs:
            base = (os.path.splitext(f["name"])[0][:14] + "_" if multi else "") + sheet.strip()
            ws = wb.add_worksheet(clean_name(used, base))
            ws.freeze_panes(1, 0)
            order = {c["idx"]: n for n, c in enumerate(cols)}
            dcols = {n for n, c in enumerate(cols) if is_date_label(c["label"])}
            for n, c in enumerate(cols):
                ws.write(0, n, c["label"], hdr)
                ws.set_column(n, n, max(12, min(len(c["label"]) + 4, 40)))
            out = 1
            for r, d in iter_rows(f, sheet, order.keys(), hrow):
                if r == hrow - 1 or not d:
                    continue
                if out > MAX_ROWS:
                    truncated = True
                    break
                for i, v in d.items():
                    write_cell(ws, out, order[i], v, order[i] in dcols)
                out += 1
                if out % 20000 == 0:
                    progress(f"{f['name']} / {sheet.strip()}: {out:,} rows written...")
            ws.autofilter(0, 0, max(out - 1, 1), len(cols) - 1)
            progress(f"{sheet.strip()}: {out - 1:,} rows done")
    progress("Saving file...")
    wb.close()
    return json.dumps({"path": dest, "name": out_name, "truncated": truncated})
