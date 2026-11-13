"""
NPD Sourcing Tracker — Streamlit application
E&E BOP alternate development, parts MPN/price coverage, supplier tracking and cost-saving comparisons.

Converted from the HTML/JavaScript tracker (index.html + app.js). All data stays with the user:
upload the master Excel or a JSON backup, edit in the app, then download the updated Excel/backup.
"""
from __future__ import annotations

import copy
import hashlib
import io
import json
import os
import re
from datetime import date, datetime, timedelta

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

# ----------------------------------------------------------------------------------------------
# Constants
# ----------------------------------------------------------------------------------------------
APP_TITLE = "NPD Sourcing Tracker"
SHEETS = {
    "mqis": "MQIS AND SINGLE PHASE PART",
    "ds6": "6DS PROJECT PARTS",
    "projects": "ALTERNATE DEVELOPNMENT",
    "fan": "DC Fan Comparison",
    "teknic": "Teknic Parts (Domestic)",
}
SHEET_INDEX = {"mqis": 0, "ds6": 1, "projects": 2, "fan": 3, "teknic": 4}
MQIS_HEADERS = ["Sl.No", "Part No", "Part Name", "BOM Qty/Set", "UOM", "Buyer", "Supplier", "Project",
                "MAKE/ MPN-1", "MAKE/ MPN -2", "MAKE/MPN-3", "Price"]
DS6_HEADERS = ["S. No.", "Part No.", "Part Name", "Rev", "Qty", "UOM", "Qyt Required for 50  Sets", "Vendor",
               "MAKE/MPN", "Unit Cost"]
# per parts-sheet config (indexes into the row list) — same as app.js PCFG
PCFG = {
    "mqis": {"num": [3, 11], "mpn": [8, 9, 10], "price": 11, "supplier": 6, "qty": 3, "proj": 7, "ncol": 12,
             "label": "MQIS & Single Phase Parts"},
    "ds6": {"num": [4, 6, 9], "mpn": [8], "price": 9, "supplier": 7, "qty": 4, "proj": None, "ncol": 10,
            "label": "6DS Project Parts"},
}
STATUSES = ["Ongoing", "Completed", "On Hold", "Dropped"]
PROJECT_CHOICES = ["Common", "P1R5", "Single Phs"]
DEFAULT_STAGE_GATE = [
    "Alternate change (trigger / start)", "Receipt of Latest SOR/Drawing from Engineering",
    "Share SOR & Float RFQ to Shortlisted Suppliers", "Receipt of Quotation & Technical Compliance",
    "Technical Evaluation & Shortlisting (Engineering)", "Commercial Negotiation & Finalisation",
    "Sample PO Release", "Sample Receipt", "Sample Inspection (Quality)", "Sample Validation (Engineering)",
    "Field Trial", "Field Trial Report & Approval", "PPAP / Documentation", "Vendor Code Creation",
    "Part Code Creation / BOM Update", "Pilot Lot PO", "Pilot Lot Receipt & Inspection",
    "Production Approval (Engg + Quality)", "Bulk PO Release", "Closure & Handover to Regular Buyer",
]
COLORS = {
    "navy": "#0B5C8A", "ink": "#18222F", "muted": "#5D6B7C", "line": "#D8DFE8", "bg": "#F3F5F8",
    "ok": "#1D8A4E", "ok_soft": "#DCF2E5", "warn": "#B86E00", "warn_soft": "#FDECCC",
    "bad": "#C2372D", "bad_soft": "#FBE1DE", "idle": "#6B7686", "idle_soft": "#E7EAEF",
    "Common": "#6A4FC7", "P1R5": "#0D8C8C", "Single": "#D0611B", "6DS": "#2F6FD6",
}
STATUS_COLORS = {"Completed": COLORS["ok"], "Ongoing": COLORS["warn"], "On Hold": COLORS["idle"],
                 "Dropped": COLORS["idle"], "Not set": COLORS["idle"]}
ACT_COLORS = {"Done": COLORS["ok"], "Delayed": COLORS["bad"], "Due in 7 days": COLORS["warn"],
              "Open": COLORS["navy"], "Not planned": COLORS["idle"]}
AUTOSAVE_FILE = os.environ.get("NPD_AUTOSAVE_FILE", "")  # optional: set for local single-user use
# Default data shipped with the (PRIVATE) repository. Loaded automatically when the app starts.
DEFAULT_DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "npd_data.json")


# ----------------------------------------------------------------------------------------------
# Generic helpers
# ----------------------------------------------------------------------------------------------
def today() -> date:
    return date.today()


def blank(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


def to_date(v):
    """Accept date, datetime, ISO string → date or None."""
    if blank(v):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.date()
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def iso(v):
    d = to_date(v)
    return d.isoformat() if d else None


def num(v):
    if blank(v):
        return None
    if isinstance(v, (int, float)):
        return float(v) if not isinstance(v, bool) else None
    try:
        return float(str(v).replace("₹", "").replace(",", "").strip())
    except ValueError:
        return None


def jsonable(v):
    if isinstance(v, (datetime, date)):
        return v.isoformat()[:10] if isinstance(v, date) and not isinstance(v, datetime) else v.isoformat()[:10]
    if isinstance(v, float) and pd.isna(v):
        return None
    if hasattr(v, "item"):  # numpy scalars
        try:
            return v.item()
        except Exception:
            return str(v)
    return v


def part_no(v):
    """Keep numeric part numbers as int, everything else as trimmed text."""
    if blank(v):
        return None
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return int(s) if s.isdigit() else s


def inr(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = f"{abs(v):,.2f}"
    # Indian grouping
    whole, frac = s.split(".")
    whole = whole.replace(",", "")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        head = re.sub(r"(\d)(?=(\d{2})+$)", r"\1,", head)
        whole = head + "," + tail
    return ("-" if v < 0 else "") + "₹ " + whole + "." + frac


def proj_key(p) -> str:
    p = str(p or "")
    if re.search("single", p, re.I):
        return "Single"
    if re.search("p1r5", p, re.I):
        return "P1R5"
    if re.search("6ds", p, re.I):
        return "6DS"
    return "Common"


def fmt_d(d) -> str:
    d = to_date(d)
    return d.strftime("%d %b %y") if d else ""


# ----------------------------------------------------------------------------------------------
# Data model (same structure as app.js / JSON backup, so backups are interchangeable)
# ----------------------------------------------------------------------------------------------
def empty_data() -> dict:
    return {"mqis": [MQIS_HEADERS[:], []], "ds6": [DS6_HEADERS[:], []], "projects": [], "fan": [], "teknic": [],
            "tk": {"title": "Teknic Parts (Domestic)", "parts": [], "stations": None}, "fanExtra": [],
            "stdActs": DEFAULT_STAGE_GATE[:]}


def teknic_from_raw(raw) -> dict:
    rows = [r for r in raw if r and any(not blank(v) for v in r)]
    title = rows[0][0] if rows else "Teknic Parts (Domestic)"
    parts = []
    for r in rows:
        r = list(r) + [None] * 12
        if isinstance(r[0], (int, float)) and not isinstance(r[0], bool) and not blank(r[1]):
            parts.append({"pn": part_no(r[0]), "name": r[1], "make": r[2], "mpn": r[3], "bom": num(r[4]),
                          "exs": r[5], "alts": r[6], "exp": num(r[7]), "qp": num(r[8])})
    stations = None
    for r in rows:
        if re.search("Stations per year", str(r[0] or ""), re.I):
            stations = num(r[1]) if len(r) > 1 else None
    return {"title": title or "Teknic Parts (Domestic)", "parts": parts, "stations": stations}


def norm(d: dict) -> dict:
    d = copy.deepcopy(d)
    if not isinstance(d.get("tk"), dict):
        d["tk"] = teknic_from_raw(d.get("teknic") or [])
    base = empty_data()
    for k, v in base.items():
        d.setdefault(k, v)
    if not d.get("tk") or not isinstance(d["tk"], dict):
        d["tk"] = teknic_from_raw(d.get("teknic") or [])
    d["tk"].setdefault("parts", [])
    d["tk"].setdefault("stations", None)
    d["tk"].setdefault("title", "Teknic Parts (Domestic)")
    if not d.get("stdActs"):
        d["stdActs"] = DEFAULT_STAGE_GATE[:]
    for key, hdr in (("mqis", MQIS_HEADERS), ("ds6", DS6_HEADERS)):
        if not d.get(key) or len(d[key]) != 2:
            d[key] = [hdr[:], []]
        ncol = PCFG[key]["ncol"]
        d[key][0] = (list(d[key][0]) + hdr[len(d[key][0]):])[:ncol]
        d[key][1] = [_norm_row(r, ncol) for r in d[key][1]]
    for p in d["projects"]:
        p.setdefault("status", "")
        p["acts"] = [_norm_row(a, 5) for a in p.get("acts", [])]
    d.setdefault("meta", {})
    return d


def _norm_row(r, n):
    """Core values in fixed positions + optional dict of extra (user-added) column values at position n."""
    r = list(r)
    ex = r[n] if len(r) > n and isinstance(r[n], dict) else None
    core = (r[:n] + [None] * n)[:n]
    return core + [ex] if ex else core


def std_acts(d) -> list:
    dls = next((p for p in d["projects"] if p["name"] == "DLS"), None)
    if dls is None and d["projects"]:
        dls = max(d["projects"], key=lambda p: len(p["acts"]))
    if dls and dls["acts"]:
        return [a[1] for a in dls["acts"]]
    return d.get("stdActs") or DEFAULT_STAGE_GATE[:]


# ----------------------------------------------------------------------------------------------
# Calculations (identical rules to app.js)
# ----------------------------------------------------------------------------------------------
def is_trigger(name) -> bool:
    return bool(re.search("trigger", str(name or ""), re.I))


def activity_status(a) -> str:
    step, name, plan, actual = a[0], a[1], to_date(a[2]), to_date(a[3])
    if actual:
        return "Done"
    if not isinstance(step, (int, float)) or is_trigger(name) or not plan:
        return "Not planned" if not plan else "Open"
    if plan < today():
        return "Delayed"
    if (plan - today()).days <= 7:
        return "Due in 7 days"
    return "Open"


def proj_stats(p) -> dict:
    total = done = soon = 0
    late = []
    for a in p["acts"]:
        if not isinstance(a[0], (int, float)) or is_trigger(a[1]):
            continue
        pl, ac = to_date(a[2]), to_date(a[3])
        if not pl and not ac:
            continue
        total += 1
        if ac:
            done += 1
        elif pl and pl < today():
            late.append({"act": a[1], "plan": pl, "days": (today() - pl).days, "rem": a[4]})
        elif pl and (pl - today()).days <= 7:
            soon += 1
    return {"total": total, "done": done, "late": late, "soon": soon, "pct": done / total if total else 0.0}


def has_mpn(key, r) -> bool:
    return any(not blank(r[i]) for i in PCFG[key]["mpn"])


def has_price(key, r) -> bool:
    v = r[PCFG[key]["price"]]
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not pd.isna(v)


def fan_calc(f: dict) -> dict:
    g = lambda k: num(f.get(k)) or 0.0
    p_inr = g("pUSD") * g("fx")
    fr = g("freight")
    ins = (p_inr + fr) * g("insPct") / 100
    cif = p_inr + fr + ins
    bcd = cif * g("bcdPct") / 100
    sws = bcd * g("swsPct") / 100
    landed = cif + bcd + sws + g("cha")
    alt = g("altPrice") + g("altFreight")
    save = landed - alt
    return {"pINR": p_inr, "ins": ins, "cif": cif, "bcd": bcd, "sws": sws, "landed": landed, "alt": alt,
            "save": save, "pct": save / landed if landed else 0.0, "annual": save * g("vol")}


def teknic_calc(tk: dict) -> dict:
    rows, te, tq, ts = [], 0.0, 0.0, 0.0
    for p in tk["parts"]:
        exp, qp, bom = num(p.get("exp")) or 0, num(p.get("qp")) or 0, num(p.get("bom")) or 0
        s = exp - qp
        te += exp * bom
        tq += qp * bom
        ts += s * bom
        rows.append({**p, "save": s, "pct": s / exp if exp else 0.0, "save_stn": s * bom})
    st_ = num(tk.get("stations"))
    return {"rows": rows, "te": te, "tq": tq, "ts": ts, "pct": (te - tq) / te if te else 0.0,
            "annual": ts * st_ if st_ is not None else None}


# ----------------------------------------------------------------------------------------------
# Excel import (mirrors the "Load Excel" logic of app.js)
# ----------------------------------------------------------------------------------------------
def _sheet(wb, key):
    name = SHEETS[key]
    if name in wb.sheetnames:
        return wb[name]
    idx = SHEET_INDEX[key]
    return wb.worksheets[idx] if idx < len(wb.worksheets) else None


def _grid(ws):
    return [[c for c in row] for row in ws.iter_rows(values_only=True)]


def _slug(label, taken):
    base = "x_" + (re.sub(r"[^a-z0-9]+", "_", str(label).lower()).strip("_") or "col")
    sid, n = base, 2
    while sid in taken:
        sid, n = f"{base}_{n}", n + 1
    return sid


def _extra_meta(labels, values_by_col):
    """Build meta for user-added columns found in Excel: {id: type}, {id: label}, [ids]."""
    custom, lab, ids = {}, {}, []
    for label, vals in zip(labels, values_by_col):
        cid = _slug(label, set(ids))
        nonblank = [v for v in vals if not blank(v)]
        typ = "number" if nonblank and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in nonblank) else (
            "date" if nonblank and all(isinstance(v, (datetime, date)) for v in nonblank) else "text")
        custom[cid], lab[cid] = typ, label
        ids.append(cid)
    return custom, lab, ids


SYNONYMS = {
    "mqis": {"slno": ["slno", "sno", "srno", "serialno"], "partno": ["partno", "partnumber", "itemcode"],
             "partname": ["partname", "description", "partdescription"], "qty": ["bomqtyset", "bomqty", "qty", "qtyset"],
             "uom": ["uom", "unit"], "buyer": ["buyer"], "supplier": ["supplier", "vendor"], "project": ["project"],
             "mpn1": ["makempn1", "makempn"], "mpn2": ["makempn2"], "mpn3": ["makempn3"],
             "price": ["price", "unitprice", "unitcost", "cost"]},
    "ds6": {"slno": ["sno", "slno", "srno"], "partno": ["partno", "partnumber", "itemcode"], "partname": ["partname", "description"],
            "rev": ["rev", "revision"], "qty": ["qty", "qtyset"], "uom": ["uom", "unit"],
            "qty50": ["qytrequiredfor50sets", "qtyrequiredfor50sets", "qtyfor50sets", "qty50sets"],
            "vendor": ["vendor", "supplier"], "mpn": ["makempn", "mpn"], "price": ["unitcost", "price", "unitprice", "cost"]},
}


def _nk(v):
    return re.sub(r"[^a-z0-9]", "", str(v or "").lower())


def _map_header(labels, t, cur_labels=None):
    """Match sheet header labels to built-in column ids by name (so added, removed or moved Excel columns still map).
    Returns ({cid: index in labels}, [indexes not matched])."""
    syn = {cid: set(v) for cid, v in SYNONYMS.get(t, {}).items()}
    for cid, lab in (cur_labels or {}).items():
        if cid in syn and not blank(lab):
            syn[cid].add(_nk(lab))
    found, rest = {}, []
    for i, lab in enumerate(labels):
        k = _nk(lab)
        cid = next((c for c in TABLE_DEFS[t]["core"] if c not in found and k in syn.get(c, ())), None)
        if cid:
            found[cid] = i
        elif not blank(lab):
            rest.append(i)
    return found, rest


def _parts(ws, hdr_re, ncol, meta_out=None, key=None, cur_meta=None):
    g = _grid(ws)
    hi = next((i for i, r in enumerate(g) if any(re.fullmatch(hdr_re, str(v or "").strip(), re.I) for v in r)), -1)
    if hi < 0:
        raise ValueError(f'Could not find the header row in sheet "{ws.title}".')
    c0 = next(i for i, v in enumerate(g[hi]) if not blank(v))
    labels = [str(v or "").replace("\xa0", " ").strip() for v in g[hi][c0:]]
    while labels and blank(labels[-1]):
        labels.pop()
    found, rest = _map_header(labels, key, (cur_meta or {}).get("labels"))
    core = TABLE_DEFS[key]["core"]
    defaults = MQIS_HEADERS if key == "mqis" else DS6_HEADERS
    hdr = [labels[found[c]] if c in found else defaults[i] for i, c in enumerate(core)]
    xlabels = [labels[i] for i in rest]
    body = [list(r[c0:]) + [None] * (len(labels) + 1) for r in g[hi + 1:]]
    custom, lab, ids = _extra_meta(xlabels, [[r[i] for r in body] for i in rest])
    rows = []
    for r in body:
        vals = [jsonable(r[found[c]]) if c in found else None for c in core]
        ex = {cid: jsonable(r[i]) for cid, i in zip(ids, rest) if not blank(r[i])}
        if any(not blank(v) for v in vals) or ex:
            rows.append(vals + ([ex] if ex else []))
    if meta_out is not None and key:
        meta_out[key] = {"custom": custom, "labels": {**{c: hdr[i] for i, c in enumerate(core)}, **lab},
                         "hidden": [c for c in core if c not in found]}
    return [hdr, rows]


def parse_workbook(file_bytes: bytes, current: dict | None = None) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    nd = {"fanExtra": (current or {}).get("fanExtra", []), "stdActs": (current or {}).get("stdActs", DEFAULT_STAGE_GATE[:]),
          "meta": {}}
    old_meta = (current or {}).get("meta", {})
    ws = _sheet(wb, "mqis")
    if ws is None:
        raise ValueError("The MQIS / Single Phase parts sheet was not found.")
    nd["mqis"] = _parts(ws, r"Part No", 12, nd["meta"], "mqis", old_meta.get("mqis"))
    ws = _sheet(wb, "ds6")
    if ws is None:
        raise ValueError("The 6DS parts sheet was not found.")
    nd["ds6"] = _parts(ws, r"Part No\.?", 10, nd["meta"], "ds6", old_meta.get("ds6"))
    ws = _sheet(wb, "projects")
    if ws is None:
        raise ValueError("The ALTERNATE DEVELOPNMENT sheet was not found.")
    g = _grid(ws)
    si = next((i for i, r in enumerate(g) if r and str(r[0] or "").strip() == "Stage"), -1)
    if si < 0:
        raise ValueError('Could not find the "Stage" header in the Alternate Development sheet.')
    head = list(g[si]) + [None]
    xcols = [j for j in range(7, len(g[si])) if not blank(g[si][j])]  # columns after the hidden key column G
    xlabels = [str(g[si][j]).strip() for j in xcols]
    body = [list(r) + [None] * (len(head) + 6) for r in g[si + 1:]]
    custom, lab, ids = _extra_meta(xlabels, [[r[j] for r in body if not blank(r[1])] for j in xcols])
    nd["meta"]["acts"] = {"custom": custom, "labels": lab}
    projects, cur = [], None
    for r in body:
        a, b = r[0], r[1]
        if isinstance(a, str) and a.strip() and blank(b):
            cur = {"name": a.strip(), "status": str(r[4] or "").strip(), "acts": []}
            projects.append(cur)
        elif cur is not None and not blank(b):
            ex = {cid: (iso(r[j]) if custom[cid] == "date" else jsonable(r[j])) for cid, j in zip(ids, xcols) if not blank(r[j])}
            cur["acts"].append([jsonable(a), b, iso(r[2]) or jsonable(r[2]), iso(r[3]) or jsonable(r[3]), jsonable(r[5])]
                               + ([ex] if ex else []))
    nd["projects"] = projects
    if "teknic" in old_meta:
        nd["meta"]["teknic"] = old_meta["teknic"]
    ws = _sheet(wb, "fan")
    nd["fan"] = [[jsonable(v) for v in (list(r) + [None] * 5)[1:5]] for r in _grid(ws)] if ws is not None else (current or {}).get("fan", [])
    ws = _sheet(wb, "teknic")
    if ws is not None:
        nd["teknic"] = [[jsonable(v) for v in (list(r) + [None] * 13)[1:13]] for r in _grid(ws)]
        nd["tk"] = teknic_from_raw(nd["teknic"])
    else:
        nd["teknic"] = (current or {}).get("teknic", [])
        nd["tk"] = (current or {}).get("tk") or teknic_from_raw([])
    return norm(nd)


def parse_backup(file_bytes: bytes) -> dict:
    o = json.loads(file_bytes.decode("utf-8"))
    d = o.get("data", o) if isinstance(o, dict) else None
    if not isinstance(d, dict) or "projects" not in d:
        raise ValueError("This is not an NPD tracker backup file.")
    return norm(d)


# ----------------------------------------------------------------------------------------------
# Excel export 1: clean data workbook (same sheets as app.js "Export Excel")
# ----------------------------------------------------------------------------------------------
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Calibri", size=10)
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _write_table(ws, aoa, widths, money_cols=(), pct_cols=(), date_cols=()):
    for r, row in enumerate(aoa, 1):
        for c, v in enumerate(row, 1):
            if isinstance(v, str) and c in date_cols and r > 1 and to_date(v):
                v = datetime.combine(to_date(v), datetime.min.time())
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if r == 1:
                cell.font, cell.fill = HDR_FONT, HDR_FILL
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.font = BODY_FONT
                cell.alignment = Alignment(vertical="center", wrap_text=isinstance(v, str) and len(v) > 30)
                if c in money_cols:
                    cell.number_format = '"₹ "#,##0.00'
                elif c in pct_cols:
                    cell.number_format = "0.00%"
                elif c in date_cols:
                    cell.number_format = "dd-mmm-yy"
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if len(aoa) > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(aoa[0]))}{len(aoa)}"
    ws.row_dimensions[1].height = 30


def clean_export(d: dict) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    stats = [proj_stats(p) for p in d["projects"]]
    _write_table(wb.create_sheet("Project Status"),
                 [["#", "Project Name", "Status", "Activities Done", "Total", "Progress %", "Delayed"]] +
                 [[i + 1, p["name"], p["status"], s["done"], s["total"], round(s["pct"] * 100), len(s["late"])]
                  for i, (p, s) in enumerate(zip(d["projects"], stats))], [5, 32, 12, 14, 8, 11, 9])
    for key in ("mqis", "ds6"):
        m = tmeta(d, key)
        recs = [_row_to_rec(r, TABLE_DEFS[key]["core"]) for r in d[key][1]]
        cols = vis(m)
        aoa = [[m["labels"][c] for c in cols]] + [[recs_v(r, c) for c in cols] for r in recs]
        _write_table(wb.create_sheet(SHEETS[key][:31]), aoa, [max(8, min(40, len(str(h)) + 6)) for h in aoa[0]],
                     money_cols=(cols.index("price") + 1,) if "price" in cols else ())
    m = tmeta(d, "acts")
    cols = vis(m)
    det = [["Project", "Project Status"] + [m["labels"][c] for c in cols]]
    for p in d["projects"]:
        for a in p["acts"]:
            rec = _row_to_rec(a, TABLE_DEFS["acts"]["core"])
            det.append([p["name"], p["status"]] + [recs_v(rec, c) for c in cols])
    dcols = tuple(i + 3 for i, c in enumerate(cols) if _ctype(m, "acts", c) == "date")
    _write_table(wb.create_sheet(SHEETS["projects"][:31]), det, [26, 12] + [max(8, min(60, len(str(h)) + 8)) for h in det[0][2:]],
                 date_cols=dcols)
    ws = wb.create_sheet(SHEETS["fan"])
    for row in d["fan"]:
        ws.append(row)
    for i, w in enumerate([38, 26, 28, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for i, f in enumerate(d["fanExtra"]):
        _fan_extra_sheet(wb, f, i + 2)
    t = d["tk"]
    c = teknic_calc(t)
    tm = tmeta(d, "teknic")
    xc = [cid for cid in tm["order"] if cid in tm["custom"]]
    tk = [["Part No.", "Part Name", "Make", "MPN", "BOM / Station", "Existing Supplier", "Alternate Supplier",
           "Existing Unit Price (₹)", "Quoted Unit Price (₹)", "Saving / Unit (₹)", "Saving %", "Saving / Station (₹)"]
          + [tm["labels"][cid] for cid in xc]]
    for p in c["rows"]:
        tk.append([p["pn"], p["name"], p["make"], p["mpn"], p["bom"], p["exs"], p["alts"], p["exp"], p["qp"],
                   p["save"], p["pct"], p["save_stn"]] + [p.get(cid) for cid in xc])
    tk.append(["TOTAL PER STATION", None, None, None, None, None, None, c["te"], c["tq"], None, c["pct"], c["ts"]])
    ws = wb.create_sheet(SHEETS["teknic"])
    _write_table(ws, tk, [14, 28, 10, 30, 8, 16, 14, 14, 14, 12, 9, 14] + [16] * len(xc), money_cols=(8, 9, 10, 12), pct_cols=(11,))
    ws.append([])
    ws.append(["Stations per year", t.get("stations")])
    ws.append(["Annual saving (₹)", c["annual"]])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _fan_extra_sheet(wb, f: dict, n: int):
    """Comparison sheet with live Excel formulas (inputs in column B/C)."""
    name = re.sub(r"[\\/?*\[\]:]", " ", f"DC Fan {n} {f.get('part', '')}")[:31]
    ws = wb.create_sheet(name)
    rows = [
        [f.get("part"), None, None, None],
        ["Cost element", "Existing supplier (import)", "Alternate supplier (domestic)", "Basis / note"],
        ["Supplier", f.get("exName"), f.get("altName"), None],
        ["Unit price (US$)", num(f.get("pUSD")), None, "Input"],
        ["Exchange rate (₹ per US$)", num(f.get("fx")), None, "Input"],
        ["Unit price (₹)", "=B4*B5", num(f.get("altPrice")), "Existing = US$ × rate"],
        ["Freight (₹)", num(f.get("freight")) or 0, num(f.get("altFreight")) or 0, "Input"],
        ["Insurance %", (num(f.get("insPct")) or 0) / 100, None, "% of price + freight"],
        ["Insurance (₹)", "=(B6+B7)*B8", None, None],
        ["CIF value (₹)", "=B6+B7+B9", None, "Price + freight + insurance"],
        ["BCD %", (num(f.get("bcdPct")) or 0) / 100, None, "Basic Customs Duty"],
        ["BCD (₹)", "=B10*B11", None, None],
        ["SWS % (on BCD)", (num(f.get("swsPct")) or 0) / 100, None, "Social Welfare Surcharge"],
        ["SWS (₹)", "=B12*B13", None, None],
        ["CHA / clearance / local freight (₹)", num(f.get("cha")) or 0, None, None],
        ["LANDED COST PER UNIT (₹)", "=B10+B12+B14+B15", "=C6+C7", "IGST excluded on both sides"],
        ["Saving per unit with alternate (₹)", "=B16-C16", None, None],
        ["Saving %", "=IFERROR(B17/B16,0)", None, None],
        ["Annual volume (nos)", num(f.get("vol")) or 0, None, None],
        ["Annual saving (₹)", "=B17*B19", None, None],
        ["Make / MPN", f.get("exMpn"), f.get("altMpn"), None],
        ["Lead time (weeks)", f.get("exLt"), f.get("altLt"), None],
        ["MOQ", f.get("exMoq"), f.get("altMoq"), None],
        ["Warranty", f.get("exWar"), f.get("altWar"), None],
        ["Payment terms", f.get("exPay"), f.get("altPay"), None],
        ["Note", f.get("note"), None, None],
    ]
    for r in rows:
        ws.append(r)
    ws["A1"].font = Font(bold=True, size=12, color="1F4E78")
    for c in ws[2]:
        c.font, c.fill = HDR_FONT, HDR_FILL
    for r in range(3, ws.max_row + 1):
        for c in ws[r]:
            c.border = BORDER
            c.font = Font(name="Calibri", size=10, bold=r in (16, 20))
    for ref in ["B6", "C6", "B7", "C7", "B9", "B10", "B12", "B14", "B15", "B16", "C16", "B17", "B20"]:
        ws[ref].number_format = '"₹ "#,##0.00'
    for ref in ["B8", "B11", "B13", "B18"]:
        ws[ref].number_format = "0.00%"
    for i, w in enumerate([36, 26, 26, 34], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ----------------------------------------------------------------------------------------------
# Excel export 2: write changes back INTO the uploaded master workbook (keeps formatting/formulas)
# ----------------------------------------------------------------------------------------------
def _style_of(cell):
    return {"font": copy.copy(cell.font), "fill": copy.copy(cell.fill), "border": copy.copy(cell.border),
            "alignment": copy.copy(cell.alignment), "number_format": cell.number_format,
            "protection": copy.copy(cell.protection)}


def _apply_style(cell, s):
    if not s:
        return
    cell.font, cell.fill, cell.border = copy.copy(s["font"]), copy.copy(s["fill"]), copy.copy(s["border"])
    cell.alignment, cell.number_format = copy.copy(s["alignment"]), s["number_format"]
    cell.protection = copy.copy(s["protection"])


def _row_styles(ws, r, c1, c2):
    return {c: _style_of(ws.cell(r, c)) for c in range(c1, c2 + 1)}


def _unmerge_from(ws, row):
    for m in list(ws.merged_cells.ranges):
        if m.max_row >= row:
            ws.unmerge_cells(str(m))


def _find_header(ws, pattern):
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and re.fullmatch(pattern, str(c.value).strip(), re.I):
                return c.row, c.column
    return None, None


def _write_parts_sheet(ws, hdr_pat, t, d):
    """Write the parts list back into the user's sheet, matching columns by header name (built-in + added columns)."""
    hr, _ = _find_header(ws, hdr_pat)
    if hr is None:
        return
    m = tmeta(d, t)
    core = TABLE_DEFS[t]["core"]
    c0 = next(c for c in range(1, ws.max_column + 1) if ws.cell(hr, c).value not in (None, ""))
    last_hdr = max(c for c in range(c0, ws.max_column + 1) if ws.cell(hr, c).value not in (None, ""))
    labels = [ws.cell(hr, c).value for c in range(c0, last_hdr + 1)]
    found, rest = _map_header(labels, t, m["labels"])
    colof = {cid: c0 + i for cid, i in found.items()}
    free = {_nk(labels[i]): c0 + i for i in rest}
    nxt = last_hdr + 1
    head_style = _style_of(ws.cell(hr, last_hdr))
    for cid in [c for c in m["order"] if c in m["custom"]]:
        k = _nk(m["labels"][cid])
        if k in free:
            colof[cid] = free.pop(k)
        else:
            colof[cid] = nxt
            nxt += 1
    template = _row_styles(ws, hr + 1, c0, max(last_hdr, nxt))
    body_style = template.get(c0 + 2)
    height = ws.row_dimensions[hr + 1].height
    cols = sorted(set(colof.values()))
    old_last = hr
    for r in range(hr + 1, ws.max_row + 1):
        if any(ws.cell(r, c).value not in (None, "") for c in cols):
            old_last = r
    _unmerge_from(ws, hr + 1)
    for cid, c in colof.items():  # headers: renamed built-ins + added columns
        cell = ws.cell(hr, c)
        if c > last_hdr:
            _apply_style(cell, head_style)
            ws.column_dimensions[get_column_letter(c)].width = max(12, min(40, len(str(m["labels"][cid])) + 6))
        cell.value = m["labels"][cid]
    rows = d[t][1]
    for r in range(hr + 1, max(old_last, hr + len(rows)) + 1):
        for c in cols:
            ws.cell(r, c).value = None
    for i, row in enumerate(rows):
        rec = _row_to_rec(row, core)
        r = hr + 1 + i
        for cid, c in colof.items():
            cell = ws.cell(r, c)
            _apply_style(cell, template.get(c) if c <= last_hdr else body_style)
            v = rec.get(cid)
            cell.value = datetime.combine(to_date(v), datetime.min.time()) if isinstance(v, str) and re.match(r"\d{4}-\d{2}-\d{2}$", v) else jsonable(v)
        if height:
            ws.row_dimensions[r].height = height
    lastc = max(cols) if cols else last_hdr
    ws.auto_filter.ref = f"{get_column_letter(c0)}{hr}:{get_column_letter(lastc)}{hr + max(len(rows), 1)}"


def _remap(ref, old_lo, old_hi, new_lo, new_hi):
    out = []
    for part in str(ref).split():
        cr = CellRange(part)
        if cr.min_row >= old_lo and cr.max_row <= old_hi + 1:
            out.append(f"{get_column_letter(cr.min_col)}{new_lo}:{get_column_letter(cr.max_col)}{new_hi}")
        else:
            out.append(part)
    return " ".join(out)


def _write_projects_sheet(ws, projects, extras=()):
    # locate structure of the existing sheet
    stage_row = next((r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip() == "Stage"), None)
    if stage_row is None:
        return
    top_hdr = next((r for r in range(1, stage_row) if str(ws.cell(r, 2).value or "").strip() == "Project Name"), None)
    first_proj = next((r for r in range(stage_row + 1, ws.max_row + 1)
                       if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip()
                       and ws.cell(r, 2).value in (None, "")), stage_row + 1)
    act_row = first_proj + 1
    blank_row = next((r for r in range(act_row, ws.max_row + 1)
                      if all(ws.cell(r, c).value in (None, "") for c in range(1, 7))), act_row + 1)
    ncol = 7
    st_top = _row_styles(ws, top_hdr + 1, 1, 4) if top_hdr else None
    st_title = _style_of(ws.cell(stage_row - 1, 1))
    title_text = ws.cell(stage_row - 1, 1).value
    st_head = _row_styles(ws, stage_row, 1, ncol)
    head_vals = [ws.cell(stage_row, c).value for c in range(1, ncol + 1)]
    head_h = ws.row_dimensions[stage_row].height
    st_proj = _row_styles(ws, first_proj, 1, ncol)
    st_act = _row_styles(ws, act_row, 1, ncol)
    st_blank = _row_styles(ws, blank_row, 1, ncol)
    old_details_last = ws.max_row
    old_top_lo, old_top_hi = (top_hdr + 1, stage_row - 3) if top_hdr else (0, 0)
    old_det_lo = first_proj
    cf_rules = [(str(cf.sqref), list(cf.rules)) for cf in ws.conditional_formatting]
    dvs = list(ws.data_validations.dataValidation)

    start_clear = (top_hdr + 1) if top_hdr else stage_row - 1
    _unmerge_from(ws, start_clear)
    for r in range(start_clear, old_details_last + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.hyperlink = None
            _apply_style(cell, st_blank.get(c))
    # layout
    n = len(projects)
    top_lo = (top_hdr + 1) if top_hdr else None
    top_hi = (top_hdr + n) if top_hdr else None
    title_r = (top_hi + 2) if top_hdr else stage_row - 1
    head_r = title_r + 1
    r = head_r + 1
    hdr_rows = []
    for p in projects:
        hdr_rows.append(r)
        acts = p["acts"]
        end = r + len(acts)
        a_rng = f"$A${r + 1}:$A${max(end, r + 1)}"
        b_rng = f"$B${r + 1}:$B${max(end, r + 1)}"
        c_rng = f"$C${r + 1}:$C${max(end, r + 1)}"
        d_rng = f"$D${r + 1}:$D${max(end, r + 1)}"
        done = f'SUMPRODUCT(ISNUMBER({a_rng})*ISERROR(SEARCH("trigger",{b_rng}))*ISNUMBER({d_rng}))'
        tot = f'SUMPRODUCT(ISNUMBER({a_rng})*ISERROR(SEARCH("trigger",{b_rng}))*((({c_rng}<>"")+({d_rng}<>""))>0))'
        vals = [p["name"], None, f'={done}&" of "&{tot}&" activities done"', None, p["status"] or None, None, f"=A{r}"]
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c)
            _apply_style(cell, st_proj.get(c))
            cell.value = vals[c - 1]
        for a in acts:
            r += 1
            for c in range(1, ncol + 1):
                _apply_style(ws.cell(r, c), st_act.get(c))
            ws.cell(r, 1).value = a[0]
            ws.cell(r, 2).value = a[1]
            for c, v in ((3, a[2]), (4, a[3])):
                dv = to_date(v)
                ws.cell(r, c).value = datetime.combine(dv, datetime.min.time()) if dv else (v if not blank(v) else None)
            ws.cell(r, 6).value = a[4]
            ws.cell(r, 7).value = f"=G{r - 1}"
            ex = a[5] if len(a) > 5 and isinstance(a[5], dict) else {}
            for k, (cid, lab) in enumerate(extras):
                cell = ws.cell(r, 8 + k)
                _apply_style(cell, st_act.get(6))
                v = ex.get(cid)
                cell.value = datetime.combine(to_date(v), datetime.min.time()) if isinstance(v, str) and re.match(r"\d{4}-\d{2}-\d{2}$", v) else v
        r += 2  # blank separator row
    last = max(r - 2, head_r + 1)
    # section title + header
    ws.cell(title_r, 1).value = title_text or "2. Project-wise Details"
    _apply_style(ws.cell(title_r, 1), st_title)
    for c in range(1, ncol + 1):
        _apply_style(ws.cell(head_r, c), st_head.get(c))
        ws.cell(head_r, c).value = head_vals[c - 1]
    for k in range(30):  # extra columns header (clear old, write current)
        ws.cell(head_r, 8 + k).value = None
    for k, (cid, lab) in enumerate(extras):
        _apply_style(ws.cell(head_r, 8 + k), st_head.get(6))
        ws.cell(head_r, 8 + k).value = lab
        ws.column_dimensions[get_column_letter(8 + k)].width = max(12, min(40, len(str(lab)) + 6))
    if head_h:
        ws.row_dimensions[head_r].height = head_h
    # top status table
    if top_hdr:
        for i, h in enumerate(hdr_rows):
            rr = top_lo + i
            for c in range(1, 5):
                _apply_style(ws.cell(rr, c), st_top.get(c))
            ws.cell(rr, 1).value = i + 1
            ws.cell(rr, 2).value = f"=TRIM(A{h})"
            ws.cell(rr, 2).hyperlink = f"#'{ws.title}'!A{h}"
            ws.cell(rr, 3).value = f'=IF(TRIM(E{h})="","(not set)",TRIM(E{h}))'
            ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4)
    # conditional formatting & validations moved to the new ranges
    ws.conditional_formatting = type(ws.conditional_formatting)()
    for ref, rules in cf_rules:
        if top_hdr and CellRange(ref.split()[0]).min_row <= old_top_hi:
            new = _remap(ref, old_top_lo, old_top_hi, top_lo, max(top_hi, top_lo))
        else:
            new = _remap(ref, old_det_lo, old_details_last, head_r + 1, last)
        for rule in rules:
            ws.conditional_formatting.add(new, rule)
    for dv in dvs:
        dv.sqref = openpyxl.worksheet.cell_range.MultiCellRange(_remap(str(dv.sqref), old_det_lo, old_details_last, head_r + 1, last))
    ws.auto_filter.ref = f"A{head_r}:{get_column_letter(6 if not extras else 7 + len(extras))}{last}"
    ws.print_title_rows = f"{head_r}:{head_r}"


def _write_teknic_sheet(ws, tk, extras=()):
    hr, _ = _find_header(ws, r"Part No\.?")
    if hr is None:
        return
    c0 = next(c for c in range(1, ws.max_column + 1) if ws.cell(hr, c).value not in (None, ""))
    st_part = _row_styles(ws, hr + 1, c0, c0 + 11)
    tot_r = next((r for r in range(hr + 1, ws.max_row + 1) if re.search("TOTAL", str(ws.cell(r, c0).value or ""), re.I)), hr + 2)
    st_tot = _row_styles(ws, tot_r, c0, c0 + 11)
    stn_r = next((r for r in range(tot_r, ws.max_row + 1) if re.search("Stations per year", str(ws.cell(r, c0).value or ""), re.I)), tot_r + 2)
    st_stn = _row_styles(ws, stn_r, c0, c0 + 2)
    st_ann = _row_styles(ws, stn_r + 1, c0, c0 + 2)
    stn_note = ws.cell(stn_r, c0 + 2).value
    _unmerge_from(ws, hr + 1)
    for r in range(hr + 1, ws.max_row + 1):
        for c in range(c0, c0 + 12):
            ws.cell(r, c).value = None
            ws.cell(r, c).style = "Normal"
    L = lambda k: get_column_letter(c0 + k)
    parts = tk["parts"]
    for i, p in enumerate(parts):
        r = hr + 1 + i
        vals = [p.get("pn"), p.get("name"), p.get("make"), p.get("mpn"), num(p.get("bom")), p.get("exs"), p.get("alts"),
                num(p.get("exp")), num(p.get("qp")), f"={L(7)}{r}-{L(8)}{r}", f"=IFERROR({L(9)}{r}/{L(7)}{r},0)",
                f"={L(9)}{r}*{L(4)}{r}"]
        for k, v in enumerate(vals):
            cell = ws.cell(r, c0 + k)
            _apply_style(cell, st_part.get(c0 + k))
            cell.value = v
    lo, hi = hr + 1, hr + max(len(parts), 1)
    t = hi + 1
    tot_vals = {0: "TOTAL PER STATION", 7: f"=SUMPRODUCT({L(4)}{lo}:{L(4)}{hi},{L(7)}{lo}:{L(7)}{hi})",
                8: f"=SUMPRODUCT({L(4)}{lo}:{L(4)}{hi},{L(8)}{lo}:{L(8)}{hi})", 10: f"=IFERROR({L(11)}{t}/{L(7)}{t},0)",
                11: f"=SUM({L(11)}{lo}:{L(11)}{hi})"}
    for k in range(12):
        cell = ws.cell(t, c0 + k)
        _apply_style(cell, st_tot.get(c0 + k))
        cell.value = tot_vals.get(k)
    s = t + 2
    for k in range(3):
        _apply_style(ws.cell(s, c0 + k), st_stn.get(c0 + k))
        _apply_style(ws.cell(s + 1, c0 + k), st_ann.get(c0 + k))
    ws.cell(s, c0).value = "Stations per year"
    ws.cell(s, c0 + 1).value = num(tk.get("stations"))
    ws.cell(s, c0 + 2).value = stn_note
    ws.cell(s + 1, c0).value = "Annual saving (₹)"
    ws.cell(s + 1, c0 + 1).value = f'=IF({L(1)}{s}="","Enter stations",{L(11)}{t}*{L(1)}{s})'
    head_style = _style_of(ws.cell(hr, c0 + 11))
    for k in range(20):
        ws.cell(hr, c0 + 12 + k).value = None
    for k, (cid, lab) in enumerate(extras):
        _apply_style(ws.cell(hr, c0 + 12 + k), head_style)
        ws.cell(hr, c0 + 12 + k).value = lab
        ws.column_dimensions[L(12 + k)].width = max(12, min(40, len(str(lab)) + 6))
        for i, p in enumerate(parts):
            cell = ws.cell(hr + 1 + i, c0 + 12 + k)
            _apply_style(cell, st_part.get(c0 + 1))
            cell.value = p.get(cid)


def formatted_export(orig: bytes, d: dict) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(orig))
    def xcols(t):
        m = tmeta(d, t)
        return [(c, m["labels"][c]) for c in m["order"] if c in m["custom"]]

    for key, pat in (("mqis", r"Part No"), ("ds6", r"Part No\.?")):
        ws = _sheet(wb, key)
        if ws is not None:
            _write_parts_sheet(ws, pat, key, d)
    ws = _sheet(wb, "projects")
    if ws is not None:
        _write_projects_sheet(ws, d["projects"], xcols("acts"))
    ws = _sheet(wb, "teknic")
    if ws is not None:
        _write_teknic_sheet(ws, d["tk"], xcols("teknic"))
    for name in [n for n in wb.sheetnames if n.startswith("DC Fan ") and n != SHEETS["fan"]]:
        del wb[name]
    for i, f in enumerate(d["fanExtra"]):
        _fan_extra_sheet(wb, f, i + 2)
    wb.calculation.fullCalcOnLoad = True  # Excel recalculates all formulas when opened
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def backup_bytes(d: dict) -> bytes:
    return json.dumps({"app": APP_TITLE, "savedAt": datetime.now().isoformat(timespec="seconds"), "data": d},
                      ensure_ascii=False, indent=1, default=jsonable).encode("utf-8")


# ----------------------------------------------------------------------------------------------
# Session state
# ----------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------
# Permanent storage: GitHub repository (private) via the GitHub Contents API
# Configure in Streamlit Cloud → App settings → Secrets:
#   [github]
#   token  = "github_pat_..."      # fine-grained token, this repo only, Contents: Read and write
#   repo   = "hemanthv83/NPD"
#   branch = "data"                # separate branch → saving does not restart the app
#   path   = "npd_data.json"
# ---------------------------------------------------------------------------------------------
def gh_cfg():
    try:
        g = st.secrets.get("github", None)
    except Exception:  # noqa: BLE001  (no secrets file)
        return None
    if not g or not g.get("token") or not g.get("repo"):
        return None
    return {"token": g["token"], "repo": g["repo"], "branch": g.get("branch", "data"), "path": g.get("path", "npd_data.json"),
            "base": g.get("base_branch", "main")}


def _gh(method, url, cfg, **kw):
    import requests
    h = {"Authorization": f"Bearer {cfg['token']}", "Accept": "application/vnd.github+json",
         "X-GitHub-Api-Version": "2022-11-28"}
    return requests.request(method, f"https://api.github.com/repos/{cfg['repo']}{url}", headers=h, timeout=20, **kw)


def gh_ensure_branch(cfg):
    r = _gh("GET", f"/git/ref/heads/{cfg['branch']}", cfg)
    if r.status_code == 200:
        return
    if r.status_code != 404:
        raise RuntimeError(f"GitHub {r.status_code}: {r.json().get('message', r.text)[:200]}")
    base = _gh("GET", f"/git/ref/heads/{cfg['base']}", cfg)
    if base.status_code != 200:
        raise RuntimeError(f"Cannot read branch {cfg['base']}: {base.json().get('message', '')}")
    c = _gh("POST", "/git/refs", cfg, json={"ref": f"refs/heads/{cfg['branch']}", "sha": base.json()["object"]["sha"]})
    if c.status_code not in (200, 201):
        raise RuntimeError(f"Cannot create branch {cfg['branch']}: {c.json().get('message', '')}")


def gh_load(cfg):
    """Return (data, sha) from GitHub, or (None, None) if the file does not exist yet."""
    import base64
    gh_ensure_branch(cfg)
    r = _gh("GET", f"/contents/{cfg['path']}", cfg, params={"ref": cfg["branch"]})
    if r.status_code == 404:
        return None, None
    if r.status_code != 200:
        raise RuntimeError(f"GitHub {r.status_code}: {r.json().get('message', r.text)[:200]}")
    j = r.json()
    if j.get("content"):
        raw = base64.b64decode(j["content"])
    else:  # files > 1 MB come without inline content
        r2 = _gh("GET", f"/git/blobs/{j['sha']}", cfg)
        raw = base64.b64decode(r2.json()["content"])
    return parse_backup(raw), j["sha"]


def gh_save(cfg, data, sha, message):
    """Write data to GitHub; returns new sha. Retries once if the file changed meanwhile."""
    import base64
    body = {"message": f"NPD tracker: {message}", "branch": cfg["branch"],
            "content": base64.b64encode(backup_bytes(data)).decode()}
    for attempt in (1, 2):
        if sha:
            body["sha"] = sha
        r = _gh("PUT", f"/contents/{cfg['path']}", cfg, json=body)
        if r.status_code in (200, 201):
            return r.json()["content"]["sha"]
        if r.status_code in (409, 422) and attempt == 1:  # stale sha → fetch latest sha and retry
            g = _gh("GET", f"/contents/{cfg['path']}", cfg, params={"ref": cfg["branch"]})
            sha = g.json().get("sha") if g.status_code == 200 else None
            continue
        raise RuntimeError(f"GitHub {r.status_code}: {r.json().get('message', r.text)[:200]}")


# ----------------------------------------------------------------------------------------------
# Session state
# ----------------------------------------------------------------------------------------------
def init_state():
    ss = st.session_state
    if "data" in ss:
        return
    ss.data = empty_data()
    ss.source = "No data loaded"
    ss.orig_bytes = None
    ss.orig_name = None
    ss.changed = False
    ss.last_change = None
    ss.gh_sha = None
    ss.gh_status = None
    if os.path.exists(DEFAULT_DATA_FILE):
        try:
            with open(DEFAULT_DATA_FILE, "rb") as fh:
                ss.data = parse_backup(fh.read())
            ss.source = "npd_data.json (repository)"
        except Exception as e:  # noqa: BLE001
            ss.load_error = f"Could not read npd_data.json: {e}"
    cfg = gh_cfg()
    if cfg:
        try:
            data, sha = gh_load(cfg)
            if data is not None:
                ss.data, ss.gh_sha = data, sha
                ss.source = f"GitHub · {cfg['repo']} ({cfg['branch']})"
                ss.gh_status = ("ok", "Permanent saving ON (GitHub)")
            else:  # first run: seed the data branch with the current data
                ss.gh_sha = gh_save(cfg, ss.data, None, "initial data")
                ss.source = f"GitHub · {cfg['repo']} ({cfg['branch']})"
                ss.gh_status = ("ok", "Permanent saving ON (GitHub) — data branch created")
        except Exception as e:  # noqa: BLE001
            ss.gh_status = ("err", f"GitHub saving not working: {e}")
    if AUTOSAVE_FILE and os.path.exists(AUTOSAVE_FILE):
        try:
            with open(AUTOSAVE_FILE, "rb") as fh:
                ss.data = parse_backup(fh.read())
            ss.source = f"Auto-save file ({os.path.basename(AUTOSAVE_FILE)})"
        except Exception as e:  # noqa: BLE001
            ss.load_error = f"Could not read the auto-save file: {e}"


def commit(msg: str):
    """Record a change and save it permanently (GitHub) or to the local auto-save file."""
    ss = st.session_state
    ss.last_change = datetime.now()
    ss.changed = True
    cfg = gh_cfg()
    if cfg:
        try:
            ss.gh_sha = gh_save(cfg, ss.data, ss.get("gh_sha"), msg)
            ss.changed = False
            ss.gh_status = ("ok", f"Saved to GitHub at {datetime.now():%H:%M}")
            msg += " · saved permanently"
        except Exception as e:  # noqa: BLE001
            ss.gh_status = ("err", f"NOT saved to GitHub: {e}")
            msg += " · NOT saved permanently — download a backup"
    if AUTOSAVE_FILE:
        try:
            with open(AUTOSAVE_FILE, "wb") as fh:
                fh.write(backup_bytes(ss.data))
            ss.changed = False
            msg += " · auto-saved"
        except OSError as e:
            msg += f" · auto-save failed ({e})"
    ss.flash = msg


def D() -> dict:
    return st.session_state.data


def has_data() -> bool:
    d = D()
    return bool(d["projects"] or d["mqis"][1] or d["ds6"][1] or d["tk"]["parts"] or d["fan"] or d["fanExtra"])


# ----------------------------------------------------------------------------------------------
# UI helpers
# ----------------------------------------------------------------------------------------------
CSS = f"""
<style>
:root {{ --navy:{COLORS['navy']}; --ink:{COLORS['ink']}; --muted:{COLORS['muted']}; --line:{COLORS['line']}; }}
.block-container {{ padding-top: 3.2rem; max-width: 1400px; }}
[data-testid="stSidebar"] {{ background: #0F2537; }}
[data-testid="stSidebar"] * {{ color: #E4E9F0 !important; }}
[data-testid="stSidebar"] .stRadio label p {{ font-size: 0.95rem; }}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {{ background:#16324A; }}
[data-testid="stSidebar"] button {{ background:#16324A !important; border-color:#2B4B66 !important; }}
.npd-title {{ font-size: 1.7rem; line-height:1.2; font-weight: 800; color: var(--ink); margin: 0; letter-spacing:-.01em; }}
.npd-sub {{ color: var(--muted); font-size: .85rem; margin: 2px 0 14px; }}
.kpi {{ background:#fff; border:1px solid var(--line); border-radius:10px; padding:14px 16px; height:100%; }}
.kpi .lab {{ font-size:.68rem; text-transform:uppercase; letter-spacing:.07em; color:var(--muted); font-weight:700; }}
.kpi .val {{ font-size:1.9rem; font-weight:800; color:var(--ink); line-height:1.15; font-variant-numeric:tabular-nums; }}
.kpi .sub {{ font-size:.78rem; color:var(--muted); }}
.bar {{ height:6px; background:#EEF2F7; border-radius:3px; overflow:hidden; margin:6px 0 4px; }}
.bar i {{ display:block; height:100%; border-radius:3px; }}
.pill {{ display:inline-block; padding:2px 10px; border-radius:999px; font-size:.72rem; font-weight:700; white-space:nowrap; line-height:1.5; height:fit-content; flex-shrink:0; align-self:flex-start; }}
.chip {{ display:inline-block; padding:1px 8px; border-radius:4px; font-size:.7rem; font-weight:700; color:#fff; margin-right:4px; }}
.pcard {{ background:#fff; border:1px solid var(--line); border-radius:10px; padding:12px 14px; margin-bottom:12px; }}
.pcard h4 {{ margin:0; font-size:.98rem; color:var(--ink); }}
.pcard .meta {{ font-size:.78rem; color:var(--muted); display:flex; justify-content:space-between; }}
.flag {{ font-size:.7rem; font-weight:700; padding:1px 6px; border-radius:4px; }}
.section {{ font-size:1.05rem; font-weight:700; color:var(--ink); margin:18px 0 8px; }}
.note {{ color:var(--muted); font-size:.8rem; }}
</style>
"""


def pill(status: str) -> str:
    s = status or "Not set"
    color = STATUS_COLORS.get(s, COLORS["idle"])
    soft = {COLORS["ok"]: COLORS["ok_soft"], COLORS["warn"]: COLORS["warn_soft"]}.get(color, COLORS["idle_soft"])
    return f'<span class="pill" style="background:{soft};color:{color}">● {s}</span>'


def kpi(label, value, sub="", pct=None, bar_color=None, value_color=None):
    bar = f'<div class="bar"><i style="width:{pct:.0f}%;background:{bar_color or COLORS["navy"]}"></i></div>' if pct is not None else ""
    vc = f"color:{value_color};" if value_color else ""
    st.markdown(f'<div class="kpi"><div class="lab">{label}</div><div class="val" style="{vc}">{value}</div>{bar}'
                f'<div class="sub">{sub}</div></div>', unsafe_allow_html=True)


def header(title, sub=""):
    st.markdown(f'<div class="npd-title">{title}</div><div class="npd-sub">{sub}</div>', unsafe_allow_html=True)


def no_data_notice():
    st.info("No data loaded yet. Use **Data Management** (sidebar) to upload your NPD Master Tracker Excel or a JSON backup.")


def flash():
    msg = st.session_state.pop("flash", None)
    if msg:
        st.toast(msg, icon="✅")
    err = st.session_state.pop("flash_err", None)
    if err:
        st.error(err, icon="⚠️")


def plot_layout(fig, h=360):
    fig.update_layout(height=h, margin=dict(l=10, r=10, t=30, b=10), paper_bgcolor="white", plot_bgcolor="white",
                      font=dict(family="Segoe UI, Arial", size=12, color=COLORS["ink"]),
                      legend=dict(orientation="h", y=-0.15))
    fig.update_xaxes(gridcolor="#EEF2F7", zeroline=False)
    fig.update_yaxes(gridcolor="#EEF2F7", zeroline=False)
    return fig


def activities_df(d) -> pd.DataFrame:
    rows = []
    for p in d["projects"]:
        for a in p["acts"]:
            stt = activity_status(a)
            pl = to_date(a[2])
            rows.append({"Project": p["name"], "Project Status": p["status"] or "Not set", "Step": a[0],
                         "Activity": a[1], "Plan Date": pl, "Actual Date": to_date(a[3]), "Status": stt,
                         "Days Late": (today() - pl).days if stt == "Delayed" and pl else None, "Remarks": a[4]})
    return pd.DataFrame(rows)


def parts_union(d) -> pd.DataFrame:
    rows = []
    for key in ("mqis", "ds6"):
        cfg = PCFG[key]
        for r in d[key][1]:
            qty = num(r[cfg["qty"]])
            price = r[cfg["price"]] if has_price(key, r) else None
            rows.append({"Sheet": cfg["label"], "Part No": r[1], "Part Name": r[2],
                         "Supplier": (str(r[cfg["supplier"]]).strip() if not blank(r[cfg["supplier"]]) else "(not set)"),
                         "Project": r[cfg["proj"]] if cfg["proj"] is not None else "6DS",
                         "Qty": qty, "Price": price, "Has MPN": has_mpn(key, r), "Has Price": has_price(key, r),
                         "Value": (qty or 0) * price if price is not None else None,
                         "MAKE/MPN": next((r[i] for i in cfg["mpn"] if not blank(r[i])), None)})
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------------------------
# Editable-table engine: auto-save on every edit; insert / move / duplicate / delete rows anywhere;
# insert / move / rename / delete columns. Used by activities, parts lists and Teknic parts.
# ----------------------------------------------------------------------------------------------
TABLE_DEFS = {
    "mqis": {"core": ["slno", "partno", "partname", "qty", "uom", "buyer", "supplier", "project", "mpn1", "mpn2", "mpn3", "price"],
             "types": {"qty": "number", "price": "number"}, "renumber": "slno", "label_col": "partname",
             "required": [], "unique": "partno", "title": "MQIS & Single Phase parts"},
    "ds6": {"core": ["slno", "partno", "partname", "rev", "qty", "uom", "qty50", "vendor", "mpn", "price"],
            "types": {"qty": "number", "qty50": "number", "price": "number"}, "renumber": "slno", "label_col": "partname",
            "required": [], "unique": "partno", "title": "6DS parts"},
    "acts": {"core": ["step", "activity", "plan", "actual", "remarks"],
             "labels": ["Step", "Activity", "Plan Date", "Actual Date", "Remarks"],
             "types": {"step": "number", "plan": "date", "actual": "date"}, "renumber": "step", "label_col": "activity",
             "required": [], "unique": None, "title": "activities"},
    "teknic": {"core": ["pn", "name", "make", "mpn", "bom", "exs", "alts", "exp", "qp"],
               "labels": ["Part No.", "Part Name", "Make", "MPN", "BOM / Stn", "Existing supplier", "Alternate supplier",
                          "Existing ₹", "Quoted ₹"],
               "types": {"bom": "number", "exp": "number", "qp": "number"}, "renumber": None, "label_col": "name",
               "required": [], "unique": "pn", "title": "Teknic parts"},
}
TYPE_NAMES = {"Text": "text", "Number": "number", "Date": "date"}


def tmeta(d, t):
    """Column order, labels and user-added columns for table t (stored in the data, so it is saved and shared)."""
    td = TABLE_DEFS[t]
    m = d.setdefault("meta", {}).setdefault(t, {})
    m.setdefault("custom", {})
    labels = m.setdefault("labels", {})
    base = d[t][0] if t in ("mqis", "ds6") else td["labels"]
    for i, cid in enumerate(td["core"]):
        if blank(labels.get(cid)):
            labels[cid] = base[i] if i < len(base) and not blank(base[i]) else cid
    order = [c for c in m.get("order", []) if c in td["core"] or c in m["custom"]]
    for cid in td["core"] + list(m["custom"]):
        if cid not in order:
            order.append(cid)
    m["order"] = order
    m.setdefault("renumber", td["renumber"] is not None)
    m["hidden"] = [c for c in m.get("hidden", []) if c in td["core"]]
    return m


def vis(m):
    """Columns shown (display order minus hidden built-in columns)."""
    return [c for c in m["order"] if c not in m["hidden"]]


def _ctype(m, t, cid):
    return m["custom"].get(cid) or TABLE_DEFS[t]["types"].get(cid, "text")


def _row_to_rec(row, core):
    rec = {c: (row[i] if i < len(row) else None) for i, c in enumerate(core)}
    if len(row) > len(core) and isinstance(row[len(core)], dict):
        rec.update(row[len(core)])
    return rec


def _rec_to_row(rec, core, custom):
    row = [rec.get(c) for c in core]
    ex = {c: rec.get(c) for c in custom if not blank(rec.get(c))}
    return row + [ex] if ex else row


def recs_v(rec, cid):
    return rec.get(cid)


def _to_cell(v, typ):
    if typ == "number":
        return num(v)
    if typ == "date":
        return to_date(v)
    return None if blank(v) else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))


def _from_cell(v, typ):
    if blank(v):
        return None
    if typ == "number":
        n = num(v)
        return int(n) if n is not None and float(n).is_integer() else n
    if typ == "date":
        return iso(v)
    return str(v).strip()


def spec_parts(key):
    core = TABLE_DEFS[key]["core"]

    def get():
        return [_row_to_rec(r, core) for r in D()[key][1]]

    def put(recs):
        custom = tmeta(D(), key)["custom"]
        D()[key][1] = [_rec_to_row(r, core, custom) for r in recs]
    return {"t": key, "key": key, "get": get, "set": put}


def spec_acts(pi):
    core = TABLE_DEFS["acts"]["core"]

    def get():
        return [_row_to_rec(a, core) for a in D()["projects"][pi]["acts"]]

    def put(recs):
        custom = tmeta(D(), "acts")["custom"]
        D()["projects"][pi]["acts"] = [_rec_to_row(r, core, custom) for r in recs]
    return {"t": "acts", "key": f"acts_{pi}", "get": get, "set": put}


def spec_teknic():
    def get():
        return [dict(p) for p in D()["tk"]["parts"]]

    def put(recs):
        D()["tk"]["parts"] = [{k: v for k, v in r.items() if not blank(v) or k in TABLE_DEFS["teknic"]["core"]} for r in recs]
    return {"t": "teknic", "key": "teknic", "get": get, "set": put}


def _bump(spec):
    k = f"ver_{spec['key']}"
    st.session_state[k] = st.session_state.get(k, 0) + 1


def _validate(t, recs):
    u = TABLE_DEFS[t]["unique"]
    if not u:
        return None
    seen = {}
    for i, r in enumerate(recs):
        v = part_no(r.get(u))
        if v is None:
            continue
        r[u] = v
        if str(v) in seen:
            return f"Not saved: {TABLE_DEFS[t]['title']} — No. {v} is already in row {seen[str(v)] + 1}. Use a different number."
        seen[str(v)] = i
    return None


def _renumber(t, recs):
    f = TABLE_DEFS[t]["renumber"]
    if f:
        for i, r in enumerate(recs):
            r[f] = i + 1


def _finish(spec, recs, structural, msg):
    d = D()
    t = spec["t"]
    err = _validate(t, recs)
    if err:
        st.session_state.flash_err = err
        _bump(spec)
        return False
    if structural and tmeta(d, t).get("renumber"):
        _renumber(t, recs)
    spec["set"](recs)
    commit(msg)
    _bump(spec)
    return True


# ----------------------------------------------------------------------------------------------
# Excel-style grid (jspreadsheet CE, MIT licence) as a Streamlit component:
# right-click insert/delete/rename, drag rows & columns, Ctrl+Z / Ctrl+Y undo-redo, copy-paste from Excel,
# auto-save of every change. The component's HTML is written at start-up, so the app stays one file.
# ----------------------------------------------------------------------------------------------
GRID_HTML = r"""<!doctype html><html><head><meta charset="utf-8">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/jsuites@5.4.6/dist/jsuites.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/jspreadsheet-ce@4.13.4/dist/jspreadsheet.css">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Material+Icons&display=block">
<script src="https://cdn.jsdelivr.net/npm/jsuites@5.4.6/dist/jsuites.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jspreadsheet-ce@4.13.4/dist/index.js"></script>
<style>
 body{margin:0;font-family:"Segoe UI",Arial,sans-serif;font-size:13px;color:#18222F;background:#fff}
 #bar{display:flex;gap:6px;align-items:center;flex-wrap:wrap;padding:6px 2px 8px}
 #bar button{font:600 12px "Segoe UI",Arial;border:1px solid #D8DFE8;background:#F3F5F8;color:#18222F;border-radius:6px;padding:5px 10px;cursor:pointer}
 #bar button:hover{background:#DCEBF5;border-color:#0B5C8A}
 #bar .sep{width:1px;height:20px;background:#D8DFE8;margin:0 4px}
 #st{margin-left:auto;font-size:12px;color:#5D6B7C}
 #st.err{color:#C2372D;font-weight:600}
 #hint{font-size:11.5px;color:#5D6B7C;padding:0 2px 6px}
 .jexcel > thead > tr > td{background:#1F4E78 !important;color:#fff !important;font-weight:600;border-color:#3d6a93 !important}
 .jexcel > tbody > tr > td:first-child{background:#EEF2F7}
 .jexcel > tbody > tr > td{font-size:12.5px}
 .jexcel > tbody > tr > td.readonly{color:#5D6B7C;background:#FAFBFC}
 .jexcel_content{border:1px solid #D8DFE8}
</style></head><body>
<div id="bar">
 <button id="b-undo" title="Ctrl+Z">↶ Undo</button><button id="b-redo" title="Ctrl+Y">↷ Redo</button><span class="sep"></span>
 <button id="b-rowa" title="Insert a row above the selected cell">+ Row above</button><button id="b-rowb" title="Insert a row below the selected cell">+ Row below</button>
 <button id="b-coll" title="Insert a column left of the selected cell">+ Column left</button><button id="b-colr" title="Insert a column right of the selected cell">+ Column right</button>
 <button id="b-ren" style="display:none">Renumber</button><span id="st"></span></div>
<div id="hint">Right-click a cell or a column header to insert, delete or rename. Drag a row number or a header to move it. Double-click a header to rename. Changes save automatically.</div>
<div id="sheet"></div>
<script>
const post=(type,o)=>window.parent.postMessage(Object.assign({isStreamlitMessage:true,type:type},o||{}),"*");
let table=null, rev=null, timer=null, args=null, sending=false;
const $=id=>document.getElementById(id);
function height(){post("streamlit:setFrameHeight",{height:document.body.scrollHeight+8});}
function status(t,err){const s=$("st");s.textContent=t||"";s.className=err?"err":"";}
let lastSel=null;
function sel(){return lastSel;}
function payload(){const cols=table.options.columns;return{rev:rev,nonce:Date.now()+"-"+Math.random(),headers:table.getHeaders(true),
  ids:cols.map(c=>c.name||null),data:table.getData(false)};}
function changed(){if(table&&table.ignoreEvents)return;status("Saving…");clearTimeout(timer);
  timer=setTimeout(()=>{post("streamlit:setComponentValue",{value:payload(),dataType:"json"});},700);}
function isLocked(i){const c=table.options.columns[i];return c&&c.name&&(c.locked||String(c.name).startsWith("__comp_"));}
function renumber(){if(!args||!args.renumber)return;const ci=table.options.columns.findIndex(c=>c.name===args.renumber);if(ci<0)return;
  for(let r=0;r<table.rows.length;r++)table.setValueFromCoords(ci,r,r+1,true);changed();}
function updateComputed(a){if(!table||!a.computed)return;const cols=table.options.columns;table.ignoreEvents=true;table.ignoreHistory=true;
  try{a.computed.forEach((row,r)=>{if(r>=table.rows.length)return;cols.forEach((c,ci)=>{if(c.name&&String(c.name).startsWith("__comp_")){
   const v=row[String(c.name).slice(7)];table.setValueFromCoords(ci,r,v==null?"":v,true);}});});}
  finally{table.ignoreEvents=false;table.ignoreHistory=false;}}
function build(a){lastSel=null;const el=$("sheet");if(table){try{jspreadsheet.destroy(el,true);}catch(e){}table=null;}el.innerHTML="";
  table=jspreadsheet(el,{data:a.data.length?a.data:[a.columns.map(()=>"")],columns:a.columns,tableOverflow:true,
   tableHeight:a.height+"px",tableWidth:(document.body.clientWidth-4)+"px",rowDrag:true,columnDrag:true,columnSorting:false,
   allowInsertRow:true,allowDeleteRow:true,allowInsertColumn:true,allowDeleteColumn:true,allowRenameColumn:true,
   allowManualInsertRow:true,allowManualInsertColumn:true,wordWrap:false,
   onselection:(el,x1,y1,x2,y2)=>{lastSel={x1:+Math.min(x1,x2),y1:+Math.min(y1,y2),x2:+Math.max(x1,x2),y2:+Math.max(y1,y2)};},
   onbeforedeletecolumn:(el,col,num)=>{for(let i=col;i<col+num;i++){if(isLocked(i)){status("Built-in columns cannot be deleted. You can rename or move them.",true);return false;}}},
   onbeforechange:(el,cell,x,y,v)=>{if(isLocked(+x)&&String(table.options.columns[+x].name).startsWith("__comp_"))return false;},
   onchange:changed,oninsertrow:changed,ondeleterow:changed,oninsertcolumn:changed,ondeletecolumn:changed,onmoverow:changed,
   onmovecolumn:changed,onchangeheader:changed,onundo:changed,onredo:changed,onpaste:changed,onsort:changed});
  $("b-ren").style.display=a.renumber?"":"none";}
$("b-undo").onclick=()=>{table&&table.undo();};$("b-redo").onclick=()=>{table&&table.redo();};
$("b-rowa").onclick=()=>{if(!table)return;const s=sel();table.insertRow(1,s?s.y1:0,true);};
$("b-rowb").onclick=()=>{if(!table)return;const s=sel();table.insertRow(1,s?s.y2:table.rows.length-1,false);};
$("b-coll").onclick=()=>{if(!table)return;const s=sel();table.insertColumn(1,s?s.x1:0,true,{title:"New column",type:"text",width:130});};
$("b-colr").onclick=()=>{if(!table)return;const s=sel();table.insertColumn(1,s?s.x2:table.options.columns.length-1,false,{title:"New column",type:"text",width:130});};
$("b-ren").onclick=renumber;
window.addEventListener("message",e=>{const d=e.data;if(!d||d.type!=="streamlit:render")return;args=d.args;
  if(typeof jspreadsheet==="undefined"){status("Spreadsheet library could not load (no internet?).",true);height();return;}
  if(!table||args.rev!==rev){build(args);rev=args.rev;}else{updateComputed(args);}
  if(args.status)status(args.status,args.status_err);else if($("st").textContent==="Saving…")status("Saved");height();});
window.addEventListener("resize",height);
post("streamlit:componentReady",{apiVersion:1});
</script></body></html>"""

_GRID_FN = None


def _grid_component():
    global _GRID_FN
    if _GRID_FN is None:
        import tempfile
        import streamlit.components.v1 as components
        folder = os.path.join(tempfile.gettempdir(), "npd_excel_grid")
        os.makedirs(folder, exist_ok=True)
        with open(os.path.join(folder, "index.html"), "w", encoding="utf-8") as fh:
            fh.write(GRID_HTML)
        _GRID_FN = components.declare_component("npd_excel_grid", path=folder)
    return _GRID_FN


def _cellval(v, typ):
    if blank(v):
        return ""
    if typ == "date":
        dd = to_date(v)
        return dd.isoformat() if dd else str(v)
    if typ == "number":
        n = num(v)
        return n if n is not None else str(v)
    return v


def _grid_build(spec, computed=None, dropdowns=None):
    d = D()
    t = spec["t"]
    td = TABLE_DEFS[t]
    m = tmeta(d, t)
    recs = spec["get"]()
    columns = []
    for cid in m["order"]:
        typ = _ctype(m, t, cid)
        lab = m["labels"][cid]
        col = {"title": lab, "name": cid, "locked": cid in td["core"],
               "width": 70 if cid in ("slno", "step") else 300 if cid in ("activity", "remarks") else
               220 if cid in ("partname", "name", "supplier", "vendor", "mpn1", "mpn", "exs", "alts") else
               110 if typ in ("number", "date") else 150}
        if cid in m["hidden"]:
            col["type"] = "hidden"
        elif dropdowns and cid in dropdowns:
            col.update(type="dropdown", source=dropdowns[cid], autocomplete=True)
        elif typ == "number":
            col["type"] = "numeric"
        elif typ == "date":
            col.update(type="calendar", options={"format": "DD-MMM-YY"})
        else:
            col["type"] = "text"
        col["align"] = "right" if typ == "number" else "center" if typ == "date" or cid in ("slno", "step") else "left"
        columns.append(col)
    data = [[_cellval(r.get(cid), _ctype(m, t, cid)) for cid in m["order"]] for r in recs]
    comp_rows = None
    if computed:
        comp_rows = [computed(r) for r in recs]
        titles = list(computed({}).keys())
        for ti in titles:
            columns.append({"title": ti, "name": "__comp_" + ti, "type": "text", "readOnly": True, "width": 110, "locked": True,
                            "align": "center"})
        for row, cr in zip(data, comp_rows):
            row.extend("" if cr.get(ti) is None else cr.get(ti) for ti in titles)
    sig = hashlib.md5(json.dumps([[c["title"], c["name"], c["type"]] for c in columns] +
                                 [[x for x in row[:len(m["order"])]] for row in data], default=str).encode()).hexdigest()
    return columns, data, comp_rows, sig


def _grid_apply(spec, val):
    d = D()
    t = spec["t"]
    td = TABLE_DEFS[t]
    m = tmeta(d, t)
    headers = val.get("headers") or []
    ids = val.get("ids") or [None] * len(headers)
    lab2id = {_nk(lab): cid for cid, lab in m["labels"].items()}
    used, colids = set(), []
    for h, i in zip(headers, ids):
        if i and str(i).startswith("__comp_"):
            colids.append(None)
            continue
        cid = i if (i in td["core"] or i in m["custom"]) and i not in used else None
        if not cid:
            c2 = lab2id.get(_nk(h))
            cid = c2 if c2 and c2 not in used else None
        title = str(h).strip() if not blank(h) else "New column"
        if not cid:
            cid = _slug(title, set(m["order"]) | set(m["custom"]) | used)
            m["custom"][cid] = "text"
        m["labels"][cid] = title
        used.add(cid)
        colids.append(cid)
    for cid in list(m["custom"]):
        if cid not in used:
            m["custom"].pop(cid, None)
            m["labels"].pop(cid, None)
    m["hidden"] = sorted(set(m["hidden"]) | {c for c in td["core"] if c not in used}, key=td["core"].index)
    m["order"] = [c for c in colids if c] + [c for c in td["core"] if c not in used]
    if t in ("mqis", "ds6"):
        d[t][0] = [m["labels"][c] for c in td["core"]]
    recs = []
    for row in val.get("data") or []:
        rec = {}
        for cid, v in zip(colids, row):
            if cid:
                rec[cid] = _from_cell(v, _ctype(m, t, cid))
        recs.append(rec)
    while recs and all(blank(v) for v in recs[-1].values()):
        recs.pop()
    err = _validate(t, recs)
    if err:
        return False, err
    spec["set"](recs)
    commit(f"{td['title'].capitalize()} saved")
    return True, None


def excel_grid(spec, computed=None, dropdowns=None, height=460, help_text=None):
    """Excel-like editable grid with auto-save, undo/redo, insert/move/rename rows and columns."""
    ss = st.session_state
    k = spec["key"]
    columns, data, comp_rows, sig = _grid_build(spec, computed, dropdowns)
    if ss.get(f"gsig_{k}") != sig:  # data changed outside the grid (upload, other page) → reload grid
        ss[f"gsig_{k}"] = sig
        ss[f"grev_{k}"] = ss.get(f"grev_{k}", 0) + 1
    if help_text:
        st.caption(help_text)
    err = ss.get(f"gerr_{k}")
    td = TABLE_DEFS[spec["t"]]
    val = _grid_component()(rev=ss[f"grev_{k}"], columns=columns, data=data, computed=comp_rows, height=height,
                            renumber=td["renumber"], status=err or "", status_err=bool(err), key=f"grid_{k}", default=None)
    if val and val.get("nonce") != ss.get(f"gnonce_{k}") and val.get("rev") == ss[f"grev_{k}"]:
        ss[f"gnonce_{k}"] = val.get("nonce")
        ok, msg = _grid_apply(spec, val)
        if ok:
            ss.pop(f"gerr_{k}", None)
            ss[f"gsig_{k}"] = _grid_build(spec, computed, dropdowns)[3]
            st.rerun()
        else:
            ss[f"gerr_{k}"] = msg
            st.error(msg)


def _rename_activity(scope_pi):
    d = D()
    old = st.session_state.get("ren_old")
    new = (st.session_state.get("ren_new") or "").strip()
    if not old or not new:
        st.session_state.flash_err = "Pick the activity and type the new name."
        return
    allp = st.session_state.get("ren_scope", "This project only") == "All projects"
    n = 0
    for i, p in enumerate(d["projects"]):
        if allp or i == scope_pi:
            for a in p["acts"]:
                if (a[1] or "").strip() == old:
                    a[1] = new
                    n += 1
    if allp:
        d["stdActs"] = [new if x == old else x for x in (d.get("stdActs") or [])]
    st.session_state["ren_new"] = ""
    commit(f'Renamed "{old}" → "{new}" in {n} place(s)')


def _move_project(pi, delta):
    d = D()
    j = pi + delta
    if 0 <= j < len(d["projects"]):
        d["projects"][pi], d["projects"][j] = d["projects"][j], d["projects"][pi]
        commit("Project order changed")


# ----------------------------------------------------------------------------------------------
# Pages
# ----------------------------------------------------------------------------------------------
def page_dashboard():
    d = D()
    header("NPD Dashboard", f"Data: {st.session_state.source} · Today {today():%d %b %Y} · "
                            "Delay = plan date passed with no actual date")
    if not has_data():
        no_data_notice()
        return
    stats = [(p, proj_stats(p)) for p in d["projects"]]
    comp = sum(1 for p, _ in stats if p["status"] == "Completed")
    ong = sum(1 for p, _ in stats if p["status"] == "Ongoing")
    late = sorted([{**l, "proj": p["name"]} for p, s in stats for l in s["late"]], key=lambda x: -x["days"])
    tot = sum(s["total"] for _, s in stats)
    dn = sum(s["done"] for _, s in stats)
    pu = parts_union(d)
    parts = len(pu)
    w_mpn = int(pu["Has MPN"].sum()) if parts else 0
    w_price = int(pu["Has Price"].sum()) if parts else 0
    pct = lambda a, b: round(a / b * 100) if b else 0
    fan_annual = sum(fan_calc(f)["annual"] for f in d["fanExtra"])
    fan_raw_annual = next((num(r[1]) for r in d["fan"] if r and re.match("Annual saving", str(r[0] or ""), re.I)), None) or 0
    tk_annual = teknic_calc(d["tk"])["annual"] or 0

    c = st.columns(6)
    with c[0]:
        kpi("Projects", len(d["projects"]), f'<span style="color:{COLORS["ok"]}">{comp} completed</span> · '
                                            f'<span style="color:{COLORS["warn"]}">{ong} ongoing</span>')
    with c[1]:
        kpi("Activities done", f"{dn}<span style='font-size:.9rem;color:{COLORS['muted']}'> / {tot}</span>",
            f"{pct(dn, tot)}% across all projects", pct(dn, tot))
    with c[2]:
        kpi("Delayed activities", len(late), "plan date passed, not closed", value_color=COLORS["bad"])
    with c[3]:
        by = {}
        for key in ("mqis",):
            for r in d[key][1]:
                k = proj_key(r[7])
                by[k] = by.get(k, 0) + 1
        by["6DS"] = len(d["ds6"][1])
        chips = "".join(f'<span class="chip" style="background:{COLORS[k]}">{"Single Phs" if k == "Single" else k} {v}</span>'
                        for k, v in by.items() if v)
        kpi("Parts tracked", parts, chips)
    with c[4]:
        kpi("MPN coverage", f"{pct(w_mpn, parts)}%", f"{parts - w_mpn} parts without MAKE/MPN", pct(w_mpn, parts), COLORS["ok"])
    with c[5]:
        kpi("Price coverage", f"{pct(w_price, parts)}%", f"{parts - w_price} parts without price", pct(w_price, parts), COLORS["warn"])

    st.markdown('<div class="section">Project status</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    for i, (p, s) in enumerate(stats):
        color = COLORS["ok"] if p["status"] == "Completed" else (COLORS["bad"] if s["late"] else COLORS["warn"])
        flag = (f'<span class="flag" style="background:{COLORS["bad_soft"]};color:{COLORS["bad"]}">{len(s["late"])} delayed</span>'
                if s["late"] else f'<span class="flag" style="background:{COLORS["warn_soft"]};color:{COLORS["warn"]}">{s["soon"]} due ≤7d</span>'
                if s["soon"] else f'<span class="flag" style="background:{COLORS["ok_soft"]};color:{COLORS["ok"]}">on track</span>')
        with cols[i % 4]:
            st.markdown(f'<div class="pcard"><div style="display:flex;justify-content:space-between;gap:8px"><h4>{p["name"]}</h4>'
                        f'{pill(p["status"])}</div><div class="bar"><i style="width:{s["pct"] * 100:.0f}%;background:{color}"></i></div>'
                        f'<div class="meta"><span>{s["done"]} of {s["total"]} activities done</span>{flag}</div></div>',
                        unsafe_allow_html=True)

    g1, g2 = st.columns([3, 2])
    with g1:
        st.markdown('<div class="section">Progress by project</div>', unsafe_allow_html=True)
        if stats:
            df = pd.DataFrame([{"Project": p["name"], "Progress %": round(s["pct"] * 100), "Delayed": len(s["late"]),
                                "Done": s["done"], "Total": s["total"],
                                "Color": "Completed" if p["status"] == "Completed" else ("Has delays" if s["late"] else "On track")}
                               for p, s in stats])
            fig = px.bar(df, x="Progress %", y="Project", orientation="h", color="Color", text="Progress %",
                         hover_data=["Done", "Total", "Delayed"],
                         color_discrete_map={"Completed": COLORS["ok"], "Has delays": COLORS["bad"], "On track": COLORS["warn"]})
            fig.update_traces(texttemplate="%{text}%", textposition="outside")
            fig.update_yaxes(autorange="reversed", title=None)
            fig.update_xaxes(range=[0, 110], title=None)
            st.plotly_chart(plot_layout(fig, 40 * len(df) + 120), width="stretch")
    with g2:
        st.markdown('<div class="section">Activity status (all projects)</div>', unsafe_allow_html=True)
        adf = activities_df(d)
        if not adf.empty:
            cnt = adf[adf["Status"] != "Not planned"]["Status"].value_counts().reset_index()
            cnt.columns = ["Status", "Count"]
            fig = px.pie(cnt, names="Status", values="Count", hole=0.55, color="Status", color_discrete_map=ACT_COLORS)
            fig.update_traces(textinfo="value+label")
            st.plotly_chart(plot_layout(fig, 340), width="stretch")

    s1, s2 = st.columns(2)
    with s1:
        st.markdown('<div class="section">Delayed activities</div>', unsafe_allow_html=True)
        if late:
            st.dataframe(pd.DataFrame([{"Days late": l["days"], "Project": l["proj"], "Activity": l["act"],
                                        "Plan": l["plan"], "Remarks": l["rem"]} for l in late]),
                         hide_index=True, width="stretch", height=360,
                         column_config={"Plan": st.column_config.DateColumn(format="DD MMM YY"),
                                        "Days late": st.column_config.NumberColumn(format="%d d")})
        else:
            st.success("No delayed activities.")
    with s2:
        st.markdown('<div class="section">Parts missing data</div>', unsafe_allow_html=True)
        if parts:
            miss = pu[~pu["Has MPN"] | ~pu["Has Price"]].copy()
            if miss.empty:
                st.success("All parts have MAKE/MPN and price.")
            else:
                miss["Missing"] = miss.apply(lambda r: ", ".join(x for x, ok in (("MPN", r["Has MPN"]), ("Price", r["Has Price"])) if not ok), axis=1)
                st.dataframe(miss[["Part No", "Part Name", "Project", "Supplier", "Missing"]], hide_index=True,
                             width="stretch", height=360)
    if fan_annual or fan_raw_annual or tk_annual:
        st.markdown('<div class="section">Cost savings (annual)</div>', unsafe_allow_html=True)
        k = st.columns(3)
        with k[0]:
            kpi("DC fan comparisons", inr(fan_raw_annual + fan_annual), f"{(1 if fan_raw_annual else 0) + len(d['fanExtra'])} comparison(s)")
        with k[1]:
            kpi("Teknic parts (domestic)", inr(tk_annual), f"{len(d['tk']['parts'])} parts · {int(num(d['tk'].get('stations')) or 0):,} stations/yr")
        with k[2]:
            kpi("Total identified saving", inr(fan_raw_annual + fan_annual + tk_annual), "per year", value_color=COLORS["ok"])


def page_projects():
    d = D()
    header("Project Tracking — Alternate Development",
           "Stage-gate activities per project. Every change saves automatically.")
    with st.expander("➕ Add project", expanded=not d["projects"]):
        with st.form("add_project", clear_on_submit=True):
            c1, c2 = st.columns([3, 1])
            name = c1.text_input("Project / part name *")
            status = c2.selectbox("Status", STATUSES)
            c3, c4 = st.columns([1, 3])
            start = c3.date_input("Start (trigger) date", value=None, format="DD/MM/YYYY")
            use_std = c4.checkbox(f"Add the standard stage-gate activity list ({len(std_acts(d))} steps)", value=True)
            remark = st.text_area("Trigger remark", height=70)
            if st.form_submit_button("Add project", type="primary"):
                if not name.strip():
                    st.error("Enter a project name.")
                elif any(p["name"].lower() == name.strip().lower() for p in d["projects"]):
                    st.error("A project with this name already exists.")
                else:
                    acts = ([[i + 1, n, None, None, None] for i, n in enumerate(std_acts(d))] if use_std
                            else [[1, "Alternate change (trigger / start)", None, None, None]])
                    acts[0][2] = iso(start)
                    acts[0][4] = remark.strip() or None
                    d["projects"].append({"name": name.strip(), "status": status, "acts": acts})
                    st.session_state.sel_project = name.strip()
                    commit("Project added")
                    st.rerun()
    if not d["projects"]:
        no_data_notice()
        return
    names = [p["name"] for p in d["projects"]]
    if st.session_state.get("sel_project") not in names:
        st.session_state.sel_project = names[0]
    c1, c2 = st.columns([3, 1])
    sel = c1.selectbox("Project", names, key="sel_project")
    status_filter = c2.multiselect("Show activity status", list(ACT_COLORS), default=[])
    pi = names.index(sel)
    p = d["projects"][pi]
    s = proj_stats(p)
    k = st.columns(4)
    with k[0]:
        st.markdown(f'<div class="kpi"><div class="lab">Program status</div><div style="margin-top:8px">{pill(p["status"])}</div></div>',
                    unsafe_allow_html=True)
    with k[1]:
        color = COLORS["ok"] if p["status"] == "Completed" else (COLORS["bad"] if s["late"] else COLORS["warn"])
        kpi("Progress", f'{s["pct"] * 100:.0f}%', f'{s["done"]} of {s["total"]} activities done', s["pct"] * 100, color)
    with k[2]:
        kpi("Delayed", len(s["late"]), "plan passed, no actual", value_color=COLORS["bad"] if s["late"] else None)
    with k[3]:
        kpi("Due in 7 days", s["soon"], "open activities")

    rows = []
    for a in p["acts"]:
        rows.append({"Step": num(a[0]), "Activity": a[1], "Plan Date": to_date(a[2]), "Actual Date": to_date(a[3]),
                     "Status": activity_status(a),
                     "Days Late": (today() - to_date(a[2])).days if activity_status(a) == "Delayed" else None,
                     "Remarks": a[4]})
    df = pd.DataFrame(rows, columns=["Step", "Activity", "Plan Date", "Actual Date", "Status", "Days Late", "Remarks"])
    st.markdown('<div class="section">Activities</div>', unsafe_allow_html=True)
    if status_filter:
        st.dataframe(df[df["Status"].isin(status_filter)], hide_index=True, width="stretch",
                     column_config={"Plan Date": st.column_config.DateColumn(format="DD MMM YY"),
                                    "Actual Date": st.column_config.DateColumn(format="DD MMM YY")})
        st.caption("Clear the status filter to edit activities.")
    else:
        def comp(rec):
            a = [rec.get("step"), rec.get("activity"), rec.get("plan"), rec.get("actual"), rec.get("remarks")]
            stt = activity_status(a)
            return {"Status": stt, "Days Late": (today() - to_date(a[2])).days if stt == "Delayed" else None}
        excel_grid(spec_acts(pi), computed=comp, height=520)
        with st.expander("🔤 Rename an activity in this project or in all projects"):
            names = sorted({(a[1] or "").strip() for pp in d["projects"] for a in pp["acts"] if not blank(a[1])})
            c1, c2 = st.columns(2)
            c1.selectbox("Activity", names, key="ren_old")
            c2.text_input("New name", key="ren_new")
            st.radio("Apply to", ["This project only", "All projects"], horizontal=True, key="ren_scope")
            st.button("Rename activity", on_click=_rename_activity, args=(pi,), type="primary")

    # timeline
    tl = df.dropna(subset=["Plan Date"]).copy()
    if not tl.empty:
        st.markdown('<div class="section">Timeline — plan vs actual</div>', unsafe_allow_html=True)
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=tl["Plan Date"], y=tl["Activity"], mode="markers", name="Plan",
                                 marker=dict(symbol="circle-open", size=11, color=COLORS["navy"], line=dict(width=2))))
        act = df.dropna(subset=["Actual Date"])
        fig.add_trace(go.Scatter(x=act["Actual Date"], y=act["Activity"], mode="markers", name="Actual",
                                 marker=dict(size=9, color=COLORS["ok"])))
        dl = tl[tl["Status"] == "Delayed"]
        fig.add_trace(go.Scatter(x=dl["Plan Date"], y=dl["Activity"], mode="markers", name="Delayed",
                                 marker=dict(symbol="x", size=11, color=COLORS["bad"])))
        fig.add_shape(type="line", x0=today(), x1=today(), y0=0, y1=1, yref="paper", line=dict(dash="dot", color=COLORS["muted"]))
        fig.update_yaxes(autorange="reversed", title=None)
        st.plotly_chart(plot_layout(fig, 28 * max(len(tl), len(act)) + 140), width="stretch")

    with st.expander("✏️ Edit project name / status, change project order, or delete project"):
        with st.form(f"edit_proj_{pi}"):
            c1, c2 = st.columns([3, 1])
            nn = c1.text_input("Project / part name", value=p["name"])
            opts = STATUSES + ([p["status"]] if p["status"] and p["status"] not in STATUSES else [])
            ns = c2.selectbox("Status", opts, index=opts.index(p["status"]) if p["status"] in opts else 0)
            if st.form_submit_button("Save project", type="primary"):
                if any(i != pi and x["name"].lower() == nn.strip().lower() for i, x in enumerate(d["projects"])):
                    st.error("Another project already has this name.")
                elif not nn.strip():
                    st.error("Enter a project name.")
                else:
                    p["name"], p["status"] = nn.strip(), ns
                    st.session_state.sel_project = p["name"]
                    commit("Project updated")
                    st.rerun()
        c1, c2, _ = st.columns([1, 1, 3])
        c1.button("⬆ Move project up", key=f"pup_{pi}", on_click=_move_project, args=(pi, -1), disabled=pi == 0)
        c2.button("⬇ Move project down", key=f"pdn_{pi}", on_click=_move_project, args=(pi, 1),
                  disabled=pi == len(d["projects"]) - 1)
        conf = st.checkbox(f'I want to delete "{p["name"]}" and all its activities', key=f"del_conf_{pi}")
        if st.button("🗑 Delete project", disabled=not conf, key=f"del_{pi}"):
            d["projects"].pop(pi)
            st.session_state.pop("sel_project", None)
            commit("Project deleted")
            st.rerun()


def page_activities():
    d = D()
    header("Pending & Completed Activities", "All stage-gate activities across projects. Sort by clicking a column header.")
    adf = activities_df(d)
    if adf.empty:
        no_data_notice()
        return
    c1, c2, c3 = st.columns([2, 2, 2])
    projs = c1.multiselect("Project", sorted(adf["Project"].unique()))
    pstat = c2.multiselect("Project status", sorted(adf["Project Status"].unique()))
    q = c3.text_input("Search activity / remarks")
    f = adf.copy()
    if projs:
        f = f[f["Project"].isin(projs)]
    if pstat:
        f = f[f["Project Status"].isin(pstat)]
    if q:
        m = f["Activity"].astype(str).str.contains(q, case=False, na=False) | f["Remarks"].astype(str).str.contains(q, case=False, na=False)
        f = f[m]
    pending = f[f["Status"].isin(["Delayed", "Due in 7 days", "Open"])].sort_values(["Status", "Days Late"], ascending=[True, False])
    completed = f[f["Status"] == "Done"].sort_values("Actual Date", ascending=False)
    k = st.columns(4)
    with k[0]:
        kpi("Pending", len(pending), "open activities with a plan date")
    with k[1]:
        kpi("Delayed", int((f["Status"] == "Delayed").sum()), "plan passed, no actual", value_color=COLORS["bad"])
    with k[2]:
        kpi("Due in 7 days", int((f["Status"] == "Due in 7 days").sum()), "coming up")
    with k[3]:
        kpi("Completed", len(completed), "actual date recorded", value_color=COLORS["ok"])
    cfg = {"Plan Date": st.column_config.DateColumn(format="DD MMM YY"),
           "Actual Date": st.column_config.DateColumn(format="DD MMM YY"),
           "Days Late": st.column_config.NumberColumn(format="%d"), "Step": st.column_config.NumberColumn(format="%d")}
    t1, t2, t3 = st.tabs([f"⏳ Pending ({len(pending)})", f"✅ Completed ({len(completed)})", "📊 By project"])
    with t1:
        st.dataframe(pending[["Project", "Step", "Activity", "Plan Date", "Status", "Days Late", "Remarks"]],
                     hide_index=True, width="stretch", column_config=cfg, height=480)
    with t2:
        st.dataframe(completed[["Project", "Step", "Activity", "Plan Date", "Actual Date", "Remarks"]],
                     hide_index=True, width="stretch", column_config=cfg, height=480)
    with t3:
        g = f[f["Status"] != "Not planned"].groupby(["Project", "Status"]).size().reset_index(name="Count")
        if not g.empty:
            fig = px.bar(g, x="Count", y="Project", color="Status", orientation="h", color_discrete_map=ACT_COLORS)
            fig.update_yaxes(autorange="reversed", title=None)
            st.plotly_chart(plot_layout(fig, 40 * g["Project"].nunique() + 140), width="stretch")


def page_parts(key):
    d = D()
    cfg = PCFG[key]
    rows = d[key][1]
    m = tmeta(d, key)
    core = TABLE_DEFS[key]["core"]
    header(cfg["label"], "Search, filter and sort below. Switch on Edit mode to add, insert, move or delete rows and columns.")
    if not rows and not st.session_state.get(f"edit_{key}"):
        st.info("No parts yet. Upload your Excel in Data Management, or switch on Edit mode to add parts.")
    edit = st.toggle("✏️ Edit mode (changes save automatically)", key=f"edit_{key}")
    if edit:
        drops = {}
        if cfg["proj"] is not None:
            drops["project"] = sorted(set(PROJECT_CHOICES) | {str(r[cfg["proj"]]) for r in rows if not blank(r[cfg["proj"]])})
        excel_grid(spec_parts(key), dropdowns=drops, height=560,
                   help_text="Part No must be unique. MAKE/MPN format: MAKE/MPN.")
        return
    recs = [_row_to_rec(r, core) for r in rows]
    order = vis(m)
    labs = [m["labels"][c] for c in order]
    df = pd.DataFrame([[_to_cell(r.get(c), _ctype(m, key, c)) for c in order] for r in recs], columns=labs)
    for c, lab in zip(order, labs):
        if _ctype(m, key, c) == "number":
            df[lab] = pd.to_numeric(df[lab], errors="coerce")
    c1, c2, c3 = st.columns([3, 2, 2])
    q = c1.text_input("Search part no, name, supplier, MPN…", key=f"q_{key}")
    pf = []
    if cfg["proj"] is not None:
        pf = c2.multiselect("Project", sorted({str(r[cfg["proj"]]) for r in rows if not blank(r[cfg["proj"]])}), key=f"pf_{key}")
    sup = c3.multiselect("Supplier", sorted({str(r[cfg["supplier"]]).strip() for r in rows if not blank(r[cfg["supplier"]])}), key=f"sf_{key}")
    miss = st.checkbox("Only parts missing MPN or price", key=f"mf_{key}")
    mask = pd.Series(True, index=df.index)
    if q:
        mask &= df.astype(str).apply(lambda col: col.str.contains(q, case=False, na=False)).any(axis=1)
    if pf:
        mask &= pd.Series([str(r[cfg["proj"]]) in pf for r in rows], index=df.index)
    if sup:
        mask &= pd.Series([str(r[cfg["supplier"]]).strip() in sup for r in rows], index=df.index)
    if miss:
        mask &= pd.Series([not (has_mpn(key, r) and has_price(key, r)) for r in rows], index=df.index)
    f = df[mask]
    st.caption(f"{len(f)} of {len(df)} parts")
    mpn_pos = order.index(core[cfg["mpn"][0]]) if core[cfg["mpn"][0]] in order else None
    price_pos = order.index(core[cfg["price"]]) if core[cfg["price"]] in order else None

    def colour(row):
        r = rows[row.name]
        out = [""] * len(row)
        if mpn_pos is not None and not has_mpn(key, r):
            out[mpn_pos] = f"background-color:{COLORS['bad_soft']}"
        if price_pos is not None and not has_price(key, r):
            out[price_pos] = f"background-color:{COLORS['bad_soft']}"
        return out
    st.dataframe(f.style.apply(colour, axis=1), hide_index=True, width="stretch", height=560,
                 column_config={labs[price_pos]: st.column_config.NumberColumn(format="₹ %.2f")} if price_pos is not None else None)
    st.caption("Red cells = missing MAKE/MPN or price.")


def page_suppliers():
    d = D()
    header("Supplier-wise Tracking", "Parts, MPN/price coverage and BOM value by supplier across MQIS, Single Phase and 6DS.")
    pu = parts_union(d)
    if pu.empty:
        no_data_notice()
        return
    agg = pu.groupby("Supplier").agg(Parts=("Part No", "count"), With_MPN=("Has MPN", "sum"), With_Price=("Has Price", "sum"),
                                     BOM_Value=("Value", "sum"), Projects=("Project", lambda s: ", ".join(sorted({str(x) for x in s if x})))).reset_index()
    agg["MPN %"] = (agg["With_MPN"] / agg["Parts"] * 100).round(0)
    agg["Price %"] = (agg["With_Price"] / agg["Parts"] * 100).round(0)
    agg = agg.sort_values("Parts", ascending=False)
    k = st.columns(4)
    with k[0]:
        kpi("Suppliers", len(agg), "distinct names in parts lists")
    with k[1]:
        kpi("Top supplier", agg.iloc[0]["Supplier"], f'{int(agg.iloc[0]["Parts"])} parts')
    with k[2]:
        kpi("Suppliers with gaps", int((agg["With_MPN"] < agg["Parts"]).sum() + 0), "missing MPN on ≥1 part")
    with k[3]:
        kpi("BOM value (per set)", inr(float(pu["Value"].sum(skipna=True))), "qty × unit price, priced parts only")
    c1, c2 = st.columns([3, 2])
    with c1:
        top = agg.head(15)
        fig = px.bar(top, x="Parts", y="Supplier", orientation="h", color="Price %",
                     color_continuous_scale=[[0, COLORS["bad"]], [0.5, COLORS["warn"]], [1, COLORS["ok"]]], range_color=[0, 100],
                     hover_data=["With_MPN", "With_Price", "Projects"])
        fig.update_yaxes(autorange="reversed", title=None)
        st.plotly_chart(plot_layout(fig, 30 * len(top) + 140), width="stretch")
    with c2:
        val = agg.dropna(subset=["BOM_Value"])
        val = val[val["BOM_Value"] > 0].head(10)
        if not val.empty:
            fig = px.pie(val, names="Supplier", values="BOM_Value", hole=0.5,
                         color_discrete_sequence=[COLORS[k] for k in ("navy", "ok", "warn", "Common", "P1R5", "Single", "6DS", "bad", "idle", "muted")])
            fig.update_traces(textinfo="percent")
            st.plotly_chart(plot_layout(fig, 380), width="stretch")
    st.dataframe(agg.rename(columns={"With_MPN": "With MPN", "With_Price": "With Price", "BOM_Value": "BOM Value"}),
                 hide_index=True, width="stretch",
                 column_config={"BOM Value": st.column_config.NumberColumn(format="₹ %.2f"),
                                "MPN %": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%d%%"),
                                "Price %": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%d%%")})
    sel = st.selectbox("Show parts for supplier", ["—"] + agg["Supplier"].tolist())
    if sel != "—":
        st.dataframe(pu[pu["Supplier"] == sel][["Sheet", "Part No", "Part Name", "Project", "Qty", "Price", "MAKE/MPN", "Has MPN", "Has Price"]],
                     hide_index=True, width="stretch", column_config={"Price": st.column_config.NumberColumn(format="₹ %.2f")})


def _fan_raw_table(rows):
    rows = [r for r in rows if r and any(not blank(v) for v in r)]
    if len(rows) < 2:
        return None, None
    title = rows[0][0]
    out = []
    for r in rows[1:]:
        r = (list(r) + [None] * 4)[:4]
        lab = str(r[0] or "")
        vals = []
        for i, v in enumerate(r):
            if i == 0 or not isinstance(v, (int, float)) or isinstance(v, bool):
                vals.append("" if v is None else str(v))
            elif "%" in lab:
                vals.append(f"{v * 100:.2f}%")
            elif re.search(r"US\$|Exchange|volume|Lead time", lab, re.I):
                vals.append(f"{v:,.2f}".rstrip("0").rstrip("."))
            else:
                vals.append(inr(v))
        out.append(vals)
    hdr = out[0] if out and out[0][0] == "Cost Element" else ["Item", "Existing", "Alternate", "Note"]
    body = out[1:] if out and out[0][0] == "Cost Element" else out
    return title, pd.DataFrame(body, columns=hdr)


def _fan_form(idx=None):
    d = D()
    f = d["fanExtra"][idx] if idx is not None else {
        "fx": next((num(r[1]) for r in d["fan"] if r and re.search("Exchange", str(r[0] or ""))), None),
        "insPct": 0.13, "bcdPct": 10, "swsPct": 10, "freight": 0, "altFreight": 0, "cha": 0}
    with st.form(f"fan_form_{idx}"):
        c = st.columns(3)
        part = c[0].text_input("Part / description *", value=f.get("part") or "")
        exn = c[1].text_input("Existing supplier (import)", value=f.get("exName") or "")
        altn = c[2].text_input("Alternate supplier (domestic)", value=f.get("altName") or "")
        c = st.columns(4)
        pusd = c[0].number_input("Existing unit price (US$) *", value=num(f.get("pUSD")), min_value=0.0, format="%.4f")
        fx = c[1].number_input("Exchange rate (₹ per US$) *", value=num(f.get("fx")), min_value=0.0, format="%.2f")
        fr = c[2].number_input("Existing freight (₹/unit)", value=num(f.get("freight")) or 0.0, min_value=0.0)
        cha = c[3].number_input("CHA / clearance (₹/unit)", value=num(f.get("cha")) or 0.0, min_value=0.0)
        c = st.columns(3)
        ins = c[0].number_input("Insurance %", value=num(f.get("insPct")) or 0.0, min_value=0.0, format="%.3f")
        bcd = c[1].number_input("BCD %", value=num(f.get("bcdPct")) or 0.0, min_value=0.0)
        sws = c[2].number_input("SWS % (on BCD)", value=num(f.get("swsPct")) or 0.0, min_value=0.0)
        c = st.columns(3)
        alt = c[0].number_input("Alternate unit price (₹) *", value=num(f.get("altPrice")), min_value=0.0)
        altfr = c[1].number_input("Alternate freight (₹/unit)", value=num(f.get("altFreight")) or 0.0, min_value=0.0)
        vol = c[2].number_input("Annual volume (nos)", value=num(f.get("vol")) or 0.0, min_value=0.0, step=100.0)
        st.markdown("**Before switching — confirm**")
        vals = {}
        for lab, a, b in (("Make / MPN", "exMpn", "altMpn"), ("Lead time (wks)", "exLt", "altLt"), ("MOQ", "exMoq", "altMoq"),
                          ("Warranty", "exWar", "altWar"), ("Payment terms", "exPay", "altPay")):
            c = st.columns(2)
            vals[a] = c[0].text_input(f"Existing {lab.lower()}", value=str(f.get(a) or ""))
            vals[b] = c[1].text_input(f"Alternate {lab.lower()}", value=str(f.get(b) or ""))
        note = st.text_area("Note", value=f.get("note") or "", height=60)
        if st.form_submit_button("Save comparison", type="primary"):
            if not part.strip() or pusd is None or fx is None or alt is None:
                st.error("Fill in: part, existing US$ price, exchange rate and alternate price.")
                return
            if any(i != idx and x.get("part", "").lower() == part.strip().lower() for i, x in enumerate(d["fanExtra"])):
                st.error("A comparison for this part already exists.")
                return
            rec = {"part": part.strip(), "exName": exn, "altName": altn, "pUSD": pusd, "fx": fx, "freight": fr, "insPct": ins,
                   "bcdPct": bcd, "swsPct": sws, "cha": cha, "altPrice": alt, "altFreight": altfr, "vol": vol, "note": note, **vals}
            if idx is None:
                d["fanExtra"].append(rec)
                commit("Comparison added")
            else:
                d["fanExtra"][idx] = rec
                commit("Comparison updated")
            st.rerun()


def page_fan():
    d = D()
    header("DC Fan Comparison — Landed Cost", "Existing (import) vs alternate (domestic) supplier, per unit. "
                                              "Landed cost and savings are calculated automatically.")
    with st.expander("➕ Add comparison for another part"):
        _fan_form(None)
    n = 0
    title, raw = _fan_raw_table(d["fan"])
    if raw is not None:
        n += 1
        st.markdown(f'<div class="section">{n} · {title}</div>', unsafe_allow_html=True)
        st.dataframe(raw, hide_index=True, width="stretch", height=min(38 * len(raw) + 40, 900))
        st.caption("From the uploaded Excel sheet. To change it, update the Excel and upload it again.")
    summary = []
    for i, f in enumerate(d["fanExtra"]):
        n += 1
        c = fan_calc(f)
        summary.append({"Part": f["part"], "Existing landed": c["landed"], "Alternate landed": c["alt"], "Annual saving": c["annual"]})
        st.markdown(f'<div class="section">{n} · {f["part"]}</div>', unsafe_allow_html=True)
        k = st.columns(4)
        with k[0]:
            kpi("Existing landed / unit", inr(c["landed"]), f.get("exName") or "import")
        with k[1]:
            kpi("Alternate landed / unit", inr(c["alt"]), f.get("altName") or "domestic")
        with k[2]:
            kpi("Saving / unit", inr(c["save"]), f'{c["pct"] * 100:.2f}%', value_color=COLORS["ok"] if c["save"] >= 0 else COLORS["bad"])
        with k[3]:
            kpi("Annual saving", inr(c["annual"]), f'{int(num(f.get("vol")) or 0):,} nos/yr', value_color=COLORS["ok"] if c["annual"] >= 0 else COLORS["bad"])
        tbl = pd.DataFrame([
            ["Unit price (US$)", f"{num(f.get('pUSD')) or 0:g}", "–", ""],
            ["Exchange rate (₹ per US$)", f"{num(f.get('fx')) or 0:g}", "–", ""],
            ["Unit price (₹)", inr(c["pINR"]), inr(num(f.get("altPrice"))), ""],
            ["Freight", inr(num(f.get("freight")) or 0), inr(num(f.get("altFreight")) or 0), ""],
            ["Insurance", inr(c["ins"]), "–", f'{num(f.get("insPct")) or 0:g}% of price + freight'],
            ["CIF value", inr(c["cif"]), "–", ""],
            ["BCD", inr(c["bcd"]), "–", f'{num(f.get("bcdPct")) or 0:g}%'],
            ["SWS (on BCD)", inr(c["sws"]), "–", f'{num(f.get("swsPct")) or 0:g}%'],
            ["CHA / clearance / local freight", inr(num(f.get("cha")) or 0), "–", ""],
            ["LANDED COST PER UNIT (₹)", inr(c["landed"]), inr(c["alt"]), "IGST excluded"],
            ["Make / MPN", f.get("exMpn") or "", f.get("altMpn") or "", ""],
            ["Lead time (weeks)", f.get("exLt") or "", f.get("altLt") or "", ""],
            ["MOQ", f.get("exMoq") or "", f.get("altMoq") or "", ""],
            ["Warranty", f.get("exWar") or "", f.get("altWar") or "", ""],
            ["Payment terms", f.get("exPay") or "", f.get("altPay") or "", ""],
        ], columns=["Cost element", "Existing (import)", "Alternate (domestic)", "Basis / note"])
        st.dataframe(tbl, hide_index=True, width="stretch")
        if f.get("note"):
            st.caption(f"Note: {f['note']}")
        with st.expander(f"✏️ Edit or delete — {f['part']}"):
            _fan_form(i)
            conf = st.checkbox("Confirm delete", key=f"fan_del_conf_{i}")
            if st.button("🗑 Delete comparison", disabled=not conf, key=f"fan_del_{i}"):
                d["fanExtra"].pop(i)
                commit("Comparison deleted")
                st.rerun()
    if n == 0:
        st.info("No comparisons yet. Upload your Excel (Data Management) or add a comparison above.")
    if len(summary) > 1:
        fig = px.bar(pd.DataFrame(summary), x="Part", y=["Existing landed", "Alternate landed"], barmode="group",
                     color_discrete_sequence=[COLORS["navy"], COLORS["ok"]])
        st.plotly_chart(plot_layout(fig, 340), width="stretch")


def page_teknic():
    d = D()
    t = d["tk"]
    header("Teknic Parts (Domestic)", f"{t.get('title') or ''} · Savings and totals are calculated automatically.")
    c = teknic_calc(t)
    k = st.columns(4)
    with k[0]:
        kpi("Existing cost / station", inr(c["te"]))
    with k[1]:
        kpi("Quoted cost / station", inr(c["tq"]))
    with k[2]:
        kpi("Saving / station", inr(c["ts"]), f'{c["pct"] * 100:.2f}%', value_color=COLORS["ok"] if c["ts"] >= 0 else COLORS["bad"])
    with k[3]:
        kpi("Annual saving", inr(c["annual"]) if c["annual"] is not None else "Enter stations",
            f'{int(num(t.get("stations")) or 0):,} stations / year', value_color=COLORS["ok"])
    with st.form("stations"):
        cc = st.columns([1, 3])
        stn = cc[0].number_input("Stations per year", value=num(t.get("stations")), min_value=0.0, step=100.0)
        if cc[1].form_submit_button("Update stations"):
            t["stations"] = stn
            commit("Station volume updated")
            st.rerun()
    st.markdown('<div class="section">Parts — edit inputs here (changes save automatically)</div>', unsafe_allow_html=True)
    excel_grid(spec_teknic(), height=300)
    c = teknic_calc(D()["tk"])
    if c["rows"]:
        st.markdown('<div class="section">Savings (calculated)</div>', unsafe_allow_html=True)
        out = pd.DataFrame([{"Part No.": p["pn"], "Part Name": p["name"], "BOM / Stn": p.get("bom"), "Existing ₹": p.get("exp"),
                             "Quoted ₹": p.get("qp"), "Saving / unit": p["save"], "Saving %": p["pct"] * 100,
                             "Saving / station": p["save_stn"]} for p in c["rows"]])
        out.loc[len(out)] = {"Part No.": "TOTAL PER STATION", "Part Name": "", "BOM / Stn": None, "Existing ₹": c["te"],
                             "Quoted ₹": c["tq"], "Saving / unit": None, "Saving %": c["pct"] * 100, "Saving / station": c["ts"]}
        out["Part No."] = out["Part No."].astype(str)
        money = st.column_config.NumberColumn(format="₹ %.2f")
        st.dataframe(out, hide_index=True, width="stretch",
                     column_config={"Existing ₹": money, "Quoted ₹": money, "Saving / unit": money, "Saving / station": money,
                                    "Saving %": st.column_config.NumberColumn(format="%.2f%%")})


def page_data():
    ss = st.session_state
    header("Data Management", "Load your data, then download the updated Excel or a backup when you finish.")
    if gh_cfg() and ss.get("gh_status", ("", ""))[0] == "ok":
        st.success("Permanent saving is ON: every Save writes to GitHub. Downloads below are for Excel copies and backups.")
    else:
        st.warning("**Permanent saving is OFF.** Streamlit Cloud clears memory on refresh/restart. Set up GitHub saving "
                   "(README) or download the updated Excel / JSON backup before you close the browser.", icon="⚠️")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section">1 · Load</div>', unsafe_allow_html=True)
        up = st.file_uploader("NPD Master Tracker Excel (.xlsx)", type=["xlsx", "xlsm"], key="up_xlsx")
        if up is not None and st.button("Load Excel", type="primary"):
            try:
                b = up.getvalue()
                ss.data = parse_workbook(b, ss.data)
                ss.orig_bytes, ss.orig_name = b, up.name
                ss.source = up.name
                commit(f"Loaded {up.name}: {len(ss.data['projects'])} projects, "
                       f"{len(ss.data['mqis'][1]) + len(ss.data['ds6'][1])} parts")
                st.rerun()
            except Exception as e:  # noqa: BLE001
                st.error(f"Could not read that file: {e}")
        upj = st.file_uploader("…or restore a JSON backup (from this app or the old HTML tracker)", type=["json"], key="up_json")
        if upj is not None and st.button("Restore backup"):
            try:
                ss.data = parse_backup(upj.getvalue())
                ss.source = upj.name
                commit("Backup restored")
                st.rerun()
            except Exception as e:  # noqa: BLE001
                st.error(f"Could not restore: {e}")
        if ss.orig_name:
            st.caption(f"Master workbook in memory: **{ss.orig_name}** (used for the formatted download).")
    with c2:
        st.markdown('<div class="section">2 · Download</div>', unsafe_allow_html=True)
        stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
        if not has_data():
            st.info("Load data first.")
        else:
            if ss.orig_bytes:
                try:
                    st.download_button("⬇ Updated master Excel (keeps your formatting)", formatted_export(ss.orig_bytes, D()),
                                       file_name=f"NPD_Master_Tracker_{stamp}.xlsx", type="primary", width="stretch",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.caption("Writes your changes into the workbook you uploaded: parts sheets, Alternate Development "
                               "(status table + project details), Teknic sheet, and one sheet per added DC fan comparison. "
                               "Other sheets are untouched. Formulas recalculate when opened in Excel.")
                except Exception as e:  # noqa: BLE001
                    st.error(f"Formatted export failed ({e}). Use the clean export below.")
            else:
                st.caption("Upload the master Excel to enable the formatted download.")
            st.download_button("⬇ Clean data export (.xlsx)", clean_export(D()), file_name=f"NPD_Tracker_Export_{stamp}.xlsx",
                               width="stretch", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.download_button("⬇ JSON backup", backup_bytes(D()), file_name=f"NPD_Tracker_Backup_{stamp}.json",
                               width="stretch", mime="application/json")
            st.download_button("⬇ npd_data.json (make changes permanent)", backup_bytes(D()), file_name="npd_data.json",
                               width="stretch", mime="application/json")
            st.caption("To keep today's changes for everyone: download **npd_data.json** and upload it to the GitHub "
                       "repository, replacing the old one. The app reloads with the new data in about a minute.")
    st.markdown('<div class="section">3 · Current data</div>', unsafe_allow_html=True)
    d = D()
    st.dataframe(pd.DataFrame([
        ["Projects (Alternate Development)", len(d["projects"]), sum(len(p["acts"]) for p in d["projects"])],
        ["MQIS & Single Phase parts", len(d["mqis"][1]), ""],
        ["6DS project parts", len(d["ds6"][1]), ""],
        ["DC fan comparisons", (1 if _fan_raw_table(d["fan"])[1] is not None else 0) + len(d["fanExtra"]), ""],
        ["Teknic parts", len(d["tk"]["parts"]), ""],
    ], columns=["Section", "Records", "Activities"]).astype(str), hide_index=True, width="stretch")
    if ss.get("last_change"):
        st.caption(f"Last change in this session: {ss.last_change:%d %b %Y %H:%M}. Download before closing.")
    with st.expander("Clear all data from this session"):
        conf = st.checkbox("I have downloaded what I need — clear everything")
        if st.button("Clear data", disabled=not conf):
            for k in ("data", "source", "orig_bytes", "orig_name", "changed", "last_change"):
                ss.pop(k, None)
            init_state()
            st.rerun()


# ----------------------------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------------------------
PAGES = {
    "📊 Dashboard": page_dashboard,
    "🗂 Project Tracking": page_projects,
    "⏳ Pending & Completed": page_activities,
    "🔩 MQIS & Single Phase Parts": lambda: page_parts("mqis"),
    "🔧 6DS Project Parts": lambda: page_parts("ds6"),
    "🏭 Supplier-wise": page_suppliers,
    "🌀 DC Fan Comparison": page_fan,
    "⚡ Teknic Parts (Domestic)": page_teknic,
    "💾 Data Management": page_data,
}


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="⚙️", layout="wide", initial_sidebar_state="expanded")
    st.markdown(CSS, unsafe_allow_html=True)
    init_state()
    ss = st.session_state
    with st.sidebar:
        st.markdown(f"### ⚙️ {APP_TITLE}")
        st.caption("E&E BOP · Alternate development · Sun Mobility CPMO")
        page = st.radio("Navigate", list(PAGES), key="nav", label_visibility="collapsed")
        st.divider()
        st.caption(f"**Data:** {ss.source}")
        gs = ss.get("gh_status")
        if gs:
            st.caption(("🟢 " if gs[0] == "ok" else "🔴 ") + gs[1])
        elif not AUTOSAVE_FILE:
            st.caption("⚪ Permanent saving OFF — changes are lost on refresh. See README: GitHub saving.")
        if ss.get("changed"):
            st.caption("🟠 Unsaved changes — download from Data Management")
        if ss.get("load_error"):
            st.caption(f"⚠️ {ss.load_error}")
        if has_data():
            st.download_button("⬇ Quick JSON backup", backup_bytes(D()),
                               file_name=f"NPD_Tracker_Backup_{datetime.now():%Y-%m-%d_%H%M}.json",
                               mime="application/json", width="stretch")
    flash()
    PAGES[page]()


if __name__ == "__main__":
    main()
