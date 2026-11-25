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
import math
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
APP_TITLE = "NPD & Sourcing Management"
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
            "stdActs": DEFAULT_STAGE_GATE[:], "suppliers": [], "rfq": [], "pos": [], "inventory": [], "dv": [], "ecn": [], "pbom": [],
            "emails": [], "email_log": [], "settings": dict(DEFAULT_SETTINGS)}


DEFAULT_SETTINGS = {"company": "Sun Mobility", "team": "NPD Team", "sender_name": "", "sender_email": "", "default_cc": "",
                    "signature": "", "ack_days": 2, "rfq_due_days": 7, "horizon_weeks": 12, "ai_model": "gpt-4o-mini"}
COLLECTIONS = ["suppliers", "rfq", "pos", "inventory", "dv", "ecn", "pbom"]


def new_id() -> str:
    import uuid
    return uuid.uuid4().hex[:10]


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
    for c in COLLECTIONS + ["emails", "email_log"]:
        if not isinstance(d.get(c), list):
            d[c] = []
    for c in COLLECTIONS:
        for r in d[c]:
            if blank(r.get("_id")):
                r["_id"] = new_id()
    st_ = d.get("settings") if isinstance(d.get("settings"), dict) else {}
    d["settings"] = {**DEFAULT_SETTINGS, **st_}
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
    return bool(d["projects"] or d["mqis"][1] or d["ds6"][1] or d["tk"]["parts"] or d["fan"] or d["fanExtra"]
                or any(d.get(c) for c in COLLECTIONS))


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
    info = st.session_state.pop("flash_info", None)
    if info:
        st.info(info, icon="✉️")


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
SUPPLIER_STATUS = ["Approved", "Under development", "On hold", "Blocked"]
RFQ_STATUS = ["Draft", "Sent", "Quotation pending", "Clarification pending", "Quotation received", "Closed"]
PO_APPROVAL = ["Draft", "Pending approval", "Approved", "Rejected", "Cancelled"]
PO_ACK = ["Pending", "Acknowledged", "Not required"]
DV_STATUS = ["Not started", "In progress", "Passed", "Failed", "On hold"]
ECN_STATUS = ["Open", "Under review", "Approved", "Implemented", "Closed", "Rejected"]
TABLE_DEFS.update({
    "suppliers": {"core": ["_id", "name", "code", "category", "contact", "email", "cc", "phone", "location", "status", "rating", "remarks"],
                  "labels": ["ID", "Supplier Name", "Vendor Code", "Category", "Contact Person", "Email", "CC Emails", "Phone",
                             "Location", "Status", "Rating (1-5)", "Remarks"],
                  "types": {"rating": "number"}, "renumber": None, "label_col": "name", "unique": ["name"],
                  "title": "suppliers", "coll": "suppliers"},
    "rfq": {"core": ["_id", "rfq_no", "rfq_date", "project", "part_no", "description", "supplier", "sent_date", "quote_due",
                     "quote_received", "status", "remarks"],
            "labels": ["ID", "RFQ No", "RFQ Date", "Project", "Part No", "Description", "Supplier", "Sent Date", "Quote Due",
                       "Quote Received", "Status", "Remarks"],
            "types": {"rfq_date": "date", "sent_date": "date", "quote_due": "date", "quote_received": "date"},
            "renumber": None, "label_col": "rfq_no", "unique": ["rfq_no", "supplier", "part_no"], "title": "RFQs", "coll": "rfq"},
    "pos": {"core": ["_id", "po_no", "po_date", "project", "part_no", "description", "supplier", "supplier_email", "mpn", "qty",
                     "unit_price", "payment_terms", "delivery_address", "attachment", "approval", "ack", "promised_date",
                     "actual_date", "received_qty", "remarks"],
            "labels": ["ID", "PO Number", "PO Date", "Project", "Part Number", "Part Description", "Supplier Name",
                       "Supplier Email", "MPN / Make", "PO Qty", "Unit Price", "Payment Terms", "Delivery Address",
                       "PO Attachment", "PO Approval Status", "Supplier Acknowledgement", "Promised Delivery Date",
                       "Actual Delivery Date", "Received Qty", "Remarks"],
            "types": {"po_date": "date", "qty": "number", "unit_price": "number", "promised_date": "date", "actual_date": "date",
                      "received_qty": "number"},
            "renumber": None, "label_col": "po_no", "unique": ["po_no", "part_no"], "title": "purchase orders", "coll": "pos"},
    "inventory": {"core": ["_id", "part_no", "part_name", "stock", "monthly", "open_po_qty", "lead_time", "confirmed_date", "safety",
                           "project_req", "remarks"],
                  "labels": ["ID", "Part Number", "Part Name", "Current Stock", "Monthly Consumption", "Open PO Qty (manual)",
                             "Supplier Lead Time (days)", "Confirmed Delivery Date", "Safety Stock", "Project Requirement", "Remarks"],
                  "types": {"stock": "number", "monthly": "number", "open_po_qty": "number", "lead_time": "number",
                            "confirmed_date": "date", "safety": "number", "project_req": "number"},
                  "renumber": None, "label_col": "part_name", "unique": ["part_no"], "title": "inventory", "coll": "inventory"},
    "dv": {"core": ["_id", "project", "part_no", "supplier", "samples_handed", "dv_plan", "dv_actual", "dv_result", "dvvp_plan",
                    "dvvp_actual", "dvvp_result", "report", "status", "remarks"],
           "labels": ["ID", "Project", "Part No", "Supplier", "Samples Handed to Engg", "DV Plan", "DV Actual", "DV Result",
                      "DVVP Plan", "DVVP Actual", "DVVP Result", "Report Ref", "Status", "Remarks"],
           "types": {"samples_handed": "date", "dv_plan": "date", "dv_actual": "date", "dvvp_plan": "date", "dvvp_actual": "date"},
           "renumber": None, "label_col": "project", "unique": ["project", "part_no", "supplier"], "title": "DV/DVVP records",
           "coll": "dv"},
    "ecn": {"core": ["_id", "ecn_no", "ecn_date", "type", "project", "part_no", "description", "from_make", "to_make", "reason",
                     "approver", "status", "impl_date", "remarks"],
            "labels": ["ID", "ECN/ECR No", "Date", "Type", "Project", "Part No", "Change Description", "From (Make/MPN)",
                       "To (Make/MPN)", "Reason", "Approver", "Status", "Implementation Date", "Remarks"],
            "types": {"ecn_date": "date", "impl_date": "date"}, "renumber": None, "label_col": "ecn_no", "unique": ["ecn_no"],
            "title": "ECN/ECR records", "coll": "ecn"},
})
TABLE_DEFS["pbom"] = {
    "core": ["_id", "sl", "part_no", "description", "p1r5_qty", "sp_qty", "uom", "buyer", "supplier", "project", "insp",
             "store_stock", "wip", "remarks"],
    "labels": ["ID", "Sl", "Part No", "Description", "P1R5 Qty", "Sing Phs Qty", "Uom", "Buyer", "Supplier", "Project", "Insp",
               "Store Stock", "WIP", "Remarks"],
    "types": {"sl": "number", "p1r5_qty": "number", "sp_qty": "number", "insp": "number", "store_stock": "number", "wip": "number"},
    "renumber": None, "label_col": "part_no", "unique": ["part_no"], "title": "PBOM parts", "coll": "pbom"}
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
    if "_id" in td["core"] and "_id" not in m["hidden"]:
        m["hidden"].append("_id")
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


def spec_coll(t, after=None):
    coll = TABLE_DEFS[t]["coll"]

    def get():
        return [dict(r) for r in D()[coll]]

    def put(recs):
        out = []
        for r in recs:
            r = {k: v for k, v in r.items() if not blank(v) or k in TABLE_DEFS[t]["core"]}
            if blank(r.get("_id")):
                r["_id"] = new_id()
            out.append(r)
        D()[coll] = out
    return {"t": t, "key": t, "get": get, "set": put, "after": after}


def _bump(spec):
    k = f"ver_{spec['key']}"
    st.session_state[k] = st.session_state.get(k, 0) + 1


def _validate(t, recs):
    u = TABLE_DEFS[t]["unique"]
    if not u:
        return None
    if isinstance(u, list):
        seen = {}
        labels = tmeta(D(), t)["labels"]
        for i, r in enumerate(recs):
            key = tuple(str(r.get(f)).strip().lower() if not blank(r.get(f)) else "" for f in u)
            if not key[0]:
                continue
            if key in seen:
                what = " + ".join(f"{labels.get(f, f)} '{r.get(f)}'" for f in u if not blank(r.get(f)))
                return f"Not saved: duplicate {TABLE_DEFS[t]['title']} — {what} is already in row {seen[key] + 1}."
            seen[key] = i
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
   tableHeight:a.height+"px",tableWidth:Math.max(document.documentElement.clientWidth||0,window.innerWidth||0,320)-4+"px",rowDrag:true,columnDrag:true,columnSorting:false,
   allowInsertRow:true,allowDeleteRow:true,allowInsertColumn:true,allowDeleteColumn:true,allowRenameColumn:true,
   allowManualInsertRow:true,allowManualInsertColumn:true,wordWrap:false,
   onselection:(el,x1,y1,x2,y2)=>{lastSel={x1:+Math.min(x1,x2),y1:+Math.min(y1,y2),x2:+Math.max(x1,x2),y2:+Math.max(y1,y2)};},
   onbeforedeletecolumn:(el,col,num)=>{for(let i=col;i<col+num;i++){if(isLocked(i)){status("Built-in columns cannot be deleted. You can rename or move them.",true);return false;}}},
   onbeforechange:(el,cell,x,y,v)=>{if(isLocked(+x)&&String(table.options.columns[+x].name).startsWith("__comp_"))return false;},
   onchange:changed,oninsertrow:changed,ondeleterow:changed,oninsertcolumn:changed,ondeletecolumn:changed,onmoverow:changed,
   onmovecolumn:changed,onchangeheader:changed,onundo:changed,onredo:changed,onpaste:changed,onsort:changed});
  $("b-ren").style.display=a.renumber?"":"none";
  [50,300,1000,2500].forEach(t=>setTimeout(height,t));}
function fitWidth(){const c=document.querySelector(".jexcel_content");if(c){c.style.width=(Math.max(window.innerWidth,320)-4)+"px";}height();}
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
window.addEventListener("resize",fitWidth);
setTimeout(()=>{if(typeof jspreadsheet==="undefined"){status("Spreadsheet library could not load. Check internet / firewall, then reload.",true);height();}},6000);
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
               "width": 70 if cid in ("slno", "step") else 300 if cid in ("activity", "remarks", "description") else
               220 if cid in ("partname", "name", "supplier", "vendor", "mpn1", "mpn", "exs", "alts", "part_name",
                              "delivery_address", "supplier_email", "email", "cc") else
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
    old = spec["get"]()
    spec["set"](recs)
    note = ""
    if spec.get("after"):
        note = spec["after"](old, spec["get"]()) or ""
    commit(f"{td['title'].capitalize()} saved" + (f" · {note}" if note else ""))
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
    with st.expander("Grid not showing? View the data as a plain table"):
        plain = pd.DataFrame([row[:len(columns)] for row in data], columns=[c["title"] for c in columns]) if data else pd.DataFrame(
            columns=[c["title"] for c in columns])
        plain = plain[[c["title"] for c in columns if c.get("type") != "hidden"]]
        st.dataframe(plain.astype(str).replace({"": None}), hide_index=True, width="stretch")
        st.caption("The editable grid loads from the internet (cdn.jsdelivr.net). If it stays blank, a firewall or "
                   "browser extension may be blocking it.")
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
    procurement_overview()
    st.markdown('<div class="section">NPD alternate development</div>', unsafe_allow_html=True)
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


# ==============================================================================================
# PROCUREMENT MODULES: suppliers, RFQ, PO, email approval, follow-ups, inventory, DV/DVVP, ECN, AI, reports
# ==============================================================================================
EMAIL_RE = re.compile(r"^[^@\s,;<>]+@[^@\s,;<>]+\.[^@\s,;<>]+$")


def S_() -> dict:
    return D()["settings"]


def split_emails(txt) -> list:
    return [e.strip() for e in re.split(r"[,;\s]+", str(txt or "")) if e.strip()]


def bad_emails(txt) -> list:
    return [e for e in split_emails(txt) if not EMAIL_RE.match(e)]


def supplier_rec(name):
    k = _nk(name)
    return next((s for s in D()["suppliers"] if _nk(s.get("name")) == k and k), None)


def supplier_options():
    return sorted({s["name"] for s in D()["suppliers"] if not blank(s.get("name"))})


def project_options():
    return sorted({p["name"] for p in D()["projects"]})


# ---------------------------------------------------------------------------------------------- PO calculations
def po_calc(r: dict) -> dict:
    qty = num(r.get("qty")) or 0
    price = num(r.get("unit_price")) or 0
    rec = num(r.get("received_qty")) or 0
    bal = max(qty - rec, 0)
    prom, act = to_date(r.get("promised_date")), to_date(r.get("actual_date"))
    appr = (r.get("approval") or "Draft").strip()
    ack = (r.get("ack") or "Pending").strip()
    delay = None
    if prom and act:
        delay = (act - prom).days
    elif prom and bal > 0 and today() > prom:
        delay = (today() - prom).days
    if appr in ("Rejected", "Cancelled"):
        status = appr
    elif appr != "Approved":
        status = "Awaiting approval"
    elif qty and rec >= qty:
        status = "Closed – received"
    elif rec > 0:
        status = "Partial – overdue" if prom and today() > prom else "Partially received"
    elif ack not in ("Acknowledged", "Not required"):
        status = "Ack pending"
    elif not prom:
        status = "ETA pending"
    elif today() > prom:
        status = "Overdue"
    else:
        status = "Open – on schedule"
    return {"Total Value": round(qty * price, 2) if qty and price else None, "Balance Qty": bal if qty else None,
            "Delay (days)": delay, "PO Status": status}


def po_lines():
    return [{**r, **po_calc(r)} for r in D()["pos"]]


def po_groups():
    """PO number -> list of lines."""
    g = {}
    for r in D()["pos"]:
        if not blank(r.get("po_no")):
            g.setdefault(str(r["po_no"]).strip(), []).append(r)
    return g


# ---------------------------------------------------------------------------------------------- attachments
def att_save(name: str, data: bytes) -> str:
    """Store a file permanently (GitHub data branch) or for this session only. Returns where it was stored."""
    import base64
    cfg = gh_cfg()
    if cfg:
        path = f"attachments/{name}"
        g = _gh("GET", f"/contents/{path}", cfg, params={"ref": cfg["branch"]})
        body = {"message": f"NPD tracker: attachment {name}", "branch": cfg["branch"],
                "content": base64.b64encode(data).decode()}
        if g.status_code == 200:
            body["sha"] = g.json()["sha"]
        r = _gh("PUT", f"/contents/{path}", cfg, json=body)
        if r.status_code not in (200, 201):
            raise RuntimeError(f"GitHub {r.status_code}: {r.json().get('message', '')[:150]}")
        return "github"
    st.session_state.setdefault("att_mem", {})[name] = data
    return "session"


def att_load(name: str):
    import base64
    mem = st.session_state.get("att_mem", {})
    if name in mem:
        return mem[name]
    cfg = gh_cfg()
    if cfg and name:
        r = _gh("GET", f"/contents/attachments/{name}", cfg, params={"ref": cfg["branch"]})
        if r.status_code == 200:
            j = r.json()
            if j.get("content"):
                return base64.b64decode(j["content"])
            b = _gh("GET", f"/git/blobs/{j['sha']}", cfg)
            if b.status_code == 200:
                return base64.b64decode(b.json()["content"])
    return None


# ---------------------------------------------------------------------------------------------- email drafts
def _log_email(e, event, detail=""):
    D()["email_log"].append({"at": datetime.now().isoformat(timespec="seconds"), "event": event, "email_id": e["id"],
                             "type": e.get("type"), "ref": e.get("ref"), "supplier": e.get("supplier"), "to": e.get("to"),
                             "cc": e.get("cc"), "subject": e.get("subject"), "detail": detail})


def make_draft(kind, ref, supplier, to, cc, subject, body, attachments=None, source="system"):
    e = {"id": new_id(), "type": kind, "ref": ref, "supplier": supplier, "to": to or "", "cc": cc or "", "subject": subject,
         "body": body, "attachments": attachments or [], "status": "Draft", "created": datetime.now().isoformat(timespec="seconds"),
         "updated": None, "sent_at": None, "method": None, "error": None, "source": source}
    D()["emails"].insert(0, e)
    _log_email(e, "Draft created", source)
    return e


def _sig_block():
    s = S_()
    lines = ["Regards,", s.get("sender_name") or s.get("team") or "NPD Team"]
    if s.get("sender_name") and s.get("team"):
        lines.append(s["team"])
    lines.append(s.get("company") or "Sun Mobility")
    if s.get("signature"):
        lines.append(s["signature"])
    return "\n".join(lines)


def po_email(po_no, lines):
    first = lines[0]
    sup = first.get("supplier") or "Supplier"
    srec = supplier_rec(sup) or {}
    to = first.get("supplier_email") or srec.get("email") or ""
    cc = ", ".join(x for x in [srec.get("cc"), S_().get("default_cc")] if not blank(x))
    att = sorted({str(r["attachment"]).strip() for r in lines if not blank(r.get("attachment"))})
    greet = f"Dear {srec.get('contact')}," if srec.get("contact") else f"Dear {sup} Team,"
    po_date = fmt_d(first.get("po_date"))
    rows = "\n".join(f"  {i + 1}. {r.get('part_no') or ''} – {r.get('description') or ''} | Qty: {r.get('qty') or ''}"
                     + (f" | Unit price: {inr(num(r.get('unit_price')))}" if num(r.get('unit_price')) else "")
                     for i, r in enumerate(lines))
    intro = (f"Please find attached Purchase Order {po_no}" + (f" dated {po_date}" if po_date else "") + " for your reference."
             if att else f"Purchase Order {po_no}" + (f" dated {po_date}" if po_date else "") +
             " has been released to you. The PO copy will follow separately.")
    extra = []
    if not blank(first.get("payment_terms")):
        extra.append(f"Payment terms: {first['payment_terms']}")
    if not blank(first.get("delivery_address")):
        extra.append(f"Delivery address: {first['delivery_address']}")
    body = (f"{greet}\n\nGreetings from {S_().get('company') or 'Sun Mobility'}!\n\n{intro}\n\nPO details:\n{rows}\n"
            + ("\n" + "\n".join(extra) + "\n" if extra else "") +
            "\nKindly acknowledge the order and confirm the expected delivery date (ETA) at the earliest.\n\n"
            "Please ensure the material is delivered as per the agreed specifications and commercial terms.\n\n" + _sig_block())
    subject = f"Purchase Order {po_no} – Acknowledgement and Delivery Confirmation"
    return to, cc, subject, body, att


def po_after_save(old, new):
    """When a PO line turns APPROVED: create ONE email draft per PO number. Never sends."""
    was = {(str(r.get("po_no")).strip(), str(r.get("part_no")).strip()): (r.get("approval") or "") for r in old}
    newly = sorted({str(r.get("po_no")).strip() for r in new
                    if (r.get("approval") or "") == "Approved" and not blank(r.get("po_no"))
                    and was.get((str(r.get("po_no")).strip(), str(r.get("part_no")).strip())) != "Approved"})
    made = []
    groups = po_groups()
    for po in newly:
        if any(e["type"] == "PO approval" and e["ref"] == po and e["status"] in ("Draft", "Approved – send from Outlook", "Sent")
               for e in D()["emails"]):
            continue
        to, cc, subject, body, att = po_email(po, groups.get(po, []))
        make_draft("PO approval", po, groups[po][0].get("supplier"), to, cc, subject, body, att, source="PO approved")
        made.append(po)
    if made:
        st.session_state.flash_info = (f"Email draft created for PO {', '.join(made)} in the Email Approval Centre. "
                                       "Nothing has been sent.")
        return f"email draft created for {', '.join(made)}"
    return ""


def build_eml(e) -> bytes:
    from email.message import EmailMessage
    from email.utils import make_msgid, formataddr
    msg = EmailMessage()
    s = S_()
    if s.get("sender_email"):
        msg["From"] = formataddr((s.get("sender_name") or "", s["sender_email"]))
    msg["To"] = ", ".join(split_emails(e["to"]))
    if split_emails(e["cc"]):
        msg["Cc"] = ", ".join(split_emails(e["cc"]))
    msg["Subject"] = e["subject"]
    msg["Message-ID"] = make_msgid(domain="npd-tracker")
    msg["X-Unsent"] = "1"  # Outlook opens it as a draft ready to send
    msg.set_content(e["body"])
    for name in e.get("attachments") or []:
        data = att_load(name)
        if data:
            sub = "pdf" if name.lower().endswith(".pdf") else "octet-stream"
            msg.add_attachment(data, maintype="application", subtype=sub, filename=name)
    return bytes(msg)


def smtp_cfg():
    try:
        c = st.secrets.get("smtp", None)
    except Exception:  # noqa: BLE001
        return None
    if not c or not c.get("host") or not c.get("username") or not c.get("password"):
        return None
    return {"host": c["host"], "port": int(c.get("port", 587)), "username": c["username"], "password": c["password"],
            "from_addr": c.get("from_addr", c["username"]), "ssl": bool(c.get("ssl", False))}


def smtp_send(e):
    import smtplib
    from email import message_from_bytes
    cfg = smtp_cfg()
    raw = build_eml(e)
    msg = message_from_bytes(raw)
    del msg["X-Unsent"]
    del msg["From"]
    msg["From"] = cfg["from_addr"]
    rcpt = split_emails(e["to"]) + split_emails(e["cc"])
    if cfg["ssl"]:
        srv = smtplib.SMTP_SSL(cfg["host"], cfg["port"], timeout=30)
    else:
        srv = smtplib.SMTP(cfg["host"], cfg["port"], timeout=30)
        srv.starttls()
    try:
        srv.login(cfg["username"], cfg["password"])
        srv.sendmail(cfg["from_addr"], rcpt, msg.as_bytes())
    finally:
        try:
            srv.quit()
        except Exception:  # noqa: BLE001
            pass


# ---------------------------------------------------------------------------------------------- follow-ups
FOLLOW_TYPES = ["RFQ pending", "Quotation pending", "Technical clarification pending", "Sample ETA pending",
                "PO acknowledgement pending", "Supplier ETA pending", "Delivery confirmation pending", "Overdue material"]


def followup_items():
    d = D()
    out = []
    s = S_()
    ack_days = int(num(s.get("ack_days")) or 2)
    for r in d["rfq"]:
        stt = (r.get("status") or "").strip()
        sup = r.get("supplier")
        ref = r.get("rfq_no") or ""
        due = to_date(r.get("quote_due"))
        base = {"supplier": sup, "ref": f"RFQ {ref}", "part": r.get("part_no"), "project": r.get("project"),
                "desc": r.get("description")}
        if stt == "Draft":
            created = to_date(r.get("rfq_date"))
            out.append({**base, "type": "RFQ pending", "detail": "RFQ prepared but not sent", "days": (today() - created).days
                        if created else None, "internal": True})
        elif stt in ("Sent", "Quotation pending") and not to_date(r.get("quote_received")):
            late = (today() - due).days if due else None
            out.append({**base, "type": "Quotation pending", "detail": f"Quote due {fmt_d(due)}" if due else "No due date",
                        "days": late if late and late > 0 else 0})
        elif stt == "Clarification pending":
            out.append({**base, "type": "Technical clarification pending", "detail": r.get("remarks") or "Clarification awaited",
                        "days": None})
    for r in po_lines():
        if (r.get("approval") or "") != "Approved" or r["PO Status"].startswith("Closed"):
            continue
        base = {"supplier": r.get("supplier"), "ref": f"PO {r.get('po_no')}", "part": r.get("part_no"),
                "project": r.get("project"), "desc": r.get("description")}
        pod = to_date(r.get("po_date"))
        age = (today() - pod).days if pod else None
        if r["PO Status"] == "Ack pending" and (age is None or age >= ack_days):
            out.append({**base, "type": "PO acknowledgement pending", "detail": f"PO dated {fmt_d(pod)}" if pod else "",
                        "days": age})
        elif r["PO Status"] == "ETA pending":
            out.append({**base, "type": "Supplier ETA pending", "detail": "Acknowledged, no delivery date", "days": age})
        elif r["PO Status"] in ("Overdue", "Partial – overdue"):
            out.append({**base, "type": "Overdue material",
                        "detail": f"Promised {fmt_d(r.get('promised_date'))}, balance {r['Balance Qty']}",
                        "days": r["Delay (days)"]})
        else:
            prom = to_date(r.get("promised_date"))
            if prom and 0 <= (prom - today()).days <= 3:
                out.append({**base, "type": "Delivery confirmation pending", "detail": f"Due {fmt_d(prom)}", "days": 0})
    for p in d["projects"]:
        if p.get("status") != "Ongoing":
            continue
        for a in p["acts"]:
            if re.search(r"sample", str(a[1] or ""), re.I) and activity_status(a) == "Delayed":
                out.append({"supplier": None, "ref": f"NPD {p['name']}", "part": None, "project": p["name"],
                            "desc": a[1], "type": "Sample ETA pending", "detail": f"Plan {fmt_d(a[2])}",
                            "days": (today() - to_date(a[2])).days})
    for it in out:
        srec = supplier_rec(it.get("supplier")) if it.get("supplier") else None
        it["email"] = (srec or {}).get("email") or next((x.get("supplier_email") for x in d["pos"]
                                                         if _nk(x.get("supplier")) == _nk(it.get("supplier"))
                                                         and not blank(x.get("supplier_email"))), "") if it.get("supplier") else ""
    return out


FOLLOW_TEXT = {
    "RFQ pending": "our RFQ is ready to be shared with you",
    "Quotation pending": "we are yet to receive your quotation for the RFQ below",
    "Technical clarification pending": "the technical clarifications below are still open from your side",
    "Sample ETA pending": "we are awaiting confirmation of the sample dispatch / ETA",
    "PO acknowledgement pending": "we have not yet received your acknowledgement of the Purchase Order below",
    "Supplier ETA pending": "please confirm the delivery date (ETA) for the Purchase Order below",
    "Delivery confirmation pending": "please confirm that the material below will be delivered as scheduled",
    "Overdue material": "the material below is overdue against the committed delivery date",
}


def followup_email(supplier, items, escalate=False):
    srec = supplier_rec(supplier) or {}
    to = srec.get("email") or next((i["email"] for i in items if i.get("email")), "")
    cc = ", ".join(x for x in [srec.get("cc"), S_().get("default_cc")] if not blank(x))
    greet = f"Dear {srec.get('contact')}," if srec.get("contact") else f"Dear {supplier or 'Supplier'} Team,"
    kinds = sorted({i["type"] for i in items})
    lines = []
    for i, it in enumerate(items, 1):
        bits = [it["ref"], it.get("part") or "", it.get("desc") or "", it.get("detail") or ""]
        if it.get("days"):
            bits.append(f"{it['days']} day(s) pending")
        lines.append(f"  {i}. " + " | ".join(b for b in bits if b))
    reason = FOLLOW_TEXT.get(kinds[0], "the points below are pending") if len(kinds) == 1 else "the points below are pending from your side"
    if escalate:
        subject = f"ESCALATION – Pending actions from {supplier} ({len(items)} item(s))"
        body = (f"{greet}\n\nDespite earlier reminders, {reason}. This is now affecting our development and production plan.\n\n"
                + "\n".join(lines) +
                "\n\nPlease treat this as top priority and share a written closure plan with committed dates within 24 hours.\n\n"
                + _sig_block())
    else:
        subject = f"Follow-up: {', '.join(kinds)} – {supplier}"
        body = (f"{greet}\n\nGreetings from {S_().get('company') or 'Sun Mobility'}!\n\nThis is a follow-up as {reason}:\n\n"
                + "\n".join(lines) + "\n\nKindly respond with the required confirmation / documents at the earliest.\n\n"
                + _sig_block())
    return to, cc, subject, body


# ---------------------------------------------------------------------------------------------- inventory projection
def inventory_projection():
    d = D()
    weeks = int(num(S_().get("horizon_weeks")) or 12)
    lines_by_part = {}
    for r in po_lines():
        if not blank(r.get("part_no")) and (r.get("approval") or "") == "Approved" and (r.get("Balance Qty") or 0) > 0:
            lines_by_part.setdefault(_nk(r["part_no"]), []).append(r)
    out = []
    for inv in d["inventory"]:
        stock = num(inv.get("stock")) or 0
        monthly = num(inv.get("monthly")) or 0
        daily = monthly / 30 if monthly else 0
        safety = num(inv.get("safety")) or 0
        lt = int(num(inv.get("lead_time")) or 0)
        confirmed, unconfirmed = [], 0
        pols = lines_by_part.get(_nk(inv.get("part_no")), [])
        if pols:
            for r in pols:
                prom = to_date(r.get("promised_date"))
                if prom and prom >= today() and (r.get("ack") or "") in ("Acknowledged", "Not required"):
                    confirmed.append((prom, r["Balance Qty"], f"PO {r.get('po_no')}"))
                else:  # no acknowledgement, no date, or the promised date has already been missed
                    unconfirmed += r["Balance Qty"] or 0
        else:
            q, cd = num(inv.get("open_po_qty")) or 0, to_date(inv.get("confirmed_date"))
            if q and cd and cd >= today():
                confirmed.append((cd, q, "Open PO (manual)"))
            elif q:
                unconfirmed += q
        # daily simulation (shortages between weekly points are not missed); weekly view = end-of-week level
        days = 7 * weeks
        level, lv = stock, [stock]
        inc_by_day = {}
        for dt, q, _ in confirmed:
            inc_by_day[(dt - today()).days] = inc_by_day.get((dt - today()).days, 0) + q
        for dd in range(1, days + 1):
            level = level + inc_by_day.get(dd, 0) - daily
            lv.append(level)
        lv[0] = stock + inc_by_day.get(0, 0)
        proj = [round(lv[7 * w], 1) for w in range(weeks + 1)]
        projmin = [round(min(lv[max(0, 7 * w - 6):7 * w + 1]), 1) for w in range(weeks + 1)]
        cov = stock / daily if daily else None
        short_d = next((i for i, v in enumerate(lv) if v < 0), None)
        below_d = next((i for i, v in enumerate(lv) if v < safety), None)
        short_w = None if short_d is None else short_d // 7
        # reorder qty: what a new order (arriving after the lead time) must add so that projected stock
        # stays at or above safety stock for the 30 days after it arrives
        win = lv[min(lt, len(lv) - 1): lt + 31]
        reorder = max(math.ceil(safety - min(win)), 0) if daily and win else 0
        expedite = below_d is not None and below_d < lt  # a new order cannot arrive in time
        overdue = any(r["PO Status"] in ("Overdue", "Partial – overdue") for r in pols)
        sup_risk = overdue or unconfirmed > 0
        req = num(inv.get("project_req")) or 0
        prod_risk = bool(req) and req > stock + sum(q for _, q, _ in confirmed)
        if short_d is not None and (short_d <= max(lt, 7) or stock <= 0):
            risk = "RED"  # stock-out before a new order could arrive
        elif short_d is not None or below_d is not None or prod_risk:
            risk = "YELLOW"
        else:
            risk = "GREEN"
        flags = []
        if stock <= safety and daily:
            flags.append("Critical stock")
        if short_w is not None:
            flags.append(f"Stock-out in {short_d} day(s)" if short_d else "Out of stock")
        if reorder:
            flags.append(f"Reorder {reorder}")
        if expedite:
            flags.append("Expedite open POs / spot buy (gap before lead time)")
        if sup_risk:
            flags.append("Supplier delivery risk")
        if prod_risk:
            flags.append("Production shortage risk")
        out.append({"inv": inv, "daily": round(daily, 2), "coverage": round(cov, 1) if cov is not None else None,
                    "proj": proj, "projmin": projmin, "safety": safety, "risk": risk, "flags": flags, "reorder": reorder,
                    "confirmed": confirmed, "unconfirmed": unconfirmed, "short_w": short_w})
    return out


def inv_calc(rec):
    monthly = num(rec.get("monthly")) or 0
    daily = monthly / 30 if monthly else 0
    stock = num(rec.get("stock")) or 0
    return {"Daily Consumption": round(daily, 2) if daily else None,
            "Stock Coverage (days)": round(stock / daily, 1) if daily else None}


# ---------------------------------------------------------------------------------------------- KPIs used by dashboard + AI
def proc_kpis():
    d = D()
    lines = po_lines()
    pos = po_groups()
    act = activities_df(d)
    pend_samples = set()
    pend_dv = set()
    if not act.empty:
        for _, r in act.iterrows():
            if r["Status"] in ("Delayed", "Due in 7 days", "Open") and re.search("sample", str(r["Activity"]), re.I):
                pend_samples.add(r["Project"])
            if r["Status"] in ("Delayed", "Due in 7 days", "Open") and re.search(r"\bDV", str(r["Activity"])):
                pend_dv.add(r["Project"])
    for r in d["dv"]:
        if (r.get("status") or "") not in ("Passed", "Failed"):
            pend_dv.add(r.get("project") or r.get("part_no"))
    return {
        "projects": len(d["projects"]),
        "active_parts": sum(1 for p in d["projects"] if p.get("status") == "Ongoing"),
        "pos": len(pos),
        "pending_approval": len({str(r["po_no"]).strip() for r in lines if r.get("approval") in ("Draft", "Pending approval", None, "")
                                 and not blank(r.get("po_no"))}),
        "ack_pending": len({str(r["po_no"]) for r in lines if r["PO Status"] == "Ack pending"}),
        "eta_pending": len({str(r["po_no"]) for r in lines if r["PO Status"] == "ETA pending"}),
        "overdue": sum(1 for r in lines if r["PO Status"] in ("Overdue", "Partial – overdue")),
        "samples": len(pend_samples),
        "dv": len({x for x in pend_dv if x}),
        "ecn": sum(1 for r in d["ecn"] if (r.get("status") or "Open") not in ("Implemented", "Closed", "Rejected")),
        "drafts": sum(1 for e in d["emails"] if e["status"] in ("Draft", "Approved – send from Outlook")),
    }


def procurement_overview():
    k = proc_kpis()
    st.markdown('<div class="section">Procurement & development control</div>', unsafe_allow_html=True)
    rows = [[("NPD projects", k["projects"], "in tracker", None), ("Active development parts", k["active_parts"], "ongoing projects", None),
             ("Purchase orders", k["pos"], "PO numbers", None), ("Pending PO approvals", k["pending_approval"], "draft / pending", COLORS["warn"]),
             ("PO ack pending", k["ack_pending"], "approved, not acknowledged", COLORS["warn"])],
            [("Supplier ETA pending", k["eta_pending"], "no delivery date", COLORS["warn"]),
             ("Overdue deliveries", k["overdue"], "PO lines past promised date", COLORS["bad"]),
             ("Pending samples", k["samples"], "projects with open sample steps", None),
             ("Pending DV/DVVP", k["dv"], "projects / parts", None), ("Open ECN/ECR", k["ecn"], "not implemented / closed", None)]]
    for row in rows:
        cols = st.columns(5)
        for c, (lab, v, sub, col) in zip(cols, row):
            with c:
                kpi(lab, v, sub, value_color=col if v else None)
    if k["drafts"]:
        st.info(f"✉️ {k['drafts']} email draft(s) waiting for your approval in the Email Approval Centre. Nothing is sent "
                "without your click.")
    lines = po_lines()
    fu = followup_items()
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section">Supplier-wise pending activities</div>', unsafe_allow_html=True)
        if fu:
            df = pd.DataFrame([{"Supplier": i.get("supplier") or i.get("project") or "-", "Type": i["type"]} for i in fu])
            g = df.groupby(["Supplier", "Type"]).size().reset_index(name="Items")
            fig = px.bar(g, x="Items", y="Supplier", color="Type", orientation="h")
            fig.update_yaxes(autorange="reversed", title=None)
            st.plotly_chart(plot_layout(fig, 30 * g["Supplier"].nunique() + 160), width="stretch")
        else:
            st.caption("No pending supplier actions (add RFQs and POs to track them).")
    with c2:
        st.markdown('<div class="section">PO delivery status</div>', unsafe_allow_html=True)
        if lines:
            g = pd.DataFrame(lines)["PO Status"].value_counts().reset_index()
            g.columns = ["Status", "Lines"]
            cmap = {"Overdue": COLORS["bad"], "Partial – overdue": COLORS["bad"], "Ack pending": COLORS["warn"],
                    "ETA pending": COLORS["warn"], "Awaiting approval": COLORS["idle"], "Open – on schedule": COLORS["navy"],
                    "Partially received": COLORS["P1R5"], "Closed – received": COLORS["ok"], "Rejected": "#999",
                    "Cancelled": "#bbb"}
            fig = px.pie(g, names="Status", values="Lines", hole=0.55, color="Status", color_discrete_map=cmap)
            st.plotly_chart(plot_layout(fig, 320), width="stretch")
        else:
            st.caption("No purchase orders yet.")
    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="section">Monthly procurement value (approved POs)</div>', unsafe_allow_html=True)
        vals = [(to_date(r.get("po_date")), r["Total Value"]) for r in lines
                if r.get("approval") == "Approved" and to_date(r.get("po_date")) and r["Total Value"]]
        if vals:
            df = pd.DataFrame(vals, columns=["date", "value"])
            df["Month"] = pd.to_datetime(df["date"]).dt.to_period("M").dt.to_timestamp()
            g = df.groupby("Month")["value"].sum().reset_index()
            fig = px.bar(g, x="Month", y="value", labels={"value": "₹"}, color_discrete_sequence=[COLORS["navy"]])
            fig.update_xaxes(dtick="M1", tickformat="%b %y")
            st.plotly_chart(plot_layout(fig, 300), width="stretch")
        else:
            st.caption("Needs approved POs with PO date, qty and unit price.")
    with c4:
        st.markdown('<div class="section">Delayed material analysis</div>', unsafe_allow_html=True)
        late = [r for r in lines if r["PO Status"] in ("Overdue", "Partial – overdue")]
        if late:
            df = pd.DataFrame([{"Item": f"{r.get('part_no')} · {r.get('supplier')}", "Days late": r["Delay (days)"],
                                "PO": r.get("po_no")} for r in late]).sort_values("Days late", ascending=False).head(15)
            fig = px.bar(df, x="Days late", y="Item", orientation="h", hover_data=["PO"], color_discrete_sequence=[COLORS["bad"]])
            fig.update_yaxes(autorange="reversed", title=None)
            st.plotly_chart(plot_layout(fig, 30 * len(df) + 140), width="stretch")
        else:
            st.caption("No overdue PO lines.")


# ---------------------------------------------------------------------------------------------- pages
def page_suppliers_master():
    header("Supplier Management", "Supplier master used for PO and follow-up emails. Edit like Excel; it saves automatically.")
    view = st.radio("View", ["Supplier list", "Supplier-wise parts analysis"], horizontal=True, label_visibility="collapsed",
                    key="sup_view")
    if view == "Supplier list":
        excel_grid(spec_coll("suppliers"), dropdowns={"status": SUPPLIER_STATUS}, height=420,
                   help_text="Email / CC Emails are used in PO and follow-up drafts. Separate several addresses with a comma.")
        bad = [(s.get("name"), x) for s in D()["suppliers"] for x in bad_emails(s.get("email")) + bad_emails(s.get("cc"))]
        if bad:
            st.warning("Invalid email address(es): " + "; ".join(f"{n}: {x}" for n, x in bad))
    else:
        page_suppliers()


def page_rfq():
    header("RFQ Management", "Track RFQs, quotations and clarifications per supplier. Pending items feed Supplier Follow-ups.")
    sups, projs = supplier_options(), project_options()
    excel_grid(spec_coll("rfq"), dropdowns={k: v for k, v in {"status": RFQ_STATUS, "supplier": sups, "project": projs}.items() if v},
               height=440)


def page_po():
    header("Purchase Order Management",
           "One row per PO line. Total value, balance, delay and status calculate automatically. Setting a PO to "
           "Approved creates an email draft for your review. It is never sent automatically.")
    d = D()
    lines = po_lines()
    if lines:
        k = proc_kpis()
        c = st.columns(5)
        for col, (lab, v, colr) in zip(c, [("PO numbers", k["pos"], None), ("Pending approval", k["pending_approval"], COLORS["warn"]),
                                            ("Ack pending", k["ack_pending"], COLORS["warn"]), ("ETA pending", k["eta_pending"], COLORS["warn"]),
                                            ("Overdue lines", k["overdue"], COLORS["bad"])]):
            with col:
                kpi(lab, v, "", value_color=colr if v else None)
        tot = sum(r["Total Value"] or 0 for r in lines if r.get("approval") == "Approved")
        st.caption(f"Approved PO value: **{inr(tot)}**")
    excel_grid(spec_coll("pos", after=po_after_save), computed=po_calc,
               dropdowns={k: v for k, v in {"approval": PO_APPROVAL, "ack": PO_ACK, "supplier": supplier_options(),
                                             "project": project_options()}.items() if v}, height=480,
               help_text="Supplier Email: leave blank to use the email from Supplier Management. Total Value, Balance Qty, "
                         "Delay (days) and PO Status are calculated.")
    groups = po_groups()
    with st.expander("📎 Attach PO PDF"):
        if not groups:
            st.caption("Add a PO first.")
        else:
            po = st.selectbox("PO number", sorted(groups), key="att_po")
            f = st.file_uploader("PO document (PDF, max 10 MB)", type=["pdf"], key="att_file")
            if f is not None and st.button("Upload and link to this PO", type="primary"):
                if f.size > 10 * 1024 * 1024:
                    st.error("File is larger than 10 MB.")
                else:
                    name = re.sub(r"[^A-Za-z0-9._-]+", "_", f"PO_{po}.pdf")
                    try:
                        where = att_save(name, f.getvalue())
                        for r in d["pos"]:
                            if str(r.get("po_no")).strip() == po:
                                r["attachment"] = name
                        for e in d["emails"]:
                            if e["ref"] == po and e["type"] == "PO approval" and e["status"] == "Draft":
                                if e.get("updated") is None:  # untouched draft: rebuild text with the attachment
                                    _, _, _, body, att = po_email(po, groups[po])
                                    e.update(body=body, attachments=att)
                                elif name not in e["attachments"]:
                                    e["attachments"].append(name)
                        commit(f"PO {po}: attachment {name} uploaded")
                        if where == "session":
                            st.warning("Stored for this session only (GitHub saving is off). Re-upload after a restart.")
                        st.rerun()
                    except Exception as ex:  # noqa: BLE001
                        st.error(f"Upload failed: {ex}")
    with st.expander("✉️ Create the PO email draft again for a PO"):
        if groups:
            po2 = st.selectbox("PO number ", sorted(groups), key="redraft_po")
            if st.button("Create draft"):
                to, cc, subject, body, att = po_email(po2, groups[po2])
                make_draft("PO approval", po2, groups[po2][0].get("supplier"), to, cc, subject, body, att, source="manual")
                commit(f"Email draft created for PO {po2}")
                st.rerun()


EMAIL_STATES = ["Draft", "Approved – send from Outlook", "Sent", "Rejected", "Failed"]


def _email_set(e, **kw):
    e.update(kw)
    e["updated"] = datetime.now().isoformat(timespec="seconds")


def page_email_centre():
    d = D()
    header("Email Approval Centre", "Every supplier email waits here. Nothing is sent until you review it and click Approve.")
    smtp = smtp_cfg()
    st.caption(("🟢 Direct sending is configured (SMTP). Approve & Send sends the email." if smtp else
                "⚪ Direct sending is not configured. Approve creates an Outlook-ready .eml file, which opens as a draft "
                "in Outlook for you to send. Then click Mark as sent.") )
    tab = st.radio("Show", ["Pending", "Sent", "Rejected / failed", "Email history"], horizontal=True, key="em_tab",
                   label_visibility="collapsed")
    if tab == "Email history":
        log = d["email_log"]
        if not log:
            st.caption("No email activity yet.")
        else:
            st.dataframe(pd.DataFrame(log[::-1])[["at", "event", "type", "ref", "supplier", "to", "cc", "subject", "detail"]],
                         hide_index=True, width="stretch", height=480)
        return
    want = {"Pending": ("Draft", "Approved – send from Outlook"), "Sent": ("Sent",), "Rejected / failed": ("Rejected", "Failed")}[tab]
    items = [e for e in d["emails"] if e["status"] in want]
    if not items:
        st.info("No emails in this list.")
        return
    lab = {e["id"]: f"{e['status']} · {e['type']} · {e.get('ref') or ''} · {e.get('supplier') or ''} · {e['created'][:10]}"
           for e in items}
    eid = st.selectbox("Email", [e["id"] for e in items], format_func=lambda i: lab[i], key="em_sel")
    e = next(x for x in items if x["id"] == eid)
    editable = e["status"] == "Draft"
    c1, c2 = st.columns([3, 2])
    with c1:
        with st.form(f"edit_{eid}"):
            to = st.text_input("To", value=e["to"], disabled=not editable)
            cc = st.text_input("CC", value=e["cc"], disabled=not editable)
            subject = st.text_input("Subject", value=e["subject"], disabled=not editable)
            body = st.text_area("Body", value=e["body"], height=360, disabled=not editable)
            saved = st.form_submit_button("💾 Save as draft", disabled=not editable)
        if saved:
            problems = bad_emails(to) + bad_emails(cc)
            if problems:
                st.error("Invalid email address: " + ", ".join(problems))
            else:
                _email_set(e, to=to, cc=cc, subject=subject, body=body)
                _log_email(e, "Draft edited")
                commit("Email draft saved")
                st.rerun()
    with c2:
        st.markdown("**Preview**")
        atts = e.get("attachments") or []
        st.markdown(f"<div class='pcard'><div class='meta'><span><b>To:</b> {e['to'] or '—'}</span></div>"
                    f"<div class='meta'><span><b>CC:</b> {e['cc'] or '—'}</span></div>"
                    f"<div class='meta'><span><b>Subject:</b> {e['subject']}</span></div>"
                    f"<div class='meta'><span><b>Attachments:</b> {', '.join(atts) if atts else 'none'}</span></div></div>",
                    unsafe_allow_html=True)
        st.text(e["body"][:3000])
        missing = [a for a in atts if att_load(a) is None]
        if missing:
            st.warning("Attachment not found: " + ", ".join(missing) + ". Upload it on the Purchase Orders page.")
        problems = (["No recipient (To)"] if not split_emails(e["to"]) else []) + [f"Invalid: {x}" for x in bad_emails(e["to"]) + bad_emails(e["cc"])]
        if problems:
            st.error(" · ".join(problems))
        if e.get("error"):
            st.error(f"Last send error: {e['error']}")
        if e["status"] == "Sent":
            st.success(f"Sent {e.get('sent_at')} ({e.get('method')}) to {e['to']}")
        st.divider()
        if editable:
            ok = st.checkbox("I have reviewed recipient, CC, subject, body and attachments", key=f"rev_{eid}")
            if smtp:
                if st.button("✅ Approve & Send", type="primary", disabled=not ok or bool(problems), key=f"send_{eid}"):
                    if e["status"] == "Sent":
                        st.warning("Already sent.")
                    else:
                        try:
                            smtp_send(e)
                            _email_set(e, status="Sent", sent_at=datetime.now().isoformat(timespec="seconds"), method="SMTP",
                                       error=None)
                            _log_email(e, "Sent", "SMTP")
                            commit(f"Email sent: {e['subject']}")
                            st.rerun()
                        except Exception as ex:  # noqa: BLE001
                            _email_set(e, error=str(ex)[:300])
                            _log_email(e, "Send failed", str(ex)[:300])
                            commit("Email send failed – draft kept")
                            st.error(f"Sending failed, the draft is kept: {ex}")
            else:
                if st.button("✅ Approve – prepare for Outlook", type="primary", disabled=not ok or bool(problems), key=f"appr_{eid}"):
                    _email_set(e, status="Approved – send from Outlook")
                    _log_email(e, "Approved", "for manual sending from Outlook")
                    commit(f"Email approved: {e['subject']}")
                    st.rerun()
            with st.popover("✖ Reject"):
                reason = st.text_input("Reason", key=f"rej_{eid}")
                if st.button("Confirm reject", key=f"rejb_{eid}"):
                    _email_set(e, status="Rejected")
                    _log_email(e, "Rejected", reason)
                    commit("Email rejected")
                    st.rerun()
        if e["status"] == "Approved – send from Outlook":
            st.download_button("⬇ Download .eml (opens as a draft in Outlook)", build_eml(e),
                               file_name=re.sub(r"[^A-Za-z0-9._-]+", "_", e["subject"])[:80] + ".eml", mime="message/rfc822",
                               type="primary", key=f"eml_{eid}")
            st.caption("Open the file → check it → Send from Outlook. Then come back and click Mark as sent.")
            if st.button("📨 Mark as sent", key=f"mark_{eid}"):
                _email_set(e, status="Sent", sent_at=datetime.now().isoformat(timespec="seconds"), method="Outlook (manual)")
                _log_email(e, "Sent", "marked as sent from Outlook")
                commit(f"Email marked as sent: {e['subject']}")
                st.rerun()
            if st.button("↩ Back to draft", key=f"back_{eid}"):
                _email_set(e, status="Draft")
                _log_email(e, "Returned to draft")
                commit("Email returned to draft")
                st.rerun()
        if e["status"] in ("Sent", "Rejected", "Failed"):
            if st.button("⧉ Copy as a new draft", key=f"copy_{eid}"):
                make_draft(e["type"], e["ref"], e["supplier"], e["to"], e["cc"], e["subject"], e["body"], list(e["attachments"]),
                           source=f"copy of {e['id']}")
                commit("Draft copied")
                st.rerun()


def page_followups():
    header("Supplier Follow-ups", "Pending supplier actions found in RFQs, POs and NPD activities. Create follow-up drafts. "
                                  "They go to the Email Approval Centre for your approval.")
    items = followup_items()
    if not items:
        st.success("No pending supplier actions.")
        st.caption("Follow-ups come from: RFQ status / quote due date, PO approval + acknowledgement + promised date, "
                   "and delayed sample steps in NPD projects.")
        return
    df = pd.DataFrame([{"Type": i["type"], "Supplier": i.get("supplier") or "(not linked)", "Reference": i["ref"],
                        "Part": i.get("part"), "Project": i.get("project"), "Detail": i.get("detail"), "Days": i.get("days"),
                        "Email": i.get("email")} for i in items])
    c1, c2 = st.columns(2)
    tf = c1.multiselect("Type", sorted(df["Type"].unique()), key="fu_t")
    sf = c2.multiselect("Supplier", sorted(df["Supplier"].unique()), key="fu_s")
    f = df.copy()
    if tf:
        f = f[f["Type"].isin(tf)]
    if sf:
        f = f[f["Supplier"].isin(sf)]
    k = st.columns(4)
    for col, (lab, v) in zip(k, [("Pending items", len(f)), ("Suppliers", f["Supplier"].nunique()),
                                  ("Overdue material", int((f["Type"] == "Overdue material").sum())),
                                  ("Ack / ETA pending", int(f["Type"].isin(["PO acknowledgement pending", "Supplier ETA pending"]).sum()))]):
        with col:
            kpi(lab, v)
    st.dataframe(f, hide_index=True, width="stretch", height=380,
                 column_config={"Days": st.column_config.NumberColumn(format="%d")})
    st.markdown('<div class="section">Create follow-up drafts</div>', unsafe_allow_html=True)
    sups = sorted(s for s in f["Supplier"].unique() if s != "(not linked)")
    pick = st.multiselect("Suppliers to follow up (one email per supplier)", sups, default=sups, key="fu_pick")
    esc = st.checkbox("Escalation tone", key="fu_esc")
    use_ai = ai_cfg() is not None and st.checkbox("Improve wording with AI (OpenAI)", key="fu_ai")
    if st.button("Create drafts", type="primary", disabled=not pick):
        n = 0
        for sup in pick:
            its = [i for i in items if (i.get("supplier") or "") == sup and (not tf or i["type"] in tf)]
            if not its:
                continue
            to, cc, subject, body = followup_email(sup, its, escalate=esc)
            if use_ai:
                body = ai_rewrite(body) or body
            make_draft("Escalation" if esc else "Follow-up", ", ".join(sorted({i["ref"] for i in its}))[:120], sup, to, cc,
                       subject, body, source="follow-up")
            n += 1
        commit(f"{n} follow-up draft(s) created")
        st.session_state.flash_info = f"{n} draft(s) created. Review them in the Email Approval Centre. Nothing has been sent."
        st.rerun()
    nl = df[df["Supplier"] == "(not linked)"]
    if len(nl):
        st.caption(f"{len(nl)} item(s) come from NPD projects with no supplier linked; follow these up from the project.")


def page_inventory():
    header("Inventory & Shortage Prediction",
           "Projected stock = current stock + confirmed incoming − forecast consumption. Only POs that are acknowledged "
           "and have a promised date count as incoming. Unconfirmed quantities are shown but not counted.")
    view = st.radio("View", ["Risk overview", "Edit inventory data"], horizontal=True, label_visibility="collapsed", key="inv_view")
    if view == "Edit inventory data":
        excel_grid(spec_coll("inventory"), computed=inv_calc, height=440,
                   help_text="Open PO Qty (manual) and Confirmed Delivery Date are used only when the part has no lines in "
                             "Purchase Orders.")
        return
    res = inventory_projection()
    if not res:
        st.info("No inventory data yet. Switch to 'Edit inventory data' and add parts (or import from Excel in Settings & Data).")
        return
    k = st.columns(5)
    cnt = {r: sum(1 for x in res if x["risk"] == r) for r in ("RED", "YELLOW", "GREEN")}
    vals = [("Critical (RED)", cnt["RED"], COLORS["bad"]), ("Watch (YELLOW)", cnt["YELLOW"], COLORS["warn"]),
            ("OK (GREEN)", cnt["GREEN"], COLORS["ok"]), ("Reorder needed", sum(1 for x in res if x["reorder"]), COLORS["warn"]),
            ("Supplier delivery risk", sum(1 for x in res if "Supplier delivery risk" in x["flags"]), COLORS["bad"])]
    for col, (lab, v, c) in zip(k, vals):
        with col:
            kpi(lab, v, "parts", value_color=c if v else None)
    weeks = len(res[0]["proj"]) - 1
    heads = "".join(f"<th>{'Today' if w == 0 else (today() + timedelta(days=7 * w)).strftime('%d %b')}</th>" for w in range(weeks + 1))
    rows = []
    order = {"RED": 0, "YELLOW": 1, "GREEN": 2}
    for x in sorted(res, key=lambda x: (order[x["risk"]], x["short_w"] if x["short_w"] is not None else 99)):
        inv = x["inv"]
        cells = []
        for v, vmin in zip(x["proj"], x["projmin"]):
            c = COLORS["bad"] if vmin < 0 else (COLORS["warn"] if vmin < x["safety"] else COLORS["ok"])
            cells.append(f"<td title='end of week {v:,.0f} · lowest {vmin:,.0f}' style='background:{c};color:#fff;"
                         f"text-align:center;font-size:11px'>{v:,.0f}</td>")
        badge = {"RED": COLORS["bad"], "YELLOW": COLORS["warn"], "GREEN": COLORS["ok"]}[x["risk"]]
        rows.append(f"<tr><td style='white-space:nowrap'><b>{inv.get('part_no') or ''}</b><br><span style='color:#5D6B7C;font-size:11px'>"
                    f"{(inv.get('part_name') or '')[:40]}</span></td><td style='text-align:right'>{num(inv.get('stock')) or 0:,.0f}</td>"
                    f"<td style='text-align:right'>{x['safety']:,.0f}</td><td style='text-align:right'>"
                    f"{'' if x['coverage'] is None else x['coverage']}</td>"
                    f"<td><span class='pill' style='background:{badge};color:#fff'>{x['risk']}</span></td>" + "".join(cells) + "</tr>")
    st.markdown('<div class="section">Projected inventory by week</div>', unsafe_allow_html=True)
    st.markdown("<div style='overflow-x:auto'><table style='border-collapse:separate;border-spacing:3px;font-size:12px;width:100%'>"
                "<thead><tr style='color:#5D6B7C;font-size:11px'><th style='text-align:left'>Part</th><th>Stock</th><th>Safety</th>"
                f"<th>Cover (days)</th><th>Risk</th>{heads}</tr></thead><tbody>{''.join(rows)}</tbody></table></div>",
                unsafe_allow_html=True)
    st.caption("Colour = lowest level during the week (green ≥ safety stock · yellow below safety · red stock-out). "
               "Number = projected units at week end. Overdue or unacknowledged POs are not counted.")
    st.markdown('<div class="section">Risk details</div>', unsafe_allow_html=True)
    det = pd.DataFrame([{"Part No": x["inv"].get("part_no"), "Part Name": x["inv"].get("part_name"), "Risk": x["risk"],
                         "Daily Cons.": x["daily"], "Coverage (days)": x["coverage"], "Reorder Qty": x["reorder"] or None,
                         "Confirmed incoming": sum(q for _, q, _ in x["confirmed"]) or None,
                         "Unconfirmed (not counted)": x["unconfirmed"] or None, "Flags": ", ".join(x["flags"])} for x in res])
    st.dataframe(det, hide_index=True, width="stretch")
    pick = st.selectbox("Incoming supply detail for part", [x["inv"].get("part_no") for x in res], key="inv_pick")
    x = next(r for r in res if r["inv"].get("part_no") == pick)
    if x["confirmed"]:
        st.dataframe(pd.DataFrame([{"Date": dt, "Qty": q, "Source": src} for dt, q, src in x["confirmed"]]), hide_index=True)
    else:
        st.caption("No confirmed incoming supply.")
    if x["unconfirmed"]:
        st.warning(f"{x['unconfirmed']:,.0f} units on order are not confirmed (no acknowledgement or promised date) "
                   "and are not counted.")


def page_dv():
    header("DV / DVVP Tracking", "Design validation status per part and supplier.")
    excel_grid(spec_coll("dv"), dropdowns={k: v for k, v in {"status": DV_STATUS, "project": project_options(),
                                                            "supplier": supplier_options()}.items() if v}, height=440)


def page_ecn():
    header("ECN / ECR Tracking", "Engineering change notes and requests linked to alternate development.")
    excel_grid(spec_coll("ecn"), dropdowns={k: v for k, v in {"status": ECN_STATUS, "type": ["ECN", "ECR"],
                                                             "project": project_options()}.items() if v}, height=440)


# ---------------------------------------------------------------------------------------------- AI assistant
def ai_cfg():
    try:
        c = st.secrets.get("openai", None)
    except Exception:  # noqa: BLE001
        return None
    return {"key": c["api_key"], "model": c.get("model") or S_().get("ai_model") or "gpt-4o-mini"} if c and c.get("api_key") else None


def _ai_call(messages):
    import requests
    cfg = ai_cfg()
    r = requests.post("https://api.openai.com/v1/chat/completions", timeout=60,
                      headers={"Authorization": f"Bearer {cfg['key']}"},
                      json={"model": cfg["model"], "messages": messages, "temperature": 0.2})
    if r.status_code != 200:
        raise RuntimeError(f"OpenAI {r.status_code}: {r.text[:200]}")
    return r.json()["choices"][0]["message"]["content"]


def ai_rewrite(body):
    try:
        return _ai_call([{"role": "system", "content": "Rewrite this supplier email to be clear, professional and firm. "
                                                       "Keep every fact, number, PO/RFQ reference and date exactly. Do not add facts."},
                         {"role": "user", "content": body}])
    except Exception as ex:  # noqa: BLE001
        st.warning(f"AI rewrite failed, template text kept: {ex}")
        return None


def ai_context():
    lines = po_lines()
    fu = followup_items()
    inv = inventory_projection()
    acts = activities_df(D())
    delayed = acts[acts["Status"] == "Delayed"][["Project", "Activity", "Plan Date", "Days Late"]] if not acts.empty else pd.DataFrame()
    return {
        "today": today().isoformat(),
        "kpis": proc_kpis(),
        "overdue_po_lines": [{k: r.get(k) for k in ("po_no", "part_no", "supplier", "promised_date", "Balance Qty", "Delay (days)")}
                             for r in lines if r["PO Status"] in ("Overdue", "Partial – overdue")],
        "ack_pending": [{k: r.get(k) for k in ("po_no", "supplier", "po_date")} for r in lines if r["PO Status"] == "Ack pending"],
        "eta_pending": [{k: r.get(k) for k in ("po_no", "supplier", "part_no")} for r in lines if r["PO Status"] == "ETA pending"],
        "followups": [{k: i.get(k) for k in ("type", "supplier", "ref", "detail", "days")} for i in fu][:80],
        "inventory_risk": [{"part_no": x["inv"].get("part_no"), "risk": x["risk"], "flags": x["flags"]} for x in inv if x["risk"] != "GREEN"],
        "delayed_npd_activities": json.loads(delayed.head(60).to_json(orient="records", date_format="iso")) if len(delayed) else [],
        "projects": [{"name": p["name"], "status": p["status"], **{k: v for k, v in proj_stats(p).items() if k in ("done", "total")},
                      "delayed": len(proj_stats(p)["late"])} for p in D()["projects"]],
    }


def _answer(topic):
    """Rule-based answers computed from the app data (no AI needed)."""
    lines = po_lines()
    if topic == "eta":
        rows = [r for r in lines if r["PO Status"] == "ETA pending"]
        return "Suppliers that have acknowledged but not confirmed an ETA", pd.DataFrame(
            [{"Supplier": r.get("supplier"), "PO": r.get("po_no"), "Part": r.get("part_no"), "PO date": r.get("po_date")} for r in rows])
    if topic == "overdue":
        rows = [r for r in lines if r["PO Status"] in ("Overdue", "Partial – overdue")]
        return "Overdue PO lines", pd.DataFrame([{"PO": r.get("po_no"), "Supplier": r.get("supplier"), "Part": r.get("part_no"),
                                                  "Promised": r.get("promised_date"), "Balance": r["Balance Qty"],
                                                  "Days late": r["Delay (days)"]} for r in rows]).sort_values("Days late", ascending=False) if rows else pd.DataFrame()
    if topic == "stock":
        res = [x for x in inventory_projection() if x["risk"] != "GREEN" or "Critical stock" in x["flags"]]
        return "Materials with critical stock / shortage risk", pd.DataFrame(
            [{"Part": x["inv"].get("part_no"), "Name": x["inv"].get("part_name"), "Risk": x["risk"], "Coverage (days)": x["coverage"],
              "Flags": ", ".join(x["flags"])} for x in res])
    if topic == "npd":
        rows = []
        for p in D()["projects"]:
            s = proj_stats(p)
            if s["late"]:
                worst = max(s["late"], key=lambda l: l["days"])
                rows.append({"Project": p["name"], "Status": p["status"], "Delayed activities": len(s["late"]),
                             "Worst": f"{worst['act']} ({worst['days']} d)"})
        return "NPD projects with delayed activities", pd.DataFrame(rows).sort_values("Delayed activities", ascending=False) if rows else pd.DataFrame()
    if topic == "ack":
        rows = [r for r in lines if r["PO Status"] == "Ack pending"]
        return "Suppliers with pending PO acknowledgements", pd.DataFrame(
            [{"Supplier": r.get("supplier"), "PO": r.get("po_no"), "PO date": r.get("po_date"), "Part": r.get("part_no")} for r in rows])
    if topic == "summary":
        fu = followup_items()
        k = proc_kpis()
        txt = (f"**{k['pos']}** POs · **{k['pending_approval']}** awaiting approval · **{k['ack_pending']}** ack pending · "
               f"**{k['eta_pending']}** ETA pending · **{k['overdue']}** overdue lines · **{len(fu)}** supplier follow-ups · "
               f"**{k['samples']}** projects with pending samples · **{k['dv']}** pending DV/DVVP · **{k['ecn']}** open ECN/ECR · "
               f"**{k['drafts']}** emails awaiting your approval.")
        g = pd.DataFrame(fu).groupby("type").size().reset_index(name="items").rename(columns={"type": "Pending activity"}) if fu else pd.DataFrame()
        return txt, g
    return None, None


QUICK_Q = {"Which suppliers have not confirmed their ETA?": "eta", "Which POs are overdue?": "overdue",
           "Which materials have critical stock?": "stock", "Which NPD projects are delayed?": "npd",
           "Which suppliers have pending acknowledgements?": "ack", "Summarize pending procurement activities": "summary"}


def _route(q):
    ql = q.lower()
    for pat, t in ((r"eta|delivery date", "eta"), (r"overdue|late po|delay.*po", "overdue"), (r"stock|shortage|inventory|material", "stock"),
                   (r"npd|project|develop", "npd"), (r"ack", "ack"), (r"summar|pending|status", "summary")):
        if re.search(pat, ql):
            return t
    return None


def page_ai():
    header("AI Assistant", "Answers come from the data in this app. Nothing is invented.")
    cfg = ai_cfg()
    st.caption("🟢 OpenAI connected: free-text questions are answered from your data." if cfg else
               "⚪ OpenAI not connected: built-in answers from your data. Add [openai] api_key in Streamlit Secrets to enable free-text AI.")
    cols = st.columns(3)
    for i, (q, t) in enumerate(QUICK_Q.items()):
        if cols[i % 3].button(q, key=f"q_{t}", width="stretch"):
            st.session_state.ai_q = q
            st.session_state.ai_topic = t
    q = st.text_input("Ask a question", value=st.session_state.get("ai_q", ""), key="ai_free")
    if st.button("Ask", type="primary") and q:
        st.session_state.ai_q = q
        st.session_state.ai_topic = None
    q = st.session_state.get("ai_q")
    if q:
        topic = st.session_state.get("ai_topic") or _route(q)
        st.markdown(f"**Q:** {q}")
        if cfg and not st.session_state.get("ai_topic"):
            try:
                ans = _ai_call([{"role": "system", "content": "You are a procurement analyst for Sun Mobility. Answer ONLY from the JSON "
                                                              "data provided. If the data does not contain the answer, say so. Use "
                                                              "concise bullets and include PO/part/supplier references."},
                                {"role": "user", "content": f"DATA:\n{json.dumps(ai_context(), default=str)[:60000]}\n\nQUESTION: {q}"}])
                st.markdown(ans)
                st.caption("AI answer based on the app data sent with the question.")
            except Exception as ex:  # noqa: BLE001
                st.error(f"AI call failed: {ex}")
        if topic:
            title, df = _answer(topic)
            st.markdown(f"**{title}**")
            if df is not None and len(df):
                st.dataframe(df, hide_index=True, width="stretch")
            elif topic != "summary":
                st.success("Nothing found in the current data.")
        elif not cfg:
            st.info("I can answer: ETA pending, overdue POs, critical stock, delayed NPD projects, pending acknowledgements, "
                    "summary, or generate an escalation email (below).")
    st.markdown('<div class="section">Generate an escalation email</div>', unsafe_allow_html=True)
    fu = followup_items()
    sups = sorted({i["supplier"] for i in fu if i.get("supplier")})
    if not sups:
        st.caption("No supplier has pending items to escalate.")
        return
    sup = st.selectbox("Supplier", sups, key="esc_sup")
    its = [i for i in fu if i.get("supplier") == sup]
    st.caption(f"{len(its)} pending item(s): " + ", ".join(sorted({i['type'] for i in its})))
    if st.button("Create escalation draft"):
        to, cc, subject, body = followup_email(sup, its, escalate=True)
        if cfg:
            body = ai_rewrite(body) or body
        make_draft("Escalation", ", ".join(sorted({i["ref"] for i in its}))[:120], sup, to, cc, subject, body, source="AI assistant")
        commit(f"Escalation draft created for {sup}")
        st.session_state.flash_info = "Escalation draft created in the Email Approval Centre. Nothing has been sent."
        st.rerun()


# ---------------------------------------------------------------------------------------------- reports
def _table_export(title_rows, sheet):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, aoa, widths, money in title_rows:
        _write_table(wb.create_sheet(name[:31]), aoa, widths, money_cols=money)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def coll_aoa(t, extra=None):
    m = tmeta(D(), t)
    cols = vis(m)
    recs = D()[TABLE_DEFS[t]["coll"]]
    ex_titles = list(extra({}).keys()) if extra else []
    aoa = [[m["labels"][c] for c in cols] + ex_titles]
    for r in recs:
        e = extra(r) if extra else {}
        aoa.append([r.get(c) for c in cols] + [e.get(x) for x in ex_titles])
    return aoa


def page_reports():
    header("Reports", "Excel downloads. Original Excel files are never overwritten: every download is a new file.")
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    xl = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ss = st.session_state
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section">NPD</div>', unsafe_allow_html=True)
        if ss.get("orig_bytes"):
            st.download_button("⬇ NPD master Excel (your format, updated)", formatted_export(ss.orig_bytes, D()),
                               file_name=f"NPD_Master_Tracker_{stamp}.xlsx", mime=xl, width="stretch")
        else:
            st.caption("Upload the master Excel in Settings & Data to enable the formatted NPD download.")
        st.download_button("⬇ NPD report (all sheets, clean)", clean_export(D()), file_name=f"NPD_Report_{stamp}.xlsx",
                           mime=xl, width="stretch")
        st.markdown('<div class="section">Purchase orders</div>', unsafe_allow_html=True)
        st.download_button("⬇ PO report", _table_export([("PO Lines", coll_aoa("pos", po_calc), [14] * 24, ())], "pos"),
                           file_name=f"PO_Report_{stamp}.xlsx", mime=xl, width="stretch")
    with c2:
        st.markdown('<div class="section">Suppliers</div>', unsafe_allow_html=True)
        fu = followup_items()
        aoa = [["Type", "Supplier", "Reference", "Part", "Project", "Detail", "Days pending", "Email"]] + [
            [i["type"], i.get("supplier"), i["ref"], i.get("part"), i.get("project"), i.get("detail"), i.get("days"), i.get("email")]
            for i in fu]
        st.download_button("⬇ Supplier pending report", _table_export([("Supplier Pending", aoa, [26, 24, 18, 14, 22, 40, 12, 26], ())], "fu"),
                           file_name=f"Supplier_Pending_{stamp}.xlsx", mime=xl, width="stretch")
        st.markdown('<div class="section">Inventory</div>', unsafe_allow_html=True)
        res = inventory_projection()
        weeks = len(res[0]["proj"]) if res else 0
        aoa = [["Part No", "Part Name", "Risk", "Stock", "Safety", "Daily Cons.", "Coverage (days)", "Reorder Qty",
                "Confirmed incoming", "Unconfirmed (not counted)", "Flags"] +
               [("Today" if w == 0 else (today() + timedelta(days=7 * w)).strftime("%d-%b")) for w in range(weeks)]]
        for x in res:
            aoa.append([x["inv"].get("part_no"), x["inv"].get("part_name"), x["risk"], num(x["inv"].get("stock")), x["safety"],
                        x["daily"], x["coverage"], x["reorder"], sum(q for _, q, _ in x["confirmed"]), x["unconfirmed"],
                        ", ".join(x["flags"])] + x["proj"])
        st.download_button("⬇ Material shortage report", _table_export([("Shortage", aoa, [14, 28, 9] + [12] * (8 + weeks), ())], "inv"),
                           file_name=f"Material_Shortage_{stamp}.xlsx", mime=xl, width="stretch")
        st.markdown('<div class="section">Other trackers</div>', unsafe_allow_html=True)
        st.download_button("⬇ RFQ, DV/DVVP, ECN and suppliers (one workbook)",
                           _table_export([("Suppliers", coll_aoa("suppliers"), [16] * 12, ()), ("RFQ", coll_aoa("rfq"), [16] * 12, ()),
                                          ("DV-DVVP", coll_aoa("dv"), [16] * 14, ()), ("ECN-ECR", coll_aoa("ecn"), [16] * 14, ())], "all"),
                           file_name=f"Trackers_{stamp}.xlsx", mime=xl, width="stretch")


# ---------------------------------------------------------------------------------------------- settings, import, versions
IMPORT_TARGETS = {"Suppliers": "suppliers", "RFQs": "rfq", "Purchase orders": "pos", "Inventory": "inventory",
                  "DV/DVVP": "dv", "ECN/ECR": "ecn"}


def _import_table():
    st.markdown('<div class="section">Import a tracker (Excel / CSV)</div>', unsafe_allow_html=True)
    tgt_lab = st.selectbox("Import into", list(IMPORT_TARGETS), key="imp_t")
    t = IMPORT_TARGETS[tgt_lab]
    f = st.file_uploader("File", type=["xlsx", "xlsm", "csv"], key="imp_f")
    if f is None:
        return
    try:
        if f.name.lower().endswith(".csv"):
            raw = pd.read_csv(f, dtype=object)
        else:
            xls = pd.ExcelFile(f)
            sh = st.selectbox("Sheet", xls.sheet_names, key="imp_sh")
            hdr = st.number_input("Header row number", min_value=1, value=1, step=1, key="imp_hr")
            raw = pd.read_excel(xls, sheet_name=sh, header=int(hdr) - 1, dtype=object)
    except Exception as ex:  # noqa: BLE001
        st.error(f"Could not read the file: {ex}")
        return
    raw = raw.dropna(how="all")
    m = tmeta(D(), t)
    fields = [c for c in TABLE_DEFS[t]["core"] if c != "_id"]
    cols = ["(skip)"] + [str(c) for c in raw.columns]
    st.caption("Match your columns to the app fields (auto-matched by name, change where needed).")
    mapping = {}
    grid = st.columns(3)
    for i, fld in enumerate(fields):
        lab = m["labels"][fld]
        guess = next((c for c in raw.columns if _nk(c) == _nk(lab) or _nk(c) == _nk(fld)), None)
        mapping[fld] = grid[i % 3].selectbox(lab, cols, index=cols.index(str(guess)) if guess is not None else 0, key=f"imp_{t}_{fld}")
    mode = st.radio("Mode", ["Append (skip duplicates)", "Replace all existing rows"], horizontal=True, key="imp_mode")
    st.dataframe(raw.head(5), hide_index=True)
    if st.button(f"Import {len(raw)} row(s)", type="primary"):
        recs = []
        for _, row in raw.iterrows():
            rec = {"_id": new_id()}
            for fld, col in mapping.items():
                if col != "(skip)":
                    v = row[col]
                    rec[fld] = _from_cell(None if (isinstance(v, float) and pd.isna(v)) else jsonable(v.to_pydatetime() if hasattr(v, "to_pydatetime") else v),
                                          _ctype(m, t, fld))
            if any(not blank(rec.get(f)) for f in fields):
                recs.append(rec)
        coll = TABLE_DEFS[t]["coll"]
        base = [] if mode.startswith("Replace") else list(D()[coll])
        u = TABLE_DEFS[t]["unique"] or []
        keyf = (lambda r: tuple(str(r.get(x) or "").strip().lower() for x in u))
        have = {keyf(r) for r in base}
        added = skipped = 0
        for r in recs:
            if u and keyf(r)[0] and keyf(r) in have:
                skipped += 1
                continue
            base.append(r)
            have.add(keyf(r))
            added += 1
        D()[coll] = base
        commit(f"Imported {added} row(s) into {tgt_lab} ({skipped} duplicate(s) skipped)")
        st.rerun()


def _versions():
    st.markdown('<div class="section">Version history (backups)</div>', unsafe_allow_html=True)
    cfg = gh_cfg()
    if not cfg:
        st.caption("Every save is a backup when GitHub saving is on. It is off now. Use the JSON backup download instead.")
        return
    if st.button("Load recent versions"):
        r = _gh("GET", "/commits", cfg, params={"sha": cfg["branch"], "path": cfg["path"], "per_page": 20})
        st.session_state.versions = [{"sha": c["sha"], "date": c["commit"]["author"]["date"], "msg": c["commit"]["message"]}
                                     for c in r.json()] if r.status_code == 200 else []
    vs = st.session_state.get("versions") or []
    if vs:
        pick = st.selectbox("Version", vs, format_func=lambda v: f"{v['date'][:16].replace('T', ' ')} · {v['msg'][:70]}", key="ver_pick")
        conf = st.checkbox("Restore this version (the current data is kept in history)", key="ver_conf")
        if st.button("Restore", disabled=not conf):
            import base64
            r = _gh("GET", f"/contents/{cfg['path']}", cfg, params={"ref": pick["sha"]})
            if r.status_code == 200:
                raw = base64.b64decode(r.json()["content"]) if r.json().get("content") else base64.b64decode(
                    _gh("GET", f"/git/blobs/{r.json()['sha']}", cfg).json()["content"])
                st.session_state.data = parse_backup(raw)
                commit(f"Restored version from {pick['date'][:16]}")
                st.rerun()
            else:
                st.error(f"Could not load that version ({r.status_code}).")


def page_settings():
    header("Settings & Data", "Company details for emails, follow-up rules, integrations, imports, backups.")
    view = st.radio("Section", ["Settings", "Load / download NPD data", "Import trackers", "Version history"], horizontal=True,
                    key="set_view", label_visibility="collapsed")
    if view == "Load / download NPD data":
        page_data()
        return
    if view == "Import trackers":
        _import_table()
        return
    if view == "Version history":
        _versions()
        return
    s = S_()
    with st.form("settings"):
        c1, c2 = st.columns(2)
        company = c1.text_input("Company name", s.get("company"))
        team = c2.text_input("Team name (email signature)", s.get("team"))
        sname = c1.text_input("Sender name", s.get("sender_name"))
        semail = c2.text_input("Sender email (From)", s.get("sender_email"))
        dcc = c1.text_input("Default CC on supplier emails", s.get("default_cc"))
        sig = c2.text_input("Extra signature line (phone etc.)", s.get("signature"))
        ack = c1.number_input("Follow up PO acknowledgement after (days)", 0, 60, int(num(s.get("ack_days")) or 2))
        hz = c2.number_input("Inventory projection horizon (weeks)", 4, 26, int(num(s.get("horizon_weeks")) or 12))
        model = c1.text_input("OpenAI model", s.get("ai_model") or "gpt-4o-mini")
        if st.form_submit_button("Save settings", type="primary"):
            problems = bad_emails(semail) + bad_emails(dcc)
            if problems:
                st.error("Invalid email: " + ", ".join(problems))
            else:
                s.update(company=company, team=team, sender_name=sname, sender_email=semail, default_cc=dcc, signature=sig,
                         ack_days=int(ack), horizon_weeks=int(hz), ai_model=model)
                commit("Settings saved")
                st.rerun()
    st.markdown('<div class="section">Integrations</div>', unsafe_allow_html=True)
    rows = [("Permanent saving (GitHub)", "🟢 On" if gh_cfg() else "⚪ Off", "Secrets → [github]"),
            ("Email sending (SMTP)", "🟢 On" if smtp_cfg() else "⚪ Off — drafts open in Outlook (.eml)",
             "Secrets → [smtp] host, port, username, password, from_addr"),
            ("AI (OpenAI)", "🟢 On" if ai_cfg() else "⚪ Off — built-in answers", "Secrets → [openai] api_key")]
    st.dataframe(pd.DataFrame(rows, columns=["Integration", "Status", "How to enable"]), hide_index=True, width="stretch")
    st.caption("Passwords and API keys are read only from Streamlit Secrets. They are never stored in the data file.")


# ---------------------------------------------------------------------------------------------- Split AC PBOM (buyer view)
PBOM_PARAMS = {"p1r5_plan": ("P1R5 plan (month)", 211), "p1r5_fg": ("P1R5 FG", 73), "wip": ("WIP", 24),
               "sp_wip": ("SP WIP", 0), "sp_plan": ("Single Phs plan (month)", 29), "sp_fg": ("Single Phs FG", 4)}


def pbom_params() -> dict:
    p = S_().setdefault("pbom", {})
    for k, (_, dv) in PBOM_PARAMS.items():
        p.setdefault(k, dv)
    p.setdefault("buyer", "Hemanth")
    p.setdefault("title", "Split AC + Delta Charger – Master PBOM")
    p["bal_p1r5"] = (num(p["p1r5_plan"]) or 0) - (num(p["p1r5_fg"]) or 0) - (num(p["wip"]) or 0)
    p["bal_sp"] = (num(p["sp_plan"]) or 0) - (num(p["sp_fg"]) or 0)
    return p


def _pbom_kind(project) -> str:
    s = str(project or "").strip().lower()
    return "common" if s.startswith("common") else ("sp" if s.startswith("single") else "p1r5")


def pbom_calc(r: dict) -> dict:
    """Same formulas as the Master PBOM sheet, chosen by Project:
    Common:     Build = (WIP+SP WIP)×P1R5 Qty · Coverage = SMPL-3 ÷ P1R5 Qty · Shortage = Coverage − (Bal P1R5 + Bal SP)
    P1R5:       Build = WIP×P1R5 Qty          · Coverage = SMPL-3 ÷ P1R5 Qty · Shortage = Coverage − Bal P1R5
    Single Phs: Build = SP WIP×Sing Phs Qty   · Coverage = SMPL-3 ÷ Sing Phs Qty · Shortage = Coverage − Bal SP
    SMPL-3 Stock = (Store Stock + WIP) − Build."""
    out = {"P1R5 Build": None, "SMPL-3 Stock": None, "Coverage in Sets": None, "Total Shortage in Set": None}
    if not r or blank(r.get("part_no")):
        return out
    p = pbom_params()
    kind = _pbom_kind(r.get("project"))
    q1, qs = num(r.get("p1r5_qty")), num(r.get("sp_qty"))
    stock, wip = num(r.get("store_stock")) or 0, num(r.get("wip")) or 0
    if kind == "common":
        build, per, need = ((p["wip"] or 0) + (p["sp_wip"] or 0)) * (q1 or 0), q1, p["bal_p1r5"] + p["bal_sp"]
    elif kind == "sp":
        build, per, need = (p["sp_wip"] or 0) * (qs or 0), qs, p["bal_sp"]
    else:
        build, per, need = (p["wip"] or 0) * (q1 or 0), q1, p["bal_p1r5"]
    smpl = stock + wip - build
    cov = smpl / per if per else None
    out.update({"P1R5 Build": round(build, 2), "SMPL-3 Stock": round(smpl, 2),
                "Coverage in Sets": None if cov is None else round(cov, 2),
                "Total Shortage in Set": None if cov is None else round(cov - need, 2)})
    return out


PBOM_HDR = {"sl": ("sl",), "part_no": ("partno", "partnumber"), "description": ("description",),
            "p1r5_qty": ("p1r5qty",), "sp_qty": ("singphsqty", "singlephsqty", "singlephaseqty"), "uom": ("uom",),
            "buyer": ("buyer",), "supplier": ("supplier",), "project": ("project", "projtype"), "insp": ("insp",),
            "store_stock": ("storestock", "storestockasondate"), "wip": ("wip", "wipasondate"), "remarks": ("remarks",)}


def pbom_read(file_bytes, sheet=None):
    """Read a PBOM sheet (values, not formulas). Returns (sheet names, chosen sheet, rows as dicts)."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    names = wb.sheetnames
    sheet = sheet or next((n for n in names if n.strip().lower() == "master pbom"), names[0])
    ws = wb[sheet]
    hdr_row, cols, rows = None, {}, []
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=40, max_col=40, values_only=True), 1):
        keys = [re.sub(r"[^a-z0-9]", "", str(v or "").lower()) for v in row]
        if keys and keys[0] == "sl" and any(k in ("partno", "partnumber") for k in keys):
            hdr_row = i
            for j, k in enumerate(keys):
                for fid, syn in PBOM_HDR.items():
                    if k in syn and fid not in cols:
                        cols[fid] = j
            break
    if hdr_row is None:
        return names, sheet, None
    for row in ws.iter_rows(min_row=hdr_row + 1, max_col=max(cols.values()) + 1, values_only=True):
        rec = {fid: row[j] if j < len(row) else None for fid, j in cols.items()}
        if blank(rec.get("part_no")) or not isinstance(rec.get("sl"), (int, float)):
            continue  # skip the 'Receipt plan' / blank helper rows under each part
        rec["part_no"] = part_no(rec["part_no"])
        for f in ("buyer", "supplier", "project", "description", "uom", "remarks"):
            if isinstance(rec.get(f), str):
                rec[f] = rec[f].strip()
        rows.append(rec)
    wb.close()
    return names, sheet, rows


def pbom_excel(recs, p) -> bytes:
    """Excel with live formulas (same logic as the source sheet)."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PBOM – " + str(p.get("buyer") or "Buyer")[:20]
    navy, grey = PatternFill("solid", fgColor="1F3864"), PatternFill("solid", fgColor="D9D9D9")
    th = Side(style="thin", color="9E9E9E")
    bd = Border(left=th, right=th, top=th, bottom=th)
    ws["A1"] = p.get("title")
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    params = [("P1R5 plan", p["p1r5_plan"]), ("P1R5 FG", p["p1r5_fg"]), ("WIP", p["wip"]), ("SP WIP", p["sp_wip"]),
              ("Single Phs plan", p["sp_plan"]), ("Single Phs FG", p["sp_fg"]),
              ("Balance P1R5", "=C3-C4-C5"), ("Balance Single Phs", "=C7-C8")]
    for i, (lab, v) in enumerate(params):
        r = 3 + i
        ws.cell(r, 2, lab).font = Font(bold=True)
        c = ws.cell(r, 3, v)
        c.font = Font(color="000000" if str(v).startswith("=") else "0000FF", bold=str(v).startswith("="))
        for cc in (2, 3):
            ws.cell(r, cc).border = bd
            ws.cell(r, cc).fill = grey if cc == 2 else PatternFill()
    # C3 plan, C4 FG, C5 WIP, C6 SP WIP, C7 SP plan, C8 SP FG, C9 Bal P1R5, C10 Bal SP
    H = 12
    heads = ["Sl", "Part No", "Description", "P1R5 Qty", "Sing Phs Qty", "Uom", "Buyer", "Supplier", "Project", "Insp",
             "Store Stock", "WIP", "P1R5 Build", "SMPL-3 Stock", "Coverage in Sets", "Total Shortage in Set", "Remarks"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(H, j, h)
        c.font, c.fill, c.border = Font(bold=True, color="FFFFFF"), navy, bd
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[H].height = 32
    red = Font(color="C00000", bold=True)
    for i, r in enumerate(recs):
        x = H + 1 + i
        kind = _pbom_kind(r.get("project"))
        vals = [r.get("sl"), r.get("part_no"), r.get("description"), num(r.get("p1r5_qty")), num(r.get("sp_qty")), r.get("uom"),
                r.get("buyer"), r.get("supplier"), r.get("project"), num(r.get("insp")), num(r.get("store_stock")), num(r.get("wip"))]
        if kind == "common":
            f = [f"=($C$5+$C$6)*D{x}", f"=(K{x}+L{x})-M{x}", f"=IFERROR(N{x}/D{x},\"\")", f"=IFERROR(O{x}-($C$9+$C$10),\"\")"]
        elif kind == "sp":
            f = [f"=$C$6*E{x}", f"=(K{x}+L{x})-M{x}", f"=IFERROR(N{x}/E{x},\"\")", f"=IFERROR(O{x}-$C$10,\"\")"]
        else:
            f = [f"=$C$5*D{x}", f"=(K{x}+L{x})-M{x}", f"=IFERROR(N{x}/D{x},\"\")", f"=IFERROR(O{x}-$C$9,\"\")"]
        for j, v in enumerate(vals + f + [r.get("remarks")], 1):
            c = ws.cell(x, j, v)
            c.border = bd
            if j in (15, 16):
                c.number_format = "0.00;[Red]-0.00"
            if j in (3, 8, 17):
                c.alignment = Alignment(wrap_text=True, vertical="center")
    widths = [6, 13, 34, 9, 9, 7, 11, 30, 11, 7, 10, 9, 10, 11, 11, 12, 26]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.freeze_panes = ws.cell(H + 1, 3)
    ws.auto_filter.ref = f"A{H}:Q{H + max(len(recs), 1)}"
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def page_pbom():
    p = pbom_params()
    header("Split AC PBOM – Buyer View", f"{p['title']} · buyer {p['buyer']} · columns Sl to Remarks with the sheet's formulas")
    with st.expander("Plan inputs (the sheet's header cells)", expanded=False):
        with st.form("pbom_par"):
            c = st.columns(6)
            vals = {k: c[i].number_input(lab, value=float(num(p[k]) or 0), step=1.0, key=f"pp_{k}")
                    for i, (k, (lab, _)) in enumerate(PBOM_PARAMS.items())}
            if st.form_submit_button("Save plan inputs", type="primary"):
                p.update(vals)
                commit("PBOM plan inputs updated")
                st.rerun()
        st.caption(f"Balance P1R5 = Plan − FG − WIP = **{p['bal_p1r5']:g}** · "
                   f"Balance Single Phs = SP plan − SP FG = **{p['bal_sp']:g}**")
    if st.session_state.pop("pbom_goto", None):
        st.session_state.pbom_view = "Sheet view"
    view = st.radio("View", ["Sheet view", "Edit (grid)", "Import from Master PBOM"], horizontal=True,
                    label_visibility="collapsed", key="pbom_view")
    recs = D()["pbom"]
    if view == "Import from Master PBOM":
        st.caption("Upload the Split AC + Delta Charger workbook. Only main part rows are read (helper 'Receipt plan' rows "
                   "are skipped). Store Stock, WIP and Insp come from the sheet's last calculated values, because in Excel "
                   "they are VLOOKUPs into another PBOM file.")
        f = st.file_uploader("Workbook (.xlsx)", type=["xlsx", "xlsm"], key="pbom_up")
        if not f:
            return
        names, sheet, _ = pbom_read(f.getvalue())
        sheet = st.selectbox("Sheet", names, index=names.index(sheet))
        _, _, rows = pbom_read(f.getvalue(), sheet)
        if rows is None:
            st.error("No header row starting with 'Sl' and 'Part No' found in the first 40 rows of this sheet.")
            return
        buyers = sorted({str(r.get("buyer") or "").strip() for r in rows if r.get("buyer")})
        bi = next((i for i, b in enumerate(buyers) if b.lower().startswith(str(p["buyer"]).lower())), 0)
        buyer = st.selectbox("Buyer", buyers, index=bi)
        sel = [r for r in rows if str(r.get("buyer") or "").strip().lower() == buyer.lower()]
        st.write(f"**{len(sel)}** parts for {buyer} in '{sheet}'.")
        prev = pd.DataFrame([{**{"Sl": r.get("sl"), "Part No": r.get("part_no"), "Description": r.get("description"),
                                 "Project": r.get("project"), "Store Stock": r.get("store_stock"), "WIP": r.get("wip")},
                              **pbom_calc(r)} for r in sel])
        st.dataframe(prev, hide_index=True, width="stretch", height=300)
        mode = st.radio("Import mode", ["Replace the current list", "Add new parts only (skip existing Part No)"], horizontal=True)
        if st.button(f"Import {len(sel)} parts", type="primary", disabled=not sel):
            if mode.startswith("Replace"):
                D()["pbom"] = [{"_id": new_id(), **r} for r in sel]
                n = len(sel)
            else:
                have = {str(r.get("part_no")) for r in recs}
                new = [{"_id": new_id(), **r} for r in sel if str(r.get("part_no")) not in have]
                D()["pbom"].extend(new)
                n = len(new)
            p["buyer"] = buyer
            commit(f"PBOM: imported {n} parts ({buyer})")
            st.session_state.pbom_goto = "Sheet view"
            st.rerun()
        return
    if view == "Edit (grid)":
        excel_grid(spec_coll("pbom"), computed=pbom_calc, height=520,
                   dropdowns={"project": ["Common", "P1R5", "Single Phs"]},
                   help_text="Grey columns on the right are calculated with the sheet's formulas. "
                             "Project decides the formula: Common / P1R5 / Single Phs.")
        return
    if not recs:
        st.info("No parts yet. Use 'Import from Master PBOM' to load the buyer's parts from the workbook.")
        return
    rows = []
    for r in recs:
        c = pbom_calc(r)
        rows.append({"Sl": r.get("sl"), "Part No": r.get("part_no"), "Description": r.get("description"),
                     "P1R5 Qty": num(r.get("p1r5_qty")), "Sing Phs Qty": num(r.get("sp_qty")), "Uom": r.get("uom"),
                     "Buyer": r.get("buyer"), "Supplier": r.get("supplier"), "Project": r.get("project"),
                     "Insp": num(r.get("insp")), "Store Stock": num(r.get("store_stock")), "WIP": num(r.get("wip")),
                     **c, "Remarks": r.get("remarks")})
    df = pd.DataFrame(rows)
    q = st.text_input("Search", key="pbom_q", placeholder="Part no, description, supplier…")
    only_short = st.checkbox("Only parts short (Total Shortage in Set < 0)", key="pbom_short")
    if q:
        df = df[df.apply(lambda x: q.lower() in " ".join(map(str, x.values)).lower(), axis=1)]
    if only_short:
        df = df[pd.to_numeric(df["Total Shortage in Set"], errors="coerce") < 0]
    sh = pd.to_numeric(pd.DataFrame(rows)["Total Shortage in Set"], errors="coerce")
    c = st.columns(4)
    c[0].metric("Parts", len(rows))
    c[1].metric("Short (< 0 sets)", int((sh < 0).sum()))
    c[2].metric("Covered", int((sh >= 0).sum()))
    c[3].metric("Worst shortage (sets)", f"{sh.min():,.1f}" if sh.notna().any() else "—")
    st.dataframe(df.style.map(lambda v: "color:#C00000;font-weight:600" if isinstance(v, (int, float)) and v < 0 else "",
                              subset=["Total Shortage in Set"]).format(precision=2, na_rep=""),
                 hide_index=True, width="stretch", height=560)
    st.download_button("⬇ Download Excel (with formulas)", pbom_excel(recs, p),
                       f"PBOM_{p['buyer']}_{datetime.now():%Y%m%d}.xlsx", type="primary")


# ----------------------------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------------------------
def page_parts_master():
    v = st.radio("List", ["MQIS & Single Phase", "6DS Project"], horizontal=True, key="pm_view", label_visibility="collapsed")
    page_parts("mqis" if v.startswith("MQIS") else "ds6")


def page_tech():
    v = st.radio("Comparison", ["DC Fan Comparison", "Teknic Parts (Domestic)"], horizontal=True, key="tc_view",
                 label_visibility="collapsed")
    page_fan() if v.startswith("DC") else page_teknic()


PAGES = {
    "📊 Dashboard": page_dashboard,
    "🗂 NPD Management": page_projects,
    "⏳ NPD Pending & Completed": page_activities,
    "🔩 Parts Master": page_parts_master,
    "📋 Split AC PBOM (Buyer)": page_pbom,
    "🏭 Supplier Management": page_suppliers_master,
    "📨 RFQ Management": page_rfq,
    "🧾 Purchase Orders": page_po,
    "✉️ Email Approval Centre": page_email_centre,
    "🔔 Supplier Follow-ups": page_followups,
    "📦 Inventory & Shortage": page_inventory,
    "⚖️ Technical Comparison": page_tech,
    "🧪 DV/DVVP Tracking": page_dv,
    "🔁 ECN/ECR Tracking": page_ecn,
    "🤖 AI Assistant": page_ai,
    "📑 Reports": page_reports,
    "⚙️ Settings & Data": page_settings,
}


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="⚙️", layout="wide", initial_sidebar_state="expanded")
    st.markdown(CSS, unsafe_allow_html=True)
    init_state()
    ss = st.session_state
    with st.sidebar:
        st.markdown(f"### ⚙️ {APP_TITLE}")
        st.caption("E&E BOP · NPD, sourcing & procurement · Sun Mobility CPMO")
        nd = proc_kpis()["drafts"] if has_data() else 0
        if nd:
            st.caption(f"✉️ {nd} email(s) awaiting approval")
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
