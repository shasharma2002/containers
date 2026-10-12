#!/usr/bin/env python3
"""
Parse the multi-company master workbook (all_companies.xlsx) into flat shipment
records, one per sheet row, ready to join against carrier tracking data on
bl_number.

    python read_master.py all_companies.xlsx
    python read_master.py all_companies.xlsx --sheet NICPL --json
    python read_master.py all_companies.xlsx --show-skipped

One sheet per company (NICPL, PBT, PROCHEM, ...), so the COMPANY IS THE SHEET
NAME -- there is no company column to read. New company sheets are picked up
automatically; nothing here lists them.

WHAT THIS WORKBOOK WILL DO TO A NAIVE PARSER
--------------------------------------------
1. NOT EVERY SHEET IS A COMPANY. The same file holds three other layouts:

     archive        'material reach 82-83', 'material reach 79-80' -- company
                    is a COLUMN, and a 'Bir arriv' column says the cargo already
                    reached Birganj. Historical, 2022-23, and carriers purge
                    tracking data long before that.
     fiscal archive '078-79' (header GROUP/SHIPPER/BL NO., and it DOES have a
                    SHIPPING LINE column), and 'material reached 78-79', which
                    has no header row at all.
     bulk tanker    'OIL TANKER' -- 'BL NO' values are 39, 40, 87 and vessels
                    are 'MT ...' motor tankers. Bulk liquid parcels in shore
                    tanks; there is no container and nothing to track.

   Sheets are therefore CLASSIFIED by what their headers contain, not by name,
   and anything not a live company sheet is skipped and REPORTED. Never
   silently: `skipped` in the result says which sheet and why.

2. NO SHIPPING LINE COLUMN. bl_list.xlsx had one and that is the only reason
   bare-numeric Maersk BLs ever routed correctly. These sheets have none, so
   the carrier has to come from the BL's own shape -- see carrier_id.classify,
   which now returns a medium-confidence Maersk hint for a bare 9-12 digit BL.
   Without that, 11 of the ~23 live BLs here are dropped as "no ocean carrier"
   without a lookup ever being attempted.

   Occasionally the carrier IS in the file, jammed into the BL cell:
   'HLCUMTR260616918,\\nhappag'. That text is pulled out into
   shipping_line_raw rather than thrown away.

3. THE BL CELL IS NOT ALWAYS JUST A BL. Seen in this file:
       'HLCUMTR260616918,\\nhappag'            BL + carrier name
       'HLCUMTR260830603,\\nCONFCIU6646940'    BL + a container number
       'ELG/RGN/CCU-1638/26,ELAN,TLPL'        forwarder references, no BL
   Cells are split on commas and newlines, the most BL-looking token wins, and
   whatever is left over is kept in bl_cell_extra so nothing is lost.

4. NICPL CARRIES 1,048,485 PHANTOM BLANK ROWS from stale Excel formatting.
   Reading them all is slow and pointless, so the scan stops after a run of
   consecutive empty rows (BLANK_RUN_LIMIT).

5. 'FREE TIME' MIXES TWO UNITS. Some rows hold a day count ('14'), others an
   expiry date ('04.10.26'). Both are real, so they go to different fields --
   free_days and free_time_expiry -- rather than one being coerced into the
   other.

6. 'RATE/AMT' IS A RATE, NOT A TOTAL. Values are 590, 645, 617 -- USD per
   tonne against quantities in the hundreds. bl_list.xlsx's AMOUNT column was a
   total (223574.94). Mapping one onto the other would misstate cargo value by
   orders of magnitude, so the rate is kept as `rate` and the product is
   exposed separately as `amount_derived`, named so nobody mistakes it for a
   figure the sheet actually stated.

7. THE DATE COLUMNS ARE NOT RELIABLE AND ARE NOT USED. 'ETA Viz/Kol', 'ETA
   BRJ', 'LC DATE', 'BL DATE' and a dated 'FREE TIME' cell are all typed by
   hand and are not maintained against the carriers. They are read, kept as
   raw text for audit, and then suppressed -- see TRUST_SHEET_DATES. Arrival
   dates on the report come from the carriers only.

   The one number that survives is `free_days`: a day COUNT, not a date, and a
   contractual term rather than a guess. Carrier discharge + free_days still
   gives a sound detention clock.
"""

from __future__ import annotations

import argparse
import datetime
import json
import re
import sys
from pathlib import Path

HEADER_SCAN_ROWS = 8
# Stop after this many consecutive empty rows. NICPL's phantom formatting
# produces over a million of them.
BLANK_RUN_LIMIT = 40

# THE DATES ON THIS WORKBOOK ARE NOT TRUSTED.
#
# The sheet owner's instruction, and the data agrees with it: the ETA columns
# are filled in by hand, go stale as soon as a vessel reschedules, and are not
# maintained against the carriers. Every date on this workbook is therefore
# parsed, kept as RAW TEXT for audit, and then deliberately NOT turned into a
# date the report can calculate with. The only arrival dates the report uses
# come from the carriers themselves.
#
# What this flag does NOT suppress is `free_days` -- the number in the FREE TIME
# column. That is a count of days, not a date: a contractual term agreed with
# the line, which does not rot the way a typed ETA does. Combined with the
# CARRIER's discharge event it still gives a sound free-time expiry, so
# detention tracking keeps working. See _row_record().
#
# Set True to put the sheet's own dates back into play; every suppressed field
# is listed in SHEET_DATE_FIELDS so the effect is visible in one place.
TRUST_SHEET_DATES = False

# Parsed date fields that TRUST_SHEET_DATES governs. The matching `*_raw` text
# is always kept, so turning the flag off loses nothing -- it only stops a
# hand-typed date being treated as fact.
SHEET_DATE_FIELDS = ("sheet_eta_date", "eta_port_date", "free_time_expiry",
                     "lc_date", "bl_date")

# Header text, normalised to letters+digits, that identifies the BL column.
BLNO_ALIASES = {"BLNO", "BLNUMBER", "BLNOS", "BILLOFLADINGNO", "BILLOFLADING",
                "BLNOLRNO", "BLNOLRNO", "BLLRNO"}

# Columns that mark a sheet as one of the ARCHIVE layouts rather than a live
# per-company sheet. Company-as-a-column is the giveaway: a live sheet gets its
# company from the tab name.
ARCHIVE_MARKERS = {"COMPANY", "GROUP"}
# Arrival-at-destination columns -- the cargo is already here.
ARRIVED_MARKERS = {"BIRARRIV", "BIRARRIVAL", "BIRARRIVED"}
# Bulk-tanker markers: shore tanks and custody transfer, not containers.
TANKER_MARKERS = {"TANKNO", "CTDNO", "LOADEDQTY", "BALANCEQTY"}

# Everything we know how to read off a live company sheet. Generous aliases,
# because the next revision of this workbook will not spell them identically.
HEADER_ALIASES: dict[str, str] = {
    "SNO": "serial", "SN": "serial", "SLNO": "serial",
    "SUPPLIER": "supplier", "SHIPPER": "supplier", "BENEFICIARY": "supplier",
    "LCNO": "lc_no", "LCDETAILS": "lc_no", "LCNUMBER": "lc_no",
    "LCDATE": "lc_date",
    "PINO": "pi_no", "PINUMBER": "pi_no",
    "ORIGIN": "origin",
    "BLDATE": "bl_date",
    "PORTOFLOADING": "port_of_loading", "POL": "port_of_loading",
    "MATERIAL": "commodity", "COMMODITY": "commodity", "COMMODITIES": "commodity",
    "ITEM": "commodity",
    "QTYMT": "qty", "QTY": "qty", "QUANTITY": "qty",
    "RATEAMT": "rate", "RATE": "rate", "AMOUNT": "amount", "AMT": "amount",
    "CONTNOS": "container_count", "NOOFCONT": "container_count",
    "NOOFCONTAINER": "container_count", "NOOFCONTAINERS": "container_count",
    "CONTAINER": "container_count",
    "CONTSIZE": "container_size", "CONTAINERSIZE": "container_size",
    "VIZKOL": "eta_port_raw", "ETAVIZKOL": "eta_port_raw", "KOLVIZ": "eta_port_raw",
    "ETABRJ": "eta_raw", "ETA": "eta_raw", "BIRARRIV": "eta_raw",
    "FREETIME": "free_time_raw", "FREEDAYS": "free_time_raw",
    "DOCUMENTPOSITION": "document_position", "DOCPOSITION": "document_position",
    "ORIGINALDOCRECDATE": "doc_received", "ORGDOCRECEIVING": "doc_received",
    "IPNOEXPDT": "ip_no", "IPNOEXPDT2": "ip_no",
    "SHIPPINGLINE": "shipping_line_raw", "CARRIER": "shipping_line_raw",
    "DELIVERY": "delivery_terms",
    "BROKER": "broker",
}
for _alias in BLNO_ALIASES:
    HEADER_ALIASES[_alias] = "bl_number_raw"

# A BL looks like one of these. Ordered by how strongly it identifies a carrier,
# because the best-looking token in a messy cell wins.
BL_PREFIXED = re.compile(r"^[A-Z]{4}[A-Z0-9]{6,14}$")
BL_NUMERIC = re.compile(r"^\d{9,12}$")
CONTAINER_RE = re.compile(r"^[A-Z]{4}\d{7}$")

# Free-standing carrier words that turn up inside a BL cell.
CARRIER_WORDS = re.compile(
    r"^(HAPPAG|HAPAG|HAPAGLLOYD|MAERSK|MSC|COSCO|CMACGM|CMA|ONE|EVERGREEN|"
    r"YANGMING|YML|OOCL|HMM|ZIM|CONF|PENDING|PEND)$", re.I)

DATE_PATTERNS = (
    # 31.07.26 / 04-10-26 / 15/7/2022 -- day first, which is what this sheet uses
    re.compile(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$"),
)


class MasterSheetError(RuntimeError):
    pass


# --------------------------------------------------------------------------- #
# cell helpers
# --------------------------------------------------------------------------- #

def _norm(value) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value if value is not None else "").upper())


def _clean(v):
    """Strip strings, pass numbers through, make dates JSON-safe."""
    if v is None:
        return None
    if isinstance(v, str):
        s = " ".join(v.split())
        return s or None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()
    return v


def _cell(row: tuple, idx: int | None):
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def _token_strings(value) -> list[str]:
    """One messy cell -> its candidate tokens, upper-cased, spaces removed.

    Excel stores a numeric-looking BL as a float, so 272573330 arrives as
    272573330.0 and str() would give '272573330.0'.
    """
    if value is None:
        return []
    if isinstance(value, bool):
        return []
    if isinstance(value, int):
        return [str(value)]
    if isinstance(value, float):
        return [str(int(value)) if value.is_integer() else str(value)]
    if isinstance(value, (datetime.datetime, datetime.date)):
        return []
    parts = re.split(r"[,\n\r;/|]+", str(value))
    return [p.strip().upper().replace(" ", "") for p in parts if p and p.strip()]


def split_bl_cell(value) -> tuple[str | None, str | None, str | None]:
    """Messy BL cell -> (bl_number, carrier_text, leftover).

    The cell is split on commas and newlines and the best-looking BL token is
    taken. A recognisable carrier word becomes carrier_text, which is the only
    place on these sheets the shipping line is ever stated. Anything else is
    returned as leftover so a forwarder reference is visible rather than
    discarded.

        'HLCUMTR260616918,\\nhappag'        -> ('HLCUMTR260616918', 'happag', None)
        'HLCUMTR260830603,\\nCONFCIU6646940'-> ('HLCUMTR260830603', 'CONF', 'CIU6646940')
        '272573330'                        -> ('272573330', None, None)
        'ELG/RGN/CCU-1638/26,ELAN,TLPL'    -> (None, None, 'ELG, RGN, ...')
    """
    tokens = _token_strings(value)
    if not tokens:
        return None, None, None

    bl = carrier = None
    leftover: list[str] = []
    for token in tokens:
        if bl is None and (BL_PREFIXED.match(token) and not CONTAINER_RE.match(token)):
            bl = token
        elif bl is None and BL_NUMERIC.match(token):
            bl = token
        elif carrier is None and CARRIER_WORDS.match(token):
            carrier = token
        else:
            leftover.append(token)

    # A container number is not a BL, but if the cell holds nothing better it is
    # still trackable -- the carriers accept either.
    if bl is None:
        for token in list(leftover):
            if CONTAINER_RE.match(token):
                bl = token
                leftover.remove(token)
                break

    return bl, carrier, (", ".join(leftover) or None)


def parse_sheet_date(raw) -> str | None:
    """'31.07.26' -> '2026-07-31'. Day-first, two-digit years windowed to 2000s.

    Returns None rather than guessing when the value is not a date -- callers
    keep the raw text alongside, so an unparseable entry stays visible.
    """
    if raw is None:
        return None
    if isinstance(raw, (datetime.datetime, datetime.date)):
        return raw.strftime("%Y-%m-%d")
    text = str(raw).strip()
    if not text:
        return None
    # A cell can hold 'ok, 20.09.26' or '19 / 20.01.022' -- take the first thing
    # shaped like a date.
    for chunk in re.split(r"[,\s]+", text):
        for pattern in DATE_PATTERNS:
            m = pattern.match(chunk)
            if not m:
                continue
            day, month, year = (int(g) for g in m.groups())
            # This sheet also writes years as '022' meaning 2022.
            if year < 100:
                year += 2000
            elif 100 <= year < 1000:
                year += 2000
            if not (1 <= month <= 12 and 1 <= day <= 31):
                continue
            try:
                return datetime.date(year, month, day).isoformat()
            except ValueError:
                continue
    return None


def _leading_int(value) -> int | None:
    """'1(12)' -> 1, '20' -> 20, '19 x 40' -> 19. None when there is no number."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    m = re.match(r"\s*(\d+)", str(value))
    return int(m.group(1)) if m else None


def _number(value) -> float | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    m = re.search(r"-?\d+(?:\.\d+)?", str(value).replace(",", ""))
    return float(m.group(0)) if m else None


# --------------------------------------------------------------------------- #
# sheet classification
# --------------------------------------------------------------------------- #

def classify_sheet(rows: list[tuple]) -> tuple[str, int | None, dict, str]:
    """-> (layout, header_row_index, {field: column}, reason)

    layout is one of:
        "company"  a live per-company INTRANSIT sheet -- parse it
        "archive"  company-as-a-column, or an arrival date: historical
        "tanker"   bulk liquid in shore tanks: no containers
        "noheader" no BL column found at all
    """
    for index, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        present = {_norm(c): j for j, c in enumerate(row) if c is not None}
        if not (present.keys() & BLNO_ALIASES):
            continue

        columns = {}
        for norm_header, column in present.items():
            field = HEADER_ALIASES.get(norm_header)
            if field and field not in columns:
                columns[field] = column

        if present.keys() & TANKER_MARKERS:
            return "tanker", index, columns, (
                "bulk liquid cargo in shore tanks (tank/CTD columns present) -- "
                "no containers, nothing a carrier can track")
        if present.keys() & ARCHIVE_MARKERS:
            return "archive", index, columns, (
                "company is a column, not the sheet name -- this is an archive "
                "layout, not a live company sheet")
        if present.keys() & ARRIVED_MARKERS:
            return "archive", index, columns, (
                "has an arrival-at-Birganj column -- this cargo has already "
                "landed, and carriers purge tracking data within months")
        return "company", index, columns, "live per-company INTRANSIT sheet"

    return "noheader", None, {}, (
        "no BL column found in the first %d rows -- refusing to guess which "
        "column holds the BL" % HEADER_SCAN_ROWS)


# --------------------------------------------------------------------------- #
# parsing
# --------------------------------------------------------------------------- #

def _sheet_rows(ws) -> list[tuple]:
    """Rows up to the real end of the data, ignoring phantom formatting rows."""
    out: list[tuple] = []
    blank_run = 0
    for row in ws.iter_rows(values_only=True):
        if any(c is not None and str(c).strip() for c in row):
            blank_run = 0
            out.append(row)
        else:
            blank_run += 1
            if blank_run >= BLANK_RUN_LIMIT:
                break
            out.append(row)
    while out and not any(c is not None and str(c).strip() for c in out[-1]):
        out.pop()
    return out


def _row_record(company: str, row: tuple, columns: dict, row_number: int) -> dict | None:
    """One sheet row -> one shipment record, or None when there is nothing in it."""
    bl_raw = _cell(row, columns.get("bl_number_raw"))
    bl_number, carrier_text, leftover = split_bl_cell(bl_raw)

    supplier = _clean(_cell(row, columns.get("supplier")))
    commodity = _clean(_cell(row, columns.get("commodity")))
    qty = _number(_cell(row, columns.get("qty")))

    if not any([bl_number, supplier, commodity, qty]):
        return None

    rate = _number(_cell(row, columns.get("rate")))
    amount = _number(_cell(row, columns.get("amount")))

    free_raw = _clean(_cell(row, columns.get("free_time_raw")))
    # FREE TIME holds either a day count ('14') or an expiry date ('04.10.26').
    # The count survives -- it is a contractual term, not a typed date. The date
    # is parsed only so it can be told apart from a count, then suppressed with
    # the rest of the sheet's dates.
    free_expiry = parse_sheet_date(free_raw)
    free_days = None
    if free_expiry is None:
        candidate = _leading_int(free_raw)
        if candidate is not None and candidate <= 365:
            free_days = candidate

    eta_raw = _clean(_cell(row, columns.get("eta_raw")))
    eta_port_raw = _clean(_cell(row, columns.get("eta_port_raw")))

    # shipping_line_raw is normally absent from these sheets; when present it is
    # either a real column or a carrier word found inside the BL cell.
    shipping_line = (_clean(_cell(row, columns.get("shipping_line_raw")))
                     or carrier_text)

    record = {
        "company": company,
        "row_number": row_number,
        "bl_number": bl_number,
        "bl_cell_raw": _clean(bl_raw) if bl_raw is not None else None,
        "bl_cell_extra": leftover,
        # No SHIPPING LINE column on these sheets -- carrier resolution is left
        # to track_all.py via carrier_id.classify. Deliberately not guessed here.
        "carrier": None,
        "shipping_line_raw": shipping_line,
        "beneficiary": supplier,
        "broker": _clean(_cell(row, columns.get("broker"))),
        "commodity": commodity,
        "pi_no": _clean(_cell(row, columns.get("pi_no"))),
        "qty": qty,
        "rate": rate,
        # The sheet states a RATE per tonne, not a total. This is the product,
        # named so it is never mistaken for a figure the sheet gave us.
        "amount": amount,
        "amount_derived": (round(qty * rate, 2)
                           if amount is None and qty and rate else None),
        "lc_no": _clean(_cell(row, columns.get("lc_no"))),
        "lc_date": parse_sheet_date(_cell(row, columns.get("lc_date"))),
        "bl_date": parse_sheet_date(_cell(row, columns.get("bl_date"))),
        "origin": _clean(_cell(row, columns.get("origin"))),
        "port_of_loading": _clean(_cell(row, columns.get("port_of_loading"))),
        "container_count": _leading_int(_cell(row, columns.get("container_count"))),
        "container_size": _clean(_cell(row, columns.get("container_size"))),
        "delivery_terms": _clean(_cell(row, columns.get("delivery_terms"))),
        "free_days": free_days,
        "free_time_expiry": free_expiry,
        "free_time_raw": free_raw,
        # sheet_eta_* keeps the names read_sheet.py uses, so track_all.py and
        # the reporting layer take either workbook without a special case.
        "sheet_eta_raw": eta_raw,
        "sheet_eta_date": parse_sheet_date(eta_raw),
        "sheet_eta_place": "BRJ" if eta_raw else None,
        "eta_port_raw": eta_port_raw,
        "eta_port_date": parse_sheet_date(eta_port_raw),
        "document_position": _clean(_cell(row, columns.get("document_position"))),
        "doc_received": _clean(_cell(row, columns.get("doc_received"))),
        "ip_no": _clean(_cell(row, columns.get("ip_no"))),
    }

    if not TRUST_SHEET_DATES:
        # Blank every parsed sheet date, keeping the raw text beside it. Done
        # here, once, at the edge of the system -- so nothing downstream has to
        # remember which dates came from a spreadsheet and which came from a
        # carrier. See TRUST_SHEET_DATES for why.
        for field in SHEET_DATE_FIELDS:
            record[field] = None
        record["sheet_eta_place"] = None
    return record


def get_shipments(path: Path, sheet: str | None = None) -> dict:
    """Parse the master workbook.

    -> {"shipments": [...], "sheets": [...], "skipped": [...],
        "skipped_no_bl": n, "skipped_blank": n}

    `sheets` describes every tab that was parsed; `skipped` every tab that was
    not, with the reason. Callers should surface both -- a company sheet that
    silently stopped being read is the failure mode worth guarding against.
    """
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        names = [sheet] if sheet else list(wb.sheetnames)
        if sheet and sheet not in wb.sheetnames:
            raise MasterSheetError(
                "No sheet named %r. Found: %s" % (sheet, ", ".join(wb.sheetnames)))

        shipments: list[dict] = []
        parsed: list[dict] = []
        skipped: list[dict] = []
        skipped_no_bl = skipped_blank = 0

        for name in names:
            rows = _sheet_rows(wb[name])
            if not rows:
                skipped.append({"sheet": name, "layout": "empty",
                                "reason": "no data rows", "rows": 0})
                continue

            layout, header_index, columns, reason = classify_sheet(rows)
            if layout != "company":
                skipped.append({"sheet": name, "layout": layout,
                                "reason": reason, "rows": len(rows)})
                continue

            company = name.strip()
            before = len(shipments)
            for offset, row in enumerate(rows[header_index + 1:],
                                         start=header_index + 2):
                record = _row_record(company, row, columns, offset)
                if record is None:
                    skipped_blank += 1
                    continue
                if not record["bl_number"]:
                    # Real commercial data with no BL yet (or a forwarder
                    # reference instead). Kept -- the row still belongs in the
                    # report, it just cannot be tracked.
                    skipped_no_bl += 1
                shipments.append(record)

            trackable = sum(1 for s in shipments[before:] if s["bl_number"])
            parsed.append({"sheet": name, "company": company,
                           "rows": len(shipments) - before,
                           "trackable": trackable,
                           "header_row": header_index + 1})
    finally:
        wb.close()

    return {"shipments": shipments, "sheets": parsed, "skipped": skipped,
            "skipped_no_bl": skipped_no_bl, "skipped_blank": skipped_blank}


def looks_like_master(path: Path) -> bool:
    """Is this the multi-company master workbook?

    True when at least one tab classifies as a live per-company sheet. Used by
    track_all.py to pick a reader without being told which format it was given.
    """
    from openpyxl import load_workbook

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return False
    try:
        for name in wb.sheetnames:
            rows = _sheet_rows(wb[name])
            if rows and classify_sheet(rows)[0] == "company":
                return True
        return False
    finally:
        wb.close()


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Parse the multi-company master workbook.")
    ap.add_argument("path", type=Path)
    ap.add_argument("--sheet", help="parse only this tab")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--show-skipped", action="store_true",
                    help="explain every tab that was not parsed")
    args = ap.parse_args()

    try:
        result = get_shipments(args.path, args.sheet)
    except MasterSheetError as exc:
        print("\nFAILED: %s" % exc, file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps(result["shipments"], indent=2, ensure_ascii=False))
        return 0

    shipments = result["shipments"]
    print("\nParsed %d company sheet(s), %d row(s), %d with a BL:\n"
          % (len(result["sheets"]), len(shipments),
             sum(1 for s in shipments if s["bl_number"])))
    for info in result["sheets"]:
        print("  %-24s %3d row(s), %3d trackable   (header row %d)"
              % (info["company"], info["rows"], info["trackable"],
                 info["header_row"]))

    if result["skipped"]:
        print("\n%d sheet(s) skipped:" % len(result["skipped"]))
        for info in result["skipped"]:
            print("  %-24s %-9s %s"
                  % (info["sheet"], info["layout"],
                     info["reason"] if args.show_skipped
                     else info["reason"].split(" -- ")[0]))

    print("\n%-10s %-20s %-26s %-18s %8s  %-10s %s"
          % ("COMPANY", "BL", "SUPPLIER", "COMMODITY", "QTY", "ETA BRJ", "NOTE"))
    for s in shipments:
        note = []
        if not s["bl_number"]:
            note.append("no BL")
        if s["shipping_line_raw"]:
            note.append("line=%s" % s["shipping_line_raw"])
        if s["bl_cell_extra"]:
            note.append("extra=%s" % s["bl_cell_extra"][:22])
        print("%-10s %-20s %-26s %-18s %8s  %-12s %s"
              % (s["company"], (s["bl_number"] or "-")[:20],
                 (s["beneficiary"] or "")[:26], (s["commodity"] or "")[:18],
                 s["qty"] if s["qty"] is not None else "",
                 str(s["sheet_eta_raw"] or "-")[:12], "; ".join(note)))

    print("\n%d row(s) with no BL (kept, not trackable), %d blank row(s) skipped."
          % (result["skipped_no_bl"], result["skipped_blank"]), file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
