#!/usr/bin/env python3
"""
Work out which carrier a container number belongs to, and split a spreadsheet
into per-carrier lists.

    python carrier_id.py TCLU7686420 HASU1149124 MRKU7621507
    python carrier_id.py --excel containers.xlsx --outdir ./lists

HOW RELIABLE IS THE PREFIX?
---------------------------
The first four letters are the BIC owner code -- who OWNS the box, which is not
the same as who is CARRYING it.

  * Carrier-owned prefixes (MSCU, MAEU, MRKU, ...) are a strong hint. A box
    Maersk owns is nearly always on a Maersk bill of lading.

  * LEASING prefixes (TCLU, CAIU, TEMU, TGHU, ...) tell you nothing about the
    carrier. Your own data proves it: TCLU7686420 and CAIU3056659 are both
    leasing codes and both tracked fine on MSC.

So treat this as routing hint plus fallback, never as proof. The only real
confirmation is a successful lookup. The classifier returns a confidence level
so you can decide whether to trust it or just try every carrier.

WHY YOUR SIX FAILED
-------------------
    HASU1149124   Hamburg Sud (Maersk group)   -> Maersk
    MRKU7621507   Maersk                        -> Maersk
    MRSU0444608   Maersk                        -> Maersk
    MRSU0787675   Maersk                        -> Maersk
    MSKU2876636   Maersk                        -> Maersk
    TLLU2467264   leasing                       -> unknown, try both

Five are Maersk. The sixth is a leased box that could be on anything.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path

CONTAINER_RE = re.compile(r"^([A-Z]{4})(\d{7})$")

# Bill of lading / booking formats seen in the wild:
#   MEDUXK845261   MSC, 4 letters + alnum
#   MAEU123456789  Maersk prefixed
#   260029935      Maersk bare numeric, 9-12 digits
BL_ALNUM_RE = re.compile(r"^[A-Z]{3,4}[A-Z0-9]{6,13}$")
BL_NUMERIC_RE = re.compile(r"^\d{9,12}$")

HEADER_HINTS = ("container", "cntr", "equipment", "box", "bl", "b/l",
                "bill of lading", "booking", "tracking", "number", "no.")


def is_trackable_bl(value: str) -> bool:
    v = value.strip().upper()
    return bool(BL_ALNUM_RE.match(v) or BL_NUMERIC_RE.match(v))

# Carrier-owned prefixes. High confidence: if the carrier owns the box, it is
# almost certainly moving on their service.
CARRIER_PREFIXES: dict[str, str] = {}

for _prefix in ("MAEU", "MRKU", "MSKU", "MRSU", "MSFU", "MWCU", "MNBU",
                "POCU", "MIEU", "MVIU"):
    CARRIER_PREFIXES[_prefix] = "MAEU"          # Maersk
for _prefix in ("HASU", "SUDU"):
    CARRIER_PREFIXES[_prefix] = "MAEU"          # Hamburg Sud, Maersk group
for _prefix in ("MSCU", "MEDU", "MSBU", "MSDU", "MSMU"):
    CARRIER_PREFIXES[_prefix] = "MSCU"          # MSC
for _prefix in ("CMAU", "CGMU", "ECMU", "APZU"):
    CARRIER_PREFIXES[_prefix] = "CMDU"          # CMA CGM
for _prefix in ("HLCU", "HLXU", "UACU"):
    CARRIER_PREFIXES[_prefix] = "HLCU"          # Hapag-Lloyd
for _prefix in ("OOLU", "OOCL"):
    CARRIER_PREFIXES[_prefix] = "OOLU"
for _prefix in ("COSU", "CBHU", "CCLU"):
    CARRIER_PREFIXES[_prefix] = "COSU"          # COSCO
for _prefix in ("ONEU", "NYKU", "MOLU", "KKLU"):
    CARRIER_PREFIXES[_prefix] = "ONEY"          # ONE
for _prefix in ("EGHU", "EISU", "EGSU", "EMCU"):
    CARRIER_PREFIXES[_prefix] = "EGLV"          # Evergreen
for _prefix in ("YMLU", "YMMU", "YMAU"):
    CARRIER_PREFIXES[_prefix] = "YMLU"          # Yang Ming
for _prefix in ("HDMU",):
    CARRIER_PREFIXES[_prefix] = "HDMU"          # HMM
for _prefix in ("ZIMU",):
    CARRIER_PREFIXES[_prefix] = "ZIMU"

# Leasing companies. These say nothing about the carrier.
LEASING_PREFIXES = {
    "TCLU", "TGHU", "TRHU", "TRLU", "TEMU", "TGBU",   # Triton / Textainer
    "CAIU", "CAXU",                                    # CAI
    "GESU", "GATU", "GLDU",                            # SeaCube / others
    "BEAU", "BMOU", "BSIU",                            # Beacon
    "FCIU", "FSCU",                                    # Florens
    "SEGU", "SEKU",                                    # Seaco
    "DFSU", "DRYU",                                    # Dong Fang
    "UETU", "UACU",                                    # Caru / others
    "TLLU",                                            # leasing; see note below
    "CARU", "CRXU", "CRSU", "INKU", "INBU", "XINU",
}

CARRIER_NAMES = {
    "MAEU": "Maersk (incl. Hamburg Sud)",
    "MSCU": "MSC",
    "CMDU": "CMA CGM",
    "HLCU": "Hapag-Lloyd",
    "OOLU": "OOCL",
    "COSU": "COSCO",
    "ONEY": "ONE",
    "EGLV": "Evergreen",
    "YMLU": "Yang Ming",
    "HDMU": "HMM",
    "ZIMU": "ZIM",
}

# Bill of lading prefixes -- more reliable than container prefixes, because the
# BL is issued BY the carrier.
BL_PREFIXES = {
    "MEDU": "MSCU",     # MSC, e.g. MEDUXK845261
    "MSCU": "MSCU",
    "MAEU": "MAEU",
    "MRKU": "MAEU",
    "SUDU": "MAEU",
    "CMDU": "CMDU",
    "HLCU": "HLCU",
    "OOLU": "OOLU",
    "COSU": "COSU",
    "ONEY": "ONEY",
    "EGLV": "EGLV",
    "YMLU": "YMLU",
}


def classify(number: str) -> dict:
    """-> {number, kind, carrier, carrier_name, confidence, note}"""
    n = number.strip().upper().replace(" ", "")

    m = CONTAINER_RE.match(n)
    if m:
        prefix = m.group(1)
        if prefix in CARRIER_PREFIXES:
            carrier = CARRIER_PREFIXES[prefix]
            return {"number": n, "kind": "container", "carrier": carrier,
                    "carrier_name": CARRIER_NAMES.get(carrier, carrier),
                    "confidence": "high",
                    "note": f"{prefix} is a carrier-owned prefix"}
        if prefix in LEASING_PREFIXES:
            return {"number": n, "kind": "container", "carrier": None,
                    "carrier_name": None, "confidence": "unknown",
                    "note": f"{prefix} is a leasing prefix -- could be any carrier"}
        return {"number": n, "kind": "container", "carrier": None,
                "carrier_name": None, "confidence": "unknown",
                "note": f"{prefix} not in the lookup table"}

    # Not a container number -- try BL/booking prefixes.
    for prefix, carrier in BL_PREFIXES.items():
        if n.startswith(prefix):
            return {"number": n, "kind": "bill_of_lading", "carrier": carrier,
                    "carrier_name": CARRIER_NAMES.get(carrier, carrier),
                    "confidence": "high",
                    "note": f"{prefix} is a {CARRIER_NAMES.get(carrier)} BL prefix"}

    return {"number": n, "kind": "unknown", "carrier": None, "carrier_name": None,
            "confidence": "unknown", "note": "not a recognisable container or BL"}


def read_excel(path: Path, column: str | None, sheet: str | None) -> list[str]:
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        raise SystemExit("Empty sheet.")

    header = [str(c).strip() if c is not None else "" for c in rows[0]]

    def values(idx: int) -> list[str]:
        return [str(r[idx]).strip().upper().replace(" ", "")
                for r in rows[1:]
                if idx < len(r) and r[idx] is not None and str(r[idx]).strip()]

    if column:
        idx = next((i for i, h in enumerate(header) if h.lower() == column.lower()), None)
        if idx is None:
            raise SystemExit(f"No column {column!r}. Found: {[h for h in header if h]}")
    else:
        # Score containers AND bills of lading. A BL-only sheet used to fail
        # here because only CONTAINER_RE was counted.
        idx, best = None, 0
        for i in range(len(header)):
            vals = values(i)
            score = sum(1 for v in vals if CONTAINER_RE.match(v) or is_trackable_bl(v))
            if any(h in header[i].lower() for h in HEADER_HINTS):
                score += 0.5
            if score > best:
                idx, best = i, score
        if idx is None or best < 1:
            raise SystemExit(
                "No column of container or BL numbers found; use --column.\n"
                f"Headers seen: {[h for h in header if h]}")
        print(f"# column: {header[idx]!r}", file=sys.stderr)

    seen, out = set(), []
    for v in values(idx):
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Route container numbers to carriers.")
    ap.add_argument("numbers", nargs="*", help="numbers to classify")
    ap.add_argument("--excel", type=Path, help="read numbers from a spreadsheet")
    ap.add_argument("--column")
    ap.add_argument("--sheet")
    ap.add_argument("--outdir", type=Path, help="write one .txt list per carrier")
    args = ap.parse_args()

    numbers = list(args.numbers)
    if args.excel:
        numbers += read_excel(args.excel, args.column, args.sheet)
    if not numbers:
        ap.print_help()
        return 2

    buckets: dict[str, list[str]] = defaultdict(list)
    print(f"\n{'NUMBER':<15} {'KIND':<15} {'CARRIER':<28} {'CONF':<8} NOTE")
    print("-" * 100)
    for raw in numbers:
        r = classify(raw)
        buckets[r["carrier"] or "UNKNOWN"].append(r["number"])
        print(f"{r['number']:<15} {r['kind']:<15} "
              f"{(r['carrier_name'] or '?'):<28} {r['confidence']:<8} {r['note']}")

    print("\nSummary:")
    for carrier, items in sorted(buckets.items()):
        name = CARRIER_NAMES.get(carrier, carrier)
        print(f"  {name:<30} {len(items)}")

    if args.outdir:
        args.outdir.mkdir(parents=True, exist_ok=True)
        for carrier, items in buckets.items():
            path = args.outdir / f"{carrier.lower()}.txt"
            path.write_text("\n".join(items) + "\n", encoding="utf-8")
            print(f"  wrote {path} ({len(items)})")

    if "UNKNOWN" in buckets:
        print(f"\n{len(buckets['UNKNOWN'])} unresolved -- leased or unlisted prefixes.")
        print("Try them against each carrier you ship with; first hit wins.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())