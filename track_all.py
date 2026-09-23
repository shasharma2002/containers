#!/usr/bin/env python3
"""
Multi-carrier batch tracker. One list in, flat JSON out.

    python track_all.py containers.txt
    python track_all.py containers.xlsx
    python track_all.py TCLU7686420 HASU1149124 TLLU2467264
    python track_all.py containers.xlsx --outdir ./out --append

    # commercial sheet, tracked by Bill of Lading (auto-detected):
    python track_all.py "LC OPEN AND CARGO IN TRANSIT.xlsx"

Needs msc_track.py, maersk_track.py, carrier_id.py and read_sheet.py in the
same folder.

HOW ROUTING WORKS -- container-list mode
-----------------------------------------
1. Classify by BIC prefix (carrier_id.py).
2. Carrier-owned prefix -> go straight to that carrier.
3. Leasing prefix, or a carrier miss -> try every carrier until one answers.

Step 3 is the important one. The prefix tells you who OWNS the box, not who is
carrying it. Your own data proves the point: TCLU7686420 and CAIU3056659 are
leasing prefixes and both live on MSC. So an "unknown" is not a failure -- it
just means we search instead of jumping.

HOW ROUTING WORKS -- sheet mode (BL NO. header auto-detected)
---------------------------------------------------------------
Routing comes from the sheet's own SHIPPING LINE column, not the BL prefix --
see read_sheet.py for why (Maersk BLs on that sheet are bare numbers with no
prefix at all). carrier_id.classify is used only when SHIPPING LINE is blank
or unrecognised. A known-but-unsupported line (HAPAG) never reaches a fetch
at all -- it's recorded as unsupported_carrier in errors.json.

Maersk's direct API sits behind Akamai Bot Manager and always 403s without a
browser-generated telemetry header -- see maersk_track.py. DO NOT reintroduce
maersk_track.fetch as a first attempt: every Maersk-hinted number goes
straight to one shared browser session (maersk_browser.py), batched together
rather than one launch per number. MSC still goes over plain HTTP, since
that endpoint isn't blocked.

Output
------
containers.json / events.json: one row per container / per event. The row shape
lives in pipeline.py, which every entry point imports -- see the note there
about why it is not duplicated. shipments.json (sheet mode only): one row per COMMERCIAL sheet line,
joined to what we tracked on bl_number. Duplicate BLs across sheet rows are
looked up once and joined back to every matching row.
"""

from __future__ import annotations

import argparse
import asyncio
import random
import re
import sys
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path

import carrier_id
import maersk_track
import msc_track
import read_sheet
from pipeline import CARRIER_NAMES, flatten, write_json

try:
    import maersk_browser
except ImportError:
    maersk_browser = None

CONTAINER_RE = re.compile(r"^[A-Z]{4}\d{7}$")
BL_RE = re.compile(r"^[A-Z]{3,4}[A-Z0-9]{6,13}$")
# Maersk BLs are often bare digits, e.g. 260029935
BL_NUMERIC_RE = re.compile(r"^\d{9,12}$")

# carrier code -> (fetch, parse, extra kwargs)
ADAPTERS = {
    "MSCU": (msc_track.fetch, msc_track.parse, {"mode": msc_track.MODE_CONTAINER_OR_BL}),
    "MAEU": (maersk_track.fetch, maersk_track.parse, {"operator": "MAEU"}),
}

# Known SHIPPING LINE values we recognise but have no adapter for. These are
# real shipments, not failed lookups -- flagged distinctly in errors.json so
# they don't get confused with "wrong carrier" or "number unknown".
UNSUPPORTED_SHIPPING_LINES = {"HAPAG": "Hapag-Lloyd"}


# --------------------------------------------------------------------------- #
# input -- legacy container-list mode
# --------------------------------------------------------------------------- #

def load_numbers(args) -> list[str]:
    raw: list[str] = []

    for item in args.inputs:
        path = Path(item)
        if path.exists() and path.suffix.lower() in (".xlsx", ".xlsm"):
            raw += carrier_id.read_excel(path, args.column, args.sheet)
        elif path.exists() and path.suffix.lower() in (".txt", ".csv"):
            for line in path.read_text(encoding="utf-8-sig").splitlines():
                token = line.split(",")[0].strip().upper().replace(" ", "")
                # Must contain a digit -- kills header rows like
                # "CONTAINERNUMBERS", which otherwise match the loose BL regex.
                if (token and any(ch.isdigit() for ch in token)
                        and (CONTAINER_RE.match(token) or BL_RE.match(token)
                             or BL_NUMERIC_RE.match(token))):
                    raw.append(token)
        else:
            raw.append(item.strip().upper().replace(" ", ""))

    seen, out = set(), []
    for n in raw:
        if n and n not in seen:
            seen.add(n)
            out.append(n)
    return out


def load_sheet(args) -> list[dict] | None:
    """Try BL-sheet mode on a single .xlsx input.

    Returns the shipment list if a BL NO. header was found (or --sheet-mode
    forced it), None if this isn't that kind of file at all -- the caller
    falls back to the legacy container-list path. If --sheet-mode was passed
    explicitly and the header still isn't there, that's a real error and we
    let read_sheet.SheetError propagate instead of silently falling back.
    """
    if len(args.inputs) != 1:
        return None
    path = Path(args.inputs[0])
    if not (path.exists() and path.suffix.lower() in (".xlsx", ".xlsm")):
        return None

    if args.sheet_mode:
        result = read_sheet.get_shipments(path, args.sheet)
    else:
        try:
            result = read_sheet.get_shipments(path, args.sheet)
        except read_sheet.SheetError:
            return None

    print(f"# sheet-mode: {len(result['shipments'])} shipment row(s)  "
          f"(skipped {result['skipped_bulk']} bulk cargo, "
          f"{result['skipped_no_bl']} no BL / no block, "
          f"{result['skipped_blank']} blank)", file=sys.stderr)
    return result["shipments"]


# --------------------------------------------------------------------------- #
# routing
# --------------------------------------------------------------------------- #

def try_carrier(carrier: str, number: str, verbose: bool):
    fetch, parse, kwargs = ADAPTERS[carrier]
    data = fetch(number, verbose=verbose, **kwargs)
    return parse(data, query=number)


def track_one(number: str, hint: str | None, order: list[str], verbose: bool):
    """
    -> (result, carrier, attempts, reasons)

    `order` is the fallback sequence, mutated by the caller so carriers that
    have already worked get tried first next time.

    `reasons` maps carrier -> why it failed. Always collected, because "NOT
    FOUND" alone can't distinguish "wrong carrier" from "your key expired",
    and those need completely different fixes.
    """
    attempts: list[str] = []
    reasons: dict[str, str] = {}

    sequence = [hint] if hint in ADAPTERS else []
    sequence += [c for c in order if c not in sequence]

    for carrier in sequence:
        attempts.append(carrier)
        try:
            result = try_carrier(carrier, number, verbose)
            if result.get("containers"):
                return result, carrier, attempts, reasons
            reasons[carrier] = "empty response"
        except Exception as exc:
            reasons[carrier] = f"{type(exc).__name__}: {str(exc)[:120]}"
            if verbose:
                print(f"#     {CARRIER_NAMES.get(carrier, carrier)}: "
                      f"{reasons[carrier]}", file=sys.stderr)
            continue
    return None, None, attempts, reasons


# --------------------------------------------------------------------------- #
# shared tracking core -- used by both legacy and sheet mode
# --------------------------------------------------------------------------- #

def track_numbers(pairs: list[tuple[str, str | None]], labels: dict[str, str],
                   args, run_id: str, snapshot: str):
    """Track a deduplicated list of (number, hint) pairs.

    -> (all_c, all_e, errors, tracked)
    `tracked` maps number -> the parsed result dict (msc_track.parse /
    maersk_track.parse shape), for callers that need to join back onto
    something else (shipments.json does; the legacy path doesn't).

    Maersk's direct API always 403s (see module docstring), so a number
    hinted MAEU never touches it -- it's queued and looked up in one shared
    browser session at the end, alongside any non-Maersk-hinted number that
    missed on MSC and might still be a Maersk box under an unrecognised
    prefix/shipping line.
    """
    browser_active = maersk_browser is not None and not args.no_browser_fallback
    order = ["MSCU"] if browser_active else list(ADAPTERS)
    hits: dict[str, int] = {c: 0 for c in ADAPTERS}

    all_c, all_e, errors = [], [], []
    tracked: dict[str, dict] = {}
    # Held back from `errors` until the browser batch below either recovers
    # them or genuinely fails, so a self-healing hiccup never gets logged as
    # if it were a real "not found".
    pending_maersk: list[dict] = []

    def record_not_found(number: str, attempts: list[str], reasons: dict[str, str]) -> None:
        print("        NOT FOUND", file=sys.stderr)
        for tried in attempts:
            print(f"          {CARRIER_NAMES.get(tried, tried)}: "
                  f"{reasons.get(tried, 'unknown')}", file=sys.stderr)
        errors.append({
            "run_id": run_id, "snapshot_at": snapshot,
            "query": number, "classified_as": labels.get(number),
            "confidence": "high" if number in labels and labels[number] in CARRIER_NAMES.values() else "unknown",
            "carriers_tried": ", ".join(CARRIER_NAMES.get(a, a) for a in attempts),
            "error": " | ".join(f"{CARRIER_NAMES.get(k, k)}: {v}"
                                for k, v in reasons.items()),
        })

    for i, (number, hint) in enumerate(pairs, 1):
        label = labels.get(number) or (CARRIER_NAMES.get(hint, hint) if hint else "?")
        print(f"[{i}/{len(pairs)}] {number:<20} -> {label}", file=sys.stderr)

        if hint == "MAEU" and browser_active:
            print("        Maersk -- direct API is blocked, tracking via browser",
                  file=sys.stderr)
            pending_maersk.append({"number": number, "attempts": [], "reasons": {}})
            if i < len(pairs):
                time.sleep(args.delay + random.uniform(0, args.delay * 0.4))
            continue

        result, carrier, attempts, reasons = track_one(number, hint, order, args.verbose)

        if result:
            hits[carrier] = hits.get(carrier, 0) + 1
            order.sort(key=lambda c: -hits.get(c, 0))
            tracked[number] = result
            crows, erows = flatten(result, run_id, snapshot)
            all_c.extend(crows)
            all_e.extend(erows)
            for c in crows:
                found = "" if hint == carrier else f"  [found on {CARRIER_NAMES[carrier]}]"
                eta = f"  ETA {c['pod_eta'][:10]}" if c["pod_eta"] else ""
                print(f"        {c['container_id']:<13} {c['current_event']} @ {c['current_city']}"
                      f"  ({c['event_count']} events){eta}{found}", file=sys.stderr)
        elif browser_active:
            print("        not on MSC -- trying Maersk via browser", file=sys.stderr)
            pending_maersk.append({"number": number, "attempts": attempts, "reasons": reasons})
        else:
            record_not_found(number, attempts, reasons)

        if i < len(pairs):
            time.sleep(args.delay + random.uniform(0, args.delay * 0.4))

    if pending_maersk:
        pending_numbers = [p["number"] for p in pending_maersk]
        print(f"\n# Tracking {len(pending_maersk)} Maersk number(s) via browser\n",
              file=sys.stderr)
        recovered = asyncio.run(
            maersk_browser.track_many(pending_numbers, verbose=args.verbose,
                                      timeout_ms=int(args.browser_timeout * 1000)))

        for p in pending_maersk:
            number = p["number"]
            payload = recovered.get(number)
            result = None
            if payload:
                try:
                    result = maersk_track.parse(payload, query=number)
                except Exception as exc:
                    print(f"#   {number}: parse failed after browser fetch: "
                          f"{type(exc).__name__}: {exc}", file=sys.stderr)

            if result:
                hits["MAEU"] = hits.get("MAEU", 0) + 1
                tracked[number] = result
                crows, erows = flatten(result, run_id, snapshot)
                all_c.extend(crows)
                all_e.extend(erows)
                for c in crows:
                    eta = f"  ETA {c['pod_eta'][:10]}" if c["pod_eta"] else ""
                    print(f"        [browser] {number}  {c['container_id']:<13} "
                          f"{c['current_event']} @ {c['current_city']}"
                          f"  ({c['event_count']} events){eta}",
                          file=sys.stderr)
            else:
                attempts = list(p["attempts"]) + ["MAEU"]
                reasons = dict(p["reasons"])
                reasons["MAEU"] = ("no tracking JSON captured via browser -- "
                                   "number unknown, page changed, or a challenge blocked it")
                print(f"\n[{number}] still not found", file=sys.stderr)
                record_not_found(number, attempts, reasons)

    return all_c, all_e, errors, tracked


def build_shipments(shipments: list[dict], tracked: dict[str, dict],
                    run_id: str, snapshot: str) -> list[dict]:
    """Join sheet rows onto tracked results, one output row per sheet row.

    A shipment whose BL wasn't tracked (unsupported carrier, or genuinely
    not found) still gets a row -- with every tracked_* field blank -- since
    it's still real commercial data the sheet owner needs to see.

    run_id / snapshot_at are stamped on every row for the same reason
    containers.json carries them: with --append this file accumulates one copy
    of each sheet row PER RUN, and without a run stamp there is no way to tell
    which copy is current. Readers previously had to assume append order was
    chronological to work that out.
    """
    out = []
    for s in shipments:
        result = tracked.get(s["bl_number"])
        tracked_eta = tracked_status = tracked_location = None
        eta_variance_days = None
        container_count = 0
        containers_str = None

        if result:
            containers = result.get("containers") or []
            container_count = len(containers)
            containers_str = ", ".join(
                c["container_id"] for c in containers if c.get("container_id")) or None
            tracked_eta = result.get("final_pod_eta")
            tracked_status = ("DELIVERED" if containers and all(c.get("delivered") for c in containers)
                              else "IN TRANSIT")

            def _latest_date(c):
                dates = [e["date"] for e in (c.get("events") or []) if e.get("date")]
                return max(dates) if dates else ""

            latest = max(containers, key=_latest_date, default=None)
            if latest:
                tracked_location = latest.get("latest_move")

            sheet_eta = s.get("sheet_eta_date")
            if tracked_eta and sheet_eta:
                try:
                    carrier_date = datetime.fromisoformat(tracked_eta[:10])
                    sheet_date = datetime.fromisoformat(sheet_eta)
                    eta_variance_days = (carrier_date - sheet_date).days
                except ValueError:
                    eta_variance_days = None

        out.append({
            "run_id": run_id,
            "snapshot_at": snapshot,
            "row_number": s["row_number"],
            "bl_number": s["bl_number"],
            "carrier_name": CARRIER_NAMES.get(s["carrier"]) or s.get("shipping_line_raw"),
            "beneficiary": s["beneficiary"],
            "broker": s["broker"],
            "commodity": s["commodity"],
            "pi_no": s["pi_no"],
            "qty": s["qty"],
            "amount": s["amount"],
            "delivery_terms": s["delivery_terms"],
            "free_days": s["free_days"],
            "sheet_eta_date": s["sheet_eta_date"],
            "sheet_eta_place": s["sheet_eta_place"],
            "tracked_eta": tracked_eta,
            "tracked_status": tracked_status,
            "tracked_location": tracked_location,
            "eta_variance_days": eta_variance_days,
            "container_count": container_count,
            "containers": containers_str,
        })
    return out


# --------------------------------------------------------------------------- #
# main -- legacy container-list mode
# --------------------------------------------------------------------------- #

def run_legacy_mode(args, numbers: list[str]) -> int:
    if args.limit:
        numbers = numbers[:args.limit]
    print(f"\n{len(numbers)} numbers to track\n", file=sys.stderr)

    run_id = uuid.uuid4().hex[:12]
    snapshot = datetime.now(timezone.utc).isoformat()

    pairs, labels = [], {}
    for number in numbers:
        info = carrier_id.classify(number)
        hint = info["carrier"] if info["confidence"] == "high" else None
        labels[number] = (CARRIER_NAMES.get(hint, hint) if hint
                          else f"? ({info['note'].split(' -- ')[0]})")
        pairs.append((number, hint))

    all_c, all_e, errors, _tracked = track_numbers(pairs, labels, args, run_id, snapshot)

    print(file=sys.stderr)
    write_json(args.outdir / "containers.json", all_c, args.append)
    write_json(args.outdir / "events.json", all_e, args.append)
    if errors:
        write_json(args.outdir / "errors.json", errors, args.append)

    by_carrier: dict[str, set] = {}
    for c in all_c:
        by_carrier.setdefault(c["carrier_name"], set()).add(c["container_id"])
    forecastable = sum(1 for c in all_c if c["has_forecast"])

    print(f"\n{'-' * 52}", file=sys.stderr)
    for name, ids in sorted(by_carrier.items()):
        print(f"  {name:<12} {len(ids)} containers", file=sys.stderr)
    if errors:
        print(f"  {'not found':<12} {len(errors)}", file=sys.stderr)
    print(f"  {'with ETA':<12} {forecastable} (forward schedule available)",
          file=sys.stderr)
    print(f"  run_id       {run_id}", file=sys.stderr)

    return 1 if errors and not all_c else 0


# --------------------------------------------------------------------------- #
# main -- sheet mode (BL NO.)
# --------------------------------------------------------------------------- #

def run_sheet_mode(args, shipments: list[dict]) -> int:
    if args.limit:
        shipments = shipments[:args.limit]
    print(f"\n{len(shipments)} shipment row(s) from the sheet\n", file=sys.stderr)

    run_id = uuid.uuid4().hex[:12]
    snapshot = datetime.now(timezone.utc).isoformat()
    errors: list[dict] = []

    # One entry per UNIQUE BL -- split shipments (same BL on multiple sheet
    # rows) get looked up once, then build_shipments() joins the result back
    # onto every matching row.
    resolved: dict[str, dict] = {}
    for s in shipments:
        bl = s["bl_number"]
        if bl in resolved:
            continue
        shipping_raw = (s["shipping_line_raw"] or "").strip().upper()
        if s["carrier"]:
            resolved[bl] = {"hint": s["carrier"],
                            "label": CARRIER_NAMES.get(s["carrier"], s["carrier"])}
        elif shipping_raw in UNSUPPORTED_SHIPPING_LINES:
            resolved[bl] = {"hint": None, "unsupported": shipping_raw}
        else:
            # SHIPPING LINE blank or unrecognised -- fall back to BIC/BL
            # prefix classification, same as legacy mode.
            info = carrier_id.classify(bl)
            hint = info["carrier"] if info["confidence"] == "high" else None
            if hint is None:
                # Neither SHIPPING LINE nor the BL's own shape identifies an
                # ocean carrier at all. On this sheet that's not a carrier we
                # failed to recognise -- it's inland/truck cargo (seen:
                # Bangladesh/India deliveries) that never gets a BL because
                # there's no ocean leg to issue one for. Don't burn a live
                # fetch attempt (and a 20s browser timeout) confirming that.
                resolved[bl] = {"hint": None, "no_carrier": s["shipping_line_raw"] or ""}
            else:
                resolved[bl] = {"hint": hint, "label": CARRIER_NAMES.get(hint, hint)}

    unsupported = {bl: v for bl, v in resolved.items() if "unsupported" in v}
    for bl, v in unsupported.items():
        name = UNSUPPORTED_SHIPPING_LINES[v["unsupported"]]
        print(f"  {bl:<20} -> unsupported carrier: {name}", file=sys.stderr)
        errors.append({
            "run_id": run_id, "snapshot_at": snapshot, "query": bl,
            "classified_as": name, "confidence": "high", "carriers_tried": "",
            "error": f"unsupported_carrier: {v['unsupported']}",
        })

    no_carrier = {bl: v for bl, v in resolved.items() if "no_carrier" in v}
    for bl, v in no_carrier.items():
        shown = v["no_carrier"] or "(blank)"
        print(f"  {bl:<20} -> no ocean carrier (inland/truck movement, no BL): "
              f"{shown}", file=sys.stderr)
        errors.append({
            "run_id": run_id, "snapshot_at": snapshot, "query": bl,
            "classified_as": "Inland/Truck", "confidence": "high", "carriers_tried": "",
            "error": f"no_ocean_carrier: SHIPPING LINE={shown!r} -- "
                     "likely inland/truck movement, no BL issued",
        })

    skip = set(unsupported) | set(no_carrier)
    pairs = [(bl, v["hint"]) for bl, v in resolved.items() if bl not in skip]
    labels = {bl: v["label"] for bl, v in resolved.items() if bl not in skip}
    print(f"\n# {len(pairs)} unique BL(s) to look up "
          f"({len(unsupported)} unsupported carrier, "
          f"{len(no_carrier)} no ocean carrier -- skipped)\n", file=sys.stderr)

    all_c, all_e, track_errors, tracked = track_numbers(pairs, labels, args, run_id, snapshot)
    errors.extend(track_errors)

    shipment_rows = build_shipments(shipments, tracked, run_id, snapshot)

    print(file=sys.stderr)
    write_json(args.outdir / "containers.json", all_c, args.append)
    write_json(args.outdir / "events.json", all_e, args.append)
    write_json(args.outdir / "shipments.json", shipment_rows, args.append)
    if errors:
        write_json(args.outdir / "errors.json", errors, args.append)

    tracked_bls = {s["bl_number"] for s in shipments if s["bl_number"] in tracked}
    print(f"\n{'-' * 52}", file=sys.stderr)
    print(f"  shipment rows   {len(shipment_rows)}", file=sys.stderr)
    print(f"  unique BLs      {len(pairs)}", file=sys.stderr)
    print(f"  tracked         {len(tracked_bls)}", file=sys.stderr)
    print(f"  unsupported     {len(unsupported)}", file=sys.stderr)
    print(f"  no ocean carrier {len(no_carrier)}", file=sys.stderr)
    print(f"  not found       {len(track_errors)}", file=sys.stderr)
    print(f"  run_id          {run_id}", file=sys.stderr)

    return 0


# --------------------------------------------------------------------------- #
# entry point
# --------------------------------------------------------------------------- #

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Track containers/BLs across MSC and Maersk; emit flat JSON.")
    ap.add_argument("inputs", nargs="+",
                    help=".xlsx / .txt file, or numbers typed directly")
    ap.add_argument("--column", help="Excel column header (else auto-detect)")
    ap.add_argument("--sheet")
    ap.add_argument("--sheet-mode", action="store_true",
                    help="force BL-based commercial-sheet parsing "
                         "(auto-detected for a single .xlsx otherwise)")
    ap.add_argument("--outdir", type=Path, default=Path("."))
    ap.add_argument("--delay", type=float, default=4.0,
                    help="seconds between lookups (default 4)")
    ap.add_argument("--append", action="store_true", help="build history")
    ap.add_argument("--limit", type=int)
    ap.add_argument("-v", "--verbose", action="store_true",
                    help="show each adapter attempt")
    ap.add_argument("--no-browser-fallback", action="store_true",
                    help="skip the real-browser retry when Maersk's direct API 403s")
    ap.add_argument("--browser-timeout", type=float, default=20.0,
                    help="seconds to wait per number in the browser fallback "
                         "before giving up (default 20; lower moves through "
                         "genuinely-bad numbers faster, at the risk of giving "
                         "up on a slow-loading real one)")
    args = ap.parse_args()

    args.outdir.mkdir(parents=True, exist_ok=True)

    shipments = load_sheet(args)
    if shipments is not None:
        if not shipments:
            raise SystemExit("Sheet parsed but no usable shipment rows found.")
        return run_sheet_mode(args, shipments)

    numbers = load_numbers(args)
    if not numbers:
        raise SystemExit("No usable numbers found.")
    return run_legacy_mode(args, numbers)


if __name__ == "__main__":
    raise SystemExit(main())
