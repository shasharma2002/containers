#!/usr/bin/env python3
"""
THE output schema. One definition, imported by every entry point.

Both carrier adapters produce the same parsed shape (see msc_track.parse /
maersk_track.parse); this module is the single place that turns that shape into
the flat containers.json / events.json rows, and the single place that writes
them to disk with --append semantics.

track_all.py and maersk_browser.py both import flatten() and write_json() from
here. Do NOT copy either function into a caller.

WHY THAT RULE EXISTS
--------------------
It was broken once, and the cost was real. track_all.py grew its own copy of
flatten() when `tracking_number` was added, while maersk_browser.py kept
importing the older one from here. The result was two schemas in one
containers.json: rows written by track_all.py had a tracking_number column and
rows written by maersk_browser.py did not. Anything reading the file -- Power
BI, the Excel report -- then had to detect and repair the difference.

One definition. If a column is added, it is added here, and every entry point
gets it at once.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

# Carrier code -> display name, for the carriers this pipeline can FETCH.
# The reporting layer keeps a wider registry (reporting/normalize.py) covering
# carriers it can report on but the tracker cannot yet retrieve; this one is
# deliberately limited to what has an adapter, so a name appearing here means
# a lookup is actually possible.
CARRIER_NAMES = {"MSCU": "MSC", "MAEU": "Maersk"}


def flatten(result: dict, run_id: str, snapshot: str) -> tuple[list[dict], list[dict]]:
    """Parsed carrier result -> (container rows, event rows).

    Maersk's API never returns a BL number at all (see maersk_track.py), so
    bill_of_lading is always None on a Maersk row -- even when `query` WAS the
    real BL. tracking_number is the one column safe to group or join on
    regardless of carrier: the confirmed BL when the carrier gave one,
    otherwise whatever was searched with.
    """
    tracking_number = result["bill_of_lading"] or result["query"]

    shared = {
        "run_id": run_id,
        "snapshot_at": snapshot,
        "carrier": result["carrier"],
        "carrier_name": CARRIER_NAMES.get(result["carrier"], result["carrier"]),
        "query": result["query"],
        "bill_of_lading": result["bill_of_lading"],
        "tracking_number": tracking_number,
    }

    crows, erows = [], []
    for c in result["containers"]:
        cid = c["container_id"]
        events = c["events"]
        actuals = [e for e in events if e["status"] in ("COMPLETED", "CURRENT")]
        planned = [e for e in events if e["status"] == "PLANNED"]
        current = actuals[-1] if actuals else (events[-1] if events else {})

        crows.append({
            **shared,
            "container_id": cid,
            "container_type": c["container_type"],
            "delivered": c["delivered"],
            "latest_move": c["latest_move"],
            "pod_eta": c["pod_eta"],
            "final_pod_eta": result["final_pod_eta"],
            "origin_city": result["origin_city"],
            "origin_country": result["origin_country"],
            "destination_city": result["destination_city"],
            "destination_country": result["destination_country"],
            "port_of_load": result["port_of_load"],
            "port_of_discharge": result["port_of_discharge"],
            "transshipment_count": len(result["transshipments"]),
            "transshipments": " -> ".join(result["transshipments"]) or None,
            "event_count": len(events),
            "actual_event_count": len(actuals),
            # Only Maersk returns a forward schedule; this flags which rows can
            # support planned-vs-actual analysis at all.
            "has_forecast": len(planned) > 0,
            "current_event": current.get("description"),
            "current_date": current.get("date"),
            "current_city": current.get("city"),
            "current_country": current.get("country"),
            "current_unlocode": current.get("unlocode"),
            "current_vessel": current.get("vessel_name"),
            "current_voyage": current.get("voyage"),
        })

        for e in events:
            erows.append({
                "run_id": run_id,
                "snapshot_at": snapshot,
                "carrier": result["carrier"],
                "container_id": cid,
                "bill_of_lading": result["bill_of_lading"],
                "tracking_number": tracking_number,
                "event_key": f"{cid}|{e['sequence']}|{e['raw_date']}|{e['description']}",
                "sequence": e["sequence"],
                "status": e["status"],
                "event_date": e["date"],
                "description": e["description"],
                "city": e["city"],
                "country": e["country"],
                "unlocode": e["unlocode"],
                "terminal": e["terminal"],
                "terminal_code": e["terminal_code"],
                "mode": e["mode"],
                "vessel_name": e["vessel_name"],
                "voyage": e["voyage"],
                "vessel_imo": e["vessel_imo"],
                "vessel_flag": e["vessel_flag"],
                "laden_status": e["laden_status"],
            })

    return crows, erows


def write_json(path: Path, rows: list[dict], append: bool) -> None:
    """Write rows as a JSON array, optionally appending to what is there.

    --append is what makes delay detection possible: the reporting layer
    compares a container's ETA between snapshots to find out whether the
    carrier moved its own date. A single run has nothing to compare against.
    """
    existing: list[dict] = []
    if append and path.exists():
        try:
            loaded = json.loads(path.read_text(encoding="utf-8"))
            existing = loaded if isinstance(loaded, list) else []
        except (json.JSONDecodeError, OSError):
            print(f"# warn: {path} unreadable, starting fresh", file=sys.stderr)
    path.write_text(json.dumps(existing + rows, indent=2, ensure_ascii=False),
                    encoding="utf-8")
    print(f"# {path.name:<18} {len(existing) + len(rows)} rows"
          + (f"  (+{len(existing)} existing)" if existing else ""), file=sys.stderr)
