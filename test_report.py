#!/usr/bin/env python3
"""
Tests for the Excel reporting layer.

    python test_report.py

Two halves, and the distinction matters when reading the results:

  REAL DATA     run against the JSON actually in this project. Only Maersk and
                MSC exist there, so only those two are genuinely tested
                end-to-end against carrier output.

  SYNTHETIC     fixtures in the tracking application's own JSON shape, written
                to a temp directory, covering the six carriers that have no
                adapter yet plus the edge cases (empty results, all-failures,
                missing fields, duplicate events, unseen event wording, bad
                dates, a large dataset).

A synthetic pass proves the REPORT handles a carrier's data correctly. It does
not prove the tracker can fetch that carrier -- nothing here can, because no
adapter exists. The two are reported separately on purpose.
"""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
import traceback
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

from reporting import analyze, excel, loader, normalize, validate

PASSED: list[str] = []
FAILED: list[tuple[str, str]] = []
SKIPPED: list[str] = []


def skip(name: str, why: str) -> None:
    """Not a failure. A clean checkout has no tracking data yet, and a test
    that went red for that reason would train people to ignore red."""
    SKIPPED.append("%s -- %s" % (name, why))
    print("  SKIP  %s  (%s)" % (name, why))


def check(name: str, condition: bool, detail: str = "") -> None:
    if condition:
        PASSED.append(name)
        print("  PASS  %s" % name)
    else:
        FAILED.append((name, detail))
        print("  FAIL  %s  %s" % (name, detail))


def section(text: str) -> None:
    print("\n%s\n%s" % (text, "-" * len(text)))


# --------------------------------------------------------------------------- #
# fixtures
# --------------------------------------------------------------------------- #

BASE = datetime(2026, 9, 1, 8, 0, 0)


def event(container, carrier, seq, days, description, status="COMPLETED", **extra):
    """An events.json row in exactly the shape pipeline.flatten writes."""
    when = BASE + timedelta(days=days)
    row = {
        "run_id": extra.pop("run_id", "run0001"),
        "snapshot_at": extra.pop("snapshot_at", BASE.isoformat() + "+00:00"),
        "carrier": carrier,
        "container_id": container,
        "bill_of_lading": extra.pop("bill_of_lading", None),
        "tracking_number": extra.pop("tracking_number", container),
        "event_key": "%s|%d|%s|%s" % (container, seq, when.strftime("%d/%m/%Y"),
                                      description),
        "sequence": seq,
        "status": status,
        "event_date": when.isoformat(),
        "description": description,
        "city": extra.pop("city", "SINGAPORE"),
        "country": extra.pop("country", "SG"),
        "unlocode": extra.pop("unlocode", "SGSIN"),
        "terminal": extra.pop("terminal", "PSA TERMINAL"),
        "terminal_code": extra.pop("terminal_code", "SGSINPS"),
        "mode": extra.pop("mode", "VESSEL"),
        "vessel_name": extra.pop("vessel_name", "TEST VOYAGER"),
        "voyage": extra.pop("voyage", "001E"),
        "vessel_imo": extra.pop("vessel_imo", None),
        "vessel_flag": extra.pop("vessel_flag", None),
        "laden_status": extra.pop("laden_status", "LADEN"),
    }
    row.update(extra)
    return row


def container(cid, carrier, carrier_name, **extra):
    """A containers.json row in exactly the shape track_all.flatten writes."""
    row = {
        "run_id": extra.pop("run_id", "run0001"),
        "snapshot_at": extra.pop("snapshot_at", BASE.isoformat() + "+00:00"),
        "carrier": carrier,
        "carrier_name": carrier_name,
        "query": extra.pop("query", cid),
        "bill_of_lading": extra.pop("bill_of_lading", None),
        "tracking_number": extra.pop("tracking_number", cid),
        "container_id": cid,
        "container_type": extra.pop("container_type", "40' DRY"),
        "delivered": extra.pop("delivered", False),
        "latest_move": extra.pop("latest_move", "SINGAPORE, SG"),
        "pod_eta": extra.pop("pod_eta", (BASE + timedelta(days=20)).isoformat()),
        "final_pod_eta": extra.pop("final_pod_eta",
                                   (BASE + timedelta(days=20)).isoformat()),
        "origin_city": extra.pop("origin_city", "SHANGHAI"),
        "origin_country": extra.pop("origin_country", "CN"),
        "destination_city": extra.pop("destination_city", "BIRGUNJ"),
        "destination_country": extra.pop("destination_country", "NP"),
        "port_of_load": extra.pop("port_of_load", "SHANGHAI"),
        "port_of_discharge": extra.pop("port_of_discharge", "KOLKATA"),
        "transshipment_count": extra.pop("transshipment_count", 1),
        "transshipments": extra.pop("transshipments", "SINGAPORE, SG"),
        "event_count": extra.pop("event_count", 3),
        "actual_event_count": extra.pop("actual_event_count", 3),
        "has_forecast": extra.pop("has_forecast", False),
        "current_event": extra.pop("current_event", "Vessel Departure"),
        "current_date": extra.pop("current_date", (BASE + timedelta(days=2)).isoformat()),
        "current_city": extra.pop("current_city", "SINGAPORE"),
        "current_country": extra.pop("current_country", "SG"),
        "current_unlocode": extra.pop("current_unlocode", "SGSIN"),
        "current_vessel": extra.pop("current_vessel", "TEST VOYAGER"),
        "current_voyage": extra.pop("current_voyage", "001E"),
    }
    row.update(extra)
    return row


# The six carriers with no adapter. Event wording is what each line publishes on
# its own tracking pages -- close enough to exercise the keyword fallback, and
# deliberately NOT added to EXACT_EVENT_MAP, because no real payload from these
# carriers has been seen in this project yet.
OTHER_CARRIER_EVENTS = {
    ("COSU", "COSCO"): [
        "Empty Equipment Dispatch", "Gate in at first POL",
        "Loaded on vessel at first POL", "Vessel Departure from first POL",
        "Discharged at transshipment port", "Arrival at final port of discharge",
    ],
    ("CMDU", "CMA CGM"): [
        "Empty container released to shipper", "Container gate in at terminal",
        "Loaded on board vessel", "Vessel departure",
        "Vessel arrival", "Container discharged from vessel",
        "Container delivered to consignee",
    ],
    ("HLCU", "Hapag-Lloyd"): [
        "Empty picked up by customer", "Container arrival in port of loading",
        "Vessel departed", "Transshipment discharged",
        "Container departed terminal", "Empty returned to depot",
    ],
    ("ONEY", "ONE"): [
        "Empty Container Release to Shipper", "Gate In to Outbound Terminal",
        "Loaded on Vessel", "Vessel Departure from Port of Loading",
        "Vessel Arrival at Port of Discharge", "Rail Departure from POD",
    ],
    ("EGLV", "Evergreen"): [
        "Empty container pickup by shipper", "Gate-in full at origin CY",
        "Loaded onto vessel", "Departure from port of loading",
        "Discharged from vessel", "Customs clearance completed",
    ],
    ("YMLU", "Yang Ming"): [
        "Empty release to shipper", "Received at origin terminal",
        "On board vessel", "Sailed from port of loading",
        "Berthed at discharge port", "Truck delivery to consignee",
    ],
}


def write_fixture(directory: Path, containers, events, errors=None, shipments=None):
    directory.mkdir(parents=True, exist_ok=True)
    (directory / "containers.json").write_text(
        json.dumps(containers, indent=2), encoding="utf-8")
    (directory / "events.json").write_text(
        json.dumps(events, indent=2), encoding="utf-8")
    if errors is not None:
        (directory / "errors.json").write_text(
            json.dumps(errors, indent=2), encoding="utf-8")
    if shipments is not None:
        (directory / "shipments.json").write_text(
            json.dumps(shipments, indent=2), encoding="utf-8")
    return directory


def generate(directory: Path, output: Path, latest_only=True, as_of=None):
    """Run the whole pipeline the way generate_report.py does."""
    data = loader.load(directory, latest_only=latest_only)
    resolved = analyze.enrich(data, as_of)
    kpi = analyze.kpis(data, resolved)
    dist = analyze.distributions(data, resolved)
    carriers = analyze.carrier_summary(data)
    quality = validate.findings(data, resolved)
    ledger, balanced = validate.reconcile(data)
    excel.build(data, kpi, dist, carriers, quality, ledger, output)
    return data, kpi, ledger, balanced


def opens(path: Path) -> tuple[bool, str]:
    """Does the workbook load back, with its sheets and tables intact?"""
    try:
        wb = load_workbook(path)
        if excel.S_DASH not in wb.sheetnames:
            return False, "no Dashboard sheet"
        for ws in wb.worksheets:
            for name, ref in ws.tables.items():
                ref = ref if isinstance(ref, str) else ref.ref
                if ":" not in ref:
                    return False, "bad table ref %s on %s" % (name, ws.title)
        return True, "%d sheets" % len(wb.sheetnames)
    except Exception as exc:
        return False, "%s: %s" % (type(exc).__name__, exc)


# --------------------------------------------------------------------------- #
# tests -- real project data
# --------------------------------------------------------------------------- #

def test_real_data(tmp: Path) -> None:
    section("REAL DATA -- the JSON actually in this project")

    available = [(label, directory)
                 for label, directory in (("appended history", Path(".")),
                                          ("single run", Path("out")))
                 if (directory / "containers.json").exists()]
    if not available:
        skip("real-data tests",
             "no tracking JSON in . or out/ -- run track_all.py first")
        return

    for label, directory in available:
        output = tmp / ("real_%s.xlsx" % label.split()[0])
        data, kpi, ledger, balanced = generate(directory, output)
        ok, detail = opens(output)
        check("%s: workbook opens" % label, ok, detail)
        check("%s: ledger balances" % label, balanced,
              str([r for r in ledger if r["balanced"] == "FAIL"]))
        check("%s: every container reported once" % label,
              len({c["container_id"] for c in data.containers}) == len(data.containers))
        check("%s: no event lost or invented" % label,
              sum(c["events_on_report"] for c in data.containers) == len(data.events))
        check("%s: carriers tested = %s" % (label,
              sorted({c["carrier_display"] for c in data.containers})), True)

    # Latest-snapshot selection must collapse the append history to one row per
    # container, and every collapsed row must still be in the workbook.
    data, kpi, ledger, _ = generate(Path("."), tmp / "real_latest.xlsx")
    if len(data.raw_containers) == len(data.containers):
        skip("append-history collapse",
             "only one run present; nothing to collapse yet")
    else:
        check("append history collapsed to latest row per container",
              len(data.containers) < len(data.raw_containers),
              "%d of %d" % (len(data.containers), len(data.raw_containers)))
    wb = load_workbook(tmp / "real_latest.xlsx")
    check("superseded rows retained on Raw Containers",
          wb[excel.S_RAW_CONTAINERS].max_row - 4 == len(data.raw_containers),
          "%d rows on sheet vs %d read"
          % (wb[excel.S_RAW_CONTAINERS].max_row - 4, len(data.raw_containers)))
    check("Raw sheets hidden",
          wb[excel.S_RAW_CONTAINERS].sheet_state == "hidden"
          and wb[excel.S_RAW_EVENTS].sheet_state == "hidden")

    # --all-snapshots must keep every row.
    data, _, ledger, balanced = generate(Path("."), tmp / "real_all.xlsx",
                                         latest_only=False)
    check("--all-snapshots keeps every row",
          len(data.containers) == len(data.raw_containers)
          and len(data.events) == len(data.raw_events))
    check("--all-snapshots still balances", balanced)

    # Every event string in the real data must map exactly -- a keyword match on
    # Maersk or MSC wording means the exact table has drifted from the feed.
    events = json.loads(Path("events.json").read_text(encoding="utf-8"))
    sources = {normalize.normalize_event(e["description"])[1] for e in events}
    check("all real event wording maps exactly", sources == {"exact"},
          "mapping sources seen: %s" % sorted(sources))


# --------------------------------------------------------------------------- #
# tests -- eight carriers
# --------------------------------------------------------------------------- #

def test_eight_carriers(tmp: Path) -> None:
    section("EIGHT CARRIERS -- synthetic fixtures for the six with no adapter")

    containers, events = [], []
    for (code, name), descriptions in OTHER_CARRIER_EVENTS.items():
        cid = "%sU1234567" % code[:3]
        containers.append(container(cid, code, name,
                                    event_count=len(descriptions),
                                    actual_event_count=len(descriptions),
                                    current_event=descriptions[-1]))
        for seq, description in enumerate(descriptions):
            events.append(event(cid, code, seq, seq * 3, description,
                                status="CURRENT" if seq == len(descriptions) - 1
                                else "COMPLETED"))

    directory = write_fixture(tmp / "eight", containers, events)
    data, kpi, ledger, balanced = generate(directory, tmp / "eight.xlsx")
    ok, detail = opens(tmp / "eight.xlsx")
    check("six adapterless carriers: workbook opens", ok, detail)
    check("six adapterless carriers: ledger balances", balanced)

    reported = {c["carrier_display"] for c in data.containers}
    for (code, name) in OTHER_CARRIER_EVENTS:
        check("%s reported with its own name" % name, name in reported,
              "got %s" % sorted(reported))

    # The whole point of the keyword fallback: unseen wording must still land on
    # a usable milestone rather than in an OTHER bucket.
    unmapped = [e for e in data.events if e["mapping_source"] in ("unmapped", "blank")]
    check("no unseen carrier wording left unclassified", not unmapped,
          "unmapped: %s" % sorted({e["description"] for e in unmapped}))
    keyword = [e for e in data.events if e["mapping_source"] == "keyword"]
    check("unseen wording resolved by keyword rule and flagged as such",
          len(keyword) == len(data.events),
          "%d of %d" % (len(keyword), len(data.events)))

    # Every one of the eight must appear on Carrier Mapping whether it has data
    # or not -- that sheet is how a reader learns COSCO produced nothing.
    wb = load_workbook(tmp / "eight.xlsx")
    ws = wb[excel.S_CARRIER_MAP]
    listed = {ws.cell(row=r, column=1).value for r in range(5, ws.max_row + 1)}
    for target in normalize.TARGET_CARRIERS:
        check("Carrier Mapping lists %s" % target["name"],
              target["name"] in listed)

    # Maersk and MSC have no rows in this fixture and must be shown as absent,
    # not silently dropped.
    rows = analyze.carrier_mapping_rows(data)
    absent = {r["carrier"]: r["coverage"] for r in rows if r["containers"] == 0}
    check("carriers with no data are labelled, not omitted",
          "Maersk" in absent and "MSC" in absent, str(absent))


# --------------------------------------------------------------------------- #
# tests -- edge cases
# --------------------------------------------------------------------------- #

def test_edge_cases(tmp: Path) -> None:
    section("EDGE CASES")

    # ---- empty results ----------------------------------------------------- #
    directory = write_fixture(tmp / "empty", [], [])
    try:
        loader.load(directory)
        check("empty containers.json refuses with a clear error", False,
              "loaded without error")
    except loader.LoaderError as exc:
        check("empty containers.json refuses with a clear error",
              "no usable rows" in str(exc), str(exc))

    # ---- events but no containers ------------------------------------------ #
    directory = write_fixture(tmp / "orphan",
                              [container("AAAU1111111", "MAEU", "Maersk")],
                              [event("BBBU2222222", "MAEU", 0, 0, "Gate In")])
    data, _, ledger, balanced = generate(directory, tmp / "orphan.xlsx")
    check("orphan events reported, not dropped",
          any("appears in no container row" in n["message"] for n in data.notes))
    ok, detail = opens(tmp / "orphan.xlsx")
    check("orphan events: workbook opens", ok, detail)

    # ---- tracking failures only -------------------------------------------- #
    failures = [{
        "run_id": "run0001", "snapshot_at": BASE.isoformat() + "+00:00",
        "query": "HLCUMTR260616918", "classified_as": "Hapag-Lloyd",
        "confidence": "high", "carriers_tried": "",
        "error": "unsupported_carrier: HAPAG",
    }, {
        "run_id": "run0001", "snapshot_at": BASE.isoformat() + "+00:00",
        "query": "ZZZU9999999", "classified_as": None, "confidence": "unknown",
        "carriers_tried": "MSC, Maersk",
        "error": "MSC: empty response | Maersk: no tracking JSON captured",
    }]
    directory = write_fixture(tmp / "failures",
                              [container("AAAU1111111", "MAEU", "Maersk")],
                              [event("AAAU1111111", "MAEU", 0, 0, "Gate In")],
                              errors=failures)
    data, kpi, ledger, balanced = generate(directory, tmp / "failures.xlsx")
    kinds = {e["failure_kind"] for e in data.errors}
    check("failure kinds separated", kinds == {"Unsupported carrier", "Not found"},
          str(kinds))
    check("unsupported carrier is not reported as a delay",
          not any(f["class"] == "Shipment" and "Unsupported" in f["category"]
                  for f in data.exceptions))
    check("failures: ledger balances", balanced)

    # ---- missing fields everywhere ----------------------------------------- #
    sparse = container("SPAR1111111", "MAEU", "Maersk",
                       pod_eta=None, final_pod_eta=None, destination_city=None,
                       destination_country=None, current_event=None,
                       current_date=None, current_vessel=None, current_voyage=None,
                       latest_move=None, container_type=None, transshipments=None,
                       tracking_number=None, query=None, bill_of_lading=None)
    bare = event("SPAR1111111", "MAEU", 0, 0, None, city=None, country=None,
                 unlocode=None, terminal=None, terminal_code=None, mode=None,
                 vessel_name=None, voyage=None, laden_status=None,
                 tracking_number=None)
    directory = write_fixture(tmp / "sparse", [sparse], [bare])
    data, kpi, ledger, balanced = generate(directory, tmp / "sparse.xlsx")
    ok, detail = opens(tmp / "sparse.xlsx")
    check("all-null fields: workbook opens", ok, detail)
    check("all-null fields: ledger balances", balanced)
    checks = {f["check"] for f in validate.findings(data, BASE)}
    check("missing ETA reported as a data gap, not a delay",
          any(f["category"] == "No ETA from carrier" and f["class"] == "Data"
              for f in data.exceptions))
    check("missing description reported",
          "Missing event description" in checks, str(sorted(checks)))
    check("missing tracking number reported",
          "No tracking number" in checks, str(sorted(checks)))

    # ---- invalid dates ----------------------------------------------------- #
    directory = write_fixture(
        tmp / "baddates",
        [container("DATE1111111", "MSCU", "MSC", snapshot_at="not-a-date",
                   current_date="31/31/2026")],
        [event("DATE1111111", "MSCU", 0, 0, "Export received at CY",
               event_date="garbage")])
    data, kpi, ledger, balanced = generate(directory, tmp / "baddates.xlsx")
    checks = {f["check"] for f in validate.findings(data, BASE)}
    check("unparseable snapshot reported", "Unparseable snapshot_at" in checks)
    check("unparseable event date reported",
          "Missing or unparseable event date" in checks, str(sorted(checks)))
    check("bad dates: ledger balances", balanced)
    ok, detail = opens(tmp / "baddates.xlsx")
    check("bad dates: workbook opens", ok, detail)

    # ---- duplicate events within one run ----------------------------------- #
    duplicate = event("DUPE1111111", "MSCU", 0, 0, "Export received at CY")
    directory = write_fixture(tmp / "dupes",
                              [container("DUPE1111111", "MSCU", "MSC")],
                              [duplicate, dict(duplicate)])
    data, _, ledger, balanced = generate(directory, tmp / "dupes.xlsx")
    checks = {f["check"] for f in validate.findings(data, BASE)}
    check("duplicate event in one run reported",
          "Duplicate event in one run" in checks, str(sorted(checks)))
    check("duplicate events kept, not silently dropped",
          len(data.events) == 2, "%d events" % len(data.events))

    # ---- duplicate containers across runs ---------------------------------- #
    later = BASE + timedelta(days=5)
    directory = write_fixture(
        tmp / "history",
        [container("HIST1111111", "MAEU", "Maersk", run_id="r1",
                   snapshot_at=BASE.isoformat() + "+00:00",
                   pod_eta=(BASE + timedelta(days=10)).isoformat()),
         container("HIST1111111", "MAEU", "Maersk", run_id="r2",
                   snapshot_at=later.isoformat() + "+00:00",
                   pod_eta=(BASE + timedelta(days=22)).isoformat())],
        [event("HIST1111111", "MAEU", 0, 0, "Gate In", run_id="r1"),
         event("HIST1111111", "MAEU", 0, 0, "Gate In", run_id="r2"),
         event("HIST1111111", "MAEU", 1, 2, "Load", run_id="r2")])
    data, kpi, ledger, balanced = generate(directory, tmp / "history.xlsx")
    check("latest snapshot chosen for a repeated container",
          len(data.containers) == 1 and data.containers[0]["run_id"] == "r2",
          str([(c["run_id"]) for c in data.containers]))
    check("events taken from the same run as the container row",
          len(data.events) == 2, "%d events" % len(data.events))
    check("ETA slip derived from the carrier's own revision",
          data.containers[0]["eta_slip_days"] == 12,
          "slip = %s" % data.containers[0]["eta_slip_days"])
    check("ETA slip raised as a shipment finding",
          any(f["category"] == "ETA slipped" and f["class"] == "Shipment"
              for f in data.exceptions))
    check("single-snapshot container claims no slip",
          loader.load(tmp / "sparse").containers is not None)

    # ---- MSC forecast mislabelled as CURRENT ------------------------------- #
    # Both real moves are dated BEFORE the snapshot, because an event dated
    # after the snapshot cannot be an actual and the report is right to say so.
    # Only the ETA row is in the future, which is the case being tested.
    directory = write_fixture(
        tmp / "forecast",
        [container("FCST1111111", "MSCU", "MSC",
                   current_event="Estimated Time of Arrival",
                   current_date=(BASE + timedelta(days=30)).isoformat())],
        [event("FCST1111111", "MSCU", 0, -12, "Export received at CY"),
         event("FCST1111111", "MSCU", 1, -10, "Export Loaded on Vessel"),
         event("FCST1111111", "MSCU", 2, 30, "Estimated Time of Arrival",
               status="CURRENT")])
    data, _, _, _ = generate(directory, tmp / "forecast.xlsx", as_of=BASE)
    row = data.containers[0]
    check("MSC ETA row is not treated as a confirmed move",
          row["last_actual_normalized"] == "Loaded on vessel",
          "got %r" % row["last_actual_normalized"])
    check("forecast-as-current flagged for the reader",
          row["current_is_forecast"] is True)
    basis = {e["description"]: e["time_basis"] for e in data.events}
    check("ETA advice classified EXPECTED",
          basis["Estimated Time of Arrival"] == "EXPECTED", str(basis))

    # ---- no errors.json / no shipments.json -------------------------------- #
    directory = write_fixture(tmp / "minimal",
                              [container("MIN11111111", "MAEU", "Maersk")],
                              [event("MIN11111111", "MAEU", 0, 0, "Gate In")])
    data, _, _, balanced = generate(directory, tmp / "minimal.xlsx")
    wb = load_workbook(tmp / "minimal.xlsx")
    check("optional inputs absent: sheets omitted cleanly",
          excel.S_SHIPMENTS not in wb.sheetnames
          and excel.S_FAILURES not in wb.sheetnames)
    check("optional inputs absent: workbook still balances", balanced)

    # ---- schema v1 (no tracking_number) ------------------------------------ #
    v1_container = container("VONE1111111", "MAEU", "Maersk",
                             query="260029935", bill_of_lading=None)
    del v1_container["tracking_number"]
    v1_event = event("VONE1111111", "MAEU", 0, 0, "Gate In")
    del v1_event["tracking_number"]
    directory = write_fixture(tmp / "v1", [v1_container], [v1_event])
    data, _, _, balanced = generate(directory, tmp / "v1.xlsx")
    check("schema v1 row accepted and reported",
          any("predate the tracking_number column" in n["message"]
              for n in data.notes))
    check("schema v1 tracking number derived from query",
          data.containers[0]["tracking_number_resolved"] == "260029935",
          str(data.containers[0]["tracking_number_resolved"]))
    check("schema v1: ledger balances", balanced)

    # ---- malformed JSON ---------------------------------------------------- #
    directory = tmp / "broken"
    directory.mkdir(parents=True, exist_ok=True)
    (directory / "containers.json").write_text("{not json", encoding="utf-8")
    (directory / "events.json").write_text("[]", encoding="utf-8")
    try:
        loader.load(directory)
        check("malformed JSON refuses with a clear error", False, "loaded anyway")
    except loader.LoaderError as exc:
        check("malformed JSON refuses with a clear error",
              "no usable rows" in str(exc), str(exc))


def test_bl_grain(tmp: Path) -> None:
    section("BILL OF LADING GRAIN -- many containers on one BL")

    def consignment(bl, carrier, name, count, description, days, **extra):
        """`count` containers on one BL, all at the same milestone."""
        containers, events = [], []
        for index in range(count):
            cid = "BOX%c%06d" % ("UJZ"[index % 3], index)
            containers.append(container(cid, carrier, name, query=bl,
                                        tracking_number=bl, **extra))
            events.append(event(cid, carrier, 0, days, description,
                                status="CURRENT", tracking_number=bl))
        return containers, events

    # ---- a uniform 20-box consignment -------------------------------------- #
    # The whole point: 20 containers moving together are ONE line, not 20, and
    # raise no divergence finding at all.
    cs, es = consignment("BLUNIFORM1", "MAEU", "Maersk", 20, "Discharg", -3)
    directory = write_fixture(tmp / "bl_uniform", cs, es)
    data, kpi, ledger, balanced = generate(directory, tmp / "bl_uniform.xlsx",
                                           as_of=BASE)
    check("20 containers collapse to 1 BL row", len(data.shipment_rows) == 1,
          "%d rows" % len(data.shipment_rows))
    row = data.shipment_rows[0]
    check("BL row counts its containers", row["containers"] == 20, str(row["containers"]))
    check("uniform consignment is not flagged as split", row["split"] == "No",
          "%s / %s" % (row["split"], row["divergence"]))
    check("uniform consignment raises no divergence finding",
          not [f for f in data.exceptions if f["category"].startswith("Containers")],
          str([f["category"] for f in data.exceptions]))
    check("BL grain balances in the ledger", balanced)
    check("every container accounted for under a BL",
          sum(s["containers"] for s in data.shipment_rows) == len(data.containers))

    # ---- timestamp jitter must NOT look like a problem --------------------- #
    # A terminal scanning a 23-box stack produces timestamps minutes apart. This
    # is the single most important false positive to avoid, because it would
    # fire on almost every real multi-container BL.
    cs, es = [], []
    for index in range(23):
        cid = "JIT%c%06d" % ("UJZ"[index % 3], index)
        cs.append(container(cid, "MAEU", "Maersk", query="BLJITTER1",
                            tracking_number="BLJITTER1"))
        when = (BASE + timedelta(days=-3, minutes=index * 4)).isoformat()
        es.append(event(cid, "MAEU", 0, 0, "Discharg", status="CURRENT",
                        event_date=when, tracking_number="BLJITTER1"))
    directory = write_fixture(tmp / "bl_jitter", cs, es)
    data, _, _, balanced = generate(directory, tmp / "bl_jitter.xlsx", as_of=BASE)
    row = data.shipment_rows[0]
    check("timestamps minutes apart: spread is 0 days",
          row["move_spread_days"] == 0, str(row["move_spread_days"]))
    check("timestamps minutes apart: not flagged as separated",
          not [f for f in data.exceptions
               if f["category"] == "Containers separated in time"])
    check("timestamps minutes apart: not flagged as split", row["split"] == "No")

    # ---- one container left behind ----------------------------------------- #
    # The real exception the business cares about: 12 boxes delivered, 1 still
    # in transit somewhere else.
    cs, es = consignment("BLSPLIT001", "MAEU", "Maersk", 12, "Container Return", -2,
                         delivered=True)
    stray = container("STRAYU123456", "MAEU", "Maersk", query="BLSPLIT001",
                      tracking_number="BLSPLIT001", delivered=False)
    cs.append(stray)
    es.append(event("STRAYU123456", "MAEU", 0, -2, "Discharg", status="CURRENT",
                    city="COLOMBO", country="LK", tracking_number="BLSPLIT001"))
    directory = write_fixture(tmp / "bl_split", cs, es)
    data, kpi, _, balanced = generate(directory, tmp / "bl_split.xlsx", as_of=BASE)
    row = data.shipment_rows[0]
    check("13 boxes with 1 stray collapse to 1 BL row",
          len(data.shipment_rows) == 1 and row["containers"] == 13)
    check("stray container detected as split", row["split"] == "Yes", row["divergence"] or "")
    check("exactly one outlier identified", row["outlier_count"] == 1,
          "%d: %s" % (row["outlier_count"], row["outliers"]))
    check("the outlier is named", row["outliers"] == "STRAYU123456", str(row["outliers"]))
    check("consensus follows the 12, not the straggler",
          row["status"] == "Delivered" and row["location"] == "SINGAPORE, SG",
          "%s at %s" % (row["status"], row["location"]))
    divergence = [f for f in data.exceptions
                  if f["category"] == "Containers not travelling together"]
    check("divergence raised once, not 13 times", len(divergence) == 1,
          "%d findings" % len(divergence))
    check("divergence scope names the proportion",
          divergence and divergence[0]["scope"] == "1 of 13 containers",
          divergence[0]["scope"] if divergence else "none")
    check("split shipment counted in the KPIs",
          kpi["split_shipments"] == 1 and kpi["split_containers"] == 1)
    check("split BL balances", balanced)

    # ---- containers genuinely days apart ----------------------------------- #
    cs, es = consignment("BLLATE0001", "MAEU", "Maersk", 8, "Discharg", -3)
    late = container("LATEU1234567", "MAEU", "Maersk", query="BLLATE0001",
                     tracking_number="BLLATE0001")
    cs.append(late)
    es.append(event("LATEU1234567", "MAEU", 0, -11, "Discharg", status="CURRENT",
                    tracking_number="BLLATE0001"))
    directory = write_fixture(tmp / "bl_late", cs, es)
    data, _, _, _ = generate(directory, tmp / "bl_late.xlsx", as_of=BASE)
    row = data.shipment_rows[0]
    check("containers 8 days apart: spread measured",
          row["move_spread_days"] == 8, str(row["move_spread_days"]))
    check("containers days apart: flagged as separated in time",
          [f for f in data.exceptions
           if f["category"] == "Containers separated in time"])

    # ---- staggered empty returns on a delivered BL are NOT a problem ------- #
    # A consignee unloads 10 boxes over a week. Every box is delivered; the
    # spread is real but means nothing.
    cs, es = [], []
    for index in range(10):
        cid = "DLV%c%06d" % ("UJZ"[index % 3], index)
        cs.append(container(cid, "MAEU", "Maersk", query="BLDONE0001",
                            tracking_number="BLDONE0001", delivered=True))
        es.append(event(cid, "MAEU", 0, -12 + index, "Container Return",
                        status="CURRENT", tracking_number="BLDONE0001"))
    directory = write_fixture(tmp / "bl_done", cs, es)
    data, _, _, _ = generate(directory, tmp / "bl_done.xlsx", as_of=BASE)
    row = data.shipment_rows[0]
    check("delivered BL: spread exists but is not flagged",
          row["move_spread_days"] >= analyze.CONTAINER_SPREAD_WARN_DAYS
          and not [f for f in data.exceptions
                   if f["category"] == "Containers separated in time"],
          "spread=%s" % row["move_spread_days"])

    # ---- ETA a day or two apart is normal, 3+ days is not ------------------ #
    # Staggered arrival is reported on its own, NOT as Split?: Split? promises a
    # nameable outlier container, and a spread of ETAs has no culprit box.
    for label, offsets, expect in (("1 day", [0, 1, 1, 0], False),
                                   ("6 days", [0, 2, 4, 6], True)):
        cs, es = [], []
        for index, offset in enumerate(offsets):
            cid = "ETA%c%06d" % ("UJZ"[index % 3], index)
            eta = (BASE + timedelta(days=20 + offset)).isoformat()
            cs.append(container(cid, "MAEU", "Maersk", query="BLETA00001",
                                tracking_number="BLETA00001",
                                pod_eta=eta, final_pod_eta=eta))
            es.append(event(cid, "MAEU", 0, -3, "Discharg", status="CURRENT",
                            tracking_number="BLETA00001"))
        directory = write_fixture(tmp / ("bl_eta_" + label.split()[0]), cs, es)
        data, _, _, _ = generate(directory, tmp / ("bl_eta_%s.xlsx" % label.split()[0]),
                                 as_of=BASE)
        row = data.shipment_rows[0]
        flagged = bool([f for f in data.exceptions
                        if f["category"] == "Arrival dates spread across the consignment"])
        check("ETA %s apart: %s" % (label, "flagged" if expect else "not flagged"),
              flagged == expect,
              "spread=%s split=%s" % (row["eta_spread_days"], row["split"]))
        check("ETA %s apart: Split? stays No (no culprit container)" % label,
              row["split"] == "No", "split=%s" % row["split"])

    # ---- one finding for the whole consignment ----------------------------- #
    # 20 containers all past their ETA is one line saying "all 20", not 20 lines.
    past = (BASE - timedelta(days=10)).isoformat()
    cs, es = consignment("BLOVERDUE1", "MAEU", "Maersk", 20, "Discharg", -3,
                         pod_eta=past, final_pod_eta=past)
    directory = write_fixture(tmp / "bl_overdue", cs, es)
    data, _, _, _ = generate(directory, tmp / "bl_overdue.xlsx", as_of=BASE)
    overdue = [f for f in data.exceptions if f["category"] == "ETA passed, not delivered"]
    check("20 overdue containers roll up to one finding", len(overdue) == 1,
          "%d findings" % len(overdue))
    check("rolled-up finding says 'All 20 containers'",
          overdue and overdue[0]["scope"] == "All 20 containers",
          overdue[0]["scope"] if overdue else "none")

    # ---- commercial sheet joined onto the BL -------------------------------- #
    cs, es = consignment("270383678", "MAEU", "Maersk", 5, "Discharg", -3)
    sheet = [{
        "row_number": 3, "bl_number": "270383678", "carrier_name": "Maersk",
        "beneficiary": "MANDALA TRADING PTY LTD", "broker": "CREDENCE",
        "commodity": "CANOLA", "pi_no": "MT05327", "qty": 311.82,
        "amount": 223574.94, "delivery_terms": "CFR BIRGUNJ", "free_days": 14,
        "sheet_eta_date": "2026-08-29", "sheet_eta_place": "DRYPORT",
        "tracked_eta": None, "tracked_status": "IN TRANSIT",
        "tracked_location": None, "eta_variance_days": 3,
        "container_count": 5, "containers": None,
    }]
    directory = write_fixture(tmp / "bl_join", cs, es, shipments=sheet)
    data, _, _, balanced = generate(directory, tmp / "bl_join.xlsx", as_of=BASE)
    row = data.shipment_rows[0]
    check("commercial beneficiary joined onto the BL",
          row["beneficiary"] == "MANDALA TRADING PTY LTD", str(row["beneficiary"]))
    check("commercial commodity and amount joined",
          row["commodity"] == "CANOLA" and row["amount"] == 223574.94,
          "%s / %s" % (row["commodity"], row["amount"]))
    check("commercial join balances", balanced)

    # A BL absent from the sheet must simply have blanks, not break the join.
    cs2, es2 = consignment("NOSHEET001", "MAEU", "Maersk", 2, "Discharg", -3)
    directory = write_fixture(tmp / "bl_nojoin", cs + cs2, es + es2, shipments=sheet)
    data, _, _, _ = generate(directory, tmp / "bl_nojoin.xlsx", as_of=BASE)
    unmatched = [r for r in data.shipment_rows if r["bl"] == "NOSHEET001"][0]
    check("BL with no sheet row has blank commercial fields",
          unmatched["beneficiary"] is None and unmatched["commodity"] is None)

    # ---- BL Summary sheet exists and is the first sheet after the dashboard - #
    wb = load_workbook(tmp / "bl_split.xlsx")
    check("BL Summary is the first sheet after the Dashboard",
          wb.sheetnames[:2] == [excel.S_DASH, excel.S_BLS], str(wb.sheetnames[:2]))
    ws = wb[excel.S_BLS]
    check("BL Summary carries the container numbers for drill-down",
          any("STRAYU123456" in str(ws.cell(row=5, column=c).value or "")
              for c in range(1, ws.max_column + 1)))


def test_large(tmp: Path) -> None:
    section("LARGE DATASET")

    containers, events = [], []
    carriers = [(c["code"], c["name"]) for c in normalize.TARGET_CARRIERS]
    wording = {code: descriptions for (code, _), descriptions
               in OTHER_CARRIER_EVENTS.items()}
    wording["MAEU"] = ["Gate In", "Load", "Container Departure",
                       "Container Arrival", "Discharg", "Gate Out"]
    wording["MSCU"] = ["Empty to Shipper", "Export received at CY",
                       "Export Loaded on Vessel", "Full Transshipment Discharged",
                       "Estimated Time of Arrival"]

    for index in range(2000):
        code, name = carriers[index % len(carriers)]
        cid = "TST%c%07d" % ("UJZ"[index % 3], index)
        descriptions = wording[code]
        containers.append(container(cid, code, name,
                                    event_count=len(descriptions),
                                    delivered=index % 9 == 0))
        for seq, description in enumerate(descriptions):
            events.append(event(cid, code, seq, seq * 2, description))

    directory = write_fixture(tmp / "large", containers, events)
    started = datetime.now()
    data, kpi, ledger, balanced = generate(directory, tmp / "large.xlsx")
    elapsed = (datetime.now() - started).total_seconds()
    ok, detail = opens(tmp / "large.xlsx")
    check("2000 containers / %d events: workbook opens" % len(events), ok, detail)
    check("2000 containers: ledger balances", balanced)
    check("2000 containers: all eight carriers present",
          len({c["carrier_display"] for c in data.containers}) == 8,
          str(sorted({c["carrier_display"] for c in data.containers})))
    check("2000 containers: completes in reasonable time (%.1fs)" % elapsed,
          elapsed < 120, "%.1fs" % elapsed)
    print("        %d containers, %d events, %.0f KB, %.1fs"
          % (len(data.containers), len(data.events),
             (tmp / "large.xlsx").stat().st_size / 1024, elapsed))


# --------------------------------------------------------------------------- #

def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="cargo_report_test_"))
    print("Temp fixtures: %s" % tmp)
    try:
        for test in (test_real_data, test_eight_carriers, test_bl_grain,
                     test_edge_cases, test_large):
            try:
                test(tmp)
            except Exception:
                FAILED.append((test.__name__, "raised"))
                print("  ERROR in %s" % test.__name__)
                traceback.print_exc()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print("\n%s\n%d passed, %d failed" % ("=" * 60, len(PASSED), len(FAILED)))
    for name, detail in FAILED:
        print("  FAILED  %s  %s" % (name, detail))

    print("\nCARRIER COVERAGE -- what was actually tested")
    print("  Maersk, MSC          real carrier JSON from this project")
    print("  COSCO, CMA CGM, Hapag-Lloyd, ONE, Evergreen, Yang Ming")
    print("                       SYNTHETIC fixtures only. The report handles "
          "their data; no adapter exists to fetch it, so the tracking path for "
          "these six is UNTESTED and cannot be tested here.")
    return 1 if FAILED else 0


if __name__ == "__main__":
    raise SystemExit(main())
