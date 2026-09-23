#!/usr/bin/env python3
"""
Offline tests for the tracking layer. No network, no browser.

    python test_tracking.py

Replaces the old test_pipeline.py, which had been broken for some time: it
loaded a captured payload named HASU1149124.json that is not in the project, so
it raised FileNotFoundError before asserting anything.

The parity check it was written for is worth keeping, so it is rebuilt here from
payloads constructed to match each adapter's documented envelope rather than a
captured file that can go missing.

WHAT IS COVERED
---------------
  1. SCHEMA PARITY -- MSC and Maersk rows must carry identical key sets. They
     land in the same containers.json and the same Power BI table; if one
     carrier omits a column, an append silently drops it for every row.
  2. ONE SCHEMA DEFINITION -- track_all.py and maersk_browser.py must share
     pipeline.flatten. They each had their own copy once, and the result was
     two different shapes in one output file.
  3. CARRIER ROUTING -- all eight target carriers classify from their BIC
     prefixes, and leasing prefixes stay deliberately unresolved.
  4. SHEET PARSING -- the commercial workbook parser, against
     test_fixture_sheet.xlsx.

WHAT IS NOT COVERED
-------------------
Anything that needs a live carrier. fetch() is never called. Whether Maersk's
consumer key still works, or whether the browser path still captures the XHR,
can only be found out by running for real -- use maersk_debug.py for that.
"""

from __future__ import annotations

import inspect
import sys
import traceback
from pathlib import Path

import carrier_id
import maersk_browser
import maersk_track
import msc_track
import pipeline
import read_sheet
import track_all

PASSED: list[str] = []
FAILED: list[tuple[str, str]] = []


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
# payloads -- shaped to each adapter's documented envelope
# --------------------------------------------------------------------------- #

MSC_PAYLOAD = {
    "TrackingNumber": "MEDUXK845261",
    "TrackingType": "Bill of Lading",
    "TrackingResultsLabel": "Data as of 18/09/2026 09:00",
    "BillOfLadings": [{
        "BillOfLadingNumber": "MEDUXK845261",
        "GeneralTrackingInfo": {
            "ShippedFrom": "MERSIN, TR",
            "ShippedTo": "BIRGUNJ, NP",
            "PortOfLoad": "MERSIN, TR",
            "PortOfDischarge": "KOLKATA, IN",
            "Transshipments": ["GIOIA TAURO, IT", "COLOMBO, LK"],
            "FinalPodEtaDate": "05/10/2026",
        },
        "ContainersInfo": [{
            "ContainerNumber": "TCLU7686420",
            "ContainerType": "20' DRY VAN",
            "LatestMove": "BIRGUNJ, NP",
            "Delivered": True,
            "PodEtaDate": "05/10/2026",
            "Events": [
                {
                    "Order": 1, "Date": "14/08/2026",
                    "Description": "Export received at CY",
                    "Location": "MERSIN, TR", "UnLocationCode": "TRMER",
                    "EquipmentHandling": {"Name": "MIP TERMINAL", "Smdg": "MERSIN"},
                    "Detail": ["LADEN"], "Vessel": {},
                },
                {
                    "Order": 2, "Date": "16/08/2026",
                    "Description": "Export Loaded on Vessel",
                    "Location": "MERSIN, TR", "UnLocationCode": "TRMER",
                    "EquipmentHandling": {"Name": "MIP TERMINAL", "Bic": "MERSIN"},
                    "Detail": ["MSC AKITETA II", "SH629R"],
                    "Vessel": {"IMO": "9306263", "FlagName": "LIBERIA",
                               "Built": "2005"},
                },
            ],
        }],
    }],
}

MAERSK_PAYLOAD = {
    "is_container_search": True,
    "last_update_time": "2026-09-18T09:30:00.000",
    "origin": {"city": "NORFOLK", "country_code": "US"},
    "destination": {"city": "BIRGANJ", "country_code": "NP"},
    "containers": [{
        "container_num": "TCNU7075603",
        "container_size": "40",
        "container_type": "DRY",
        "status": "IN TRANSIT",
        "eta_final_delivery": "2026-10-08T03:10:00.000",
        "locations": [
            {
                "city": "NORFOLK", "country_code": "US",
                "location_code": "USORF01", "terminal": "NIT SOUTH",
                "events": [
                    {"event_time": "2026-08-20T10:00:00.000",
                     "event_time_type": "ACTUAL", "activity": "GATE-IN",
                     "transport_mode": None, "stempty": False},
                    {"event_time": "2026-08-22T18:00:00.000",
                     "event_time_type": "ACTUAL", "activity": "LOAD",
                     "transport_mode": "FEF", "vessel_name": "ULSAN EXPRESS",
                     "voyage_num": "631W", "stempty": False},
                ],
            },
            {
                "city": "BIRGANJ", "country_code": "NP",
                "location_code": "NPBIRHT",
                "terminal": "HIMALAYAN TERMINALS PRIVATE LIMITED",
                "events": [
                    {"event_time": "2026-10-08T03:10:00.000",
                     "event_time_type": "EXPECTED",
                     "activity": "RAIL_ARRIVAL_AT_DESTINATION",
                     "transport_mode": "RR", "stempty": False},
                ],
            },
        ],
    }],
}


# --------------------------------------------------------------------------- #

def test_schema_parity() -> None:
    section("SCHEMA PARITY -- both carriers must produce identical key sets")

    msc = msc_track.parse(MSC_PAYLOAD, query="MEDUXK845261")
    maersk = maersk_track.parse(MAERSK_PAYLOAD, query="274031020")

    check("MSC payload parses", bool(msc["containers"]))
    check("Maersk payload parses", bool(maersk["containers"]))

    # The intermediate parsed shape both adapters promise.
    check("adapters agree on the parsed top-level shape",
          set(msc) - {"query_type"} == set(maersk) - {"query_type"},
          "MSC only: %s | Maersk only: %s"
          % (sorted(set(msc) - set(maersk)), sorted(set(maersk) - set(msc))))

    msc_c, msc_e = pipeline.flatten(msc, "run0001", "2026-09-18T09:00:00+00:00")
    mae_c, mae_e = pipeline.flatten(maersk, "run0001", "2026-09-18T09:00:00+00:00")

    check("container rows have identical columns",
          set(msc_c[0]) == set(mae_c[0]),
          "MSC only: %s | Maersk only: %s"
          % (sorted(set(msc_c[0]) - set(mae_c[0])),
             sorted(set(mae_c[0]) - set(msc_c[0]))))
    check("event rows have identical columns",
          set(msc_e[0]) == set(mae_e[0]),
          "MSC only: %s | Maersk only: %s"
          % (sorted(set(msc_e[0]) - set(mae_e[0])),
             sorted(set(mae_e[0]) - set(msc_e[0]))))

    check("tracking_number present on every row",
          all("tracking_number" in r for r in msc_c + mae_c + msc_e + mae_e))
    check("MSC tracking_number is the confirmed BL",
          msc_c[0]["tracking_number"] == "MEDUXK845261",
          msc_c[0]["tracking_number"])
    check("Maersk tracking_number falls back to the query",
          mae_c[0]["tracking_number"] == "274031020"
          and mae_c[0]["bill_of_lading"] is None,
          "%s / %s" % (mae_c[0]["tracking_number"], mae_c[0]["bill_of_lading"]))

    # Forward schedule: only Maersk supplies one, and has_forecast must say so.
    check("Maersk forward schedule flagged", mae_c[0]["has_forecast"] is True)
    check("MSC has no forward schedule", msc_c[0]["has_forecast"] is False)
    check("planned events are kept as PLANNED",
          any(e["status"] == "PLANNED" for e in mae_e))

    # FEF used to pass through raw and show as a transport code in reports.
    load = [e for e in mae_e if e["description"] == "Load"]
    check("FEF transport mode resolves to VESSEL",
          load and load[0]["mode"] == "VESSEL",
          load[0]["mode"] if load else "no Load event")

    # UN/LOCODE is what lets the two carriers join on location.
    check("Maersk location_code truncated to a UN/LOCODE",
          {e["unlocode"] for e in mae_e} == {"USORF", "NPBIR"},
          str({e["unlocode"] for e in mae_e}))
    check("MSC UN/LOCODE passed through",
          {e["unlocode"] for e in msc_e} == {"TRMER"},
          str({e["unlocode"] for e in msc_e}))

    # MSC's Detail array is vessel/voyage for a sailing, laden status for a yard
    # move -- indexed, never regex-matched.
    loaded = [e for e in msc_e if e["description"] == "Export Loaded on Vessel"]
    check("MSC vessel and voyage read from the Detail array",
          loaded and loaded[0]["vessel_name"] == "MSC AKITETA II"
          and loaded[0]["voyage"] == "SH629R",
          str(loaded[0]) if loaded else "missing")
    check("MSC vessel IMO captured where Maersk has none",
          loaded and loaded[0]["vessel_imo"] == "9306263"
          and all(e["vessel_imo"] is None for e in mae_e))


def test_one_schema_definition() -> None:
    section("ONE SCHEMA DEFINITION -- no entry point may keep its own copy")

    check("track_all uses pipeline.flatten",
          track_all.flatten is pipeline.flatten)
    check("maersk_browser uses pipeline.flatten",
          maersk_browser.flatten is pipeline.flatten)
    check("track_all uses pipeline.write_json",
          track_all.write_json is pipeline.write_json)
    check("carrier names have one definition",
          track_all.CARRIER_NAMES is pipeline.CARRIER_NAMES)

    # A local def would shadow the import and reintroduce the drift.
    for module in (track_all, maersk_browser):
        source = Path(module.__file__).read_text(encoding="utf-8")
        check("%s does not redefine flatten" % Path(module.__file__).name,
              "\ndef flatten(" not in source)

    check("shipment rows carry a run stamp",
          "run_id" in inspect.signature(track_all.build_shipments).parameters)


def test_carrier_routing() -> None:
    section("CARRIER ROUTING -- all eight target carriers")

    expected = {
        "MAEU1234567": "Maersk (incl. Hamburg Sud)",
        "MSCU1234567": "MSC",
        "COSU1234567": "COSCO",
        "CMAU1234567": "CMA CGM",
        "HLCU1234567": "Hapag-Lloyd",
        "ONEU1234567": "ONE",
        "EGHU1234567": "Evergreen",
        "YMLU1234567": "Yang Ming",
    }
    for number, name in expected.items():
        result = carrier_id.classify(number)
        check("%s -> %s" % (number, name),
              result["carrier_name"] == name and result["confidence"] == "high",
              "got %s (%s)" % (result["carrier_name"], result["confidence"]))

    # A leasing prefix says nothing about the carrier, and must NOT resolve --
    # the project's own data has leased boxes on MSC.
    for number in ("TCLU7686420", "CAIU3056659", "TLLU2467264"):
        result = carrier_id.classify(number)
        check("%s stays unresolved (leasing prefix)" % number,
              result["carrier"] is None and result["confidence"] == "unknown",
              "got %s" % result["carrier"])

    check("a bare numeric Maersk BL is not mistaken for a container",
          carrier_id.classify("270383678")["kind"] != "container")
    check("an MSC BL prefix resolves",
          carrier_id.classify("MEDUXK845261")["carrier"] == "MSCU")


def test_sheet_parsing() -> None:
    section("SHEET PARSING -- the commercial workbook")

    fixture = Path("test_fixture_sheet.xlsx")
    if not fixture.exists():
        check("fixture sheet present", False, "%s missing" % fixture)
        return

    result = read_sheet.get_shipments(fixture)
    shipments = result["shipments"]
    check("fixture sheet parses", bool(shipments), "%d rows" % len(shipments))
    check("every row has a BL", all(s["bl_number"] for s in shipments))
    check("commercial fields captured",
          any(s["beneficiary"] and s["commodity"] for s in shipments))

    # The header is not on row 1 -- there are title rows above it. Finding it by
    # matching the BL NO. header rather than assuming an offset is the point.
    check("header located below title rows", bool(shipments))

    # SHIPPING LINE, not the BL prefix, decides the carrier on this sheet.
    check("MAERSK maps to MAEU",
          read_sheet.classify_shipping_line("MAERSK") == "MAEU")
    check("MSC maps to MSCU",
          read_sheet.classify_shipping_line("MSC") == "MSCU")
    check("HAPAG deliberately does not map here",
          read_sheet.classify_shipping_line("HAPAG") is None)
    check("blank shipping line does not map",
          read_sheet.classify_shipping_line(None) is None)

    # Refusing to guess the BL column is the safety property worth testing.
    try:
        read_sheet.get_shipments(Path("bl_list.xlsx"))
        real_sheet_ok = True
    except read_sheet.SheetError:
        real_sheet_ok = False
    check("the live bl_list.xlsx still parses", real_sheet_ok,
          "BL NO. header not found -- check the sheet layout")


def main() -> int:
    for test in (test_schema_parity, test_one_schema_definition,
                 test_carrier_routing, test_sheet_parsing):
        try:
            test()
        except Exception:
            FAILED.append((test.__name__, "raised"))
            print("  ERROR in %s" % test.__name__)
            traceback.print_exc()

    print("\n%s\n%d passed, %d failed" % ("=" * 60, len(PASSED), len(FAILED)))
    for name, detail in FAILED:
        print("  FAILED  %s  %s" % (name, detail))
    print("\nNo network was used. Whether the carriers still ANSWER can only be "
          "found out by running for real -- see maersk_debug.py.")
    return 1 if FAILED else 0


if __name__ == "__main__":
    raise SystemExit(main())
