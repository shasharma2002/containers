#!/usr/bin/env python3
"""
Maersk DCSA Track & Trace adapter -- the DOCUMENTED developer.maersk.com API.

    pip install requests
    python maersk_dcsa.py APZU4812090
    python maersk_dcsa.py APZU4812090 --json
    python maersk_dcsa.py --bl 260029935

This is a different API from maersk_track.py. Which to use:

    maersk_track.py    api.maersk.com/synergy/tracking/...     internal, powers
                       maersk.com. Rich, but behind Akamai bot protection --
                       returns 403 to any non-browser client.

    maersk_dcsa.py     api.maersk.com/track-and-trace-private/events
                       THIS FILE. Published, supported, standards-based.
                       Auth instead of bot-blocking.

If this one works for you, use it and ignore the other two Maersk scripts.

AUTHENTICATION -- TWO CREDENTIALS, NOT ONE
------------------------------------------
The sample curl sends both:

    Consumer-Key:   <your app key from the portal>
    Authorization:  Bearer <access token>

The Consumer-Key identifies your app. The Bearer token is a short-lived OAuth
token you exchange your client credentials for -- it is NOT the same string as
the Consumer-Key, and it expires (usually ~1 hour).

Set CONSUMER_KEY below. Then either:

  (a) paste a token into ACCESS_TOKEN for a quick test -- it will expire; or
  (b) fill in CLIENT_ID / CLIENT_SECRET and let the script fetch tokens itself.

I could NOT verify the token URL -- it is not in the sample you have. Check
your portal for the OAuth / "Get access token" section. TOKEN_URL below is the
common Maersk pattern; if it 404s, replace it with the one from your docs.

Some portals issue only a Consumer-Key with no OAuth at all. If so, set
ACCESS_TOKEN = None and the script will send just the key.

WHAT "PRIVATE" MEANS
--------------------
`track-and-trace-private` is the customer-scoped product: it generally returns
events for shipments linked to YOUR Maersk account. Boxes on someone else's
bill of lading may return an empty list rather than an error. If you get 200
with no events for a container you can see on maersk.com, that is probably
what happened -- and it is an entitlement question for your Maersk rep, not a
bug in this script.

RESPONSE FORMAT -- READ THIS BEFORE TRUSTING THE PARSER
-------------------------------------------------------
This endpoint returns DCSA-standard events, which look nothing like the
internal API's nested containers/locations/events. DCSA is a flat array:

    [{
      "eventID": "...",
      "eventType": "EQUIPMENT",
      "eventClassifierCode": "ACT",          ACT=actual PLN=planned EST=estimated
      "eventDateTime": "2026-07-29T07:15:00Z",
      "equipmentEventTypeCode": "GTIN",      GTIN=gate-in LOAD=load DISC=discharge
      "equipmentReference": "APZU4812090",
      "transportCall": {
         "UNLocationCode": "AUFRE",
         "facilityCode": "...",
         "modeOfTransport": "VESSEL",
         "vessel": {"vesselIMONumber": "9321483", "vesselName": "..."},
         "exportVoyageNumber": "2103S"
      }
    }]

I have NOT seen your actual response, so the parser below is written against
the DCSA 2.0 spec and defensively (multiple key spellings, missing fields
tolerated). Run with --raw first and compare. If fields differ, paste me one
raw event and I will correct it exactly.

The good news: eventClassifierCode gives ACT/PLN/EST, so you still get the
forward schedule and real ETA-slip detection.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone

import requests

# --------------------------------------------------------------------------- #
# CREDENTIALS -- edit these
# --------------------------------------------------------------------------- #

CONSUMER_KEY = "jknZrwPsO463AW1SCXVcaLj7AgVnTfjM"

# Quick test: paste a token from the portal's "Try it" panel. Expires in ~1h.
ACCESS_TOKEN = "paste_a_token_from_the_portal"   # quick test, expires

# Preferred: client credentials, so the script refreshes tokens itself.
CLIENT_ID: str | None = None
CLIENT_SECRET: str | None = None

# UNVERIFIED -- confirm against your portal's OAuth section.
TOKEN_URL = "https://api.maersk.com/oauth2/access_token"

# --------------------------------------------------------------------------- #

EVENTS_URL = "https://api.maersk.com/track-and-trace-private/events"

# DCSA equipment event codes -> readable text
EQUIPMENT_EVENTS = {
    "GTIN": "Gate In", "GTOT": "Gate Out",
    "LOAD": "Loaded on Vessel", "DISC": "Discharged from Vessel",
    "STUF": "Stuffed", "STRP": "Stripped",
    "PICK": "Picked Up", "DROP": "Dropped Off",
    "INSP": "Inspected", "RSEA": "Resealed", "RMVD": "Removed",
    "AVFR": "Available for Release",
}
TRANSPORT_EVENTS = {"ARRI": "Arrival", "DEPA": "Departure"}
SHIPMENT_EVENTS = {
    "RECE": "Received", "DRFT": "Drafted", "PENA": "Pending Approval",
    "PENU": "Pending Update", "REJE": "Rejected", "APPR": "Approved",
    "ISSU": "Issued", "SURR": "Surrendered", "SUBM": "Submitted",
    "CONF": "Confirmed", "CMPL": "Completed", "HOLD": "On Hold",
}
MODE_MAP = {
    "VESSEL": "VESSEL", "RAIL": "RAIL", "TRUCK": "TRUCK",
    "BARGE": "BARGE", "AIR": "AIR",
    "1": "VESSEL", "2": "RAIL", "3": "TRUCK", "4": "AIR", "8": "BARGE",
}
CLASSIFIER = {"ACT": "COMPLETED", "PLN": "PLANNED", "EST": "PLANNED"}


class MaerskDcsaError(RuntimeError):
    pass


def _dig(obj: dict, *names, default=None):
    """First matching key, tolerating camelCase/PascalCase drift."""
    for n in names:
        if isinstance(obj, dict) and obj.get(n) is not None:
            return obj[n]
    return default


def get_token(verbose: bool = True) -> str | None:
    if ACCESS_TOKEN:
        return ACCESS_TOKEN
    if not (CLIENT_ID and CLIENT_SECRET):
        return None
    if verbose:
        print(f"# fetching token from {TOKEN_URL}", file=sys.stderr)
    resp = requests.post(
        TOKEN_URL,
        data={"grant_type": "client_credentials",
              "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET},
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30)
    if resp.status_code != 200:
        raise MaerskDcsaError(
            f"Token request failed ({resp.status_code}): {resp.text[:300]}\n"
            "Check TOKEN_URL against your portal's OAuth documentation.")
    data = resp.json()
    token = _dig(data, "access_token", "accessToken")
    if not token:
        raise MaerskDcsaError(f"No access_token in token response: {data}")
    return token


def fetch(number: str, by: str = "equipment", limit: int = 100,
          verbose: bool = True) -> list[dict]:
    """`by` is 'equipment' (container), 'bl', or 'booking'."""
    def log(m: str) -> None:
        if verbose:
            print(f"# {m}", file=sys.stderr)

    param = {"equipment": "equipmentReference",
             "bl": "transportDocumentReference",
             "booking": "carrierBookingReference"}[by]

    params = {
        param: number,
        "eventType": "SHIPMENT,EQUIPMENT,TRANSPORT",
        "limit": limit,
    }

    headers = {
        "API-Version": "1",
        "Consumer-Key": CONSUMER_KEY,
        "Accept": "application/json",
    }
    token = get_token(verbose)
    if token:
        headers["Authorization"] = f"Bearer {token}"
    else:
        log("no Bearer token set -- sending Consumer-Key only. "
            "If this 401s, you need an OAuth token (see the header of this file).")

    log(f"GET {EVENTS_URL}  {param}={number}")
    resp = requests.get(EVENTS_URL, params=params, headers=headers, timeout=45)
    log(f"status {resp.status_code}")

    if resp.status_code == 401:
        raise MaerskDcsaError(
            "401 -- credentials rejected. Either the Bearer token is missing or "
            "expired, or the Consumer-Key isn't entitled to this API product. "
            "Check the portal shows Track & Trace as subscribed for your app.")
    if resp.status_code == 403:
        body = resp.text[:200]
        if "<HTML" in body.upper():
            raise MaerskDcsaError(
                "403 from Akamai (HTML body) -- blocked at the edge, same as the "
                "internal API. Your key may not be provisioned for this host yet.")
        raise MaerskDcsaError(f"403 -- not entitled to this data: {body}")
    if resp.status_code == 404:
        raise MaerskDcsaError(
            f"404 -- endpoint not found. The path may differ for your subscription; "
            f"compare against the portal's sample request.")
    resp.raise_for_status()

    data = resp.json()
    events = data if isinstance(data, list) else (
        _dig(data, "events", "data", "items", default=[]))
    log(f"{len(events)} events")
    return events


# --------------------------------------------------------------------------- #
# parsing: DCSA events -> the same flat dict msc_track.parse returns
# --------------------------------------------------------------------------- #

def _iso(raw) -> str | None:
    if not raw:
        return None
    try:
        return datetime.fromisoformat(str(raw).replace("Z", "+00:00")).isoformat()
    except ValueError:
        return None


def _describe(ev: dict) -> str:
    etype = (_dig(ev, "eventType", "event_type") or "").upper()
    if etype == "EQUIPMENT":
        code = _dig(ev, "equipmentEventTypeCode", "equipment_event_type_code")
        return EQUIPMENT_EVENTS.get(code, code or "Equipment Event")
    if etype == "TRANSPORT":
        code = _dig(ev, "transportEventTypeCode", "transport_event_type_code")
        base = TRANSPORT_EVENTS.get(code, code or "Transport Event")
        tc = _dig(ev, "transportCall", "transport_call", default={}) or {}
        mode = (_dig(tc, "modeOfTransport", "mode_of_transport") or "").title()
        return f"{mode} {base}".strip()
    if etype == "SHIPMENT":
        code = _dig(ev, "shipmentEventTypeCode", "shipment_event_type_code")
        return SHIPMENT_EVENTS.get(code, code or "Shipment Event")
    return etype.title() or "Event"


def parse(events: list[dict], query: str = "") -> dict:
    if not events:
        raise MaerskDcsaError(
            f"No events returned for {query!r}. With the 'private' product this "
            "often means the shipment isn't linked to your account.")

    rows = []
    for ev in events:
        tc = _dig(ev, "transportCall", "transport_call", default={}) or {}
        loc = _dig(tc, "location", default={}) or {}
        vessel = _dig(tc, "vessel", default={}) or {}
        when = _iso(_dig(ev, "eventDateTime", "event_date_time"))
        classifier = (_dig(ev, "eventClassifierCode", "event_classifier_code")
                      or "ACT").upper()
        mode_raw = _dig(tc, "modeOfTransport", "mode_of_transport")

        rows.append({
            "when": when,
            "status": CLASSIFIER.get(classifier, "COMPLETED"),
            "description": _describe(ev),
            "unlocode": _dig(tc, "UNLocationCode", "unLocationCode",
                             default=_dig(loc, "UNLocationCode")),
            "city": _dig(loc, "locationName", "location_name"),
            "country": _dig(loc, "countryCode", "country_code"),
            "facility": _dig(tc, "facilityCode", "facility_code"),
            "mode": MODE_MAP.get(str(mode_raw).upper(), mode_raw) if mode_raw else None,
            "vessel_name": _dig(vessel, "vesselName", "vessel_name"),
            "vessel_imo": _dig(vessel, "vesselIMONumber", "vessel_imo_number"),
            "voyage": _dig(tc, "exportVoyageNumber", "export_voyage_number",
                           "importVoyageNumber"),
            "equipment": _dig(ev, "equipmentReference", "equipment_reference"),
            "iso_code": _dig(ev, "ISOEquipmentCode", "isoEquipmentCode"),
            "empty": _dig(ev, "emptyIndicatorCode", "empty_indicator_code"),
        })

    rows.sort(key=lambda r: r["when"] or "")

    actual_idx = [i for i, r in enumerate(rows) if r["status"] == "COMPLETED"]
    current_idx = actual_idx[-1] if actual_idx else None

    container_id = next((r["equipment"] for r in rows if r["equipment"]), query)
    iso_code = next((r["iso_code"] for r in rows if r["iso_code"]), None)

    out_events = []
    for i, r in enumerate(rows):
        out_events.append({
            "sequence": i,
            "status": "CURRENT" if i == current_idx else r["status"],
            "date": r["when"],
            "raw_date": (datetime.fromisoformat(r["when"]).strftime("%d/%m/%Y")
                         if r["when"] else None),
            "description": r["description"],
            "city": r["city"],
            "country": r["country"],
            "unlocode": r["unlocode"],
            "terminal": r["facility"],
            "terminal_code": r["facility"],
            "mode": r["mode"],
            "vessel_name": r["vessel_name"],
            "voyage": r["voyage"],
            "laden_status": ("EMPTY" if (r["empty"] or "").upper() == "EMPTY"
                             else "LADEN" if r["empty"] else None),
            "vessel_imo": r["vessel_imo"],
            "vessel_flag": None,
            "vessel_built": None,
        })

    planned = [e for e in out_events if e["status"] == "PLANNED"]
    current = out_events[current_idx] if current_idx is not None else {}

    return {
        "carrier": "MAEU",
        "query": query or container_id,
        "query_type": "Container",
        "bill_of_lading": None,
        "scraped_at": datetime.now(timezone.utc).isoformat(),
        "source_timestamp": None,
        "origin_city": out_events[0]["city"] if out_events else None,
        "origin_country": out_events[0]["country"] if out_events else None,
        "destination_city": out_events[-1]["city"] if out_events else None,
        "destination_country": out_events[-1]["country"] if out_events else None,
        "port_of_load": out_events[0]["unlocode"] if out_events else None,
        "port_of_discharge": out_events[-1]["unlocode"] if out_events else None,
        "transshipments": [],
        "final_pod_eta": planned[-1]["date"] if planned else None,
        "containers": [{
            "container_id": container_id,
            "container_type": iso_code,
            "latest_move": (f"{current.get('city')}, {current.get('country')}"
                            if current else None),
            "delivered": False,
            "pod_eta": planned[-1]["date"] if planned else None,
            "events": out_events,
        }],
    }


def main() -> int:
    ap = argparse.ArgumentParser(description="Maersk DCSA Track & Trace.")
    ap.add_argument("number")
    ap.add_argument("--bl", action="store_true", help="number is a BL reference")
    ap.add_argument("--booking", action="store_true", help="number is a booking ref")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--raw", action="store_true",
                    help="print the untouched DCSA events -- DO THIS FIRST")
    args = ap.parse_args()

    by = "bl" if args.bl else "booking" if args.booking else "equipment"
    try:
        events = fetch(args.number.strip().upper(), by=by)
    except MaerskDcsaError as exc:
        print(f"\nFAILED: {exc}", file=sys.stderr)
        return 1

    if args.raw:
        print(json.dumps(events, indent=2))
        return 0

    result = parse(events, query=args.number.strip().upper())
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        import maersk_track
        maersk_track.show(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())