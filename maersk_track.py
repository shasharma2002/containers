#!/usr/bin/env python3
"""
Maersk container tracker -- single file, same output shape as msc_track.py so
both feed the same Power BI tables.

    pip install requests
    python maersk_track.py HASU1149124
    python maersk_track.py MRKU7621507 --json

Built from a captured request:

    GET https://api.maersk.com/synergy/tracking/{number}?operator=MAEU
    accept: application/json
    consumer-key: <key>
    api-version: v2
    origin:  https://www.maersk.com
    referer: https://www.maersk.com/

ABOUT THE CONSUMER KEY -- READ THIS
-----------------------------------
The key below is the one Maersk's own website ships inside its JavaScript
bundle. It is not a secret, but it is not yours either, and there are two
practical problems with leaning on it:

  1. It rotates. When Maersk redeploys their frontend, this stops working and
     you have to re-capture it.
  2. Its rate limit is shared with every visitor to maersk.com. You can't see
     the budget, and heavy use may get the key throttled for everyone -- which
     is a good way to get the endpoint locked down.

The CONSUMER_KEY below is a key registered at developer.maersk.com. Edit that
one line if it ever changes -- no environment variables involved.

NOTE ON akamai-bm-telemetry
---------------------------
The browser sends a large `akamai-bm-telemetry` header. We do not, because it's
a JS-generated bot-sensor blob that can't be replayed. api.maersk.com appears to
be less defended than www.maersk.com (CORS is wide open: access-control-allow-
origin: *). If you start getting 403s, that assumption has changed.

WHAT MAERSK GIVES YOU THAT MSC DOESN'T
--------------------------------------
`event_time_type` is ACTUAL or EXPECTED, so you get the whole forward
itinerary, plus `eta_final_delivery`. That means genuine delay detection:
compare today's ETA against last week's. MSC returns actuals only.

The trade the other way: Maersk gives no vessel IMO (only an internal
`vessel_num` like "I2O"), where MSC gives real IMO numbers.

SHAPE OF THE RESPONSE
---------------------
Events are nested two levels deep -- containers[] > locations[] > events[] --
so a location's city/country applies to every event beneath it. We flatten that
and sort globally by event_time.

`location_code` is a 7-char Maersk internal code (AUFRE15, MYTPPTM). The first
five characters are the UN/LOCODE (AUFRE, MYTPP), which is what lets Maersk and
MSC rows join on the same `unlocode` column in Power BI.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone

try:
    from curl_cffi import requests as curl_requests
    HAS_CURL_CFFI = True
except ImportError:
    curl_requests = None
    HAS_CURL_CFFI = False

import requests

API = "https://api.maersk.com/synergy/tracking/{number}"

# Public key from Maersk's web bundle. Override with the env var -- see above.
# ---------------------------------------------------------------------------
# YOUR KEY. Edit this one line if it ever changes.
# ---------------------------------------------------------------------------
CONSUMER_KEY = "jknZrwPsO463AW1SCXVcaLj7AgVnTfjM"

FALLBACK_CONSUMER_KEY = "UtMm6JCDcGTnMGErNGvS2B98kt1Wl25H"   # Maersk's public web key

IMPERSONATE_PROFILES = ("chrome124", "chrome", "chrome110")

HEADERS = {
    "accept": "application/json",
    "accept-language": "en-US,en;q=0.9",
    "api-version": "v2",
    "consumer-key": CONSUMER_KEY,
    "origin": "https://www.maersk.com",
    "referer": "https://www.maersk.com/",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
}

# Maersk transport_mode -> our canonical mode. Anything not in here is passed
# through in upper case, which is why an unmapped code shows up raw in reports.
#
# FEF is not in Maersk's published list, but every one of the 1,316 FEF events
# in this project's own history carried an activity of Load, Discharg,
# Container Departure or Container Arrival -- all unambiguously vessel
# operations, and never a gate or rail move. Mapped on that evidence.
TRANSPORT_MODE = {
    "MVS": "VESSEL", "VSL": "VESSEL", "FEF": "VESSEL",
    "RR": "RAIL", "RAIL": "RAIL",
    "TRK": "TRUCK", "TRC": "TRUCK", "RD": "TRUCK",
    "BAR": "BARGE", "AIR": "AIR",
}

# Activities that are yard/gate moves rather than a transport leg
GATE_ACTIVITIES = {"GATE-IN", "GATE-OUT", "GATE IN", "GATE OUT",
                   "DELIVERY", "PICKUP", "EMPTY RETURN"}


class MaerskError(RuntimeError):
    pass


# --------------------------------------------------------------------------- #
# transport
# --------------------------------------------------------------------------- #

def fetch(number: str, operator: str = "MAEU",
          timeout: int = 30, verbose: bool = True) -> dict:
    def log(msg: str) -> None:
        if verbose:
            print(f"# {msg}", file=sys.stderr)

    log(f"consumer-key: {CONSUMER_KEY[:8]}...{CONSUMER_KEY[-4:]}")

    url = API.format(number=number)
    params = {"operator": operator}
    profiles = list(IMPERSONATE_PROFILES) if HAS_CURL_CFFI else [None]
    last = None

    for profile in profiles:
        label = f"impersonate={profile}" if profile else "plain requests"
        if profile and HAS_CURL_CFFI:
            session = curl_requests.Session(impersonate=profile)
            session.headers.update(HEADERS)
        else:
            session = requests.Session()
            session.headers.update(HEADERS)
            session.headers["user-agent"] = (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36")

        try:
            resp = session.get(url, params=params, timeout=timeout)
            last = resp.status_code
            log(f"{label}: GET -> {resp.status_code}")

            if resp.status_code == 403:
                continue
            if resp.status_code == 401:
                raise MaerskError(
                    "401 -- consumer-key rejected. Maersk rotated their web key. "
                    "Re-capture it from DevTools, or register at developer.maersk.com.")
            if resp.status_code == 404:
                raise MaerskError(
                    f"404 -- Maersk has no record of {number!r}. Check the number, "
                    "or try --operator SEAU / SEJJ / MCPU for other group lines.")
            resp.raise_for_status()

            data = resp.json()
            if not data.get("containers"):
                raise MaerskError(f"No containers in response for {number!r}")
            return data

        except MaerskError:
            raise
        except Exception as exc:
            log(f"{label}: {type(exc).__name__}: {exc}")
            continue

    raise MaerskError(
        f"All transports failed (last status {last}). "
        + ("Try: pip install curl_cffi" if not HAS_CURL_CFFI
           else "api.maersk.com may now require the akamai-bm-telemetry header, "
                "which cannot be replayed without a browser."))


# --------------------------------------------------------------------------- #
# parsing
# --------------------------------------------------------------------------- #

def _iso(raw: str | None) -> str | None:
    """Maersk sends '2026-08-02T10:12:00.000' and sometimes a trailing Z."""
    if not raw or not str(raw).strip():
        return None
    text = str(raw).strip().replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(text).isoformat()
    except ValueError:
        return None


def _clean(v):
    if v is None:
        return None
    if isinstance(v, str) and not v.strip():
        return None
    return v


def _mode(activity: str | None, transport_mode: str | None) -> str | None:
    act = (activity or "").upper()
    if act in GATE_ACTIVITIES or act.startswith("GATE"):
        return "GATE"
    if transport_mode:
        return TRANSPORT_MODE.get(transport_mode.upper(), transport_mode.upper())
    if "RAIL" in act:
        return "RAIL"
    if "VESSEL" in act or "CONTAINER" in act or act in ("LOAD", "DISCHARGE"):
        return "VESSEL"
    return None


def _humanise(activity: str | None) -> str:
    """'RAIL_ARRIVAL_AT_DESTINATION' -> 'Rail Arrival At Destination'."""
    if not activity:
        return ""
    return activity.replace("_", " ").replace("-", " ").title()


def parse(data: dict, query: str = "") -> dict:
    """Maersk payload -> the same flat dict msc_track.parse returns."""
    containers = data.get("containers") or []
    if not containers:
        raise MaerskError(f"No containers for {query!r}")

    top_origin = data.get("origin") or {}
    top_dest = data.get("destination") or {}

    out = {
        "carrier": "MAEU",
        "query": query or _clean(containers[0].get("container_num")),
        "query_type": "Container" if data.get("is_container_search") else "Shipment",
        "bill_of_lading": None,          # not returned by this endpoint
        "scraped_at": datetime.now(timezone.utc).isoformat(),
        "source_timestamp": _iso(data.get("last_update_time")),
        "origin_city": _clean(top_origin.get("city")),
        "origin_country": _clean(top_origin.get("country_code")),
        "destination_city": _clean(top_dest.get("city")),
        "destination_country": _clean(top_dest.get("country_code")),
        "port_of_load": _clean(top_origin.get("city")),
        "port_of_discharge": None,       # filled from the last vessel arrival
        "transshipments": [],
        "final_pod_eta": None,
        "containers": [],
    }

    for c in containers:
        locations = c.get("locations") or []

        # Flatten locations > events, carrying the location down onto each event.
        flat = []
        for loc in locations:
            city = _clean(loc.get("city"))
            country = _clean(loc.get("country_code"))
            loc_code = _clean(loc.get("location_code"))
            for ev in loc.get("events") or []:
                flat.append({"loc": loc, "city": city, "country": country,
                             "loc_code": loc_code, "ev": ev})

        flat.sort(key=lambda r: _iso(r["ev"].get("event_time")) or "")

        # Last ACTUAL event is the container's current position.
        actual_idx = [i for i, r in enumerate(flat)
                      if (r["ev"].get("event_time_type") or "").upper() == "ACTUAL"]
        current_idx = actual_idx[-1] if actual_idx else None

        events = []
        for i, r in enumerate(flat):
            ev, loc_code = r["ev"], r["loc_code"]
            is_actual = (ev.get("event_time_type") or "").upper() == "ACTUAL"
            if i == current_idx:
                status = "CURRENT"
            elif is_actual:
                status = "COMPLETED"
            else:
                status = "PLANNED"

            iso = _iso(ev.get("event_time"))
            events.append({
                "sequence": i,
                "status": status,
                "date": iso,
                "raw_date": (datetime.fromisoformat(iso).strftime("%d/%m/%Y")
                             if iso else None),
                "description": _humanise(ev.get("activity")),
                "city": r["city"],
                "country": r["country"],
                # First 5 chars of Maersk's internal code = the UN/LOCODE,
                # which is what makes these rows joinable with the MSC ones.
                "unlocode": loc_code[:5] if loc_code else None,
                "terminal": _clean(r["loc"].get("terminal")),
                "terminal_code": loc_code,
                "mode": _mode(ev.get("activity"), ev.get("transport_mode")),
                "vessel_name": _clean(ev.get("vessel_name")),
                "voyage": _clean(ev.get("voyage_num")),
                "laden_status": "EMPTY" if ev.get("stempty") else "LADEN",
                "vessel_imo": None,          # Maersk gives no IMO here
                "vessel_flag": None,
                "vessel_built": None,
            })

        # POD = city of the last vessel arrival on the itinerary.
        vessel_arrivals = [e for e in events
                           if e["mode"] == "VESSEL" and "Arrival" in e["description"]]
        if vessel_arrivals and not out["port_of_discharge"]:
            out["port_of_discharge"] = vessel_arrivals[-1]["city"]

        # Transshipments = intermediate locations between origin and destination.
        if not out["transshipments"] and len(locations) > 2:
            out["transshipments"] = [
                f"{_clean(l.get('city'))}, {_clean(l.get('country_code'))}"
                for l in locations[1:-1]
            ]

        eta = _iso(c.get("eta_final_delivery"))
        if eta and not out["final_pod_eta"]:
            out["final_pod_eta"] = eta

        size, ctype = _clean(c.get("container_size")), _clean(c.get("container_type"))
        current = events[current_idx] if current_idx is not None else {}

        out["containers"].append({
            "container_id": _clean(c.get("container_num")),
            "container_type": f"{size}' {ctype}" if size and ctype else (ctype or None),
            "latest_move": (f"{current.get('city')}, {current.get('country')}"
                            if current else None),
            "delivered": (c.get("status") or "").upper() in
                         ("COMPLETED", "DELIVERED", "FINISHED"),
            "pod_eta": eta,
            "events": events,
        })

    return out


# --------------------------------------------------------------------------- #
# output
# --------------------------------------------------------------------------- #

def show(r: dict) -> None:
    print(f"\n{r['carrier']}  {r['query_type']} {r['query']}")
    print(f"{r['origin_city']}, {r['origin_country']}  ->  "
          f"{r['destination_city']}, {r['destination_country']}")
    if r["transshipments"]:
        print("via " + "  ->  ".join(r["transshipments"]))
    if r["final_pod_eta"]:
        print(f"ETA final delivery: {r['final_pod_eta'][:10]}")
    if r["source_timestamp"]:
        print(f"carrier last updated: {r['source_timestamp'][:16]}")

    for c in r["containers"]:
        print(f"\n  {c['container_id']}  {c['container_type']}"
              f"   {'DELIVERED' if c['delivered'] else 'in transit'}")
        for e in c["events"]:
            marker = {"CURRENT": ">", "PLANNED": "~"}.get(e["status"], " ")
            loc = f"{e['city']}, {e['country']}"
            ship = f"{e['vessel_name']} {e['voyage']}" if e["vessel_name"] else ""
            print(f"  {marker} {e['raw_date']}  {loc:<18} {(e['unlocode'] or ''):<6} "
                  f"{(e['mode'] or ''):<7} {e['description']:<28} {ship}")
    print("\n  legend:  (blank) done   > current   ~ expected\n")


def main() -> int:
    ap = argparse.ArgumentParser(description="Track a Maersk container.")
    ap.add_argument("number")
    ap.add_argument("--operator", default="MAEU",
                    help="MAEU (default), SEAU=Sealand, SEJJ, MCPU")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--raw", action="store_true")
    ap.add_argument("-q", "--quiet", action="store_true")
    args = ap.parse_args()

    number = args.number.strip().upper()
    try:
        data = fetch(number, operator=args.operator, verbose=not args.quiet)
    except MaerskError as exc:
        print(f"\nFAILED: {exc}", file=sys.stderr)
        return 1

    if args.raw:
        print(json.dumps(data, indent=2))
        return 0

    result = parse(data, query=number)
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        show(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())