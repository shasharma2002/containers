#!/usr/bin/env python3
"""
MSC container / BL tracker -- single file, no other project files needed.

    pip install requests
    pip install curl_cffi          # optional, only if you get 403s

    python msc_track.py TCLU7686420
    python msc_track.py MEDUXK845261
    python msc_track.py TCLU7686420 --json
    python msc_track.py 123456789 --booking

Everything below is built from a real captured request, not guesswork:

  POST https://www.msc.com/api/feature/tools/TrackingInfo
  Content-Type: application/json
  X-Requested-With: XMLHttpRequest
  body: {"trackingMode":"0","trackingNumber":"TCLU7686420"}

Three things worth knowing, because they are not what the page markup implied:

1. NO CSRF TOKEN. The page renders <input name="__RequestVerificationToken">,
   but the actual XHR never sends it. Don't bother scraping it.

2. trackingMode and trackingNumber are STRINGS ("0", not 0). Content-Length was
   51 bytes, which only works with quoted values. ASP.NET model binding is
   usually lenient here, but send what the browser sends.

3. Detail is an ARRAY, not a string: ["MSC AKITETA II","SH628A"] for a vessel
   leg, ["LADEN"] or ["EMPTY"] for a yard move. So vessel/voyage need no regex
   at all -- just index the list.

Response envelope is {"IsSuccess":bool,"Data":{...}}; the useful part is
Data.BillOfLadings[].ContainersInfo[].Events[].

Events arrive NEWEST FIRST but carry an `Order` integer that ascends
chronologically, so we sort on Order rather than trusting array position.

MSC returns ACTUALS ONLY -- no forward schedule. `PodEtaDate` and
`FinalPodEtaDate` are the only forward-looking fields and are often "".
If you build delay detection, diff those across polls; you cannot diff planned
events the way you can with Maersk.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone

try:
    from curl_cffi import requests as curl_requests
    HAS_CURL_CFFI = True
except ImportError:
    curl_requests = None
    HAS_CURL_CFFI = False

import requests

BASE = "https://www.msc.com"
PAGE = f"{BASE}/en/track-a-shipment"
API = f"{BASE}/api/feature/tools/TrackingInfo"

MODE_CONTAINER_OR_BL = "0"
MODE_BOOKING = "1"

IMPERSONATE_PROFILES = ("chrome124", "chrome", "chrome110", "safari17_0")

FIREFOX_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:153.0) "
              "Gecko/20100101 Firefox/153.0")

# Copied from the captured XHR. Origin + Referer + X-Requested-With are the ones
# that matter: this is a same-origin CORS call and the server checks it looks
# like one.
API_HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Content-Type": "application/json",
    "Origin": BASE,
    "Referer": PAGE,
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "X-Requested-With": "XMLHttpRequest",
}


class MscError(RuntimeError):
    pass


# --------------------------------------------------------------------------- #
# transport
# --------------------------------------------------------------------------- #

def _make_session(profile: str | None):
    """
    curl_cffi impersonation when available, else plain requests.

    When impersonating, curl_cffi sends a header set matching its TLS
    fingerprint -- so we do NOT override User-Agent there. Overriding it desyncs
    the handshake from the headers, which is a louder bot signal than either.
    """
    if profile and HAS_CURL_CFFI:
        return curl_requests.Session(impersonate=profile)
    session = requests.Session()
    session.headers["User-Agent"] = FIREFOX_UA
    return session


def fetch(number: str, mode: str = MODE_CONTAINER_OR_BL,
          timeout: int = 30, verbose: bool = True) -> dict:
    """POST to the tracking API, warming cookies from the page first."""
    def log(msg: str) -> None:
        if verbose:
            print(f"# {msg}", file=sys.stderr)

    profiles = list(IMPERSONATE_PROFILES) if HAS_CURL_CFFI else [None]
    if not HAS_CURL_CFFI:
        log("curl_cffi not installed; using plain requests "
            "(pip install curl_cffi if you hit 403s)")

    payload = {"trackingMode": mode, "trackingNumber": number}
    last = None

    for profile in profiles:
        label = f"impersonate={profile}" if profile else "plain requests"
        session = _make_session(profile)
        try:
            # Warm-up GET: banks ASP.NET_SessionId and friends. A 403 here is
            # not fatal -- the API path is sometimes less protected than the
            # page, so we try the POST regardless.
            try:
                warm = session.get(PAGE, timeout=timeout)
                log(f"{label}: page GET -> {warm.status_code}")
            except Exception as exc:
                log(f"{label}: page GET failed ({type(exc).__name__}), continuing")

            resp = session.post(API, json=payload, headers=API_HEADERS, timeout=timeout)
            last = resp.status_code
            log(f"{label}: API POST -> {resp.status_code}")

            if resp.status_code == 403:
                continue
            resp.raise_for_status()

            try:
                data = resp.json()
            except ValueError:
                log(f"{label}: non-JSON body: {resp.text[:200]!r}")
                continue

            if not data.get("IsSuccess"):
                raise MscError(
                    f"MSC returned IsSuccess=false for {number!r}. "
                    "Usually means the number is unknown, or it's a BL when you "
                    "sent booking mode (or vice versa)."
                )
            return data["Data"]

        except MscError:
            raise
        except Exception as exc:
            log(f"{label}: {type(exc).__name__}: {exc}")
            continue

    raise MscError(
        f"All transports blocked (last status {last}).\n"
        + ("Install curl_cffi and retry: pip install curl_cffi"
           if not HAS_CURL_CFFI else
           "curl_cffi was blocked too, so Akamai is enforcing the _abck sensor\n"
           "cookie on this path. That needs a real browser (Playwright) or a\n"
           "different egress IP. Check the page loads in your normal browser first.")
    )


# --------------------------------------------------------------------------- #
# parsing
# --------------------------------------------------------------------------- #

def _iso(raw: str | None) -> str | None:
    """MSC dates are dd/MM/yyyy. Empty string means 'no value'."""
    if not raw or not raw.strip():
        return None
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw.strip(), fmt).isoformat()
        except ValueError:
            continue
    return None


def _split_location(raw: str | None) -> tuple[str | None, str | None]:
    """'HALDIA, IN' -> ('HALDIA', 'IN')"""
    if not raw:
        return None, None
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if len(parts) >= 2:
        return parts[0], parts[-1]
    return (parts[0] if parts else None), None


def _clean(value):
    """MSC uses '' and null interchangeably for missing. Normalise to None."""
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    return value


def _mode(description: str | None, has_vessel: bool) -> str | None:
    d = (description or "").lower()
    if any(k in d for k in ("vessel", "transshipment")):
        return "VESSEL"
    if "rail" in d:
        return "RAIL"
    if "truck" in d or "road" in d:
        return "TRUCK"
    if any(k in d for k in ("empty to", "received at cy", "gate", "delivered")):
        return "GATE"
    return "VESSEL" if has_vessel else None


def parse(data: dict, query: str = "") -> dict:
    """MSC payload -> flat, stable dict. This is what you persist."""
    bols = data.get("BillOfLadings") or []
    if not bols:
        raise MscError(f"No BillOfLadings for {query!r}")

    general = bols[0].get("GeneralTrackingInfo") or {}
    o_city, o_country = _split_location(general.get("ShippedFrom"))
    d_city, d_country = _split_location(general.get("ShippedTo"))
    pol_city, _ = _split_location(general.get("PortOfLoad"))
    pod_city, _ = _split_location(general.get("PortOfDischarge"))

    out = {
        "carrier": "MSCU",
        "query": query or _clean(data.get("TrackingNumber")),
        "query_type": _clean(data.get("TrackingType")),
        "bill_of_lading": _clean(bols[0].get("BillOfLadingNumber")),
        "scraped_at": datetime.now(timezone.utc).isoformat(),
        "source_timestamp": _clean(data.get("TrackingResultsLabel")),
        "origin_city": o_city, "origin_country": o_country,
        "destination_city": d_city, "destination_country": d_country,
        "port_of_load": pol_city, "port_of_discharge": pod_city,
        "transshipments": [t for t in (general.get("Transshipments") or []) if t],
        "final_pod_eta": _iso(general.get("FinalPodEtaDate")),
        "containers": [],
    }

    for bol in bols:
        for c in bol.get("ContainersInfo") or []:
            events = []
            # Order ascends chronologically; the array itself is newest-first.
            raw_events = sorted(c.get("Events") or [],
                                key=lambda e: e.get("Order", 0))
            for i, ev in enumerate(raw_events):
                detail = ev.get("Detail") or []
                if isinstance(detail, str):            # defensive
                    detail = [detail]
                vessel_name = voyage = None
                if len(detail) >= 2:
                    vessel_name, voyage = detail[0], detail[1]
                laden = detail[0] if len(detail) == 1 else None

                vessel = ev.get("Vessel") or {}
                handling = ev.get("EquipmentHandling") or {}
                city, country = _split_location(ev.get("Location"))

                events.append({
                    "sequence": i,
                    "status": "CURRENT" if i == len(raw_events) - 1 else "COMPLETED",
                    "date": _iso(ev.get("Date")),
                    "raw_date": _clean(ev.get("Date")),
                    "description": _clean(ev.get("Description")),
                    "city": city,
                    "country": country,
                    "unlocode": _clean(ev.get("UnLocationCode")),
                    "terminal": _clean(handling.get("Name")),
                    "terminal_code": (_clean(handling.get("Smdg"))
                                      or _clean(handling.get("Bic"))
                                      or _clean(handling.get("Code"))),
                    "mode": _mode(ev.get("Description"), bool(_clean(vessel.get("IMO")))),
                    "vessel_name": vessel_name,
                    "voyage": voyage,
                    "laden_status": laden,
                    "vessel_imo": _clean(vessel.get("IMO")),
                    "vessel_flag": _clean(vessel.get("FlagName")),
                    "vessel_built": _clean(vessel.get("Built")),
                })

            out["containers"].append({
                "container_id": _clean(c.get("ContainerNumber")),
                "container_type": _clean(c.get("ContainerType")),
                "latest_move": _clean(c.get("LatestMove")),
                "delivered": bool(c.get("Delivered")),
                "pod_eta": _iso(c.get("PodEtaDate")),
                "events": events,
            })

    return out


# --------------------------------------------------------------------------- #
# output
# --------------------------------------------------------------------------- #

def show(r: dict) -> None:
    print(f"\n{r['carrier']}  {r['query_type'] or ''} {r['query']}"
          f"   BL: {r['bill_of_lading'] or '-'}")
    print(f"{r['origin_city']}, {r['origin_country']}  ->  "
          f"{r['destination_city']}, {r['destination_country']}")
    print(f"POL {r['port_of_load']}  POD {r['port_of_discharge']}")
    if r["transshipments"]:
        print("via " + "  ->  ".join(r["transshipments"]))
    if r["final_pod_eta"]:
        print(f"final POD ETA: {r['final_pod_eta'][:10]}")

    for c in r["containers"]:
        print(f"\n  {c['container_id']}  {c['container_type']}"
              f"   latest: {c['latest_move']}"
              f"   {'DELIVERED' if c['delivered'] else 'in transit'}"
              + (f"   ETA {c['pod_eta'][:10]}" if c["pod_eta"] else ""))
        for e in c["events"]:
            marker = ">" if e["status"] == "CURRENT" else " "
            loc = f"{e['city']}, {e['country']}"
            ship = f"{e['vessel_name']} {e['voyage']}" if e["vessel_name"] else (e["laden_status"] or "")
            print(f"  {marker} {e['raw_date']}  {loc:<20} {(e['unlocode'] or ''):<6} "
                  f"{(e['mode'] or ''):<7} {e['description']:<32} {ship}")
    print()


def main() -> int:
    ap = argparse.ArgumentParser(description="Track an MSC container, BL, or booking.")
    ap.add_argument("number", help="container / bill of lading / booking number")
    ap.add_argument("--booking", action="store_true", help="treat as a booking number")
    ap.add_argument("--json", action="store_true", help="emit raw JSON")
    ap.add_argument("--raw", action="store_true", help="emit MSC's unparsed payload")
    ap.add_argument("-q", "--quiet", action="store_true", help="suppress diagnostics")
    args = ap.parse_args()

    mode = MODE_BOOKING if args.booking else MODE_CONTAINER_OR_BL
    number = args.number.strip().upper()

    try:
        data = fetch(number, mode=mode, verbose=not args.quiet)
    except MscError as exc:
        print(f"\nFAILED: {exc}", file=sys.stderr)
        return 1

    if args.raw:
        print(json.dumps(data, indent=2))
        return 0

    result = parse(data, query=number)
    print(json.dumps(result, indent=2) if args.json else "", end="")
    if not args.json:
        show(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
