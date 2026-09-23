#!/usr/bin/env python3
"""
Diagnose why the Maersk API isn't answering.

    python maersk_debug.py HASU1149124

Prints the raw status, headers and body for each variation it tries, so you can
see the actual cause instead of "NOT FOUND". Most likely causes, in order:

  401 / 403  -> the consumer-key in maersk_track.py has rotated. Maersk ships
                that key in their web bundle and changes it on redeploy.
                Fix: re-capture from DevTools, or register at
                developer.maersk.com and set MAERSK_CONSUMER_KEY.

  404        -> right key, wrong operator. Hamburg Sud boxes sometimes need
                operator=SUDU rather than MAEU. We try several.

  Timeout    -> corporate network or proxy blocking api.maersk.com. Check the
                URL loads in your browser on the same machine.
"""

from __future__ import annotations

import json
import os
import sys

import requests

try:
    from curl_cffi import requests as curl_requests
    HAS_CURL = True
except ImportError:
    HAS_CURL = False

URL = "https://api.maersk.com/synergy/tracking/{n}"
# ---------------------------------------------------------------------------
# YOUR KEY. Must match the one in maersk_track.py.
# ---------------------------------------------------------------------------
KEY = "jknZrwPsO463AW1SCXVcaLj7AgVnTfjM"

BASE_HEADERS = {
    "accept": "application/json",
    "accept-language": "en-US,en;q=0.9",
    "api-version": "v2",
    "consumer-key": KEY,
    "origin": "https://www.maersk.com",
    "referer": "https://www.maersk.com/",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "user-agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36"),
}


def attempt(label: str, number: str, params: dict,
            headers: dict, use_curl: bool) -> bool:
    print(f"\n--- {label} ---")
    print(f"    params : {params}")
    try:
        if use_curl and HAS_CURL:
            s = curl_requests.Session(impersonate="chrome")
            s.headers.update(headers)
            r = s.get(URL.format(n=number), params=params, timeout=30)
        else:
            s = requests.Session()
            s.headers.update(headers)
            r = s.get(URL.format(n=number), params=params, timeout=30)
    except Exception as exc:
        print(f"    EXCEPTION {type(exc).__name__}: {exc}")
        return False

    print(f"    status : {r.status_code}")
    for h in ("content-type", "x-akamai-reference", "www-authenticate",
              "retry-after", "x-ratelimit-remaining"):
        if h in r.headers:
            print(f"    {h}: {r.headers[h]}")

    body = r.text
    print(f"    bytes  : {len(body)}")
    if r.status_code == 200:
        try:
            data = r.json()
            n = len(data.get("containers") or [])
            print(f"    OK -- {n} container(s)")
            if n:
                c = data["containers"][0]
                print(f"    {c.get('container_num')}  {c.get('status')}  "
                      f"ETA {c.get('eta_final_delivery')}")
                print("\n    >>> THIS COMBINATION WORKS. Use these settings.")
                return True
            print("    200 but no containers -- number unknown to this operator")
        except ValueError:
            print(f"    200 but not JSON: {body[:200]!r}")
    else:
        print(f"    body   : {body[:400]!r}")
    return False


def main() -> int:
    number = (sys.argv[1] if len(sys.argv) > 1 else "HASU1149124").strip().upper()
    print(f"Diagnosing Maersk API for {number}")
    print(f"consumer-key: {KEY[:8]}...{KEY[-4:]}"
          + "  (hardcoded)")
    print(f"curl_cffi   : {'available' if HAS_CURL else 'NOT installed'}")

    # Operators: MAEU=Maersk, SUDU=Hamburg Sud, SEAU=Sealand, SEJJ, MCPU
    for op in ("MAEU", "SUDU", "SEAU"):
        if attempt(f"operator={op}", number, {"operator": op},
                   BASE_HEADERS, use_curl=False):
            return 0

    if attempt("no operator param", number, {}, BASE_HEADERS, use_curl=False):
        return 0

    if HAS_CURL and attempt("curl_cffi + operator=MAEU", number,
                            {"operator": "MAEU"}, BASE_HEADERS, use_curl=True):
        return 0

    no_key = {k: v for k, v in BASE_HEADERS.items() if k != "consumer-key"}
    attempt("without consumer-key (is the key even required?)",
            number, {"operator": "MAEU"}, no_key, use_curl=False)

    print("\n" + "=" * 60)
    print("Nothing worked. Read the statuses above:")
    print("  401/403 -> key rotated. Re-capture it:")
    print("             maersk.com/tracking/<number> > DevTools > Network >")
    print("             the api.maersk.com request > Headers > consumer-key")
    print("             then:  set MAERSK_CONSUMER_KEY=<value>")
    print("  404     -> number genuinely not on Maersk, or wrong operator")
    print("  timeout -> your network blocks api.maersk.com")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())