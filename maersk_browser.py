#!/usr/bin/env python3
"""
Maersk tracker via a real browser. Free, and works where the direct API 403s.

    pip install playwright
    playwright install chromium

    python maersk_browser.py HASU1149124
    python maersk_browser.py HASU1149124 MRKU7621507 --json
    python maersk_browser.py --file maersk_list.txt --outdir ./out

WHY THIS WORKS WHEN requests DOESN'T
------------------------------------
api.maersk.com sits behind Akamai Bot Manager, which requires an
`akamai-bm-telemetry` header. That value is produced by obfuscated JavaScript
from mouse movement, screen metrics and timing. It cannot be forged, replayed
or reverse-engineered in any stable way -- which is why every requests/curl_cffi
attempt returns the same edge-level 403 even with no credentials at all.

A real browser just... runs the script. So instead of calling the API
ourselves, we open the normal tracking page, let Maersk's own JavaScript make
its own authenticated call, and read the JSON out of the response as it goes by.
We never touch the endpoint directly.

The parsing is unchanged -- it hands the captured payload to
maersk_track.parse(), so the output is identical to the direct-API path and
still matches msc_track.py field for field.

TRADE-OFFS, HONESTLY
--------------------
  * ~8-15 seconds per container instead of ~1. A browser has to start and a
    page has to render.
  * Chromium download is ~150 MB, one time.
  * Breaks if Maersk restructures their frontend (the API path did not change
    for months, but this is a heavier dependency).
  * headless=False by default. Headless Chrome is itself a bot signal, and
    since the whole point is looking like a real browser, we don't fight that.
    A window will open and close. Set HEADLESS=True below if you want to try.

Use this if the documented developer.maersk.com endpoint isn't available to
you. If it is, prefer it -- it's faster, lighter and supported.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import random
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path

from playwright.async_api import async_playwright

import maersk_track
from pipeline import flatten, write_json

PAGE = "https://www.maersk.com/tracking/{number}"

# A window will open. See the note above about headless being a bot signal.
HEADLESS = False

# Persisted browser profile: cookies and the Akamai _abck sensor survive between
# runs, so later lookups are less likely to be challenged.
PROFILE_DIR = Path("./.maersk-profile")

# The response we're waiting for.
API_MARKER = "/synergy/tracking/"


async def track_many(numbers: list[str], timeout_ms: int = 45_000,
                     verbose: bool = True) -> dict[str, dict]:
    """Open the tracking page per number; capture the JSON its own JS fetches."""
    results: dict[str, dict] = {}

    def log(msg: str) -> None:
        if verbose:
            print(f"# {msg}", file=sys.stderr)

    async with async_playwright() as pw:
        ctx = await pw.chromium.launch_persistent_context(
            str(PROFILE_DIR),
            headless=HEADLESS,
            viewport={"width": 1440, "height": 900},
            locale="en-US",
            args=["--disable-blink-features=AutomationControlled"],
        )

        page = await ctx.new_page()
        captured: dict[str, dict] = {}

        async def on_response(resp):
            if API_MARKER not in resp.url:
                return
            try:
                if "json" not in (resp.headers or {}).get("content-type", ""):
                    return
                data = await resp.json()
            except Exception:
                return
            if isinstance(data, dict) and data.get("containers"):
                for c in data["containers"]:
                    num = (c.get("container_num") or "").upper()
                    if num:
                        captured[num] = data

        page.on("response", on_response)

        for i, number in enumerate(numbers, 1):
            log(f"[{i}/{len(numbers)}] {number}")
            captured.clear()
            try:
                await page.goto(PAGE.format(number=number),
                                wait_until="domcontentloaded", timeout=timeout_ms)
                # Wait for the app's own XHR rather than a fixed sleep.
                for _ in range(int(timeout_ms / 500)):
                    if captured:
                        break
                    await page.wait_for_timeout(500)
            except Exception as exc:
                log(f"    navigation failed: {type(exc).__name__}: {exc}")

            if captured:
                payload = captured.get(number.upper()) or next(iter(captured.values()))
                results[number] = payload
                n = len(payload.get("containers") or [])
                log(f"    captured ({n} container(s))")
            else:
                log("    no tracking JSON seen -- number unknown, or the page "
                    "changed, or a challenge appeared (watch the window)")

            if i < len(numbers):
                await page.wait_for_timeout(int(random.uniform(2000, 4000)))

        await ctx.close()

    return results


def load_numbers(args) -> list[str]:
    nums = [n.strip().upper() for n in args.numbers if n.strip()]
    if args.file:
        for line in Path(args.file).read_text(encoding="utf-8-sig").splitlines():
            token = line.split(",")[0].strip().upper().replace(" ", "")
            if token and any(ch.isdigit() for ch in token):
                nums.append(token)
    seen, out = set(), []
    for n in nums:
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Track Maersk containers through a real browser.")
    ap.add_argument("numbers", nargs="*")
    ap.add_argument("--file", help="text file of numbers, one per line")
    ap.add_argument("--outdir", type=Path,
                    help="where containers.json / events.json go (default: here)")
    ap.add_argument("--rawdir", type=Path, help="also save untouched payloads")
    ap.add_argument("--append", action="store_true",
                    help="add to existing files instead of overwriting")
    ap.add_argument("--json", action="store_true", help="print parsed JSON")
    ap.add_argument("--show", action="store_true", help="print the full timeline")
    args = ap.parse_args()

    numbers = load_numbers(args)
    if not numbers:
        ap.print_help()
        return 2

    print(f"\nOpening a browser window for {len(numbers)} number(s). "
          f"Leave it alone while it works.\n", file=sys.stderr)

    raw = asyncio.run(track_many(numbers))
    if not raw:
        print("\nNothing captured. If a Maersk page loaded but no JSON appeared, "
              "watch the window on the next run -- a cookie banner or bot "
              "challenge may need one manual click. The profile is saved, so "
              "clicking it once usually fixes later runs.", file=sys.stderr)
        return 1

    outdir = args.outdir or Path(".")
    outdir.mkdir(parents=True, exist_ok=True)
    if args.rawdir:
        args.rawdir.mkdir(parents=True, exist_ok=True)

    run_id = uuid.uuid4().hex[:12]
    snapshot = datetime.now(timezone.utc).isoformat()
    all_c, all_e = [], []

    for number, payload in raw.items():
        if args.rawdir:
            (args.rawdir / f"{number}.json").write_text(
                json.dumps(payload, indent=2), encoding="utf-8")

        result = maersk_track.parse(payload, query=number)
        crows, erows = flatten(result, run_id, snapshot)
        all_c.extend(crows)
        all_e.extend(erows)

        if args.json:
            print(json.dumps(result, indent=2))
        elif args.show:
            maersk_track.show(result)
        else:
            c = crows[0]
            eta = c["pod_eta"][:10] if c["pod_eta"] else "-"
            print(f"#   {c['container_id']}  {c['current_event']} @ "
                  f"{c['current_city']}  ({c['event_count']} ev)  ETA {eta}",
                  file=sys.stderr)

    print(file=sys.stderr)
    if all_c:
        write_json(outdir / "containers.json", all_c, args.append)
        write_json(outdir / "events.json", all_e, args.append)
    print(f"# {len(raw)}/{len(numbers)} captured, run_id {run_id}", file=sys.stderr)

    missed = [n for n in numbers if n not in raw]
    if missed:
        print(f"# retry: python maersk_browser.py {' '.join(missed)}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())