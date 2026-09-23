#!/usr/bin/env python3
"""
Carrier registry and event-vocabulary normalisation.

TWO RULES THIS MODULE FOLLOWS
-----------------------------
1. The original carrier wording is never destroyed. Every normalised value is
   an ADDITIONAL field; `description` comes through untouched so the report
   can show both, and so an ops user can always see what the carrier actually
   said.

2. Every mapping below was read off the real event data in this project, not
   invented. EXACT_EVENT_MAP contains exactly the strings the two live
   adapters produce (9 Maersk, 12 MSC). Run

       python -m reporting.normalize --vocabulary events.json

   to re-derive that list from a fresh run and see whether anything new has
   appeared.

   KEYWORD_RULES is the deliberate exception. When a COSCO or Hapag adapter
   lands it will bring wording nobody here has seen, and the alternative to a
   keyword fallback is dropping those events into an OTHER bucket no report
   can use. The fallback is recorded per event in `mapping_source`
   ("exact" / "keyword" / "unmapped"), so the Status Mapping sheet shows
   exactly which rows were guessed at -- a guess you can see is fine, a guess
   you cannot see is not.

WHY time_basis EXISTS
---------------------
msc_track.parse() labels whatever event sorts last as CURRENT, without asking
what it is. MSC's last event is very often "Estimated Time of Arrival" or
"Full Intended Transshipment" -- forecasts, not moves. In one real run, 17 of
19 ETA events carried status=CURRENT. Trusting that makes `current_date` a
FUTURE date and turns "days since last move" into nonsense.

So the report derives `time_basis` (ACTUAL or EXPECTED) from three signals, in
order: the tracker's own PLANNED status, a normalised milestone that is
inherently a forecast, or an event dated after the snapshot. Everything the
report calls a "last actual move" is computed from ACTUAL rows only.

This is a reporting-layer correction. It does not modify the tracker, and it
does not overwrite the tracker's `status` field -- both are shown.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from datetime import date, datetime, timezone
from pathlib import Path

# --------------------------------------------------------------------------- #
# carriers
# --------------------------------------------------------------------------- #

# The eight carriers the reporting system is required to support. Order is the
# order they appear on the Carrier Mapping sheet. `adapter` records whether the
# tracking application can actually FETCH this carrier today -- the report says
# so plainly rather than implying coverage that does not exist.
TARGET_CARRIERS: list[dict] = [
    {"code": "MAEU", "name": "Maersk", "adapter": "browser (Playwright)"},
    {"code": "MSCU", "name": "MSC", "adapter": "HTTP API"},
    {"code": "COSU", "name": "COSCO", "adapter": None},
    {"code": "CMDU", "name": "CMA CGM", "adapter": None},
    {"code": "HLCU", "name": "Hapag-Lloyd", "adapter": None},
    {"code": "ONEY", "name": "ONE", "adapter": None},
    {"code": "EGLV", "name": "Evergreen", "adapter": None},
    {"code": "YMLU", "name": "Yang Ming", "adapter": None},
]

# Carriers outside the eight that carrier_id.py can already classify. Kept so a
# stray OOCL/HMM/ZIM row is labelled properly instead of showing a raw code.
OTHER_CARRIERS = {
    "OOLU": "OOCL",
    "HDMU": "HMM",
    "ZIMU": "ZIM",
}

CARRIER_NAMES: dict[str, str] = {c["code"]: c["name"] for c in TARGET_CARRIERS}
CARRIER_NAMES.update(OTHER_CARRIERS)

# Free-text carrier names -> code. Sources: the commercial sheet's SHIPPING
# LINE column, and errors.json `classified_as`. Matched on letters-only so
# "CMA-CGM", "CMA CGM" and "CMACGM" all land in the same bucket.
CARRIER_ALIASES: dict[str, str] = {
    "MAERSK": "MAEU", "MAERSKLINE": "MAEU", "SEALAND": "MAEU",
    "HAMBURGSUD": "MAEU",
    "MSC": "MSCU", "MEDITERRANEANSHIPPING": "MSCU",
    "COSCO": "COSU", "COSCOSHIPPING": "COSU", "COSCON": "COSU",
    "CMACGM": "CMDU", "CMA": "CMDU", "ANL": "CMDU", "APL": "CMDU",
    "HAPAG": "HLCU", "HAPAGLLOYD": "HLCU", "HAPAGLLOYDAG": "HLCU",
    "ONE": "ONEY", "OCEANNETWORKEXPRESS": "ONEY", "NYK": "ONEY",
    "MOL": "ONEY", "KLINE": "ONEY",
    "EVERGREEN": "EGLV", "EVERGREENLINE": "EGLV",
    "YANGMING": "YMLU", "YML": "YMLU", "YANGMINGLINE": "YMLU",
    "OOCL": "OOLU", "HMM": "HDMU", "HYUNDAI": "HDMU", "ZIM": "ZIMU",
}

# Longest alias first, so "HAPAGLLOYD" is preferred over "HAPAG" and
# "COSCOSHIPPING" over "COSCO" when the substring pass runs.
_ALIASES_BY_LENGTH = sorted(CARRIER_ALIASES.items(), key=lambda kv: -len(kv[0]))


def carrier_code(raw: str | None) -> str | None:
    """Free text or a code -> our carrier code. None when unrecognisable."""
    if not raw:
        return None
    text = str(raw).strip().upper()
    if text in CARRIER_NAMES:
        return text
    letters = re.sub(r"[^A-Z]", "", text)
    if not letters:
        return None
    if letters in CARRIER_ALIASES:
        return CARRIER_ALIASES[letters]
    # Substring pass, so "MAERSK (INCL. HAMBURG SUD)" and "SHIPPED VIA
    # EVERGREEN" still resolve. Minimum length 3 keeps it from matching noise.
    for alias, code in _ALIASES_BY_LENGTH:
        if len(alias) >= 3 and alias in letters:
            return code
    return None


def carrier_name(code: str | None, fallback: str | None = None) -> str:
    """Display name for a carrier code, never blank."""
    if code and str(code).strip().upper() in CARRIER_NAMES:
        return CARRIER_NAMES[str(code).strip().upper()]
    if fallback and str(fallback).strip():
        return str(fallback).strip()
    if code and str(code).strip():
        return str(code).strip()
    return "Unknown carrier"


# --------------------------------------------------------------------------- #
# event vocabulary
# --------------------------------------------------------------------------- #

# Normalised milestone -> (label, phase group, is_forecast).
#
# `phase` groups milestones into the stages an ops user thinks in. It is
# deliberately NOT a linear progress index: Maersk reports "Gate Out" both at
# origin (empty pickup) and at destination (final collection), so any single
# number claiming "% complete" from these events would be wrong.
MILESTONES: dict[str, tuple[str, str, bool]] = {
    "EMPTY_DISPATCHED": ("Empty released to shipper", "Origin", False),
    "GATE_IN": ("Gate in / received at yard", "Origin", False),
    "GATE_OUT": ("Gate out of yard", "Inland", False),
    "LOADED_ON_VESSEL": ("Loaded on vessel", "Ocean", False),
    "DEPARTURE": ("Vessel departure", "Ocean", False),
    "ARRIVAL": ("Vessel arrival", "Ocean", False),
    "DISCHARGED_FROM_VESSEL": ("Discharged from vessel", "Ocean", False),
    "TRANSSHIPMENT_LOADED": ("Transshipment loaded", "Transshipment", False),
    "TRANSSHIPMENT_DISCHARGED": ("Transshipment discharged", "Transshipment", False),
    "TRANSSHIPMENT_PLANNED": ("Transshipment planned", "Transshipment", True),
    "RAIL_DEPARTURE": ("Rail departure", "Inland", False),
    "RAIL_ARRIVAL": ("Rail arrival", "Inland", False),
    "TRUCK_MOVE": ("Road movement", "Inland", False),
    "CUSTOMS": ("Customs / documentation", "Customs", False),
    "DELIVERED": ("Delivered to consignee", "Destination", False),
    "EMPTY_RETURN": ("Empty returned", "Destination", False),
    "ETA_ADVICE": ("Carrier ETA advice", "Forecast", True),
    "OTHER": ("Other / unclassified", "Other", False),
}

# Exact matches on the carrier wording seen in this project's own event data.
# Left side is the string as it arrives in events.json `description`.
EXACT_EVENT_MAP: dict[str, str] = {
    # --- Maersk (maersk_track._humanise of the API `activity` field) ------- #
    "Gate Out": "GATE_OUT",
    "Gate In": "GATE_IN",
    "Container Departure": "DEPARTURE",
    "Container Arrival": "ARRIVAL",
    "Load": "LOADED_ON_VESSEL",
    # Maersk's source activity really is truncated to "DISCHARG"; kept verbatim
    # so it matches the live feed rather than a tidied-up version of it.
    "Discharg": "DISCHARGED_FROM_VESSEL",
    "Discharge": "DISCHARGED_FROM_VESSEL",
    "Rail Departure": "RAIL_DEPARTURE",
    "Rail Arrival At Destination": "RAIL_ARRIVAL",
    "Container Return": "EMPTY_RETURN",

    # --- MSC (Events[].Description, verbatim from the API) ---------------- #
    "Empty to Shipper": "EMPTY_DISPATCHED",
    "Export received at CY": "GATE_IN",
    "Export Loaded on Vessel": "LOADED_ON_VESSEL",
    "Full Intended Transshipment": "TRANSSHIPMENT_PLANNED",
    "Full Transshipment Loaded": "TRANSSHIPMENT_LOADED",
    "Full Transshipment Discharged": "TRANSSHIPMENT_DISCHARGED",
    "Import Discharged from Vessel": "DISCHARGED_FROM_VESSEL",
    "Import Loaded on Rail": "RAIL_DEPARTURE",
    "Import Unloaded from Rail": "RAIL_ARRIVAL",
    "Import to consignee": "DELIVERED",
    "Empty received at CY": "EMPTY_RETURN",
    "Estimated Time of Arrival": "ETA_ADVICE",
}

# Fallback for carriers with no adapter yet. First rule that matches wins, so
# order matters: specific phrases sit above generic single words.
KEYWORD_RULES: list[tuple[str, str]] = [
    (r"estimated.*arriv|^eta\b|\beta\b.*(update|revis)", "ETA_ADVICE"),
    (r"(intended|planned|expected).*transship", "TRANSSHIPMENT_PLANNED"),
    (r"transship\w*.*load|load\w*.*transship", "TRANSSHIPMENT_LOADED"),
    (r"transship\w*.*discharg|discharg\w*.*transship", "TRANSSHIPMENT_DISCHARGED"),
    (r"empty.*(return|received at|back to depot)", "EMPTY_RETURN"),
    (r"empty.*(shipper|pick|release|dispatch|out)", "EMPTY_DISPATCHED"),
    (r"deliver\w*|to consignee|consignee.*receiv", "DELIVERED"),
    (r"customs|clearance|documentation|manifest", "CUSTOMS"),
    (r"rail.*(arriv|unload|discharg)", "RAIL_ARRIVAL"),
    (r"rail.*(depart|load)", "RAIL_DEPARTURE"),
    (r"truck|road|barge|dray", "TRUCK_MOVE"),
    (r"discharg\w*|unload\w*.*vessel", "DISCHARGED_FROM_VESSEL"),
    # "On board vessel" (Yang Ming) and "Laden on board" (several lines) name no
    # loading action at all, so a load-word rule misses them entirely.
    (r"load\w*.*vessel|vessel.*load|\bloaded\b|on[\s-]?board", "LOADED_ON_VESSEL"),
    (r"depart\w*|sail\w*|left port", "DEPARTURE"),
    (r"arriv\w*|berth\w*", "ARRIVAL"),
    (r"gate[\s-]?out|departed.*terminal|out.*terminal", "GATE_OUT"),
    (r"gate[\s-]?in|received at|received.*(cy|yard|terminal|depot)", "GATE_IN"),
]

_COMPILED_RULES = [(re.compile(pat, re.I), code) for pat, code in KEYWORD_RULES]


def normalize_event(description: str | None) -> tuple[str, str]:
    """Carrier wording -> (milestone code, mapping_source).

    mapping_source is "exact" when the string is one this project has actually
    observed, "keyword" when a fallback rule fired, "unmapped" when nothing
    matched, and "blank" when the carrier sent nothing at all. The Status
    Mapping sheet reports these counts, so a growing keyword/unmapped
    population is visible rather than silent.
    """
    if description is None or not str(description).strip():
        return "OTHER", "blank"
    text = str(description).strip()
    if text in EXACT_EVENT_MAP:
        return EXACT_EVENT_MAP[text], "exact"
    for pattern, code in _COMPILED_RULES:
        if pattern.search(text):
            return code, "keyword"
    return "OTHER", "unmapped"


def milestone_label(code: str) -> str:
    return MILESTONES.get(code, MILESTONES["OTHER"])[0]


def milestone_phase(code: str) -> str:
    return MILESTONES.get(code, MILESTONES["OTHER"])[1]


def milestone_is_forecast(code: str) -> bool:
    return MILESTONES.get(code, MILESTONES["OTHER"])[2]


# --------------------------------------------------------------------------- #
# dates
# --------------------------------------------------------------------------- #

def parse_dt(raw) -> datetime | None:
    """ISO string -> naive datetime, or None. Never raises.

    Naive on purpose. The tracker writes snapshot_at with a UTC offset but
    event dates without one, and Excel has no concept of a timezone-aware
    datetime -- mixing the two produces "can't subtract offset-naive and
    offset-aware" at the first piece of date arithmetic. Everything is
    flattened to naive UTC here, once, so nothing downstream has to care.
    """
    if raw is None or raw == "":
        return None
    if isinstance(raw, datetime):
        return raw.astimezone(timezone.utc).replace(tzinfo=None) if raw.tzinfo else raw
    if isinstance(raw, date):
        return datetime(raw.year, raw.month, raw.day)
    text = str(raw).strip()
    if not text:
        return None
    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M", "%d/%m/%Y",
                    "%d-%m-%Y", "%d.%m.%Y"):
            try:
                dt = datetime.strptime(text, fmt)
                break
            except ValueError:
                continue
        else:
            return None
    if dt.tzinfo is not None:
        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
    return dt


def days_between(later: datetime | None, earlier: datetime | None) -> int | None:
    """Whole days from `earlier` to `later`, comparing dates not instants."""
    if later is None or earlier is None:
        return None
    return (later.date() - earlier.date()).days


def location_text(city, country) -> str | None:
    """('HALDIA', 'IN') -> 'HALDIA, IN'. Blank parts dropped, never ', IN'."""
    city = str(city).strip() if city not in (None, "") else ""
    country = str(country).strip() if country not in (None, "") else ""
    if city and country:
        return f"{city}, {country}"
    return city or country or None


# --------------------------------------------------------------------------- #
# event-level derivation
# --------------------------------------------------------------------------- #

FORECAST_STATUSES = {"PLANNED", "EXPECTED", "ESTIMATED", "FORECAST"}


def derive_event(row: dict, snapshot: datetime | None) -> dict:
    """One raw events.json row -> the row the report uses.

    Adds normalised fields alongside everything the tracker wrote. The raw
    `description` and `status` are carried through unchanged.
    """
    milestone, source = normalize_event(row.get("description"))
    event_dt = parse_dt(row.get("event_date"))
    tracker_status = str(row.get("status") or "").upper()

    # time_basis: is this a move that HAPPENED, or one the carrier expects?
    # Three independent signals -- see the module docstring for why the
    # tracker's own status is not sufficient on its own.
    if tracker_status in FORECAST_STATUSES:
        basis, reason = "EXPECTED", "tracker status PLANNED"
    elif milestone_is_forecast(milestone):
        basis, reason = "EXPECTED", "forecast milestone"
    elif snapshot is not None and event_dt is not None and event_dt.date() > snapshot.date():
        basis, reason = "EXPECTED", "dated after snapshot"
    else:
        basis, reason = "ACTUAL", "tracker status " + (tracker_status or "?")

    code = row.get("carrier")
    return {
        **row,
        "carrier_code": code,
        "carrier_display": carrier_name(code),
        "event_dt": event_dt,
        "milestone": milestone,
        "normalized_status": milestone_label(milestone),
        "phase": milestone_phase(milestone),
        "mapping_source": source,
        "time_basis": basis,
        "time_basis_reason": reason,
        "location": location_text(row.get("city"), row.get("country")),
    }


# --------------------------------------------------------------------------- #
# CLI -- re-derive the vocabulary from a fresh events.json
# --------------------------------------------------------------------------- #

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Show the event vocabulary in an events.json and how it maps.")
    ap.add_argument("--vocabulary", type=Path, default=Path("events.json"))
    args = ap.parse_args()

    rows = json.loads(args.vocabulary.read_text(encoding="utf-8"))
    seen: Counter = Counter()
    for r in rows:
        seen[(r.get("carrier"), r.get("description"))] += 1

    unmapped = 0
    print("\n%-10s %7s  %-34s %-9s %s"
          % ("CARRIER", "COUNT", "ORIGINAL", "SOURCE", "NORMALISED"))
    print("-" * 100)
    for (code, desc), n in sorted(seen.items(), key=lambda kv: (-kv[1], str(kv[0]))):
        milestone, source = normalize_event(desc)
        if source in ("unmapped", "blank"):
            unmapped += n
        print("%-10s %7d  %-34s %-9s %s"
              % (carrier_name(code), n, str(desc)[:34], source,
                 milestone_label(milestone)))
    print("\n%d distinct (carrier, description) pairs; %d event rows unmapped."
          % (len(seen), unmapped), file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
