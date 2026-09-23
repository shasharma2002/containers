#!/usr/bin/env python3
"""
Turn the loaded dataset into the numbers the report shows.

WHAT COUNTS AS A DELAY HERE
---------------------------
Only two things, and both are things the carrier itself told us:

  1. ETA SLIP -- the carrier's own ETA for this container moved later between
     two snapshots of the appended history. That is the carrier revising its
     own promise, not an inference. Needs at least two snapshots carrying an
     ETA; containers tracked once have no slip and are not called delayed.

  2. ETA PASSED -- the carrier's latest ETA is in the past and the container
     still has no delivery or final-arrival event. The date came from the
     carrier; the absence of an arrival event came from the carrier.

Everything else that might look like lateness is reported as something else on
purpose:

  * No movement for N days is STALENESS, not delay. A box can legitimately sit
    on a vessel for two weeks. It is flagged because it needs a human to look,
    not because it is provably late.
  * A missing ETA is a DATA GAP. MSC supplies no forward schedule on most
    rows (see powerbi_setup.md section 9), so treating "no ETA" as "late"
    would mark most MSC shipments delayed for a reason that is the carrier's
    API design.

Every row on the Exceptions sheet therefore carries a `Class` of "Shipment" or
"Data", and the evidence column says which field the finding came from.

THREE DIFFERENT KINDS OF STALE
------------------------------
  idle_at_snapshot_days   how long the box had gone unscanned at the moment the
                          carrier was last queried. A fact about the CARGO, and
                          the only one the no-movement finding uses.
  snapshot_age_days       how long ago the TRACKER last looked. A fact about
                          our refresh cadence.
  days_since_last_move    the two added together -- what an ops user sees when
                          they ask "when did this last move?" today.

Keeping them apart is not cosmetic. Run the generator against a three-week-old
snapshot and days_since_last_move exceeds the critical threshold for every
container in the file, which would report the entire fleet as stalled when the
only real fact is that the pipeline has not run. Measured at the snapshot, the
same dataset flags 60 of 93 -- and the other 33 get a "tracking data not
refreshed" finding instead, which is the problem that actually needs fixing.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import datetime

from . import normalize as nz

# Thresholds. Collected here so they are tuned in one place, and every one of
# them appears in the report's own legend so a reader knows what produced a flag.
STALE_WARN_DAYS = 7
STALE_CRITICAL_DAYS = 14
SNAPSHOT_STALE_DAYS = 7
ETA_SLIP_WARN_DAYS = 3
ETA_SLIP_CRITICAL_DAYS = 7
RECENT_ACTIVITY_DAYS = 7
ACTIVITY_CHART_DAYS = 14
TOP_N = 10

# How far apart the boxes on one BL may last have moved before it means
# something. Containers on the same bill of lading are scanned individually, so
# a 23-container BL routinely shows a dozen different timestamps minutes apart
# -- that is the terminal working down a stack, not a problem. Whole days apart
# is different: it means part of the consignment did not travel with the rest.
CONTAINER_SPREAD_WARN_DAYS = 2
CONTAINER_SPREAD_CRITICAL_DAYS = 5

# Fields where ANY difference between containers on the same BL is real: these
# describe where a box physically is, so two different values mean two
# different places. `last_actual_dt` is deliberately absent -- see above.
DIVERGENCE_FIELDS = (
    ("status_bucket", "status"),
    ("last_actual_milestone", "confirmed milestone"),
    ("last_actual_location", "location"),
    ("vessel", "vessel"),
)

# ETA is NOT a strict field. Carriers issue a per-container delivery date, so a
# 12-box consignment routinely carries ETAs a day or two apart while every box
# is on the same vessel -- testing ETA for equality flagged 6 of 12 containers
# on BL 270383678 as outliers when nothing was wrong. It is measured as a
# spread instead, and only a spread of whole days means the consignment is
# genuinely being split across arrivals.
ETA_SPREAD_WARN_DAYS = 3

# How many container numbers to name inline before summarising the rest.
MAX_NAMED_CONTAINERS = 8

# Milestones that mean the box has reached the consignee. Used for "delivered"
# when the carrier's own `delivered` boolean is not set -- MSC sets it, Maersk
# only sets it once its container status reaches COMPLETED.
ARRIVED_MILESTONES = {"DELIVERED", "EMPTY_RETURN"}
# Milestones that mean the box has left origin and is on its way.
IN_TRANSIT_MILESTONES = {"LOADED_ON_VESSEL", "DEPARTURE", "ARRIVAL",
                         "DISCHARGED_FROM_VESSEL", "TRANSSHIPMENT_LOADED",
                         "TRANSSHIPMENT_DISCHARGED", "RAIL_DEPARTURE",
                         "RAIL_ARRIVAL", "GATE_OUT"}


# --------------------------------------------------------------------------- #
# container enrichment
# --------------------------------------------------------------------------- #

def enrich(data, as_of: datetime | None = None) -> datetime:
    """Attach per-container derived fields, in place. Returns the as-of time.

    `as_of` is the moment the report speaks for -- normally now. It is a
    parameter so a report can be regenerated for an earlier date and produce
    the same ageing numbers it would have produced then.
    """
    as_of = as_of or datetime.utcnow()

    # Keyed on (container, run), not container alone. In the normal latest-only
    # mode there is exactly one run per container so the two are equivalent --
    # but under --all-snapshots a container has one row per run, and keying on
    # the container alone would hand every one of those rows the whole event
    # history, inflating each row's event count by the number of runs.
    by_container: dict[tuple, list[dict]] = defaultdict(list)
    for event in data.events:
        by_container[(str(event.get("container_id")),
                      str(event.get("run_id")))].append(event)
    for events in by_container.values():
        # None-dated events sort last so they never win "latest".
        events.sort(key=lambda e: (e["event_dt"] is None,
                                   e["event_dt"] or datetime.min,
                                   e.get("sequence") or 0))

    for container in data.containers:
        cid = str(container.get("container_id"))
        events = by_container.get((cid, str(container.get("run_id"))), [])
        actual = [e for e in events if e["time_basis"] == "ACTUAL" and e["event_dt"]]
        expected = [e for e in events if e["time_basis"] == "EXPECTED" and e["event_dt"]]
        last = actual[-1] if actual else None
        first = actual[0] if actual else None
        # Next expected move: the earliest forecast event still in the future.
        upcoming = [e for e in expected if e["event_dt"].date() >= as_of.date()]
        nxt = upcoming[0] if upcoming else (expected[-1] if expected else None)

        history = data.eta_history.get(cid, {})
        eta = (history.get("latest_eta") or container["pod_eta_dt"]
               or container["final_pod_eta_dt"])

        container["events_on_report"] = len(events)
        container["actual_events"] = len(actual)
        container["expected_events"] = len(expected)

        container["last_actual_event"] = last.get("description") if last else None
        container["last_actual_normalized"] = last.get("normalized_status") if last else None
        container["last_actual_milestone"] = last.get("milestone") if last else None
        container["last_actual_dt"] = last["event_dt"] if last else None
        container["last_actual_location"] = last.get("location") if last else None
        container["last_actual_vessel"] = last.get("vessel_name") if last else None
        container["last_actual_voyage"] = last.get("voyage") if last else None
        container["last_actual_mode"] = last.get("mode") if last else None
        container["first_actual_dt"] = first["event_dt"] if first else None

        # "What vessel is it on?" is not answered by the latest event -- 42 of
        # 116 containers in the sample last moved through a gate or rail event,
        # which carries no vessel. The answer the ops team wants is the most
        # recent confirmed event that DID name a vessel.
        carried = [e for e in actual if e.get("vessel_name")]
        vessel_event = carried[-1] if carried else None
        container["vessel"] = vessel_event.get("vessel_name") if vessel_event else None
        container["voyage"] = vessel_event.get("voyage") if vessel_event else None
        container["vessel_event_dt"] = vessel_event["event_dt"] if vessel_event else None
        container["vessel_event"] = vessel_event.get("description") if vessel_event else None

        container["next_event"] = nxt.get("description") if nxt else None
        container["next_event_normalized"] = nxt.get("normalized_status") if nxt else None
        container["next_event_dt"] = nxt["event_dt"] if nxt else None
        container["next_event_location"] = nxt.get("location") if nxt else None

        container["days_since_last_move"] = nz.days_between(as_of, container["last_actual_dt"])
        container["snapshot_age_days"] = nz.days_between(as_of, container["snapshot_dt"])
        # How long the box had been idle AT THE MOMENT WE LAST LOOKED. This is
        # the number the no-movement finding uses, and the distinction is not
        # cosmetic: on a three-week-old snapshot, days_since_last_move is at
        # least three weeks for every container, which would flag the whole
        # fleet critical when the only real fact is that the tracker has not
        # run. Idleness measured at the snapshot is a fact about the cargo;
        # the gap since then is a fact about our refresh cadence, and the two
        # get separate findings.
        container["idle_at_snapshot_days"] = nz.days_between(
            container["snapshot_dt"], container["last_actual_dt"])
        container["transit_days_so_far"] = nz.days_between(
            container["last_actual_dt"], container["first_actual_dt"])

        container["eta_dt"] = eta
        # Date, not datetime. Two containers on one BL whose ETAs differ by
        # hours are on the same ETA; comparing datetimes would call that a
        # divergence.
        container["eta_date"] = eta.date() if eta else None
        container["first_eta_dt"] = history.get("first_eta")
        container["eta_slip_days"] = history.get("eta_slip_days")
        container["eta_slip_since_previous"] = history.get("eta_slip_since_previous")
        container["eta_observations"] = history.get("eta_observations", 0)
        container["snapshots_seen"] = history.get("snapshots", 1)
        container["days_to_eta"] = nz.days_between(eta, as_of)

        container["status_bucket"] = _status_bucket(container, actual)
        # The tracker's `current_event` can be a forecast (MSC labels its ETA
        # row CURRENT). Say so rather than presenting it as a confirmed move.
        current_milestone, _ = nz.normalize_event(container.get("current_event"))
        container["current_is_forecast"] = nz.milestone_is_forecast(current_milestone)
        container["current_normalized"] = nz.milestone_label(current_milestone)

    _build_shipments(data, as_of)
    _attach_exceptions(data, as_of)
    return as_of


# --------------------------------------------------------------------------- #
# bill of lading -- the grain the business actually works in
# --------------------------------------------------------------------------- #

def _build_shipments(data, as_of: datetime) -> None:
    """Group containers into one row per bill of lading. Sets data.shipment_rows.

    WHY THIS IS THE PRIMARY GRAIN
    -----------------------------
    Cargo is booked, tracked and chased by BL, not by box. In this project's
    data 116 containers are 29 BLs, and one BL carries up to 23 containers that
    travel as a single consignment -- same vessel, same voyage, same ETA,
    scanned minutes apart at the same terminal. Reporting those 23 identically
    moving boxes as 23 lines buries the one thing worth reading: the BL.

    So each BL gets one row showing the CONSENSUS position (the value most of
    its containers share), plus an explicit count of any container that does
    not match. The exception is the interesting case -- a box left behind,
    damaged, or rolled to a later vessel -- and it is named rather than
    averaged away.

    Consensus is the modal value, not the latest or the first. On BL 271560680,
    12 containers are delivered at Birganj and 1 is still in transit at
    Visakhapatnam: the shipment's honest position is "delivered", with one
    outlier called out by number. Taking the latest or the worst would describe
    the whole consignment by its straggler.
    """
    groups: dict[str, list[dict]] = defaultdict(list)
    for container in data.containers:
        key = container.get("tracking_number_resolved") or "(no tracking number)"
        groups[key].append(container)

    commercial = _commercial_index(data)
    rows: list[dict] = []

    for bl, containers in groups.items():
        for container in containers:
            container["bl"] = bl

        consensus = {}
        divergence = []
        for field, label in DIVERGENCE_FIELDS:
            values = Counter(c.get(field) for c in containers)
            winner, count = values.most_common(1)[0]
            consensus[field] = winner
            consensus[field + "_agreeing"] = count
            if len(values) > 1:
                divergence.append("%s (%d different)" % (label, len(values)))

        # Containers that differ from the consensus on any meaningful field.
        outliers = [c for c in containers
                    if any(c.get(field) != consensus[field]
                           for field, _ in DIVERGENCE_FIELDS)]

        moves = [c["last_actual_dt"] for c in containers if c["last_actual_dt"]]
        spread = nz.days_between(max(moves), min(moves)) if moves else None

        etas = [c["eta_dt"] for c in containers if c["eta_dt"]]
        eta_spread = nz.days_between(max(etas), min(etas)) if etas else None
        # Deliberately NOT appended to `divergence`: that column backs the
        # Split? flag, which promises a nameable outlier container. A staggered
        # ETA has no culprit box -- it gets its own column and its own finding.
        # ETA is not a strict divergence field, so its consensus is taken here
        # rather than coming out of the loop above.
        eta_dates = Counter(c["eta_date"] for c in containers if c["eta_date"])
        consensus_eta = eta_dates.most_common(1)[0][0] if eta_dates else None
        slips = [c["eta_slip_days"] for c in containers
                 if c["eta_slip_days"] is not None]
        idles = [c["idle_at_snapshot_days"] for c in containers
                 if c["idle_at_snapshot_days"] is not None]
        snapshots = [c["snapshot_dt"] for c in containers if c["snapshot_dt"]]
        buckets = Counter(c["status_bucket"] for c in containers)

        leader = _leader(containers, consensus)
        sheet = commercial.get(bl, {})

        rows.append({
            "bl": bl,
            "carrier": leader["carrier_display"],
            "carrier_code": leader["carrier_code"],
            "containers": len(containers),
            "container_list": _name_containers([c["container_id"] for c in containers]),
            "single_container": len(containers) == 1,

            "status": consensus["status_bucket"],
            "status_agreeing": consensus["status_bucket_agreeing"],
            "in_transit": buckets.get("In Transit", 0),
            "delivered_count": buckets.get("Delivered", 0),
            "at_origin": buckets.get("At Origin", 0),

            "confirmed_status": leader["last_actual_normalized"],
            "carrier_wording": leader["last_actual_event"],
            "location": consensus["last_actual_location"],
            "vessel": consensus["vessel"],
            "voyage": leader["voyage"],

            "latest_move": max(moves) if moves else None,
            "earliest_move": min(moves) if moves else None,
            "move_spread_days": spread,
            "idle_days": min(idles) if idles else None,

            "eta": consensus_eta,
            "eta_spread_days": eta_spread,
            "eta_earliest": min(etas).date() if etas else None,
            "eta_latest": max(etas).date() if etas else None,
            "eta_slip_days": max(slips) if slips else None,
            "days_to_eta": nz.days_between(max(etas), as_of) if etas else None,

            "next_event": leader["next_event_normalized"],
            "next_event_dt": leader["next_event_dt"],
            "next_event_location": leader["next_event_location"],

            "origin": leader.get("origin"),
            "destination": leader.get("destination"),
            "port_of_load": leader.get("port_of_load"),
            "port_of_discharge": leader.get("port_of_discharge"),
            "transshipments": leader.get("transshipments"),

            "split": "Yes" if outliers else "No",
            "outlier_count": len(outliers),
            "outliers": _name_containers([c["container_id"] for c in outliers]),
            "divergence": "; ".join(divergence) or None,

            "events": sum(c["events_on_report"] for c in containers),
            "snapshot_dt": max(snapshots) if snapshots else None,
            "snapshot_age_days": (nz.days_between(as_of, max(snapshots))
                                  if snapshots else None),

            # Joined from the commercial sheet when that BL is on it. Blank is
            # normal -- a container-number lookup has no sheet row behind it.
            "beneficiary": sheet.get("beneficiary"),
            "commodity": sheet.get("commodity"),
            "broker": sheet.get("broker"),
            "sheet_eta": sheet.get("sheet_eta"),
            "eta_variance_days": sheet.get("eta_variance_days"),
            "amount": sheet.get("amount"),
        })

    rows.sort(key=lambda r: (-r["containers"], str(r["bl"])))
    data.shipment_rows = rows


def _leader(containers: list[dict], consensus: dict) -> dict:
    """The container whose row best represents the BL.

    The first one matching the consensus status, so the descriptive fields that
    are not worth a divergence test of their own -- voyage, origin, next
    expected event -- are taken from a container that is where most of the
    consignment is, rather than from whichever happened to sort first.
    """
    for container in containers:
        if container["status_bucket"] == consensus["status_bucket"]:
            return container
    return containers[0]


def _name_containers(ids: list[str]) -> str | None:
    """'A, B, C' -- or 'A, B, ... +15 more' once a cell would stop being readable."""
    clean = [str(i) for i in ids if i]
    if not clean:
        return None
    clean.sort()
    if len(clean) <= MAX_NAMED_CONTAINERS:
        return ", ".join(clean)
    return "%s, +%d more" % (", ".join(clean[:MAX_NAMED_CONTAINERS]),
                             len(clean) - MAX_NAMED_CONTAINERS)


def _commercial_index(data) -> dict:
    """BL -> the commercial sheet facts for it.

    One BL can occupy several sheet rows (a split shipment with two commodity
    lines), so distinct values are joined rather than one row winning. Amounts
    are summed, since two lines of the same BL are two parts of one value.
    """
    index: dict[str, dict] = {}
    for shipment in data.shipments:
        bl = shipment.get("bl_number")
        if not bl:
            continue
        entry = index.setdefault(bl, {"beneficiary": set(), "commodity": set(),
                                      "broker": set(), "sheet_eta": None,
                                      "eta_variance_days": None, "amount": 0.0})
        for field in ("beneficiary", "commodity", "broker"):
            if shipment.get(field):
                entry[field].add(str(shipment[field]))
        if shipment.get("sheet_eta_dt") and entry["sheet_eta"] is None:
            entry["sheet_eta"] = shipment["sheet_eta_dt"].date()
        if shipment.get("eta_variance_days") is not None:
            entry["eta_variance_days"] = shipment["eta_variance_days"]
        if isinstance(shipment.get("amount"), (int, float)):
            entry["amount"] += float(shipment["amount"])

    return {bl: {**entry,
                 "beneficiary": "; ".join(sorted(entry["beneficiary"])) or None,
                 "commodity": "; ".join(sorted(entry["commodity"])) or None,
                 "broker": "; ".join(sorted(entry["broker"])) or None,
                 "amount": entry["amount"] or None}
            for bl, entry in index.items()}


def _status_bucket(container: dict, actual: list[dict]) -> str:
    """Delivered / In Transit / At Origin / No Events -- the operational state.

    The carrier's own `delivered` boolean is trusted first. Beyond that the
    bucket is decided by which milestones have actually happened, so it means
    the same thing for a carrier that reports a status field and one that does
    not.
    """
    if container.get("delivered"):
        return "Delivered"
    milestones = {e["milestone"] for e in actual}
    if milestones & ARRIVED_MILESTONES:
        return "Delivered"
    if not actual:
        return "No Events"
    if milestones & IN_TRANSIT_MILESTONES:
        return "In Transit"
    return "At Origin"


# --------------------------------------------------------------------------- #
# exceptions
# --------------------------------------------------------------------------- #

def _attach_exceptions(data, as_of: datetime) -> None:
    """Build data.exceptions and mark each container's attention level.

    Findings are detected per container, then ROLLED UP to the bill of lading.
    That rollup is the difference between a usable sheet and an unusable one:
    a 23-container BL whose carrier moved the ETA produces one finding saying
    "all 23 containers", not 23 identical lines. When only some containers are
    affected the row says "1 of 13 containers" and names them, which is exactly
    the case worth reading -- the box that did not travel with the rest.
    """
    findings: list[dict] = []

    for container in data.containers:
        cid = container.get("container_id")
        carrier = container["carrier_display"]
        tracking = container.get("tracking_number_resolved")
        delivered = container["status_bucket"] == "Delivered"

        def add(cls, severity, category, detail, evidence, value=None):
            findings.append({
                "container_id": cid,
                "bl": container.get("bl") or tracking,
                "carrier": carrier,
                "tracking_number": tracking,
                "class": cls,
                "severity": severity,
                "category": category,
                "detail": detail,
                "evidence": evidence,
                "value": value,
                "status": container["status_bucket"],
                "last_move": container["last_actual_dt"],
                "location": container["last_actual_location"],
                "eta": container["eta_dt"],
                "destination": container.get("destination"),
            })

        # ---- shipment findings -------------------------------------------- #
        slip = container["eta_slip_days"]
        if slip is not None and slip >= ETA_SLIP_WARN_DAYS and not delivered:
            severity = "Critical" if slip >= ETA_SLIP_CRITICAL_DAYS else "Warning"
            add("Shipment", severity, "ETA slipped",
                "Carrier moved its own ETA %d day(s) later, from %s to %s."
                % (slip, _d(container["first_eta_dt"]), _d(container["eta_dt"])),
                "pod_eta across %d snapshots" % container["snapshots_seen"], slip)

        eta = container["eta_dt"]
        if eta is not None and not delivered and eta.date() < as_of.date():
            overdue = nz.days_between(as_of, eta)
            add("Shipment", "Critical" if overdue >= ETA_SLIP_CRITICAL_DAYS else "Warning",
                "ETA passed, not delivered",
                "Carrier ETA was %s (%d day(s) ago) and no delivery or final "
                "arrival event has been reported." % (_d(eta), overdue),
                "pod_eta vs event history", overdue)

        idle = container["idle_at_snapshot_days"]
        if idle is not None and idle >= STALE_WARN_DAYS and not delivered:
            add("Shipment", "Critical" if idle >= STALE_CRITICAL_DAYS else "Warning",
                "No recent movement",
                "When the carrier was last checked (%s) this container had not "
                "moved for %d day(s) -- last seen %s at %s. Staleness needs a "
                "look; it is not proof of a delay."
                % (_d(container["snapshot_dt"]), idle,
                   _d(container["last_actual_dt"]),
                   container["last_actual_location"] or "an unreported location"),
                "latest ACTUAL event vs snapshot_at", idle)

        # ---- data findings ------------------------------------------------- #
        if container["actual_events"] == 0:
            add("Data", "Critical", "No confirmed events",
                "The carrier returned no actual movement for this container -- "
                "only forecasts, or nothing at all.",
                "events.json", container["events_on_report"])

        if eta is None and not delivered:
            add("Data", "Warning", "No ETA from carrier",
                "Neither pod_eta nor final_pod_eta is populated. For MSC this is "
                "normal -- that API supplies no forward schedule. Use last "
                "movement as the progress signal instead.",
                "pod_eta / final_pod_eta")

        if container.get("current_is_forecast"):
            add("Data", "Warning", "Latest status is a forecast",
                "The tracker's current_event is '%s', which is a forecast rather "
                "than a confirmed move. The confirmed position is '%s' on %s."
                % (container.get("current_event"),
                   container["last_actual_normalized"] or "nothing reported",
                   _d(container["last_actual_dt"])),
                "current_event vs time_basis")

        age = container["snapshot_age_days"]
        if age is not None and age >= SNAPSHOT_STALE_DAYS:
            add("Data", "Warning", "Tracking data not refreshed",
                "This container was last looked up %d day(s) ago (%s). Its status "
                "may have changed since. Re-run the tracker."
                % (age, _d(container["snapshot_dt"])), "snapshot_at", age)

        if container.get("destination") is None:
            add("Data", "Info", "No destination",
                "The carrier returned no destination city/country, so no "
                "arrival expectation can be formed.", "destination_city")

        if container.get("carrier_code") is None:
            add("Data", "Critical", "No carrier",
                "This container row carries no carrier code, so it cannot be "
                "attributed to a shipping line.", "carrier")

        if container["snapshot_dt"] is None:
            add("Data", "Warning", "Unparseable snapshot date",
                "snapshot_at could not be read as a date (%r), so this row's age "
                "is unknown." % container.get("snapshot_at"), "snapshot_at")

    # ---- BL-level findings ------------------------------------------------ #
    # These only exist at the consignment grain: they are statements about how
    # the containers on one bill of lading relate to EACH OTHER, which no
    # single container row can express.
    for shipment in data.shipment_rows:
        if shipment["single_container"]:
            continue

        if shipment["outlier_count"]:
            severity = ("Critical" if shipment["outlier_count"] * 2 <= shipment["containers"]
                        else "Warning")
            findings.append({
                "container_id": None,
                "bl": shipment["bl"],
                "carrier": shipment["carrier"],
                "tracking_number": shipment["bl"],
                "class": "Shipment",
                "severity": severity,
                "category": "Containers not travelling together",
                "detail": "%d of %d container(s) on this BL differ from the rest "
                          "on %s. The consignment's main body is '%s' at %s; the "
                          "odd container(s) are %s. This is the case worth "
                          "checking by hand -- a box damaged, rolled to a later "
                          "vessel, or held."
                          % (shipment["outlier_count"], shipment["containers"],
                             shipment["divergence"] or "at least one field",
                             shipment["confirmed_status"] or "unreported",
                             shipment["location"] or "an unreported location",
                             shipment["outliers"]),
                "evidence": "per-container status/milestone/location/vessel/ETA",
                "value": shipment["outlier_count"],
                "status": shipment["status"],
                "last_move": shipment["latest_move"],
                "location": shipment["location"],
                "eta": shipment["eta"],
                "destination": shipment["destination"],
                "_bl_level": True,
                "_affected": shipment["outliers"],
                "_scope": "%d of %d containers" % (shipment["outlier_count"],
                                                   shipment["containers"]),
            })

        eta_spread = shipment["eta_spread_days"]
        if (eta_spread is not None and eta_spread >= ETA_SPREAD_WARN_DAYS
                and shipment["delivered_count"] < shipment["containers"]):
            findings.append({
                "container_id": None,
                "bl": shipment["bl"],
                "carrier": shipment["carrier"],
                "tracking_number": shipment["bl"],
                "class": "Shipment",
                "severity": ("Critical" if eta_spread >= ETA_SPREAD_WARN_DAYS * 3
                             else "Warning"),
                "category": "Arrival dates spread across the consignment",
                "detail": "The carrier expects the %d containers on this BL to "
                          "arrive over %d day(s) (%s to %s). Plan delivery and "
                          "free-time in waves rather than as one arrival. A day "
                          "or two apart is normal and is not reported."
                          % (shipment["containers"], eta_spread,
                             shipment["eta_earliest"], shipment["eta_latest"]),
                "evidence": "spread of pod_eta across containers",
                "value": eta_spread,
                "status": shipment["status"],
                "last_move": shipment["latest_move"],
                "location": shipment["location"],
                "eta": shipment["eta"],
                "destination": shipment["destination"],
                "_bl_level": True,
                "_affected": None,
                "_scope": "Consignment-wide",
            })

        spread = shipment["move_spread_days"]
        fully_delivered = shipment["delivered_count"] == shipment["containers"]
        if (spread is not None and spread >= CONTAINER_SPREAD_WARN_DAYS
                and not fully_delivered):
            findings.append({
                "container_id": None,
                "bl": shipment["bl"],
                "carrier": shipment["carrier"],
                "tracking_number": shipment["bl"],
                "class": "Shipment",
                "severity": ("Critical" if spread >= CONTAINER_SPREAD_CRITICAL_DAYS
                             else "Warning"),
                "category": "Containers separated in time",
                "detail": "The %d containers on this BL last moved %d day(s) "
                          "apart (%s to %s). Minutes apart is normal -- a "
                          "terminal working down a stack. Days apart means part "
                          "of the consignment did not travel with the rest."
                          % (shipment["containers"], spread,
                             _d(shipment["earliest_move"]),
                             _d(shipment["latest_move"])),
                "evidence": "spread of latest ACTUAL event across containers",
                "value": spread,
                "status": shipment["status"],
                "last_move": shipment["latest_move"],
                "location": shipment["location"],
                "eta": shipment["eta"],
                "destination": shipment["destination"],
                "_bl_level": True,
                "_affected": None,
                "_scope": "Consignment-wide",
            })

    # ---- dataset-level failures, carried onto the same sheet -------------- #
    for error in data.errors:
        if error["failure_kind"] == "No ocean carrier":
            severity, cls = "Info", "Data"
        elif error["failure_kind"] == "Unsupported carrier":
            severity, cls = "Warning", "Data"
        else:
            severity, cls = "Critical", "Data"
        findings.append({
            "container_id": None,
            "bl": error.get("query"),
            "carrier": error["carrier_display"],
            "tracking_number": error.get("query"),
            "class": cls,
            "severity": severity,
            "category": "Tracking failure: " + error["failure_kind"],
            "detail": error["suggested_action"],
            "evidence": str(error.get("error") or "")[:200],
            "value": None,
            "status": "Not tracked",
            "last_move": None,
            "location": None,
            "eta": None,
            "destination": None,
            "_bl_level": True,
            "_affected": None,
            "_scope": "Never tracked",
        })

    severity_rank = {"Critical": 0, "Warning": 1, "Info": 2}

    # ---- per-container attention, from the UNROLLED findings -------------- #
    # Done before the rollup, because the Container Summary sheet still needs
    # to say which individual box is flagged.
    worst: dict[str, str] = {}
    for finding in findings:
        cid = finding.get("container_id")
        if cid is None:
            continue
        if (cid not in worst
                or severity_rank.get(finding["severity"], 3) < severity_rank.get(worst[cid], 3)):
            worst[cid] = finding["severity"]
    counts = Counter(f["container_id"] for f in findings if f.get("container_id"))
    shipment_counts = Counter(f["container_id"] for f in findings
                              if f.get("container_id") and f["class"] == "Shipment")
    for container in data.containers:
        cid = container.get("container_id")
        container["attention"] = worst.get(cid, "OK")
        container["issue_count"] = counts.get(cid, 0)
        container["shipment_issue_count"] = shipment_counts.get(cid, 0)

    data.exceptions = _rollup(findings, data, severity_rank)
    _attach_shipment_attention(data, severity_rank)


def _rollup(findings: list[dict], data, severity_rank: dict) -> list[dict]:
    """Collapse per-container findings to one row per (BL, class, category).

    The representative row is the most severe one, and among equals the one
    with the largest value -- so "ETA slipped" on a BL reports the worst slip
    of its containers, not an arbitrary one. Every affected container is still
    named in the Containers column, so nothing is lost by collapsing.
    """
    sizes = {s["bl"]: s["containers"] for s in data.shipment_rows}

    groups: dict[tuple, list[dict]] = defaultdict(list)
    passthrough: list[dict] = []
    for finding in findings:
        if finding.get("_bl_level"):
            passthrough.append(finding)
            continue
        groups[(finding["bl"], finding["class"], finding["category"])].append(finding)

    rolled: list[dict] = []
    for (bl, _cls, _category), members in groups.items():
        total = sizes.get(bl, len(members))
        affected = sorted({m["container_id"] for m in members if m["container_id"]})
        lead = min(members, key=lambda f: (severity_rank.get(f["severity"], 3),
                                           -(f["value"] if isinstance(f["value"], (int, float))
                                             else 0)))
        if total <= 1:
            scope = "Single container"
        elif len(affected) >= total:
            scope = "All %d containers" % total
        else:
            scope = "%d of %d containers" % (len(affected), total)

        rolled.append({**lead,
                       "container_id": affected[0] if len(affected) == 1 else None,
                       "scope": scope,
                       "containers": _name_containers(affected),
                       "affected_count": len(affected)})

    for finding in passthrough:
        rolled.append({**finding,
                       "scope": finding.get("_scope"),
                       "containers": finding.get("_affected"),
                       "affected_count": finding.get("value")
                       if isinstance(finding.get("value"), int) else None})

    for row in rolled:
        row.pop("_bl_level", None)
        row.pop("_affected", None)
        row.pop("_scope", None)

    rolled.sort(key=lambda f: (severity_rank.get(f["severity"], 3),
                               f["class"], f["category"], str(f["bl"] or "")))
    return rolled


def _attach_shipment_attention(data, severity_rank: dict) -> None:
    """Mark each BL row with its worst finding and how many it has."""
    worst: dict[str, str] = {}
    counts: Counter = Counter()
    shipment_counts: Counter = Counter()
    for finding in data.exceptions:
        bl = finding.get("bl")
        if bl is None:
            continue
        counts[bl] += 1
        if finding["class"] == "Shipment":
            shipment_counts[bl] += 1
        if (bl not in worst
                or severity_rank.get(finding["severity"], 3) < severity_rank.get(worst[bl], 3)):
            worst[bl] = finding["severity"]

    for shipment in data.shipment_rows:
        bl = shipment["bl"]
        shipment["attention"] = worst.get(bl, "OK")
        shipment["issue_count"] = counts.get(bl, 0)
        shipment["shipment_issue_count"] = shipment_counts.get(bl, 0)


def _d(value: datetime | None) -> str:
    return value.strftime("%Y-%m-%d") if value else "no date"


# --------------------------------------------------------------------------- #
# carrier rollup
# --------------------------------------------------------------------------- #

def carrier_summary(data) -> list[dict]:
    """One row per carrier. Carriers are derived from the data, never hard-coded.

    The eight target carriers are then unioned in, so a carrier with no rows
    shows as "no data in this run" instead of being invisible -- the report has
    to be able to say "COSCO: nothing here" rather than silently omitting it.
    """
    codes: dict[str, str] = {}
    for container in data.containers:
        codes[container["carrier_code"] or "?"] = container["carrier_display"]
    for error in data.errors:
        codes.setdefault(error["carrier_code"] or "?", error["carrier_display"])
    for shipment in data.shipments:
        codes.setdefault(shipment["carrier_code"] or "?", shipment["carrier_display"])
    for target in nz.TARGET_CARRIERS:
        codes.setdefault(target["code"], target["name"])

    adapters = {t["code"]: t["adapter"] for t in nz.TARGET_CARRIERS}
    targets = {t["code"] for t in nz.TARGET_CARRIERS}

    rows = []
    for code, name in codes.items():
        containers = [c for c in data.containers if (c["carrier_code"] or "?") == code]
        events = [e for e in data.events if (e["carrier_code"] or "?") == code]
        failures = [e for e in data.errors if (e["carrier_code"] or "?") == code]
        buckets = Counter(c["status_bucket"] for c in containers)
        slips = [c["eta_slip_days"] for c in containers
                 if c["eta_slip_days"] is not None and c["eta_slip_days"] > 0]
        idles = [c["idle_at_snapshot_days"] for c in containers
                 if c["idle_at_snapshot_days"] is not None]

        rows.append({
            "carrier": name,
            "code": code,
            "in_target_eight": "Yes" if code in targets else "No",
            "adapter": adapters.get(code) or "none",
            "containers": len(containers),
            "in_transit": buckets.get("In Transit", 0),
            "at_origin": buckets.get("At Origin", 0),
            "delivered": buckets.get("Delivered", 0),
            "no_events": buckets.get("No Events", 0),
            "delayed": len(slips),
            "avg_eta_slip": round(sum(slips) / len(slips), 1) if slips else None,
            "stale": sum(1 for c in containers
                         if (c["idle_at_snapshot_days"] or 0) >= STALE_WARN_DAYS
                         and c["status_bucket"] != "Delivered"),
            "missing_eta": sum(1 for c in containers if c["eta_dt"] is None),
            "with_forecast": sum(1 for c in containers if c.get("has_forecast")),
            "needs_attention": sum(1 for c in containers if c["attention"] != "OK"),
            "events": len(events),
            "avg_events": round(len(events) / len(containers), 1) if containers else None,
            "avg_idle_days": round(sum(idles) / len(idles), 1) if idles else None,
            "failures": len(failures),
            "coverage": _coverage(code, containers, failures, adapters, targets),
        })

    rows.sort(key=lambda r: (-r["containers"], r["carrier"]))
    return rows


def _coverage(code, containers, failures, adapters, targets) -> str:
    """Why this carrier has the data it has -- the reader's first question.

    The failure kinds are inspected rather than just counted, because "we tried
    and got nothing" and "there is nothing to try" are different answers. The
    'Inland/Truck' pseudo-carrier from errors.json is the second: inland cargo
    never gets a bill of lading, so no lookup was attempted or possible.
    """
    if containers:
        return "Tracked"
    kinds = {f["failure_kind"] for f in failures}
    if kinds and kinds <= {"No ocean carrier"}:
        return "Not an ocean carrier -- nothing to track"
    if kinds == {"Unsupported carrier"}:
        return "Real shipments, no adapter to fetch them"
    if failures:
        return "Attempted, no data returned"
    if code in targets and not adapters.get(code):
        return "No adapter built"
    if code in targets:
        return "Adapter ready, no data in this run"
    return "No data in this run"


# --------------------------------------------------------------------------- #
# dashboard
# --------------------------------------------------------------------------- #

def kpis(data, as_of: datetime) -> dict:
    """The management-level headline numbers."""
    containers = data.containers
    shipments = data.shipment_rows
    buckets = Counter(c["status_bucket"] for c in containers)
    bl_buckets = Counter(s["status"] for s in shipments)
    attention = Counter(c["attention"] for c in containers)
    bl_attention = Counter(s["attention"] for s in shipments)

    actual_events = [e for e in data.events if e["time_basis"] == "ACTUAL" and e["event_dt"]]
    recent = [e for e in actual_events
              if nz.days_between(as_of, e["event_dt"]) is not None
              and 0 <= nz.days_between(as_of, e["event_dt"]) <= RECENT_ACTIVITY_DAYS]

    slipped = [c for c in containers
               if c["eta_slip_days"] is not None and c["eta_slip_days"] > 0
               and c["status_bucket"] != "Delivered"]
    overdue = [c for c in containers
               if c["eta_dt"] and c["status_bucket"] != "Delivered"
               and c["eta_dt"].date() < as_of.date()]
    stale = [c for c in containers
             if (c["idle_at_snapshot_days"] or 0) >= STALE_WARN_DAYS
             and c["status_bucket"] != "Delivered"]

    carriers_with_data = {c["carrier_code"] for c in containers if c["carrier_code"]}
    targets = {t["code"] for t in nz.TARGET_CARRIERS}
    event_dates = [e["event_dt"] for e in actual_events]
    snapshots = [c["snapshot_dt"] for c in containers if c["snapshot_dt"]]

    # Shipment-level rollups. A BL counts as delayed/overdue/stale if ANY of
    # its containers is -- the consignment has a problem even when only part of
    # it does, which is the whole point of tracking at this grain.
    flagged = {key: {c["bl"] for c in group}
               for key, group in (("delayed", slipped), ("overdue", overdue),
                                  ("stale", stale))}
    split = [s for s in shipments if s["split"] == "Yes"]

    return {
        "as_of": as_of,
        "total_shipments": len(shipments),
        "shipments_in_transit": bl_buckets.get("In Transit", 0),
        "shipments_delivered": bl_buckets.get("Delivered", 0),
        "shipments_at_origin": bl_buckets.get("At Origin", 0),
        "shipments_delayed": len(flagged["delayed"]),
        "shipments_overdue": len(flagged["overdue"]),
        "shipments_stale": len(flagged["stale"]),
        "shipments_attention": len(shipments) - bl_attention.get("OK", 0),
        "shipments_critical": bl_attention.get("Critical", 0),
        "split_shipments": len(split),
        "split_containers": sum(s["outlier_count"] for s in shipments),
        "multi_container_bls": sum(1 for s in shipments if s["containers"] > 1),
        "largest_bl": max((s["containers"] for s in shipments), default=0),
        "avg_containers_per_bl": (round(len(containers) / len(shipments), 1)
                                  if shipments else None),
        "total_containers": len(containers),
        "in_transit": buckets.get("In Transit", 0),
        "at_origin": buckets.get("At Origin", 0),
        "delivered": buckets.get("Delivered", 0),
        "no_events": buckets.get("No Events", 0),
        "delayed": len(slipped),
        "overdue": len(overdue),
        "stale": len(stale),
        "needs_attention": len(containers) - attention.get("OK", 0),
        "critical": attention.get("Critical", 0),
        "warning": attention.get("Warning", 0),
        "missing_eta": sum(1 for c in containers if c["eta_dt"] is None),
        "total_events": len(data.events),
        "actual_events": len(actual_events),
        "expected_events": len(data.events) - len(actual_events),
        "recent_events": len(recent),
        "recent_containers": len({e["container_id"] for e in recent}),
        "shipments": len(data.shipments),
        "tracking_failures": len(data.errors),
        "carriers_with_data": len(carriers_with_data),
        "target_carriers": len(targets),
        "target_carriers_with_data": len(carriers_with_data & targets),
        "vessels": len({c["vessel"] for c in containers if c["vessel"]}),
        "locations": len({c["last_actual_location"] for c in containers
                          if c["last_actual_location"]}),
        "runs": len(data.runs),
        "earliest_event": min(event_dates) if event_dates else None,
        "latest_event": max(event_dates) if event_dates else None,
        "oldest_snapshot": min(snapshots) if snapshots else None,
        "newest_snapshot": max(snapshots) if snapshots else None,
        "avg_idle_days": _avg([c["idle_at_snapshot_days"] for c in containers
                               if c["idle_at_snapshot_days"] is not None]),
        # Averaged over DELAYED containers only. Averaging slips and
        # pull-ins together produced 0.9 days on a dataset whose delayed
        # containers average 12 -- a number that reads as "nothing wrong".
        "avg_eta_slip": _avg([c["eta_slip_days"] for c in containers
                              if c["eta_slip_days"] is not None
                              and c["eta_slip_days"] > 0]),
        "eta_pulled_in": sum(1 for c in containers
                             if c["eta_slip_days"] is not None
                             and c["eta_slip_days"] < 0),
    }


def _avg(values) -> float | None:
    return round(sum(values) / len(values), 1) if values else None


def distributions(data, as_of: datetime) -> dict:
    """The breakdowns the dashboard charts and tables read from.

    Each one answers a question someone actually asks, which is the test for
    whether it belongs on the dashboard at all:
      by_carrier      who is carrying our cargo?
      by_status       how much is moving vs sitting vs done?
      by_location     where is our cargo piling up?
      by_vessel       which sailings would one delay hit?
      by_destination  which inbound markets are affected?
      by_phase        which stage of the journey is the fleet in?
      activity        is tracking data still arriving, or has it gone quiet?
    """
    containers = data.containers
    shipments = data.shipment_rows
    status_order = ["In Transit", "At Origin", "Delivered", "No Events"]

    by_status = [(s, sum(1 for c in containers if c["status_bucket"] == s))
                 for s in status_order]
    by_status = [row for row in by_status if row[1]]

    by_carrier = Counter(c["carrier_display"] for c in containers)
    by_location = Counter(c["last_actual_location"] for c in containers
                          if c["last_actual_location"])
    by_vessel = Counter(c["vessel"] for c in containers if c["vessel"])
    by_destination = Counter(c["destination"] for c in containers if c.get("destination"))
    by_phase = Counter(e["phase"] for e in data.events if e["time_basis"] == "ACTUAL")

    activity = []
    for offset in range(ACTIVITY_CHART_DAYS - 1, -1, -1):
        day = (as_of - _days(offset)).date()
        count = sum(1 for e in data.events
                    if e["time_basis"] == "ACTUAL" and e["event_dt"]
                    and e["event_dt"].date() == day)
        activity.append((day, count))

    bl_by_status = [(s, sum(1 for r in shipments if r["status"] == s))
                    for s in status_order]
    bl_by_status = [row for row in bl_by_status if row[1]]

    # Size bands rather than raw counts: "how much cargo rides on one BL" is a
    # shape question, and 29 individual bar heights answer it worse than five.
    bands = [("1 container", lambda n: n == 1),
             ("2-5", lambda n: 2 <= n <= 5),
             ("6-10", lambda n: 6 <= n <= 10),
             ("11-20", lambda n: 11 <= n <= 20),
             ("21+", lambda n: n >= 21)]
    bl_by_size = [(label, sum(1 for r in shipments if test(r["containers"])))
                  for label, test in bands]
    bl_by_size = [row for row in bl_by_size if row[1]]

    return {
        "bl_by_status": bl_by_status,
        "bl_by_carrier": Counter(r["carrier"] for r in shipments).most_common(),
        "bl_by_size": bl_by_size,
        "by_carrier": by_carrier.most_common(),
        "by_status": by_status,
        "by_location": by_location.most_common(TOP_N),
        "by_vessel": by_vessel.most_common(TOP_N),
        "by_destination": by_destination.most_common(TOP_N),
        "by_phase": by_phase.most_common(),
        "activity": activity,
        "by_exception": Counter(
            f["category"] for f in data.exceptions).most_common(),
    }


def _days(n: int):
    from datetime import timedelta
    return timedelta(days=n)


def status_mapping_rows(data) -> list[dict]:
    """Every (carrier, original wording) pair seen, and what it normalised to.

    This is the audit trail for section 7 of the requirement: it shows the
    original event preserved next to the normalised one, and flags any row that
    a keyword fallback produced rather than a confirmed mapping.
    """
    seen: Counter = Counter()
    for event in data.raw_events:
        seen[(event.get("carrier"), event.get("description"))] += 1

    rows = []
    for (code, description), count in seen.most_common():
        milestone, source = nz.normalize_event(description)
        rows.append({
            "carrier": nz.carrier_name(code),
            "original_event": description,
            "normalized_status": nz.milestone_label(milestone),
            "milestone": milestone,
            "phase": nz.milestone_phase(milestone),
            "is_forecast": "Yes" if nz.milestone_is_forecast(milestone) else "No",
            "mapping_source": source,
            "event_rows": count,
        })
    return rows


def carrier_mapping_rows(data) -> list[dict]:
    """The eight target carriers plus anything else seen, with real coverage."""
    summary = {row["code"]: row for row in carrier_summary(data)}
    rows = []
    for target in nz.TARGET_CARRIERS:
        found = summary.get(target["code"], {})
        rows.append({
            "carrier": target["name"],
            "code": target["code"],
            "target_carrier": "Yes",
            "tracking_adapter": target["adapter"] or "not built",
            "containers": found.get("containers", 0),
            "events": found.get("events", 0),
            "failures": found.get("failures", 0),
            "coverage": found.get("coverage", "No data in this run"),
        })
    for code, row in sorted(summary.items()):
        if code in {t["code"] for t in nz.TARGET_CARRIERS}:
            continue
        rows.append({
            "carrier": row["carrier"],
            "code": code,
            "target_carrier": "No",
            "tracking_adapter": row["adapter"],
            "containers": row["containers"],
            "events": row["events"],
            "failures": row["failures"],
            "coverage": row["coverage"],
        })
    return rows
