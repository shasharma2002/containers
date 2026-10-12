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

import re
from collections import Counter, defaultdict
from datetime import datetime, timedelta

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

# Free time (detention). Once a box is discharged at the destination port the
# consignee has N free days to collect it; after that the carrier charges per
# container per day. With 23 containers on one BL that is the largest avoidable
# cost in this data, so it gets its own sheet and its own findings.
FREE_TIME_WARN_DAYS = 5      # flag when this many days or fewer remain
FREE_TIME_CRITICAL_DAYS = 2

# Arrival planning horizons, in days.
ARRIVAL_HORIZONS = (7, 14, 30)

# Document readiness. A 'DOCUMENT POSITION' cell that reads ok/received means the
# paperwork is in hand. Anything else with cargo about to land is a problem:
# without documents the box cannot be cleared, and free time runs out anyway.
DOCS_READY = re.compile(r"\b(ok|received|rec'?d|done|clear(ed)?|complete)\b", re.I)
DOCS_DEADLINE_DAYS = 14

# A lane needs this many completed journeys before its median is worth comparing
# an in-flight shipment against. Below that it is a sample, not a benchmark.
LANE_MIN_JOURNEYS = 3
# How much slower than its lane's median a journey must run to be worth flagging.
LANE_SLOW_FACTOR = 1.25

# Discharge milestones. Whether a discharge is at the DESTINATION port or at a
# transshipment hub cannot be read off the milestone alone:
#
#   MSC     labels them distinctly -- TRANSSHIPMENT_DISCHARGED vs
#           DISCHARGED_FROM_VESSEL. Unambiguous.
#   Maersk  fires the same DISCHARG activity at both. Real examples from this
#           data: BL 274146535 discharges at COLOMBO and BL 274659478 at
#           TANJUNG PELEPAS, both transshipment hubs, both reported exactly like
#           a final discharge.
#
# So a discharge counts as the destination one only when NO further vessel leg
# follows it in the itinerary -- see _final_discharge(). Treating any discharge
# as final starts the free-time clock at Colombo, weeks early, and invents
# detention exposure that does not exist.
PORT_ARRIVAL_MILESTONES = {"DISCHARGED_FROM_VESSEL"}
# Any onward sea leg after a discharge proves that discharge was a transshipment.
# The check has to include EXPECTED events: Colombo's onward departure is still
# a forecast at the moment the box is sitting there, and ignoring forecasts would
# make the transshipment look like a final arrival.
VESSEL_LOAD_MILESTONES = {"LOADED_ON_VESSEL", "TRANSSHIPMENT_LOADED", "DEPARTURE"}
# Milestones that mean the box has LEFT the port -- the detention clock stops.
PORT_EXIT_MILESTONES = {"RAIL_DEPARTURE", "GATE_OUT", "TRUCK_MOVE", "DELIVERED",
                        "EMPTY_RETURN"}
# Milestones that open a journey, for end-to-end transit time.
JOURNEY_START_MILESTONES = {"EMPTY_DISPATCHED", "GATE_IN"}
# Milestones that close one.
JOURNEY_END_MILESTONES = {"RAIL_ARRIVAL", "DELIVERED", "EMPTY_RETURN"}

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

    _attach_journey(data, by_container, as_of)
    _build_shipments(data, as_of)
    _attach_free_time(data, as_of)
    data.lane_benchmarks = _lane_benchmarks(data)
    _attach_exceptions(data, as_of)
    return as_of


def _lane_benchmarks(data) -> dict:
    """Per lane, how long a completed journey actually takes.

    Built from this operation's OWN completed journeys, not a published transit
    time -- so it already includes the transshipment waits and rail legs these
    routes really involve. A lane with fewer than LANE_MIN_JOURNEYS completed
    journeys is kept for display but never used to flag a shipment as slow: a
    median of one is not a benchmark.
    """
    lanes: dict[str, list] = defaultdict(list)
    for container in data.containers:
        if container["transit_days"] is None or container["journey_end_dt"] is None:
            continue
        lane = _lane(container)
        if lane:
            lanes[lane].append(container)

    out = {}
    for lane, containers in lanes.items():
        days = sorted(c["transit_days"] for c in containers)
        ocean = [c["ocean_days"] for c in containers if c["ocean_days"] is not None]
        dwell = [c["port_dwell_days"] for c in containers
                 if c["port_dwell_days"] is not None]
        out[lane] = {
            "lane": lane,
            "journeys": len(days),
            "containers": len(containers),
            "median": _median(days),
            "fastest": days[0],
            "slowest": days[-1],
            "median_ocean": _median(ocean),
            "median_port_dwell": _median(dwell),
            "carriers": ", ".join(sorted({c["carrier_display"] for c in containers})),
            "usable": len(days) >= LANE_MIN_JOURNEYS,
        }
    return out


# --------------------------------------------------------------------------- #
# journey stages -- where the time actually goes
# --------------------------------------------------------------------------- #

def _attach_journey(data, by_container: dict, as_of: datetime) -> None:
    """Split each container's history into the stages the business pays for.

    Three numbers come out of this, and each answers a question the flat event
    list cannot:

      port_dwell_days   how long the box sat at the DESTINATION port between
                        being discharged and moving inland. This is where time
                        is quietly lost, and where detention is charged.
      transit_days      origin gate-in to arrival at destination. The number to
                        quote a supplier, and the baseline a slow shipment is
                        measured against.
      ocean_days        loaded on vessel to discharged at destination port.

    The destination-port discharge is the pivot for all of it, and identifying it
    correctly is the whole difficulty -- see _final_discharge() and the note on
    PORT_ARRIVAL_MILESTONES. Dwell at a transshipment hub is real and worth
    knowing, but it is not detention exposure, so it is measured separately.
    """
    for container in data.containers:
        events = by_container.get((str(container.get("container_id")),
                                   str(container.get("run_id"))), [])
        dated = [e for e in events if e["event_dt"]]
        actual = [e for e in dated if e["time_basis"] == "ACTUAL"]

        start = next((e for e in actual
                      if e["milestone"] in JOURNEY_START_MILESTONES), None)
        loaded = next((e for e in actual
                       if e["milestone"] == "LOADED_ON_VESSEL"), None)
        end = next((e for e in reversed(actual)
                    if e["milestone"] in JOURNEY_END_MILESTONES), None)

        # The discharge that actually lands the cargo, actual or still forecast.
        final = _final_discharge(dated)
        arrival = final if final and final["time_basis"] == "ACTUAL" else None
        expected_arrival = final if final and final is not arrival else None

        transship_days, transship_at = _transshipment_dwell(actual)
        container["transship_dwell_days"] = transship_days
        container["transship_location"] = transship_at

        # Sitting at a hub RIGHT NOW: the latest confirmed event is a discharge
        # that is not the final one, so the box is between vessels. A completed
        # hub wait is history; this one is still accruing and is what an ops user
        # needs to see.
        latest = actual[-1] if actual else None
        waiting = (latest is not None
                   and latest["milestone"] in ("TRANSSHIPMENT_DISCHARGED",
                                               "DISCHARGED_FROM_VESSEL")
                   and (final is None or latest is not final))
        container["at_transshipment"] = waiting
        container["transship_wait_days"] = (nz.days_between(as_of, latest["event_dt"])
                                            if waiting else None)
        container["transship_waiting_at"] = latest["location"] if waiting else None
        container["port_arrival_expected_dt"] = (expected_arrival["event_dt"]
                                                 if expected_arrival else None)
        container["port_of_discharge_expected"] = (expected_arrival["location"]
                                                   if expected_arrival else None)

        # Did it leave the port AFTER being discharged there? Only an exit that
        # follows the arrival counts -- an origin gate-out does not.
        exit_event = None
        if arrival:
            exit_event = next((e for e in actual
                               if e["milestone"] in PORT_EXIT_MILESTONES
                               and e["event_dt"] > arrival["event_dt"]), None)

        container["journey_start_dt"] = start["event_dt"] if start else None
        container["loaded_dt"] = loaded["event_dt"] if loaded else None
        container["port_arrival_dt"] = arrival["event_dt"] if arrival else None
        container["port_arrival_location"] = arrival["location"] if arrival else None
        container["port_exit_dt"] = exit_event["event_dt"] if exit_event else None
        container["journey_end_dt"] = end["event_dt"] if end else None
        container["at_port"] = bool(arrival and not exit_event)

        container["transit_days"] = nz.days_between(
            end["event_dt"] if end else None, start["event_dt"] if start else None)
        container["ocean_days"] = nz.days_between(
            arrival["event_dt"] if arrival else None,
            loaded["event_dt"] if loaded else None)
        # Still at the port: dwell is measured to now, and is still growing.
        container["port_dwell_days"] = nz.days_between(
            exit_event["event_dt"] if exit_event else (as_of if arrival else None),
            arrival["event_dt"] if arrival else None)
        container["inland_days"] = nz.days_between(
            end["event_dt"] if end else None,
            exit_event["event_dt"] if exit_event else None)


# --------------------------------------------------------------------------- #
# free time / detention
# --------------------------------------------------------------------------- #

def _attach_free_time(data, as_of: datetime) -> None:
    """Work out each BL's detention exposure. Sets free_time_* on the BL row.

    The clock is simple and entirely from real inputs:

        free time expires = destination-port discharge + free days from the sheet

    with one override: where the commercial sheet states an explicit expiry date,
    that wins, because somebody has already negotiated it and the sheet is the
    authority on what was agreed.

    Only containers STILL AT THE PORT are counted as exposed. A box that has
    railed out is no longer accruing port detention, so including it would
    inflate the number and train people to ignore the sheet. Where free days are
    not stated, no expiry is claimed -- the row says so instead of assuming an
    industry-standard 14.
    """
    free_by_bl: dict[str, dict] = {}
    for shipment in data.shipments:
        bl = shipment.get("bl_number")
        if not bl:
            continue
        entry = free_by_bl.setdefault(bl, {"free_days": None})
        if entry["free_days"] is None and shipment.get("free_days") is not None:
            entry["free_days"] = shipment["free_days"]

    by_bl: dict[str, list] = defaultdict(list)
    for container in data.containers:
        by_bl[container.get("bl")].append(container)

    for shipment in data.shipment_rows:
        containers = by_bl.get(shipment["bl"], [])
        stated = free_by_bl.get(shipment["bl"], {})
        free_days = stated.get("free_days")

        at_port = [c for c in containers if c["at_port"]]
        arrivals = [c["port_arrival_dt"] for c in containers if c["port_arrival_dt"]]
        # Earliest arrival: the first box ashore starts the clock for the BL.
        first_arrival = min(arrivals) if arrivals else None

        # One basis only. The sheet's own expiry DATE used to override this, and
        # no longer does: those cells are typed by hand and go stale. The free
        # days count still comes from the sheet -- a contractual term, not a
        # date -- but the clock it starts is the carrier's discharge event.
        expiry = None
        if first_arrival is not None and free_days is not None:
            expiry = first_arrival + timedelta(days=free_days)
            basis = "carrier discharge %s + %d free day(s) from the sheet" % (
                first_arrival.strftime("%Y-%m-%d"), free_days)
        elif first_arrival is not None:
            basis = "no free days on the sheet (its FREE TIME cell is a date, "                     "which is not used, or is empty)"
        else:
            basis = "not yet discharged at the destination port"

        dwell = [c["port_dwell_days"] for c in at_port
                 if c["port_dwell_days"] is not None]

        shipment["free_days"] = free_days
        shipment["port_arrival_dt"] = first_arrival
        shipment["port_arrival_location"] = next(
            (c["port_arrival_location"] for c in containers
             if c["port_arrival_location"]), None)
        shipment["free_time_expiry"] = expiry.date() if expiry else None
        shipment["free_time_basis"] = basis
        shipment["free_days_left"] = nz.days_between(expiry, as_of) if expiry else None
        shipment["containers_at_port"] = len(at_port)
        shipment["max_port_dwell_days"] = max(dwell) if dwell else None
        shipment["port_dwell_days"] = (max(
            [c["port_dwell_days"] for c in containers
             if c["port_dwell_days"] is not None] or [0]) or None)
        shipment["transit_days"] = _median(
            [c["transit_days"] for c in containers if c["transit_days"] is not None])
        shipment["ocean_days"] = _median(
            [c["ocean_days"] for c in containers if c["ocean_days"] is not None])
        shipment["lane"] = _lane(containers[0]) if containers else None
        shipment["docs_ready"] = _docs_ready(shipment.get("document_position"))

        # Expected arrival at the destination port, for planning cargo that is
        # still at sea. Earliest box ashore is what the yard has to be ready for.
        expected = [c["port_arrival_expected_dt"] for c in containers
                    if c["port_arrival_expected_dt"]]
        shipment["port_arrival_expected"] = min(expected).date() if expected else None
        shipment["port_of_discharge_expected"] = next(
            (c["port_of_discharge_expected"] for c in containers
             if c["port_of_discharge_expected"]), None)
        worst_hub = max((c for c in containers
                         if c.get("transship_dwell_days") is not None),
                        key=lambda c: c["transship_dwell_days"], default=None)
        shipment["transship_dwell_days"] = (worst_hub["transship_dwell_days"]
                                            if worst_hub else None)
        shipment["transship_location"] = (worst_hub["transship_location"]
                                          if worst_hub else None)
        shipment["free_time_status"] = _free_time_status(shipment)
        if shipment["containers_at_port"] == 0:
            # Nothing on the ground, so there is no countdown to report. Leaving
            # a stale "-38" beside "Cleared the port" invites someone to chase an
            # invoice that will never arrive.
            shipment["free_days_left"] = None


def _free_time_status(shipment: dict) -> str:
    """One phrase an ops user can act on, instead of a bare signed number.

    A negative "days left" on its own is ambiguous: a BL whose free time expired
    three weeks ago but whose boxes all railed out owes nothing, while one that
    expired yesterday with 21 boxes still on the ground is an invoice forming.
    The distinction is whether any container is STILL AT THE PORT, so the status
    leads with that.
    """
    at_port = shipment["containers_at_port"]
    left = shipment["free_days_left"]

    if shipment["port_arrival_dt"] is None:
        return "Not yet ashore"
    if at_port == 0:
        return "Cleared the port"
    if left is None:
        return "%d at port, no free days stated" % at_port
    if left < 0:
        return "EXPIRED %d day(s) ago, %d still at port" % (abs(left), at_port)
    if left <= FREE_TIME_CRITICAL_DAYS:
        return "Expires in %d day(s), %d still at port" % (left, at_port)
    return "%d day(s) left, %d at port" % (left, at_port)


def _final_discharge(dated: list[dict]) -> dict | None:
    """The discharge that actually lands the cargo, or None if there isn't one.

    A discharge is final when no vessel leg follows it anywhere in the itinerary,
    forecast included. Worked through on real rows:

        BL 274146535  discharged COLOMBO 13 Sep, then an EXPECTED departure from
                      COLOMBO on 28 Sep -> transshipment, not final.
        BL 274659478  discharged TANJUNG PELEPAS 18 Sep, then LOADED_ON_VESSEL
                      there on 20 Sep -> transshipment, not final.

    MSC's own TRANSSHIPMENT_DISCHARGED rows are excluded by milestone, so this
    only has to reason about Maersk's ambiguous DISCHARG. The returned event may
    be EXPECTED -- a box still at sea has a forecast destination discharge, which
    is exactly what the arrivals planner needs and what the free-time clock must
    NOT start from.
    """
    candidate = None
    for index, event in enumerate(dated):
        if event["milestone"] not in PORT_ARRIVAL_MILESTONES:
            continue
        if any(later["milestone"] in VESSEL_LOAD_MILESTONES
               for later in dated[index + 1:]):
            continue                      # another sea leg follows: transshipment
        candidate = event                 # keep the last qualifying discharge
    return candidate


def _transshipment_dwell(actual: list[dict]) -> tuple[int | None, str | None]:
    """Longest completed wait at a transshipment hub -> (days, location).

    Measured discharge-to-reload at the same hub. This is where these routes
    quietly lose a week, and it is NOT detention: the box is never available for
    collection, so no free time is running. Reported, never charged.
    """
    worst, where = None, None
    for index, event in enumerate(actual):
        if event["milestone"] not in ("TRANSSHIPMENT_DISCHARGED",
                                      "DISCHARGED_FROM_VESSEL"):
            continue
        reload = next((later for later in actual[index + 1:]
                       if later["milestone"] in VESSEL_LOAD_MILESTONES), None)
        if reload is None:
            continue
        days = nz.days_between(reload["event_dt"], event["event_dt"])
        if days is not None and (worst is None or days > worst):
            worst, where = days, event["location"]
    return worst, where


def _median(values):
    if not values:
        return None
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return round((ordered[middle - 1] + ordered[middle]) / 2, 1)


def _lane(container: dict) -> str | None:
    origin, destination = container.get("origin"), container.get("destination")
    if origin and destination:
        return "%s -> %s" % (origin, destination)
    return origin or destination


def _docs_ready(position) -> str:
    """'ok, 20.09.26' -> Ready. Anything else with text -> Pending.

    Deliberately conservative: only wording that positively says the paperwork
    is in hand counts as ready. 'PENDING, req sign' must not read as ready, and
    an empty cell is unknown rather than either.
    """
    if position in (None, ""):
        return "Not stated"
    return "Ready" if DOCS_READY.search(str(position)) else "Pending"


# --------------------------------------------------------------------------- #
# bill of lading -- the grain the business actually works in
# --------------------------------------------------------------------------- #

def _build_shipments(data, as_of: datetime) -> None:  # noqa: C901
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
            container["company"] = commercial.get(bl, {}).get("company")

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
            "company": sheet.get("company"),
            "beneficiary": sheet.get("beneficiary"),
            "commodity": sheet.get("commodity"),
            "broker": sheet.get("broker"),
            "document_position": sheet.get("document_position"),
            "amount": sheet.get("amount"),
            "amount_derived": sheet.get("amount_derived"),
        })

    data.shipment_rows = rows


def sort_shipments(rows: list[dict]) -> list[dict]:
    """Worst first, then soonest arriving.

    Sorted on attention rather than size, because the question the sheet answers
    is "what do I deal with now", not "which BL is biggest".
    """
    rank = {"Critical": 0, "Warning": 1, "OK": 2}
    return sorted(rows, key=lambda r: (rank.get(r.get("attention"), 3),
                                       r["days_to_eta"] is None,
                                       r["days_to_eta"] if r["days_to_eta"] is not None else 0,
                                       str(r["bl"])))


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
                                      "broker": set(), "company": set(),
                                      "document_position": None,
                                      "amount": 0.0,
                                      "amount_derived": 0.0})
        for field in ("beneficiary", "commodity", "broker", "company"):
            if shipment.get(field):
                entry[field].add(str(shipment[field]))
        if shipment.get("document_position") and entry["document_position"] is None:
            entry["document_position"] = shipment["document_position"]
        for field in ("amount", "amount_derived"):
            if isinstance(shipment.get(field), (int, float)):
                entry[field] += float(shipment[field])

    return {bl: {**entry,
                 "beneficiary": "; ".join(sorted(entry["beneficiary"])) or None,
                 "commodity": "; ".join(sorted(entry["commodity"])) or None,
                 "broker": "; ".join(sorted(entry["broker"])) or None,
                 "company": "; ".join(sorted(entry["company"])) or None,
                 "amount": entry["amount"] or None,
                 "amount_derived": entry["amount_derived"] or None}
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
                "company": container.get("company"),
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

        # No recent movement -- but ONLY where sitting still is actually odd.
        #
        # A box mid-ocean is not scanned for weeks at a time: these routes run
        # 40-day sea legs, so "no event for 7 days" is the normal state of a
        # loaded vessel and says nothing. Flagging it produced 7 findings out of
        # 8 that no one could act on, which is how a report teaches people to
        # ignore it.
        #
        # Sitting still is only worth a human's time when the box is somewhere
        # it should be MOVING THROUGH -- on the ground at a destination port, or
        # between vessels at a transshipment hub. Those are the places where a
        # stalled box costs money and someone can actually push it along.
        idle = container["idle_at_snapshot_days"]
        where = ("the destination port" if container["at_port"]
                 else "a transshipment hub" if container.get("at_transshipment")
                 else None)
        if (where and idle is not None and idle >= STALE_WARN_DAYS
                and not delivered):
            add("Shipment", "Critical" if idle >= STALE_CRITICAL_DAYS else "Warning",
                "Sitting still at %s" % where,
                "No movement for %d day(s) at %s, as of the last check (%s). A "
                "box in open water goes unscanned for weeks and is not flagged; "
                "one standing at a port or hub is somewhere it should be passing "
                "through."
                % (idle, container["last_actual_location"] or "an unreported location",
                   _d(container["snapshot_dt"])),
                "latest confirmed event vs snapshot, at port or hub only", idle)

    # ---- BL-level findings ------------------------------------------------ #
    # Two kinds live here, with different guards:
    #
    #   commercial         free time, documents, lane speed. These are about the
    #                      shipment as a commercial thing and apply to a BL
    #                      carrying one box exactly as much as twenty-three.
    #   consignment shape  containers diverging, containers separated in time,
    #                      arrival dates spread. These describe how the boxes on
    #                      one BL relate to EACH OTHER, so a single-container BL
    #                      cannot have them.
    for shipment in data.shipment_rows:
        _commercial_findings(findings, shipment, data, as_of)

        if shipment["single_container"]:
            continue

        if shipment["outlier_count"]:
            severity = ("Critical" if shipment["outlier_count"] * 2 <= shipment["containers"]
                        else "Warning")
            findings.append(_bl_finding(
                shipment, severity, "Containers not travelling together",
                "%d of %d container(s) on this BL differ from the rest on %s. "
                "The main body is '%s' at %s; the odd one(s) are %s. This is the "
                "case worth checking by hand -- a box damaged, rolled to a later "
                "vessel, or held."
                % (shipment["outlier_count"], shipment["containers"],
                   shipment["divergence"] or "at least one field",
                   shipment["confirmed_status"] or "unreported",
                   shipment["location"] or "an unreported location",
                   shipment["outliers"]),
                "per-container status/milestone/location/vessel",
                shipment["outlier_count"],
                scope="%d of %d containers" % (shipment["outlier_count"],
                                               shipment["containers"])))

        eta_spread = shipment["eta_spread_days"]
        if (eta_spread is not None and eta_spread >= ETA_SPREAD_WARN_DAYS
                and shipment["delivered_count"] < shipment["containers"]):
            findings.append(_bl_finding(
                shipment,
                "Critical" if eta_spread >= ETA_SPREAD_WARN_DAYS * 3 else "Warning",
                "Arrival dates spread across the consignment",
                "The carrier expects the %d containers on this BL to arrive over "
                "%d day(s) (%s to %s). Plan delivery and free time in waves. A "
                "day or two apart is normal and is not reported."
                % (shipment["containers"], eta_spread,
                   shipment["eta_earliest"], shipment["eta_latest"]),
                "spread of carrier ETAs across containers", eta_spread))

        spread = shipment["move_spread_days"]
        fully_delivered = shipment["delivered_count"] == shipment["containers"]
        if (spread is not None and spread >= CONTAINER_SPREAD_WARN_DAYS
                and not fully_delivered):
            findings.append(_bl_finding(
                shipment,
                "Critical" if spread >= CONTAINER_SPREAD_CRITICAL_DAYS else "Warning",
                "Containers separated in time",
                "The %d containers on this BL last moved %d day(s) apart (%s to "
                "%s). Minutes apart is normal -- a terminal working down a stack. "
                "Days apart means part of the consignment did not travel with "
                "the rest."
                % (shipment["containers"], spread, _d(shipment["earliest_move"]),
                   _d(shipment["latest_move"])),
                "spread of latest confirmed event across containers", spread))

    # ---- commercial lines with no bill of lading yet ---------------------- #
    # These cannot be tracked and never will be until a BL exists, so they are
    # an action (chase it) rather than a tracking failure. Grouped, because a
    # company with six of them does not want six rows.
    pending = defaultdict(list)
    for row in data.shipments:
        if not row.get("bl_number"):
            pending[row.get("company") or "(not stated)"].append(row)
    for company, rows in pending.items():
        findings.append({
            "container_id": None, "bl": None, "company": company,
            "carrier": None, "tracking_number": None,
            "class": "Shipment", "severity": "Warning",
            "category": "No bill of lading yet",
            "detail": "%d commercial line(s) for %s have no BL on the sheet, so "
                      "nothing can be tracked for them. Suppliers: %s."
                      % (len(rows), company,
                         ", ".join(sorted({str(r.get("beneficiary"))
                                           for r in rows if r.get("beneficiary")})[:4])
                         or "not stated"),
            "evidence": "BL NO. column is empty", "value": len(rows),
            "status": "Not tracked", "last_move": None, "location": None,
            "eta": None, "destination": None,
            "_bl_level": True, "_affected": None,
            "_scope": "%d sheet line(s)" % len(rows),
        })

    # Tracking failures are NOT repeated here. They have their own sheet, which
    # carries the carrier's own reason and what to do about each kind -- putting
    # them on the action list as well just made it longer.

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

    # Split by CLASS, into two sheets with two different readers.
    #
    #   actions    something about the cargo needs a person: chase the carrier,
    #              clear the box, find the paperwork. This is the sheet someone
    #              opens in the morning, so nothing that cannot be acted on is
    #              allowed on it.
    #   data notes something about the INFORMATION needs fixing, or is simply a
    #              carrier limitation. Real, worth recording, and useless to an
    #              ops user at 9am -- so it goes to Data Quality instead.
    #
    # Mixing them is what made the old Exceptions sheet 40% unactionable.
    rolled = _rollup(findings, data, severity_rank)
    data.exceptions = [f for f in rolled if f["class"] == "Shipment"]
    # Anything not actionable would go here; nothing does today, because every
    # data problem is reported once by validate.findings() instead.
    data.data_findings = [f for f in rolled if f["class"] != "Shipment"]
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
    summaries: dict[str, list] = defaultdict(list)
    for finding in data.exceptions:
        bl = finding.get("bl")
        if bl is None:
            continue
        counts[bl] += 1
        shipment_counts[bl] += 1
        summaries[bl].append(finding["category"])
        if (bl not in worst
                or severity_rank.get(finding["severity"], 3) < severity_rank.get(worst[bl], 3)):
            worst[bl] = finding["severity"]

    for shipment in data.shipment_rows:
        bl = shipment["bl"]
        shipment["attention"] = worst.get(bl, "OK")
        shipment["issue_count"] = counts.get(bl, 0)
        shipment["shipment_issue_count"] = shipment_counts.get(bl, 0)
        # The headline reason, on the shipment row itself. Without it a reader
        # sees "Critical" and has to go hunting on another sheet to find out
        # why -- which is the single most common thing to do with this report.
        shipment["needs"] = "; ".join(dict.fromkeys(summaries.get(bl, []))) or None


def _commercial_findings(findings: list[dict], shipment: dict, data, as_of) -> None:
    """Findings about the shipment as a commercial thing, not as a set of boxes.

    Detention, documents and lane speed all apply to a BL carrying one container
    exactly as much as to one carrying twenty-three, so these are raised for
    every BL -- unlike the consignment-relationship findings, which need at
    least two containers to mean anything.
    """
    # ---- detention: the largest avoidable cost in this data ------------ #
    at_port = shipment["containers_at_port"]
    left = shipment["free_days_left"]
    if at_port and left is not None:
        if left < 0:
            findings.append(_bl_finding(
                shipment, "Critical", "Free time expired, cargo still at port",
                "Free time ran out on %s, %d day(s) ago, and %d container(s) "
                "are still at %s. Detention is accruing per container per "
                "day. Basis: %s."
                % (shipment["free_time_expiry"], abs(left), at_port,
                   shipment["port_arrival_location"] or "the port",
                   shipment["free_time_basis"]),
                "port discharge + free days from the commercial sheet", left))
        elif left <= FREE_TIME_WARN_DAYS:
            findings.append(_bl_finding(
                shipment,
                "Critical" if left <= FREE_TIME_CRITICAL_DAYS else "Warning",
                "Free time expiring",
                "Free time expires %s -- %d day(s) away -- with %d "
                "container(s) still at %s. Clear them or negotiate an "
                "extension before detention starts. Basis: %s."
                % (shipment["free_time_expiry"], left, at_port,
                   shipment["port_arrival_location"] or "the port",
                   shipment["free_time_basis"]),
                "port discharge + free days from the commercial sheet", left))
    elif at_port and shipment["free_days"] is None:
        findings.append(_bl_finding(
            shipment, "Warning", "Cargo at port, no free time on record",
            "%d container(s) are sitting at %s but the commercial sheet "
            "states no free days, so no expiry can be worked out. Detention "
            "may already be running. Fill in FREE TIME on the sheet."
            % (at_port, shipment["port_arrival_location"] or "the port"),
            "FREE TIME column is empty", at_port))

    # ---- documents: cargo cannot be cleared without them --------------- #
    arriving = shipment["port_arrival_dt"] is not None or (
        shipment["port_arrival_expected"] is not None
        and nz.days_between(
            datetime.combine(shipment["port_arrival_expected"],
                             datetime.min.time()), as_of) is not None
        and nz.days_between(
            datetime.combine(shipment["port_arrival_expected"],
                             datetime.min.time()), as_of) <= DOCS_DEADLINE_DAYS)
    if shipment["docs_ready"] == "Pending" and arriving:
        ashore = shipment["port_arrival_dt"] is not None
        findings.append(_bl_finding(
            shipment, "Critical" if ashore else "Warning",
            "Documents not ready and cargo arriving",
            "Document position reads %r while the cargo %s. Without "
            "documents the consignment cannot be cleared, and free time runs "
            "regardless."
            % (shipment.get("document_position"),
               "is already at %s" % (shipment["port_arrival_location"] or "port")
               if ashore else "is due at %s on %s"
               % (shipment["port_of_discharge_expected"] or "port",
                  shipment["port_arrival_expected"])),
            "DOCUMENT POSITION column", None))

    # ---- running slower than this lane's own history ------------------- #
    benchmark = data.lane_benchmarks.get(shipment["lane"])
    actual_transit = shipment["transit_days"]
    if (benchmark and actual_transit
            and benchmark["journeys"] >= LANE_MIN_JOURNEYS
            and actual_transit > benchmark["median"] * LANE_SLOW_FACTOR):
        findings.append(_bl_finding(
            shipment, "Warning", "Slower than this lane normally runs",
            "This shipment has taken %s days against a median of %s over %d "
            "completed journeys on %s. The benchmark is this lane's own "
            "history, not a target."
            % (actual_transit, benchmark["median"], benchmark["journeys"],
               shipment["lane"]),
            "transit days vs lane median", actual_transit))

def _bl_finding(shipment: dict, severity: str, category: str, detail: str,
                evidence: str, value, scope: str | None = None) -> dict:
    """A finding that belongs to the whole consignment, not one container.

    `scope` says who on the BL is affected and is the caller's to choose: a
    detention finding wants "7 at port", a divergence finding wants "1 of 13
    containers". Defaulting it for everyone put "1 at port" on findings that had
    nothing to do with the port.
    """
    return {
        "container_id": None,
        "bl": shipment["bl"],
        "company": shipment.get("company"),
        "carrier": shipment["carrier"],
        "tracking_number": shipment["bl"],
        "class": "Shipment",
        "severity": severity,
        "category": category,
        "detail": detail,
        "evidence": evidence,
        "value": value,
        "status": shipment["status"],
        "last_move": shipment["latest_move"],
        "location": shipment["port_arrival_location"] or shipment["location"],
        "eta": shipment["eta"],
        "destination": shipment["destination"],
        "_bl_level": True,
        "_affected": None,
        "_scope": scope or "Consignment-wide",
    }


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
        "companies": len({s["company"] for s in shipments if s.get("company")}),
        # Detention exposure: containers on the ground past, or close to, free time.
        "containers_at_port": sum(s["containers_at_port"] for s in shipments),
        "free_time_expired": sum(
            s["containers_at_port"] for s in shipments
            if s["containers_at_port"] and s["free_days_left"] is not None
            and s["free_days_left"] < 0),
        "free_time_expiring": sum(
            s["containers_at_port"] for s in shipments
            if s["containers_at_port"] and s["free_days_left"] is not None
            and 0 <= s["free_days_left"] <= FREE_TIME_WARN_DAYS),
        "free_time_unknown": sum(
            s["containers_at_port"] for s in shipments
            if s["containers_at_port"] and s["free_days_left"] is None),
        "docs_pending": sum(1 for s in shipments if s["docs_ready"] == "Pending"),
        "at_hub": sum(1 for c in containers if c.get("at_transshipment")),
        # Same rows and the same basis as the Arrivals Planner sheet, so the
        # tile and the sheet can never disagree.
        "arriving_7": sum(1 for r in arrival_rows(data, as_of)
                          if r["days_away"] is not None and 0 <= r["days_away"] <= 7),
        "ashore_now": sum(1 for r in arrival_rows(data, as_of)
                          if r["arrival_window"] == "Already ashore"),
        "median_transit": _median([s["transit_days"] for s in shipments
                                   if s["transit_days"] is not None]),
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

    # Company comes from the master workbook's sheet names. A single-company
    # file (or the older bl_list.xlsx, which states none) would make this a
    # one-bar chart, so it is emitted only when there is something to compare.
    by_company = Counter(r["company"] for r in shipments if r.get("company"))
    containers_by_company = Counter()
    for row in shipments:
        if row.get("company"):
            containers_by_company[row["company"]] += row["containers"]

    # Dwell by location, worst first -- the "where are we losing time" chart.
    dwell = [(r["location"], r["median_dwell"]) for r in port_rows(data)
             if r["median_dwell"]][:TOP_N]

    free_time = [
        ("Expired, still at port", sum(
            s["containers_at_port"] for s in shipments
            if s["containers_at_port"] and s["free_days_left"] is not None
            and s["free_days_left"] < 0)),
        ("Expiring within %d days" % FREE_TIME_WARN_DAYS, sum(
            s["containers_at_port"] for s in shipments
            if s["containers_at_port"] and s["free_days_left"] is not None
            and 0 <= s["free_days_left"] <= FREE_TIME_WARN_DAYS)),
        ("Comfortable", sum(
            s["containers_at_port"] for s in shipments
            if s["containers_at_port"] and s["free_days_left"] is not None
            and s["free_days_left"] > FREE_TIME_WARN_DAYS)),
        ("No free days on record", sum(
            s["containers_at_port"] for s in shipments
            if s["containers_at_port"] and s["free_days_left"] is None)),
    ]
    free_time = [row for row in free_time if row[1]]

    arrivals = arrival_rows(data, as_of)
    window_order = ["Already ashore"] + ["Next %d days" % h for h in ARRIVAL_HORIZONS]         + ["Beyond %d days" % ARRIVAL_HORIZONS[-1], "Unknown"]
    windows = Counter(r["arrival_window"] for r in arrivals)
    by_window = [(w, windows[w]) for w in window_order if windows.get(w)]

    lanes = [(r["lane"], r["median"]) for r in lane_rows(data) if r["median"]]

    return {
        "by_dwell": dwell,
        "by_free_time": free_time,
        "by_arrival_window": by_window,
        "by_lane": lanes,
        "bl_by_company": by_company.most_common() if len(by_company) > 1 else [],
        "containers_by_company": (containers_by_company.most_common()
                                  if len(containers_by_company) > 1 else []),
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


def free_time_rows(data) -> list[dict]:
    """Detention exposure, worst first. Only BLs with cargo ashore or due."""
    rows = [dict(s) for s in data.shipment_rows
            if s["port_arrival_dt"] is not None
            or s["port_arrival_expected"] is not None]
    rows.sort(key=lambda r: (r["containers_at_port"] == 0,
                             r["free_days_left"] is None,
                             r["free_days_left"] if r["free_days_left"] is not None else 9999))
    return rows


def arrival_rows(data, as_of: datetime) -> list[dict]:
    """What is landing, soonest first -- the daily planning list.

    Uses the confirmed arrival where there is one and the carrier's forecast
    otherwise, and says which it used. Mixing the two without labelling them
    would let a forecast read as a fact.
    """
    rows = []
    for shipment in data.shipment_rows:
        if shipment["status"] == "Delivered":
            continue
        ashore = shipment["port_arrival_dt"]
        expected = shipment["port_arrival_expected"]
        eta = shipment["eta"]

        if ashore is not None:
            when, basis = ashore.date(), "arrived (confirmed)"
        elif expected is not None:
            when, basis = expected, "carrier forecast of port discharge"
        elif eta is not None:
            when, basis = eta, "carrier ETA"
        else:
            continue

        away = nz.days_between(datetime.combine(when, datetime.min.time()), as_of)
        rows.append({**shipment,
                     "arrival_date": when,
                     "arrival_basis": basis,
                     "arrival_window": _window(away),
                     "days_away": nz.days_between(
                         datetime.combine(when, datetime.min.time()), as_of),
                     "arrival_port": (shipment["port_arrival_location"]
                                      or shipment["port_of_discharge_expected"]
                                      or shipment.get("port_of_discharge"))})
    rows.sort(key=lambda r: r["arrival_date"])
    return rows


def _window(days_away: int | None) -> str:
    """Bucket an arrival into something filterable. Negative means already ashore."""
    if days_away is None:
        return "Unknown"
    if days_away < 0:
        return "Already ashore"
    for horizon in ARRIVAL_HORIZONS:
        if days_away <= horizon:
            return "Next %d days" % horizon
    return "Beyond %d days" % ARRIVAL_HORIZONS[-1]


def port_rows(data) -> list[dict]:
    """Time spent at each port and hub, and what is sitting there now.

    Two kinds of location, deliberately labelled, because only one of them costs
    detention:

      Destination port   the box is ashore and collectable. Free time runs here.
      Transshipment hub  the box is between vessels and cannot be collected.
                         Real delay, no detention.
    """
    stats: dict[tuple, dict] = {}

    def bucket(location, kind):
        return stats.setdefault((location, kind),
                                {"location": location, "kind": kind,
                                 "dwells": [], "here_now": 0, "containers": 0,
                                 "longest_now": None, "companies": set()})

    for container in data.containers:
        company = next((s["company"] for s in data.shipment_rows
                        if s["bl"] == container.get("bl")), None)
        if container["port_arrival_location"]:
            entry = bucket(container["port_arrival_location"], "Destination port")
            entry["containers"] += 1
            if company:
                entry["companies"].add(company)
            if container["port_dwell_days"] is not None:
                entry["dwells"].append(container["port_dwell_days"])
            if container["at_port"]:
                entry["here_now"] += 1
                days = container["port_dwell_days"]
                if days is not None and (entry["longest_now"] is None
                                         or days > entry["longest_now"]):
                    entry["longest_now"] = days
        if container.get("transship_location"):
            entry = bucket(container["transship_location"], "Transshipment hub")
            entry["containers"] += 1
            if company:
                entry["companies"].add(company)
            entry["dwells"].append(container["transship_dwell_days"])
        if container.get("at_transshipment"):
            entry = bucket(container["transship_waiting_at"], "Transshipment hub")
            entry["here_now"] += 1
            # A hub whose only waits are still in progress has no completed dwell
            # to measure, but the boxes are unquestionably there. Counting them
            # keeps "Containers Seen" from reading 0 next to "Here Now" of 10.
            if not container.get("transship_dwell_days") or                     container.get("transship_location") != container["transship_waiting_at"]:
                entry["containers"] += 1
            days = container["transship_wait_days"]
            if days is not None and (entry["longest_now"] is None
                                     or days > entry["longest_now"]):
                entry["longest_now"] = days

    rows = []
    for entry in stats.values():
        dwells = entry["dwells"]
        rows.append({
            "location": entry["location"],
            "kind": entry["kind"],
            "containers": entry["containers"],
            "here_now": entry["here_now"],
            "longest_now": entry["longest_now"],
            "median_dwell": _median(dwells),
            "max_dwell": max(dwells) if dwells else None,
            "min_dwell": min(dwells) if dwells else None,
            "measured": len(dwells),
            "companies": ", ".join(sorted(entry["companies"])) or None,
        })
    rows.sort(key=lambda r: (-(r["median_dwell"] or 0), -r["containers"]))
    return rows


def lane_rows(data) -> list[dict]:
    """Transit performance per lane, slowest first."""
    rows = [dict(v) for v in data.lane_benchmarks.values()]
    rows.sort(key=lambda r: -(r["median"] or 0))
    return rows


def company_summary(data) -> list[dict]:
    """Per-company rollup. Companies come from the master workbook's tab names."""
    by_company: dict[str, list] = defaultdict(list)
    for shipment in data.shipment_rows:
        by_company[shipment.get("company") or "(not stated)"].append(shipment)

    rows = []
    for company, shipments in by_company.items():
        containers = sum(s["containers"] for s in shipments)
        at_port = sum(s["containers_at_port"] for s in shipments)
        expiring = [s for s in shipments
                    if s["containers_at_port"] and s["free_days_left"] is not None
                    and s["free_days_left"] <= FREE_TIME_WARN_DAYS]
        docs_pending = [s for s in shipments if s["docs_ready"] == "Pending"]
        transit = [s["transit_days"] for s in shipments
                   if s["transit_days"] is not None]
        rows.append({
            "company": company,
            "shipments": len(shipments),
            "containers": containers,
            "in_transit": sum(1 for s in shipments if s["status"] == "In Transit"),
            "delivered": sum(1 for s in shipments if s["status"] == "Delivered"),
            "at_port": at_port,
            "free_time_at_risk": len(expiring),
            "docs_pending": len(docs_pending),
            "needs_attention": sum(1 for s in shipments if s["attention"] != "OK"),
            "split": sum(1 for s in shipments if s["split"] == "Yes"),
            "carriers": ", ".join(sorted({s["carrier"] for s in shipments})),
            "commodities": len({s["commodity"] for s in shipments if s["commodity"]}),
            "median_transit_days": _median(transit),
        })
    rows.sort(key=lambda r: -r["containers"])
    return rows


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
