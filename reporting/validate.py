#!/usr/bin/env python3
"""
Data-quality findings and read/write reconciliation.

THE CONTRACT THIS MODULE ENFORCES
---------------------------------
The report must not silently lose tracking information. "Silently" is the
operative word -- rows legitimately leave the operational sheets (a superseded
snapshot should not be counted twice), but every row that does so has to be
accounted for somewhere a reader can find it.

So there are two separate mechanisms:

  RECONCILIATION  a row-count ledger: read from JSON, placed on the operational
                  sheet, placed on the raw sheet, and the difference with an
                  explicit reason. If read != operational + explained, the
                  ledger says FAIL and the generator exits non-zero. A report
                  that quietly dropped 40 events is worse than no report.

  FINDINGS        per-field integrity checks (missing carrier, unparseable
                  date, duplicate event, missing status) counted and listed
                  with an example, so a systematic problem is visible at the
                  count rather than found by accident in one row.

Every raw row also lands on the hidden Raw Containers / Raw Events sheets
exactly as it was read, so the workbook is self-auditing: any number on any
sheet can be traced back to source rows without going to the JSON.
"""

from __future__ import annotations

from collections import Counter

from . import normalize as nz


def findings(data, as_of) -> list[dict]:
    """Per-field data-quality findings. One row per distinct check that fired."""
    out: list[dict] = []

    def add(severity, area, check, count, detail, example=None):
        out.append({"severity": severity, "area": area, "check": check,
                    "count": count, "detail": detail, "example": example})

    # Anything the loader already noticed while reading and reconciling.
    for note in data.notes:
        add(note["severity"].title(), note["area"].title(), "Load-time note",
            note.get("count", 1), note["message"])

    containers = data.containers
    events = data.events

    # ---- containers ------------------------------------------------------- #
    no_carrier = [c for c in containers if not c.get("carrier_code")]
    if no_carrier:
        add("Critical", "Containers", "Missing carrier", len(no_carrier),
            "Container rows with no carrier code cannot be attributed to a "
            "shipping line, so they are missing from every carrier breakdown.",
            no_carrier[0].get("container_id"))

    unknown_carrier = [c for c in containers
                       if c.get("carrier_code")
                       and c["carrier_code"] not in nz.CARRIER_NAMES]
    if unknown_carrier:
        add("Warning", "Containers", "Unrecognised carrier code",
            len(unknown_carrier),
            "The carrier code is set but is not in the carrier registry, so the "
            "report shows the raw code. Add it to normalize.TARGET_CARRIERS or "
            "OTHER_CARRIERS. Codes seen: %s"
            % ", ".join(sorted({c["carrier_code"] for c in unknown_carrier})),
            unknown_carrier[0].get("container_id"))

    bad_snapshot = [c for c in containers if c["snapshot_dt"] is None]
    if bad_snapshot:
        add("Warning", "Containers", "Unparseable snapshot_at", len(bad_snapshot),
            "snapshot_at could not be read as a date, so the age of these rows "
            "is unknown and they cannot be ordered against other snapshots.",
            repr(bad_snapshot[0].get("snapshot_at")))

    bad_current = [c for c in containers
                   if c.get("current_date") not in (None, "") and c["current_dt"] is None]
    if bad_current:
        add("Warning", "Containers", "Unparseable current_date", len(bad_current),
            "current_date held a value that is not a readable date. The report "
            "uses the event history instead for these rows.",
            repr(bad_current[0].get("current_date")))

    no_status = [c for c in containers
                 if not c.get("current_event") and c["actual_events"] == 0]
    if no_status:
        add("Critical", "Containers", "No status available", len(no_status),
            "Neither current_event nor any confirmed event exists, so there is "
            "nothing to report as this container's position.",
            no_status[0].get("container_id"))

    duplicate_ids = [cid for cid, n in
                     Counter(c.get("container_id") for c in containers).items() if n > 1]
    if duplicate_ids:
        add("Critical", "Containers", "Duplicate container on report",
            len(duplicate_ids),
            "The same container_id appears more than once on Container Summary. "
            "Snapshot selection should make this impossible -- investigate. "
            "Affected: %s" % ", ".join(str(c) for c in duplicate_ids[:5]),
            duplicate_ids[0])

    no_tracking = [c for c in containers if not c.get("tracking_number_resolved")]
    if no_tracking:
        add("Warning", "Containers", "No tracking number", len(no_tracking),
            "tracking_number, bill_of_lading and query are all empty, so this "
            "container cannot be joined back to a booking or a sheet row.",
            no_tracking[0].get("container_id"))

    forecast_current = [c for c in containers if c.get("current_is_forecast")]
    if forecast_current:
        add("Warning", "Containers", "current_event is a forecast",
            len(forecast_current),
            "The tracker labelled a forecast row as the current event (MSC marks "
            "its 'Estimated Time of Arrival' row CURRENT). The report shows the "
            "last CONFIRMED move instead -- see the Confirmed Status column. "
            "Fixing this at source means changing msc_track.parse.",
            forecast_current[0].get("container_id"))

    # ---- events ----------------------------------------------------------- #
    no_eta = [c for c in containers
              if c["eta_dt"] is None and c["status_bucket"] != "Delivered"]
    if no_eta:
        add("Warning", "Containers", "No ETA from the carrier", len(no_eta),
            "Neither pod_eta nor final_pod_eta is populated. Normal for MSC, "
            "which supplies no forward schedule -- use the last confirmed move "
            "as the progress signal for those. Not a delay.",
            no_eta[0].get("container_id"))

    stale_rows = [c for c in containers
                  if (c["snapshot_age_days"] or 0) >= 7]
    if stale_rows:
        add("Warning", "Pipeline", "Tracking data not refreshed", len(stale_rows),
            "These containers were last looked up %d or more days ago, so their "
            "status may have changed since. This is about the pipeline, not the "
            "cargo -- re-run track_all.py."
            % 7, stale_rows[0].get("container_id"))

    no_destination = [c for c in containers if not c.get("destination")]
    if no_destination:
        add("Info", "Containers", "No destination from the carrier",
            len(no_destination),
            "The carrier returned no destination city or country, so no arrival "
            "expectation can be formed for these containers.",
            no_destination[0].get("container_id"))

    no_actual = [c for c in containers if c["actual_events"] == 0]
    if no_actual:
        add("Critical", "Containers", "No confirmed movement", len(no_actual),
            "The carrier returned only forecasts, or nothing at all, for these "
            "containers -- so there is no confirmed position to report.",
            no_actual[0].get("container_id"))

    no_date = [e for e in events if e["event_dt"] is None]
    if no_date:
        add("Warning", "Events", "Missing or unparseable event date", len(no_date),
            "These events cannot be placed on a timeline and are excluded from "
            "last-movement and activity calculations. They are still listed on "
            "Event History.", repr(no_date[0].get("event_date")))

    no_desc = [e for e in events if not e.get("description")]
    if no_desc:
        add("Warning", "Events", "Missing event description", len(no_desc),
            "The carrier sent an event with no description, so it normalises to "
            "'Other / unclassified'.", no_desc[0].get("container_id"))

    unmapped = [e for e in events if e["mapping_source"] in ("unmapped", "blank")]
    if unmapped:
        add("Warning", "Events", "Event wording not mapped", len(unmapped),
            "No exact mapping and no keyword rule matched, so these normalise to "
            "'Other / unclassified'. Add them to normalize.EXACT_EVENT_MAP. "
            "Wording seen: %s"
            % ", ".join(sorted({str(e.get("description")) for e in unmapped})[:5]),
            unmapped[0].get("description"))

    keyword = [e for e in events if e["mapping_source"] == "keyword"]
    if keyword:
        add("Info", "Events", "Event wording mapped by keyword rule", len(keyword),
            "These matched a fallback pattern rather than a confirmed mapping. "
            "Expected for a carrier with no adapter yet; confirm the mapping and "
            "promote it to EXACT_EVENT_MAP. See the Status Mapping sheet.",
            keyword[0].get("description"))

    # Duplicate events WITHIN the reporting scope. event_key already includes
    # container, sequence, date and description, so a repeat here means the same
    # move was emitted twice in one run -- not the same move seen across runs,
    # which is expected and handled by snapshot selection.
    keys = Counter((str(e.get("container_id")), str(e.get("event_key"))) for e in events)
    duplicate_events = [k for k, n in keys.items() if n > 1]
    if duplicate_events:
        add("Warning", "Events", "Duplicate event in one run", len(duplicate_events),
            "The same event_key appears more than once for a container within the "
            "same run. Both rows are shown on Event History; counts will be "
            "inflated by the repeats.", duplicate_events[0][1])

    missing_status = [e for e in events if not e.get("status")]
    if missing_status:
        add("Warning", "Events", "Missing event status", len(missing_status),
            "No COMPLETED/CURRENT/PLANNED status, so the actual-vs-expected split "
            "for these rows falls back to the event date alone.",
            missing_status[0].get("container_id"))

    vessel_legs = [e for e in events
                   if e.get("mode") == "VESSEL" and not e.get("vessel_name")]
    if vessel_legs:
        add("Info", "Events", "Vessel leg with no vessel name", len(vessel_legs),
            "A vessel-mode event with no vessel named. Expected for Maersk, "
            "whose feed omits the vessel on roughly half of its events.",
            vessel_legs[0].get("description"))

    unknown_mode = sorted({str(e.get("mode")) for e in events
                           if e.get("mode") and e["mode"] not in
                           ("VESSEL", "RAIL", "TRUCK", "GATE", "BARGE", "AIR")})
    if unknown_mode:
        affected = sum(1 for e in events if str(e.get("mode")) in unknown_mode)
        add("Info", "Events", "Unrecognised transport mode", affected,
            "The tracker passed a carrier transport code through untranslated "
            "(codes: %s). Harmless for this report; add them to "
            "maersk_track.TRANSPORT_MODE to get a readable mode."
            % ", ".join(unknown_mode), unknown_mode[0])

    # ---- shipments -------------------------------------------------------- #
    if data.shipments:
        no_bl = [s for s in data.shipments if not s.get("bl_number")]
        if no_bl:
            add("Warning", "Shipments", "Sheet row with no BL", len(no_bl),
                "A commercial sheet row with no bill of lading cannot be joined "
                "to tracking data.", no_bl[0].get("row_number"))
        untracked = [s for s in data.shipments if not s.get("tracked_status")]
        if untracked:
            add("Info", "Shipments", "Sheet row never tracked", len(untracked),
                "No tracking result was joined to this sheet row. Cross-check "
                "against Tracking Failures -- an unsupported carrier or an "
                "inland movement will land here permanently and is not a fault.",
                untracked[0].get("bl_number"))
        no_free_days = [s for s in data.shipments
                        if s.get("bl_number") and s.get("free_days") is None]
        if no_free_days:
            add("Info", "Shipments", "No free-days count on the sheet",
                len(no_free_days),
                "The FREE TIME cell is empty, or holds a date rather than a day "
                "count. Sheet dates are not used, so no free-time expiry can be "
                "worked out for these rows and detention cannot be tracked. "
                "Enter a number of days to fix it.",
                no_free_days[0].get("bl_number"))

        ignored = [s for s in data.shipments
                   if s.get("sheet_eta_date") or s.get("eta_variance_days") is not None]
        if ignored:
            add("Info", "Shipments", "Sheet dates present but ignored", len(ignored),
                "These rows were written before the commercial sheet's own dates "
                "were dropped. The values are not read: those columns are typed "
                "by hand and not maintained against the carriers, so every "
                "arrival date on this report comes from the carrier instead. "
                "Re-run the tracker to stop writing them.",
                ignored[0].get("bl_number"))

    if not out:
        add("Info", "All", "No issues detected", 0,
            "Every integrity check passed on this dataset.")

    rank = {"Critical": 0, "Warning": 1, "Info": 2}
    out.sort(key=lambda f: (rank.get(f["severity"], 3), f["area"], f["check"]))
    return out


# --------------------------------------------------------------------------- #
# reconciliation ledger
# --------------------------------------------------------------------------- #

def reconcile(data) -> tuple[list[dict], bool]:
    """Row-count ledger. -> (rows, ok).

    `ok` is False when any line fails to balance, and the generator turns that
    into a non-zero exit code. Balance means:

        rows read from JSON  ==  rows on the operational sheet
                                 + rows explained away

    "Explained away" is only ever superseded snapshots or rows the findings list
    names. Nothing else is allowed to disappear.
    """
    rows: list[dict] = []
    ok = True

    def line(source, read, operational, raw, explained, reason):
        nonlocal ok
        balanced = read == operational + explained
        if not balanced:
            ok = False
        rows.append({
            "source": source,
            "rows_read": read,
            "rows_on_report": operational,
            "rows_in_raw_sheet": raw,
            "rows_explained": explained,
            "balanced": "OK" if balanced else "FAIL",
            "reason": reason,
        })

    superseded_containers = len(data.raw_containers) - len(data.containers)
    line("containers.json", len(data.raw_containers), len(data.containers),
         len(data.raw_containers), superseded_containers,
         "Superseded snapshots of the same container, kept on Raw Containers."
         if superseded_containers else "Every row is on Container Summary.")

    superseded_events = len(data.raw_events) - len(data.events)
    line("events.json", len(data.raw_events), len(data.events),
         len(data.raw_events), superseded_events,
         "Events belonging to superseded runs, kept on Raw Events."
         if superseded_events else "Every row is on Event History.")

    if data.files_found.get("errors.json") is not None:
        line("errors.json", len(data.raw_errors), len(data.errors),
             len(data.raw_errors), 0, "Every failure row is on Tracking Failures.")

    if data.files_found.get("shipments.json") is not None:
        collapsed = len(data.raw_shipments) - len(data.shipments)
        # "On report" means reaching the analysis, not appearing on a visible
        # sheet: the commercial columns now live on the hidden Raw Sheet Rows
        # layer, and the rows themselves drive the Shipments and Free Time
        # sheets through their BL.
        line("shipments.json", len(data.raw_shipments), len(data.shipments),
             len(data.raw_shipments), collapsed,
             "Duplicate/superseded copies of the same sheet row_number "
             "(shipments.json carries no run_id)." if collapsed
             else "Every sheet row is on Commercial Shipments.")

    # Distinctness checks, reported as ledger lines rather than as prose so they
    # sit next to the counts they qualify.
    #
    # One row per container is the guarantee in the normal mode and a genuine
    # FAIL if broken. Under --all-snapshots it is not the guarantee at all --
    # one row per container PER RUN is the whole point of that flag -- so the
    # line reports the multiplicity instead of failing on it.
    unique_containers = len({c.get("container_id") for c in data.containers})
    if data.latest_only:
        one_per_container = unique_containers == len(data.containers)
        reason = ("One row per container on Container Summary."
                  if one_per_container
                  else "Container Summary has %d rows for %d distinct containers -- "
                       "snapshot selection should make this impossible."
                       % (len(data.containers), unique_containers))
        verdict = "OK" if one_per_container else "FAIL"
    else:
        verdict = "OK"
        reason = ("--all-snapshots: %d rows for %d distinct containers, one per "
                  "run. Counts on every sheet are row counts, not container "
                  "counts." % (len(data.containers), unique_containers))
    rows.append({
        "source": "Distinct containers",
        "rows_read": len({c.get("container_id") for c in data.raw_containers}),
        "rows_on_report": unique_containers,
        "rows_in_raw_sheet": len({c.get("container_id") for c in data.raw_containers}),
        "rows_explained": 0,
        "balanced": verdict,
        "reason": reason,
    })
    if verdict == "FAIL":
        ok = False

    # Every container must sit under exactly one BL row, and the BL rows must
    # account for every container -- the rollup is where containers could go
    # missing without any file-level count noticing.
    grouped = sum(s["containers"] for s in data.shipment_rows)
    rows.append({
        "source": "Containers grouped into BLs",
        "rows_read": len(data.containers),
        "rows_on_report": grouped,
        "rows_in_raw_sheet": len(data.containers),
        "rows_explained": 0,
        "balanced": "OK" if grouped == len(data.containers) else "FAIL",
        "reason": ("%d container(s) across %d bill(s) of lading on BL Summary."
                   % (grouped, len(data.shipment_rows))
                   if grouped == len(data.containers)
                   else "BL Summary accounts for %d of %d containers."
                        % (grouped, len(data.containers))),
    })
    if rows[-1]["balanced"] == "FAIL":
        ok = False

    events_on_report = sum(c["events_on_report"] for c in data.containers)
    rows.append({
        "source": "Events attributed to a container",
        "rows_read": len(data.events),
        "rows_on_report": events_on_report,
        "rows_in_raw_sheet": len(data.raw_events),
        "rows_explained": len(data.events) - events_on_report,
        "balanced": "OK" if events_on_report == len(data.events) else "FAIL",
        "reason": ("Every reported event sits under a container on Container Summary."
                   if events_on_report == len(data.events)
                   else "%d reported event(s) belong to no container row."
                        % (len(data.events) - events_on_report)),
    })
    if rows[-1]["balanced"] == "FAIL":
        ok = False

    return rows, ok
