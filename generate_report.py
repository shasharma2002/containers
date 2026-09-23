#!/usr/bin/env python3
"""
Build the cargo tracking Excel report from the tracking application's JSON.

    python generate_report.py
    python generate_report.py --input-dir out
    python generate_report.py --output reports/september.xlsx
    python generate_report.py --all-snapshots      # history, not latest-only

Reads containers.json / events.json (required) and errors.json /
shipments.json (used when present) and writes ONE workbook to
output/cargo_tracking_report.xlsx.

This is a read-only consumer of the tracking output. It does not import
track_all.py, msc_track.py, maersk_track.py or maersk_browser.py, does not
touch the network, and never writes to the tracking JSON -- so it cannot
affect a tracking run. Run it as often as you like.

EXIT CODES
----------
    0   workbook written and the row-count ledger balanced
    1   workbook written but the ledger did NOT balance -- the report is
        missing data it could not account for. Open the Data Quality sheet.
    2   could not run at all (no input found, or the file could not be saved)

The distinction matters for scheduling: a non-zero exit from a nightly job
should page someone, and "the report silently lost 400 events" deserves that
just as much as "the input was missing".
"""

from __future__ import annotations

import argparse
import sys
import traceback
from datetime import datetime
from pathlib import Path

from reporting import analyze, excel, loader, normalize, validate

DEFAULT_OUTPUT = Path("output") / "cargo_tracking_report.xlsx"


def log(message: str = "") -> None:
    print(message, file=sys.stderr)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Generate the cargo tracking Excel report from tracking JSON.")
    ap.add_argument("--input-dir", type=Path,
                    help="folder holding containers.json and events.json "
                         "(default: '.', then 'out')")
    ap.add_argument("--output", type=Path, default=DEFAULT_OUTPUT,
                    help="workbook path (default: %s)" % DEFAULT_OUTPUT)
    ap.add_argument("--all-snapshots", action="store_true",
                    help="report every appended run instead of the latest row "
                         "per container. Totals become row counts, not "
                         "container counts -- for trend analysis, not for "
                         "sending to anyone.")
    ap.add_argument("--as-of", type=str,
                    help="treat this ISO date/time as 'now' when ageing rows "
                         "(default: current UTC time). Useful for reproducing "
                         "an earlier report.")
    ap.add_argument("-q", "--quiet", action="store_true")
    args = ap.parse_args()

    def say(message: str = "") -> None:
        if not args.quiet:
            log(message)

    # ---- read -------------------------------------------------------------- #
    try:
        data = loader.load(args.input_dir, latest_only=not args.all_snapshots)
    except loader.LoaderError as exc:
        log("FAILED: %s" % exc)
        return 2

    say("\nReading from %s" % data.input_dir.resolve())
    for name, count in data.files_found.items():
        say("  %-18s %s" % (name, "not present" if count is None
                            else "%d row(s)" % count))

    as_of = None
    if args.as_of:
        as_of = normalize.parse_dt(args.as_of)
        if as_of is None:
            log("FAILED: --as-of %r is not a readable date." % args.as_of)
            return 2

    # ---- analyse ----------------------------------------------------------- #
    as_of = analyze.enrich(data, as_of)
    kpi = analyze.kpis(data, as_of)
    dist = analyze.distributions(data, as_of)
    carriers = analyze.carrier_summary(data)
    quality = validate.findings(data, as_of)
    ledger, balanced = validate.reconcile(data)

    say("\nReporting as of %s UTC" % as_of.strftime("%Y-%m-%d %H:%M"))
    say("  %-28s %d" % ("containers", kpi["total_containers"]))
    say("  %-28s %d" % ("events", kpi["total_events"]))
    say("  %-28s %d in transit, %d delivered, %d at origin"
        % ("status", kpi["in_transit"], kpi["delivered"], kpi["at_origin"]))
    say("  %-28s %d (%d critical, %d warning)"
        % ("needing attention", kpi["needs_attention"], kpi["critical"],
           kpi["warning"]))
    say("  %-28s %d delayed, %d ETA passed, %d not moved %dd+"
        % ("exceptions", kpi["delayed"], kpi["overdue"], kpi["stale"],
           analyze.STALE_WARN_DAYS))
    say("  %-28s %d of %d target carriers"
        % ("carriers with data", kpi["target_carriers_with_data"],
           kpi["target_carriers"]))

    # ---- write ------------------------------------------------------------- #
    try:
        written = excel.build(data, kpi, dist, carriers, quality, ledger, args.output)
    except PermissionError:
        log("\nFAILED: %s is open in Excel. Close it and run again."
            % args.output)
        return 2
    except Exception:
        log("\nFAILED while writing the workbook:")
        traceback.print_exc()
        return 2

    say("\nReconciliation")
    for row in ledger:
        say("  %-34s %-5s read=%-7d report=%-7d explained=%d"
            % (row["source"], row["balanced"], row["rows_read"],
               row["rows_on_report"], row["rows_explained"]))

    critical = [f for f in quality if f["severity"] == "Critical"]
    if critical:
        say("\n%d critical data-quality finding(s) -- see the Data Quality sheet:"
            % len(critical))
        for finding in critical[:5]:
            say("  %s: %s (%d row(s))"
                % (finding["area"], finding["check"], finding["count"]))

    say("\nWrote %s  (%.0f KB)"
        % (written.resolve(), written.stat().st_size / 1024))

    if not balanced:
        log("\nWARNING: the row-count ledger did not balance. The workbook was "
            "written, but it is missing rows it could not account for. Open the "
            "Data Quality sheet.")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
