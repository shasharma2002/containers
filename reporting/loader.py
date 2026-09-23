#!/usr/bin/env python3
"""
Find, read and reconcile the tracking JSON.

THE APPEND PROBLEM -- READ THIS FIRST
-------------------------------------
track_all.py --append writes history into the SAME files. The containers.json
in this project holds 767 rows that are really 116 containers across 9 runs.
Counting rows instead of containers overstates every number by roughly the
number of runs you have ever done.

powerbi_setup.md section 6 already settled the correct semantics for this, and
this module reproduces them exactly so the Excel report and the Power BI model
never disagree:

    the reporting row for a container is the row from that container's OWN
    latest snapshot -- not the globally latest run

Per-container rather than global, so a container that failed to refresh in the
most recent run still shows its last good result instead of vanishing from the
report. `snapshot_age_days` on every container row says how stale that
particular row is, so a silently-aging container is visible rather than
presented as current.

Events are then taken from THE SAME run_id as the container row they belong to.
Mixing runs would let a container show 14 events from run A next to a current
status from run B.

WHAT "SCHEMA v1" MEANS HERE
---------------------------
pipeline.flatten() is an older copy of track_all.flatten() that predates the
`tracking_number` column. maersk_browser.py and maersk_burst.py import from
pipeline, so running either of those standalone writes v1 rows while
track_all.py writes v2. Both shapes are loaded; a v1 row gets
tracking_number derived the same way v2 derives it (bill_of_lading or query),
and the substitution is reported on the Data Quality sheet rather than being
done silently.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

from . import normalize as nz

# Filenames the tracking application writes. `required` is False for the two
# that only some run modes produce -- shipments.json is sheet-mode only, and
# errors.json is written only when something actually failed.
INPUT_FILES = {
    "containers": {"name": "containers.json", "required": True},
    "events": {"name": "events.json", "required": True},
    "errors": {"name": "errors.json", "required": False},
    "shipments": {"name": "shipments.json", "required": False},
}

# Directories searched, in order, when --input-dir is not given. "." first
# because that is where track_all.py writes by default; "out" second because
# that is the --outdir the project's own examples use.
DEFAULT_SEARCH_DIRS = (".", "out")


class LoaderError(RuntimeError):
    pass


@dataclass
class Dataset:
    """Everything the report needs, already reconciled.

    `notes` collects anything worth telling the reader about how the data was
    interpreted. It feeds the Data Quality sheet -- nothing is dropped or
    substituted without leaving a line here.
    """

    input_dir: Path
    raw_containers: list[dict] = field(default_factory=list)
    raw_events: list[dict] = field(default_factory=list)
    raw_errors: list[dict] = field(default_factory=list)
    raw_shipments: list[dict] = field(default_factory=list)

    containers: list[dict] = field(default_factory=list)
    events: list[dict] = field(default_factory=list)
    errors: list[dict] = field(default_factory=list)
    shipments: list[dict] = field(default_factory=list)

    runs: list[dict] = field(default_factory=list)
    eta_history: dict = field(default_factory=dict)
    notes: list[dict] = field(default_factory=list)
    files_found: dict = field(default_factory=dict)
    latest_only: bool = True

    def note(self, severity: str, area: str, message: str, count: int = 1) -> None:
        self.notes.append({"severity": severity, "area": area,
                           "message": message, "count": count})


# --------------------------------------------------------------------------- #
# file discovery
# --------------------------------------------------------------------------- #

def find_input_dir(explicit: Path | None) -> Path:
    """Pick the directory holding the tracking JSON.

    An explicit --input-dir is used as given and must contain the required
    files. Otherwise the default search order is tried, and the first directory
    holding BOTH containers.json and events.json wins.
    """
    if explicit is not None:
        path = Path(explicit)
        if not path.is_dir():
            raise LoaderError("--input-dir %s is not a directory." % path)
        missing = [spec["name"] for spec in INPUT_FILES.values()
                   if spec["required"] and not (path / spec["name"]).exists()]
        if missing:
            raise LoaderError(
                "%s has no %s. Point --input-dir at the folder track_all.py "
                "wrote to." % (path, " or ".join(missing)))
        return path

    for candidate in DEFAULT_SEARCH_DIRS:
        path = Path(candidate)
        if all((path / spec["name"]).exists()
               for spec in INPUT_FILES.values() if spec["required"]):
            return path

    raise LoaderError(
        "Found no containers.json + events.json in %s. Run track_all.py first, "
        "or pass --input-dir." % " or ".join(repr(d) for d in DEFAULT_SEARCH_DIRS))


def _read_json_list(path: Path, data: Dataset) -> list[dict]:
    """Read a JSON array of objects. Bad rows are reported, never dropped quietly."""
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        data.note("error", "input", "%s could not be read: %s" % (path.name, exc))
        return []
    if not text.strip():
        data.note("warning", "input", "%s is empty." % path.name)
        return []
    try:
        loaded = json.loads(text)
    except json.JSONDecodeError as exc:
        data.note("error", "input",
                  "%s is not valid JSON (%s) -- no rows read from it." % (path.name, exc))
        return []
    if isinstance(loaded, dict):
        loaded = [loaded]
    if not isinstance(loaded, list):
        data.note("error", "input",
                  "%s holds a %s, not a list of rows." % (path.name, type(loaded).__name__))
        return []
    rows = [r for r in loaded if isinstance(r, dict)]
    if len(rows) != len(loaded):
        data.note("error", "input",
                  "%s: %d entries were not objects and could not be used."
                  % (path.name, len(loaded) - len(rows)), len(loaded) - len(rows))
    return rows


# --------------------------------------------------------------------------- #
# snapshot selection
# --------------------------------------------------------------------------- #

def _tracking_number(row: dict) -> str | None:
    """v2 writes tracking_number; v1 does not. Derive it the same way v2 does."""
    value = row.get("tracking_number")
    if value not in (None, ""):
        return str(value)
    for key in ("bill_of_lading", "query"):
        if row.get(key) not in (None, ""):
            return str(row[key])
    return None


def _select_latest(rows: list[dict], data: Dataset) -> tuple[list[dict], set[tuple[str, str]]]:
    """Latest row per container_id, plus the (container_id, run_id) pairs kept.

    The pairs are what the event selection joins on, so events always come from
    the same run as the container row above them.
    """
    best: dict[str, dict] = {}
    unkeyed = 0
    for row in rows:
        cid = row.get("container_id")
        if cid in (None, ""):
            unkeyed += 1
            continue
        cid = str(cid)
        current = best.get(cid)
        if current is None:
            best[cid] = row
            continue
        # Compare on the parsed snapshot so a v1/v2 offset difference cannot
        # flip the ordering. Rows with no parseable snapshot lose to rows that
        # have one, and otherwise the later-appearing row wins (append order
        # is chronological).
        new_dt = nz.parse_dt(row.get("snapshot_at"))
        old_dt = nz.parse_dt(current.get("snapshot_at"))
        if old_dt is None and new_dt is not None:
            best[cid] = row
        elif new_dt is not None and old_dt is not None and new_dt >= old_dt:
            best[cid] = row
        elif new_dt is None and old_dt is None:
            best[cid] = row

    if unkeyed:
        data.note("error", "containers",
                  "%d container row(s) had no container_id and could not be "
                  "placed on the report. They remain on Raw Containers." % unkeyed,
                  unkeyed)

    kept = list(best.values())
    pairs = {(str(r.get("container_id")), str(r.get("run_id"))) for r in kept}
    return kept, pairs


def _eta_history(rows: list[dict]) -> dict[str, dict]:
    """Per container, how pod_eta moved across the appended snapshots.

    This is the one genuine, auditable delay signal in the data: the carrier's
    own ETA for the same container, compared between the first snapshot that
    carried one and the latest. It is NOT a guess about lateness -- it is the
    carrier revising its own promise, which is why it is safe to report as a
    delay. Containers tracked only once simply have no history and get None.
    """
    by_container: dict[str, list[tuple]] = {}
    for row in rows:
        cid = row.get("container_id")
        if cid in (None, ""):
            continue
        snap = nz.parse_dt(row.get("snapshot_at"))
        eta = nz.parse_dt(row.get("pod_eta")) or nz.parse_dt(row.get("final_pod_eta"))
        by_container.setdefault(str(cid), []).append((snap, eta))

    history: dict[str, dict] = {}
    for cid, points in by_container.items():
        points.sort(key=lambda p: (p[0] is None, p[0]))
        etas = [(snap, eta) for snap, eta in points if eta is not None]
        entry = {
            "snapshots": len(points),
            "eta_observations": len(etas),
            "first_eta": etas[0][1] if etas else None,
            "first_eta_seen_at": etas[0][0] if etas else None,
            "previous_eta": etas[-2][1] if len(etas) > 1 else None,
            "latest_eta": etas[-1][1] if etas else None,
            "eta_slip_days": None,
            "eta_slip_since_previous": None,
            "eta_changes": 0,
        }
        if len(etas) > 1:
            entry["eta_slip_days"] = nz.days_between(etas[-1][1], etas[0][1])
            entry["eta_slip_since_previous"] = nz.days_between(etas[-1][1], etas[-2][1])
            entry["eta_changes"] = len({e.date() for _, e in etas}) - 1
        history[cid] = entry
    return history


def _runs(containers: list[dict], events: list[dict]) -> list[dict]:
    """One row per run_id -- what the appended history actually contains."""
    runs: dict[str, dict] = {}
    for row in containers:
        rid = str(row.get("run_id") or "(none)")
        entry = runs.setdefault(rid, {"run_id": rid, "snapshot_at": None,
                                      "containers": 0, "events": 0,
                                      "carriers": set(), "schema": set()})
        entry["containers"] += 1
        entry["carriers"].add(nz.carrier_name(row.get("carrier")))
        entry["schema"].add("v2" if "tracking_number" in row else "v1")
        snap = nz.parse_dt(row.get("snapshot_at"))
        if snap and (entry["snapshot_at"] is None or snap > entry["snapshot_at"]):
            entry["snapshot_at"] = snap
    for row in events:
        rid = str(row.get("run_id") or "(none)")
        if rid in runs:
            runs[rid]["events"] += 1

    out = []
    for entry in runs.values():
        out.append({
            "run_id": entry["run_id"],
            "snapshot_at": entry["snapshot_at"],
            "containers": entry["containers"],
            "events": entry["events"],
            "carriers": ", ".join(sorted(entry["carriers"])),
            "schema": "/".join(sorted(entry["schema"])),
        })
    out.sort(key=lambda r: (r["snapshot_at"] is None, r["snapshot_at"] or ""), reverse=True)
    return out


# --------------------------------------------------------------------------- #
# shipments
# --------------------------------------------------------------------------- #

def _dedupe_shipments(rows: list[dict], data: Dataset) -> list[dict]:
    """shipments.json has no run_id, so --append leaves the same sheet row over.

    The file in this project holds 8 passes over the same 17 sheet rows: 79
    byte-identical copies plus 40 rows that differ because the tracked ETA
    moved between runs.

    `row_number` is the sheet line the row came from, so it is the natural
    identity. Picking the CURRENT copy of each one depends on what the file
    carries:

      snapshot_at present -- the newest snapshot wins, the same rule
                             containers.json uses. Authoritative.
      no snapshot_at      -- fall back to APPEND ORDER being chronological,
                             since write_json() concatenates new rows after
                             existing ones. That is an inference, so it is
                             reported on the Data Quality sheet rather than
                             being made quietly.

    Older files predate the run stamp (it was added to track_all.build_shipments
    after this reporting layer was built), so both paths have to work.

    Rows with no row_number at all cannot be identified this way, so they are
    deduplicated on full-field equality instead and all survivors are kept.
    """
    latest_by_row: dict = {}
    order: list = []
    identical = superseded = 0
    unkeyed: dict[str, dict] = {}
    stamped = sum(1 for r in rows if r.get("snapshot_at"))

    for row in rows:
        key = row.get("row_number")
        if key in (None, ""):
            fingerprint = json.dumps(row, sort_keys=True, default=str)
            if fingerprint in unkeyed:
                identical += 1
            else:
                unkeyed[fingerprint] = row
            continue

        if key not in latest_by_row:
            order.append(key)
            latest_by_row[key] = row
            continue

        incumbent = latest_by_row[key]
        if (json.dumps(incumbent, sort_keys=True, default=str)
                == json.dumps(row, sort_keys=True, default=str)):
            identical += 1
            continue

        superseded += 1
        new_dt = nz.parse_dt(row.get("snapshot_at"))
        old_dt = nz.parse_dt(incumbent.get("snapshot_at"))
        if new_dt is not None and old_dt is not None:
            # Both stamped: the later snapshot is the current one.
            if new_dt >= old_dt:
                latest_by_row[key] = row
        else:
            # Unstamped on at least one side -- fall back to append order.
            latest_by_row[key] = row

    kept = [latest_by_row[k] for k in order] + list(unkeyed.values())

    if identical or superseded:
        if stamped == len(rows) and rows:
            basis = ("Collapsed to the newest snapshot_at per sheet row_number.")
            severity = "info"
        else:
            basis = ("These rows carry no snapshot_at, so the report fell back "
                     "to append order being chronological and kept the LAST "
                     "occurrence of each sheet row_number. Re-run the tracker "
                     "to get run-stamped rows.")
            severity = "warning"
        data.note(severity, "shipments",
                  "shipments.json accumulates one copy of each sheet row per "
                  "run with --append. %s %d row(s) reported, %d identical "
                  "copies and %d superseded older versions dropped."
                  % (basis, len(kept), identical, superseded),
                  identical + superseded)
    if unkeyed:
        data.note("warning", "shipments",
                  "%d shipment row(s) have no row_number, so the newest version "
                  "of them cannot be identified. All distinct versions are shown."
                  % len(unkeyed), len(unkeyed))
    return kept


# --------------------------------------------------------------------------- #
# entry point
# --------------------------------------------------------------------------- #

def load(input_dir: Path | None = None, latest_only: bool = True) -> Dataset:
    """Read the tracking JSON and return a reconciled Dataset."""
    resolved = find_input_dir(input_dir)
    data = Dataset(input_dir=resolved, latest_only=latest_only)

    for key, spec in INPUT_FILES.items():
        path = resolved / spec["name"]
        if not path.exists():
            data.files_found[spec["name"]] = None
            if spec["required"]:
                raise LoaderError("%s not found in %s." % (spec["name"], resolved))
            data.note("info", "input",
                      "%s not present -- the sheets that depend on it are "
                      "omitted. This is normal for a container-list run."
                      % spec["name"])
            continue
        rows = _read_json_list(path, data)
        data.files_found[spec["name"]] = len(rows)
        setattr(data, "raw_" + key, rows)

    if not data.raw_containers:
        raise LoaderError(
            "containers.json in %s holds no usable rows -- nothing to report on."
            % resolved)

    # --- schema version ---------------------------------------------------- #
    v1 = [r for r in data.raw_containers if "tracking_number" not in r]
    if v1:
        data.note("warning", "schema",
                  "%d container row(s) predate the tracking_number column "
                  "(written by pipeline.flatten via maersk_browser.py or "
                  "maersk_burst.py). tracking_number was derived from "
                  "bill_of_lading or query for those rows." % len(v1), len(v1))
    null_tn = [r for r in data.raw_containers
               if "tracking_number" in r and r.get("tracking_number") in (None, "")]
    if null_tn:
        data.note("warning", "schema",
                  "%d container row(s) carry tracking_number = null. Derived "
                  "from bill_of_lading or query instead." % len(null_tn), len(null_tn))

    # --- history before collapsing ----------------------------------------- #
    data.eta_history = _eta_history(data.raw_containers)
    data.runs = _runs(data.raw_containers, data.raw_events)

    # --- snapshot selection ------------------------------------------------ #
    if latest_only:
        selected, pairs = _select_latest(data.raw_containers, data)
        if len(data.raw_containers) > len(selected):
            data.note("info", "snapshot",
                      "containers.json holds %d row(s) across %d run(s); the "
                      "report shows the latest row per container -- %d "
                      "container(s). Full history stays on Raw Containers."
                      % (len(data.raw_containers), len(data.runs), len(selected)))
        selected_events = [e for e in data.raw_events
                           if (str(e.get("container_id")), str(e.get("run_id"))) in pairs]
    else:
        selected = list(data.raw_containers)
        selected_events = list(data.raw_events)
        data.note("info", "snapshot",
                  "--all-snapshots was used: every appended run is on the "
                  "report, so a container tracked N times contributes N rows. "
                  "Totals are row counts, not container counts.")

    orphan_events = len(data.raw_events) - len(selected_events)
    if latest_only and orphan_events:
        data.note("info", "snapshot",
                  "%d event row(s) belong to superseded runs and are not on "
                  "Event History. They remain on Raw Events." % orphan_events,
                  orphan_events)

    # Events whose container has no container row at all: a real inconsistency,
    # not just an older snapshot.
    known = {str(r.get("container_id")) for r in data.raw_containers}
    orphaned = [e for e in data.raw_events if str(e.get("container_id")) not in known]
    if orphaned:
        data.note("error", "events",
                  "%d event row(s) reference a container_id that appears in no "
                  "container row. They are on Raw Events but cannot be placed "
                  "under a container." % len(orphaned), len(orphaned))

    snapshot_by_run = {r["run_id"]: r["snapshot_at"] for r in data.runs}

    # --- derive ------------------------------------------------------------ #
    data.containers = [_derive_container(r, snapshot_by_run, data) for r in selected]
    data.events = [nz.derive_event(r, snapshot_by_run.get(str(r.get("run_id"))))
                   for r in selected_events]
    data.errors = [_derive_error(r) for r in data.raw_errors]
    data.shipments = [_derive_shipment(r)
                      for r in _dedupe_shipments(data.raw_shipments, data)]
    return data


def _derive_container(row: dict, snapshot_by_run: dict, data: Dataset) -> dict:
    code = row.get("carrier")
    snapshot = (nz.parse_dt(row.get("snapshot_at"))
                or snapshot_by_run.get(str(row.get("run_id"))))
    return {
        **row,
        "carrier_code": code,
        "carrier_display": nz.carrier_name(code, row.get("carrier_name")),
        "tracking_number_resolved": _tracking_number(row),
        "snapshot_dt": snapshot,
        "current_dt": nz.parse_dt(row.get("current_date")),
        "pod_eta_dt": nz.parse_dt(row.get("pod_eta")),
        "final_pod_eta_dt": nz.parse_dt(row.get("final_pod_eta")),
        "origin": nz.location_text(row.get("origin_city"), row.get("origin_country")),
        "destination": nz.location_text(row.get("destination_city"),
                                        row.get("destination_country")),
    }


def _derive_error(row: dict) -> dict:
    """errors.json row -> report row, with the failure kind pulled out.

    track_all.py encodes the kind as a prefix on `error`: "unsupported_carrier",
    "no_ocean_carrier", or a carrier-by-carrier reason list for a genuine
    not-found. Those three need completely different responses, so the report
    separates them rather than showing one "failed" bucket.
    """
    text = str(row.get("error") or "")
    if text.startswith("unsupported_carrier"):
        kind = "Unsupported carrier"
        action = "No adapter exists for this line. Track manually or build an adapter."
    elif text.startswith("no_ocean_carrier"):
        kind = "No ocean carrier"
        action = ("Inland/truck movement -- no bill of lading is ever issued. "
                  "Nothing to track; exclude from pending work.")
    elif not text:
        kind = "Unclassified"
        action = "No reason recorded by the tracker."
    else:
        kind = "Not found"
        action = "Check the number, or the carrier's response in the reason column."

    code = nz.carrier_code(row.get("classified_as"))
    return {
        **row,
        "failure_kind": kind,
        "suggested_action": action,
        "carrier_code": code,
        "carrier_display": nz.carrier_name(code, row.get("classified_as")),
        "snapshot_dt": nz.parse_dt(row.get("snapshot_at")),
    }


def _derive_shipment(row: dict) -> dict:
    code = nz.carrier_code(row.get("carrier_name"))
    return {
        **row,
        "carrier_code": code,
        "carrier_display": nz.carrier_name(code, row.get("carrier_name")),
        "sheet_eta_dt": nz.parse_dt(row.get("sheet_eta_date")),
        "tracked_eta_dt": nz.parse_dt(row.get("tracked_eta")),
    }
