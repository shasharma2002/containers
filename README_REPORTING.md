# Excel reporting layer

Turns the JSON that `track_all.py` writes into one professional workbook you can
send to the people who need it.

```
track_all.py  ->  containers.json / events.json / errors.json / shipments.json
                                      |
                              generate_report.py
                                      |
                      output/cargo_tracking_report.xlsx
```

Email integration is not part of this — the generator only writes the file.

---

## Running it

```bash
python generate_report.py
```

Reads the JSON from the current directory (then `out/` if not found there) and
writes `output/cargo_tracking_report.xlsx`.

| Option | What it does |
|---|---|
`--input-dir DIR` | Read the JSON from `DIR` instead of searching |
`--output PATH` | Write somewhere other than `output/cargo_tracking_report.xlsx` |
`--all-snapshots` | Report every appended run instead of the latest row per container. Totals become row counts, not container counts — for trend analysis, not for sending to anyone |
`--as-of ISO` | Treat that date/time as "now" when ageing rows, to reproduce an earlier report |
`-q` | Suppress the console summary |

**Exit codes.** `0` = written and the row-count ledger balanced. `1` = written
but the ledger did **not** balance, meaning the report is missing rows it could
not account for — open the Data Quality sheet. `2` = could not run (no input, or
the file could not be saved; the usual cause is the workbook being open in
Excel). The `1` case exists so a scheduled job treats "silently lost 400 events"
as the failure it is.

No new dependencies: the generator uses `openpyxl`, which this project already
requires.

---

## What it does not touch

The reporting layer imports nothing from the tracking application — not
`track_all.py`, not the carrier adapters, not `maersk_browser.py`. It makes no
network calls and never writes to the tracking JSON. Running it cannot affect a
tracking run, so it is safe to run mid-run or on a schedule.

---

## The workbook

| Sheet | For | Contents |
|---|---|---|
**Dashboard** | Management | 12 KPI tiles, 5 charts, links to every sheet |
**Shipments** | Operations | **One row per BL** — the sheet to open first. Worst first |
**Action Required** | Operations | Only things a person can act on |
**Container Summary** | Operations | Drill-down to individual boxes |
**Event History** | Operations | Full movement record, carrier wording + normalised |
**Free Time & Detention** | Operations | Detention exposure |
**Company Summary** | Management | Per-company rollup |
**Carrier Summary** | Both | Per-carrier rollup and real coverage |
**Ports & Hubs** | Both | Where time is lost |
**Lane Performance** | Management | Real transit times |
**Tracking Failures** | Operations | What could not be looked up, and why |
**Data Quality** | Both | Row-count ledger + every data problem, counted once |

Hidden (unhide from the tab strip): **Status Mapping**, **Carrier Mapping**,
**Run History**, **Raw Sheet Rows**, **Raw Containers**, **Raw Events**.

### What was removed, and why

| Removed | Reason |
|---|---|
**Arrivals Planner** sheet | The same rows as Shipments, re-sorted. Shipments now sorts worst-first and carries Carrier ETA + Days To ETA, so Excel's own sort does the rest |
**Commercial Shipments** sheet | Its tracked columns now sit on Shipments; the rest moved to the hidden **Raw Sheet Rows**, so nothing is lost |
**Cargo value** (qty × rate) | Rested on reading `RATE/AMT` as a rate per tonne. Nobody asked for the figure and it was an assumption, so it went |
**Data-class findings** on the action sheet | They duplicated checks `validate.py` already did per container. Every data problem is now reported **once**, on Data Quality |
**Tracking failures** on the action sheet | They have their own sheet with the carrier's reason and what to do |
32 columns on Shipments | 50 → 18. A sheet you cannot read at a glance is not a summary |
22 columns on Container Summary | 35 → 13 |
11 columns on Event History | 23 → 12. The rest is on Raw Events |
12 KPI tiles, 9 charts | 24 → 12 tiles, 14 → 5 charts |

One file, not several. The raw sheets are hidden rather than split out, because a
report you cannot trace back to source rows is not auditable, and a second file
is a file that gets separated from the first one in an email.

---

## Dates come from the carriers, never from the sheet

The commercial workbook's date columns are typed by hand and are not kept in
step with the carriers, so **none of them is used**. Every arrival date, ETA and
free-time expiry on the report is derived from carrier events.

| Sheet column | Status |
|---|---|
`ETA BRJ`, `ETA Viz/Kol` | **Not used.** Raw text shown on Commercial Shipments, marked `NOT USED` |
`LC DATE`, `BL DATE` | **Not used** |
`FREE TIME` holding a date | **Not used** — no expiry is claimed for that row |
`FREE TIME` holding a day count (`14`) | **Used.** A contractual term, not a date |
`DOCUMENT POSITION` | **Wording used** (`ok` vs `PENDING`); any date inside the cell ignored |

Consequences worth knowing:

- **There is no ETA-variance metric.** Comparing a carrier ETA against a date
  nobody maintains produced a number that looked meaningful and was not, so it
  is gone rather than left to mislead.
- **Free time still works**, because the clock starts on the carrier's discharge
  event. The `How the expiry was worked out` column on every row reads
  *"carrier discharge YYYY-MM-DD + N free day(s) from the sheet"*, so the inputs
  are visible per row.
- **A shipment with no carrier date does not appear on the Arrivals Planner.**
  It is left off rather than given a hand-typed date — the planner's `Basis`
  column only ever says *arrived (confirmed)*, *carrier forecast of port
  discharge*, or *carrier ETA*.
- Older `shipments.json` files written before this change still contain
  `sheet_eta_date` and `eta_variance_days`. They are **ignored**, and the Data
  Quality sheet says so and tells you to re-run the tracker.

To put the sheet's dates back in play, set `TRUST_SHEET_DATES = True` in
[read_master.py](read_master.py) — `SHEET_DATE_FIELDS` beside it lists exactly
what that re-enables.

---

## What the report computes beyond "where is it"

Four things that are not in the JSON and have to be derived. All four come from
real events plus the commercial sheet — none is an assumption.

### Free time and detention

The biggest avoidable cost in this data. Once a box is discharged at the
**destination** port, the consignee has N free days; after that the carrier
charges per container per day, and a 23-container BL multiplies that fast.

```
free time expires = destination-port discharge  (a CARRIER event)
                  + free days from the sheet    (a day COUNT, not a date)
```

The sheet's own expiry *date* is **not** honoured — see the section above. Its
free-days *count* is, because a contractual term does not go stale the way a
typed date does.

Two rules stop this from crying wolf:

- **Only containers still at the port are exposed.** A BL whose free time ran out
  three weeks ago but whose boxes all railed out owes nothing. `Free Time Status`
  therefore leads with the count still on the ground, not a signed number that
  cannot tell those apart.
- **No free days on the sheet means no expiry is claimed.** The row says so and
  raises a *separate* finding, rather than assuming an industry-standard 14.
  This also covers the rows whose `FREE TIME` cell holds a date instead of a
  count, since that date is not trusted.

### Destination port vs transshipment hub

This one is subtle and it was wrong in my first cut. **Maersk fires the same
`DISCHARG` activity at a transshipment hub as at the final port** — only MSC
labels them distinctly. Real rows from your data: BL `274146535` discharges at
**Colombo** and BL `274659478` at **Tanjung Pelepas**, both hubs, both reported
exactly like a final discharge.

So a discharge counts as the destination only when **no vessel leg follows it in
the itinerary — forecast included**, since Colombo's onward departure is still an
`EXPECTED` event while the box sits there. Get this wrong and the free-time clock
starts weeks early at Colombo, inventing detention that does not exist.

Hub dwell is still measured — it is real delay — just never charged as detention.
`Ports & Hubs` labels every location as one kind or the other.

### Journey stages

| Column | Means |
|---|---|
`Ocean Days` | loaded on vessel → discharged at destination port |
`Port Dwell (days)` | discharged → moved inland. **Detention territory** |
`Hub Wait (days)` | discharged at a hub → reloaded. Delay, not detention |
`Transit So Far (days)` | origin gate-in → arrival at destination |

### Lane benchmarks

Transit time per lane from **your own completed journeys**, so it already
includes the transshipment waits and rail legs these routes really involve — not
a carrier's published port-to-port time. A lane needs ≥3 completed journeys
before it is used to flag anything, and a shipment is only called slow at ≥25%
over its own lane's median.

---

## The four things worth understanding

### 1. The bill of lading is the primary grain

Cargo is booked, tracked and chased by BL, not by box. In this project's data
**116 containers are 29 BLs**, and one BL carries up to **23 containers** that
travel as a single consignment — same vessel, same voyage, same ETA, scanned
minutes apart at the same terminal. Listing those 23 identically moving boxes as
23 lines buries the only thing worth reading: the BL.

So **BL Summary** is the first sheet and each row is one consignment, showing the
**consensus** position — the value most of its containers share:

- **Consensus is the modal value**, not the latest and not the worst. On BL
  `271560680`, 12 boxes are delivered at Birganj and 1 is still in transit at
  Visakhapatnam. The shipment reads **Delivered**, and the straggler is named in
  *Containers Not With The Rest*. Taking the latest or the worst would describe a
  whole consignment by its one outlier.
- **`Split? = Yes`** means at least one container differs from the rest on
  status, milestone, location or vessel. That is your damaged / rolled / held
  box, and it is the case worth a phone call.
- **Container Summary is still there** as the drill-down, now with a BL column.

**Findings roll up to the BL.** A 20-container consignment past its ETA produces
*one* line reading "All 20 containers", not 20 identical lines. When only part is
affected the row reads "1 of 13 containers" and names them.

**What is deliberately *not* treated as divergence** — these were the two real
false positives found while building this:

| Signal | Why it's ignored |
|---|---|
Different **timestamps** | A terminal working down a 23-box stack produces a dozen timestamps minutes apart. `274146535` shows exactly this: 12 distinct times, spread **0 days**. Tested for as a spread in *days*, never equality |
Different **ETA dates** | Carriers issue a per-container delivery date, so a 12-box BL routinely carries ETAs a day or two apart. Equality-testing ETA flagged 6 of 12 boxes on `270383678` when nothing was wrong. Now a spread, flagged only at 3+ days |
Staggered **empty returns** on a delivered BL | The consignee unloading 10 boxes over a week is unloading cadence, not boxes left behind. The "separated in time" finding is suppressed once every container is delivered |

### 2. The JSON is append-mode history

`track_all.py --append` writes into the *same* files. This project's
`containers.json` holds **767 rows that are really 116 containers across 9
runs**. Counting rows would overstate every number eightfold.

The report shows **the latest row per container** — not the globally latest run.
Per-container, so a container that failed to refresh in the most recent run still
shows its last good result instead of vanishing. This is the same rule as
`IsLatest` in `powerbi_setup.md` §6, deliberately, so the workbook and the Power
BI model never disagree.

Events come from the *same run* as the container row above them. Superseded rows
stay on the hidden raw sheets and are named in the reconciliation ledger.

### 3. What counts as a delay

Only two things, and both are things the carrier itself told us:

- **ETA slipped** — the carrier moved its own ETA later between two snapshots.
  Requires at least two snapshots carrying an ETA, which is why `--append` is
  what makes delay detection possible at all.
- **ETA passed** — the carrier's latest ETA is in the past and no delivery or
  final arrival has been reported.

Deliberately *not* called delays:

- **Sitting still** is staleness, not lateness — and it is only reported where
  standing still is actually odd: **on the ground at a destination port, or
  between vessels at a hub.** A box in open water is not scanned for weeks on
  these routes, so flagging that produced 7 findings out of 8 that nobody could
  act on. Removing that assumption is what shrank the action list most.
- **A missing ETA** is a data gap. MSC supplies no forward schedule on most rows,
  so treating "no ETA" as "late" would mark most MSC shipments delayed because of
  an API design choice.

Every Exceptions row carries a `Class` — `Shipment` (call the carrier) or `Data`
(fix the pipeline) — and an evidence column naming the source field.

### 4. Three kinds of "stale"

| Column | Means |
|---|---|
`Idle Days When Last Checked` | How long the box had gone unscanned when the carrier was last queried. **A fact about the cargo** — this is what the no-movement finding uses |
`Tracking Data Age (days)` | How long ago the tracker last ran. A fact about your refresh cadence |
`Days Since Last Move (today)` | The two added together — what an ops user sees asking "when did this last move?" |

Keeping them apart is not cosmetic. Run the generator against the 3-week-old
`out/` snapshot and "days since last move" exceeds the critical threshold for
*every* container, which would report the whole fleet as stalled when the only
real fact is that the pipeline has not run. Measured at the snapshot, the same
dataset flags 60 of 93 — and the other 33 get a "tracking data not refreshed"
finding, which is the problem that actually needs fixing.

---

## Code layout

```
generate_report.py        CLI entry point
reporting/
    loader.py             find + read the JSON, pick the reporting snapshot
    normalize.py          carrier registry, event vocabulary, date parsing
    analyze.py            KPIs, carrier rollup, exception detection
    validate.py           integrity findings + the reconciliation ledger
    excel.py              the workbook writer
test_report.py            test suite (real data + synthetic fixtures)
```

---

## How to extend it

### Add a new carrier

Two separate jobs — the reporting layer and the tracking layer.

**Reporting** (enough to make the carrier's data appear correctly):

1. Add it to `TARGET_CARRIERS` in [reporting/normalize.py](reporting/normalize.py):
   ```python
   {"code": "COSU", "name": "COSCO", "adapter": "HTTP API"},
   ```
   Set `adapter` to `None` until a tracker adapter exists — the Carrier Mapping
   sheet reads this field and will say "No adapter built", which is what you want
   it to say while that is true.
2. Add the spellings it appears under to `CARRIER_ALIASES` (the commercial
   sheet's `SHIPPING LINE` column and `errors.json` `classified_as` both feed
   through this).
3. Run the generator. Its event wording will be resolved by keyword fallback and
   flagged `keyword` on Status Mapping. Check those rows, then move the confirmed
   wording into `EXACT_EVENT_MAP` (next section).

Nothing hard-codes the eight carriers into a formula: carriers on Carrier Summary
are derived from the data, then the eight are unioned in so one with no rows shows
as absent rather than being invisible.

**Tracking** (so the carrier can actually be fetched) — see
"Changes the tracking application needs" below.

### Add an event mapping

1. See what is unmapped:
   ```bash
   python -m reporting.normalize --vocabulary events.json
   ```
   This lists every `(carrier, wording)` pair in the file with its mapping and
   the source (`exact` / `keyword` / `unmapped`).
2. Add the wording to `EXACT_EVENT_MAP` in
   [reporting/normalize.py](reporting/normalize.py), pointing at a key in
   `MILESTONES`. Use the string **exactly** as it arrives — `Discharg` is
   genuinely truncated in Maersk's feed and the mapping matches that, not a
   tidied-up version.
3. Need a milestone that does not exist? Add it to `MILESTONES` as
   `code: (label, phase, is_forecast)`. Set `is_forecast=True` only for events
   that are predictions, since that drives `time_basis`.

The original wording is never replaced — both columns appear on Event History.

### Change a dashboard metric

- **Thresholds** — the constants at the top of
  [reporting/analyze.py](reporting/analyze.py) (`STALE_WARN_DAYS`,
  `ETA_SLIP_WARN_DAYS`, `CONTAINER_SPREAD_WARN_DAYS`, `ETA_SPREAD_WARN_DAYS`,
  `FREE_TIME_WARN_DAYS`, `LANE_MIN_JOURNEYS`, `DOCS_DEADLINE_DAYS`, …).
  They are interpolated into the sheet legends, so the report explains itself
  after you change them.
- **What counts as containers diverging on a BL** — `DIVERGENCE_FIELDS` in
  [reporting/analyze.py](reporting/analyze.py). Only add a field here if any
  difference in it is genuinely a problem; anything with natural per-container
  variation (a timestamp, a per-box delivery date) belongs as a *spread*
  threshold instead, or it will fire on every multi-container BL you have.
- **A KPI tile** — add a key in `analyze.kpis()`, then a `(label, value, note,
  alert)` entry in the `tiles` list in `excel._dashboard()`. Tiles flow four per
  band automatically.
- **A chart** — add an entry to `analyze.distributions()`, then a line to the
  `sections` list in `excel._dashboard()`. Each section is a table plus a chart
  of the same numbers, with the question it answers printed above it. Keep that
  discipline: if you cannot write the question, the chart does not belong.
- **A column** — add one `Column(header, key, width, fmt)` line to the relevant
  list in [reporting/excel.py](reporting/excel.py). Widths, formats, filters and
  freeze panes are handled by `_table()`.

---

## Testing

```bash
python test_report.py
```

Covers two clearly separated things:

- **Real data** — the project's own JSON, both the 9-run appended history and the
  single-run `out/` directory. This genuinely tests **Maersk and MSC only**,
  because those are the only carriers that exist in the data.
- **Synthetic fixtures** — written in the tracking application's own JSON shape,
  covering **COSCO, CMA CGM, Hapag-Lloyd, ONE, Evergreen and Yang Ming**, plus:
  empty results, orphan events, failures-only, all-null fields, unparseable
  dates, duplicate events, a repeated container across runs, MSC's
  forecast-labelled-as-current case, missing optional inputs, schema v1 rows,
  malformed JSON, and a 2,000-container load.
- **BL grain** — a uniform 20-box consignment collapsing to one row; 23 boxes
  scanned minutes apart NOT flagged; one stray box among 13 detected and named;
  consensus following the 12 rather than the straggler; containers genuinely 8
  days apart flagged; staggered empty returns on a delivered BL not flagged; ETA
  1 day apart ignored and 6 days apart flagged; 20 overdue containers rolling up
  to one finding; the commercial sheet joined onto the BL.

**A synthetic pass proves the report handles that carrier's data. It does not
prove the tracker can fetch it** — no adapter exists for those six, so the
tracking path for them is untested and cannot be tested from here. The suite
prints this distinction at the end rather than reporting eight carriers as
"tested".

---

## Known limitations

- **Six of the eight carriers have no tracking adapter.** The report handles
  their data and lists them as absent with a reason, but the tracking application
  cannot fetch COSCO, CMA CGM, Hapag-Lloyd, ONE, Evergreen or Yang Ming today.
- **Delay detection needs `--append`.** ETA slip is a diff across snapshots. With
  a single run there is no history, so the report reports no slips — correctly,
  rather than guessing. `out/` demonstrates this: 0 delayed, because one run
  cannot show a revision.
- **Multi-company row numbering.** Each company tab in the master workbook
  restarts its row numbers, so a shipment row is identified by
  **(company, row_number)**. Rows from the older `bl_list.xlsx`, which states no
  company, fall back to `row_number` alone — correct for a single-sheet
  workbook.
- **`shipments.json` run stamps.** Rows written by current code carry
  `run_id`/`snapshot_at`, and the newest snapshot per (company, row) wins. Files
  written before that was added have no stamp, and the report falls back to
  append order being chronological — flagged as a warning on the Data Quality
  sheet, not assumed silently.
- **Maersk rows never carry a bill of lading.** That endpoint does not return
  one, so the Bill of Lading column is always blank for Maersk. Join on
  `Tracking No. / BL` instead.
- **Maersk supplies no vessel IMO or flag,** and no vessel name on roughly half
  its events. MSC supplies IMO and flag but leaves `laden_status` blank on most
  vessel legs. Both are carrier feed limitations, reported as `Info` on Data
  Quality rather than as faults.
- **`mode` can contain a raw carrier code.** `FEF` appears on 232 reported events
  — a Maersk `transport_mode` not in `maersk_track.TRANSPORT_MODE`. Harmless
  here; adding it there makes the Mode column readable.
- **Charts are static.** openpyxl writes chart definitions, not pivot caches, so
  the charts do not respond to the table filters. Filter a table and read the
  numbers; the charts show the whole dataset.
- **Cargo value is derived, not stated.** The master sheet holds a RATE per
  tonne, so value is quantity × rate. Treat it as an indication of exposure, not
  an invoice figure — and it is blank wherever either input is missing.
- **Detention rates are not in the data**, so the report says how many containers
  are exposed and for how long, never a currency amount. Add a per-diem per
  carrier to the sheet and that becomes computable.
- **Free time is assumed to start at discharge.** Some carriers start it at
  gate-in or on a combined demurrage/detention clock. If yours differ, the basis
  column makes it obvious what was assumed on every row.
- **The commercial sheet's ETA columns are unused**, so there is no comparison
  between what the commercial team expected and what the carrier says. If those
  columns are ever maintained properly, flip `TRUST_SHEET_DATES` and the
  variance metric can come back.
- **Lane benchmarks need finished journeys.** A lane appears only once cargo has
  completed it, so this sheet fills out over weeks rather than being useful on
  day one.
- **No trend sheet.** The history is used for ETA slip, but the workbook does not
  chart a container's position over time. `--all-snapshots` plus a pivot table is
  the workaround.

---

## Changes the tracking application needs

None of these are required for the report to work, and none have been made. They
are recorded here because each one would make the reported data better.

| # | Change | Why |
|---|---|---|
1 | Add Yang Ming prefixes (`YMLU`, `YMMU`) to `carrier_id.CARRIER_PREFIXES` | One of the eight target carriers cannot be classified at all today — a Yang Ming container falls through to "unknown prefix" and gets tried against every adapter |
2 | Build adapters for COSCO, CMA CGM, Hapag-Lloyd, ONE, Evergreen, Yang Ming | The only way those six ever appear in the report with real data. `ADAPTERS` in `track_all.py` is the extension point; each needs a `fetch`/`parse` pair returning the existing parsed shape |
3 | Delete `pipeline.flatten` and have `maersk_browser.py` / `maersk_burst.py` import from `track_all.py` | `pipeline.flatten` is an older copy that predates `tracking_number`, so those two entry points write schema v1 while `track_all.py` writes v2. The report handles both and says so, but one schema is better than two |
4 | Stop `msc_track.parse` labelling a forecast row `CURRENT` | MSC's last event is often `Estimated Time of Arrival` — a prediction. Labelling it `CURRENT` makes `current_event`/`current_date` a future date. The report corrects this with `time_basis`, but Power BI reading the same JSON does not |
5 | Add `run_id` / `snapshot_at` to `build_shipments` rows | Removes the append-order assumption described above |
6 | Add `FEF` to `maersk_track.TRANSPORT_MODE` | Turns a raw carrier code into a readable mode on 232 events |

Items 1 and 2 are the ones that matter for the eight-carrier requirement: **the
reporting layer supports all eight today; the tracking layer supports two.**
