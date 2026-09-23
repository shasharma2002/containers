# Power BI setup

Paste-ready M and DAX. Skips the click-through and, more importantly, sets the
data types correctly — the JSON connector imports every column as text, so
`event_date` sorts alphabetically and date maths silently fails.

Four files now come out of `track_all.py`: `containers.json`, `events.json`,
`errors.json`, and (only when tracking a commercial sheet by Bill of Lading)
`shipments.json`. This doc covers all four and how they relate.

---

## 0. Which column joins to which — read this first

Every container/event row carries **three** different "what is this shipment"
columns, and they mean different things:

| Column | Meaning |
|---|---|
| `query` | Whatever string was searched with — a container number, a BL, or a booking number. |
| `bill_of_lading` | The carrier's own confirmed BL number. **MSC returns this. Maersk's API never does** — it's `null` on every single Maersk row, no exceptions. |
| `tracking_number` | `bill_of_lading` if the carrier gave one, else `query`. **Always populated. Use this one.** |

If you build a relationship on `bill_of_lading`, every Maersk container silently
drops out of it — Power BI won't error, the numbers will just quietly be wrong.
`tracking_number` is the column meant for grouping and joining; keep `query`/
`bill_of_lading` around only if you need to audit what was searched vs. what the
carrier confirmed.

**The relationship map:**

```
Shipments[bl_number]  (commercial sheet, one row per line item)
        │  1
        │
        ▼  *
Containers[tracking_number]  (one row per container per snapshot)
        │  1
        │
        ▼  *
Events[container_id]  (one row per event, ever)
```

Containers → Events joins on `container_id` (unchanged from before). Shipments
→ Containers joins on `bl_number` = `tracking_number` — see §5 for why that
needs a small bridge table rather than a direct relationship.

---

## 1. Parameter for the folder

Home → Transform data → Home → Manage Parameters → New:

| | |
|---|---|
| Name | `DataFolder` |
| Type | Text |
| Current Value | `C:\tracking\out` |

Set it once; every query below follows. Moving the files later is a one-field edit.

---

## 2. Containers query

New Source → Blank Query → Advanced Editor → replace everything:

```m
let
    Source     = Json.Document(File.Contents(DataFolder & "\containers.json")),
    ToTable    = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Expand     = Table.ExpandRecordColumn(ToTable, "Column1", {
                    "run_id", "snapshot_at", "carrier", "carrier_name", "query",
                    "bill_of_lading", "tracking_number", "container_id",
                    "container_type", "delivered", "latest_move", "pod_eta",
                    "final_pod_eta", "origin_city", "origin_country",
                    "destination_city", "destination_country", "port_of_load",
                    "port_of_discharge", "transshipment_count", "transshipments",
                    "event_count", "actual_event_count", "has_forecast",
                    "current_event", "current_date", "current_city",
                    "current_country", "current_unlocode", "current_vessel",
                    "current_voyage"}),
    Types      = Table.TransformColumnTypes(Expand, {
                    {"snapshot_at", type datetimezone},
                    {"pod_eta", type datetime},
                    {"final_pod_eta", type datetime},
                    {"current_date", type datetime},
                    {"delivered", type logical},
                    {"has_forecast", type logical},
                    {"transshipment_count", Int64.Type},
                    {"event_count", Int64.Type},
                    {"actual_event_count", Int64.Type}
                 })
in
    Types
```

Rename the query to `Containers`.

`carrier_name` is the display-friendly version of `carrier` ("Maersk (incl.
Hamburg Sud)" vs `MAEU`) — use it on visuals, keep `carrier` for filtering/DAX.

---

## 3. Events query

Same again, rename to `Events`:

```m
let
    Source  = Json.Document(File.Contents(DataFolder & "\events.json")),
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Expand  = Table.ExpandRecordColumn(ToTable, "Column1", {
                 "run_id", "snapshot_at", "carrier", "container_id",
                 "bill_of_lading", "tracking_number", "event_key", "sequence",
                 "status", "event_date", "description", "city", "country",
                 "unlocode", "terminal", "terminal_code", "mode", "vessel_name",
                 "voyage", "vessel_imo", "vessel_flag", "laden_status"}),
    Types   = Table.TransformColumnTypes(Expand, {
                 {"snapshot_at", type datetimezone},
                 {"event_date", type datetime},
                 {"sequence", Int64.Type}
              }),
    Dedupe  = Table.Distinct(Types, {"event_key"})
in
    Dedupe
```

The `Dedupe` step matters once you start using `--append`. Historical events never
change, so re-fetching them every run would multiply your event count by the number of
runs. Deduping on `event_key` keeps exactly one row per real-world event while
`Containers` still gets one row per snapshot — which is what you want, because the
container's *position* is the thing that changes.

---

## 4. Shipments query (only if you're tracking by BL from a commercial sheet)

`shipments.json` only exists after a sheet-mode run (`track_all.py "your sheet.xlsx"`).
Skip this if you're only tracking container lists.

```m
let
    Source  = Json.Document(File.Contents(DataFolder & "\shipments.json")),
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Expand  = Table.ExpandRecordColumn(ToTable, "Column1", {
                 "row_number", "bl_number", "carrier_name", "beneficiary",
                 "broker", "commodity", "pi_no", "qty", "amount",
                 "delivery_terms", "free_days", "sheet_eta_date",
                 "sheet_eta_place", "tracked_eta", "tracked_status",
                 "tracked_location", "eta_variance_days", "container_count",
                 "containers"}),
    Types   = Table.TransformColumnTypes(Expand, {
                 {"row_number", Int64.Type},
                 {"qty", type number},
                 {"amount", type number},
                 {"free_days", Int64.Type},
                 {"sheet_eta_date", type date},
                 {"tracked_eta", type datetime},
                 {"eta_variance_days", Int64.Type},
                 {"container_count", Int64.Type}
              })
in
    Types
```

Rename to `Shipments`. `eta_variance_days` is `tracked_eta − sheet_eta_date`:
positive means the carrier is now reporting a later ETA than your commercial
team recorded. It's blank whenever either side is missing — MSC gives no
forward ETA on most rows, so expect blanks there; that's the carrier, not a
bug.

**Also load `errors.json`** if you want failures visible in the report itself
rather than just in the terminal — same pattern, one field list:

```m
let
    Source  = Json.Document(File.Contents(DataFolder & "\errors.json")),
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Expand  = Table.ExpandRecordColumn(ToTable, "Column1", {
                 "run_id", "snapshot_at", "query", "classified_as",
                 "confidence", "carriers_tried", "error"}),
    Types   = Table.TransformColumnTypes(Expand, {{"snapshot_at", type datetimezone}})
in
    Types
```

Rename to `Errors`. Two `error` prefixes are not failures and shouldn't be
colored the same as a real "not found":
- `unsupported_carrier:` (e.g. HAPAG) — a real shipment with no adapter yet.
- `no_ocean_carrier:` — inland/truck cargo with no BL to track at all (seen on
  this sheet: Bangladesh/India deliveries by road). There's no carrier API to
  call for a truck, so this will never resolve and isn't meant to.

Only a bare "not found" (no such prefix) means the pipeline actually tried and
failed — that's the one worth chasing.

Close & Apply.

---

## 5. Relationships

**Containers → Events** (unchanged): Model view → drag `Containers[container_id]`
onto `Events[container_id]`.
- Cardinality: **One to many**
- Cross-filter direction: **Single**

If Power BI refuses "one to many" here, you're appending and `Containers` has
repeat `container_id` values across snapshots — apply the `IsLatest` filter in
§6 first, or accept many-to-many.

**Shipments → Containers** needs one extra step, because `bl_number` is **not
unique** in `Shipments` — a BL with multiple commodity lines (continuation
rows) appears once per line, e.g. `MEDUY5025638` shows up three times for
three different products under the same BL. Power BI can't use a non-unique
column as the "one" side of a relationship, so build a tiny bridge table first:

New Source → Blank Query → Advanced Editor:

```m
let
    Source = Table.Distinct(Shipments, {"bl_number"}),
    Kept   = Table.SelectColumns(Source, {"bl_number"})
in
    Kept
```

Rename to `BL`. Then:
- `BL[bl_number]` **1** → `Shipments[bl_number]` **\***, single direction
- `BL[bl_number]` **1** → `Containers[tracking_number]` **\***, single direction

Now a slicer on `Shipments[beneficiary]` or `Shipments[commodity]` correctly
filters `Containers` and `Events` down through `BL`, and a container's
tracking status correctly rolls up to every commodity line sharing its BL —
without Power BI ever seeing a many-to-many relationship.

---

## 6. Latest-snapshot column

**Do this before building any visual if you plan to use `--append`.** Otherwise
every count is multiplied by the number of runs.

Table view → `Containers` → New column:

```dax
IsLatest =
VAR LatestForThisContainer =
    CALCULATE(
        MAX( Containers[snapshot_at] ),
        ALLEXCEPT( Containers, Containers[container_id] )
    )
RETURN
    Containers[snapshot_at] = LatestForThisContainer
```

Per-container rather than global, so a container that failed in the most recent
run still shows its last good result instead of vanishing.

Drop `IsLatest` into the Filters pane as a page-level filter set to `True`.

---

## 7. Measures

New measure on `Containers`:

```dax
Containers Tracked = DISTINCTCOUNT( Containers[container_id] )

In Transit =
CALCULATE( [Containers Tracked], Containers[delivered] = FALSE )

Delivered =
CALCULATE( [Containers Tracked], Containers[delivered] = TRUE )

Days Since Last Move =
VAR LastMove = MAX( Containers[current_date] )
RETURN IF( NOT ISBLANK( LastMove ), DATEDIFF( LastMove, TODAY(), DAY ) )

Stalled 7d+ =
CALCULATE(
    [Containers Tracked],
    FILTER( VALUES( Containers[container_id] ), [Days Since Last Move] >= 7 )
)

Transit Days So Far =
VAR FirstEvent = CALCULATE( MIN( Events[event_date] ) )
VAR LastEvent  = CALCULATE( MAX( Events[event_date] ) )
RETURN DATEDIFF( FirstEvent, LastEvent, DAY )
```

New measure on `Shipments` (only if you loaded it — this is the one your
manager actually wants):

```dax
Avg ETA Variance (days) = AVERAGE( Shipments[eta_variance_days] )

Shipments Slipping =
CALCULATE(
    COUNTROWS( Shipments ),
    Shipments[eta_variance_days] > 0
)

Shipments Improving =
CALCULATE(
    COUNTROWS( Shipments ),
    Shipments[eta_variance_days] < 0
)
```

`Stalled 7d+` and `Shipments Slipping` are the two that earn their keep — see below.

---

## 8. Visuals worth building first

**Exception table** (not a map — start with what needs action):

Table visual, `Containers`, columns:
`container_id`, `current_city`, `current_event`, `current_date`,
`Days Since Last Move`, `destination_city`

Sort descending by `Days Since Last Move`, conditional-format that column red
above 7. That's your morning check.

**Cards:** `Containers Tracked`, `In Transit`, `Stalled 7d+`

**ETA variance table** (if you loaded `Shipments`): Table visual, `Shipments`,
columns: `bl_number`, `beneficiary`, `commodity`, `sheet_eta_date`,
`tracked_eta`, `eta_variance_days`. Conditional-format `eta_variance_days`
red when positive (carrier running later than the commercial team expected).
Blank rows are BLs with no `tracked_eta` yet — mostly MSC, some genuinely
not-yet-found (cross-check against `Errors`).

**Journey timeline:** Line/scatter, `Events[event_date]` on X, `Events[city]` on
Y, legend `Events[vessel_name]`, filtered to one container via a slicer.

**Port dwell:** Matrix, rows `Events[city]`, values `Transit Days So Far` —
shows where boxes actually lose time.

---

## 9. Things worth knowing about this data

**One BL can mean many containers, automatically.** `track_all.py` dedupes BLs
before fetching — a BL that covers 16 containers is one lookup, not 16, and
all 16 come back in `Containers` sharing the same `tracking_number`. You don't
need to do anything for this; it's why the bridge table in §5 exists.

**MSC gives actuals only — no forward ETA on most rows.** `pod_eta` and
`final_pod_eta` (and therefore `Shipments[tracked_eta]`/`eta_variance_days`)
are frequently blank for MSC shipments. That's the carrier, not missing data
on your end — `Days Since Last Move` is the delay signal that works
regardless of carrier.

**`unsupported_carrier` in `Errors` is not a failure.** It means the sheet has
a real shipment on a line with no adapter (currently: Hapag-Lloyd). It will
never resolve on its own — it needs either a Hapag adapter built, or manual
tracking outside this pipeline.

**A `Shipments` row with everything blank after `sheet_eta_place` is usually
inland/truck cargo, not a data problem.** Deliveries arriving overland from
Bangladesh/India never get a Bill of Lading — that's an ocean-freight
document, and there's no carrier API to call for a truck. `track_all.py`
recognises this (no `SHIPPING LINE` and no recognisable BL shape) and records
it as `no_ocean_carrier` in `Errors` rather than attempting a lookup. These
rows are done, permanently — filter them out of any "needs tracking" visual
rather than treating them as pending.
