#!/usr/bin/env python3
"""
The workbook writer.

WHY openpyxl
------------
It is already a dependency of this project (carrier_id.read_excel and
read_sheet.get_shipments both use it), so the report adds nothing to install.
It covers every requirement here: Excel tables with filters, freeze panes,
conditional formatting, number and date formats, bar/pie/line charts, and
internal hyperlinks. xlsxwriter has a nicer chart API and writes faster, but
adding a second Excel library to save a few lines on ~100 containers is not a
trade worth making.

ONE FILE, NOT SEVERAL
---------------------
Everything goes in one workbook, including the raw layer. The raw sheets are
hidden rather than split out, because a report that cannot be traced back to
its source rows is not auditable -- and a second file is a file that gets
separated from the first one in an email.

HOW A SHEET IS DEFINED
----------------------
Every tabular sheet is a list of Column(header, key, width, fmt). `_table()`
does the rest: writes the header, writes rows in order, sets widths, freezes
the header, wraps it in a real Excel table so filters and banding come for
free, and returns the row range so conditional formatting can be applied on
top. Adding a column to a sheet means adding one line to its column list.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table, TableStyleInfo

from . import analyze

# --------------------------------------------------------------------------- #
# house style
# --------------------------------------------------------------------------- #

NAVY = "1F3864"
BLUE = "2E75B6"
LIGHT = "DDEBF7"
GREY = "F2F2F2"
BORDER_GREY = "BFBFBF"

RED_FILL = PatternFill("solid", fgColor="FFC7CE")
RED_FONT = Font(name="Calibri", size=10, color="9C0006")
AMBER_FILL = PatternFill("solid", fgColor="FFEB9C")
AMBER_FONT = Font(name="Calibri", size=10, color="9C6500")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
GREEN_FONT = Font(name="Calibri", size=10, color="006100")

TITLE_FONT = Font(name="Calibri", size=18, bold=True, color=NAVY)
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
SECTION_FONT = Font(name="Calibri", size=12, bold=True, color=NAVY)
HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
BODY_FONT = Font(name="Calibri", size=10)
BOLD_FONT = Font(name="Calibri", size=10, bold=True)
LINK_FONT = Font(name="Calibri", size=10, color="0563C1", underline="single")

TILE_LABEL_FONT = Font(name="Calibri", size=9, bold=True, color="404040")
TILE_VALUE_FONT = Font(name="Calibri", size=20, bold=True, color=NAVY)
TILE_NOTE_FONT = Font(name="Calibri", size=8, italic=True, color="595959")
TILE_FILL = PatternFill("solid", fgColor=GREY)
TILE_ALERT_FILL = PatternFill("solid", fgColor="FCE4E4")

THIN = Side(style="thin", color=BORDER_GREY)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

DATE_FMT = "yyyy-mm-dd"
DATETIME_FMT = "yyyy-mm-dd hh:mm"
INT_FMT = "#,##0"
DEC_FMT = "0.0"
MONEY_FMT = "#,##0.00"

# Sheet names, in workbook order. Kept in one place because the Dashboard's
# navigation links and the hyperlinks between sheets both reference them.
S_DASH = "Dashboard"
S_BLS = "Shipments"
S_FREETIME = "Free Time & Detention"
S_COMPANIES = "Company Summary"
S_PORTS = "Ports & Hubs"
S_LANES = "Lane Performance"
S_CONTAINERS = "Container Summary"
S_EVENTS = "Event History"
S_CARRIERS = "Carrier Summary"
S_EXCEPTIONS = "Action Required"
S_FAILURES = "Tracking Failures"
S_QUALITY = "Data Quality"
S_STATUS_MAP = "Status Mapping"
S_CARRIER_MAP = "Carrier Mapping"
S_RUNS = "Run History"
S_RAW_CONTAINERS = "Raw Containers"
S_RAW_EVENTS = "Raw Events"
S_RAW_SHIPMENTS = "Raw Sheet Rows"


@dataclass
class Column:
    header: str
    key: str
    width: int = 14
    fmt: str | None = None


# --------------------------------------------------------------------------- #
# primitives
# --------------------------------------------------------------------------- #

def _title(ws, text: str, subtitle: str, span: int) -> None:
    ws["A1"] = text
    ws["A1"].font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    ws["A2"] = subtitle
    ws["A2"].font = SUBTITLE_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
    ws.row_dimensions[1].height = 24


def _link(ws, row: int, column: int, text: str, sheet: str) -> None:
    """An internal link to another sheet.

    Set via `location` rather than by assigning a "#'Sheet'!A1" string. The
    string form makes openpyxl emit an EXTERNAL relationship per link, which
    some mail gateways strip and some viewers refuse; `location` produces the
    canonical internal hyperlink with no relationship at all. This workbook is
    meant to be emailed, so that difference matters.
    """
    cell = ws.cell(row=row, column=column, value=text)
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location="'%s'!A1" % sheet)
    cell.font = LINK_FONT


def _nav(ws, row: int, targets: list[str], span: int) -> int:
    """Internal links to every sheet, wrapping so none are dropped.

    Each label needs two columns to read without being clipped, so a 12-column
    dashboard fits five per row and the rest wrap onto the next line.
    """
    ws.cell(row=row, column=1, value="Go to:").font = BOLD_FONT
    column = 3
    for name in targets:
        if column + 1 > span:
            row += 1
            column = 3
        _link(ws, row, column, name, name)
        column += 2
    return row + 1


def _cell_value(value):
    """Anything -> something Excel can store, without losing information."""
    if value is None or isinstance(value, (str, int, float, bool, datetime, date)):
        return value
    if isinstance(value, (list, tuple, set)):
        return ", ".join(str(v) for v in value) or None
    return str(value)


def _table(ws, columns: list[Column], rows: list[dict], name: str,
           start_row: int = 4, freeze: bool = True) -> tuple[int, int]:
    """Write a sheet body as a real Excel table. -> (first data row, last row).

    A table gives filters, banding and a structured reference for free. When
    there are no rows an Excel table is invalid, so a single explanatory line
    is written instead -- an empty sheet that says why is better than a sheet
    that fails to open.
    """
    for index, column in enumerate(columns, start=1):
        cell = ws.cell(row=start_row, column=index, value=column.header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BOX
        ws.column_dimensions[get_column_letter(index)].width = column.width
    ws.row_dimensions[start_row].height = 28

    for offset, row in enumerate(rows, start=1):
        for index, column in enumerate(columns, start=1):
            cell = ws.cell(row=start_row + offset, column=index,
                           value=_cell_value(row.get(column.key)))
            cell.font = BODY_FONT
            if column.fmt:
                cell.number_format = column.fmt
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=column.width >= 30)

    first = start_row + 1
    last = start_row + len(rows)

    if rows:
        ref = "A%d:%s%d" % (start_row, get_column_letter(len(columns)), last)
        table = Table(displayName=name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleLight9", showRowStripes=True,
            showFirstColumn=False, showLastColumn=False, showColumnStripes=False)
        ws.add_table(table)
    else:
        cell = ws.cell(row=first, column=1,
                       value="No rows for this sheet in the current tracking data.")
        cell.font = Font(name="Calibri", size=10, italic=True, color="808080")

    if freeze:
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    return first, last


def _severity_rules(ws, letter: str | None, first: int, last: int,
                    critical=("Critical",), warning=("Warning",), good=("OK",)) -> None:
    """Colour a text column by severity.

    Silent when the sheet is empty, and silent when `letter` is None -- which is
    what _letter() returns for a column that has been removed. Without that
    guard, trimming a column fails inside an openpyxl descriptor with a
    TypeError that names neither the sheet nor the column.
    """
    if letter is None or last < first:
        return
    span = "%s%d:%s%d" % (letter, first, letter, last)
    for values, fill, font in ((critical, RED_FILL, RED_FONT),
                               (warning, AMBER_FILL, AMBER_FONT),
                               (good, GREEN_FILL, GREEN_FONT)):
        for value in values:
            ws.conditional_formatting.add(span, CellIsRule(
                operator="equal", formula=['"%s"' % value], fill=fill, font=font))


def _threshold_rules(ws, letter: str | None, first: int, last: int,
                     warn: int, critical: int) -> None:
    """Amber at or above `warn`, red at or above `critical`. Red rule first.

    Silent for a removed column -- see _severity_rules.
    """
    if letter is None or last < first:
        return
    span = "%s%d:%s%d" % (letter, first, letter, last)
    ws.conditional_formatting.add(span, CellIsRule(
        operator="greaterThanOrEqual", formula=[str(critical)],
        fill=RED_FILL, font=RED_FONT, stopIfTrue=True))
    ws.conditional_formatting.add(span, CellIsRule(
        operator="greaterThanOrEqual", formula=[str(warn)],
        fill=AMBER_FILL, font=AMBER_FONT))


def _letter(columns: list[Column], key: str) -> str | None:
    for index, column in enumerate(columns, start=1):
        if column.key == key:
            return get_column_letter(index)
    return None


def _legend(ws, row: int, lines: list[str], span: int) -> int:
    """A short legend under a sheet, so a colour never needs explaining twice."""
    cell = ws.cell(row=row, column=1, value="How to read this sheet")
    cell.font = BOLD_FONT
    for offset, text in enumerate(lines, start=1):
        target = ws.cell(row=row + offset, column=1, value=text)
        target.font = TILE_NOTE_FONT
        ws.merge_cells(start_row=row + offset, start_column=1,
                       end_row=row + offset, end_column=max(span, 2))
    return row + len(lines) + 1


# --------------------------------------------------------------------------- #
# dashboard
# --------------------------------------------------------------------------- #

TILE_WIDTH = 3
TILES_PER_BAND = 4


def _tile(ws, row: int, column: int, label: str, value, note: str,
          alert: bool = False) -> None:
    end = column + TILE_WIDTH - 1
    for offset, (text, font) in enumerate((
            (label, TILE_LABEL_FONT), (value, TILE_VALUE_FONT), (note, TILE_NOTE_FONT))):
        cell = ws.cell(row=row + offset, column=column, value=text)
        cell.font = font
        cell.alignment = Alignment(horizontal="left", vertical="center",
                                   wrap_text=offset == 2)
        ws.merge_cells(start_row=row + offset, start_column=column,
                       end_row=row + offset, end_column=end)
        for fill_column in range(column, end + 1):
            target = ws.cell(row=row + offset, column=fill_column)
            target.fill = TILE_ALERT_FILL if alert else TILE_FILL
            target.border = BOX
    if isinstance(value, (int, float)):
        ws.cell(row=row + 1, column=column).number_format = INT_FMT
    ws.row_dimensions[row].height = 14
    ws.row_dimensions[row + 1].height = 28
    ws.row_dimensions[row + 2].height = 22


def _mini_table(ws, row: int, heading: str, headers: tuple[str, str],
                rows: list[tuple], total_label: str | None = None) -> int:
    """A two-column breakdown table. -> row after the table."""
    cell = ws.cell(row=row, column=1, value=heading)
    cell.font = SECTION_FONT
    row += 1
    for index, text in enumerate(headers, start=1):
        head = ws.cell(row=row, column=index, value=text)
        head.font = HEADER_FONT
        head.fill = PatternFill("solid", fgColor=BLUE)
        head.border = BOX
    header_row = row
    row += 1
    for key, value in rows:
        label = ws.cell(row=row, column=1, value=_cell_value(key))
        label.font = BODY_FONT
        label.border = BOX
        count = ws.cell(row=row, column=2, value=value)
        count.font = BODY_FONT
        count.number_format = INT_FMT
        count.border = BOX
        row += 1
    if total_label and rows:
        label = ws.cell(row=row, column=1, value=total_label)
        label.font = BOLD_FONT
        total = ws.cell(row=row, column=2, value=sum(v for _, v in rows))
        total.font = BOLD_FONT
        total.number_format = INT_FMT
        row += 1
    return header_row, row


def _chart(ws, kind: str, title: str, header_row: int, count: int,
           anchor: str, height: float = 7.5, width: float = 14.0) -> None:
    """Anchor a chart to the mini-table that starts at `header_row`."""
    if count == 0:
        return
    if kind == "pie":
        chart = PieChart()
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.dataLabels.showPercent = False
    elif kind == "line":
        chart = LineChart()
        chart.y_axis.title = "Events"
        chart.smooth = False
    else:
        chart = BarChart()
        chart.type = "bar" if kind == "hbar" else "col"
        chart.gapWidth = 60
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
    chart.title = title
    chart.height = height
    chart.width = width
    chart.style = 10
    data = Reference(ws, min_col=2, min_row=header_row, max_row=header_row + count)
    cats = Reference(ws, min_col=1, min_row=header_row + 1, max_row=header_row + count)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    if kind != "pie":
        chart.legend = None
    ws.add_chart(chart, anchor)


def _dashboard(wb, data, kpi: dict, dist: dict, available: list[str]) -> None:
    ws = wb.create_sheet(S_DASH)
    ws.sheet_view.showGridLines = False
    span = 12

    window = "no dated events"
    if kpi["earliest_event"] and kpi["latest_event"]:
        window = "%s to %s" % (kpi["earliest_event"].strftime("%d %b %Y"),
                               kpi["latest_event"].strftime("%d %b %Y"))
    tracked = "tracking data last refreshed %s" % (
        kpi["newest_snapshot"].strftime("%d %b %Y %H:%M") if kpi["newest_snapshot"]
        else "unknown")

    _title(ws, "Cargo Tracking Report",
           "Generated %s UTC   |   source: %s   |   event window: %s   |   %s"
           % (kpi["as_of"].strftime("%d %b %Y %H:%M"), data.input_dir, window, tracked),
           span)
    row = _nav(ws, 4, available, span)
    row += 1

    # Tiles. Ordered so the first band is the fleet, the second is what needs
    # action, and the third is the state of the data behind the report.
    # Tiles. Band 1 is the fleet at the grain the business works in (BLs),
    # band 2 is what needs action, band 3 is how the consignments are moving,
    # band 4 is the state of the data behind the report.
    # Ten tiles. The previous twenty-four could not be read at a glance, which
    # is the only thing a tile is for -- anything that needs studying belongs on
    # a sheet. Band 1 is the fleet, band 2 is what needs action, band 3 is the
    # state of the data behind it.
    tiles = [
        ("SHIPMENTS", kpi["total_shipments"],
         "bills of lading tracked", False),
        ("CONTAINERS", kpi["total_containers"],
         "%s boxes per BL, largest %d"
         % (kpi["avg_containers_per_bl"], kpi["largest_bl"]), False),
        ("IN TRANSIT", kpi["shipments_in_transit"],
         "%d containers moving" % kpi["in_transit"], False),
        ("DELIVERED", kpi["shipments_delivered"],
         "%d containers delivered" % kpi["delivered"], False),

        ("NEEDS ATTENTION", kpi["shipments_attention"],
         "%d critical -- see Action Required" % kpi["shipments_critical"],
         kpi["shipments_attention"] > 0),
        ("DELAYED", kpi["shipments_delayed"],
         "carrier moved its own ETA later", kpi["shipments_delayed"] > 0),
        ("ETA PASSED", kpi["shipments_overdue"],
         "ETA gone by, no arrival reported", kpi["shipments_overdue"] > 0),
        ("FREE TIME AT RISK", kpi["free_time_expired"] + kpi["free_time_expiring"],
         "containers at port, expired or within %d days"
         % analyze.FREE_TIME_WARN_DAYS,
         (kpi["free_time_expired"] + kpi["free_time_expiring"]) > 0),

        ("CONTAINERS AT PORT", kpi["containers_at_port"],
         "ashore, awaiting collection", False),
        ("WAITING AT A HUB", kpi["at_hub"],
         "between vessels", False),
        ("TRACKING FAILURES", kpi["tracking_failures"],
         "see Tracking Failures", kpi["tracking_failures"] > 0),
        ("TRACKING DATA AGE", "%s d" % (kpi["oldest_snapshot"] and
                                        (kpi["as_of"] - kpi["oldest_snapshot"]).days),
         "oldest row; re-run the tracker to refresh", False),
    ]
    for index, (label, value, note, alert) in enumerate(tiles):
        band, slot = divmod(index, TILES_PER_BAND)
        _tile(ws, row + band * 3, 1 + slot * TILE_WIDTH, label, value, note, alert)
    row += ((len(tiles) + TILES_PER_BAND - 1) // TILES_PER_BAND) * 3 + 2

    # Each section below is a table plus a chart of the same numbers: the table
    # so a figure can be read exactly, the chart so the shape is obvious. Every
    # one answers a question someone actually asks -- see analyze.distributions.
    # Five charts. Each one is on the dashboard because no sheet answers its
    # question as quickly; everything else that used to be here duplicated a
    # sheet that does it better, so it went.
    sections = [
        ("Shipments by status", ("Status", "Shipments"), dist["bl_by_status"],
         "pie", "How many consignments are moving, waiting or finished?"),
        ("Containers by company", ("Company", "Containers"),
         dist["containers_by_company"], "col",
         "Which of our companies has cargo on the water?"),
        ("Where cargo is sitting now", ("Location", "Containers"),
         dist["by_location"], "hbar",
         "Where is our cargo physically piling up?"),
        ("Where time is lost (median dwell, days)",
         ("Location", "Median days"), dist["by_dwell"], "hbar",
         "Which port or hub holds cargo longest?"),
        ("What needs doing", ("Action", "Shipments"), dist["by_exception"],
         "hbar", "What kind of problem dominates the action list?"),
    ]
    for heading, headers, table_rows, kind, question in sections:
        # A section with nothing in it is not informative -- the company
        # breakdown is empty when the file covers a single company, and a bare
        # heading over an empty table just makes the dashboard longer.
        if not table_rows:
            continue
        note = ws.cell(row=row, column=1, value=question)
        note.font = TILE_NOTE_FONT
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
        row += 1
        header_row, after = _mini_table(ws, row, heading, headers, list(table_rows),
                                        total_label=None)
        _chart(ws, kind, heading, header_row, len(table_rows), "D%d" % (row + 1))
        row = max(after, row + 17) + 1

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 3
    for letter in "DEFGHIJKL":
        ws.column_dimensions[letter].width = 11


# --------------------------------------------------------------------------- #
# operational sheets
# --------------------------------------------------------------------------- #

BL_COLUMNS = [
    Column("Company", "company", 11),
    Column("BL / Tracking No.", "bl", 18),
    Column("Carrier", "carrier", 12),
    Column("Boxes", "containers", 7, INT_FMT),
    Column("Status", "status", 11),
    Column("Attention", "attention", 10),
    Column("What Needs Doing", "needs", 38),
    Column("Last Confirmed Move", "confirmed_status", 24),
    Column("Moved On", "latest_move", 12, DATE_FMT),
    Column("Where It Is Now", "location", 22),
    Column("Vessel", "vessel", 18),
    Column("Carrier ETA", "eta", 12, DATE_FMT),
    Column("Days To ETA", "days_to_eta", 9, INT_FMT),
    Column("Next Expected", "next_event", 22),
    Column("Free Days Left", "free_days_left", 9, INT_FMT),
    Column("Boxes Not With The Rest", "outliers", 28),
    Column("Destination", "destination", 18),
    Column("Commodity", "commodity", 20),
]

ARRIVAL_COLUMNS = [
    Column("Window", "arrival_window", 15),
    Column("Arrival Date", "arrival_date", 13, DATE_FMT),
    Column("Days Away", "days_away", 10, INT_FMT),
    Column("Basis", "arrival_basis", 30),
    Column("Company", "company", 12),
    Column("BL / Tracking No.", "bl", 18),
    Column("Carrier", "carrier", 13),
    Column("Containers", "containers", 10, INT_FMT),
    Column("Arrival Port", "arrival_port", 22),
    Column("Documents", "docs_ready", 11),
    Column("Document Position (sheet)", "document_position", 26),
    Column("Free Time Expires", "free_time_expiry", 14, DATE_FMT),
    Column("Free Days Left", "free_days_left", 11, INT_FMT),
    Column("Commodity", "commodity", 22),
    Column("Qty (mt)", "qty", 10, MONEY_FMT),
    Column("Beneficiary", "beneficiary", 26),
    Column("Status", "status", 12),
    Column("Attention", "attention", 11),
    Column("Vessel", "vessel", 18),
    Column("Destination", "destination", 20),
]

FREETIME_COLUMNS = [
    Column("Free Time Status", "free_time_status", 32),
    Column("Company", "company", 11),
    Column("BL / Tracking No.", "bl", 18),
    Column("Boxes On BL", "containers", 9, INT_FMT),
    Column("Still At Port", "containers_at_port", 9, INT_FMT),
    Column("Destination Port", "port_arrival_location", 20),
    Column("Discharged On", "port_arrival_dt", 12, DATE_FMT),
    Column("Free Days", "free_days", 8, INT_FMT),
    Column("Expires", "free_time_expiry", 12, DATE_FMT),
    Column("Days Left", "free_days_left", 8, INT_FMT),
    Column("Documents", "docs_ready", 10),
    Column("How the expiry was worked out", "free_time_basis", 48),
]

COMPANY_COLUMNS = [
    Column("Company", "company", 15),
    Column("Shipments", "shipments", 10, INT_FMT),
    Column("Containers", "containers", 10, INT_FMT),
    Column("In Transit", "in_transit", 9, INT_FMT),
    Column("Delivered", "delivered", 9, INT_FMT),
    Column("At Port", "at_port", 8, INT_FMT),
    Column("Free Time At Risk", "free_time_at_risk", 11, INT_FMT),
    Column("Documents Pending", "docs_pending", 11, INT_FMT),
    Column("Needs Attention", "needs_attention", 11, INT_FMT),
    Column("Carriers", "carriers", 24),
]

PORT_COLUMNS = [
    Column("Location", "location", 24),
    Column("Kind", "kind", 18),
    Column("Containers Here Now", "here_now", 12, INT_FMT),
    Column("Longest Box Here (days)", "longest_now", 12, INT_FMT),
    Column("Median Dwell (days)", "median_dwell", 12, DEC_FMT),
    Column("Shortest", "min_dwell", 9, INT_FMT),
    Column("Longest", "max_dwell", 9, INT_FMT),
    Column("Dwells Measured", "measured", 11, INT_FMT),
    Column("Containers Seen", "containers", 11, INT_FMT),
    Column("Companies", "companies", 22),
]

LANE_COLUMNS = [
    Column("Lane", "lane", 40),
    Column("Completed Journeys", "journeys", 12, INT_FMT),
    Column("Median Transit (days)", "median", 13, DEC_FMT),
    Column("Fastest", "fastest", 9, INT_FMT),
    Column("Slowest", "slowest", 9, INT_FMT),
    Column("Median Ocean (days)", "median_ocean", 13, DEC_FMT),
    Column("Median Port Dwell (days)", "median_port_dwell", 13, DEC_FMT),
    Column("Usable As Benchmark", "usable", 12),
    Column("Containers", "containers", 11, INT_FMT),
    Column("Carriers", "carriers", 26),
]

CONTAINER_COLUMNS = [
    Column("Container", "container_id", 14),
    Column("BL / Tracking No.", "bl", 18),
    Column("Carrier", "carrier_display", 12),
    Column("Status", "status_bucket", 11),
    Column("Attention", "attention", 10),
    Column("Last Confirmed Move", "last_actual_normalized", 24),
    Column("Moved On", "last_actual_dt", 12, DATE_FMT),
    Column("Where It Is Now", "last_actual_location", 22),
    Column("Vessel", "vessel", 18),
    Column("Carrier ETA", "eta_dt", 12, DATE_FMT),
    Column("Idle Days", "idle_at_snapshot_days", 9, INT_FMT),
    Column("Type", "container_type", 13),
    Column("Events", "events_on_report", 7, INT_FMT),
]

EVENT_COLUMNS = [
    Column("Container", "container_id", 14),
    Column("BL / Tracking No.", "tracking_number", 18),
    Column("Carrier", "carrier_display", 12),
    Column("Date", "event_dt", 16, DATETIME_FMT),
    Column("Basis", "time_basis", 10),
    Column("Carrier Event (original)", "description", 26),
    Column("Normalised Status", "normalized_status", 24),
    Column("Journey Phase", "phase", 13),
    Column("Location", "location", 22),
    Column("Mode", "mode", 8),
    Column("Vessel", "vessel_name", 18),
    Column("Voyage", "voyage", 9),
]

CARRIER_COLUMNS = [
    Column("Carrier", "carrier", 15),
    Column("Coverage", "coverage", 34),
    Column("Containers", "containers", 10, INT_FMT),
    Column("In Transit", "in_transit", 9, INT_FMT),
    Column("Delivered", "delivered", 9, INT_FMT),
    Column("Delayed", "delayed", 9, INT_FMT),
    Column("Sitting Still", "stale", 9, INT_FMT),
    Column("Missing ETA", "missing_eta", 10, INT_FMT),
    Column("Needs Attention", "needs_attention", 11, INT_FMT),
    Column("Tracking Failures", "failures", 10, INT_FMT),
    Column("Events", "events", 9, INT_FMT),
]

EXCEPTION_COLUMNS = [
    Column("Severity", "severity", 10),
    Column("What Needs Doing", "category", 32),
    Column("Company", "company", 11),
    Column("BL / Tracking No.", "bl", 18),
    Column("Scope", "scope", 17),
    Column("Boxes Affected", "containers", 28),
    Column("Detail", "detail", 72),
    Column("Where It Is Now", "location", 20),
    Column("Carrier ETA", "eta", 12, DATE_FMT),
    Column("Evidence (source field)", "evidence", 28),
]

SHIPMENT_COLUMNS = [
    Column("Company", "company", 12),
    Column("Sheet Row", "row_number", 9, INT_FMT),
    Column("Bill of Lading", "bl_number", 20),
    Column("Carrier", "carrier_display", 14),
    Column("Beneficiary", "beneficiary", 28),
    Column("Broker", "broker", 14),
    Column("Commodity", "commodity", 20),
    Column("PI No.", "pi_no", 14),
    Column("Qty", "qty", 11, MONEY_FMT),
    Column("Amount", "amount", 14, MONEY_FMT),
    Column("Delivery Terms", "delivery_terms", 18),
    Column("Free Days", "free_days", 9, INT_FMT),
    Column("Sheet ETA text (NOT USED)", "sheet_eta_raw", 18),
    Column("Sheet Port ETA text (NOT USED)", "eta_port_raw", 20),
    Column("Free Time text", "free_time_raw", 14),
    Column("Origin", "origin", 16),
    Column("Port of Loading", "port_of_loading", 16),
    Column("Rate", "rate", 10, MONEY_FMT),
    Column("Value (qty x rate)", "amount_derived", 14, MONEY_FMT),
    Column("Containers On Sheet", "sheet_container_count", 11, INT_FMT),
    Column("Container Size", "container_size", 12),
    Column("Document Position", "document_position", 26),
    Column("LC No.", "lc_no", 20),
    Column("Carrier ETA", "tracked_eta_dt", 13, DATE_FMT),
    Column("Tracked Status", "tracked_status", 14),
    Column("Tracked Location", "tracked_location", 22),
    Column("Containers", "container_count", 10, INT_FMT),
    Column("Container Numbers", "containers", 60),
]

FAILURE_COLUMNS = [
    Column("Failure Kind", "failure_kind", 20),
    Column("Number Queried", "query", 20),
    Column("Carrier", "carrier_display", 16),
    Column("Confidence", "confidence", 11),
    Column("Carriers Tried", "carriers_tried", 18),
    Column("What to do", "suggested_action", 58),
    Column("Reason from tracker", "error", 70),
    Column("When", "snapshot_dt", 18, DATETIME_FMT),
    Column("Run ID", "run_id", 14),
]

QUALITY_COLUMNS = [
    Column("Severity", "severity", 11),
    Column("Area", "area", 13),
    Column("Check", "check", 34),
    Column("Rows", "count", 9, INT_FMT),
    Column("What it means", "detail", 86),
    Column("Example", "example", 26),
]

RECONCILE_COLUMNS = [
    Column("Source", "source", 32),
    Column("Rows Read", "rows_read", 11, INT_FMT),
    Column("Rows On Report", "rows_on_report", 12, INT_FMT),
    Column("Rows In Raw Sheet", "rows_in_raw_sheet", 13, INT_FMT),
    Column("Rows Explained", "rows_explained", 12, INT_FMT),
    Column("Balanced", "balanced", 10),
    Column("Reason for any difference", "reason", 78),
]

STATUS_MAP_COLUMNS = [
    Column("Carrier", "carrier", 14),
    Column("Original Carrier Event", "original_event", 32),
    Column("Normalised Status", "normalized_status", 28),
    Column("Milestone Code", "milestone", 26),
    Column("Journey Phase", "phase", 14),
    Column("Forecast?", "is_forecast", 10),
    Column("Mapping Source", "mapping_source", 13),
    Column("Event Rows (all runs)", "event_rows", 13, INT_FMT),
]

CARRIER_MAP_COLUMNS = [
    Column("Carrier", "carrier", 16),
    Column("Code", "code", 8),
    Column("One of the eight", "target_carrier", 12),
    Column("Tracking Adapter", "tracking_adapter", 22),
    Column("Coverage in this report", "coverage", 32),
    Column("Containers", "containers", 11, INT_FMT),
    Column("Events", "events", 10, INT_FMT),
    Column("Tracking Failures", "failures", 11, INT_FMT),
]

RUN_COLUMNS = [
    Column("Run ID", "run_id", 16),
    Column("Snapshot", "snapshot_at", 20, DATETIME_FMT),
    Column("Container Rows", "containers", 13, INT_FMT),
    Column("Event Rows", "events", 11, INT_FMT),
    Column("Carriers", "carriers", 26),
    Column("Schema", "schema", 9),
]


def _sheet(wb, name: str, heading: str, subtitle: str, columns: list[Column],
           rows: list[dict], table_name: str, legend: list[str] | None = None,
           hidden: bool = False, freeze: bool = True):
    ws = wb.create_sheet(name)
    _title(ws, heading, subtitle, min(len(columns), 10))
    # Row 3 is free between the subtitle and the header, so every sheet gets a
    # way back without the reader hunting the tab strip.
    _link(ws, 3, 1, "< Back to Dashboard", S_DASH)
    first, last = _table(ws, columns, rows, table_name, start_row=4, freeze=freeze)
    if legend:
        _legend(ws, last + 2, legend, min(len(columns), 8))
    if hidden:
        ws.sheet_state = "hidden"
    return ws, first, last


def _raw_sheet(wb, name: str, rows: list[dict], table_name: str, subtitle: str):
    """Raw rows exactly as read, so every figure can be traced to source.

    Column order is taken from the union of keys across all rows rather than
    the first row alone, so a v1 row and a v2 row can share the sheet without
    either losing a column.
    """
    keys: list[str] = []
    for row in rows:
        for key in row:
            if key not in keys:
                keys.append(key)
    columns = [Column(key, key, max(12, min(30, len(key) + 4))) for key in keys]
    if not columns:
        columns = [Column("(no data)", "_none", 20)]
    return _sheet(wb, name, name, subtitle, columns, rows, table_name, hidden=True)


# --------------------------------------------------------------------------- #
# build
# --------------------------------------------------------------------------- #

def build(data, kpi: dict, dist: dict, carriers: list[dict],
          quality: list[dict], ledger: list[dict], output: Path) -> Path:
    """Write the workbook. Returns the path written."""
    wb = Workbook()
    wb.remove(wb.active)

    available = [S_BLS, S_EXCEPTIONS, S_CONTAINERS, S_EVENTS, S_FREETIME,
                 S_COMPANIES, S_CARRIERS, S_PORTS, S_LANES]
    if data.errors:
        available.append(S_FAILURES)
    available.append(S_QUALITY)
    if not data.shipments:
        # Free time needs the commercial sheet's free-days count. Without
        # shipments.json there is nothing to put on it, so it is omitted rather
        # than shown empty.
        available = [name for name in available if name != S_FREETIME]

    _dashboard(wb, data, kpi, dist, available)

    # --- BL Summary --------------------------------------------------------- #
    # First after the dashboard because this is the grain the business works
    # in: cargo is booked and chased by bill of lading, not by box.
    ws, first, last = _sheet(
        wb, S_BLS, "Shipment Summary by Bill of Lading",
        "One row per BL. %d shipment(s) covering %d container(s). The boxes on a "
        "BL normally travel as one consignment -- this row is that consignment, "
        "and any container that is NOT with the rest is named."
        % (len(data.shipment_rows), len(data.containers)),
        BL_COLUMNS, analyze.sort_shipments(data.shipment_rows), "tblBLs",
        legend=[
            "Sorted worst first, then soonest arriving. What Needs Doing is the "
            "short version of why a row is flagged; Action Required has the full "
            "detail and the box numbers.",
            "Status, Where It Is Now, Vessel and Carrier ETA are the CONSENSUS "
            "across the BL's containers -- the value most of them share. On a "
            "13-box BL where 12 are delivered and one is still in transit, the "
            "shipment reads Delivered and the straggler is named under Boxes Not "
            "With The Rest.",
            "Every date here comes from the carrier. The commercial sheet's own "
            "ETA columns are typed by hand and are not used anywhere in this "
            "workbook.",
            "Free Days Left is carrier discharge + the free-days count from the "
            "sheet, counted only while boxes are still at the port. Blank means "
            "nothing is ashore yet, or the sheet states no free days.",
            "Drill down on Container Summary for the individual boxes, or Event "
            "History for the full movement record.",
        ])
    _severity_rules(ws, _letter(BL_COLUMNS, "attention"), first, last)
    _severity_rules(ws, _letter(BL_COLUMNS, "status"), first, last,
                    critical=("No Events",), warning=("At Origin",),
                    good=("Delivered",))
    left = _letter(BL_COLUMNS, "free_days_left")
    if left and last >= first:
        ws.conditional_formatting.add(
            "%s%d:%s%d" % (left, first, left, last),
            CellIsRule(operator="lessThanOrEqual",
                       formula=[str(analyze.FREE_TIME_CRITICAL_DAYS)],
                       fill=RED_FILL, font=RED_FONT, stopIfTrue=True))
        ws.conditional_formatting.add(
            "%s%d:%s%d" % (left, first, left, last),
            CellIsRule(operator="lessThanOrEqual",
                       formula=[str(analyze.FREE_TIME_WARN_DAYS)],
                       fill=AMBER_FILL, font=AMBER_FONT))
    eta = _letter(BL_COLUMNS, "days_to_eta")
    if eta and last >= first:
        ws.conditional_formatting.add(
            "%s%d:%s%d" % (eta, first, eta, last),
            CellIsRule(operator="lessThan", formula=["0"],
                       fill=RED_FILL, font=RED_FONT))

    # --- Free Time & Detention ---------------------------------------------- #
    if data.shipments:
        free_time = analyze.free_time_rows(data)
        ws, first, last = _sheet(
            wb, S_FREETIME, "Free Time & Detention Exposure",
            "The largest avoidable cost in this data. %d shipment(s) with cargo "
            "ashore or due. Worst first." % len(free_time),
            FREETIME_COLUMNS, free_time, "tblFreeTime",
            legend=[
                "Free time expires = discharge at the DESTINATION port (a "
                "CARRIER event) + the free days on the commercial sheet. The "
                "free days are a day COUNT -- a contractual term, not a typed "
                "date -- which is why they are still used when the sheet's date "
                "columns are not. Where the FREE TIME cell holds a date instead "
                "of a count, no expiry is claimed: that date is not trusted.",
                "Only containers STILL AT PORT are exposed. A BL whose free time "
                "ran out weeks ago but whose boxes all railed out owes nothing -- "
                "Free Time Status leads with the count still on the ground rather "
                "than with a signed number that cannot tell those apart.",
                "A transshipment discharge does NOT start the clock. Colombo and "
                "Tanjung Pelepas are hubs on these routes: the box is between "
                "vessels and was never available for collection. Only a discharge "
                "with no onward sea leg counts.",
                "Blank expiry with cargo ashore means the sheet's FREE TIME cell "
                "is empty, or holds a date rather than a day count. Detention may "
                "already be running and the report cannot tell -- that is a "
                "finding on the Exceptions sheet, not a reassurance.",
            ])
        status = _letter(FREETIME_COLUMNS, "free_time_status")
        if last >= first:
            span = "%s%d:%s%d" % (status, first, status, last)
            for text, fill, font in (("EXPIRED", RED_FILL, RED_FONT),
                                     ("Expires in", AMBER_FILL, AMBER_FONT),
                                     ("Cleared", GREEN_FILL, GREEN_FONT),
                                     ("Not yet ashore", GREEN_FILL, GREEN_FONT)):
                ws.conditional_formatting.add(span, CellIsRule(
                    operator="containsText", formula=[
                        'NOT(ISERROR(SEARCH("%s",%s%d)))' % (text, status, first)],
                    fill=fill, font=font))
            _threshold_rules(ws, _letter(FREETIME_COLUMNS, "max_port_dwell_days"),
                             first, last, 7, 14)
        _severity_rules(ws, _letter(FREETIME_COLUMNS, "docs_ready"), first, last,
                        critical=("Pending",), warning=("Not stated",),
                        good=("Ready",))

    # --- Container Summary -------------------------------------------------- #
    ws, first, last = _sheet(
        wb, S_CONTAINERS, "Container Summary",
        "One row per container, from that container's most recent tracking "
        "result. %d container(s)." % len(data.containers),
        CONTAINER_COLUMNS, data.containers, "tblContainers",
        legend=[
            "Confirmed Status / Last Confirmed Move come from the latest event the "
            "carrier reported as ACTUAL. Forecast rows are excluded, so an MSC "
            "'Estimated Time of Arrival' never appears here as a completed move.",
            "Idle Days When Last Checked is how long the box had gone unscanned "
            "at the moment the carrier was last queried -- amber from %d days, red "
            "from %d. This is the column the no-movement finding uses, because it "
            "is a fact about the cargo."
            % (analyze.STALE_WARN_DAYS, analyze.STALE_CRITICAL_DAYS),
            "Days Since Last Move (today) adds the time since the tracker last "
            "ran. If it is much larger than the idle column, the gap is our "
            "refresh cadence, not the cargo -- see Tracking Data Age.",
            "ETA Slip (days) is the carrier's latest ETA minus the first ETA it "
            "gave us. Positive = the carrier moved its own date later. Blank means "
            "the container has only been tracked once, so there is nothing to "
            "compare against.",
            "Tracking Data Age (days) is how long ago the tracker last looked this "
            "container up -- a large number means re-run the tracker, not that the "
            "cargo has stopped.",
            "Bill of Lading is always blank for Maersk: that endpoint does not "
            "return one. Use Tracking No. / BL to join rows instead.",
        ])
    _threshold_rules(ws, _letter(CONTAINER_COLUMNS, "idle_at_snapshot_days"),
                     first, last, analyze.STALE_WARN_DAYS, analyze.STALE_CRITICAL_DAYS)
    _severity_rules(ws, _letter(CONTAINER_COLUMNS, "attention"), first, last)
    _severity_rules(ws, _letter(CONTAINER_COLUMNS, "status_bucket"), first, last,
                    critical=("No Events",), warning=("At Origin",),
                    good=("Delivered",))
    # ETA slip and snapshot age used to be coloured here too. Both are gone from
    # this sheet: slip belongs to the consignment, not the box, and snapshot age
    # is a fact about the pipeline that now lives on Data Quality.

    # --- Event History ------------------------------------------------------ #
    ws, first, last = _sheet(
        wb, S_EVENTS, "Event History",
        "One row per event, for the same snapshot each container is reported on. "
        "%d event(s). The carrier's own wording is preserved next to the "
        "normalised status." % len(data.events),
        EVENT_COLUMNS, data.events, "tblEvents",
        legend=[
            "Carrier Event (original) is exactly what the carrier sent. Normalised "
            "Status is this report's common vocabulary for it -- both are shown so "
            "nothing is lost in translation.",
            "Basis says whether the event HAPPENED (ACTUAL) or is something the "
            "carrier expects (EXPECTED). Basis Reason says which signal decided "
            "it. Only ACTUAL rows feed last-movement figures.",
            "Mapping Source: 'exact' is a confirmed mapping for wording this "
            "project has seen; 'keyword' is a fallback pattern for a carrier with "
            "no adapter yet and should be confirmed. See the Status Mapping sheet.",
            "Vessel IMO and Flag are blank for every Maersk row -- that feed does "
            "not supply them. Laden is blank on most MSC vessel legs.",
        ])
    _severity_rules(ws, _letter(EVENT_COLUMNS, "time_basis"), first, last,
                    critical=(), warning=("EXPECTED",), good=("ACTUAL",))
    _severity_rules(ws, _letter(EVENT_COLUMNS, "mapping_source"), first, last,
                    critical=("unmapped", "blank"), warning=("keyword",),
                    good=("exact",))

    # --- Company Summary ---------------------------------------------------- #
    companies = analyze.company_summary(data)
    ws, first, last = _sheet(
        wb, S_COMPANIES, "Company Summary",
        "One row per company. Companies come from the master workbook's sheet "
        "names, so a new company tab appears here automatically.",
        COMPANY_COLUMNS, companies, "tblCompanies",
        legend=[
            "Cargo Value is quantity x rate from the commercial sheet. The sheet "
            "states a RATE per tonne, not a total, so this is derived -- treat it "
            "as an indication of exposure, not an invoice figure.",
            "Median Transit (days) counts only journeys that have COMPLETED, so "
            "it is blank for a company whose cargo is all still moving.",
            "'(not stated)' is cargo tracked from the older single-sheet workbook "
            "or from a container-number list, neither of which names a company.",
        ])
    _threshold_rules(ws, _letter(COMPANY_COLUMNS, "free_time_at_risk"),
                     first, last, 1, 3)
    _threshold_rules(ws, _letter(COMPANY_COLUMNS, "docs_pending"), first, last, 1, 3)
    _threshold_rules(ws, _letter(COMPANY_COLUMNS, "split"), first, last, 1, 2)
    if last >= first:
        ws.conditional_formatting.add(
            "%s%d:%s%d" % (_letter(COMPANY_COLUMNS, "containers"), first,
                           _letter(COMPANY_COLUMNS, "containers"), last),
            ColorScaleRule(start_type="min", start_color="FFFFFF",
                           end_type="max", end_color=BLUE))

    # --- Carrier Summary ---------------------------------------------------- #
    ws, first, last = _sheet(
        wb, S_CARRIERS, "Carrier Summary",
        "Carriers are derived from the data, then the eight target carriers are "
        "added so one with no data is visible rather than missing.",
        CARRIER_COLUMNS, carriers, "tblCarriers",
        legend=[
            "Coverage tells you why a carrier has no containers: 'No adapter "
            "built' means the tracking application cannot fetch it at all; "
            "'Attempted, no data' means it was tried and recorded as a failure.",
            "With Forward Schedule counts containers whose carrier supplies "
            "planned events. Maersk does; MSC does not, which is why MSC's "
            "Missing ETA count is not a fault in this pipeline.",
            "Avg ETA Slip is averaged over delayed containers only, so it is not "
            "diluted by containers whose ETA came forward.",
        ])
    if last >= first:
        ws.conditional_formatting.add(
            "%s%d:%s%d" % (_letter(CARRIER_COLUMNS, "containers"), first,
                           _letter(CARRIER_COLUMNS, "containers"), last),
            ColorScaleRule(start_type="min", start_color="FFFFFF",
                           end_type="max", end_color=BLUE))
        _threshold_rules(ws, _letter(CARRIER_COLUMNS, "failures"), first, last, 1, 5)

    # --- Ports & Hubs ------------------------------------------------------- #
    ports = analyze.port_rows(data)
    ws, first, last = _sheet(
        wb, S_PORTS, "Ports & Transshipment Hubs",
        "Where the time goes. Slowest median dwell first, measured from this "
        "operation's own event history.",
        PORT_COLUMNS, ports, "tblPorts",
        legend=[
            "Kind is the important distinction. At a DESTINATION PORT the box is "
            "ashore and collectable, so free time runs and detention can be "
            "charged. At a TRANSSHIPMENT HUB it is between vessels and cannot be "
            "collected -- real delay, but no detention.",
            "Median Dwell is measured over completed waits only. Containers Here "
            "Now and Longest Box Here (days) describe what is sitting there at "
            "this moment, and those waits are still growing.",
            "A hub with a high median is where to push the carrier on routing; a "
            "destination port with a high median is where to push your own "
            "clearance and haulage.",
        ])
    _severity_rules(ws, _letter(PORT_COLUMNS, "kind"), first, last,
                    critical=(), warning=("Transshipment hub",),
                    good=("Destination port",))
    _threshold_rules(ws, _letter(PORT_COLUMNS, "median_dwell"), first, last, 7, 10)
    _threshold_rules(ws, _letter(PORT_COLUMNS, "longest_now"), first, last, 7, 14)

    # --- Lane Performance --------------------------------------------------- #
    lanes = analyze.lane_rows(data)
    ws, first, last = _sheet(
        wb, S_LANES, "Lane Performance",
        "How long each route actually takes, from completed journeys only. "
        "%d lane(s) measured." % len(lanes),
        LANE_COLUMNS, lanes, "tblLanes",
        legend=[
            "Median Transit is origin gate-in to arrival at destination, measured "
            "from this operation's OWN completed journeys -- so it already "
            "includes the transshipment waits and rail legs these routes really "
            "involve, not a carrier's published port-to-port time.",
            "Usable As Benchmark is True only at %d or more completed journeys. "
            "Below that the median is a sample, and no shipment is flagged slow "
            "against it." % analyze.LANE_MIN_JOURNEYS,
            "A shipment is called slow only when it exceeds its own lane's median "
            "by %d%% or more. That finding is on the Exceptions sheet."
            % int((analyze.LANE_SLOW_FACTOR - 1) * 100),
            "A lane with no completed journeys yet does not appear. It will once "
            "cargo finishes -- the history is what makes this sheet work.",
        ])
    _severity_rules(ws, _letter(LANE_COLUMNS, "usable"), first, last,
                    critical=(), warning=("False",), good=("True",))
    if last >= first:
        median = _letter(LANE_COLUMNS, "median")
        ws.conditional_formatting.add(
            "%s%d:%s%d" % (median, first, median, last),
            ColorScaleRule(start_type="min", start_color="C6EFCE",
                           end_type="max", end_color="FFC7CE"))

    # --- Exceptions --------------------------------------------------------- #
    ws, first, last = _sheet(
        wb, S_EXCEPTIONS, "Exceptions / Attention Required",
        "%d finding(s). Class separates a SHIPMENT problem from a DATA problem -- "
        "they need different people." % len(data.exceptions),
        EXCEPTION_COLUMNS, data.exceptions, "tblExceptions",
        legend=[
            "Class = Shipment: something about the cargo needs action -- the "
            "carrier moved its ETA, the ETA has passed with no arrival, or the box "
            "has not been scanned for a while.",
            "Class = Data: something about the information needs fixing -- no ETA "
            "supplied, the tracker has not re-run, a carrier has no adapter. These "
            "are NOT delays and must not be reported as such.",
            "A container appears once per finding, so one container can have "
            "several rows. Filter on Class or Severity to work one list at a time.",
            "Thresholds: ETA slip flagged from %d days (red from %d); no movement "
            "flagged from %d days (red from %d); tracking data flagged stale from "
            "%d days."
            % (analyze.ETA_SLIP_WARN_DAYS, analyze.ETA_SLIP_CRITICAL_DAYS,
               analyze.STALE_WARN_DAYS, analyze.STALE_CRITICAL_DAYS,
               analyze.SNAPSHOT_STALE_DAYS),
        ])
    _severity_rules(ws, _letter(EXCEPTION_COLUMNS, "severity"), first, last,
                    critical=("Critical",), warning=("Warning",), good=())
    _severity_rules(ws, _letter(EXCEPTION_COLUMNS, "class"), first, last,
                    critical=(), warning=("Shipment",), good=("Data",))

    # --- Tracking Failures -------------------------------------------------- #
    if data.errors:
        ws, first, last = _sheet(
            wb, S_FAILURES, "Tracking Failures",
            "%d row(s) the tracker could not or would not look up. The three kinds "
            "need completely different responses." % len(data.errors),
            FAILURE_COLUMNS, data.errors, "tblFailures",
            legend=[
                "'No ocean carrier' is not a failure. Inland/truck cargo never "
                "gets a bill of lading, so there is nothing to track. These rows "
                "are finished permanently -- exclude them from pending work.",
                "'Unsupported carrier' is a real shipment on a line with no "
                "adapter. It will never resolve on its own: either build the "
                "adapter or track it by hand.",
                "'Not found' is the one to investigate -- read the tracker's own "
                "reason in the last column before re-running.",
            ])
        _severity_rules(ws, _letter(FAILURE_COLUMNS, "failure_kind"), first, last,
                        critical=("Not found",), warning=("Unsupported carrier",),
                        good=("No ocean carrier",))

    # --- Data Quality ------------------------------------------------------- #
    ws = wb.create_sheet(S_QUALITY)
    _title(ws, "Data Quality",
           "Integrity checks and the row-count ledger. Nothing leaves the report "
           "without being accounted for here.", 6)
    _link(ws, 3, 1, "< Back to Dashboard", S_DASH)
    cell = ws.cell(row=4, column=1, value="Reconciliation ledger")
    cell.font = SECTION_FONT
    first, last = _table(ws, RECONCILE_COLUMNS, ledger, "tblReconcile",
                         start_row=5, freeze=False)
    _severity_rules(ws, _letter(RECONCILE_COLUMNS, "balanced"), first, last,
                    critical=("FAIL",), warning=(), good=("OK",))
    row = last + 3
    cell = ws.cell(row=row, column=1, value="Integrity findings")
    cell.font = SECTION_FONT
    # A second table on one sheet needs its own column widths; the ledger's are
    # close enough in the first few columns, and the wide detail column is set
    # explicitly below so both tables stay readable.
    qfirst, qlast = _table(ws, QUALITY_COLUMNS, quality, "tblQuality",
                           start_row=row + 1, freeze=False)
    _severity_rules(ws, _letter(QUALITY_COLUMNS, "severity"), qfirst, qlast,
                    critical=("Critical",), warning=("Warning",), good=("Info",))
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["E"].width = 86
    ws.column_dimensions["G"].width = 78
    _legend(ws, qlast + 2, [
        "The ledger must balance: rows read = rows on the report + rows "
        "explained. A FAIL line means the report lost data and generate_report.py "
        "exits non-zero.",
        "Rows never leave the workbook. Superseded snapshots stay on the hidden "
        "Raw Containers and Raw Events sheets -- unhide them from the tab strip "
        "to check any figure against source.",
    ], 6)

    # --- Audit sheets -------------------------------------------------------- #
    _sheet(wb, S_STATUS_MAP, "Status Mapping",
           "Every (carrier, event wording) pair in the data and what it "
           "normalises to. Original wording is never replaced -- only "
           "supplemented.",
           STATUS_MAP_COLUMNS, analyze.status_mapping_rows(data), "tblStatusMap", hidden=True,
           legend=[
               "Mapping Source 'exact' means a confirmed mapping in "
               "reporting/normalize.py EXACT_EVENT_MAP.",
               "'keyword' means a fallback pattern matched -- correct often, but "
               "confirm it and promote the wording to EXACT_EVENT_MAP.",
               "'unmapped' means nothing matched and the event shows as 'Other / "
               "unclassified'. Add it to EXACT_EVENT_MAP.",
               "Event Rows counts every run in the file, not just the reported "
               "snapshot, so a rare wording is not missed.",
           ])
    ws = wb[S_STATUS_MAP]
    _severity_rules(ws, _letter(STATUS_MAP_COLUMNS, "mapping_source"),
                    5, 4 + len(analyze.status_mapping_rows(data)),
                    critical=("unmapped", "blank"), warning=("keyword",),
                    good=("exact",))

    _sheet(wb, S_CARRIER_MAP, "Carrier Mapping",
           "The eight carriers this reporting system supports, plus anything else "
           "found in the data, with the coverage the tracker actually has today.",
           CARRIER_MAP_COLUMNS, analyze.carrier_mapping_rows(data), "tblCarrierMap", hidden=True,
           legend=[
               "'not built' in Tracking Adapter means the tracking application has "
               "no way to fetch that carrier. The reporting layer will handle its "
               "data correctly as soon as an adapter exists -- see README_REPORTING.md.",
               "Codes follow the carrier's own SCAC/BL prefix, matching "
               "carrier_id.py so the two stay joinable.",
           ])

    _sheet(wb, S_RUNS, "Run History",
           "Every tracking run present in the JSON. The report uses the latest row "
           "per container, so older runs are here for context and audit.",
           RUN_COLUMNS, data.runs, "tblRuns", hidden=True,
           legend=[
               "Schema v1 rows were written by pipeline.flatten (via "
               "maersk_browser.py or maersk_burst.py) and predate the "
               "tracking_number column; the report derives it for them.",
               "More than one run here means track_all.py was used with --append. "
               "That is what makes ETA-slip detection possible.",
           ])

    _raw_sheet(wb, S_RAW_SHIPMENTS, data.raw_shipments, "tblRawShipments",
               "Every commercial sheet row as read from shipments.json, "
               "unmodified -- including the columns the report does not use and "
               "the rows with no BL. Hidden by default; unhide to audit.")
    _raw_sheet(wb, S_RAW_CONTAINERS, data.raw_containers, "tblRawContainers",
               "Every container row as read from containers.json, all runs, "
               "unmodified. Hidden by default; unhide to audit any figure.")
    _raw_sheet(wb, S_RAW_EVENTS, data.raw_events, "tblRawEvents",
               "Every event row as read from events.json, all runs, unmodified. "
               "Hidden by default; unhide to audit any figure.")

    wb.active = 0
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return output


# --------------------------------------------------------------------------- #
# self-check
# --------------------------------------------------------------------------- #

def _assert_columns_unique() -> None:
    """Every sheet's headers and keys must be unique.

    A duplicate header makes the Excel table invalid and Excel offers to repair
    the file; a duplicate key silently writes the same value into two columns.
    Both are easy to introduce when adding a column to a long list, and neither
    is obvious by eye -- this ran at import and caught a second
    "Free Time Expires" on BL Summary immediately.
    """
    sheets = {
        "BL Summary": BL_COLUMNS,
        "Arrivals Planner": ARRIVAL_COLUMNS,
        "Free Time & Detention": FREETIME_COLUMNS,
        "Company Summary": COMPANY_COLUMNS,
        "Ports & Hubs": PORT_COLUMNS,
        "Lane Performance": LANE_COLUMNS,
        "Container Summary": CONTAINER_COLUMNS,
        "Event History": EVENT_COLUMNS,
        "Carrier Summary": CARRIER_COLUMNS,
        "Exceptions": EXCEPTION_COLUMNS,
        "Commercial Shipments": SHIPMENT_COLUMNS,
        "Tracking Failures": FAILURE_COLUMNS,
        "Data Quality": QUALITY_COLUMNS,
        "Reconciliation": RECONCILE_COLUMNS,
        "Status Mapping": STATUS_MAP_COLUMNS,
        "Carrier Mapping": CARRIER_MAP_COLUMNS,
        "Run History": RUN_COLUMNS,
    }
    problems = []
    for sheet, columns in sheets.items():
        for label, values in (("header", [c.header for c in columns]),
                              ("key", [c.key for c in columns])):
            seen, repeated = set(), set()
            for value in values:
                if value in seen:
                    repeated.add(value)
                seen.add(value)
            if repeated:
                problems.append("%s has duplicate %s(s): %s"
                                % (sheet, label, ", ".join(sorted(repeated))))
            if any(not v for v in values):
                problems.append("%s has an empty %s" % (sheet, label))
    if problems:
        raise AssertionError("Column definition problems:\n  "
                             + "\n  ".join(problems))


_assert_columns_unique()
