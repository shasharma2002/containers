"""Excel reporting layer for the cargo tracking pipeline.

Reads the JSON that track_all.py writes and produces one professional
workbook. Nothing in here imports the tracking modules or touches the
network -- it is a pure read-only consumer of containers.json / events.json /
errors.json / shipments.json, so it cannot break a tracking run.

    loader.py     find and read the JSON; pick the reporting snapshot
    normalize.py  carrier registry, event vocabulary mapping, derived fields
    analyze.py    KPIs, carrier rollup, exception detection
    validate.py   data-quality findings and read/write reconciliation
    excel.py      the workbook writer
"""

__all__ = ["loader", "normalize", "analyze", "validate", "excel"]
