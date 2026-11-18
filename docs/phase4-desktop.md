# Phase 4 — Windows desktop sessions

The optional `DesktopSession` API opens or attaches to one existing `.xlsx`
workbook through xlwings and installed Windows Excel. Core model authoring and
importing `excel_model_lab` remain headless; they do not import xlwings or launch
Excel. macOS, xlwings Server/Lite, browser Excel and `.xlsm` are outside this
adapter's supported boundary.

From this source checkout, install `python -m pip install ".[desktop]"`.

```python
from excel_model_lab import DesktopSession

session = DesktopSession.open("dashboard.xlsx", visible=True)
print(session.pid)
print(session.read_range("Dashboard", "B3:B6"))  # always a 2D matrix
session.write_range("Inputs", "A10:B10", [["Requested target", 100]])
session.write_range("Inputs", "B11", [["=B10*2"]], formulas=True)
session.calculate()
session.update_pivots()
session.save()
session.close()
session.quit()  # only the app this session created, and only when empty
```

For a workbook the user already opened, use an exact path and optionally a PID:

```python
session = DesktopSession.attach("dashboard.xlsx", pid=12345)
state = session.snapshot()
print(state.saved, state.read_only, state.disk_sha256)
```

Attachment searches exact normalized full paths. Multiple matching instances
require a PID. `open` rejects an already-open matching path; it does not silently
attach or choose the active workbook. Unsaved, unnamed workbooks cannot be
attached. A later Save As invalidates the session identity. Sessions are bound
to their creating Python thread; COM calls are synchronous and Excel dialogs
may require interaction. This is not an unattended Excel service.

## Supported operations

- `read_range(sheet, address)` returns a 2D matrix using an exact worksheet name
  and a bounded A1 rectangle. Whole rows/columns, names, unions, external
  references, reversed ranges and out-of-bounds cells are rejected. The default
  `max_cells=100000` can be set at session creation.
- `write_range(sheet, address, values, formulas=False)` requires an exact-shape
  2D list/tuple. Strings beginning `=`, `+`, `-`, `@` or an apostrophe are escaped
  as literal cell text in value mode. Explicit formula mode requires every
  matrix element to be a string beginning `=`. Writes may be partially applied
  if Excel reports an error; there is no automatic rollback of live edits.
- `calculate()` requests calculation of this workbook's sheets. It does not
  invoke app-wide `CalculateFullRebuild`. Asynchronous CUBE results may still be
  pending after return. A CUBE formula can contact its configured endpoint.
- `update_pivots()` initializes the embedded model and calls `Update` for reports
  connected to `ThisWorkbookDataModel`, returning their count. It skips ordinary
  worksheet and other OLAP reports; it does not call `RefreshAll`. Model
  initialization retries the observed Excel busy HRESULT for at most five
  seconds while automatic CUBE calculation starts. Other errors propagate
  immediately, and individual PivotTable updates are never replayed.
- `refresh()` explicitly starts Excel's `RefreshAll`. This may contact external
  systems, prompt for authentication, and update source data. Return does not
  certify completion of asynchronous provider refresh. The caller must verify
  its expected results before saving or handing the workbook off.
- `save()` saves in place and returns a fresh snapshot. Read-only saves/writes
  are rejected. Save As is intentionally not part of this interface.
- `close()` rejects dirty workbooks. Use `close(save=True)` or
  `close(discard=True)` only when that behavior is intended. Closing never quits
  Excel. `reopen()` uses the same app, path and original read-only option.
- `quit()` is permitted only for an app created by this session, after its
  workbook is closed, and only if the app contains no other workbooks. Attached
  apps are never quit. The adapter never kills Excel processes.

No context manager silently saves, discards, closes, or quits on exceptions.
Keep the session handle and explicitly resolve any dirty workbook before cleanup.
If opening partially succeeds and then fails, the error identifies the app PID;
the workbook is left available for recovery rather than discarded.

## Opening and refresh boundaries

Managed open/reopen disables link updates, temporarily forces macro automation
security and disables events during opening, then restores both settings.
Before opening it rejects packages declaring connection/query-table/PivotCache
refresh-on-open or a periodic connection refresh interval, including alternate
XML part names. Embedded-model PivotCache initialization is allowed when its
connection resolves to the embedded Data Model; it is not an external import
refresh. It rejects malformed ZIP/XML content before starting Excel. This is
conservative: clear the refresh-on-open configuration intentionally before using
managed open. It does not modify that configuration behind the caller's back.

These controls do not make Excel a security sandbox. Excel add-ins, trusted
application configuration, formula calculation, and already-open attached
workbooks can have their own behavior. Attaching does not change existing app
settings or cancel already-running refreshes. No external refresh is requested
by the adapter unless `refresh()` is called; formula calculation has its own
connection behavior as described above.

## Change detection and handoff

`DesktopSnapshot` contains `path`, `pid`, Excel's `saved` and `read_only` flags,
and a SHA-256 of the disk file. The fingerprint rejects an observed concurrent
file change during reading. It does not hash unsaved in-memory workbook state.
The Excel dirty flag is not a revision counter: two edits while already dirty
cannot be distinguished using it. Pause user editing during managed operations.
`assert_closed()` also detects a workbook reopened in another instance before
file publication; it is an optimistic check, not an OS lock. Explicitly save
desired live changes, and obtain a new plan before a file-level
handoff. Phase 5 adds staged checkpoints and optimistic disk conflict detection;
it cannot make concurrent desktop editing transactional.

## Primary API references

- [xlwings Books: PID-specific collections and opening](https://docs.xlwings.org/en/latest/api/books.html)
- [xlwings Book: save, close and native API](https://docs.xlwings.org/en/latest/api/book.html)
- [xlwings App: ownership and application lifecycle](https://docs.xlwings.org/en/latest/api/app.html)

Unit tests use controlled workbook doubles and run without Excel. Native Excel
acceptance is recorded separately by the repository's phase acceptance report;
unit tests alone do not certify desktop interoperability.

Native Excel acceptance passed for exact-PID attachment, attached-app quit
rejection, literal `=`, `+` and apostrophe text, dirty-close protection, one
embedded PivotTable update, CUBE total 100, worksheet formula result 12, and
save/reopen. The owned test instance was closed and quit. Local evidence:
`work/phase45/desktop-native.json`.

When model bytes change during headless saving, embedded-model CUBE tuple
results stored in PivotCache definitions are invalidated and Excel is requested
to recalculate. External cube caches, worksheet formulas, formatting and cache
metadata remain intact. Manual workbook calculation mode is preserved; callers
using that mode must explicitly calculate. This prevents an Excel-saved total
from being reused after a staged model edit; it does not populate caches without
Excel or certify unrelated formula result caches.
