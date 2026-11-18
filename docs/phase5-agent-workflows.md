# Phase 5: agent operations and workbook handoff

`AgentWorkflow` coordinates a [DesktopSession](phase4-desktop.md) with the
headless workbook APIs. It supplies a small structured operation vocabulary,
reviewable plans, checkpoints, stale-plan detection, and a managed
close/edit/reopen sequence. It does not start an AI agent, call a model service,
evaluate supplied Python strings, or introduce another dependency.

Desktop execution currently requires Windows Excel and the optional `desktop`
extra. Plan and recovery tests run without Excel using a controlled session
double. Headless model authoring remains independent of this adapter.

## Inspect, plan, apply

```python
from excel_model_lab import DesktopSession, AgentWorkflow, AgentOperation

# Select the exact workbook, never whichever workbook happens to be active.
session = DesktopSession.attach(r"C:\Reports\sales.xlsx", pid=12345)
workflow = AgentWorkflow(session, checkpoint_directory=r"C:\Reports\checkpoints")
print(workflow.inspect())

plan = workflow.plan([
    AgentOperation.create("write_range", sheet="Dashboard", address="A1",
                          values=[["Sales summary"]]),
    AgentOperation.create("write_formula", sheet="Dashboard", address="B1",
                          formula="=SUM(B3:B6)"),
    AgentOperation.create("calculate"),
    AgentOperation.create("save"),
    AgentOperation.create("read_range", sheet="Dashboard", address="A1:B1"),
], description="Update the dashboard heading and total")
print(plan.to_dict())  # JSON-serializable review record
result = workflow.apply(plan)
print(result.results[-1])
print(result.checkpoint.workbook)
```

The allowed operations are `read_range`, `write_range`, `write_formula`,
`calculate`, `update_pivots`, and `save`. Range values use exact two-dimensional
JSON scalar matrices. `write_formula` targets one cell and requires a string
beginning with `=`. Literal writes preserve leading formula characters as text.
There is no generic attribute invocation, `eval`, `exec`, refresh, close, or quit
operation. Calculate can evaluate CUBE/external formulas and contact their
configured endpoints; it is an explicit operation, not an offline promise.

Plans validate bounded A1 addresses, matrix dimensions, scalar values, Excel
text/formula limits and argument names before executing anything. A plan contains
at most 100 operations by default, and each operation has a 10-million-character
JSON payload limit. The desktop session's range cell limit also applies.
Operation data are copied into immutable JSON, so modifying the caller's input
list cannot change an already reviewed operation. A plan is single-use and bound
to the creating workflow instance, workbook path, Excel PID, and saved file hash.
Plans are serializable for inspection, but are not a portable execution token;
create a fresh plan after restarting the agent or session.

The workbook must be saved when planning and when applying. Unsaved user changes
are never automatically saved or discarded. Save explicitly through the desktop
session first if that is the intended action, then create a new plan. Mutating
batches create a checkpoint before their first operation. Writes stay unsaved
unless the plan explicitly contains `save`. Read-only plans do not create a
checkpoint. Live operations are sequential and are **not an all-or-nothing
transaction**: a failure reports the completed results in
`WorkflowExecutionError.completed`, and earlier writes may remain in Excel.

## Change the model while handing off a desktop workbook

Do not rewrite the ZIP package while Excel still has it open. Use a staged
handoff instead:

```python
from excel_model_lab import WorkbookModel

def edit_model(input_path, output_path):
    model = WorkbookModel.open(input_path, read_only=False)
    model.replace_column("Sales", "Amount", [20, 30, 40])
    model.save(output_path)

plan = workflow.plan_edit(description="Replace the three Sales Amount values")
print(plan.to_dict())
result = workflow.apply_edit(plan, edit_model)

# The same session is open again on the published workbook.
# Calculation/layout updates and saves are separate explicit operations.
session.calculate()
session.update_pivots()
session.save()
```

The callback is trusted application Python supplied by the developer, not code
loaded from the plan or an agent-produced string. It receives a private input
copy and a separate output path; both retain the workbook's basename so authored
chart references can use the correct name. The callback must create the complete
output workbook and must only modify its output. Modifying the input copy is
detected and rejected. An optional `validate(output_path)` callback may add
domain checks, but must be read-only; output modifications are rejected.

Execution performs these steps:

1. Claim the single-use plan and acquire a cooperative per-workbook lease.
2. Verify the saved workbook identity/hash and create a hashed checkpoint.
3. Recheck the live state and close the workbook without discarding changes.
4. Run the callback against staged paths beside the destination.
5. Check required XLSX package parts, duplicate members, ZIP checksums and core
   XML parsing; run the optional read-only validator. These structural checks
   do not certify native Excel/model compatibility.
6. Preserve the candidate in the checkpoint, verify no matching workbook has
   reopened in Excel, and recheck the destination file hash.
7. Atomically replace the destination and reopen it in the same session.

An edit or validation failure preserves the original file and attempts to reopen
it. An external save during staging prevents replacement and preserves that
external file. An unsuccessful reopen **after** publication does not blindly
restore the checkpoint: the new file remains, with both the checkpoint and
candidate available for recovery. `WorkflowExecutionError` exposes `stage`,
`published`, `checkpoint`, `recovery_paths`, and any `recovery_error`; the original
exception is chained as its cause. A successful handoff does not itself calculate,
refresh sources, update report layouts, save through Excel, close other books,
or quit the Excel application.

## Checkpoints and concurrency boundaries

Checkpoints are retained under `.excelmodel-checkpoints` beside the workbook by
default, or in `checkpoint_directory`. Each has the original workbook, a JSON
manifest containing its SHA-256 and plan, and (after validation) the completed
candidate. `result.checkpoint.verify()` checks the original's hash against the
checkpoint object and manifest. Copies and plan arguments can contain sensitive
workbook data; choose storage and retention accordingly. There is no automatic
checkpoint deletion or restore over an open/current workbook.

Pause user interaction with the target workbook while applying a plan. The
adapter detects changed file hashes, path/PID changes, unexpected transitions in
Excel's Saved flag, and reopening during a file handoff. Excel exposes no edit
generation counter: two edits while a workbook is already dirty cannot be
distinguished reliably. A user edit racing between a check and the next COM
operation also cannot be excluded. The lock file coordinates workflows using
this library; arbitrary applications and user interaction do not honor it.
The final file-hash check plus atomic replace is not a filesystem
compare-and-swap transaction. Use one writer and an exclusive interaction window.

If a process crashes, its `.WORKBOOK.xlsx.excelmodel.lock` lease may remain.
Future workflows refuse to remove someone else's lease automatically. Inspect
the recorded PID, workbook state and preserved checkpoint before manually
removing a stale lease and creating a fresh plan. The checkpoint is a recovery
artifact, not an automatic crash-resume mechanism.

## Verification

The focused suite covers stale/replayed/modified plans, dirty/read-only states,
bounded operation validation, immutable payloads, exact operation dispatch,
live partial failures, checkpoint integrity, callback/validation failures,
concurrent saves and reopen attempts, failed recovery, and successful staged
publication. Native acceptance and the full integration result are recorded
separately by the release acceptance report; mocked sessions do not establish
Excel interoperability.
