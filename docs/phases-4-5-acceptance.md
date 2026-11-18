# Phase 4 and 5 acceptance

Historical checkpoint: the Phase 7 endpoint blocker mentioned below was resolved
on November 17, 2026; see [live acceptance](phase7-live-acceptance.md).

Development version: **0.5.0.dev2**. These phases add optional desktop integration
to the verified headless model/workbook features. The frozen dev1 checkpoint is
unchanged. Phase 7 live endpoint certification remains blocked separately.

## Phase 4 — desktop sessions

- [x] Exact-path/PID attach and owned app opening with lazy xlwings dependency.
- [x] Bounded range reads/writes, literal text, explicit formulas/calculation,
  embedded report update, explicit refresh, save/close/reopen and ownership guards.
- [x] Independent review of dirty state, ownership, identity and refresh boundaries.
- [x] Native Windows Excel: exact attachment, attached quit rejection, dirty close
  rejection, literals, worksheet total 12, CUBE total 100, one PivotTable update,
  save/reopen and owned instance cleanup.
- [x] Full local regression: **866 passed**, no skips/failures, 487.73 seconds.
- [x] Clean installed package: **66 focused tests passed**, dependency checks,
  optional imports, composition/CUBE/DAX/connection/CLI smoke and all 69 runtime
  files matching source/wheel/source archive.
- [x] All nine installed-wheel CI jobs pass: Windows/Linux/macOS with
  Python 3.10/3.12/3.14.

## Phase 5 — agent workflows

- [x] Validated structured operation plans, checkpoints and single-use execution.
- [x] Changed disk/identity/dirty state detection and cooperative workbook leases.
- [x] Staged close/edit/validate/publish/reopen with explicit recovery evidence.
- [x] Native live batch, user-edit rejection, failed callback reopening the original,
  and concurrent reopen rejection before publication.
- [x] Native complete model-edit handoff: model/Pivot/CUBE total changes from
  100 to 140 and remains 140 after save/reopen. The old CUBE cache is invalidated.
- [x] Runnable desktop example verified natively: CUBE total changes to 90.
- [x] Installed-package verification passes as recorded above.
- [x] All nine installed-wheel CI jobs pass: Windows/Linux/macOS with
  Python 3.10/3.12/3.14.

Native integration exposed saved CUBE tuple results surviving model edits even
when PivotTables showed the new data. The model save path now removes only
embedded-model tuple-cache payloads when model bytes change and requests full
recalculation; it preserves cache identity, metadata, external caches and manual
calculation mode. The final integrated path passed native Excel verification; see
[the evidence record](phases-4-5-native-evidence.json).

The desktop adapter is Windows/installed Excel only; core authoring remains
headless. Asynchronous calculations/refreshes require result verification.
Workflows require an exclusive interaction window, not simultaneous interactive
editing. Live batches may partially apply; checkpoints and explicit errors aid
recovery but do not provide transactional Excel editing or automatic crash resume.

Runtime revision: `532226b6ed7010a14e57de306edfb14be8a0f4b3`.
Native tests use xlwings 0.33.22 and Excel 16.0.20625.20044.

[CI run 43008893891](https://github.com/sean-palmer-coding/excel-model-lab/actions/runs/43008893891)
passed all nine jobs. The local source run has zero skips; existing platform and
unredistributed-fixture skips apply to installed-package CI. Final documentation
changes do not change the tested runtime revision. Phases 4 and 5 are complete
within the supported boundaries recorded here. Phase 7 live endpoint acceptance
remains blocked and is not claimed by this checkpoint.
