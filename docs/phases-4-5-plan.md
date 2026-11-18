# Phases 4 and 5 implementation plan

Phase 4 supplies an optional Windows Excel desktop adapter using xlwings.
The headless authoring API remains independent of Excel and xlwings. Sessions
select an exact workbook path and process, distinguish owned from attached
applications, expose bounded range reads/writes, calculation, explicit refresh,
PivotTable layout updates, save and guarded close/reopen.

Phase 5 supplies a bounded agent operation layer and managed file handoff:
inspect and plan operations, detect changed or unsaved workbooks, checkpoint
before mutation, close the selected workbook, edit a staged copy with the
headless API, validate it, publish atomically only if the original is unchanged,
and reopen. Failure must preserve the original or an explicit recovery copy.
No arbitrary code execution from JSON operations or automatic external refresh.

Acceptance requires fake-provider failure/race/ownership tests, native Excel
range/calculation/report/save/reopen checks, actual headless model editing
through handoff, user-edit rejection and callback-failure recovery, clean wheel
installation, full regression suite and the existing nine-job CI matrix.
Native tests own at most one workbook and never quit unrelated Excel instances.

Work runs in parallel: desktop agent owns session module/tests/guide; workflow
agent owns agent operations/checkpoint module/tests/guide; root owns shared
exports/dependency configuration, integration, native acceptance and packaging.
Phase 7 external endpoint verification remains separately blocked. No PyPI
publication or completion claim for live external connectivity is included.
