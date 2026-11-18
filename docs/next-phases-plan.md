# Delivery plan: phases 3, 6, 7, 4, 5

Started November 11, 2026. User-authorized order: **3 → 6 → 7 → 4 → 5**.
Implementation of 3, 6 and 7 runs in parallel where dependencies permit; each
phase must pass its own acceptance gates before being declared complete.
The verified 0.4.0rc1 Phase 2 checkpoint remains immutable.

## Phase 3 — Complete analytical workbook creation

1. Provide a supported composition API: build worksheets with XlsxWriter, close
   that workbook, then attach a newly authored embedded model and reports.
2. Preserve worksheet data, formatting, formulas, names, tables, charts and
   unrelated package parts; reject conflicting model/connection identities.
3. Add typed CUBE formula builders for embedded-model connections, escaped
   member/measure references, sets and ranked members. Distinguish formula
   syntax/reference validation from Excel calculation and arbitrary MDX.
4. Ship a complete dashboard example, API documentation and optional dependency.
5. Accept only after package-preservation tests, installed-wheel checks and native
   Excel open/model-load/CUBE-result/report-update/save/reopen verification.

## Phase 6 — Optional DAX evaluation

1. Audit reusable evaluators, licenses, actual callable APIs, dependencies and
   expression coverage. Pin evaluated revisions; do not trust parity claims.
2. Define an optional backend interface accepting decoded model tables,
   relationships and measures, with explicit unsupported/error responses.
3. Implement the strongest viable adapter without adding evaluation dependencies
   to core authoring. Do not silently fall back to guessed results.
4. Compare independent expected answers and native Excel answers for aggregates,
   BLANK, filter context, CALCULATE, relationship propagation, iterators and totals.
5. Record correctness, supported boundaries and reproducible runtime/memory
   measurements. Accept supported evaluation only for verified cases; report any
   dependency/integration blocker precisely. No claim of complete DAX parity.

## Phase 7 — Connections, report coverage and compatibility

1. Inventory existing native OLE DB/linked-table and Python source APIs; reconcile
   stale connection documentation against recorded native evidence.
2. Add a bounded external Analysis Services semantic-model connection API,
   separating embedded import/refresh from external live-query semantics.
   Validate definitions and credential handling, preserving unrelated parts.
3. Implement a query adapter for supported external semantic endpoints if a
   usable engine/client is available; do not equate ordinary SQL Server with SSAS.
4. Establish native fixtures for external connections/report querying where an
   authorized endpoint exists, and explicit unsupported results otherwise.
5. Expand meaningful compatibility/scale tests and enumerate supported Excel
   versions, types and report operations; quantify remaining native limitations.
6. External live acceptance requires actual endpoint query/refresh/save/reopen.
   XML-only tests cannot certify connectivity. Missing licensed infrastructure
   is a concrete blocker for that gate, not grounds to fake verification.

## Later phases

Phase 4 adds an optional xlwings desktop-session adapter after the preceding
authoring/evaluation/connection work. Phase 5 adds agent operations, checkpoints,
user-edit detection and managed save/close/file-edit/reopen handoff. These were
subsequently authorized and implemented in 0.5.0.dev2; see the
[Phase 4/5 plan](phases-4-5-plan.md) and
[acceptance record](phases-4-5-acceptance.md) for current verification.

## Parallel ownership and integration

- Workbook agent: composition and CUBE modules, focused tests and example.
- Evaluation agent: evaluator audit, optional backend module, conformance tests.
- Connections agent: external connections/query module and focused tests/docs.
- Root: shared exports/dependency metadata, integration, native Excel acceptance,
  full suite, package builds, GitHub CI and final acceptance records.

Agents do not edit shared exports, project metadata or run desktop Excel. Root
keeps fewer than three workbooks open, owns only its test instances, and performs
native checks sequentially. Code tests use workspace temporary directories.
No publication, credentials embedded in workbooks, or unsupported claims are
necessary to implement these phases. Release approval/account setup remains
separate from implementation.

## Final implementation acceptance

As of November 17, 2026, all agreed phases have passed their documented gates.
The last Phase 7 endpoint gate passed on local SSAS Tabular with Windows
integrated authentication; see [live acceptance](phase7-live-acceptance.md).
This does not extend the scope to cloud authentication, external PivotTable
authoring, full DAX parity or unsupported Excel operations.
