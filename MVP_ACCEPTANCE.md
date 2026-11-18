# Expanded MVP acceptance — November 8, 2026

The **supported expanded MVP has passed its release gates**. It is a local
`0.3.1rc1` candidate, not a published release. Scope boundaries are explicit below.

## Automated and distribution acceptance

| Check | Result |
| --- | --- |
| Frozen Windows suite | 512 passed |
| Full Linux installed-wheel suite | 506 passed, 6 Windows-only skips |
| Wheel/source/rebuilt-wheel code and resources | Byte-identical |
| Licenses, resource provenance and excluded legacy fixtures | Checked |
| Clean Windows installed-wheel regression | 148 passed |
| Final artifact hashes and source commit | Distribution verification report |

## Native acceptance

| Capability | Confirmed behavior |
| --- | --- |
| Stored models and data | New/existing authoring, table/column edits, conversion, growth/shrink/clear/repopulate, worksheet preservation and save/reopen |
| Empty models | Stored and calculated zero-row caches load and process; subsequent population works |
| Dictionaries | Million-row DISTINCTCOUNT=100; tiny/nullable text, integer, Currency and decimal lifecycle checks pass |
| Numeric precision | Wide integers/decimals and supported endpoint-adjacent aggregates match exactly |
| Unicode | Accented/decomposed keys and Japanese text filter correctly after save/reopen |
| Relationships | Direct/chained, one-to-one, activation/removal, unmatched members, calculated-key creation and changed-input rebuilding |
| Hierarchies and sorts | Stored/calculated members and custom sorts, including stored labels sorted by calculated keys and dependent hierarchy rebuilding |
| Measures and KPIs | Creation/update/removal, structural rename, status queries and native KPI graphic definitions survive save/reopen |
| Dates and perspectives | Stored date-key processing, alternate-key relationships, multiple hierarchy selections and perspective lifecycle |
| Calculated columns | Six explicit types, update/removal and table lifecycle; exact saved Python types/values |
| Native sources | Linked-table refresh; multiple SQL sources, rebind/detach, automatic session authentication, save/fresh-instance reopen |
| Combined model | Date table, two relationships, nullable wide integers/Currency, calculated fields, Unicode filtering, hierarchy/sort, KPI and perspective together |

The combined model returns Rows=4, Total=75, Double=150, AuditDistinct=4 and
CalendarRows=3. Native save/reopen and independent all-table Python readback pass.
The final source-preservation SQL smoke returns totals 75/300 before and after
fresh-instance reopen. Native testing used one owned workbook at a time.

Exact evidence, workbook hashes and earlier failed attempts are in
`docs/native-acceptance-2026-11-08.json` and `LOCAL_HANDOFF.md`. Query stability
alone was not treated as correctness: expected values and independent readback
were checked. KPI definitions are verified; exhaustive visual rendering of every
Excel icon/theme combination is not claimed.

## Scope and compatibility boundaries

DAX is stored for Excel to process, not evaluated in Python. Pending calculated
relationship keys have explicit deferred validation; stored keys retain immediate
validation. Calculated date-table keys are rejected with a structured error after
a corrected native reference still produced a circular dependency.

Unsupported storage shapes/locales/codepoints and unsafe report edits are rejected.
Authoring is in-memory and datetime precision is whole seconds. The documented
integer/Currency endpoint restrictions preserve native aggregate correctness.
Actual macOS and Excel 2024 acceptance remains unverified; configured CI is not
platform certification.

DAX execution, worksheet calculation, Power Query authoring, PBIX writing and new
report/Pivot layout authoring were excluded from this MVP. Existing workbook parts
must be preserved or an unsafe operation rejected. No credentials are bundled in
release artifacts, and no publication has been performed.
