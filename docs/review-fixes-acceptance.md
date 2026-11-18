# Review fixes acceptance — November 18, 2026

The follow-up fixes preserve public column/hierarchy captions while generating
legal internal identities for punctuation-bearing names. Hierarchy levels and
calculated-column invalidation use physical identities. Column cloning changes
identity fields rather than enum/scalar metadata; a source column named `Key`
must not change `OrderBy=Key`. Source-view column bindings remain unchanged.

Integer-mutation regressions now create synthetic input instead of requiring
the historical workbook excluded from release bundles.

## Local verification

Verification used Windows and Python 3.12.14 on local hardware. GitHub-hosted
Actions were not used because the account's Actions budget was exhausted.

- Full source run: 913 passed, with one outdated perspective-test expectation.
- Full installed-wheel run using tests extracted from the source distribution:
  863 passed, 50 expected historical-fixture skips, and the same expectation.
- The test had assumed an internal hierarchy ID equaled its caption, including
  `&`. It now resolves the actual ID and still checks the unchanged caption.
  All 14 perspective tests pass after this correction in both source and
  rebuilt-source-distribution/installed-wheel verification.
- No implementation code changed between these full runs and the focused
  correction. This verifies all 914 source cases and all 864 runnable packaged
  cases; the original full-run failures are recorded here rather than reported
  as clean single runs.
- Wheel metadata validates. The five historical fixtures intentionally excluded
  by `MANIFEST.in` remain absent from the source package.

## Native Excel

[Native evidence](review-fixes-native-evidence.json) records ten load/save/reopen
cases plus a combined calculated-column query. Only one test workbook was open
at a time; the test-owned Excel instances were closed afterward.

- Dotted columns and hierarchies load before and after sibling removal.
- The synthetic integer update loads with 500 rows.
- Appended PivotTables update to totals of 30 and save/reopen.
- Plain and dotted calculated-column controls load and save/reopen.
- A combined model containing `Key`, `No.`, `Calc.Dot`, and `Group.Dot` returns
  calculated values 6 and 8 after explicit native processing; those values are
  unchanged after save/reopen. This processing uses the existing native
  measure-edit acceptance harness, not an external-source refresh.

Calculated columns remain subject to the documented native-processing boundary;
these fixes do not add a complete headless DAX engine.

## Publication boundary

These are code and package acceptance results, not authorization to redistribute
unreviewed historical fixtures. The private repository still contains legacy
fixtures in Git history that are excluded from distributions. Public repository
cleanup and final release publication remain separate from this PR.
