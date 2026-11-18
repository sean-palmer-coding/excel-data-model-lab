# Optional scalar DAX evaluation

The Phase 6 adapter evaluates a deliberately bounded, schema-checked subset of
DAX over decoded model data. It does not require Excel to run, change the model,
populate calculated-column caches, or write calculated workbook/report results.
It is not a complete replacement for Excel's DAX engine.

## Install and use

From the repository root (this development snapshot is not published to PyPI):

```bash
python -m pip install ".[dax]"
```

```python
from excel_model_lab.dax_evaluation import DaxEvaluator

# Accepts a path or a live WorkbookModel, including staged stored-data/measures.
evaluator = DaxEvaluator("sales.xlsx", max_seconds=30)
answer = evaluator.evaluate("Total Amount")
print(answer.value, answer.elapsed_seconds, answer.backend_version)

filtered = evaluator.evaluate(
    "Total Amount", filters={("Products", "Category"): ["A"]}
)
ratio = evaluator.evaluate_expression(
    "DIVIDE([Total Amount], COUNTROWS(Sales), 0)", table="Sales"
)
```

Construction captures an independent snapshot. Recreate the evaluator after
editing the source model. An empty selection matches no rows; an absent member
also matches no rows. Filter keys are `(table, column)` pairs, not arbitrary
executable expressions. Measure names are resolved case-insensitively.

The result is `DaxEvaluation(value, elapsed_seconds, backend, backend_version)`.
Python `None` represents a legitimate DAX BLANK. Unsupported expressions,
binding/type errors, backend failures and timeouts raise explicit exceptions;
they are not returned as `None`. Datetime scalar results are Python datetimes.

`DaxBackend` describes the optional backend interface. `PbixMcpBackend` is the
first implementation. Importing the core library does not load the evaluator,
start an MCP server, perform network requests or install software.

## Verified expression surface

The parser permits scalar literals, parentheses, numeric arithmetic, comparisons,
Boolean operations, model table/column references and measure references, with:

`SUM`, `COUNT`, `COUNTROWS`, `COUNTBLANK`, `DISTINCTCOUNT`, `SUMX`, `MAXX`,
`CALCULATE`, `FILTER`, `ALL`, `VALUES`, `RELATED`, `DIVIDE`, `BLANK`, `IF`,
`DATE`, `TRUE`, and `FALSE`.

This is a set of supported expression shapes, not a claim that every possible
combination of those functions has been proved equivalent. A parser/type-binding
pass runs before the backend. For example:

- `SUM(T[Text])`, `SUM(T[Flag])` and `COUNT(T[Flag])` reject instead of returning
  the upstream evaluator's misleading BLANK or number.
- Unknown tables/columns and circular measures reject before execution.
- Numeric iterators require numeric/blank expressions. Unverified mixed-type
  coercion, concatenation, table constructors, VAR syntax, unsupported functions,
  and strict `==` reject. The tested legacy Excel model rejected strict equality.
- `RELATED` requires a direct active relationship. `DIVIDE`'s alternate result
  must be a numeric literal (including a signed literal).
- Integer bounds are checked conservatively across source rows and intermediate
  expressions. A filter that would reduce a wide range does not bypass this
  guard. Decimal/Currency values are not supported in this first adapter.
- Only ASCII table/column/measure identifiers using letters, digits, underscores
  and spaces are certified, beginning with a letter or underscore. The pinned
  evaluator mishandled apostrophes in quoted table names; such schemas reject.
- Calculated columns must already be processed in Excel. Pending calculated
  values or staged changes that could stale their caches reject.
- Date values require timezone-naive whole-second precision. Automatic date-table
  guessing is disabled; workbook metadata supplies the date designation.
  Multiple marked date tables, multihop date propagation, and disconnected
  date-like tables that could trigger upstream heuristic filtering reject.

Floating-point results are not promised to be bit-for-bit identical to VertiPaq
aggregation ordering. No security/RLS enforcement, arbitrary DAX query execution,
dynamic calculation groups, headless calculated-column materialization or model
cache writing is provided. The backend time budget is cooperative, not an OS
resource sandbox. Application services must supply their own process isolation
and resource limits for untrusted or large workloads.

## Backend audit and dependency policy

The selected implementation is the published **pbix-mcp 0.9.153** wheel:

`25688dda8fe1fb98c78e2dcb09bc489821c3c2b4353d62be47670638b4646fd8`

Its [license is MIT](https://github.com/d0nk3yhm/pbix-mcp/blob/main/LICENSE).
No upstream implementation code is copied into this project. The wrapper uses
the actual `DAXContext` / `DAXEngine` API and checks `eval_errors`,
`unsupported_functions`, and `timed_out`; upstream can return `None` for failures,
which would otherwise be indistinguishable from a real blank result.

The source audit inspected revision
`99be9c9f97ee5fa8f014d19f78371bd99e1ca8bb`, then independently downloaded the
published wheel. The source has two post-release date-propagation fixes absent
from that wheel despite reporting the same version. Runtime loading therefore
checks the exact version **and normalized hashes of ten audited runtime files**.
Installing an arbitrary same-version source checkout is not accepted. Backend
upgrades require a new audit and conformance run, not a widened version range.

Upstream declares Python >=3.10 and dependencies `mcp>=1.7.1,<2`, `xpress9`,
`xmhuffman>=0.3.0`, `pydantic>=2`, and `apsw`. Its complete package brings these
dependencies even though this adapter only uses its Python evaluation modules.
They stay in the optional `dax` extra. The imported evaluator path was inspected
for network/process execution and general Python eval/exec calls; none were
found in the audited path. This inspection is not a general security guarantee.

Death2BI was considered as another publicly available evaluator, but its larger
application integration was not adopted or certified. Microsoft ADOMD remains a
separate external-engine client, not a local Python DAX evaluator.

## Acceptance evidence

Independent expected results matched native Excel and the published Python
backend for **17 scalar cases**, including native save/reopen:

- Fourteen aggregate, BLANK, filter, iterator and relationship cases returned
  total `70`, category A `50`, weighted iterator `210`, RELATED-weighted `240`,
  maximum category total `50`, and the independently expected other scalars.
- Three date cases returned `60`, `60`, and `20`. Filtering an unrelated calendar
  did not invent a relationship; filtering the fact date selected the right row.
- Native evidence: workspace `work/phase6/native-oracle-v2-results/` and
  `work/phase6/native-date-results/`. The original strict-equality fixture failed
  in Excel and is not counted as a passing case.
- Python error-path and schema/type/overflow regressions are in
  `tests/test_dax_evaluation.py`: **48 focused checks pass**; model creation and
  file readback are exercised.
- The audit hashes, native comparison and reproducible small/100,000-row
  measurements are in `work/phase6/backend-audit.json` and
  `work/phase6/evaluation-verification.json`. Portable copies are
  [backend audit](phase6-backend-audit.json) and
  [evaluation evidence](phase6-evaluation-evidence.json).

The benchmark constructs an evaluator snapshot separately from timed calls,
checks answers against integer arithmetic, records three same-snapshot timings,
and measures Python allocation peaks in a separate traced call. Allocation peaks
are not process RSS and timings are not a comparison against Excel. The
100,000-row iterator is materially slower than the aggregate; no blanket speed
claim is made.

Observed on Windows/Python 3.12.14 with the pinned published backend:

| Rows | SUM median | SUMX median | Snapshot Python allocation peak |
|---|---:|---:|---:|
| 1,000 | 0.00036 s | 0.01191 s | 2.0 MB |
| 100,000 | 0.00997 s | 1.17676 s | 14.8 MB |

Reproduce the benchmark with:

```bash
python examples/dax_evaluation_benchmark.py --rows 1000 100000 --output timings.json
```

The 100,000-row iterator's separate traced evaluation peaked at approximately
23.3 MB of Python allocations. These figures describe the small synthetic
two-column benchmark, not a general model-size or latency guarantee.
