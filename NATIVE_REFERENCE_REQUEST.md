# Native reference needed for remaining authoring

When back at Excel, open compatibility-suite/api-date-perspective-reference-input.xlsx.
Do not enable external connections or refresh. It contains Calendar dates January
1–3, 2026 and Sales amounts10/20/30, with Total Amount60.

1. In Power Pivot, select Calendar, then Design > Mark as Date Table; choose Date.
2. Under Advanced > Create and Manage (Perspectives), create a perspective named
   Sales view containing Sales.Amount and the Total Amount measure.
3. Save As api-date-perspective-native-reference.xlsx in compatibility-suite.

Keep the generated input unchanged for comparison. If an option is unavailable
or Excel reports an error, report that rather than continuing. These steps create
reference metadata; they do not certify our eventual date/perspective writer.

Separate queued acceptance: explicit one-to-one and empty-column lifecycle are
listed in MVP_ACCEPTANCE.md. No need to repeat previously accepted tests.
