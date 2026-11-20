"""Create two related model tables and a PivotTable without installed Excel.

Run: python examples/quickstart.py sales-report.xlsx
Excel calculates and renders the report. See docs/phase2-reports.md for first use.
"""
from pathlib import Path
import sys

import pandas as pd
from excel_model_lab import WorkbookModel, WorkbookReport, PivotTableSpec


def build(destination):
    products = pd.DataFrame({
        "ProductID": [1, 2, 3],
        "Category": ["Camping", "Cycling", "Hiking"],
    })
    sales = pd.DataFrame({
        "ProductID": [1, 2, 1, 3, 2],
        "Revenue": [120, 300, 80, 150, 200],
    })
    model = WorkbookModel.create({"Sales": sales, "Products": products})
    model.add_relationship("Sales", "ProductID", "Products", "ProductID")
    model.add_measure("Sales", "Total Revenue", "SUM(Sales[Revenue])",
                      format_string="$#,##0.00")

    report = WorkbookReport(model)
    report.add_pivot_table(PivotTableSpec(
        "RevenueByCategory", "Category Report",
        rows=(("Products", "Category"),), measures=("Total Revenue",),
    ))
    return report.save(destination)


if __name__ == "__main__":
    print(build(Path(sys.argv[1] if len(sys.argv) > 1 else "sales-report.xlsx")))
