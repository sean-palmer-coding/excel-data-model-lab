r"""Create an experimental model-backed PivotTable and PivotChart without Excel.

Run from an environment with this source checkout installed:
    python examples/model_reports.py model-reports.xlsx

In desktop Excel, select each report's A3 cell, check Defer Layout Update in
the PivotTable Fields pane, and click Update. Ordinary Refresh/Refresh All did
not populate fresh layouts on the tested Excel build. Alternatively run:
    .\tools\update_model_reports.ps1 -Source model-reports.xlsx -Output model-reports-ready.xlsx
The optional helper requires installed Excel. This Python example does not
evaluate DAX or cache report values. See docs/phase2-reports.md for first use.
"""
from pathlib import Path
import sys

import pandas as pd
from excel_model_lab import (
    WorkbookModel, WorkbookReport, PivotTableSpec, PivotChartSpec,
    PivotMemberFilter, PivotSort, SlicerSpec, TimelineSpec,
)


def build(output):
    model = WorkbookModel.create({'Sales': pd.DataFrame({
        'Category': ['Alpha', 'Beta', 'Alpha', 'Gamma'],
        'Region': ['East', 'East', 'West', 'West'],
        'Amount': [10, 20, 30, 40],
        'OrderDate': pd.to_datetime(['2026-01-01', '2026-01-02', '2026-02-01', '2026-03-01']),
    })})
    model.add_measure('Sales', 'Total Amount', 'SUM(Sales[Amount])')
    report = WorkbookReport(model)
    report.add_pivot_table(PivotTableSpec(
        name='SalesPivot', sheet='Sales Report', cell='A3',
        rows=(('Sales', 'Category'),), columns=(('Sales', 'Region'),),
        measures=('Total Amount',),
        sorts=(PivotSort('Sales', 'Category', 'descending'),),
    ))
    report.add_pivot_table(PivotTableSpec(
        name='TotalPivot', sheet='Total', measures=('Total Amount',),
    ))
    report.add_pivot_table(PivotTableSpec(
        name='EastPivot', sheet='East Only', rows=(('Sales', 'Category'),),
        filters=(('Sales', 'Region'),), measures=('Total Amount',),
        member_filters=(PivotMemberFilter('Sales', 'Region', ('East',)),),
    ))
    report.add_pivot_chart(PivotChartSpec(
        name='SalesChart', pivot_table='SalesPivot', cell='G3',
        chart_type='column', title='Sales by category and region',
        style=10, legend_position='bottom', category_axis_title='Category',
        value_axis_title='Amount', number_format='#,##0', show_data_labels=True,
        background_color='FFFFFF', plot_background_color='F2F5FA',
    ))
    report.add_slicer(SlicerSpec(
        name='Categories', pivot_table='SalesPivot', field=('Sales', 'Category'),
        cell='G24', caption='Category', connected_pivots=('TotalPivot',),
    ))
    report.add_timeline(TimelineSpec(
        name='OrderDates', pivot_table='SalesPivot', field=('Sales', 'OrderDate'),
        cell='K24', caption='Order date', connected_pivots=('TotalPivot',),
    ))
    return report.save(output)


if __name__ == '__main__':
    print(build(Path(sys.argv[1] if len(sys.argv) > 1 else 'model-reports.xlsx')))
