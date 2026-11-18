"""Headless dashboard composition; Excel calculates CUBE formulas and reports.

Run: python examples/analytical_dashboard.py dashboard.xlsx
Also writes dashboard.expected.json for independent native acceptance.
"""
from pathlib import Path
import json
import sys
import pandas as pd

from excel_model_lab.composition import compose_workbook
from excel_model_lab.cube_formulas import CubeCell
from excel_model_lab.reports import PivotTableSpec, PivotChartSpec, SlicerSpec, TimelineSpec


def create_dashboard(destination):
    sales = pd.DataFrame({'Category': ['Alpha', 'Beta', 'Alpha', 'Gamma'],
                          'Amount': [10, 20, 30, 40],
                          'OrderDate': pd.to_datetime(['2026-01-01', '2026-01-02', '2026-02-01', '2026-03-01'])})

    def configure_model(model):
        model.add_measure('Sales', 'Total Amount', 'SUM(Sales[Amount])')

    def worksheets(writer, cubes):
        title = writer.add_format({'bold': True, 'font_size': 18, 'font_color': '#17365D'})
        currency = writer.add_format({'num_format': '$#,##0.00'})
        date = writer.add_format({'num_format': 'yyyy-mm-dd'})
        inputs = writer.add_worksheet('Inputs')
        inputs.add_table('A1:C5', {'name': 'SalesInput', 'columns': [
            {'header': 'Category'}, {'header': 'Amount', 'format': currency},
            {'header': 'OrderDate', 'format': date}],
            'data': sales.values.tolist()})
        inputs.set_column('A:C', 18)
        inputs.write('A8', '=Literal source text')
        inputs.freeze_panes(1, 0)
        inputs.conditional_format('B2:B5', {'type': 'data_bar'})
        writer.define_name('InputAmount', '=Inputs!$B$2:$B$5')
        chart = writer.add_chart({'type': 'column'})
        chart.add_series({'name': 'Input amounts', 'categories': '=Inputs!$A$2:$A$5',
                          'values': '=Inputs!$B$2:$B$5'})
        inputs.insert_chart('E2', chart)
        dashboard = writer.add_worksheet('Dashboard')
        dashboard.set_column('A:A', 24)
        dashboard.set_column('B:B', 24)
        dashboard.write('A1', 'Sales dashboard', title)
        total = cubes.measure('Total Amount')
        alpha = cubes.member('Sales', 'Category', 'Alpha')
        beta = cubes.member('Sales', 'Category', 'Beta')
        formulas = {
            'B3': ('Measure', total),
            'B4': ('Category', alpha),
            'B5': ('Total sales', cubes.value(CubeCell('B3'))),
            'B6': ('Alpha sales', cubes.value(CubeCell('B3'), CubeCell('B4'))),
            'B7': ('Selected categories', cubes.set([alpha, beta], caption='Alpha and Beta')),
            'B8': ('Selected count', cubes.set_count(CubeCell('B7', kind='set'))),
            'B9': ('Second category', cubes.ranked_member(CubeCell('B7', kind='set'), 2)),
            'B10': ('Selected sales', cubes.value(CubeCell('B3'), CubeCell('B7', kind='set'))),
            'B11': ('Member caption', cubes.member('Sales', 'Category', 'Alpha', caption='Alpha')),
        }
        for cell, (label, formula) in formulas.items():
            dashboard.write('A' + cell[1:], label)
            cubes.write(dashboard, cell, formula, currency if cell in ('B5', 'B6', 'B10') else None)
        dashboard.write('A13', 'Worksheet input sum')
        dashboard.write_formula('B13', '=SUM(InputAmount)', currency, '#N/A')
        second = writer.add_chart({'type': 'line'})
        second.add_series({'name': 'Input trend', 'values': '=Inputs!$B$2:$B$5'})
        dashboard.insert_chart('D2', second)

    def reports(report):
        report.add_pivot_table(PivotTableSpec('SalesPivot', 'Sales Report',
            rows=(('Sales', 'Category'),), measures=('Total Amount',)))
        report.add_pivot_chart(PivotChartSpec('SalesChart', 'SalesPivot', sheet='Dashboard', cell='D20'))
        report.add_slicer(SlicerSpec('Categories', 'SalesPivot', ('Sales', 'Category'), sheet='Sales Report'))
        report.add_timeline(TimelineSpec('OrderDates', 'SalesPivot', ('Sales', 'OrderDate'), sheet='Sales Report'))

    path = compose_workbook(destination, {'Sales': sales}, worksheets,
                            configure_model=configure_model, configure_reports=reports)
    expected = {'sheet_cells': {'Dashboard': {'B3': 'Total Amount', 'B4': 'Alpha', 'B5': 100,
        'B6': 40, 'B7': 'Alpha and Beta', 'B8': 2, 'B9': 'Beta', 'B10': 60,
        'B11': 'Alpha', 'B13': 100}, 'Inputs': {'A8': '=Literal source text'}},
        'pivot_totals': {'SalesPivot': 100}, 'chart_count': 3, 'slicer_cache_count': 2,
        'note': 'Excel must calculate CUBE formulas and update the native report layout.'}
    path.with_suffix('.expected.json').write_text(json.dumps(expected, indent=2), encoding='utf-8')
    return path


if __name__ == '__main__':
    print(create_dashboard(Path(sys.argv[1] if len(sys.argv) > 1 else 'dashboard.xlsx')))
