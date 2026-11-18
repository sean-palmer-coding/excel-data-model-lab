"""Reproducible synthetic composition benchmark; Excel is not required.

The adjacent expected JSON can be checked by verify_analytical_workbook.ps1.
Peak memory records Python allocations, not total process working set.
"""
import argparse
import json
from pathlib import Path
from time import perf_counter
import tracemalloc

import pandas as pd
from excel_model_lab import compose_workbook, PivotTableSpec, PivotChartSpec, SlicerSpec


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--rows', type=int, default=10000)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if args.rows < 1:
        parser.error('--rows must be positive')
    data = pd.DataFrame({'Group': [f'Group {i % 100:03d}' for i in range(args.rows)],
                         'Amount': range(args.rows)})
    total = args.rows * (args.rows - 1) // 2
    selected = sum(range(0, args.rows, 100))

    def configure(model):
        model.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')

    def sheets(writer, cubes):
        sheet = writer.add_worksheet('Dashboard')
        sheet.write('A1', 'Synthetic analytical workbook')
        sheet.write('A3', 'All rows')
        sheet.write('A4', 'Group 000')
        measure = cubes.measure('Total')
        cubes.write(sheet, 'B3', cubes.value(measure))
        cubes.write(sheet, 'B4', cubes.value(measure, cubes.member('Sales', 'Group', 'Group 000')))

    def reports(report):
        report.add_pivot_table(PivotTableSpec('GroupTotals', 'Group Report',
            rows=(('Sales', 'Group'),), measures=('Total',)))
        report.add_pivot_chart(PivotChartSpec('GroupChart', 'GroupTotals', sheet='Dashboard', cell='D3'))
        report.add_slicer(SlicerSpec('Groups', 'GroupTotals', ('Sales', 'Group')))

    tracemalloc.start()
    start = perf_counter()
    compose_workbook(args.output, {'Sales': data}, sheets,
                     configure_model=configure, configure_reports=reports)
    seconds = perf_counter() - start
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    expected = {'sheet_cells': {'Dashboard': {'B3': total, 'B4': selected}},
                'pivot_totals': {'GroupTotals': total}, 'chart_count': 1, 'slicer_cache_count': 1}
    args.output.with_suffix('.expected.json').write_text(json.dumps(expected, indent=2), encoding='utf-8')
    results = {'rows': args.rows, 'groups': min(100, args.rows), 'authoring_seconds': seconds,
               'python_peak_bytes': peak, 'file_bytes': args.output.stat().st_size,
               'expected_total': total, 'expected_group_000': selected}
    args.output.with_suffix('.benchmark.json').write_text(json.dumps(results, indent=2), encoding='utf-8')
    print(json.dumps(results))


if __name__ == '__main__':
    main()
