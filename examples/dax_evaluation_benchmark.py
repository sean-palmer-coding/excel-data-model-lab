"""Reproduce bounded scalar evaluation measurements without Excel.

Install excel-model-lab[dax], then run:
  python examples/dax_evaluation_benchmark.py --rows 1000 100000 --output timings.json

Times include validation and evaluation. They exclude model authoring/snapshot
construction. Memory figures are traced Python allocations, not process RSS.
"""
from argparse import ArgumentParser
import json
from pathlib import Path
import platform
from statistics import median
from time import perf_counter
import tracemalloc

import pandas as pd

from excel_model_lab import WorkbookModel
from excel_model_lab.dax_evaluation import DaxEvaluator


def benchmark(count):
    if count < 1:
        raise ValueError('Row counts must be positive')
    model = WorkbookModel.create({'Sales': pd.DataFrame({
        'Amount': range(1, count + 1), 'Group': [i % 100 for i in range(count)]})})
    model.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')
    model.add_measure('Sales', 'Weighted', 'SUMX(Sales, Sales[Amount] * 2)')
    tracemalloc.start()
    try:
        start = perf_counter()
        evaluator = DaxEvaluator(model)
        snapshot_seconds = perf_counter() - start
        snapshot_peak = tracemalloc.get_traced_memory()[1]
    finally:
        tracemalloc.stop()
    results = []
    for name, factor in (('Total', 1), ('Weighted', 2)):
        times = []
        for _ in range(3):
            start = perf_counter()
            result = evaluator.evaluate(name)
            times.append(perf_counter() - start)
            assert result.value == count * (count + 1) // 2 * factor
        tracemalloc.start()
        try:
            evaluator.evaluate(name)
            peak = tracemalloc.get_traced_memory()[1]
        finally:
            tracemalloc.stop()
        results.append({'measure': name, 'runs_seconds': times,
                        'median_seconds': median(times), 'python_allocation_peak_bytes': peak})
    return {'rows': count, 'snapshot_seconds': snapshot_seconds,
            'snapshot_python_allocation_peak_bytes': snapshot_peak, 'queries': results}


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--rows', type=int, nargs='+', default=[1000, 100000])
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    report = {'python': platform.python_version(), 'platform': platform.platform(),
              'backend': 'pbix-mcp==0.9.153', 'benchmarks': [benchmark(n) for n in args.rows]}
    encoded = json.dumps(report, indent=2)
    if args.output:
        args.output.write_text(encoded, encoding='utf-8')
    print(encoded)
