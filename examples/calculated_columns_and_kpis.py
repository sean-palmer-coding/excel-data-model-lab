"""Create a model whose DAX calculations can be processed in Power Pivot.

Run: python examples/calculated_columns_and_kpis.py output.xlsx
"""
from pathlib import Path
import sys

import pandas as pd
from excel_model_lab import WorkbookModel, ModelReader


def build(output):
    model = WorkbookModel.create({'Sales': pd.DataFrame({
        'ID': [1, 2, 3], 'Amount': [20, 30, 40],
        'Label': ['Alpha', 'Beta', 'Gamma'],
    })})
    definitions = [
        ('Tripled', '[Amount]*3', 'integer'),
        ('Half', '[Amount]/2', 'decimal'),
        ('Quarter', '[Amount]/4', 'currency'),
        ('Next Date', 'DATE(2026,1,[ID])', 'datetime'),
        ('Large', '[Amount]>=30', 'boolean'),
        ('Caption', '[Label]&"!"', 'text'),
    ]
    for name, expression, data_type in definitions:
        model.add_calculated_column('Sales', name, expression, data_type=data_type)
    model.add_measure('Sales', 'Total Amount', 'SUM(Sales[Amount])')
    model.add_measure('Sales', 'Target Amount', '100')
    model.add_kpi('Sales', 'Total Amount', target_measure='Target Amount',
                  mode='centered', thresholds=(0.5, 0.8, 1.2, 1.5), icon='Shapes')
    path = model.save(output)
    with ModelReader(path) as reader:
        assert reader.get_table('Sales')['Tripled'].isna().all()
        assert set(reader.excel_metadata.calculated_columns.processing_state) == {'Unprocessed'}
    return path


if __name__ == '__main__':
    print(build(Path(sys.argv[1] if len(sys.argv) > 1 else 'calculated-columns.xlsx')))
