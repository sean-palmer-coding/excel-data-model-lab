"""Run with an output filename to create a small dependency-aware example."""
from pathlib import Path
import sys
import pandas as pd
from excel_model_lab import WorkbookModel, DependencyError


def build(destination):
    model = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [10, 20], 'Unused': [1, 2]})})
    with model.transaction() as edit:
        edit.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')
        edit.remove_column('Sales', 'Unused')
    try:
        model.remove_column('Sales', 'Amount')
    except DependencyError as error:
        assert error.issues[0].dependencies == (('tables', 'Sales', 'measures', 'Total'),)
    else:
        raise AssertionError('The measure must protect its input column')
    model.rename_table('Sales', 'Sales Data')
    return model.save(destination)


if __name__ == '__main__':
    build(Path(sys.argv[1]) if len(sys.argv) > 1 else Path('dependency-example.xlsx'))
