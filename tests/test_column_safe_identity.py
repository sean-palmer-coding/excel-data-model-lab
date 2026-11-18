from decimal import Decimal
import pandas as pd
import pytest
from excel_model_lab import WorkbookModel


@pytest.mark.parametrize('values', [
    [1, 2], [1.25, 2.5], [2**60, 2**60 + 1], ['a', 'b'],
    [Decimal('1.2345'), Decimal('2.3456')], [True, False],
    pd.to_datetime(['2026-01-01', '2026-01-02']),
])
def test_dotted_column_types_roundtrip(tmp_path, values):
    model = WorkbookModel.create({'T': pd.DataFrame({'No.': values, 'No': [3, 4]})})
    identity = model._column_ids[('T', 'No.')]
    assert '.' not in identity
    model = WorkbookModel(model.save(tmp_path / 'typed.xlsx'))
    assert model._column_ids[('T', 'No.')] == identity
    assert model.get_table('T')['No.'].tolist() == list(values)
    model.remove_column('T', 'No')
    model = WorkbookModel(model.save(tmp_path / 'removed.xlsx'))
    assert model.get_table('T')['No.'].tolist() == list(values)


def test_dotted_column_add_rename_and_convert(tmp_path):
    model = WorkbookModel.create({'T': pd.DataFrame({'Key': [1, 2]})})
    model.add_column('T', 'Amount.', [3, 4])
    model.change_column_type('T', 'Amount.', 'decimal', values=[3.5, 4.5])
    model.rename_column('T', 'Amount.', 'Renamed.')
    model = WorkbookModel(model.save(tmp_path / 'added.xlsx'))
    assert model.get_table('T')['Renamed.'].tolist() == [3.5, 4.5]
    assert '.' not in model._column_ids[('T', 'Renamed.')]
