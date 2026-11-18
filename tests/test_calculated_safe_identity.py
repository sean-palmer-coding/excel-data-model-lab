"""Calculated-column lifecycle uses physical IDs independently of captions."""
from xml.dom import minidom

import pandas as pd
import pytest

from excel_model_lab import WorkbookModel
from excel_model_lab.hierarchy_storage import resolve_hierarchy
from excel_model_lab.schema import _child, _elements, _text


def test_dotted_calculated_column_lifecycle_keeps_physical_identity(tmp_path):
    model = WorkbookModel.create({'T': pd.DataFrame({'Key': [1, 2], 'Other': [3, 4]})})
    model.add_calculated_column('T', 'Calc.Dot', '[Key] + 1', data_type='integer')
    column, = model.list_calculated_columns()
    identity = column.id
    assert column.name == 'Calc.Dot'
    assert identity.startswith('_Column_') and '.' not in identity
    model = WorkbookModel(model.save(tmp_path / 'added.xlsx'))
    model.update_calculated_column('T', 'Calc.Dot', '[Key] / 2', data_type='decimal')
    assert model.list_calculated_columns()[0].id == identity
    model.replace_column('T', 'Key', [10, 20])
    model = WorkbookModel(model.save(tmp_path / 'updated.xlsx'))
    column, = model.list_calculated_columns()
    assert (column.name, column.id, column.expression, column.data_type) == (
        'Calc.Dot', identity, '[Key] / 2', 'Double')
    model.replace_table('T', pd.DataFrame({'Key': [10, 20, 30], 'Other': [4, 5, 6]}))
    model = WorkbookModel(model.save(tmp_path / 'resized.xlsx'))
    assert model.list_calculated_columns()[0].id == identity
    assert model.get_table('T')['Calc.Dot'].isna().all()
    assert len(model.get_table('T')) == 3
    model.remove_calculated_column('T', 'Calc.Dot')
    model = WorkbookModel(model.save(tmp_path / 'removed.xlsx'))
    assert model.list_calculated_columns() == ()
    assert model.get_table('T')['Key'].tolist() == [10, 20, 30]
    assert not any(f.logical_path and identity in f.logical_path for f in model._image.files)


def test_dotted_calculated_column_hierarchy_and_dependency_bind_by_identity():
    model = WorkbookModel.create({'T': pd.DataFrame({'Key': [1, 2]})})
    model.add_calculated_column('T', 'Calc.Dot', '[Key] + 1', data_type='integer')
    identity = model.list_calculated_columns()[0].id
    model.add_hierarchy('T', 'Hierarchy.Dot', ['Calc.Dot'])
    hierarchy_id, _ = resolve_hierarchy(model, 'T', 'Hierarchy.Dot')
    document = minidom.parseString(model._image.read_logical_file(model._dimension_paths['T']))
    hierarchy = next(h for h in _elements(document, 'Hierarchy')
                     if _text(_child(h, 'ID')) == hierarchy_id)
    level, = _elements(hierarchy, 'Level')
    assert _text(_child(level, 'Name')) == 'Calc.Dot'
    assert _text(_child(level, 'ID')) == identity
    assert _text(_child(level, 'SourceAttributeID')) == identity
    before = model._image.data
    with pytest.raises(ValueError):
        model.remove_calculated_column('T', 'Calc.Dot')
    assert model._image.data == before
