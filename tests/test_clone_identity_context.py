"""Column cloning changes identities without rewriting schema constants."""
from xml.dom import minidom

import pandas as pd
import pytest

from excel_model_lab import WorkbookModel
from excel_model_lab.schema import _child, _elements, _text


def _attribute(model, name):
    doc = minidom.parseString(model._image.read_logical_file(model._dimension_paths['T']))
    return next(a for a in _elements(doc, 'Attribute') if _text(_child(a, 'Name')) == name)


@pytest.mark.parametrize('source', ['Key', 'Regular', 'BigInt', 'None', 'true', 'false', '3', 'Dictionary'])
def test_cloning_stored_column_preserves_enums_and_scalar_metadata(tmp_path, source):
    values = ['a', 'b'] if source == 'Dictionary' else [1, 2]
    model = WorkbookModel.create({'T': pd.DataFrame({source: values, 'Other': [3, 4]})})
    before = _attribute(model, source)
    model.add_column('T', 'Copied', values, like=source)
    copied = _attribute(model, 'Copied')
    for property_name in ('Type', 'Usage', 'OrderBy', 'DiscretizationMethod',
                          'InstanceSelection', 'MemberNamesUnique', 'IsAggregatable',
                          'AttributeHierarchyEnabled', 'AttributeHierarchyVisible'):
        assert _text(_child(copied, property_name)) == _text(_child(before, property_name))
    assert [_text(n) for n in _elements(copied, 'DataType')] == [_text(n) for n in _elements(before, 'DataType')]
    assert _text(_child(copied, 'ID')) == 'Copied'
    model = WorkbookModel(model.save(tmp_path / 'cloned.xlsx'))
    assert model.get_table('T')[source].tolist() == values
    assert model.get_table('T')['Copied'].tolist() == values


@pytest.mark.parametrize('source', ['Key', 'Regular', 'BigInt', 'None', 'true', 'false', '3', 'Dictionary'])
def test_calculated_column_clone_keeps_legal_attribute_metadata(source):
    model = WorkbookModel.create({'T': pd.DataFrame({source: [1, 2]})})
    model.add_calculated_column('T', 'Calc.Dot', f'[{source}] * 2', data_type='integer')
    attribute = _attribute(model, 'Calc.Dot')
    assert _text(_child(attribute, 'OrderBy')) == 'Key'
    assert _text(_child(attribute, 'Type')) == 'Regular'
    assert _text(_child(attribute, 'Usage')) == 'Regular'
    assert _text(_child(attribute, 'ID')) == model.list_calculated_columns()[0].id
    assert [_text(n) for n in _elements(attribute, 'Expression')] == [f'[{source}] * 2'] * 2
