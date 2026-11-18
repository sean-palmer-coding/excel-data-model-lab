import pandas as pd
import pytest
from excel_model_lab import WorkbookModel
from excel_model_lab.dependencies import DependencyError


def test_mark_unmark_date_table_preserves_values(tmp_path):
    dates = pd.date_range('2026-01-01', periods=3)
    model = WorkbookModel.create({'Calendar': pd.DataFrame({'Date': dates, 'Day': [1, 2, 3]})})
    model.mark_date_table('Calendar', 'Date')
    graph = model.dependency_graph()
    assert any(edge.reason == 'date table key' for edge in graph.edges)
    with pytest.raises(DependencyError):
        model.remove_column('Calendar', 'Date')
    path = model.save(tmp_path / 'date-table.xlsx')
    reopened = WorkbookModel(path)
    assert list(pd.to_datetime(reopened.get_table('Calendar')['Date'])) == list(dates)
    reopened.unmark_date_table('Calendar')
    assert not any(edge.reason == 'date table key' for edge in reopened.dependency_graph().edges)


@pytest.mark.parametrize('dates, message', [
    (['2026-01-01', '2026-01-01'], 'unique'),
    (['2026-01-01', '2026-01-03'], 'contiguous'),
    (['2026-01-01 12:00', '2026-01-02 12:00'], 'whole days'),
    (['2026-01-01', None], 'blanks'),
])
def test_invalid_date_key_does_not_change_model(dates, message):
    model = WorkbookModel.create({'Calendar': pd.DataFrame({'Date': pd.to_datetime(dates)})})
    before = model._image.data
    with pytest.raises(ValueError, match=message):
        model.mark_date_table('Calendar', 'Date')
    assert model._image.data == before


def test_editing_marked_key_is_revalidated_before_save(tmp_path):
    model = WorkbookModel.create({'Calendar': pd.DataFrame({'Date': pd.date_range('2026-01-01', periods=3)})})
    model.mark_date_table('Calendar', 'Date')
    model.replace_column('Calendar', 'Date', pd.to_datetime(['2026-01-01'] * 3))
    with pytest.raises(ValueError, match='unique'):
        model.save(tmp_path / 'invalid.xlsx')
    assert not (tmp_path / 'invalid.xlsx').exists()


def test_date_key_caption_can_differ_from_physical_id(tmp_path):
    from xml.etree import ElementTree as ET
    from excel_model_lab.abf import AbfImage
    from excel_model_lab.perspectives import _tag
    source = WorkbookModel.create({'Calendar': pd.DataFrame({'DateID': pd.date_range('2026-01-01', periods=3)})}).save(tmp_path / 'base.xlsx')
    image = AbfImage.from_workbook(source)
    file = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
    root = ET.fromstring(image.read_logical_file(file))
    attribute = next(a for a in root.findall('.//' + _tag('Attributes') + '/' + _tag('Attribute'))
                     if a.findtext(_tag('ID')) == 'DateID')
    attribute.find(_tag('Name')).text = 'Visible Date'
    image = AbfImage(image.assemble({file.storage_path: ET.tostring(root)}))
    model = WorkbookModel(source, _image=image)
    model.mark_date_table('Calendar', 'Visible Date')
    assert model.validate().valid
    saved = model.save(tmp_path / 'marked.xlsx')
    reopened = WorkbookModel(saved)
    assert reopened.validate().valid
    reopened.unmark_date_table('Calendar')
    assert reopened.validate().valid


def test_direct_date_errors_are_structured_and_compatible():
    from excel_model_lab import DateTableValidationError, DateTableLookupError, ModelValidationError
    model = WorkbookModel.create({'Calendar': pd.DataFrame({'Date': pd.to_datetime(['2026-01-01', '2026-01-01'])})})
    with pytest.raises(DateTableValidationError) as invalid:
        model.mark_date_table('Calendar', 'Date')
    assert isinstance(invalid.value, (ValueError, ModelValidationError))
    assert invalid.value.issues[0].object_path == ('tables', 'Calendar', 'columns', 'Date')
    with pytest.raises(DateTableLookupError) as missing:
        model.mark_date_table('Calendar', 'Missing')
    assert isinstance(missing.value, KeyError)
    assert missing.value.issues[0].code == 'invalid_date_table'


def test_date_key_updates_processing_granularity_and_restores_row_key():
    from xml.dom import minidom
    from excel_model_lab.schema import _child, _elements, _text
    model = WorkbookModel.create({'Calendar': pd.DataFrame({'Date': pd.date_range('2026-01-01', periods=3), 'OtherDate': pd.date_range('2027-01-01', periods=3)})})
    def keys():
        file = next(f for f in model._image.files if f.logical_path and f.logical_path.endswith('.det.xml'))
        group = _elements(minidom.parseString(model._image.read_logical_file(file)), 'MeasureGroup')[0]
        binding = _elements(_child(group, 'Dimensions'), 'Dimension')[0]
        return {_text(_child(a, 'AttributeID')): _text(_child(a, 'Type')) for a in _elements(_child(binding, 'Attributes'), 'Attribute')}
    before = keys()
    row = next(k for k,v in before.items() if v == 'Granularity')
    model.mark_date_table('Calendar', 'Date')
    assert keys() == {**before, row: 'Regular', 'Date': 'Granularity'}
    model.mark_date_table('Calendar', 'OtherDate')
    assert keys() == {**before, row: 'Regular', 'OtherDate': 'Granularity'}
    model.unmark_date_table('Calendar')
    assert keys() == before


def test_calculated_date_key_reports_explicit_structured_limitation():
    from excel_model_lab import DateTableValidationError
    model = WorkbookModel.create({'Calendar': pd.DataFrame({'Day': [1, 2, 3]})})
    model.add_calculated_column('Calendar', 'Date', 'DATE(2026,1,[Day])', data_type='datetime')
    before = model._image.data
    with pytest.raises(DateTableValidationError, match='Calculated date-table keys are not supported') as error:
        model.mark_date_table('Calendar', 'Date')
    assert error.value.issues[0].object_path == ('tables', 'Calendar', 'columns', 'Date')
    assert model._image.data == before

@pytest.mark.parametrize('lookup_date', [False, True])
def test_date_key_edge_is_unique_and_unmark_retains_relationship_contract(tmp_path, lookup_date):
    from xml.dom import minidom
    from excel_model_lab.schema import _elements, _child, _text
    dates = pd.date_range('2026-01-01', periods=3)
    model = WorkbookModel.create({'Calendar': pd.DataFrame({'Date': dates, 'Day': [1,2,3]}),
                                  'Sales': pd.DataFrame({'Date': dates, 'Day': [1,2,3]})})
    if lookup_date:
        model.add_relationship('Sales', 'Date', 'Calendar', 'Date')
    model.mark_date_table('Calendar', 'Date')
    if not lookup_date:
        model.add_relationship('Sales', 'Day', 'Calendar', 'Day')
    def edges():
        doc = minidom.parseString(model._image.read_logical_file(model._dimension_paths['Calendar']))
        return {(_text(_child(a,'ID')), _text(_child(r,'AttributeID'))): _text(_child(r,'Cardinality'))
                for a in _elements(doc,'Attribute') for r in _elements(_child(a,'AttributeRelationships'),'AttributeRelationship')}
    assert edges()['Date', 'RowNumber'] == 'One'
    model = WorkbookModel(model.save(tmp_path/'marked-related.xlsx'))
    assert edges()['Date', 'RowNumber'] == 'One'
    model.unmark_date_table('Calendar')
    assert edges()['RowNumber', 'Date'] == ('One' if lookup_date else 'Many')
    assert model.validate().valid
