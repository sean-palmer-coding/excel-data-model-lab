from pathlib import Path
import pytest
from excel_model_lab.abf import AbfImage, replace_model_part, UnsupportedModelError
from excel_model_lab.calculated_columns import (add_calculated_column, update_calculated_column,
    remove_calculated_column, list_calculated_columns)
from excel_model_lab import ModelReader

FIXTURE = Path(__file__).parent / 'data' / 'native-authoring-base.xlsx'


@pytest.mark.parametrize('dtype,expression,semantic', [
    ('integer', '[ID]*2', 'BigInt'), ('decimal', '[ID]/2', 'Double'),
    ('currency', '[ID]*1.25', 'Currency'), ('datetime', 'DATE(2026,1,[ID])', 'Date'),
    ('boolean', '[ID]>1', 'Boolean'), ('text', '[Label]&"!"', 'WChar')])
def test_calculated_column_explicit_type_and_uncalculated_values(tmp_path, dtype, expression, semantic):
    image = add_calculated_column(AbfImage.from_workbook(FIXTURE), 'Table1', 'Calculated', expression, data_type=dtype)
    column, = list_calculated_columns(image)
    assert (column.expression, column.data_type, column.processing_state) == (expression, semantic, 'Unprocessed')
    from xml.dom import minidom
    from excel_model_lab.schema import _elements, _child, _text
    for file in image.files:
        if file.logical_path and file.logical_path.endswith('.dim.xml'):
            doc = minidom.parseString(image.read_logical_file(file))
            attr = next(a for a in _elements(doc, 'Attribute') if _text(_child(a, 'ID')) == 'Calculated')
            assert _text(_child(_elements(attr, 'KeyColumn')[0], 'DataType')) == semantic
    output = tmp_path / 'calculated.xlsx'
    replace_model_part(FIXTURE, output, image.data)
    from excel_model_lab import WorkbookModel
    model = WorkbookModel(output)
    with pytest.raises(UnsupportedModelError, match='calculated-column API'):
        model.replace_column('Table1', 'Calculated', [1, 2, 3])
    with ModelReader(output) as reader:
        values = reader.get_table('Table1')
        assert list(values.ID) == [1, 2, 3]
        assert values.Calculated.isna().all()
        metadata = reader.excel_metadata.calculated_columns.iloc[0]
        assert metadata['data_type'] == semantic
        assert metadata['processing_state'] == 'Unprocessed'


def test_calculated_update_remove_preserve_stored_values(tmp_path):
    original = AbfImage.from_workbook(FIXTURE)
    image = add_calculated_column(original, 'Table1', 'Calc', '[ID]*2', data_type='integer')
    updated = update_calculated_column(image, 'Table1', 'Calc', '[ID]/2', data_type='decimal')
    assert list_calculated_columns(image)[0].expression == '[ID]*2'
    assert list_calculated_columns(updated)[0].expression == '[ID]/2'
    removed = remove_calculated_column(updated, 'Table1', 'Calc')
    assert not list_calculated_columns(removed)
    output = tmp_path / 'removed.xlsx'
    replace_model_part(FIXTURE, output, removed.data)
    with ModelReader(output) as reader:
        assert list(reader.get_table('Table1').columns) == ['ID', 'Label']
        assert list(reader.get_table('Table1').ID) == [1, 2, 3]
    assert original.data == AbfImage.from_workbook(FIXTURE).data


def test_calculated_column_validates_names_types_and_dependencies():
    image = AbfImage.from_workbook(FIXTURE)
    with pytest.raises(ValueError, match='exists'):
        add_calculated_column(image, 'Table1', 'ID', '1', data_type='integer')
    with pytest.raises(ValueError, match='data_type'):
        add_calculated_column(image, 'Table1', 'Calc', '1', data_type='object')
    with pytest.raises(ValueError, match='expression'):
        add_calculated_column(image, 'Table1', 'Calc', ' ', data_type='integer')
    with pytest.raises(KeyError):
        update_calculated_column(image, 'Table1', 'ID', '1', data_type='integer')
    first = add_calculated_column(image, 'Table1', 'First', '[ID]*2', data_type='integer')
    second = add_calculated_column(first, 'Table1', 'Second', '[First]*2', data_type='integer')
    with pytest.raises(UnsupportedModelError, match='calculation'):
        remove_calculated_column(second, 'Table1', 'First')


def test_public_calculated_crud_materializes_pending_input(tmp_path):
    import pandas as pd
    from excel_model_lab import WorkbookModel
    model = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [10, 20, 30, 40]})})
    model.add_calculated_column('Sales', 'Double', '[Amount] * 2', data_type='integer')
    assert model.list_calculated_columns()[0].expression == '[Amount] * 2'
    output = tmp_path / 'public.xlsx'
    model.save(output)
    with ModelReader(output) as reader:
        data = reader.get_table('Sales')
        assert list(data.Amount) == [10, 20, 30, 40]
        assert len(data) == 4 and data.Double.isna().all()
    model.update_calculated_column('Sales', 'Double', '[Amount] / 2', data_type='decimal')
    assert model.list_calculated_columns()[0].data_type == 'Double'
    model.replace_column('Sales', 'Amount', [1, 2, 3, 4])
    model.save(output)
    with ModelReader(output) as reader:
        assert list(reader.get_table('Sales').Amount) == [1, 2, 3, 4]
        assert reader.get_table('Sales').Double.isna().all()
    model.remove_calculated_column('Sales', 'Double')
    assert model.list_calculated_columns() == ()
    model.save(output)
    with ModelReader(output) as reader:
        assert list(reader.get_table('Sales').columns) == ['Amount']


def test_public_calculated_dependencies_and_invalid_input_are_atomic(tmp_path):
    import pandas as pd
    from excel_model_lab import WorkbookModel
    model = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [10, 20]})})
    model.add_calculated_column('Sales', 'First', '[Amount] * 2', data_type='integer')
    model.add_calculated_column('Sales', 'Second', '[First] * 2', data_type='integer')
    before = model._image.data
    with pytest.raises((ValueError, RuntimeError)):
        model.remove_calculated_column('Sales', 'First')
    assert model._image.data == before
    with pytest.raises(ValueError):
        model.add_calculated_column('Sales', 'Bad', '=', data_type='integer')
    assert model._image.data == before


def test_calculated_inferred_type_uses_engine_extension_namespace():
    from xml.etree import ElementTree as ET
    image = add_calculated_column(AbfImage.from_workbook(FIXTURE), 'Table1', 'Calc', '[ID]*2', data_type='integer')
    dimension = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
    root = ET.fromstring(image.read_logical_file(dimension))
    ns = {'e': 'http://schemas.microsoft.com/analysisservices/2003/engine',
          'd': 'http://schemas.microsoft.com/analysisservices/2010/engine/200'}
    attr = next(a for a in root.findall('.//e:Attributes/e:Attribute', ns) if a.findtext('e:ID', namespaces=ns) == 'Calc')
    inferred = attr.find('d:InferredDatatype', ns)
    assert inferred is not None and inferred.text == 'BigInt'
    assert attr.find('e:InferredDatatype', ns) is None
    assert list(attr).index(inferred) < list(attr).index(attr.find('e:KeyColumns', ns))


def test_native_calculated_caption_differs_from_storage_identity(tmp_path):
    from xml.dom import minidom
    from excel_model_lab.schema import _elements, _child, _text
    image = add_calculated_column(AbfImage.from_workbook(FIXTURE), 'Table1', 'PhysicalId', '[ID]*2', data_type='integer')
    f = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
    doc = minidom.parseString(image.read_logical_file(f))
    attr = next(a for a in _elements(doc, 'Attribute') if _text(_child(a, 'ID')) == 'PhysicalId')
    _child(attr, 'Name').firstChild.data = 'Visible name'
    image = AbfImage(image.assemble({f.storage_path: doc.documentElement.toxml(encoding='utf-8')}))
    column, = list_calculated_columns(image)
    assert column.name == 'Visible name' and column.id == 'PhysicalId'
    image = update_calculated_column(image, 'Table1', 'Visible name', '[ID]*3', data_type='integer')
    assert list_calculated_columns(image)[0].expression == '[ID]*3'
    image = remove_calculated_column(image, 'Table1', 'Visible name')
    assert not list_calculated_columns(image)


def test_only_proven_local_row_calculations_preserve_unrelated_caches():
    from excel_model_lab.calculated_columns import CalculatedColumn, _local_row_expression
    image = AbfImage.from_workbook(FIXTURE)
    def local(expression):
        return _local_row_expression(image, CalculatedColumn('Table1', 'Calc', expression, 'BigInt', 'Processed'))
    assert local('[ID] * 2')
    assert local('FORMAT([ID], "000")')
    assert local('LEN("SUM(Other[ID])") + [ID]')
    assert not local('CALCULATE(SUM(Table1[ID]))')
    assert not local('[A measure]')
    assert not local('RELATED(Other[ID])')
    assert not local('NOW()')


@pytest.mark.parametrize('count', [0, 1, 2, 7])
def test_public_calculated_table_resize_rebuilds_blank_rows(tmp_path, count):
    import pandas as pd
    from excel_model_lab import WorkbookModel
    model = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [10, 20, 30]})})
    model.add_calculated_column('Sales', 'Double', '[Amount]*2', data_type='integer')
    values = list(range(100, 100 + count))
    model.replace_table('Sales', pd.DataFrame({'Amount': pd.Series(values, dtype='int64')}))
    staged = model.get_table('Sales')
    assert list(staged.Amount) == values and len(staged.Double) == count
    assert staged.Double.isna().all()
    output = model.save(tmp_path / 'resized.xlsx')
    with ModelReader(output) as reader:
        actual = reader.get_table('Sales')
        assert list(actual.Amount) == values
        assert len(actual.Double) == count and actual.Double.isna().all()
        assert reader.excel_metadata.calculated_columns.iloc[0]['processing_state'] == 'Unprocessed'


def test_zero_row_calculated_table_can_be_repopulated(tmp_path):
    import pandas as pd
    from excel_model_lab import WorkbookModel
    model = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [10, 20]})})
    model.add_calculated_column('Sales', 'Double', '[Amount]*2', data_type='integer')
    empty = pd.DataFrame({'Amount': pd.Series([], dtype='int64')})
    path = model.replace_table('Sales', empty).save(tmp_path / 'empty.xlsx')
    reopened = WorkbookModel(path)
    reopened.replace_table('Sales', pd.DataFrame({'Amount': [4, 5, 6]}))
    assert list(reopened.get_table('Sales').Amount) == [4, 5, 6]
    populated = reopened.save(tmp_path / 'populated.xlsx')
    with ModelReader(populated) as reader:
        values = reader.get_table('Sales')
        assert list(values.Amount) == [4, 5, 6]
        assert len(values.Double) == 3 and values.Double.isna().all()
        assert reader.excel_metadata.calculated_columns.iloc[0]['expression'] == '[Amount]*2'


def test_add_from_native_caption_uses_requested_name_and_rejects_caption_duplicates():
    from xml.dom import minidom
    from excel_model_lab.schema import _elements, _child, _text
    from excel_model_lab.calculated_columns import CalculatedColumnValidationError
    image = AbfImage.from_workbook(FIXTURE)
    file = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
    doc = minidom.parseString(image.read_logical_file(file))
    attribute = next(a for a in _elements(doc, 'Attribute') if _text(_child(a, 'ID')) == 'ID')
    _child(attribute, 'Name').firstChild.data = 'Visible ID'
    image = AbfImage(image.assemble({file.storage_path: doc.documentElement.toxml(encoding='utf-8')}))
    with pytest.raises(CalculatedColumnValidationError) as error:
        add_calculated_column(image, 'Table1', 'visible id', '[Visible ID]*2', data_type='integer')
    assert error.value.issues[0].code == 'duplicate_column'
    added = add_calculated_column(image, 'Table1', 'New Column', '[Visible ID]*2', data_type='integer')
    column, = list_calculated_columns(added)
    assert column.name == column.id == 'New Column'


def test_invalid_calculated_definition_exposes_structured_issue():
    from excel_model_lab.calculated_columns import CalculatedColumnValidationError
    with pytest.raises(CalculatedColumnValidationError) as error:
        add_calculated_column(AbfImage.from_workbook(FIXTURE), 'Table1', 'Bad', '[ID]', data_type=[])
    assert error.value.issues[0].object_path == ('tables', 'Table1', 'columns', 'Bad')


def test_calculated_cache_with_outgoing_custom_sort_retains_binding():
    from xml.dom import minidom
    from excel_model_lab.schema import _elements, _child, _text
    image = add_calculated_column(AbfImage.from_workbook(FIXTURE), 'Table1', 'Calc', '[ID]*2', data_type='integer')
    file = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
    doc = minidom.parseString(image.read_logical_file(file))
    attribute = next(a for a in _elements(doc, 'Attribute') if _text(_child(a, 'ID')) == 'Calc')
    order = _child(attribute, 'OrderByAttributeID')
    while order.firstChild:
        order.removeChild(order.firstChild)
    order.appendChild(doc.createTextNode('ID'))
    image = AbfImage(image.assemble({file.storage_path: doc.documentElement.toxml(encoding='utf-8')}))
    updated = update_calculated_column(image, 'Table1', 'Calc', '[ID]*3', data_type='integer')
    raw_file = next(f for f in updated.files if f.logical_path and f.logical_path.endswith('.tbl.xml')
                    and f.logical_path.replace('\\', '/').rsplit('/', 1)[-1].startswith('Table1.'))
    raw_doc = minidom.parseString(updated.read_logical_file(raw_file))
    raw = next(n for n in _elements(raw_doc, 'XMObject') if n.getAttribute('class') == 'XMRawColumn' and n.getAttribute('name') == 'Calc')
    assert _text(_child(_child(raw, 'Properties'), 'OrderByColumn')) == 'ID'
    assert list_calculated_columns(updated)[0].processing_state == 'Unprocessed'


def test_preflight_reports_unprocessed_calculated_cache_as_warning():
    import pandas as pd
    from excel_model_lab import WorkbookModel
    model = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [1, 2]})})
    model.add_calculated_column('Sales', 'Double', '[Amount]*2', data_type='integer')
    report = model.validate()
    warning, = [issue for issue in report.issues if issue.code == 'unprocessed_calculated_column']
    assert warning.severity == 'warning'
    assert warning.object_path == ('tables', 'Sales', 'columns', 'Double')
    assert report.valid


def test_public_calculated_sort_preserves_unprocessed_binding(tmp_path):
    import pandas as pd
    from excel_model_lab import WorkbookModel
    from excel_model_lab.sorting import stored_sorts
    model = WorkbookModel.create({'T': pd.DataFrame({'A': [3, 1, 2]})})
    model.add_calculated_column('T', 'Label', '[A]&"!"', data_type='text')
    model.add_calculated_column('T', 'Rank', '[A]', data_type='integer')
    model.set_sort_column('T', 'Label', 'Rank')
    path = model.save(tmp_path / 'sorted.xlsx')
    reopened = WorkbookModel(path)
    assert stored_sorts(reopened) == [('T', 'Label', 'Rank')]
    assert reopened.get_table('T').Label.isna().all()
    reopened.set_sort_column('T', 'Label', None)
    assert stored_sorts(reopened) == []
    reopened.set_sort_column('T', 'Label', 'A')
    assert stored_sorts(reopened) == [('T', 'Label', 'A')]
    before = reopened._image.data
    with pytest.raises(ValueError):
        reopened.set_sort_column('T', 'A', 'Rank')
    assert reopened._image.data == before


def test_calculated_hierarchy_invalidation_removes_only_traversal_cache():
    import pandas as pd
    from xml.dom import minidom
    from excel_model_lab import WorkbookModel
    from excel_model_lab.schema import _elements, _child, _text
    model = WorkbookModel.create({'T': pd.DataFrame({'A': [1, 2, 3]})})
    model.add_hierarchy('T', 'Detail', ['A'])
    model.add_calculated_column('T', 'Calc', '[A]*2', data_type='integer')
    image = model._image
    file = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
    doc = minidom.parseString(image.read_logical_file(file))
    hierarchy = _elements(_child(_elements(doc, 'Dimension')[0], 'Hierarchies'), 'Hierarchy')[0]
    _elements(hierarchy, 'SourceAttributeID')[0].firstChild.data = 'Calc'
    hierarchy.setAttribute('custom', 'retained')
    image = AbfImage(image.assemble({file.storage_path: doc.documentElement.toxml(encoding='utf-8')}))
    table_id = model._dimension_ids['T']
    updated = update_calculated_column(image, table_id, 'Calc', '[A]*3', data_type='integer')
    assert not any(f'U${table_id}$Detail.' in f.logical_path for f in updated.files if f.logical_path)
    doc = minidom.parseString(updated.read_logical_file(file.storage_path))
    hierarchy = _elements(_child(_elements(doc, 'Dimension')[0], 'Hierarchies'), 'Hierarchy')[0]
    assert hierarchy.getAttribute('custom') == 'retained'
    assert _text(_child(hierarchy, 'ProcessingState')) == 'Unprocessed'
    assert _text(_child(hierarchy, 'StructureType')) == 'Unknown'


def test_public_calculated_hierarchy_lifecycle(tmp_path):
    import pandas as pd
    from excel_model_lab import WorkbookModel
    model = WorkbookModel.create({'T': pd.DataFrame({'A': [1, 2, 3]})})
    model.add_calculated_column('T', 'Calc', '[A]*2', data_type='integer')
    model.add_hierarchy('T', 'Detail', ['A', 'Calc'])
    model.configure_hierarchy('T', 'Detail', description='Preserve me')
    model.update_calculated_column('T', 'Calc', '[A]*3', data_type='integer')
    model.replace_table('T', pd.DataFrame({'A': [4, 5]}))
    path = model.save(tmp_path / 'hierarchy.xlsx')
    with ModelReader(path) as reader:
        hierarchy = reader.excel_metadata.hierarchies.iloc[0]
        assert hierarchy['processing_state'] == 'Unprocessed'
        assert hierarchy['description'] == 'Preserve me'
        assert reader.get_table('T').Calc.isna().all()
    reopened = WorkbookModel(path)
    reopened.remove_hierarchy('T', 'Detail')
    reopened.remove_calculated_column('T', 'Calc')

def test_calculated_relationship_input_edit_marks_index_unprocessed(tmp_path):
    import pandas as pd
    from xml.dom import minidom
    from excel_model_lab import WorkbookModel
    from excel_model_lab.schema import _elements, _child, _text
    model=WorkbookModel.create({'Lookup':pd.DataFrame({'ID':[1,2]}), 'Facts':pd.DataFrame({'ID':[1,2,1]})})
    model.add_relationship('Facts','ID','Lookup','ID')
    model=WorkbookModel(model.save(tmp_path/'stored-relationship.xlsx'))
    for table in ('Lookup','Facts'):
        model.add_calculated_column(table,'CalculatedKey','[ID]',data_type='integer')
    changes={}
    for file in model._image.files:
        if not file.logical_path or not file.logical_path.endswith('.dim.xml'):continue
        doc=minidom.parseString(model._image.read_logical_file(file));changed=False
        for rel in _elements(doc,'Relationship'):
            for node in _elements(rel,'AttributeID'):node.firstChild.data='CalculatedKey'
            changed=True
        if changed:changes[file.storage_path]=doc.documentElement.toxml(encoding='utf-8')
    model=WorkbookModel(model.source,_image=AbfImage(model._image.assemble(changes)))
    index_files={f.storage_path:model._image.read_logical_file(f) for f in model._image.files if f.logical_path and '.INDEX.' in f.logical_path}
    model.replace_column('Facts','ID',[2,2,1])
    relationships=[]
    for file in model._image.files:
        if file.logical_path and file.logical_path.endswith('.dim.xml'):
            doc=minidom.parseString(model._image.read_logical_file(file))
            relationships.extend(_text(_child(r,'ProcessingState')) for r in _elements(doc,'Relationship'))
    assert relationships==['Unprocessed']
    for path,payload in index_files.items():assert model._image.read_logical_file(path)==payload
    assert model.get_table('Facts').CalculatedKey.isna().all()

def test_public_calculated_relationship_creation_and_input_edit(tmp_path):
    import pandas as pd
    from excel_model_lab import WorkbookModel
    from excel_model_lab.relationship_storage import stored_relationships
    model=WorkbookModel.create({'Lookup':pd.DataFrame({'ID':[1,2]}),'Facts':pd.DataFrame({'ID':[1,2,1]})})
    model.add_calculated_column('Lookup','Key','[ID]*10',data_type='integer')
    model.add_calculated_column('Facts','Key','[ID]*10',data_type='integer')
    model.add_relationship('Facts','Key','Lookup','Key')
    report=model.validate()
    assert report.valid
    assert any(issue.code=='unprocessed_relationship_keys' and issue.severity=='warning' for issue in report.issues)
    model=WorkbookModel(model.save(tmp_path/'calculated-relationship.xlsx'))
    assert stored_relationships(model)[0][1].from_column=='Key'
    model.replace_column('Facts','ID',[2,2,1])
    model=WorkbookModel(model.save(tmp_path/'changed-relationship.xlsx'))
    assert model.get_table('Facts').ID.tolist()==[2,2,1]
    assert model.get_table('Facts').Key.isna().all()


def test_pending_calculated_foreign_key_still_rejects_duplicate_stored_lookup():
    import pandas as pd
    from excel_model_lab import WorkbookModel
    model=WorkbookModel.create({'Lookup':pd.DataFrame({'ID':[1,1]}),'Facts':pd.DataFrame({'ID':[1,2]})})
    model.add_calculated_column('Facts','Key','[ID]',data_type='integer')
    before=model._image.data
    with pytest.raises(ValueError,match='unique'):
        model.add_relationship('Facts','Key','Lookup','ID')
    assert model._image.data==before and not model._pending_relationships
