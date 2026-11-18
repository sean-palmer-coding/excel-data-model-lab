from pathlib import Path
from xml.etree import ElementTree as ET

import pytest
from pbixray import PBIXRay

from excel_model_lab.abf import AbfImage, replace_model_part
from excel_model_lab.schema import clone_column
from excel_model_lab.model import ExcelModelTemplate

FIXTURE = Path(__file__).parent / 'data' / 'native-authoring-base.xlsx'


def test_add_typed_column_to_empty_model_then_populate(tmp_path):
    import pandas as pd
    empty = ExcelModelTemplate.create({'Empty':pd.DataFrame({'ID':pd.array([],dtype='Int64')})}).save(tmp_path/'empty.xlsx')
    model = ExcelModelTemplate(empty)
    with pytest.raises(ValueError, match='explicitly typed'):
        model.add_column('Empty','Label',[])
    model.add_column('Empty','Label',pd.array([],dtype='string'))
    added = model.save(tmp_path/'added.xlsx')
    assert ExcelModelTemplate(added).get_table('Empty').columns.tolist() == ['ID','Label']
    frame = pd.DataFrame({'ID':[7,8],'Label':['Seven','Eight']})
    full = ExcelModelTemplate(added).replace_table('Empty',frame).save(tmp_path/'full.xlsx')
    assert ExcelModelTemplate(full).get_table('Empty').to_dict('list') == frame.to_dict('list')


def test_empty_population_with_native_style_hyphenated_table_id(tmp_path):
    import pandas as pd
    model = ExcelModelTemplate(FIXTURE)
    frame = pd.DataFrame({'ID': [7, 8], 'Label': ['Seven', 'Eight']})
    model.add_table('Query1_66b433ff-36de-4cbf', frame, like='Table1')
    full = model.save(tmp_path/'full.xlsx')
    empty = ExcelModelTemplate(full).replace_table('Query1_66b433ff-36de-4cbf', frame.iloc[:0]).save(tmp_path/'empty.xlsx')
    output = ExcelModelTemplate(empty).replace_table('Query1_66b433ff-36de-4cbf', frame).save(tmp_path/'populated.xlsx')
    assert ExcelModelTemplate(output).get_table('Query1_66b433ff-36de-4cbf').to_dict('list') == frame.to_dict('list')


@pytest.mark.parametrize('like', [None, 'Table1'])
def test_added_table_ordinal_follows_native_gaps(tmp_path, like):
    import pandas as pd
    from xml.dom import minidom
    from excel_model_lab.schema import _elements, _text
    model = ExcelModelTemplate(FIXTURE)
    changes = {}
    for file in model._image.files:
        if file.logical_path and file.logical_path.endswith(('.dim.xml', '.det.xml')):
            doc = minidom.parseString(model._image.read_logical_file(file))
            _elements(doc, 'Ordinal')[0].firstChild.data = '7'
            changes[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    image = AbfImage(model._image.assemble(changes))
    model = ExcelModelTemplate(FIXTURE, _image=image)
    model.add_table('Added', pd.DataFrame({'ID': [8, 9, 10], 'Label': ['A', 'B', 'C']}), like=like)
    result = ExcelModelTemplate(model.save(tmp_path / 'added.xlsx'))
    ordinals = []
    for file in result._image.files:
        if file.logical_path and file.logical_path.endswith(('.dim.xml', '.det.xml')):
            doc = minidom.parseString(result._image.read_logical_file(file))
            ordinals.extend(int(_text(n)) for n in _elements(doc, 'Ordinal'))
    assert sorted(ordinals) == [7, 7, 8, 8]


@pytest.mark.parametrize('filename', ['native-authoring-added-column.xlsx', 'null_data_id.xlsx'])
def test_import_table_retains_donor_versions_and_rebinds_database(filename, tmp_path):
    import pandas as pd
    from xml.dom import minidom
    from excel_model_lab.schema import import_table, _elements, _text

    target = ExcelModelTemplate(FIXTURE.parent / filename)
    donor = ExcelModelTemplate.create({'Added': pd.DataFrame({'ID': [8, 9], 'Text': ['A', 'B']})})
    table_id = donor._dimension_ids['Added']
    result = import_table(target._image, donor._image, table_id)
    donor_groups = {g.findtext('Class'): g for g in donor._image.backup_log_root.findall('FileGroups/FileGroup')
                    if g.findtext('ID') == table_id}
    database = next(g.findtext('ID') for g in target._image.backup_log_root.findall('FileGroups/FileGroup')
                    if g.findtext('Class') == '100002')
    for group in result.backup_log_root.findall('FileGroups/FileGroup'):
        if group.findtext('ID') != table_id:
            continue
        assert group.findtext('ObjectVersion') == donor_groups[group.findtext('Class')].findtext('ObjectVersion')
        for entry in group.findall('FileList/BackupFile'):
            if entry.findtext('Path').endswith(('.dim.xml', '.det.xml', '.prt.xml')):
                doc = minidom.parseString(result.read_logical_file(entry.findtext('StoragePath')))
                assert [_text(n) for n in _elements(doc, 'DatabaseID')] == [database]


def test_created_id_member_property_is_not_reserved(tmp_path):
    import pandas as pd
    model = ExcelModelTemplate.create({'T': pd.DataFrame({'ID': [1, 2], 'ID_': [3, 4]})})
    model = ExcelModelTemplate(model.save(tmp_path / 'reserved.xlsx'))
    root = ET.fromstring(model._image.read_logical_file(model._dimension_paths['T']))
    ns = {'a': 'http://schemas.microsoft.com/analysisservices/2003/engine'}
    names = [r.findtext('a:Name', namespaces=ns)
             for r in root.findall('.//a:AttributeRelationship', ns)]
    assert 'ID' not in names
    assert len(names) == len(set(names))


@pytest.mark.parametrize('source', ['ID', 'Label'])
def test_column_clone_decodes_and_retains_original_values(tmp_path, source):
    original = AbfImage.from_workbook(FIXTURE)
    result = clone_column(original, 'Table1', source, 'CopiedColumn')
    dimension = next(f for f in result.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
    root = ET.fromstring(result.read_logical_file(dimension))
    ns = {'a': 'http://schemas.microsoft.com/analysisservices/2003/engine'}
    for relations in root.findall('.//a:AttributeRelationships', ns):
        names = [r.findtext('a:Name', namespaces=ns) for r in relations]
        assert len(names) == len(set(names)), 'Excel rejects duplicate relationship names'
    output = tmp_path / 'clone.xlsx'
    replace_model_part(FIXTURE, output, result.data)
    with PBIXRay(output) as reader:
        table = reader.get_table('Table1')
        assert table['ID'].tolist() == [1, 2, 3]
        assert table['Label'].tolist() == ['Alpha', 'Beta', 'Gamma']
        assert table['CopiedColumn'].tolist() == table[source].tolist()
    for entry in original.files:
        if entry.logical_path and entry.logical_path.endswith(('.idf', '.dictionary', '.hidx')):
            assert original.read_file(entry) == result.read_file(entry.storage_path)


def test_existing_column_name_rejected():
    image=AbfImage.from_workbook(FIXTURE)
    with pytest.raises(ValueError, match='already exists'):
        clone_column(image, 'Table1', 'ID', 'label')


def test_public_add_columns_composes_with_pending_values(tmp_path):
    model = ExcelModelTemplate(FIXTURE)
    model.replace_column('Table1', 'ID', [3, 2, 1])
    model.add_column('Table1', 'Amount', [10, 20, 30], like='ID')
    model.add_column('Table1', 'Category', ['Delta', 'Beta', 'Alpha'], like='Label')
    output = model.save(tmp_path / 'new-columns.xlsx')
    with PBIXRay(output) as reader:
        data = reader.get_table('Table1')
        assert data.to_dict('list') == {
            'ID': [3, 2, 1], 'Label': ['Alpha', 'Beta', 'Gamma'],
            'Amount': [10, 20, 30], 'Category': ['Delta', 'Beta', 'Alpha'],
        }


def test_add_column_invalid_length_is_transactional():
    model = ExcelModelTemplate(FIXTURE)
    model.replace_column('Table1', 'ID', [3, 2, 1])
    before = model._image.data
    with pytest.raises(ValueError, match='needs 3 values'):
        model.add_column('Table1', 'Amount', [10], like='ID')
    assert model._image.data == before
    assert model._pending == {('Table1', 'ID'): [3, 2, 1]}
    assert [c.name for c in model.columns] == ['ID', 'Label']


def test_row_growth_updates_hidden_row_number(tmp_path):
    import struct
    import pandas as pd
    model = ExcelModelTemplate(FIXTURE)
    expected = pd.DataFrame({'ID': list(range(10, 110, 10)), 'Label': ['Alpha', 'Beta'] * 5})
    output = model.replace_table('Table1', expected).save(tmp_path / 'grown.xlsx')
    image = AbfImage.from_workbook(output)
    hidden = next(f for f in image.files if f.logical_path and '.RowNumber.' in f.logical_path)
    assert struct.unpack_from('<I', image.read_logical_file(hidden), 12)[0] == 10
    with PBIXRay(output) as reader:
        assert reader.get_table('Table1').to_dict('list') == expected.to_dict('list')


def test_remove_column_composes_with_addition(tmp_path):
    model = ExcelModelTemplate(FIXTURE).unbind_source('Table1')
    output = model.add_column('Table1', 'Amount', [10, 20, 30], like='ID').remove_column(
        'Table1', 'Label').save(tmp_path / 'removed.xlsx')
    with PBIXRay(output) as reader:
        assert reader.get_table('Table1').to_dict('list') == {'ID': [1, 2, 3], 'Amount': [10, 20, 30]}
    image = AbfImage.from_workbook(output)
    assert not any(f.logical_path and ('$Label.' in f.logical_path or '.Label.' in f.logical_path)
                   for f in image.files)


def test_remove_last_column_is_transactional():
    model = ExcelModelTemplate(FIXTURE).unbind_source('Table1').remove_column('Table1', 'Label')
    before = model._image.data
    with pytest.raises(ValueError, match='at least one'):
        model.remove_column('Table1', 'ID')
    assert model._image.data == before


def test_add_table_has_independent_storage_and_values(tmp_path):
    import pandas as pd
    model = ExcelModelTemplate(FIXTURE)
    expected = {'ID': [10, 20, 30, 40], 'Label': ['Alpha', 'Beta', 'Alpha', 'Beta']}
    output = model.add_table('SecondTable', pd.DataFrame(expected), like='Table1').save(tmp_path / 'tables.xlsx')
    with PBIXRay(output) as reader:
        assert set(reader.tables) == {'Table1', 'SecondTable'}
        assert reader.get_table('SecondTable').to_dict('list') == expected
        assert reader.get_table('Table1').to_dict('list') == {'ID': [1, 2, 3], 'Label': ['Alpha', 'Beta', 'Gamma']}
    image = AbfImage.from_workbook(output)
    groups = image.backup_log_root.findall('FileGroups/FileGroup')
    identities = [g.findtext('ObjectID') for g in groups]
    assert len(identities) == len(set(identities))
    assert sum(g.findtext('ID') == 'SecondTable' for g in groups) == 3

def test_rename_table_preserves_identity_and_pending_data(tmp_path):
    model = ExcelModelTemplate(FIXTURE)
    model.replace_column('Table1', 'ID', [3, 2, 1])
    model.rename_table('Table1', 'Product Catalog')
    assert model._dimension_ids == {'Product Catalog': 'Table1'}
    output = model.save(tmp_path / 'renamed.xlsx')
    with PBIXRay(output) as reader:
        assert list(reader.tables) == ['Product Catalog']
        assert reader.get_table('Product Catalog')['ID'].tolist() == [3, 2, 1]

def test_rename_column_keeps_pending_values(tmp_path):
    model = ExcelModelTemplate(FIXTURE).unbind_source('Table1')
    model.replace_column('Table1', 'ID', [3, 2, 1])
    output = model.rename_column('Table1', 'ID', 'ProductID').save(tmp_path / 'renamed-column.xlsx')
    with PBIXRay(output) as reader:
        assert reader.get_table('Table1').to_dict('list') == {
            'Label': ['Alpha', 'Beta', 'Gamma'], 'ProductID': [3, 2, 1]}

def test_remove_table_deletes_storage_and_keeps_added_data(tmp_path):
    import pandas as pd
    model = ExcelModelTemplate(FIXTURE).unbind_source('Table1')
    expected = {'ID': [10, 20], 'Label': ['Alpha', 'Beta']}
    model.add_table('SecondTable', pd.DataFrame(expected), like='Table1')
    output = model.remove_table('Table1').save(tmp_path / 'removed-table.xlsx')
    with PBIXRay(output) as reader:
        assert list(reader.tables) == ['SecondTable']
        assert reader.get_table('SecondTable').to_dict('list') == expected
    image = AbfImage.from_workbook(output)
    assert not any(g.findtext('ID') == 'Table1' for g in image.backup_log_root.findall('FileGroups/FileGroup'))

def test_create_model_without_caller_template(tmp_path):
    import pandas as pd
    expected = {'ProductID': [10, 20, 30, 40], 'Category': ['Alpha', 'Beta', 'Alpha', 'Beta']}
    model = ExcelModelTemplate.create({'Products': pd.DataFrame(expected)})
    output = model.save(tmp_path / 'created.xlsx')
    with PBIXRay(output) as reader:
        assert list(reader.tables) == ['Products']
        assert reader.get_table('Products').to_dict('list') == expected

def test_column_settings_preserve_data(tmp_path):
    model = ExcelModelTemplate(FIXTURE)
    output = model.configure_column('Table1', 'ID', hidden=True, description='Product key',
        format_string='0.00').save(tmp_path / 'settings.xlsx')
    with PBIXRay(output) as reader:
        assert reader.get_table('Table1')['ID'].tolist() == [1, 2, 3]
    image = AbfImage.from_workbook(output)
    ns = {'a': 'http://schemas.microsoft.com/analysisservices/2003/engine'}
    dim = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
    root = ET.fromstring(image.read_logical_file(dim))
    attr = next(a for a in root.findall('.//a:Attribute', ns) if a.findtext('a:ID', namespaces=ns) == 'ID')
    assert attr.findtext('a:AttributeHierarchyVisible', namespaces=ns) == 'false'
    assert attr.findtext('a:Description', namespaces=ns) == 'Product key'
    assert next(n.text for n in attr if n.tag.endswith('}FormatString')) == '0.00'

def test_create_nullable_columns_and_validate_pandas_na(tmp_path):
    import pandas as pd
    expected = pd.DataFrame({'Number': pd.array([1, None, 3, 4], dtype='Int64'),
                             'Text': pd.array(['Alpha', None, 'New', 'Beta'], dtype='string')})
    model = ExcelModelTemplate.create({'Nullable': expected})
    output = model.save(tmp_path / 'nullable.xlsx')
    with PBIXRay(output) as reader:
        actual = reader.get_table('Nullable')
        assert actual['Number'].isna().tolist() == [False, True, False, False]
        assert actual['Text'].isna().tolist() == [False, True, False, False]
        assert actual.loc[[0,2,3], 'Number'].tolist() == [1,3,4]
        assert actual.loc[[0,2,3], 'Text'].tolist() == ['Alpha','New','Beta']

def test_create_zero_negative_and_null_integers(tmp_path):
    import pandas as pd
    values = pd.array([-10, 0, None, 10], dtype='Int64')
    output = ExcelModelTemplate.create({'Signed': pd.DataFrame({'Number': values})}).save(tmp_path / 'signed.xlsx')
    with PBIXRay(output) as reader:
        actual = reader.get_table('Signed')['Number']
        assert actual.isna().tolist() == [False, False, True, False]
        assert actual.dropna().tolist() == [-10, 0, 10]


def test_create_decimal_columns_and_convert_without_rounding(tmp_path):
    import pandas as pd
    model = ExcelModelTemplate.create({'Values': pd.DataFrame({'Amount': [-1.25, 0, 2.5, 3.75]})})
    output = model.save(tmp_path / 'decimal.xlsx')
    with PBIXRay(output) as reader:
        assert reader.get_table('Values')['Amount'].tolist() == [-1.25, 0, 2.5, 3.75]
    before = model._image.data
    pending = dict(model._pending)
    with pytest.raises(ValueError, match='not exactly representable'):
        model.change_column_type('Values', 'Amount', 'integer')
    assert model._image.data == before
    assert model._pending == pending
    model.replace_column('Values', 'Amount', [-2, 0, 2, 4])
    model.change_column_type('Values', 'Amount', 'integer').save(tmp_path / 'integer.xlsx')


def test_decimal_precision_is_explicit_and_nonfinite_rejected():
    import pandas as pd
    model = ExcelModelTemplate(FIXTURE)
    model.change_column_type('Table1', 'ID', 'decimal', decimal_places=2)
    with pytest.raises(ValueError, match='not exactly representable'):
        model.replace_column('Table1', 'ID', [1.001, 2, 3])
    with pytest.raises(ValueError):
        ExcelModelTemplate.create({'Values': pd.DataFrame({'Amount': [float('inf')]})})


def test_create_dates_currency_and_nullable_booleans(tmp_path):
    import pandas as pd
    from decimal import Decimal
    expected = pd.DataFrame({'Day': pd.to_datetime(['2026-01-01', None, '2026-01-03']),
        'Flag': pd.array([True, None, False], dtype='boolean'),
        'Money': [Decimal('-1.25'), None, Decimal('3.4567')]})
    output = ExcelModelTemplate.create({'Types': expected}).save(tmp_path / 'types.xlsx')
    with ExcelModelTemplate(output)._open_reader() as reader:
        actual = reader.get_table('Types')
        assert actual.isna().to_dict('list') == expected.isna().to_dict('list')
        assert actual['Flag'].dropna().tolist() == [True, False]
        assert actual['Day'].dropna().tolist() == expected['Day'].dropna().tolist()
        assert [str(v) for v in actual['Money'].dropna()] == ['-1.25', '3.4567']


def test_independent_table_and_column_without_like(tmp_path):
    import pandas as pd
    model = ExcelModelTemplate(FIXTURE)
    model.add_column('Table1', 'Amount', [-1.25, 0, 3.75])
    model.add_table('Sales', pd.DataFrame({'Quantity': [-1, 0, 4, 7], 'Price': [2.5, 0., 1.25, 3.75]}))
    output = model.save(tmp_path / 'independent.xlsx')
    with PBIXRay(output) as reader:
        assert reader.get_table('Table1').to_dict('list') == {'ID':[1,2,3], 'Label':['Alpha','Beta','Gamma'], 'Amount':[-1.25,0,3.75]}
        assert reader.get_table('Sales')['Quantity'].tolist() == [-1,0,4,7]


def test_independent_column_after_pending_resize(tmp_path):
    import pandas as pd
    model = ExcelModelTemplate(FIXTURE)
    model.replace_table('Table1', pd.DataFrame({'ID':[1,2,3,4], 'Label':['Alpha','Beta','Gamma','Delta']}))
    model.add_column('Table1', 'NewText', ['One','Two','Three','Four'])
    output = model.save(tmp_path / 'independent-growth.xlsx')
    with PBIXRay(output) as reader:
        assert reader.get_table('Table1')['NewText'].tolist() == ['One','Two','Three','Four']


def test_create_model_in_existing_workbook_preserves_other_parts(tmp_path):
    import pandas as pd
    from zipfile import ZipFile
    seed = Path(__file__).parents[1] / 'src/excel_model_lab/data/authoring-seed.xlsx'
    source = tmp_path / 'existing.xlsx'
    # An ordinary workbook, with a worksheet payload that must stay byte-exact.
    with ZipFile(seed) as original, ZipFile(source, 'w') as output:
        for part in original.infolist():
            if part.filename != 'xl/model/item.data':
                data = original.read(part.filename)
                if part.filename == 'xl/connections.xml':
                    data = b'<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
                output.writestr(part, data)
    model = ExcelModelTemplate.create({'NewTable': pd.DataFrame({'Number':[10,20,30]})}, workbook=source)
    result = model.save(tmp_path / 'inserted.xlsx')
    with ZipFile(source) as original, ZipFile(result) as output:
        assert output.read('xl/worksheets/sheet1.xml') == original.read('xl/worksheets/sheet1.xml')
        assert output.testzip() is None
    with PBIXRay(result) as reader:
        assert reader.get_table('NewTable')['Number'].tolist() == [10,20,30]


def test_datetime_seconds_roundtrip_and_edit(tmp_path):
    import pandas as pd
    values = pd.to_datetime(['2026-01-01 06:12:13', None, '2026-01-02 18:23:59'])
    model = ExcelModelTemplate.create({'Times': pd.DataFrame({'Stamp': values})})
    output = model.save(tmp_path / 'seconds.xlsx')
    reopened = ExcelModelTemplate(output)
    with reopened._open_reader() as reader:
        assert reader.get_table('Times')['Stamp'].tolist() == list(values)
    changed = pd.to_datetime(['2026-01-01 06:12:14', None, '2026-01-02 18:24:00'])
    second = reopened.replace_column('Times', 'Stamp', changed).save(tmp_path / 'edited.xlsx')
    with ExcelModelTemplate(second)._open_reader() as reader:
        assert reader.get_table('Times')['Stamp'].tolist() == list(changed)


@pytest.mark.parametrize('value, message', [
    ('2026-01-01T12:00:00Z', 'timezone-naive'),
    ('2026-01-01T12:00:00.123', 'whole-second'),
])
def test_datetime_unsupported_precision_leaves_model_unchanged(value, message):
    import pandas as pd
    from excel_model_lab.abf import UnsupportedModelError
    model = ExcelModelTemplate(FIXTURE)
    before = model._image.data
    with pytest.raises((ValueError, UnsupportedModelError), match=message):
        model.change_column_type('Table1', 'ID', 'datetime', values=[pd.Timestamp(value)] * 3)
    assert model._image.data == before
    assert not model._pending


def test_large_numeric_ids_are_rebased_below_signed_boundary(tmp_path):
    import pandas as pd
    output = ExcelModelTemplate.create({'Numbers': pd.DataFrame({'Value': [3_900_000_000, 3_900_000_001]})}).save(tmp_path / 'large.xlsx')
    reopened = ExcelModelTemplate(output)
    encoding = reopened._encodings[('Numbers', 'Value')]
    assert encoding.encode([3_900_000_000, 3_900_000_001]) == [3, 4]
    with reopened._open_reader() as reader:
        assert reader.get_table('Numbers')['Value'].tolist() == [3_900_000_000, 3_900_000_001]


@pytest.mark.parametrize('first', [None, '2026-01-01 00:00:00'])
def test_datetime_precision_uses_entire_column(tmp_path, first):
    import pandas as pd
    values = pd.to_datetime([first, '2026-01-01 12:00:01', '2026-01-01 23:59:59'])
    output = ExcelModelTemplate.create({'Times': pd.DataFrame({'Stamp': values})}).save(tmp_path / 'later-time.xlsx')
    with ExcelModelTemplate(output)._open_reader() as reader:
        actual = reader.get_table('Times')['Stamp']
        assert actual.isna().tolist() == list(pd.isna(values))
        assert actual.dropna().tolist() == list(values.dropna())


def test_spaced_unicode_names_and_all_null_columns(tmp_path):
    import pandas as pd
    data = pd.DataFrame({'Order Amount':[1,2,3], 'Coût (€)':[1.25,2.5,3.75],
        'Blank integer':pd.array([None]*3,dtype='Int64'),
        'Blank text':pd.array([None]*3,dtype='string'),
        'Blank boolean':pd.array([None]*3,dtype='boolean'),
        'Blank date':pd.to_datetime([None]*3)})
    output = ExcelModelTemplate.create({'Order details':data}).save(tmp_path/'names-null.xlsx')
    with ExcelModelTemplate(output)._open_reader() as reader:
        actual=reader.get_table('Order details')
        assert actual.columns.tolist()==data.columns.tolist()
        assert actual['Coût (€)'].tolist()==[1.25,2.5,3.75]
        assert actual.iloc[:,2:].isna().all().all()


def test_empty_creation_populate_and_clear(tmp_path):
    import pandas as pd
    data=pd.DataFrame({'Order ID':pd.array([],dtype='Int64'),'Label':pd.array([],dtype='string')})
    output=ExcelModelTemplate.create({'Empty table':data}).save(tmp_path/'empty.xlsx')
    model=ExcelModelTemplate(output)
    populated=model.replace_table('Empty table',pd.DataFrame({'Order ID':[7,8],'Label':['Seven','Eight']})).save(tmp_path/'populated.xlsx')
    with ExcelModelTemplate(populated)._open_reader() as reader:
        assert reader.get_table('Empty table').to_dict('list')=={'Order ID':[7,8],'Label':['Seven','Eight']}
    cleared=ExcelModelTemplate(populated).replace_table('Empty table',data).save(tmp_path/'cleared.xlsx')
    with ExcelModelTemplate(cleared)._open_reader() as reader:
        assert reader.get_table('Empty table').empty


def test_empty_native_storage_and_mixed_type_population(tmp_path):
    import struct
    import pandas as pd
    from decimal import Decimal
    from excel_model_lab.abf import AbfImage
    populated = pd.DataFrame({'ID': pd.array([7, 8], dtype='Int64'),
        'Text': pd.array(['Seven', None], dtype='string'),
        'Flag': pd.array([True, None], dtype='boolean'),
        'Real': [1.25, 2.5], 'Money': [Decimal('3.25'), Decimal('4.50')],
        'Stamp': pd.to_datetime(['2026-01-01 12:00:01', None])})
    first = ExcelModelTemplate.create({'Types': populated}).save(tmp_path/'first.xlsx')
    empty = ExcelModelTemplate(first).replace_table('Types', populated.iloc[:0]).save(tmp_path/'empty.xlsx')
    image = AbfImage.from_workbook(empty)
    streams = [image.read_logical_file(f) for f in image.files if f.logical_path
               and f.logical_path.endswith('.idf') and 'H$' not in f.logical_path]
    ordinary = [data for data in streams if len(data) == 152]
    assert len(ordinary) == 6
    assert all(data == struct.pack('<Q', 16) + bytes(128) + struct.pack('<QQ', 1, 0) for data in ordinary)
    model = ExcelModelTemplate(empty)
    output = model.replace_table('Types', populated).save(tmp_path/'populated.xlsx')
    actual = ExcelModelTemplate(output).get_table('Types')
    assert actual['ID'].tolist() == [7, 8]
    assert actual['Stamp'].iloc[0] == populated['Stamp'].iloc[0]
    assert actual['Real'].tolist() == [1.25, 2.5]
    assert actual['Money'].tolist() == [3.25, 4.5]
    assert actual['Flag'].isna().tolist() == [False, True]
    # A later invalid table must roll back reconstruction of an earlier empty table.
    model = ExcelModelTemplate(empty)
    before = model._image.data
    with pytest.raises(KeyError):
        model.replace_tables({'Types': populated, 'Missing': populated})
    assert model._image.data == before
    assert not model._pending
