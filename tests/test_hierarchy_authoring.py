import struct

import pandas as pd
import pytest

from excel_model_lab import ExcelModelTemplate
from excel_model_lab.hierarchy_storage import stored_hierarchies, traversal_indexes


def test_traversal_groups_repeated_children_by_parent():
    offsets, arrays = traversal_indexes([[4, 3, 4, 3, 3], [7, 7, 8, 9, 9]])
    assert offsets == [0, 2]
    assert list(arrays[0]) == [3, 4, 7, 9, 7, 8]
    assert list(arrays[1]) == [-1, -1, 0, 0, 1, 1]
    assert list(arrays[2]) == [2, 4, -1, -1, -1, -1]
    assert list(arrays[3]) == [2, 2, -1, -1, -1, -1]


def test_hierarchy_tracks_sort_column_changes(tmp_path):
    model = ExcelModelTemplate.create({'T': pd.DataFrame({
        'Label': ['Gamma', 'Beta', 'Alpha'], 'Rank': [1, 2, 3], 'ID': [1, 2, 3]})})
    model.add_hierarchy('T', 'H', ['Label', 'ID'])
    model.set_sort_column('T', 'Label', 'Rank')
    model = ExcelModelTemplate(model.save(tmp_path / 'sorted.xlsx'))
    def roots(current):
        file = next(f for f in current._image.files if f.logical_path and '.MULTI_LEVEL_ID.' in f.logical_path)
        ids = struct.unpack_from('<3I', current._image.read_logical_file(file), 8)
        members = current._encodings['T', 'Label'].original_values
        return tuple(members[value - 3] for value in ids)
    assert roots(model) == ('Gamma', 'Beta', 'Alpha')
    model.replace_column('T', 'Rank', [3, 2, 1])
    model = ExcelModelTemplate(model.save(tmp_path / 'changed.xlsx'))
    assert roots(model) == ('Alpha', 'Beta', 'Gamma')
    model.set_sort_column('T', 'Label', None)
    assert roots(model) == ('Alpha', 'Beta', 'Gamma')


def test_hierarchy_survives_save_and_rebuilds_on_resize(tmp_path):
    model = ExcelModelTemplate.create({'Products': pd.DataFrame({'Category': ['A', 'B', 'A'], 'ID': [1, 2, 3]})})
    model.add_hierarchy('Products', 'Catalog', ['Category', 'ID'])
    output = model.save(tmp_path / 'hierarchy.xlsx')
    result = ExcelModelTemplate(output)
    assert stored_hierarchies(result) == [('Products', 'Catalog', ['Category', 'ID'])]
    result.replace_table('Products', pd.DataFrame({'Category': ['A', 'A', 'A', 'A'], 'ID': [1, 2, 3, 4]}))
    result = ExcelModelTemplate(result.save(tmp_path / 'resized.xlsx'))
    file = next(f for f in result._image.files if f.logical_path and '.PARENT_POS.' in f.logical_path)
    payload = result._image.read_logical_file(file)
    assert struct.unpack('<5I', payload[8:28]) == (0, 1, 1, 1, 1)
    result.remove_hierarchy('Products', 'Catalog')
    result = ExcelModelTemplate(result.save(tmp_path / 'removed.xlsx'))
    assert stored_hierarchies(result) == []
    assert not any(f.logical_path and 'U$' in f.logical_path for f in result._image.files)
    with result._open_reader() as reader:
        assert reader.get_table('Products').shape == (4, 2)


def test_hierarchy_rejects_missing_duplicate_and_unsafe_names():
    model = ExcelModelTemplate.create({'T': pd.DataFrame({'ID': [1, 2]})})
    for name, levels in [('bad/name', ['ID']), ('H', []), ('H', ['ID', 'ID'])]:
        with pytest.raises(ValueError):
            model.add_hierarchy('T', name, levels)
    with pytest.raises(KeyError):
        model.add_hierarchy('T', 'H', ['Missing'])
    model.add_hierarchy('T', 'H', ['ID'])
    with pytest.raises(ValueError, match='already exists'):
        model.add_hierarchy('T', 'h', ['ID'])


def test_hierarchy_removal_rejects_presentation_dependencies(tmp_path):
    from zipfile import ZipFile
    from excel_model_lab.dependencies import DependencyError
    model = ExcelModelTemplate.create({'T': pd.DataFrame({'ID': [1, 2]})})
    model.add_hierarchy('T', 'H', ['ID'])
    path = model.save(tmp_path / 'presentation.xlsx')
    with ZipFile(path, 'a') as archive:
        archive.writestr('xl/pivotTables/pivotTable1.xml', '<pivotTableDefinition/>')
    model = ExcelModelTemplate(path)
    before = model._image.data
    with pytest.raises(DependencyError) as error:
        model.remove_hierarchy('T', 'H')
    assert error.value.issues[0].object_path == ('tables', 'T', 'hierarchies', 'H')
    assert error.value.issues[0].dependencies == (('workbook_parts', 'xl/pivotTables/pivotTable1.xml'),)
    assert model._image.data == before


def test_hierarchy_update_preserves_settings_and_pending_values(tmp_path):
    from xml.dom import minidom
    from excel_model_lab.schema import _elements, _child, _text
    model = ExcelModelTemplate.create({'T': pd.DataFrame({
        'Region': ['West', 'East'], 'Category': ['A', 'B'], 'ID': [1, 2]})})
    model.add_hierarchy('T', 'Catalog', ['Category', 'ID'])
    model.configure_hierarchy('T', 'Catalog', hidden=True, description='Product catalog',
                              level_names={'ID': 'Product number'})
    model.replace_column('T', 'ID', [10, 20])
    model.update_hierarchy('T', 'Catalog', new_name='Geography', levels=['Region', 'ID', 'Category'])
    model = ExcelModelTemplate(model.save(tmp_path / 'updated.xlsx'))
    assert stored_hierarchies(model) == [('T', 'Geography', ['Region', 'ID', 'Category'])]
    doc = minidom.parseString(model._image.read_logical_file(model._dimension_paths['T']))
    hierarchy = _elements(doc, 'Hierarchy')[0]
    assert _text(_child(hierarchy, 'Description')) == 'Product catalog'
    assert [_text(_child(level, 'Name')) for level in _elements(hierarchy, 'Level')] == [
        'Region', 'Product number', 'Category']
    cube = next(f for f in model._image.files if f.logical_path and f.logical_path.endswith('.cub.xml'))
    doc = minidom.parseString(model._image.read_logical_file(cube))
    hierarchy = next(h for h in _elements(doc, 'Hierarchy') if _text(_child(h, 'HierarchyID')) == 'Geography')
    assert _text(_child(hierarchy, 'Visible')) == 'false'
    with model._open_reader() as reader:
        assert reader.get_table('T')['ID'].tolist() == [10, 20]
    assert not any(f.logical_path and 'U$' in f.logical_path and '$Catalog.' in f.logical_path
                   for f in model._image.files)


def test_hierarchy_edit_failures_are_atomic():
    model = ExcelModelTemplate.create({'T': pd.DataFrame({'ID': [1, 2], 'Label': ['A', 'B']})})
    model.add_hierarchy('T', 'H', ['Label', 'ID'])
    before = model._image.data
    for kwargs in ({'levels': ['Missing']}, {'new_name': 'bad/name'}):
        with pytest.raises((KeyError, ValueError)):
            model.update_hierarchy('T', 'H', **kwargs)
        assert model._image.data == before
    for kwargs in ({'hidden': 'yes'}, {'level_names': {'ID': 'Label'}}, {'description': 7}):
        with pytest.raises((TypeError, ValueError)):
            model.configure_hierarchy('T', 'H', **kwargs)
        assert model._image.data == before
