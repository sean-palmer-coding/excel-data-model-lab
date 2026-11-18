import pandas as pd
import pytest
from excel_model_lab import WorkbookModel
from excel_model_lab.dependencies import DependencyError


def test_public_perspective_crud_and_dependencies(tmp_path):
    m = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [10, 20], 'Key': [1, 2]})})
    m.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')
    m.add_hierarchy('Sales', 'Detail', ['Key'])
    m.create_perspective('Finance', columns=[('Sales', 'Amount')], measures=[('Sales', 'Total')],
                         hierarchies=[('Sales', 'Detail')], description='Finance team')
    assert m.list_perspectives() == ('Finance',)
    with pytest.raises(DependencyError):
        m.remove_hierarchy('Sales', 'Detail')
    path = m.save(tmp_path / 'perspective.xlsx')
    reopened = WorkbookModel(path)
    assert reopened.list_perspectives() == ('Finance',)
    before = [(f.storage_path, f.logical_path) for f in reopened._image.files if f.logical_path and f.logical_path.endswith('.persp.xml')]
    reopened.rename_perspective('Finance', 'Reporting')
    assert reopened.list_perspectives() == ('Reporting',)
    assert [(f.storage_path, f.logical_path) for f in reopened._image.files if f.logical_path and f.logical_path.endswith('.persp.xml')] == before
    with pytest.raises(DependencyError):
        reopened.remove_hierarchy('Sales', 'Detail')
    reopened.update_perspective('Reporting', columns=[('Sales', 'Key')])
    reopened.remove_hierarchy('Sales', 'Detail')
    reopened.remove_perspective('Reporting')
    assert not reopened.list_perspectives()


def test_invalid_perspective_does_not_commit_pending_data():
    m = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [10, 20]})})
    pending = dict(m._pending)
    image = m._image.data
    with pytest.raises((ValueError, KeyError)):
        m.create_perspective('Broken', columns=[('Sales', 'Missing')])
    assert m._image.data == image and m._pending == pending


def test_perspective_update_preserves_unknown_metadata():
    from xml.etree import ElementTree as ET
    from excel_model_lab.abf import AbfImage
    from excel_model_lab.perspectives import set_perspective, _tag
    m = WorkbookModel.create({'T': pd.DataFrame({'A': [1], 'B': [2]})})
    m.create_perspective('View', columns=[('T', 'A')])
    f = next(f for f in m._image.files if f.logical_path and f.logical_path.endswith('.persp.xml'))
    doc = ET.fromstring(m._image.read_logical_file(f))
    perspective = doc.find('.//' + _tag('Perspective'))
    doc.set('{urn:external}wrapper', 'keep')
    perspective.set('{urn:external}attribute', 'keep')
    ET.SubElement(perspective.find(_tag('Actions')), '{urn:external}Action').text = 'keep'
    ET.SubElement(perspective, '{urn:external}Extension', value='preserve')
    ET.SubElement(perspective.find(_tag('Translations')), _tag('Translation')).text = 'localized'
    image = AbfImage(m._image.assemble({f.storage_path: ET.tostring(doc)}))
    updated = set_perspective(image, 'View', columns=[('T', 'B')], replace=True)
    result = ET.fromstring(updated.read_logical_file(next(item for item in updated.files if item.storage_path == f.storage_path)))
    assert result.find('.//{urn:external}Extension').get('value') == 'preserve'
    assert result.find('.//' + _tag('Translation')).text == 'localized'
    assert result.get('{urn:external}wrapper') == 'keep'
    assert result.find('.//' + _tag('Perspective')).get('{urn:external}attribute') == 'keep'
    assert result.find('.//{urn:external}Action').text == 'keep'


def test_renamed_perspective_caption_can_be_reused_without_identity_collision(tmp_path):
    from xml.etree import ElementTree as ET
    from excel_model_lab.perspectives import _tag
    model = WorkbookModel.create({'T': pd.DataFrame({'A': [1]})})
    model.create_perspective('Original', columns=[('T', 'A')])
    model.rename_perspective('Original', 'Renamed')
    model.create_perspective('original', columns=[('T', 'A')])
    files = [f for f in model._image.files if f.logical_path and f.logical_path.endswith('.persp.xml')]
    ids = [ET.fromstring(model._image.read_logical_file(f)).findtext('.//' + _tag('Perspective') + '/' + _tag('ID')) for f in files]
    assert len({value.casefold() for value in ids}) == 2
    assert len({f.logical_path.casefold() for f in files}) == 2
    output = model.save(tmp_path / 'reused.xlsx')
    assert set(WorkbookModel(output).list_perspectives()) == {'Renamed', 'original'}


def test_perspective_rejects_file_list_delimiter_atomically():
    model = WorkbookModel.create({'T': pd.DataFrame({'A': [1]})})
    before = model._image.data
    with pytest.raises(ValueError):
        model.create_perspective('One;Two', columns=[('T', 'A')])
    assert model._image.data == before


def test_update_preserves_retained_member_extensions_without_restoring_removed_members():
    from xml.etree import ElementTree as ET
    from excel_model_lab.abf import AbfImage
    from excel_model_lab.perspectives import set_perspective, _tag
    model = WorkbookModel.create({'T': pd.DataFrame({'A': [1], 'B': [2]})})
    model.add_measure('T', 'Total', 'SUM(T[A])')
    model.create_perspective('View', columns=[('T', 'A'), ('T', 'B')], measures=[('T', 'Total')])
    file = next(f for f in model._image.files if f.logical_path and f.logical_path.endswith('.persp.xml'))
    root = ET.fromstring(model._image.read_logical_file(file))
    for tag in ('Dimensions', 'Dimension', 'Attributes', 'Attribute', 'Calculation', 'MeasureGroup'):
        node = root.find('.//' + _tag(tag))
        node.set('{urn:external}flag', tag)
        ET.SubElement(node, '{urn:external}Extension').text = tag
    image = AbfImage(model._image.assemble({file.storage_path: ET.tostring(root)}))
    result = set_perspective(image, 'View', columns=[('T', 'A')], measures=[('T', 'Total')], replace=True)
    updated = ET.fromstring(result.read_logical_file(file.storage_path))
    for tag in ('Dimensions', 'Dimension', 'Attributes', 'Attribute', 'Calculation', 'MeasureGroup'):
        node = updated.find('.//' + _tag(tag))
        assert node.get('{urn:external}flag') == tag
        assert node.find('{urn:external}Extension').text == tag
    assert [a.text for a in updated.findall('.//' + _tag('AttributeID'))] == ['A']


@pytest.mark.parametrize('members, base_type', [(None, TypeError), ([([], 'A')], TypeError),
                                              ([('T', 'Missing')], KeyError), ([], ValueError)])
def test_direct_perspective_errors_are_structured_and_compatible(members, base_type):
    from excel_model_lab import ModelValidationError
    model = WorkbookModel.create({'T': pd.DataFrame({'A': [1]})})
    before = model._image.data
    with pytest.raises(base_type) as error:
        model.create_perspective('View', columns=members)
    assert isinstance(error.value, ModelValidationError)
    assert error.value.issues[0].object_path == ('perspectives', 'View')
    assert model._image.data == before


def test_multiple_hierarchies_roundtrip_and_precise_dependencies(tmp_path):
    from excel_model_lab import ModelReader
    from excel_model_lab.hierarchy_storage import resolve_hierarchy
    model = WorkbookModel.create({'T': pd.DataFrame({'A': [1, 2], 'B': [3, 4]})})
    for name, levels in [('First & detail', ['A']), ('Second detail', ['B']), ('Unselected', ['A', 'B'])]:
        model.add_hierarchy('T', name, levels)
    model.create_perspective('View & More', columns=[('T', 'A')],
        hierarchies=[('T', 'First & detail'), ('T', 'Second detail')])
    for name in ('First & detail', 'Second detail'):
        with pytest.raises(DependencyError):
            model.remove_hierarchy('T', name)
    model.remove_hierarchy('T', 'Unselected')
    expected_ids = {resolve_hierarchy(model, 'T', name)[0]
                    for name in ('First & detail', 'Second detail')}
    path = model.save(tmp_path / 'multiple.xlsx')
    with ModelReader(path) as reader:
        metadata = reader.excel_metadata
        assert set(metadata.perspective_hierarchies.hierarchy_id) == expected_ids
        assert metadata.perspective_hierarchy_annotations.iloc[0].empty_hierarchy_names == 'First & detail,Second detail'


def test_hierarchy_annotation_rejects_ambiguous_captions_atomically():
    model = WorkbookModel.create({'T': pd.DataFrame({'A': [1, 2]})})
    model.add_hierarchy('T', 'One,Two', ['A'])
    before = model._image.data
    with pytest.raises(ValueError, match='commas'):
        model.create_perspective('View', hierarchies=[('T', 'One,Two')])
    assert model._image.data == before


def test_apostrophe_caption_allowed_without_annotation_but_guarded_with_it():
    model = WorkbookModel.create({'T': pd.DataFrame({'A': [1, 2]})})
    model.add_hierarchy('T', 'Detail', ['A'])
    model.create_perspective("Owner's View", columns=[('T', 'A')])
    before = model._image.data
    with pytest.raises(ValueError, match='apostrophes'):
        model.create_perspective('Hierarchy', hierarchies=[('T', 'Detail')])
    assert model._image.data == before
    model.rename_perspective("Owner's View", 'Owners View')
    model.create_perspective('Hierarchy', hierarchies=[('T', 'Detail')])
    before = model._image.data
    with pytest.raises(ValueError, match='apostrophes'):
        model.rename_perspective('Owners View', "Owner's View")
    assert model._image.data == before


def test_hierarchy_annotation_update_preserves_extensions():
    from xml.dom import minidom
    from xml.etree import ElementTree as ET
    from excel_model_lab.perspectives import _update_hierarchy_annotation
    xml = '<Cube><Annotations><Annotation><Name>PerspectiveEmptyHierarchyNames</Name><Value/></Annotation></Annotations></Cube>'
    doc = minidom.parseString(xml)
    payload = '<Perspectives extra="root"><Perspective Name="View" extra="perspective"><Tables extra="tables"><Table Name="T" EmptyHierarchyNames="First" extra="table"><Extension/></Table></Tables></Perspective></Perspectives>'
    doc.getElementsByTagName('Value')[0].appendChild(doc.createTextNode(payload))
    _update_hierarchy_annotation(doc, doc.documentElement, 'View', [('T', 'id', 'First', 'first'), ('T', 'id', 'Second', 'second')])
    result = ET.fromstring(doc.getElementsByTagName('Value')[0].firstChild.data)
    assert result.get('extra') == 'root'
    assert result.find('Perspective').get('extra') == 'perspective'
    assert result.find('Perspective/Tables').get('extra') == 'tables'
    table = result.find('Perspective/Tables/Table')
    assert table.get('extra') == 'table' and table.find('Extension') is not None
    assert table.get('EmptyHierarchyNames') == 'First,Second'
