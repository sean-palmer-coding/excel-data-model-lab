from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import pytest
from excel_model_lab import WorkbookModel
from excel_model_lab.native_sources import LinkedTableSource, OleDbSource
from excel_model_lab.dependencies import DependencyError

FIXTURE = Path(__file__).parent / 'data' / 'native-authoring-base.xlsx'


def test_linked_table_binding_preserves_cells_and_snapshot_data(tmp_path):
    model = WorkbookModel(FIXTURE)
    assert model.list_native_sources()[0].kind == 'linked_table'
    model.bind_source('Table1', LinkedTableSource('Table1'))
    model.replace_column('Table1', 'ID', [10, 20, 30])
    output = model.save(tmp_path / 'linked.xlsx')
    assert WorkbookModel(output).get_table('Table1')['ID'].tolist() == [10, 20, 30]
    binding = WorkbookModel(output).list_native_sources()[0]
    assert binding.kind == 'linked_table' and binding.source_name == 'Table1'
    with ZipFile(FIXTURE) as before, ZipFile(output) as after:
        assert before.read('xl/worksheets/sheet1.xml') == after.read('xl/worksheets/sheet1.xml')
        assert before.read('xl/tables/table1.xml') == after.read('xl/tables/table1.xml')
        connections = ET.fromstring(after.read('xl/connections.xml'))
        linked = next(c for c in connections if c.get('type') == '102')
        assert linked.get('refreshOnLoad') == '0'


def test_unbind_allows_explicit_snapshot_schema_edits(tmp_path):
    model = WorkbookModel(FIXTURE)
    with pytest.raises(DependencyError):
        model.remove_column('Table1', 'Label')
    model.unbind_source('Table1').remove_column('Table1', 'Label')
    output = model.save(tmp_path / 'snapshot.xlsx')
    assert WorkbookModel(output).list_native_sources()[0].kind == 'snapshot'
    with ZipFile(output) as archive:
        assert not any(c.get('type') == '102' for c in ET.fromstring(archive.read('xl/connections.xml')))


def test_linked_source_requires_existing_table_and_matching_schema():
    model = WorkbookModel(FIXTURE)
    with pytest.raises(KeyError):
        model.bind_source('Table1', LinkedTableSource('Missing'))
    model.unbind_source('Table1').remove_column('Table1', 'Label')
    with pytest.raises(ValueError, match='columns must match'):
        model.bind_source('Table1', LinkedTableSource('Table1'))


def test_native_credentials_are_not_embedded_or_shown_in_repr():
    with pytest.raises(ValueError, match='credentials'):
        OleDbSource('Provider=MSOLEDBSQL;Pwd=secret', 'SELECT 1')
    source = OleDbSource('Provider=MSOLEDBSQL;Server=local;Integrated Security=SSPI', 'SELECT 1')
    assert 'connection_string' not in repr(source)

def _rewrite_fixture(tmp_path, changes):
    destination = tmp_path / 'custom.xlsx'
    with ZipFile(FIXTURE) as source, ZipFile(destination, 'w') as target:
        for item in source.infolist():
            target.writestr(item, changes.get(item.filename, source.read(item)))
        for name, payload in changes.items():
            if name not in source.namelist():
                target.writestr(name, payload)
    return destination


def test_binding_preserves_unrelated_connections_and_unknown_extensions(tmp_path):
    with ZipFile(FIXTURE) as source:
        connections = ET.fromstring(source.read('xl/connections.xml'))
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    linked = next(c for c in connections if c.get('type') == '102')
    linked.set('description', 'Retain this description')
    ext_list = linked.find(ns + 'extLst')
    unknown = ET.SubElement(ext_list, ns + 'ext', {'uri': 'vendor-test'})
    ET.SubElement(unknown, '{urn:vendor}payload', {'keep': 'yes'})
    unrelated = ET.SubElement(connections, ns + 'connection', {'id': '99', 'name': 'Other', 'type': '5'})
    ET.SubElement(unrelated, ns + 'dbPr', {'connection': 'opaque', 'command': 'unchanged'})
    source = _rewrite_fixture(tmp_path, {'xl/connections.xml': ET.tostring(connections)})
    model = WorkbookModel(source).bind_source('Table1', LinkedTableSource('Table1'))
    output = model.save(tmp_path / 'bound.xlsx')
    with ZipFile(output) as archive:
        result = ET.fromstring(archive.read('xl/connections.xml'))
    assert ET.tostring(next(c for c in result if c.get('name') == 'Other')) == ET.tostring(unrelated)
    rebound = next(c for c in result if c.get('type') == '102')
    assert rebound.get('description') == 'Retain this description'
    assert rebound.find('.//{urn:vendor}payload').get('keep') == 'yes'


def test_binding_rejects_shared_worksheet_connection(tmp_path):
    with ZipFile(FIXTURE) as archive:
        connections = ET.fromstring(archive.read('xl/connections.xml'))
    linked = next(c for c in connections if c.get('type') == '102')
    source = _rewrite_fixture(tmp_path, {'xl/queryTables/shared.xml':
        ('<queryTable connectionId="' + linked.get('id') + '"/>').encode()})
    model = WorkbookModel(source).bind_source('Table1', LinkedTableSource('Table1'))
    from excel_model_lab.abf import UnsupportedModelError
    with pytest.raises(UnsupportedModelError, match='referenced by workbook part'):
        model.save(tmp_path / 'rejected.xlsx')
    assert not (tmp_path / 'rejected.xlsx').exists()

def test_binding_preserves_connection_names_and_existing_defined_names(tmp_path):
    with ZipFile(FIXTURE) as archive:
        original_connections = ET.fromstring(archive.read('xl/connections.xml'))
        original_workbook = ET.fromstring(archive.read('xl/workbook.xml'))
    linked = next(c for c in original_connections if c.get('type') == '102')
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    names = {n.get('name'):n.text for n in original_workbook.findall('.//'+ns+'definedName')}
    first = WorkbookModel(FIXTURE).bind_source('Table1',LinkedTableSource('Table1')).save(tmp_path/'first.xlsx')
    second = WorkbookModel(first).bind_source('Table1',LinkedTableSource('Table1')).save(tmp_path/'second.xlsx')
    assert WorkbookModel(second).list_native_sources()[0].connection_name == linked.get('name')
    with ZipFile(second) as archive:
        connections = ET.fromstring(archive.read('xl/connections.xml'))
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
    assert next(c for c in connections if c.get('type')=='102').get('name') == linked.get('name')
    result = workbook.findall('.//'+ns+'definedName')
    assert len({n.get('name') for n in result}) == len(result)
    assert names.items() <= {n.get('name'):n.text for n in result}.items()
