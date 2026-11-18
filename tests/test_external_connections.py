from dataclasses import replace
from decimal import Decimal
from pathlib import Path
from types import ModuleType
from xml.dom import minidom
from zipfile import ZipFile
import pytest
import sys

from excel_model_lab.errors import ModelValidationError
from excel_model_lab.external_connections import (
    ExternalModelConnection, ExternalQueryError, ExternalQueryResult, AdomdQueryClient,
    add_external_connection, inspect_external_connections, S, R, C,
)


def source(tmp_path, *, additions=None):
    path = tmp_path / 'source.xlsx'
    seed = Path(__file__).parents[1] / 'src/excel_model_lab/data/authoring-seed.xlsx'
    with ZipFile(seed) as original, ZipFile(path, 'w') as archive:
        for info in original.infolist():
            archive.writestr(info, original.read(info))
        for name, data in (additions or {}).items():
            archive.writestr(name, data)
        archive.comment = b'preserved comment'
    return path


def spec(**kw):
    return replace(ExternalModelConnection('Corporate model', 'localhost', 'Sales'), **kw)


def test_add_preserves_embedded_model_and_unrelated_package(tmp_path):
    original = source(tmp_path, additions={'customXml/preserved.bin': b'\x00unchanged\xff'})
    output = add_external_connection(original, tmp_path/'external.xlsx', spec())
    with ZipFile(original) as before, ZipFile(output) as after:
        assert after.comment == before.comment
        changed = [name for name in before.namelist() if before.read(name) != after.read(name)]
        assert changed == ['xl/connections.xml']
        root = minidom.parseString(after.read('xl/connections.xml'))
        nodes = root.getElementsByTagNameNS(S, 'connection')
        added = next(n for n in nodes if n.getAttribute('name') == 'Corporate model')
        assert added.getAttribute('refreshOnLoad') == '0'
        assert added.getAttribute('savePassword') == '0'
        assert added.getElementsByTagNameNS(S, 'dbPr')[0].getAttribute('commandType') == '1'
        assert len(added.getElementsByTagNameNS(S, 'olapPr')) == 1
    result = inspect_external_connections(output)
    assert [(r.name, r.cube) for r in result] == [('Corporate model', 'Model')]
    assert not hasattr(result[0], 'connection_string')


def test_plain_workbook_gets_relationship_and_content_type(tmp_path):
    original = tmp_path/'plain.xlsx'
    with ZipFile(original, 'w') as archive:
        archive.writestr('xl/workbook.xml', f'<workbook xmlns="{S}"/>')
        archive.writestr('[Content_Types].xml', f'<Types xmlns="{C}"/>')
    assert inspect_external_connections(original) == ()
    add_external_connection(original, original, spec())
    with ZipFile(original) as archive:
        rels = minidom.parseString(archive.read('xl/_rels/workbook.xml.rels'))
        assert rels.getElementsByTagNameNS(R, 'Relationship')[0].getAttribute('Target') == 'connections.xml'
        types = minidom.parseString(archive.read('[Content_Types].xml'))
        assert types.getElementsByTagNameNS(C, 'Override')[0].getAttribute('PartName') == '/xl/connections.xml'


def test_prefixed_package_documents_keep_new_nodes_in_namespace(tmp_path):
    original = tmp_path/'prefixed.xlsx'
    relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections'
    with ZipFile(original, 'w') as archive:
        archive.writestr('xl/workbook.xml', f'<workbook xmlns="{S}"/>')
        archive.writestr('[Content_Types].xml', f'<t:Types xmlns:t="{C}"/>')
        archive.writestr('xl/_rels/workbook.xml.rels',
            f'<r:Relationships xmlns:r="{R}"><r:Relationship Id="rId1" Type="{relationship}" Target="connections.xml"/></r:Relationships>')
        archive.writestr('xl/connections.xml', f'<x:connections xmlns:x="{S}"/>')
    output = add_external_connection(original, tmp_path/'out.xlsx', spec())
    assert len(inspect_external_connections(output)) == 1
    with ZipFile(output) as archive:
        connections = minidom.parseString(archive.read('xl/connections.xml'))
        assert len(connections.getElementsByTagNameNS(S, 'dbPr')) == 1
        types = minidom.parseString(archive.read('[Content_Types].xml'))
        assert len(types.getElementsByTagNameNS(C, 'Override')) == 1


def test_existing_alternate_connection_part_and_extension_preserved(tmp_path):
    original = source(tmp_path)
    with ZipFile(original) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    parts['xl/connections-other.xml'] = parts.pop('xl/connections.xml').replace(
        b'</connections>', b'<extLst xmlns:v="urn:vendor"><ext uri="test"><v:payload v:key="unchanged"/></ext></extLst></connections>')
    parts['xl/_rels/workbook.xml.rels'] = parts['xl/_rels/workbook.xml.rels'].replace(b'Target="connections.xml"', b'Target="connections-other.xml"')
    parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(b'/xl/connections.xml', b'/xl/connections-other.xml')
    with ZipFile(original, 'w') as archive:
        for n, data in parts.items(): archive.writestr(n, data)
    out = add_external_connection(original, tmp_path/'added.xlsx', spec())
    assert inspect_external_connections(out)[0].part == 'xl/connections-other.xml'
    with ZipFile(out) as archive:
        root = minidom.parseString(archive.read('xl/connections-other.xml'))
        assert root.getElementsByTagNameNS('urn:vendor', 'payload')[0].getAttribute('v:key') == 'unchanged'
        assert [n.localName for n in root.documentElement.childNodes if n.nodeType == n.ELEMENT_NODE][-1] == 'extLst'


def test_duplicate_name_is_case_insensitive_and_atomic(tmp_path):
    original = source(tmp_path)
    destination = tmp_path/'out.xlsx'
    destination.write_bytes(b'keep')
    with pytest.raises(ModelValidationError, match='reserved'):
        spec(name='thisworkbookdatamodel')
    first = add_external_connection(original, tmp_path/'first.xlsx', spec())
    with pytest.raises(ModelValidationError, match='already exists'):
        add_external_connection(first, destination, spec(name='CORPORATE MODEL'))
    assert destination.read_bytes() == b'keep'


@pytest.mark.parametrize('kwargs', [dict(name=''), dict(server='x\x00y'), dict(authentication='password'),
    dict(server='powerbi://api.powerbi.com/myorg/workspace'), dict(cube='\ud800')])
def test_invalid_connection_inputs(kwargs):
    with pytest.raises(ModelValidationError): spec(**kwargs)


def test_quoted_values_cannot_inject_connection_properties():
    value = 'server";Password=not-a-secret;'
    connection = spec(server=value, database='db;Catalog=other')
    assert 'Data Source="server"";Password=not-a-secret;";' in connection.connection_string()
    assert 'Initial Catalog="db;Catalog=other";' in connection.connection_string()
    assert 'Persist Security Info=False;' in connection.connection_string()


def test_cloud_authentication_is_deferred_to_provider():
    connection = spec(server='powerbi://api.powerbi.com/v1.0/myorg/Workspace', authentication='interactive')
    assert 'Integrated Security=' not in connection.connection_string()
    with pytest.raises(ExternalQueryError, match='caller-configured'):
        AdomdQueryClient(connection).query('EVALUATE ROW("X", 1)')


class Resource:
    disposed = False
    def Dispose(self): self.disposed = True


class Reader(Resource):
    FieldCount = 2
    def __init__(self, rows): self.rows = rows; self.index = -1
    def GetName(self, i): return ('[Value]', '[Label]')[i]
    def Read(self): self.index += 1; return self.index < len(self.rows)
    def IsDBNull(self, i): return self.rows[self.index][i] is None
    def GetValue(self, i): return self.rows[self.index][i]
    def NextResult(self): return False


class Command(Resource):
    def __init__(self, reader): self.reader = reader
    def ExecuteReader(self): return self.reader


class Connection(Resource):
    def __init__(self, command, error=None): self.command = command; self.error = error
    def Open(self):
        if self.error: raise self.error
    def CreateCommand(self): return self.command


def test_query_results_blank_values_timeout_and_resource_cleanup():
    reader = Reader([(10, 'A'), (None, 'B')]); command = Command(reader); connection = Connection(command)
    client = AdomdQueryClient(spec(), connection_factory=lambda: connection)
    result = client.query('EVALUATE Example', timeout=7)
    assert result.columns == ('[Value]', '[Label]')
    assert result.rows == ((10, 'A'), (None, 'B'))
    assert result.to_dataframe().shape == (2, 2)
    assert command.CommandTimeout == 7 and command.CommandText == 'EVALUATE Example'
    assert reader.disposed and command.disposed and connection.disposed


def test_dataframe_preserves_nullable_int64_and_decimal_exactly():
    large = 9007199254740993
    precise = Decimal('12345678901234567890.1234')
    result = ExternalQueryResult(('ID', 'Amount'), ((large, precise), (None, None)))
    frame = result.to_dataframe()
    assert frame.iloc[0, 0] == large and isinstance(frame.iloc[0, 0], int)
    assert frame.iloc[0, 1] == precise and isinstance(frame.iloc[0, 1], Decimal)
    assert frame.iloc[1, 0] is None and frame.iloc[1, 1] is None
    assert all(dtype == object for dtype in frame.dtypes)


def test_invalid_provider_connection_string_disposes_created_connection(monkeypatch):
    class InvalidConnection(Resource):
        @property
        def ConnectionString(self):
            return ''
        @ConnectionString.setter
        def ConnectionString(self, value):
            raise ValueError('Provider rejected the connection string')
    connection = InvalidConnection()
    clr = ModuleType('clr')
    clr.AddReference = lambda name: None
    microsoft = ModuleType('Microsoft'); microsoft.__path__ = []
    services = ModuleType('Microsoft.AnalysisServices'); services.__path__ = []
    provider = ModuleType('Microsoft.AnalysisServices.AdomdClient')
    provider.AdomdConnection = lambda: connection
    for name, module in [('clr', clr), ('Microsoft', microsoft),
                         ('Microsoft.AnalysisServices', services),
                         ('Microsoft.AnalysisServices.AdomdClient', provider)]:
        monkeypatch.setitem(sys.modules, name, module)
    with pytest.raises(ExternalQueryError) as failure:
        AdomdQueryClient(spec()).query('EVALUATE ROW("X", 1)')
    assert isinstance(failure.value.__cause__, ValueError)
    assert connection.disposed


def test_query_row_bound_is_not_silent_truncation_and_disposes():
    reader = Reader([(1, 'A'), (2, 'B')]); command = Command(reader); connection = Connection(command)
    client = AdomdQueryClient(spec(), connection_factory=lambda: connection)
    with pytest.raises(ExternalQueryError, match='max_rows'):
        client.query('EVALUATE Example', max_rows=1)
    assert reader.disposed and command.disposed and connection.disposed


def test_multiple_query_results_are_not_silently_discarded():
    class MultiReader(Reader):
        def __init__(self):
            super().__init__([(1, 'first')])
            self.additional_results = iter([[(2, 'second')]])
        def NextResult(self):
            self.rows = next(self.additional_results, None)
            self.index = -1
            return self.rows is not None
    reader = MultiReader(); command = Command(reader); connection = Connection(command)
    with pytest.raises(ExternalQueryError, match='Multiple result sets'):
        AdomdQueryClient(spec(), connection_factory=lambda: connection).query(
            'EVALUATE ROW("Value", 1) EVALUATE ROW("Value", 2)')
    assert reader.disposed and command.disposed and connection.disposed


def test_query_error_redacts_provider_message_and_disposes():
    connection = Connection(None, RuntimeError('Password=private'))
    with pytest.raises(ExternalQueryError) as failure:
        AdomdQueryClient(spec(), connection_factory=lambda: connection).query('EVALUATE Example')
    assert 'private' not in str(failure.value)
    assert 'private' in str(failure.value.__cause__)
    assert connection.disposed


@pytest.mark.parametrize('dax,kw', [('CREATE DATABASE x', {}), ('EVALUATE x', {'timeout': False}),
                                  ('EVALUATE x', {'max_rows': 0})])
def test_invalid_query_does_not_open_connection(dax, kw):
    def factory(): raise AssertionError('Must not open')
    with pytest.raises(ModelValidationError):
        AdomdQueryClient(spec(), connection_factory=factory).query(dax, **kw)


def test_signed_package_rejection_preserves_destination(tmp_path):
    original = source(tmp_path, additions={'_xmlsignatures/sig1.xml': b'<signature/>'})
    destination = tmp_path/'out.xlsx'; destination.write_bytes(b'keep')
    with pytest.raises(ModelValidationError, match='signed'):
        add_external_connection(original, destination, spec())
    assert destination.read_bytes() == b'keep'
