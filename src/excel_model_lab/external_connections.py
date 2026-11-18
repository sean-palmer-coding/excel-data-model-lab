"""External Analysis Services connections, separate from embedded-model imports.

Authoring never contacts a server. Live querying uses optional ADOMD.NET and a
running semantic endpoint; it does not turn an XLSX model into a server.
"""
from dataclasses import dataclass
from pathlib import Path
from tempfile import mkstemp
from xml.dom import minidom
from zipfile import ZipFile, ZIP_DEFLATED
import os
import posixpath
import re

from .errors import ModelValidationError

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'
CONNECTION_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections'
CONNECTION_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml'


def _text(value, label):
    if not isinstance(value, str) or not value.strip():
        raise ModelValidationError(f'{label} must be nonempty text')
    if any(not (ord(c) in (9, 10, 13) or 32 <= ord(c) <= 0xD7FF or
                0xE000 <= ord(c) <= 0xFFFD or 0x10000 <= ord(c) <= 0x10FFFF) for c in value):
        raise ModelValidationError(f'{label} contains invalid XML text')
    return value


def _quote(value):
    # Quoted OLE DB values cannot introduce additional connection properties.
    return '"' + value.replace('"', '""') + '"'


@dataclass(frozen=True)
class ExternalModelConnection:
    """Credential-free external cube definition; not an imported model source.

    ``integrated`` is intended for Windows-authenticated SSAS. ``interactive``
    delegates authentication to Excel/provider; it is not silent token login.
    """
    name: str
    server: str
    database: str
    cube: str = 'Model'
    authentication: str = 'integrated'
    description: str = 'External Analysis Services semantic model'

    def __post_init__(self):
        for label in ('name', 'server', 'database', 'cube', 'description'):
            _text(getattr(self, label), label)
        if self.name.casefold() == 'thisworkbookdatamodel':
            raise ModelValidationError('ThisWorkbookDataModel is reserved for the embedded model')
        if self.authentication not in ('integrated', 'interactive'):
            raise ModelValidationError('authentication must be integrated or interactive')
        if any(c in self.server for c in '\r\n\t'):
            raise ModelValidationError('server cannot contain control whitespace')
        if self.authentication == 'integrated' and self.server.lower().startswith(('powerbi://', 'asazure://')):
            raise ModelValidationError('Cloud semantic endpoints require interactive authentication or a caller-configured query connection')

    def connection_string(self):
        properties = ['Provider=MSOLAP', 'Data Source=' + _quote(self.server),
                      'Initial Catalog=' + _quote(self.database), 'Persist Security Info=False']
        if self.authentication == 'integrated':
            properties.append('Integrated Security=SSPI')
        return ';'.join(properties) + ';'


@dataclass(frozen=True)
class ExternalConnectionInfo:
    connection_id: int
    name: str
    cube: str
    part: str


def _children(node, namespace, name):
    return [n for n in node.childNodes if n.nodeType == n.ELEMENT_NODE
            and n.namespaceURI == namespace and n.localName == name]


def _node(doc, name, attrs, namespace=S):
    prefix = doc.documentElement.prefix if doc.documentElement.namespaceURI == namespace else None
    node = doc.createElementNS(namespace, f'{prefix}:{name}' if prefix else name)
    for key, value in attrs.items():
        node.setAttribute(key, str(value))
    return node


def _connection_document(parts, create=False):
    relpath = 'xl/_rels/workbook.xml.rels'
    if relpath in parts:
        rels = minidom.parseString(parts[relpath])
    else:
        rels = minidom.parseString(f'<Relationships xmlns="{R}"/>')
    if rels.documentElement.namespaceURI != R or rels.documentElement.localName != 'Relationships':
        raise ModelValidationError('Invalid workbook relationships document')
    matches = [n for n in _children(rels.documentElement, R, 'Relationship')
               if n.getAttribute('Type') == CONNECTION_REL]
    if len(matches) > 1:
        raise ModelValidationError('Ambiguous connections relationship')
    if matches:
        relation = matches[0]
        target = relation.getAttribute('Target')
        if relation.getAttribute('TargetMode') == 'External' or any(c in target for c in ('\\', '?', '#', '%', ':')):
            raise ModelValidationError('Connections relationship must target an internal workbook part')
        path = posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('xl', target))
        if not path.startswith('xl/'):
            raise ModelValidationError('Connections relationship leaves workbook parts')
        if path not in parts:
            raise ModelValidationError('Connections relationship targets a missing part')
    elif create:
        path = 'xl/connections.xml'
        if path in parts:
            raise ModelValidationError('Unreferenced connections part conflicts with new connection')
        ids = {n.getAttribute('Id') for n in _children(rels.documentElement, R, 'Relationship')}
        index = 1
        while f'rId{index}' in ids:
            index += 1
        rels.documentElement.appendChild(_node(rels, 'Relationship',
            {'Id': f'rId{index}', 'Type': CONNECTION_REL, 'Target': 'connections.xml'}, R))
    else:
        return None
    doc = minidom.parseString(parts[path]) if path in parts else minidom.parseString(f'<connections xmlns="{S}"/>')
    if doc.documentElement.namespaceURI != S or doc.documentElement.localName != 'connections':
        raise ModelValidationError('Invalid connections document')
    return path, doc, relpath, rels, bool(matches)


def _read(source):
    with ZipFile(source) as archive:
        infos = archive.infolist()
        if len({n.filename for n in infos}) != len(infos):
            raise ModelValidationError('Duplicate workbook ZIP entries')
        parts = {n.filename: archive.read(n) for n in infos}
        comment = archive.comment
    if 'xl/workbook.xml' not in parts or '[Content_Types].xml' not in parts:
        raise ModelValidationError('Expected an XLSX workbook package')
    return parts, infos, comment


def inspect_external_connections(source):
    """Return OLAP connection identities without disclosing connection strings."""
    parts, _, _ = _read(source)
    found = _connection_document(parts)
    if found is None:
        return ()
    path, doc, *_ = found
    result = []
    for node in _children(doc.documentElement, S, 'connection'):
        db = _children(node, S, 'dbPr')
        if (node.getAttribute('type') != '5' or not _children(node, S, 'olapPr') or len(db) != 1
                or node.getAttribute('name').casefold() == 'thisworkbookdatamodel'):
            continue
        result.append(ExternalConnectionInfo(int(node.getAttribute('id')), node.getAttribute('name'),
                                            db[0].getAttribute('command'), path))
    return tuple(result)


def add_external_connection(source, destination, connection):
    """Atomically append an external OLAP connection; no server or Excel needed.

    Existing connections are not replaced. This does not create a PivotTable or
    import remote tables into the embedded model. CUBE formulas can reference
    the authored connection by name when Excel connects to the real endpoint.
    """
    if not isinstance(connection, ExternalModelConnection):
        raise TypeError('connection must be ExternalModelConnection')
    parts, infos, comment = _read(source)
    if any(n.startswith('_xmlsignatures/') for n in parts):
        raise ModelValidationError('Cannot modify a digitally signed workbook package')
    path, doc, relpath, rels, existing = _connection_document(parts, create=True)
    current = _children(doc.documentElement, S, 'connection')
    if any(n.getAttribute('name').casefold() == connection.name.casefold() for n in current):
        raise ModelValidationError('Connection name already exists')
    try:
        ids = [int(n.getAttribute('id')) for n in current]
        if len(ids) != len(set(ids)) or any(n < 1 for n in ids):
            raise ValueError()
    except ValueError as exc:
        raise ModelValidationError('Existing connection IDs are invalid or ambiguous') from exc
    next_id = max(ids, default=0) + 1
    if next_id > 2**32 - 1:
        raise ModelValidationError('Workbook connection ID range exhausted')
    node = _node(doc, 'connection', {'id': next_id, 'name': connection.name,
        'description': connection.description, 'type': 5, 'refreshedVersion': 6,
        'minRefreshableVersion': 3, 'background': 0, 'refreshOnLoad': 0,
        'savePassword': 0, 'keepAlive': 1,
        'credentials': 'integrated' if connection.authentication == 'integrated' else 'prompt'})
    node.appendChild(_node(doc, 'dbPr', {'connection': connection.connection_string(),
                                       'command': connection.cube, 'commandType': 1}))
    node.appendChild(_node(doc, 'olapPr', {'sendLocale': 1, 'rowDrillCount': 1000,
        'serverFill': 1, 'serverFont': 1, 'serverFontColor': 1, 'serverNumberFormat': 1}))
    trailing = next((n for n in doc.documentElement.childNodes if n.nodeType == n.ELEMENT_NODE
                     and not (n.namespaceURI == S and n.localName == 'connection')), None)
    doc.documentElement.insertBefore(node, trailing)
    updates = {path: doc.toxml(encoding='utf-8')}
    if not existing:
        updates[relpath] = rels.toxml(encoding='utf-8')
    types = minidom.parseString(parts['[Content_Types].xml'])
    if types.documentElement.namespaceURI != C or types.documentElement.localName != 'Types':
        raise ModelValidationError('Invalid package content types document')
    overrides = [n for n in _children(types.documentElement, C, 'Override') if n.getAttribute('PartName') == '/' + path]
    if len(overrides) > 1 or (overrides and overrides[0].getAttribute('ContentType') != CONNECTION_MIME):
        raise ModelValidationError('Conflicting connections content type')
    if not overrides:
        types.documentElement.appendChild(_node(types, 'Override',
            {'PartName': '/' + path, 'ContentType': CONNECTION_MIME}, C))
        updates['[Content_Types].xml'] = types.toxml(encoding='utf-8')
    destination = Path(destination)
    fd, temporary = mkstemp(prefix='.external-connection-', suffix='.xlsx', dir=destination.parent)
    os.close(fd)
    try:
        with ZipFile(temporary, 'w', compression=ZIP_DEFLATED) as archive:
            archive.comment = comment
            for info in infos:
                archive.writestr(info, updates.pop(info.filename, parts[info.filename]))
            for name, data in updates.items():
                archive.writestr(name, data)
        os.replace(temporary, destination)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
    return destination


class ExternalQueryError(RuntimeError):
    """External provider query failure; provider text may contain private data."""


@dataclass(frozen=True)
class ExternalQueryResult:
    columns: tuple[str, ...]
    rows: tuple[tuple, ...]

    def to_dataframe(self):
        """Preserve exact provider values and nulls; callers choose conversions."""
        import pandas as pd
        # Inferred float columns round nullable Int64 values above 2**53.
        return pd.DataFrame(self.rows, columns=self.columns, dtype=object)


class AdomdQueryClient:
    """Optional ADOMD.NET adapter for an already deployed semantic model.

    Factory injection supports a caller-configured AdomdConnection (for example
    with an in-memory AccessToken); the adapter owns and disposes that connection.
    No tokens or passwords are serialized into workbook connection definitions.
    """
    def __init__(self, connection, *, connection_factory=None, assembly_path=None):
        if not isinstance(connection, ExternalModelConnection):
            raise TypeError('connection must be ExternalModelConnection')
        self.connection = connection
        self._factory = connection_factory
        self._assembly_path = assembly_path

    def _create(self):
        if self._factory is not None:
            return self._factory()
        if self.connection.authentication != 'integrated':
            raise ExternalQueryError('Interactive authentication is not supported by the headless query adapter; supply a caller-configured connection_factory')
        try:
            import clr
            clr.AddReference(str(self._assembly_path) if self._assembly_path else 'Microsoft.AnalysisServices.AdomdClient')
            from Microsoft.AnalysisServices.AdomdClient import AdomdConnection
        except Exception as exc:
            raise ExternalQueryError('ADOMD.NET and pythonnet are required; install the optional client and provide its assembly path') from exc
        connection = AdomdConnection()
        try:
            connection.ConnectionString = self.connection.connection_string() + 'Mode=Read;'
        except Exception:
            # query() cannot own this object until _create() returns it.
            try:
                connection.Dispose()
            except Exception:
                pass
            raise
        return connection

    def query(self, dax, *, timeout=30, max_rows=100000):
        """Execute one DAX result set with bounded materialization; never truncate."""
        if not isinstance(dax, str) or not re.match(r'^\s*(?:EVALUATE|DEFINE)\b', dax, re.I):
            raise ModelValidationError('Expected a DAX query beginning with EVALUATE or DEFINE')
        for name, value in [('timeout', timeout), ('max_rows', max_rows)]:
            if isinstance(value, bool) or not isinstance(value, int) or value < 1:
                raise ModelValidationError(f'{name} must be a positive integer')
        connection = command = reader = None
        try:
            connection = self._create()
            connection.Open()
            command = connection.CreateCommand()
            command.CommandText = dax
            command.CommandTimeout = timeout
            reader = command.ExecuteReader()
            columns = tuple(str(reader.GetName(i)) for i in range(reader.FieldCount))
            rows = []
            while reader.Read():
                if len(rows) >= max_rows:
                    raise ExternalQueryError('Query exceeded max_rows; add a query limit or explicitly increase max_rows')
                rows.append(tuple(None if reader.IsDBNull(i) else reader.GetValue(i) for i in range(reader.FieldCount)))
            if reader.NextResult():
                raise ExternalQueryError('Multiple result sets are unsupported; execute one EVALUATE query at a time')
            return ExternalQueryResult(columns, tuple(rows))
        except (ModelValidationError, ExternalQueryError):
            raise
        except Exception as exc:
            # Preserve diagnostic cause without echoing potentially credentialed
            # provider messages into the public exception text.
            raise ExternalQueryError('External semantic-model query failed; inspect the exception cause for provider diagnostics') from exc
        finally:
            for resource in (reader, command, connection):
                if resource is not None:
                    try:
                        resource.Dispose()
                    except Exception:
                        pass
