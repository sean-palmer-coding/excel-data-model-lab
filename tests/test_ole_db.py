import base64
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile
import pandas as pd
import pytest
from excel_model_lab import WorkbookModel
from excel_model_lab.native_sources import OleDbSource
from excel_model_lab.source_crypto import session_key, encrypt_string

pytest.importorskip('cryptography')
N = '{http://schemas.microsoft.com/analysisservices/2003/engine}'


def source(query='SELECT 1 AS ID, 10 AS Amount UNION ALL SELECT 2, 20'):
    return OleDbSource('Provider=MSOLEDBSQL;Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=master;Integrated Security=SSPI;', query)


def test_rebind_retains_source_extensions_but_replaces_connection_and_authentication(tmp_path):
    from excel_model_lab.abf import AbfImage
    from excel_model_lab.cube_source_view import _object
    from excel_model_lab.schema import _child, _elements, _text
    from excel_model_lab.ole_db import _append
    from cryptography.hazmat.decrepit.ciphers.algorithms import TripleDES
    from cryptography.hazmat.primitives.ciphers import Cipher, modes
    from cryptography.hazmat.primitives.padding import PKCS7

    def clear(image, value):
        key = next(f for f in image.files if f.logical_path and f.logical_path.endswith('CryptKey.bin'))
        value = value.removeprefix('[EncryptedDataStart]').removesuffix('[EncryptedDataEnd]')
        payload = base64.b64decode(value)
        decryptor = Cipher(TripleDES(session_key(image.read_logical_file(key))), modes.CBC(b'\0' * 8)).decryptor()
        raw = decryptor.update(payload[8:]) + decryptor.finalize()
        unpadder = PKCS7(64).unpadder()
        return (unpadder.update(raw) + unpadder.finalize())[8:].decode('utf-16le').rstrip('\0')

    m = WorkbookModel.create({'Sales': pd.DataFrame({'ID': [1], 'Amount': [10]})})
    m.bind_source('Sales', source())
    file, doc, ds = _object(m._image, '.ds.xml', 'DataSource')
    _child(ds, 'Description').appendChild(doc.createTextNode('preserved description'))
    _child(ds, 'Timeout').firstChild.data = 'PT17S'
    ds.setAttribute('xmlns:custom', 'urn:test-extension')
    ds.setAttributeNS('urn:test-extension', 'custom:flag', 'preserved')
    extension = doc.createElementNS('urn:test-extension', 'custom:Extension')
    extension.appendChild(doc.createTextNode('payload'))
    ds.appendChild(extension)
    annotation = _append(_child(ds, 'Annotations'), 'Annotation')
    _append(annotation, 'Name', 'ExternalAnnotation')
    _append(annotation, 'Value', 'preserved annotation')
    impersonation = _child(ds, 'ImpersonationInfo')
    _child(impersonation, 'Account').appendChild(doc.createTextNode('obsolete-account'))
    _child(impersonation, 'Password').firstChild.data = 'obsolete-password-envelope'
    _append(ds, 'ConnectionStringSecurity', 'PasswordRemoved')
    m._image = AbfImage(m._image.assemble({file.storage_path: doc.documentElement.toxml(encoding='utf-8')}))
    connection = 'Provider=SQLOLEDB;Data Source=changed-synthetic;Integrated Security=SSPI;'
    m.bind_source('Sales', OleDbSource(connection, 'SELECT 2 AS ID, 20 AS Amount', name='Changed source'))
    reopened = WorkbookModel(m.save(tmp_path / 'source-extensions.xlsx'))
    _, _, ds = _object(reopened._image, '.ds.xml', 'DataSource')
    assert _text(_child(ds, 'Name')) == 'Changed source'
    assert clear(reopened._image, _text(_child(ds, 'ConnectionString'))) == connection
    assert _text(_child(ds, 'Description')) == 'preserved description'
    assert _text(_child(ds, 'Timeout')) == 'PT17S'
    assert ds.getAttributeNS('urn:test-extension', 'flag') == 'preserved'
    assert _text(ds.getElementsByTagNameNS('urn:test-extension', 'Extension')[0]) == 'payload'
    annotations = {_text(_child(n, 'Name')): _text(_child(n, 'Value')) for n in _elements(ds, 'Annotation')}
    assert annotations == {'ConnectionEditUISource': 'SqlServer', 'ExternalAnnotation': 'preserved annotation'}
    impersonation = _child(ds, 'ImpersonationInfo')
    assert _text(_child(impersonation, 'Account')) == ''
    assert clear(reopened._image, _text(_child(impersonation, 'Password'))) == ''
    assert not _elements(ds, 'ConnectionStringSecurity')


def test_rebind_rejects_ambiguous_source_fields_atomically():
    from excel_model_lab.abf import AbfImage, UnsupportedModelError
    from excel_model_lab.cube_source_view import _object
    from excel_model_lab.schema import _child
    m = WorkbookModel.create({'Sales': pd.DataFrame({'ID': [1], 'Amount': [10]})})
    m.bind_source('Sales', source())
    file, doc, ds = _object(m._image, '.ds.xml', 'DataSource')
    ds.appendChild(_child(ds, 'ConnectionString').cloneNode(True))
    m._image = AbfImage(m._image.assemble({file.storage_path: doc.documentElement.toxml(encoding='utf-8')}))
    before = m._image.data
    with pytest.raises(UnsupportedModelError, match='duplicate data source field'):
        m.bind_source('Sales', source())
    assert m._image.data == before


def test_native_oledb_binding_preserves_cached_data_and_is_rebindable(tmp_path):
    m = WorkbookModel.create({'Sales': pd.DataFrame({'ID': [1, 2], 'Amount': [10, 20]})})
    m.bind_source('Sales', source())
    path = m.save(tmp_path / 'oledb.xlsx')
    reopened = WorkbookModel(path)
    assert reopened.get_table('Sales').to_dict('list') == {'ID': [1, 2], 'Amount': [10, 20]}
    assert reopened.list_native_sources()[0].kind == 'ole_db'
    groups_before = len(reopened._image.backup_log_root.findall('FileGroups/FileGroup'))
    reopened.bind_source('Sales', source('SELECT 3 AS ID, 30 AS Amount'))
    assert len(reopened._image.backup_log_root.findall('FileGroups/FileGroup')) == groups_before
    output = reopened.save(tmp_path / 'rebound.xlsx')
    with ZipFile(output) as archive:
        connections = ET.fromstring(archive.read('xl/connections.xml'))
        bindings = [c for c in connections if c.get('type') == '100']
        assert len(bindings) == 1 and bindings[0].get('refreshOnLoad') == '0'
    reopened.unbind_source('Sales')
    for file in reopened._image.files:
        if file.logical_path and file.logical_path.endswith('.prt.xml'):
            root = ET.fromstring(reopened._image.read_logical_file(file))
            assert not root.findall('.//' + N + 'QueryDefinition')


def test_two_native_sources_have_independent_groups(tmp_path):
    m = WorkbookModel.create({'First': pd.DataFrame({'ID': [1], 'Amount': [10]}),
                              'Second': pd.DataFrame({'ID': [2], 'Amount': [20]})})
    m.bind_source('First', source()).bind_source('Second', source())
    path = m.save(tmp_path / 'two-sources.xlsx')
    assert len(WorkbookModel(path).list_native_sources()) == 2


def test_crypto_matches_native_empty_password_envelope():
    from cryptography.hazmat.decrepit.ciphers.algorithms import TripleDES
    from cryptography.hazmat.primitives.ciphers import Cipher, modes
    from cryptography.hazmat.primitives.padding import PKCS7
    from excel_model_lab.abf import AbfImage
    seed = Path(__file__).parents[1] / 'src/excel_model_lab/data/authoring-seed.xlsx'
    image = AbfImage.from_workbook(seed)
    key_file = next(f for f in image.files if f.logical_path and f.logical_path.endswith('CryptKey.bin'))
    key_data = image.read_logical_file(key_file)
    database = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.db.xml'))
    password = ET.fromstring(image.read_logical_file(database)).findtext('.//' + N + 'Password')
    for encoded in (password, encrypt_string(key_data, '', flags=0), encrypt_string(key_data, 'hello')):
        cipher = base64.b64decode(encoded)
        decryptor = Cipher(TripleDES(session_key(key_data)), modes.CBC(b'\0' * 8)).decryptor()
        padded = decryptor.update(cipher[8:]) + decryptor.finalize()
        unpadder = PKCS7(64).unpadder()
        clear = unpadder.update(padded) + unpadder.finalize()
        assert clear[8:].decode('utf-16le') in ('\0', 'hello\0')


def test_cube_source_view_tracks_multiple_sources_rebind_and_detach(tmp_path):
    m = WorkbookModel.create({'First': pd.DataFrame({'ID': [1], 'Amount': [10]}),
                              'Second': pd.DataFrame({'ID': [2], 'Amount': [20]})})
    m.bind_source('First', source()).bind_source('Second', source('SELECT 2 AS ID, 20 AS Amount'))

    def inspect(model):
        docs = {f.storage_path: ET.fromstring(model._image.read_logical_file(f))
                for f in model._image.files if f.logical_path and f.logical_path.endswith('.xml')}
        cube = next(r.find('.//' + N + 'Cube') for r in docs.values() if r.find('.//' + N + 'Cube') is not None)
        binding = cube.find(N + 'Source')
        views = {v.findtext(N + 'ID'): v for r in docs.values()
                 for v in r.findall('.//' + N + 'DataSourceView')}
        sources = {v.findtext(N + 'ID') for r in docs.values()
                   for v in r.findall('.//' + N + 'DataSource')}
        return docs, binding, views, sources

    docs, binding, views, sources = inspect(m)
    assert binding is not None
    shared_id = binding.findtext(N + 'DataSourceViewID')
    shared = views[shared_id]
    assert len(views) == 1
    assert all(node.text == shared_id for doc in docs.values()
               for node in doc.findall('.//' + N + 'DataSourceViewID'))
    XS = '{http://www.w3.org/2001/XMLSchema}'
    tables = shared.findall('.//' + XS + 'choice/' + XS + 'element')
    assert {t.get('name') for t in tables} == set(m._dimension_ids.values())
    source_ids = {r.findtext('.//' + N + 'Partition/' + N + 'Source/' + N + 'DataSourceID')
                  for r in docs.values()} - {None}
    assert len(source_ids) == 2 and source_ids <= sources
    assert {t.get('{urn:schemas-microsoft-com:xml-msprop}DataSourceID') for t in tables} == source_ids
    m.bind_source('First', source('SELECT 3 AS ID, 30 AS Amount'))
    _, binding, views, _ = inspect(m)
    assert binding.findtext(N + 'DataSourceViewID') == shared_id
    assert len(views[shared_id].findall('.//' + XS + 'choice/' + XS + 'element')) == 2
    m.unbind_source('First')
    _, binding, views, sources = inspect(m)
    shared = views[binding.findtext(N + 'DataSourceViewID')]
    assert len(sources) == 1
    assert shared.findtext(N + 'DataSourceID') in sources
    assert {t.get('name') for t in shared.findall('.//' + XS + 'choice/' + XS + 'element')} == {m._dimension_ids['Second']}
    p = m.save(tmp_path / 'remaining-source.xlsx')
    reopened = WorkbookModel(p)
    assert reopened.get_table('First').ID.tolist() == [1]
    assert reopened.get_table('Second').ID.tolist() == [2]
    reopened.unbind_source('Second')
    _, binding, views, sources = inspect(reopened)
    assert binding is None
    assert not views and not sources


def test_rebind_and_detach_native_versioned_source_file(tmp_path):
    from dataclasses import replace
    from excel_model_lab.abf import AbfImage
    m = WorkbookModel.create({'Sales': pd.DataFrame({'ID': [1], 'Amount': [10]})})
    m.bind_source('Sales', source())
    image = m._image
    file = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.ds.xml'))
    old = file.logical_path.rsplit('\\', 1)[-1]
    new = old.replace('.0.ds.xml', '.1.ds.xml')
    for node in image.backup_log_root.findall('FileGroups/FileGroup/FileList/BackupFile'):
        path = node.find('Path')
        if path is not None and path.text:
            path.text = path.text.replace(old, new)
    image.files = [replace(f, logical_path=f.logical_path.replace(old, new))
                   if f.logical_path else f for f in image.files]
    replacements = {f.storage_path: image.read_logical_file(f).replace(old.encode(), new.encode())
                    for f in image.files if f.logical_path and f.logical_path.endswith('.xml')}
    m._image = AbfImage(image.assemble(replacements))
    m.bind_source('Sales', source('SELECT 2 AS ID, 20 AS Amount'))
    files = [f for f in m._image.files if f.logical_path and f.logical_path.endswith('.ds.xml')]
    assert len(files) == 1 and files[0].logical_path.endswith('.1.ds.xml')
    m.unbind_source('Sales')
    assert not any(f.logical_path and f.logical_path.endswith(('.ds.xml', '.dsv.xml')) for f in m._image.files)
