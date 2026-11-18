import pandas as pd
from excel_model_lab import WorkbookModel
from excel_model_lab.native_sources import OleDbSource, binding_overrides
from excel_model_lab.cube_source_view import bind_cube_view, detach_cube_view, _object, _tables, _metadata
from excel_model_lab.schema import _child, _text


def test_rebind_preserves_table_and_column_schema_extensions(tmp_path):
    from excel_model_lab.abf import AbfImage
    m = WorkbookModel.create({'T': pd.DataFrame({'ID': [1]})})
    connection = 'Provider=SQLOLEDB;Data Source=synthetic;Integrated Security=SSPI;'
    m.bind_source('T', OleDbSource(connection, 'SELECT 1 AS ID'))
    cfg = binding_overrides(m._image)[m._dimension_ids['T']]
    file, doc, view = _object(m._image, '.dsv.xml', 'DataSourceView', cfg['view_id'])
    table = next(n for n in _tables(view).childNodes if getattr(n, 'localName', None) == 'element')
    table.setAttribute('xmlns:custom', 'urn:test-extension')
    column = table.getElementsByTagNameNS('http://www.w3.org/2001/XMLSchema', 'element')[0]
    for node in (table, column):
        node.setAttributeNS('urn:test-extension', 'custom:flag', 'preserved')
        extension = doc.createElementNS('urn:test-extension', 'custom:Extension')
        extension.appendChild(doc.createTextNode('extension payload'))
        node.appendChild(extension)
    m._image = AbfImage(m._image.assemble({file.storage_path: doc.documentElement.toxml(encoding='utf-8')}))
    m.bind_source('T', OleDbSource(connection, 'SELECT 2 AS ID'))
    reopened = WorkbookModel(m.save(tmp_path / 'extended.xlsx'))
    _, _, view = _object(reopened._image, '.dsv.xml', 'DataSourceView', cfg['view_id'])
    table = next(n for n in _tables(view).childNodes if getattr(n, 'localName', None) == 'element')
    column = table.getElementsByTagNameNS('http://www.w3.org/2001/XMLSchema', 'element')[0]
    assert table.getAttribute('msprop:QueryDefinition') == 'SELECT 2 AS ID'
    for node in (table, column):
        assert node.getAttributeNS('urn:test-extension', 'flag') == 'preserved'
        extension = next(n for n in node.childNodes if getattr(n, 'localName', None) == 'Extension')
        assert _text(extension) == 'extension payload'


def test_rebind_rejects_competing_schema_without_mutating_model():
    import pytest
    from excel_model_lab.abf import AbfImage, UnsupportedModelError
    m = WorkbookModel.create({'T': pd.DataFrame({'ID': [1]})})
    source = OleDbSource('Provider=SQLOLEDB;Data Source=synthetic;Integrated Security=SSPI;', 'SELECT 1 AS ID')
    m.bind_source('T', source)
    cfg = binding_overrides(m._image)[m._dimension_ids['T']]
    file, doc, view = _object(m._image, '.dsv.xml', 'DataSourceView', cfg['view_id'])
    sequence = view.getElementsByTagNameNS('http://www.w3.org/2001/XMLSchema', 'sequence')[0]
    extra = doc.createElementNS('http://www.w3.org/2001/XMLSchema', 'xs:element')
    extra.setAttribute('name', 'ExternalColumn')
    sequence.appendChild(extra)
    m._image = AbfImage(m._image.assemble({file.storage_path: doc.documentElement.toxml(encoding='utf-8')}))
    before = m._image.data
    with pytest.raises(UnsupportedModelError, match='schema reconciliation'):
        m.bind_source('T', source)
    assert m._image.data == before

def test_shared_cube_view_merge_rebind_detach_preserves_partition_sources():
    m=WorkbookModel.create({'One':pd.DataFrame({'ID':[1]}),'Two':pd.DataFrame({'ID':[2]})})
    for table in ('One','Two'):
        m.bind_source(table,OleDbSource('Provider=SQLOLEDB;Data Source=synthetic;Integrated Security=SSPI;', 'SELECT 1 AS ID',name=table+' source'))
    image=m._image
    configs=binding_overrides(image)
    partitions={f.storage_path:image.read_logical_file(f) for f in image.files if f.logical_path and f.logical_path.endswith('.prt.xml')}
    for table in ('One','Two'):
        identity=m._dimension_ids[table];cfg=configs[identity]
        image=bind_cube_view(image,identity,cfg['view_id'],cfg['data_source_id'])
    _,_,cube=_object(image,'.cub.xml','Cube')
    view_id=_text(_child(_child(cube,'Source'),'DataSourceViewID'))
    _,_,view=_object(image,'.dsv.xml','DataSourceView',view_id)
    assert set(_metadata(view)['tables'])==set(m._dimension_ids.values())
    assert len([n for n in _tables(view).childNodes if getattr(n,'localName',None)=='element'])==2
    for path,data in partitions.items():assert image.read_logical_file(path)==data
    one=m._dimension_ids['One'];two=m._dimension_ids['Two']
    image=detach_cube_view(image,one)
    _,_,view=_object(image,'.dsv.xml','DataSourceView',view_id)
    assert _text(_child(view,'DataSourceID'))==configs[two]['data_source_id']
    assert set(_metadata(view)['tables'])=={two}
    image=detach_cube_view(image,two)
    _,_,cube=_object(image,'.cub.xml','Cube')
    assert not [n for n in cube.childNodes if getattr(n,'localName',None)=='Source']
    assert not any(f.logical_path and view_id in f.logical_path for f in image.files)

def test_existing_cube_view_extension_and_unowned_table_retained():
    from excel_model_lab.abf import AbfImage
    from excel_model_lab.schema import _elements
    m=WorkbookModel.create({'One':pd.DataFrame({'ID':[1]})})
    m.bind_source('One',OleDbSource('Provider=SQLOLEDB;Data Source=synthetic;Integrated Security=SSPI;', 'SELECT 1 AS ID'))
    image=m._image;table=m._dimension_ids['One'];cfg=binding_overrides(image)[table]
    _,_,cube=_object(image,'.cub.xml','Cube')
    identity=_text(_child(_child(cube,'Source'),'DataSourceViewID'))
    file,doc,view=_object(image,'.dsv.xml','DataSourceView',identity)
    for a in list(_elements(view,'Annotation')):
        if _text(_child(a,'Name'))=='ExcelModelLab.CubeViewTables':a.parentNode.removeChild(a)
    view.setAttribute('unknown-extension','retained')
    image=AbfImage(image.assemble({file.storage_path:doc.documentElement.toxml(encoding='utf-8')}))
    image=bind_cube_view(image,table,cfg['view_id'],cfg['data_source_id'])
    image=detach_cube_view(image,table)
    _,_,view=_object(image,'.dsv.xml','DataSourceView',identity)
    assert view.getAttribute('unknown-extension')=='retained'
    assert len([n for n in _tables(view).childNodes if getattr(n,'localName',None)=='element'])==1


def test_detach_retained_cube_view_keeps_its_referenced_source():
    from excel_model_lab.abf import AbfImage
    from excel_model_lab.schema import _elements
    m=WorkbookModel.create({'One':pd.DataFrame({'ID':[1]})})
    m.bind_source('One',OleDbSource('Provider=SQLOLEDB;Data Source=synthetic;Integrated Security=SSPI;', 'SELECT 1 AS ID'))
    table=m._dimension_ids['One'];cfg=binding_overrides(m._image)[table]
    file,doc,view=_object(m._image,'.dsv.xml','DataSourceView',cfg['view_id'])
    for a in list(_elements(view,'Annotation')):
        if _text(_child(a,'Name'))=='ExcelModelLab.CubeViewTables':a.parentNode.removeChild(a)
    view.setAttribute('unknown-extension','retained')
    image=AbfImage(m._image.assemble({file.storage_path:doc.documentElement.toxml(encoding='utf-8')}))
    image=bind_cube_view(image,table,cfg['view_id'],cfg['data_source_id'])
    from excel_model_lab.ole_db import detach_table_binding
    image=detach_table_binding(image,table)
    _,_,retained=_object(image,'.dsv.xml','DataSourceView',cfg['view_id'])
    assert retained.getAttribute('unknown-extension')=='retained'
    _object(image,'.ds.xml','DataSource',cfg['data_source_id'])
