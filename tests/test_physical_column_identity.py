import struct
from xml.dom import minidom
import pandas as pd
from excel_model_lab import WorkbookModel
from excel_model_lab.abf import AbfImage
from excel_model_lab.schema import _elements, _child, _text
from excel_model_lab.relationship_storage import stored_relationships
from excel_model_lab.hierarchy_storage import stored_hierarchies, intrinsic_order

def captions(model, tmp_path):
    model = WorkbookModel(model.save(tmp_path / "base.xlsx"))
    changes = {}
    for table, path in model._dimension_paths.items():
        doc = minidom.parseString(model._image.read_logical_file(path))
        for attr in _elements(doc, 'Attribute'):
            ids = [n for n in attr.childNodes if getattr(n, 'localName', None) == 'ID']
            if ids and (table, _text(ids[0])) in model._column_ids:
                _child(attr, 'Name').firstChild.data = 'Visible ' + _text(ids[0])
        changes[path] = doc.documentElement.toxml(encoding='utf-8')
    return WorkbookModel(model.source, _image=AbfImage(model._image.assemble(changes)))

def test_relationship_identity_save_rebuild_remove(tmp_path):
    model = captions(WorkbookModel.create({
        'Products': pd.DataFrame({'ID':[1,2,3]}),
        'Sales': pd.DataFrame({'ID':[1,2,1,3]})}), tmp_path)
    model.add_relationship('Sales','Visible ID','Products','Visible ID')
    model = WorkbookModel(model.save(tmp_path/'related.xlsx'))
    assert stored_relationships(model)[0][1].from_column == 'Visible ID'
    model.replace_column('Products','Visible ID',[3,2,1])
    model = WorkbookModel(model.save(tmp_path/'reordered.xlsx'))
    index = next(f for f in model._image.files if f.logical_path and '.INDEX.' in f.logical_path)
    _, word = struct.unpack('<QQ',model._image.read_logical_file(index))
    assert [(word >> (2*i)) & 3 for i in range(6)] == [0,0,0,3,2,1]
    model.remove_relationship('Sales','Visible ID','Products','Visible ID')
    model = WorkbookModel(model.save(tmp_path/'removed.xlsx'))
    assert stored_relationships(model) == []

def test_sort_and_hierarchy_identity_rebuild(tmp_path):
    model = captions(WorkbookModel.create({'T':pd.DataFrame({'Label':['Alpha','Beta','Gamma'],'Order':[3,2,1]})}), tmp_path)
    model.set_sort_column('T','Visible Label','Visible Order')
    model.add_hierarchy('T','Groups',['Visible Label','Visible Order'])
    model = WorkbookModel(model.save(tmp_path/'sorted.xlsx'))
    assert stored_hierarchies(model)[0][2] == ['Visible Label','Visible Order']
    assert list(intrinsic_order(model,'T','Visible Label')) == [5,4,3]
    model.replace_column('T','Visible Order',[2,1,3])
    model = WorkbookModel(model.save(tmp_path/'updated.xlsx'))
    assert list(intrinsic_order(model,'T','Visible Label')) == [4,3,5]
    assert model.get_table('T')['Visible Order'].tolist() == [2,1,3]

def test_hierarchy_caption_identity_edit_and_dependencies(tmp_path):
    import pytest
    from excel_model_lab.dependencies import DependencyError
    model = captions(WorkbookModel.create({'T': pd.DataFrame({'ID':[1,2], 'Group':['A','B']})}), tmp_path)
    model.add_hierarchy('T', 'PhysicalHierarchy', ['Visible Group','Visible ID'])
    path = model._dimension_paths['T']
    doc = minidom.parseString(model._image.read_logical_file(path))
    hierarchy = next(h for h in _elements(doc,'Hierarchy') if _text(_child(h,'ID')) == 'PhysicalHierarchy')
    _child(hierarchy,'Name').firstChild.data = 'Visible Hierarchy'
    model = WorkbookModel(model.source, _image=AbfImage(model._image.assemble({path:doc.documentElement.toxml(encoding='utf-8')})))
    with pytest.raises(ValueError, match='already exists'):
        model.add_hierarchy('T','PhysicalHierarchy',['Visible ID'])
    model.configure_hierarchy('T','Visible Hierarchy',description='Retained description',hidden=True,
                              level_names={'Visible ID':'Number'})
    model.replace_column('T','Visible ID',[3,4])
    model = WorkbookModel(model.save(tmp_path/'hierarchy-configured.xlsx'))
    doc = minidom.parseString(model._image.read_logical_file(model._dimension_paths['T']))
    hierarchy = _elements(_elements(doc,'Hierarchies')[0],'Hierarchy')[0]
    assert _text(_child(hierarchy,'ID')) == 'PhysicalHierarchy'
    assert _text(_child(hierarchy,'Name')) == 'Visible Hierarchy'
    assert _text(_child(hierarchy,'Description')) == 'Retained description'
    model.create_perspective('Selected', hierarchies=[('T','Visible Hierarchy')])
    before = model._image.data
    with pytest.raises(DependencyError):
        model.remove_hierarchy('T','Visible Hierarchy')
    assert model._image.data == before
    model.remove_perspective('Selected')
    model.update_hierarchy('T','Visible Hierarchy',new_name='Updated Hierarchy',levels=['Visible ID','Visible Group'])
    model = WorkbookModel(model.save(tmp_path/'hierarchy-updated.xlsx'))
    doc = minidom.parseString(model._image.read_logical_file(model._dimension_paths['T']))
    hierarchy = _elements(_elements(doc,'Hierarchies')[0],'Hierarchy')[0]
    assert _text(_child(_elements(hierarchy,'Level')[0],'Name')) == 'Number'
    model.remove_hierarchy('T','Updated Hierarchy')
    model = WorkbookModel(model.save(tmp_path/'hierarchy-removed.xlsx'))
    assert stored_hierarchies(model) == []
