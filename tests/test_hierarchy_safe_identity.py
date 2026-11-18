"""Hierarchy captions stay public while physical bindings use legal identities."""
from xml.dom import minidom

import pandas as pd

from excel_model_lab import WorkbookModel
from excel_model_lab.abf import AbfImage
from excel_model_lab.hierarchy_storage import resolve_hierarchy, stored_hierarchies
from excel_model_lab.schema import _child, _elements, _text


def _hierarchy(model, caption):
    identity, _ = resolve_hierarchy(model, 'T', caption)
    doc = minidom.parseString(model._image.read_logical_file(model._dimension_paths['T']))
    return next(h for h in _elements(doc, 'Hierarchy') if _text(_child(h, 'ID')) == identity)


def _files(model, identity):
    stem = f'U${model._dimension_ids["T"]}${identity}.'
    return {f.logical_path for f in model._image.files if f.logical_path and stem in f.logical_path}


def _assert_bindings(model, caption):
    identity, actual_caption = resolve_hierarchy(model, 'T', caption)
    assert actual_caption == caption
    assert identity.startswith('_Hierarchy_')
    assert all(c.isalnum() or c == '_' for c in identity)
    raw = next(f for f in model._image.files if f.logical_path and
               f.logical_path.split('\\')[-1].startswith(model._dimension_ids['T'] + '.') and
               f.logical_path.endswith('.tbl.xml'))
    doc = minidom.parseString(model._image.read_logical_file(raw))
    assert any(n.getAttribute('name') == identity for n in _elements(doc, 'XMObject')
               if n.getAttribute('class') == 'XMUserHierarchy')
    cube = next(f for f in model._image.files if f.logical_path and f.logical_path.endswith('.cub.xml'))
    doc = minidom.parseString(model._image.read_logical_file(cube))
    assert identity in [_text(n) for n in _elements(doc, 'HierarchyID')]
    return identity


def test_dotted_hierarchy_caption_has_legal_consistent_identity_and_removal(tmp_path):
    model = WorkbookModel.create({'T': pd.DataFrame({'Label': ['a', 'b'], 'Key': [1, 2]})})
    model.add_hierarchy('T', 'A', ['Label'])
    model.add_hierarchy('T', 'A.B', ['Label', 'Key'])
    identity = _assert_bindings(model, 'A.B')
    files = _files(model, identity)
    assert len(files) == 5
    assert resolve_hierarchy(model, 'T', 'A')[0] == 'A'
    model = WorkbookModel(model.save(tmp_path / 'before.xlsx'))
    model.remove_hierarchy('T', 'A')
    assert stored_hierarchies(model) == [('T', 'A.B', ['Label', 'Key'])]
    assert _files(model, identity) == files
    model.remove_hierarchy('T', 'A.B')
    assert stored_hierarchies(model) == []
    assert not _files(model, identity)


def test_level_ids_use_native_column_identity_not_punctuated_caption():
    model = WorkbookModel.create({'T': pd.DataFrame({'Source': ['a', 'b'], 'Key': [1, 2]})})
    # Simulate a native caption rename without changing any physical streams.
    file = next(f for f in model._image.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
    doc = minidom.parseString(model._image.read_logical_file(file))
    attribute = next(a for a in _elements(doc, 'Attribute') if _text(_child(a, 'ID')) == 'Source')
    _child(attribute, 'Name').firstChild.data = 'Source.Caption'
    image = AbfImage(model._image.assemble({file.storage_path: doc.documentElement.toxml(encoding='utf-8')}))
    model = WorkbookModel(model.source, _image=image)
    model.add_hierarchy('T', 'A.B', ['Source.Caption', 'Key'])
    levels = _elements(_hierarchy(model, 'A.B'), 'Level')
    assert [_text(_child(n, 'Name')) for n in levels] == ['Source.Caption', 'Key']
    assert [_text(_child(n, 'ID')) for n in levels] == ['Source', 'Key']
    assert [_text(_child(n, 'SourceAttributeID')) for n in levels] == ['Source', 'Key']
    assert stored_hierarchies(model) == [('T', 'A.B', ['Source.Caption', 'Key'])]


def test_rebuild_preserves_generated_identity_and_configured_captions(tmp_path):
    model = WorkbookModel.create({'T': pd.DataFrame({'Label': ['a', 'b'], 'Key': [1, 2]})})
    model.add_hierarchy('T', 'A.B', ['Label', 'Key'])
    identity = _assert_bindings(model, 'A.B')
    model.configure_hierarchy('T', 'A.B', description='Keep description', hidden=True,
                              level_names={'Key': 'Item.Number'})
    model.replace_table('T', pd.DataFrame({'Label': ['a', 'a', 'b'], 'Key': [10, 20, 30]}))
    model = WorkbookModel(model.save(tmp_path / 'rebuilt.xlsx'))
    assert _assert_bindings(model, 'A.B') == identity
    hierarchy = _hierarchy(model, 'A.B')
    assert _text(_child(hierarchy, 'Description')) == 'Keep description'
    assert [_text(_child(n, 'Name')) for n in _elements(hierarchy, 'Level')] == ['Label', 'Item.Number']
    assert len(_files(model, identity)) == 5
    with model._open_reader() as reader:
        assert reader.get_table('T')['Key'].tolist() == [10, 20, 30]


def test_unprocessed_calculated_hierarchy_uses_safe_identity_without_indexes():
    model = WorkbookModel.create({'T': pd.DataFrame({'Key': [1, 2]})})
    model.add_calculated_column('T', 'Calculated', '[Key] + 1', data_type='integer')
    model.add_hierarchy('T', 'A.B', ['Calculated'])
    identity = _assert_bindings(model, 'A.B')
    assert not _files(model, identity)
    hierarchy = _hierarchy(model, 'A.B')
    assert _text(_child(hierarchy, 'ProcessingState')) == 'Unprocessed'
    assert _text(_child(_elements(hierarchy, 'Level')[0], 'ID')) == model._column_ids['T', 'Calculated']
