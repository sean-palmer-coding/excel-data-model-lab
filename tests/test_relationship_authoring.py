import struct
from xml.dom import minidom

import pandas as pd
import pytest

from excel_model_lab import ExcelModelTemplate
from excel_model_lab.schema import _elements, _child, _text


def related_model(tmp_path):
    model = ExcelModelTemplate.create({
        'Products': pd.DataFrame({'ID': [1, 2, 3], 'Category': ['Alpha', 'Beta', 'Gamma']}),
        'Sales': pd.DataFrame({'ID': [1, 2, 1, 3], 'Amount': [10, 20, 30, 40]}),
    })
    model.add_relationship('Sales', 'ID', 'Products', 'ID')
    return model.save(tmp_path / 'related.xlsx')


def test_relationship_key_contract_and_xml_order(tmp_path):
    model = ExcelModelTemplate(related_model(tmp_path))
    for table, path in model._dimension_paths.items():
        doc = minidom.parseString(model._image.read_logical_file(path))
        if table == 'Sales':
            children = [n.localName for n in _elements(doc, 'Dimension')[0].childNodes if n.nodeType == n.ELEMENT_NODE]
            assert children.index('Relationships') < children.index('StructureVersion')
        else:
            attribute = next(n for n in _elements(doc, 'Attribute')
                             if any(getattr(c, 'localName', None) == 'ID' and _text(c) == 'ID' for c in n.childNodes))
            assert _text(_elements(attribute, 'NullProcessing')[0]) == 'Error'


def test_lookup_index_rebuilt_after_primary_row_reorder(tmp_path):
    model = ExcelModelTemplate(related_model(tmp_path))
    model.replace_table('Products', pd.DataFrame({'ID': [3, 2, 1], 'Category': ['Gamma', 'Beta', 'Alpha']}))
    result = ExcelModelTemplate(model.save(tmp_path / 'reordered.xlsx'))
    index = next(f for f in result._image.files if f.logical_path and '.INDEX.' in f.logical_path)
    count, word = struct.unpack('<QQ', result._image.read_logical_file(index))
    assert count == 1
    assert [(word >> (2 * i)) & 3 for i in range(6)] == [0, 0, 0, 3, 2, 1]


def test_invalid_lookup_does_not_replace_destination(tmp_path):
    model = ExcelModelTemplate(related_model(tmp_path))
    model.replace_column('Products', 'ID', [1, 1, 3])
    destination = tmp_path / 'existing.xlsx'
    destination.write_bytes(b'unchanged')
    with pytest.raises(ValueError, match='unique'):
        model.save(destination)
    assert destination.read_bytes() == b'unchanged'


def test_remove_and_deactivate_relationship(tmp_path):
    from excel_model_lab.relationship_storage import stored_relationships
    source = related_model(tmp_path)
    model = ExcelModelTemplate(source)
    model.set_relationship_active('Sales', 'ID', 'Products', 'ID', active=False)
    inactive = ExcelModelTemplate(model.save(tmp_path / 'inactive.xlsx'))
    assert not stored_relationships(inactive)[0][1].active
    for file in inactive._image.files:
        if file.logical_path and file.logical_path.endswith('.det.xml'):
            assert b'ReferenceMeasureGroupDimension' not in inactive._image.read_logical_file(file)
    inactive.remove_relationship('Sales', 'ID', 'Products', 'ID')
    removed = ExcelModelTemplate(inactive.save(tmp_path / 'removed.xlsx'))
    assert stored_relationships(removed) == []
    assert not any(f.logical_path and '.INDEX.' in f.logical_path for f in removed._image.files)


def test_duplicate_and_ambiguous_relationship_rejected(tmp_path):
    model = ExcelModelTemplate(related_model(tmp_path))
    with pytest.raises(ValueError, match='already exists'):
        model.add_relationship('Sales', 'ID', 'Products', 'ID')
    with pytest.raises(ValueError, match='ambiguous'):
        model.add_relationship('Sales', 'Amount', 'Products', 'ID')


def test_unknown_hierarchy_member_propagates_and_clears(tmp_path):
    model = ExcelModelTemplate.create({
        'Regions': pd.DataFrame({'ID': [1, 2], 'Name': ['East', 'West']}),
        'Products': pd.DataFrame({'ID': [1, 2, 3], 'RegionID': [1, 2, 1]}),
        'Sales': pd.DataFrame({'ID': [1, 2, 99], 'Amount': [10, 20, 55]}),
    })
    model.add_relationship('Sales', 'ID', 'Products', 'ID')
    model.add_relationship('Products', 'RegionID', 'Regions', 'ID')
    model.add_hierarchy('Regions', 'Geography', ['Name', 'ID'])
    def save_counts(current, name):
        result = ExcelModelTemplate(current.save(tmp_path / name))
        captions = {v: k for k, v in result._dimension_ids.items()}
        counts = {}
        for f in result._image.files:
            if f.logical_path and f.logical_path.endswith('.tbl.xml'):
                doc = minidom.parseString(result._image.read_logical_file(f))
                table = captions.get(doc.documentElement.getAttribute('name'))
                if table:
                    counts[table] = int(_text(_child(_child(doc.documentElement, 'Properties'), 'RIViolationCount')))
        return result, counts
    model, counts = save_counts(model, 'unknown.xlsx')
    assert counts == {'Regions': 1, 'Products': 1, 'Sales': 0}
    def relationship_flags(current):
        flags = {}
        for file in current._image.files:
            if not file.logical_path or not file.logical_path.endswith('.tbl.xml'):
                continue
            doc = minidom.parseString(current._image.read_logical_file(file))
            for rel in _elements(doc, 'XMObject'):
                if rel.getAttribute('class') == 'XMRelationship':
                    foreign = _text(_child(_child(rel, 'Properties'), 'ForeignColumn'))
                    flags[foreign] = int(_text(_elements(rel, 'Flags')[0]))
        return flags
    assert relationship_flags(model) == {'ID': 1, 'RegionID': 0}
    model.set_relationship_active('Products', 'RegionID', 'Regions', 'ID', active=False)
    model, counts = save_counts(model, 'inactive-unknown.xlsx')
    assert counts['Regions'] == 1
    model.replace_column('Sales', 'ID', [1, 2, 3])
    model, counts = save_counts(model, 'matched.xlsx')
    assert not any(counts.values())
    assert not any(relationship_flags(model).values())
    model.replace_column('Sales', 'ID', [1, None, 3])
    model, counts = save_counts(model, 'null.xlsx')
    assert counts['Regions'] == counts['Products'] == 1
    assert relationship_flags(model)['ID'] == 1
    model.remove_relationship('Sales', 'ID', 'Products', 'ID')
    _, counts = save_counts(model, 'removed-unknown.xlsx')
    assert not any(counts.values())


@pytest.mark.parametrize('reverse_order', [False, True])
def test_chain_bindings_follow_graph_edits(tmp_path, reverse_order):
    model = ExcelModelTemplate.create({
        'Regions': pd.DataFrame({'ID': [1, 2]}),
        'Products': pd.DataFrame({'ID': [1, 2, 3], 'RegionID': [1, 2, 1]}),
        'Sales': pd.DataFrame({'ID': [1, 2, 3, 1], 'Amount': [10, 20, 30, 40]}),
    })
    edges = [('Sales', 'ID', 'Products', 'ID'), ('Products', 'RegionID', 'Regions', 'ID')]
    for edge in reversed(edges) if reverse_order else edges:
        model.add_relationship(*edge)
    model = ExcelModelTemplate(model.save(tmp_path / 'chain.xlsx'))
    def references(current):
        physical = current._dimension_ids['Sales']
        file = next(f for f in current._image.files if f.logical_path and
                    f.logical_path.split('\\')[-1].startswith(physical + '.') and f.logical_path.endswith('.det.xml'))
        doc = minidom.parseString(current._image.read_logical_file(file))
        return [n for n in _elements(doc, 'Dimension')
                if n.getAttribute('xsi:type') == 'ReferenceMeasureGroupDimension']
    refs = references(model)
    assert len(refs) == 2
    indirect = next(r for r in refs if _text(_child(r, 'Materialization')) == 'Indirect')
    assert _text(_child(indirect, 'CubeDimensionID')) == model._dimension_ids['Regions']
    assert _text(_child(indirect, 'IntermediateCubeDimensionID')) == model._dimension_ids['Products']
    assert not _elements(indirect, 'RelationshipID')
    model.set_relationship_active(*edges[1], active=False)
    model = ExcelModelTemplate(model.save(tmp_path / 'inactive-chain.xlsx'))
    assert len(references(model)) == 1
    model.set_relationship_active(*edges[1], active=True)
    model = ExcelModelTemplate(model.save(tmp_path / 'active-chain.xlsx'))
    assert len(references(model)) == 2
    model.remove_relationship(*edges[0])
    assert references(model) == []


def test_new_lookup_column_is_bound_into_fact_dimension(tmp_path):
    model = ExcelModelTemplate(related_model(tmp_path))
    model.add_column('Products', 'ID_', [10, 20, 30])
    result = ExcelModelTemplate(model.save(tmp_path / 'added.xlsx'))
    doc = minidom.parseString(result._image.read_logical_file(result._dimension_paths['Products']))
    names = [_text(_child(n, 'Name')).casefold() for n in _elements(doc, 'AttributeRelationship')]
    assert len(set(names)) == len(names)
    for file in result._image.files:
        if file.logical_path and file.logical_path.endswith('.det.xml'):
            doc = minidom.parseString(result._image.read_logical_file(file))
            for dimension in _elements(doc, 'Dimension'):
                if _text(_child(dimension, 'CubeDimensionID')) == result._dimension_ids['Products']:
                    assert 'ID_' in [_text(n) for n in _elements(dimension, 'AttributeID')]

def test_last_relationship_releases_primary_key_contract(tmp_path):
    model = ExcelModelTemplate(related_model(tmp_path))
    model.add_table('Other', pd.DataFrame({'ID':[1,2,3]}))
    model.add_relationship('Other','ID','Products','ID')
    model = ExcelModelTemplate(model.save(tmp_path/'two.xlsx'))
    model.remove_relationship('Sales','ID','Products','ID')
    assert not next(c for c in model.columns if (c.table,c.name)==('Products','ID')).nullable
    model.remove_relationship('Other','ID','Products','ID')
    assert next(c for c in model.columns if (c.table,c.name)==('Products','ID')).nullable
    model.replace_column('Products','ID',[1,None,1])
    model = ExcelModelTemplate(model.save(tmp_path/'released.xlsx'))
    with model._open_reader() as reader:
        values=reader.get_table('Products')['ID']
        assert pd.isna(values.iloc[1])
        assert values.iloc[0] == values.iloc[2] == 1
    doc=minidom.parseString(model._image.read_logical_file(model._dimension_paths['Products']))
    binding=next(n for n in _elements(doc,'AttributeRelationship') if _text(_child(n,'AttributeID'))=='ID')
    assert _text(_child(binding,'Cardinality'))=='Many'


def test_explicit_cardinality_and_revalidation(tmp_path):
    from excel_model_lab import RelationshipValidationError
    from excel_model_lab.relationship_storage import stored_relationships
    m = ExcelModelTemplate.create({'Customers': pd.DataFrame({'ID': [1, 2]}),
        'Details': pd.DataFrame({'ID': [1, 2], 'Amount': [10, 20]})})
    m.add_relationship('Details', 'ID', 'Customers', 'ID', cardinality='one-to-one')
    path = m.save(tmp_path/'one.xlsx')
    reopened = ExcelModelTemplate(path)
    assert stored_relationships(reopened)[0][1].cardinality == 'one-to-one'
    reopened.replace_table('Details', pd.DataFrame({'ID': [1, 1], 'Amount': [10, 20]}))
    destination = tmp_path/'protected.xlsx'
    destination.write_bytes(b'preserve me')
    with pytest.raises(RelationshipValidationError, match=r'Details.ID.*rows 1 and 2'):
        reopened.save(destination)
    assert destination.read_bytes() == b'preserve me'


def test_relationship_validation_is_immediate_and_atomic():
    from excel_model_lab import RelationshipValidationError
    m = ExcelModelTemplate.create({'Customers': pd.DataFrame({'ID': [1, 1]}),
        'Sales': pd.DataFrame({'ID': [1, 2]})})
    with pytest.raises(RelationshipValidationError, match=r'Customers.ID.*duplicate key 1'):
        m.add_relationship('Customers', 'ID', 'Sales', 'ID', cardinality='one-to-many')
    assert not m._pending_relationships
    with pytest.raises(RelationshipValidationError, match='many-to-many'):
        m.add_relationship('Sales', 'ID', 'Customers', 'ID', cardinality='many-to-many')
    assert not m._pending_relationships


def test_one_to_many_normalizes_direction(tmp_path):
    from excel_model_lab.relationship_storage import stored_relationships
    m = ExcelModelTemplate.create({'Customers': pd.DataFrame({'ID': [1, 2]}),
        'Sales': pd.DataFrame({'ID': [1, 1, 2]})})
    m.add_relationship('Customers', 'ID', 'Sales', 'ID', cardinality='one-to-many')
    reopened = ExcelModelTemplate(m.save(tmp_path/'many.xlsx'))
    r = stored_relationships(reopened)[0][1]
    assert (r.from_table, r.to_table, r.cardinality) == ('Sales', 'Customers', 'many-to-one')


def test_activation_preserves_one_to_one(tmp_path):
    from excel_model_lab.relationship_storage import stored_relationships
    m = ExcelModelTemplate.create({'A': pd.DataFrame({'ID': [1, 2]}),
        'B': pd.DataFrame({'ID': [1, 2]})})
    m.add_relationship('A', 'ID', 'B', 'ID', cardinality='one-to-one')
    m.set_relationship_active('A', 'ID', 'B', 'ID', active=False)
    reopened = ExcelModelTemplate(m.save(tmp_path/'inactive-one.xlsx'))
    r = stored_relationships(reopened)[0][1]
    assert r.cardinality == 'one-to-one' and not r.active
