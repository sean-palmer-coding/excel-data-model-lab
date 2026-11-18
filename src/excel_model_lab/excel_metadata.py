"""Read-only snapshots of Excel's legacy ASSL metadata.

Names and stable IDs are both retained. Processing state describes the saved
file; an expression is not evidence that its cached values are calculated.
"""
from dataclasses import dataclass
from xml.etree import ElementTree as ET
import pandas as pd


def _root(payload):
    root = ET.fromstring(payload)
    for node in root.iter():
        node.tag = node.tag.rsplit('}', 1)[-1]
    return root


def _text(node, path, default=''):
    return node.findtext(path, default) or default


def _annotations(node):
    result = {}
    for a in node.findall('Annotations/Annotation'):
        value = a.find('Value')
        result[_text(a, 'Name')] = ('' if value is None else
            ''.join(value.itertext()) if not len(value) else
            ''.join(ET.tostring(child, encoding='unicode') for child in value))
    return result


@dataclass
class ExcelMetadata:
    """Independent DataFrame snapshots; modifying these does not edit a model.

    ``available`` means ASSL dimension metadata was found. Perspective
    calculations retain qualified MDX names, including native hidden defaults.
    ``perspective_hierarchy_annotations`` preserves Excel's additional hierarchy
    selection text verbatim: it is not equivalent to ordinary hierarchy IDs.
    """
    available: bool
    tables: pd.DataFrame
    columns: pd.DataFrame
    hierarchies: pd.DataFrame
    levels: pd.DataFrame
    perspectives: pd.DataFrame
    perspective_columns: pd.DataFrame
    perspective_measures: pd.DataFrame
    perspective_hierarchies: pd.DataFrame
    perspective_hierarchy_annotations: pd.DataFrame

    @property
    def date_tables(self):
        return self.tables.loc[self.tables['is_date_table']].copy()

    @property
    def calculated_columns(self):
        return self.columns.loc[self.columns['expression'].notna()].copy()

    @classmethod
    def from_files(cls, files):
        """Parse an iterable of (logical filename, XML bytes) metadata streams."""
        tables, columns, hierarchies, levels = [], [], [], []
        perspectives, pcolumns, pmeasures, phierarchies, pannotations = [], [], [], [], []
        visibility = {}
        for name, payload in files:
            root = _root(payload)
            if name.endswith('.cub.xml'):
                for dim in root.findall('.//Cube/Dimensions/Dimension'):
                    visibility[_text(dim, 'ID')] = _text(dim, 'Visible', 'true') == 'true'
                for cube in root.findall('.//Cube'):
                    raw = _annotations(cube).get('PerspectiveEmptyHierarchyNames')
                    if raw:
                        annotation = _root(raw)
                        for perspective in annotation.findall('Perspective'):
                            for table in perspective.findall('Tables/Table'):
                                pannotations.append((perspective.get('Name', ''),
                                    table.get('Name', ''), table.get('EmptyHierarchyNames', '')))
            elif name.endswith('.dim.xml'):
                for dim in root.findall('.//ObjectDefinition/Dimension'):
                    table_id, table = _text(dim, 'ID'), _text(dim, 'Name')
                    date = _text(dim, 'Type') == 'Time'
                    keys = []
                    for attr in dim.findall('Attributes/Attribute'):
                        if _text(attr, 'Type') == 'RowNumber':
                            continue
                        column_id = _text(attr, 'ID')
                        if _text(attr, 'Usage') == 'Key':
                            keys.append(column_id)
                        expression = attr.find('KeyColumns/KeyColumn/Source/Expression')
                        columns.append((table_id, table, column_id, _text(attr, 'Name'),
                            _text(attr, 'InferredDatatype') or _text(attr, 'KeyColumns/KeyColumn/DataType'),
                            None if expression is None else expression.text or '',
                            _text(attr, 'ProcessingState'),
                            _text(attr, 'AttributeHierarchyVisible', 'true') != 'true',
                            _text(attr, 'Description'), _text(attr, 'AttributeHierarchyDisplayFolder'),
                            _text(attr, 'FormatString'), _text(attr, 'OrderByAttributeID'), _annotations(attr)))
                    tables.append((table_id, table, date, tuple(keys) if date else (),
                        _text(dim, 'Description'), _text(dim, 'ProcessingState'),
                        _text(dim, 'Language'), _text(dim, 'Collation'), False, _annotations(dim)))
                    for hierarchy in dim.findall('Hierarchies/Hierarchy'):
                        hid = _text(hierarchy, 'ID')
                        hierarchies.append((table_id, table, hid, _text(hierarchy, 'Name'),
                            _text(hierarchy, 'Visible', 'true') != 'true',
                            _text(hierarchy, 'DisplayFolder'), _text(hierarchy, 'Description'),
                            _text(hierarchy, 'ProcessingState'), _text(hierarchy, 'StructureType')))
                        for ordinal, level in enumerate(hierarchy.findall('Levels/Level')):
                            levels.append((table_id, table, hid, _text(level, 'ID'),
                                _text(level, 'Name'), ordinal, _text(level, 'SourceAttributeID')))
            elif name.endswith('.persp.xml'):
                for perspective in root.findall('.//ObjectDefinition/Perspective'):
                    pid = _text(perspective, 'ID')
                    perspectives.append((pid, _text(perspective, 'Name'), _text(perspective, 'Description'),
                        tuple(_text(m, 'MeasureGroupID') for m in perspective.findall('MeasureGroups/MeasureGroup'))))
                    for dim in perspective.findall('Dimensions/Dimension'):
                        tid = _text(dim, 'CubeDimensionID')
                        for attr in dim.findall('Attributes/Attribute'):
                            pcolumns.append((pid, tid, _text(attr, 'AttributeID'),
                                _text(attr, 'AttributeHierarchyVisible', 'true') == 'true'))
                        for hierarchy in dim.findall('Hierarchies/Hierarchy'):
                            phierarchies.append((pid, tid, _text(hierarchy, 'HierarchyID')))
                    for calc in perspective.findall('Calculations/Calculation'):
                        pmeasures.append((pid, _text(calc, 'Name'), _text(calc, 'Type')))
        perspective_ids = {row[1].casefold(): row[0] for row in perspectives}
        table_ids = {row[1].casefold(): row[0] for row in tables}
        hierarchy_ids = {(row[0], row[3].casefold()): row[2] for row in hierarchies}
        for perspective, table, selection in pannotations:
            pid, tid = perspective_ids.get(perspective.casefold()), table_ids.get(table.casefold())
            for name in selection.split(','):
                hid = hierarchy_ids.get((tid, name.casefold()))
                member = (pid, tid, hid)
                if pid is not None and tid is not None and hid is not None and member not in phierarchies:
                    phierarchies.append(member)
        t = pd.DataFrame(tables, columns=['id', 'name', 'is_date_table', 'date_key_columns',
            'description', 'processing_state', 'language', 'collation', 'hidden', 'annotations'])
        t['is_date_table'] = t['is_date_table'].astype(bool)
        t['hidden'] = [not visibility.get(tid, True) for tid in t['id']]
        return cls(bool(tables), t,
            pd.DataFrame(columns, columns=['table_id', 'table', 'id', 'name', 'data_type', 'expression',
                'processing_state', 'hidden', 'description', 'display_folder', 'format_string',
                'sort_by_column_id', 'annotations']),
            pd.DataFrame(hierarchies, columns=['table_id', 'table', 'id', 'name', 'hidden', 'display_folder', 'description', 'processing_state', 'structure_type']),
            pd.DataFrame(levels, columns=['table_id', 'table', 'hierarchy_id', 'id', 'name', 'ordinal', 'column_id']),
            pd.DataFrame(perspectives, columns=['id', 'name', 'description', 'measure_group_ids']),
            pd.DataFrame(pcolumns, columns=['perspective_id', 'table_id', 'column_id', 'visible']),
            pd.DataFrame(pmeasures, columns=['perspective_id', 'name', 'type']),
            pd.DataFrame(phierarchies, columns=['perspective_id', 'table_id', 'hierarchy_id']),
            pd.DataFrame(pannotations, columns=['perspective', 'table', 'empty_hierarchy_names']))
