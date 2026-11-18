"""Persist native Power Pivot measure cells through workbook custom XML."""
from pathlib import Path
from xml.dom import minidom
from xml.parsers.expat import ExpatError
import posixpath
import re
import uuid
from .abf import AbfImage
from .schema import _elements, _child, _text
from .measures import list_script_measures

REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
PREFIX = 'http://gemini/pivotcustomization/'

def text(node, value):
    while node.firstChild:
        node.removeChild(node.firstChild)
    node.appendChild(node.ownerDocument.createTextNode(str(value)))

def layout_parts(original, replacements, model):
    image = AbfImage(model)
    scripts = [f for f in image.files if f.logical_path and f.logical_path.endswith('.scr.xml')]
    if not scripts:
        return {}
    script = image.read_logical_file(scripts[0])
    doc = minidom.parseString(script)
    visible = {}
    kpis = set()
    for prop in _elements(doc, 'CalculationProperty'):
        ann = {}
        for a in _elements(prop, 'Annotation'):
            values = [n for n in a.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName == 'Value']
            ann[_text(_child(a, 'Name'))] = _text(values[0]) if values else ''
        ref = _text(_child(prop, 'CalculationReference'))
        if ann.get('Type') in ('User', 'Kpi') and not any(_text(n) == 'false' for n in _elements(prop, 'Visible')):
            visible[ref] = True
            if ann.get('Type') == 'Kpi':
                kpis.add(ref)
    tables = {}
    for measure in list_script_measures(script):
        if '[' + measure.name.replace(']', ']]') + ']' in visible:
            tables.setdefault(measure.table, []).append(measure.name)
    removed = set()
    if 'xl/model/item.data' in original.namelist():
        previous = AbfImage(original.read('xl/model/item.data'))
        old_scripts = [f for f in previous.files if f.logical_path and f.logical_path.endswith('.scr.xml')]
        if old_scripts:
            old = {(m.table, m.name) for m in list_script_measures(previous.read_logical_file(old_scripts[0]))}
            current = {(m.table, m.name) for m in list_script_measures(script)}
            removed = old - current
    if not tables and not removed:
        return {}
    names = set(original.namelist())
    def read(path):
        return replacements[path] if path in replacements else original.read(path)
    relpath = 'xl/_rels/workbook.xml.rels'
    rels = minidom.parseString(read(relpath))
    types = minidom.parseString(read('[Content_Types].xml'))
    result = {}
    def custom(kind):
        for rel in _elements(rels, 'Relationship'):
            if rel.getAttribute('Type') != CUSTOM or rel.getAttribute('TargetMode') == 'External':
                continue
            path = posixpath.normpath(posixpath.join('xl', rel.getAttribute('Target'))).lstrip('/')
            if path not in names:
                continue
            try:
                outer = minidom.parseString(read(path))
            except ExpatError:
                # Other add-ins may store opaque payloads in custom XML parts.
                continue
            if outer.documentElement.namespaceURI == PREFIX + kind:
                content = _child(outer.documentElement, 'CustomContent')
                payload = ''.join(n.data for n in content.childNodes
                                  if n.nodeType in (n.TEXT_NODE, n.CDATA_SECTION_NODE))
                return path, outer, minidom.parseString(payload)
        return None, None, None
    def commit(kind, path, outer, inner):
        if path is None:
            index = 1
            while f'customXml/item{index}.xml' in names or f'customXml/item{index}.xml' in result or f'customXml/itemProps{index}.xml' in names:
                index += 1
            path = f'customXml/item{index}.xml'
            outer = minidom.parseString(f'<Gemini xmlns="{PREFIX + kind}"><CustomContent/></Gemini>')
            prop = f'customXml/itemProps{index}.xml'
            ds = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml'
            result[prop] = (f'<ds:datastoreItem ds:itemID="{{{str(uuid.uuid4()).upper()}}}" xmlns:ds="{ds}">'
                            '<ds:schemaRefs/></ds:datastoreItem>').encode()
            result[f'customXml/_rels/item{index}.xml.rels'] = (f'<Relationships xmlns="{REL}"><Relationship Id="rId1" '
                f'Type="{CUSTOM}Props" Target="itemProps{index}.xml"/></Relationships>').encode()
            ids = {r.getAttribute('Id') for r in _elements(rels, 'Relationship')}
            number = 1
            while f'rId{number}' in ids:
                number += 1
            r = rels.createElementNS(REL, 'Relationship')
            for key, value in [('Id', f'rId{number}'), ('Type', CUSTOM), ('Target', '../' + path)]:
                r.setAttribute(key, value)
            rels.documentElement.appendChild(r)
            override = types.createElementNS(types.documentElement.namespaceURI, 'Override')
            override.setAttribute('PartName', '/' + prop)
            override.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml')
            types.documentElement.appendChild(override)
        content = _child(outer.documentElement, 'CustomContent')
        while content.firstChild:
            content.removeChild(content.firstChild)
        # Split CDATA terminators if a user name contains one.
        payload = inner.documentElement.toxml()
        chunks = payload.split(']]>')
        for index, chunk in enumerate(chunks):
            content.appendChild(outer.createCDATASection(('>' if index else '') + chunk + (']]' if index < len(chunks) - 1 else '')))
        result[path] = outer.toxml(encoding='utf-16')
    path, outer, diagrams = custom('Diagrams')
    template = minidom.parse(str(Path(__file__).parent / 'data' / 'measure-diagram.xml'))
    prototype = template.documentElement.firstChild
    if diagrams is None:
        diagrams = minidom.parseString(template.documentElement.toxml())
        for n in list(diagrams.documentElement.childNodes):
            diagrams.documentElement.removeChild(n)
    changed = False
    for diagram in _elements(diagrams, 'DiagramManager.SerializableDiagram'):
        table_names = {_text(n) for n in _elements(diagram, 'TableName')}
        prefixes = ['Measures\\' + name for table, name in removed if table in table_names]
        for container in _elements(diagram, 'AllKeys') + _elements(diagram, 'ViewStates'):
            for node in list(container.childNodes):
                if any(_text(k) == prefix or _text(k).startswith(prefix + '\\TagInfo\\')
                       for k in _elements(node, 'Key') for prefix in prefixes):
                    container.removeChild(node)
                    changed = True
    for table, measures in tables.items():
        matches = [d for d in _elements(diagrams, 'DiagramManager.SerializableDiagram')
                   if any(_text(n) == table for n in _elements(d, 'TableName'))
                   and any(_text(n) == 'MeasureDiagram' for n in _elements(d, 'DiagramType'))]
        if matches:
            diagram = matches[0]
        else:
            diagram = diagrams.importNode(prototype, True)
            for n in _elements(diagram, 'TableName'):
                text(n, table)
            for container in (_elements(diagram, 'AllKeys')[0], _elements(diagram, 'ViewStates')[0]):
                for n in list(container.childNodes):
                    if any(_text(k).startswith(('Measures\\', 'Columns\\')) for k in _elements(n, 'Key')):
                        container.removeChild(n)
            diagrams.documentElement.appendChild(diagram)
            changed = True
        keys = _elements(diagram, 'AllKeys')[0]
        states = _elements(diagram, 'ViewStates')[0]
        present = {_text(k) for k in _elements(keys, 'Key')}
        rows = [int(_text(r)) for r in _elements(states, 'Row')]
        next_row = max(rows + ([-1] if not any(k.startswith('Measures\\') for k in present) else [0])) + 1
        source_states = _elements(prototype, 'ViewStates')[0]
        missing = [m for m in measures if 'Measures\\' + m not in present]
        positions = {m: next_row + index for index, m in enumerate(sorted(missing,
            key=lambda m: ('[' + m.replace(']', ']]') + ']') not in kpis))}
        for measure in measures:
            if 'Measures\\' + measure in present:
                continue
            for source in list(_elements(prototype, 'AllKeys')[0].childNodes):
                if any(_text(k).startswith('Measures\\Target Amount') for k in _elements(source, 'Key')):
                    clone = diagrams.importNode(source, True)
                    for k in _elements(clone, 'Key'):
                        if not any(c.nodeType == c.ELEMENT_NODE for c in k.childNodes):
                            text(k, _text(k).replace('Measures\\Target Amount', 'Measures\\' + measure))
                    keys.appendChild(clone)
            for source in list(source_states.childNodes):
                if any(_text(k).startswith('Measures\\Target Amount') for k in _elements(source, 'Key')):
                    clone = diagrams.importNode(source, True)
                    for k in _elements(clone, 'Key'):
                        if not any(c.nodeType == c.ELEMENT_NODE for c in k.childNodes):
                            text(k, _text(k).replace('Measures\\Target Amount', 'Measures\\' + measure))
                    for r in _elements(clone, 'Row'):
                        if positions[measure] == 0:
                            r.parentNode.removeChild(r)
                        else:
                            text(r, positions[measure])
                    states.appendChild(clone)
            next_row += 1
            changed = True
        dimension = next((_elements(minidom.parseString(image.read_logical_file(f)), 'Dimension')[0]
                          for f in image.files if f.logical_path and f.logical_path.endswith('.dim.xml')
                          and _text(_child(_elements(minidom.parseString(image.read_logical_file(f)), 'Dimension')[0], 'Name')) == table), None)
        if dimension is not None:
            columns = [_text(_child(a, 'Name')) for a in _child(dimension, 'Attributes').childNodes
                       if a.nodeType == a.ELEMENT_NODE and a.localName == 'Attribute'
                       and not any(_text(t) == 'RowNumber' for t in _elements(a, 'Type'))]
            for column in columns:
                if 'Columns\\' + column in present:
                    continue
                for source_parent, target_parent in [(_elements(prototype, 'AllKeys')[0], keys), (source_states, states)]:
                    for source in source_parent.childNodes:
                        if any(_text(k) == 'Columns\\Amount' for k in _elements(source, 'Key')):
                            clone = diagrams.importNode(source, True)
                            for k in _elements(clone, 'Key'):
                                if not any(c.nodeType == c.ELEMENT_NODE for c in k.childNodes):
                                    text(k, 'Columns\\' + column)
                            target_parent.appendChild(clone)
                changed = True
    if changed:
        commit('Diagrams', path, outer, diagrams)
    path, outer, grid = custom('MeasureGridState')
    grid_template = minidom.parse(str(Path(__file__).parent / 'data' / 'measure-grid.xml'))
    grid_prototype = grid_template.documentElement.firstChild
    if grid is None:
        grid = minidom.parseString(grid_template.documentElement.toxml())
        for n in list(grid.documentElement.childNodes):
            grid.documentElement.removeChild(n)
    present = {_text(k) for k in _elements(grid, 'Key')}
    changed = False
    for file in image.files:
        if not file.logical_path or not file.logical_path.endswith('.dim.xml'):
            continue
        dimension = _elements(minidom.parseString(image.read_logical_file(file)), 'Dimension')[0]
        name, identifier = _text(_child(dimension, 'Name')), _text(_child(dimension, 'ID'))
        if name in tables and identifier not in present:
            clone = grid.importNode(grid_prototype, True)
            text(_elements(clone, 'Key')[0], identifier)
            grid.documentElement.appendChild(clone)
            changed = True
    if changed:
        commit('MeasureGridState', path, outer, grid)
    if result:
        result[relpath] = rels.toxml(encoding='utf-8')
        result['[Content_Types].xml'] = types.toxml(encoding='utf-8')
    return result
