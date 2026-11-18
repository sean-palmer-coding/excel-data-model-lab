"""Structural name rebinding with lexical DAX edits and stable physical IDs."""
from xml.dom import minidom
from xml.etree import ElementTree as ET
from xml.parsers.expat import ExpatError
from .abf import AbfImage, UnsupportedModelError
from .formula_statements import rewrite_references, formula_references, iter_measure_definitions
from .schema import _child, _elements


def _text(node):
    return ''.join(child.data for child in node.childNodes
                   if child.nodeType in (child.TEXT_NODE, child.CDATA_SECTION_NODE))


def _assign(node, value):
    while node.firstChild:
        node.removeChild(node.firstChild)
    node.appendChild(node.ownerDocument.createTextNode(value))


def rebind_table_references(image, table, name, *, table_names):
    """Rewrite DAX references and known display metadata, never source SQL."""
    replacements = {}
    for file in image.files:
        if not file.logical_path or not file.logical_path.endswith(('.dim.xml', '.scr.xml', '.cub.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        changed = False
        for node in _elements(doc, 'Expression') + _elements(doc, 'Text'):
            value = _text(node)
            updated = rewrite_references(value, table_names=table_names, table_renames={table: name})
            if updated != value:
                _assign(node, updated)
                changed = True
        for annotation in _elements(doc, 'Annotation'):
            annotation_name = _text(_child(annotation, 'Name'))
            value_node = _child(annotation, 'Value')
            value = _text(value_node)
            if annotation_name in ('Table', 'GoalMeasureTableName') and value.casefold() == table.casefold():
                _assign(value_node, name)
                changed = True
            elif annotation_name == 'PerspectiveEmptyHierarchyNames' and value:
                root = ET.fromstring(value)
                annotation_changed = False
                for selection in root.findall('.//Table'):
                    if selection.get('Name', '').casefold() == table.casefold():
                        selection.set('Name', name)
                        annotation_changed = True
                if annotation_changed:
                    _assign(value_node, '<?xml version="1.0" encoding="UTF-16"?>' + ET.tostring(root, encoding='unicode'))
                    changed = True
        if changed:
            replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements)) if replacements else image


def rebind_column_expression(expression, home_table, table, column, name, *, table_names, measure_names=()):
    """Rebind a column only when lexical resolution cannot mean a measure."""
    matching = [ref for ref in formula_references(expression, table_names=table_names)
                if ref.kind == 'field' and ref.name.casefold() == column.casefold()
                and ((ref.table is not None and ref.table.casefold() == table.casefold())
                     or (ref.table is None and home_table.casefold() == table.casefold()))]
    if matching and column.casefold() in {value.casefold() for value in measure_names}:
        raise UnsupportedModelError('Column rename cannot resolve an ambiguous column/measure reference')
    renames = {(table, column): name}
    if home_table.casefold() == table.casefold():
        renames[(None, column)] = name
    return rewrite_references(expression, table_names=table_names, field_renames=renames)


def rebind_column_references(image, table, column, name, *, table_names, measure_names=()):
    replacements = {}
    for file in image.files:
        if not file.logical_path or not file.logical_path.endswith(('.dim.xml', '.scr.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        changed = False
        if file.logical_path.endswith('.dim.xml'):
            dimension = _elements(doc, 'Dimension')[0]
            home = _text(_child(dimension, 'Name'))
            for expression in _elements(dimension, 'Expression'):
                original = _text(expression)
                updated = rebind_column_expression(original, home, table, column, name,
                    table_names=table_names, measure_names=measure_names)
                if updated != original:
                    _assign(expression, updated)
                    changed = True
        else:
            for text_node in _elements(doc, 'Text'):
                original = _text(text_node)
                updated = original
                for definition in reversed(list(iter_measure_definitions(original))):
                    home = definition.group('table').replace("''", "'")
                    expression = rebind_column_expression(definition.group('expression'), home, table, column, name,
                        table_names=table_names, measure_names=measure_names)
                    updated = updated[:definition.start('expression')] + expression + updated[definition.end('expression'):]
                if updated != original:
                    _assign(text_node, updated)
                    changed = True
        if changed:
            replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements)) if replacements else image


def _table_captions(image):
    result = {}
    for file in image.files:
        if file.logical_path and file.logical_path.endswith('.dim.xml'):
            doc = minidom.parseString(image.read_logical_file(file))
            for dimension in _elements(doc, 'Dimension'):
                if getattr(dimension.parentNode, 'localName', None) == 'ObjectDefinition':
                    result[_text(_child(dimension, 'ID'))] = _text(_child(dimension, 'Name'))
    return result


def rename_measure_references(image, table, measure, name, *, table_names):
    """Rebind a measure declaration and resolved DAX uses, preserving metadata.

    The caller guards model/report dependencies and ambiguous field resolution.
    Virtual-column constructors remain unsupported because this lexical writer
    cannot determine their row-context name binding.
    """
    import re
    from .formula_statements import protected_spans
    from .dependencies import DependencyError
    from .validation import ValidationIssue
    replacements = {}

    def expression(value, source):
        references = [r for r in formula_references(value, table_names=table_names)
                      if r.kind == 'field' and r.name.casefold() == measure.casefold()
                      and (r.table is None or r.table.casefold() == table.casefold())]
        if not references:
            return value
        masked = list(value)
        for start, end in protected_spans(value):
            masked[start:end] = ' ' * (end - start)
        if re.search(r'\b(?:ADDCOLUMNS|SELECTCOLUMNS|DATATABLE|ROW|SUMMARIZECOLUMNS|SUMMARIZE|GROUPBY|GENERATEALL|GENERATE)\s*\(', ''.join(masked), re.I):
            message = 'Measure rename cannot resolve virtual-column DAX references'
            raise DependencyError(message, issues=(ValidationIssue('ambiguous_reference', message,
                object_path=('tables', table, 'measures', measure), dependencies=(source,)),))
        return rewrite_references(value, table_names=table_names,
            field_renames={(None, measure): name, (table, measure): name})

    for file in image.files:
        if not file.logical_path or not file.logical_path.endswith(('.scr.xml', '.dim.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        changed = False
        if file.logical_path.endswith('.dim.xml'):
            dimension = next(d for d in _elements(doc, 'Dimension')
                             if getattr(d.parentNode, 'localName', None) == 'ObjectDefinition')
            home = _text(_child(dimension, 'Name'))
            for attribute in _elements(_child(dimension, 'Attributes'), 'Attribute'):
                column = _text(_child(attribute, 'Name')) or _text(_child(attribute, 'ID'))
                for node in _elements(attribute, 'Expression'):
                    old = _text(node)
                    new = expression(old, ('tables', home, 'columns', column))
                    if new != old:
                        _assign(node, new)
                        changed = True
        else:
            for node in _elements(doc, 'Text'):
                old = _text(node)
                edits = []
                definitions = list(iter_measure_definitions(old))
                remaining = list(old)
                for definition in definitions:
                    remaining[definition.start():definition.end()] = ' ' * (definition.end() - definition.start())
                if any(ref.kind == 'field' and ref.name.casefold() == measure.casefold()
                       for ref in formula_references(''.join(remaining), table_names=table_names)):
                    message = 'Measure rename cannot rewrite this non-measure script statement'
                    raise DependencyError(message, issues=(ValidationIssue('unsupported_script_reference', message,
                        object_path=('tables', table, 'measures', measure),
                        dependencies=(('scripts', file.logical_path),)),))
                for definition in definitions:
                    home = definition.group('table').replace("''", "'")
                    current = definition.group('name').replace(']]', ']')
                    new = expression(definition.group('expression'), ('tables', home, 'measures', current))
                    edits.append((definition.start('expression'), definition.end('expression'), new))
                    if home.casefold() == table.casefold() and current.casefold() == measure.casefold():
                        edits.append((definition.start('name'), definition.end('name'), name.replace(']', ']]')))
                new = old
                for start, end, value in sorted(edits, reverse=True):
                    new = new[:start] + value + new[end:]
                if new != old:
                    _assign(node, new)
                    changed = True
            for command in _elements(doc, 'Command'):
                containers = [c for c in command.childNodes if getattr(c, 'localName', None) == 'Annotations']
                if not containers:
                    continue
                annotations = {_text(_child(a, 'Name')): _child(a, 'Value')
                               for a in _elements(containers[0], 'Annotation')}
                if ('Table' in annotations and 'FullName' in annotations
                        and _text(annotations['Table']).casefold() == table.casefold()
                        and _text(annotations['FullName']).casefold() == measure.casefold()):
                    _assign(annotations['FullName'], name)
                    changed = True
            for prop in _elements(doc, 'CalculationProperty'):
                reference = _child(prop, 'CalculationReference')
                if _text(reference).casefold() == ('[' + measure.replace(']', ']]') + ']').casefold():
                    _assign(reference, '[' + name.replace(']', ']]') + ']')
                    changed = True
        if changed:
            replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements))


def renamed_table_package_parts(original, replacements, model_bytes):
    """Rebind Power Pivot's own layout captions without touching report parts."""
    if 'xl/model/item.data' not in original.namelist():
        return {}
    old = _table_captions(AbfImage(original.read('xl/model/item.data')))
    new = _table_captions(AbfImage(model_bytes))
    mapping = {caption: new[identity] for identity, caption in old.items()
               if identity in new and new[identity] != caption}
    if not mapping:
        return {}
    read = lambda path: replacements.get(path) or original.read(path)
    result = {}
    workbook = minidom.parseString(read('xl/workbook.xml'))
    changed = False
    for table in _elements(workbook, 'modelTable'):
        identifier = table.getAttribute('id')
        if identifier in old and identifier in new and old[identifier] != new[identifier]:
            table.setAttribute('name', new[identifier])
            changed = True
    if changed:
        result['xl/workbook.xml'] = workbook.toxml(encoding='utf-8')
    for path in original.namelist():
        if not path.startswith('customXml/') or not path.endswith('.xml'):
            continue
        try:
            outer = minidom.parseString(read(path))
        except ExpatError:
            continue
        if outer.documentElement.namespaceURI != 'http://gemini/pivotcustomization/Diagrams':
            continue
        content = _child(outer.documentElement, 'CustomContent')
        payload = ''.join(n.data for n in content.childNodes if n.nodeType in (n.TEXT_NODE, n.CDATA_SECTION_NODE))
        inner = minidom.parseString(payload)
        changed = False
        for node in _elements(inner, 'TableName'):
            if _text(node) in mapping:
                _assign(node, mapping[_text(node)])
                changed = True
        if changed:
            while content.firstChild:
                content.removeChild(content.firstChild)
            # Text nodes safely escape embedded XML without losing its content.
            content.appendChild(outer.createTextNode(inner.documentElement.toxml()))
            result[path] = outer.toxml(encoding='utf-16')
    return result
