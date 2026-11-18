"""Invalidate embedded-model OLAP sheet-function results after model replacement."""
import posixpath
from xml.dom import minidom

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'
CONNECTION_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections'


def _children(node, namespace, name):
    return [child for child in node.childNodes if child.nodeType == child.ELEMENT_NODE
            and child.namespaceURI == namespace and child.localName == name]


def invalidated_cube_cache_parts(original, replacements):
    """Clear only tuple result payloads bound to the embedded model.

    Preserve the tupleCache identity flag, cache metadata, namespace prefixes,
    and all external caches. Excel rebuilds these results on calculation.
    Caller must only invoke this when the model bytes changed.
    """
    names = set(original.namelist()) | set(replacements)
    def read(name):
        return replacements[name] if name in replacements else original.read(name)
    relpath = 'xl/_rels/workbook.xml.rels'
    if relpath not in names:
        return {}
    rels = minidom.parseString(read(relpath))
    relations = [n for n in _children(rels.documentElement, R, 'Relationship')
                 if n.getAttribute('Type') == CONNECTION_REL]
    if not relations:
        return {}
    if len(relations) != 1 or relations[0].getAttribute('TargetMode') == 'External':
        raise ValueError('Cannot invalidate model caches with ambiguous connections relationships')
    target = relations[0].getAttribute('Target')
    part = posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('xl', target))
    if not part.startswith('xl/') or part not in names:
        raise ValueError('Cannot resolve workbook connections for model cache invalidation')
    connections = minidom.parseString(read(part))
    model_ids = set()
    seen = set()
    for node in _children(connections.documentElement, S, 'connection'):
        cid = node.getAttribute('id')
        if not cid or cid in seen:
            raise ValueError('Workbook connection IDs must be unique for model cache invalidation')
        seen.add(cid)
        dbs = _children(node, S, 'dbPr')
        if (node.getAttribute('name').casefold() == 'thisworkbookdatamodel' and len(dbs) == 1
                and dbs[0].getAttribute('connection') == 'Data Model Connection'):
            model_ids.add(cid)
    if not model_ids:
        return {}
    updates = {}
    # Content-type discovery preserves valid nonstandard PivotCache part names.
    types = minidom.parseString(read('[Content_Types].xml'))
    T = 'http://schemas.openxmlformats.org/package/2006/content-types'
    mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'
    cache_parts = {n.getAttribute('PartName').lstrip('/')
                   for n in _children(types.documentElement, T, 'Override')
                   if n.getAttribute('ContentType') == mime}
    for name in sorted(cache_parts):
        if name not in names:
            raise ValueError('PivotCache content type points to a missing package part')
        doc = minidom.parseString(read(name))
        root = doc.documentElement
        if root.namespaceURI != S or root.localName != 'pivotCacheDefinition':
            raise ValueError('PivotCache definition has an invalid root')
        sources = _children(root, S, 'cacheSource')
        if len(sources) != 1 or sources[0].getAttribute('connectionId') not in model_ids:
            continue
        payloads = _children(root, S, 'tupleCache')
        if not payloads:
            continue
        for node in payloads:
            root.removeChild(node)
            node.unlink()
        updates[name] = doc.toxml(encoding='utf-8')
    if updates:
        workbook = minidom.parseString(read('xl/workbook.xml'))
        root = workbook.documentElement
        calc_nodes = _children(root, S, 'calcPr')
        if len(calc_nodes) > 1:
            raise ValueError('Workbook contains multiple calculation property elements')
        if calc_nodes:
            calc = calc_nodes[0]
        else:
            calc = workbook.createElementNS(S, (root.prefix + ':' if root.prefix else '') + 'calcPr')
            later_names = ('oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr',
                           'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst')
            later = next((n for n in root.childNodes if n.nodeType == n.ELEMENT_NODE
                          and n.namespaceURI == S and n.localName in later_names), None)
            root.insertBefore(calc, later)
        # Preserve intentional manual calculation mode; request a full rebuild
        # when Excel calculates instead of retaining old sheet-function results.
        for key, value in (('calcId', '0'), ('fullCalcOnLoad', '1'), ('forceFullCalc', '1')):
            calc.setAttribute(key, value)
        updates['xl/workbook.xml'] = workbook.toxml(encoding='utf-8')
    return updates
