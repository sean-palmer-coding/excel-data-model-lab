"""Model-backed OLAP slicer and timeline package parts.

The dedicated control PivotCaches are separate from report PivotCaches. Their
x14 pivotCacheId identities are not the workbook's cacheId identities.
"""
from __future__ import annotations

import posixpath
import re
from datetime import datetime
from uuid import uuid4
from xml.dom import minidom

from .errors import ModelValidationError
from .report_filters import member_unique_name, _caption, source_member_values

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'
X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
XR = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision'
XR10 = 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10'
SLE = 'http://schemas.microsoft.com/office/drawing/2010/slicer'
TSLE = 'http://schemas.microsoft.com/office/drawing/2012/timeslicer'


def _children(node, namespace=None, local=None):
    return [n for n in node.childNodes if n.nodeType == n.ELEMENT_NODE
            and (namespace is None or n.namespaceURI == namespace)
            and (local is None or n.localName == local)]


def _node(parent, namespace, qname, **attrs):
    doc = parent if parent.nodeType == parent.DOCUMENT_NODE else parent.ownerDocument
    node = doc.createElementNS(namespace, qname)
    prefix = qname.split(':')[0] if ':' in qname else None
    node.setAttribute('xmlns:'+prefix if prefix else 'xmlns', namespace)
    for key, value in attrs.items():
        if key == 'rid':
            node.setAttribute('xmlns:r', R)
            node.setAttributeNS(R, 'r:id', str(value))
        else:
            node.setAttribute(key, str(value))
    parent.appendChild(node)
    return node


def _new(namespace, name, **attrs):
    doc = minidom.Document()
    return _node(doc, namespace, name, **attrs)


def _extension(root, uri, namespace, name):
    extlst = next(iter(_children(root, S, 'extLst')), None)
    if extlst is None:
        extlst = _node(root, S, 'extLst')
    ext = next((n for n in _children(extlst) if n.getAttribute('uri').upper() == uri.upper()), None)
    if ext is None:
        ext = _node(extlst, S, 'ext', uri=uri)
    else:
        # Excel compares these extension identifiers as case-sensitive URIs,
        # despite their GUID-like spelling. Preserve the native spelling.
        ext.setAttribute('uri', uri)
    return next(iter(_children(ext, namespace, name.split(':')[-1])), None) or _node(ext, namespace, name)


def _mdx(value):
    return '['+value.replace(']', ']]')+']'


def _control_uid(node):
    node.setAttribute('xmlns:xr10', XR10)
    node.setAttribute('xmlns:mc', MC)
    node.setAttributeNS(XR10, 'xr10:uid', '{'+str(uuid4()).upper()+'}')
    node.setAttributeNS(MC, 'mc:Ignorable', 'xr10')


def _relpath(path):
    return posixpath.dirname(path)+'/_rels/'+posixpath.basename(path)+'.rels'


class _Package:
    def __init__(self, parts):
        self.parts = dict(parts)
        self.docs = {}

    def root(self, path, namespace=None, name=None):
        if path not in self.docs:
            self.docs[path] = (minidom.parseString(self.parts[path]) if path in self.parts
                               else _new(namespace, name).ownerDocument)
        return self.docs[path].documentElement

    def alloc(self, prefix):
        number = 1
        while f'{prefix}{number}.xml' in self.parts or f'{prefix}{number}.xml' in self.docs:
            number += 1
        return f'{prefix}{number}.xml'

    def relation(self, owner, kind, target):
        root = self.root(_relpath(owner), P, 'Relationships')
        ids = {n.getAttribute('Id') for n in _children(root)}
        number = 1
        while 'rId'+str(number) in ids:
            number += 1
        rid = 'rId'+str(number)
        _node(root, P, 'Relationship', Id=rid, Type=kind,
              Target=posixpath.relpath(target, posixpath.dirname(owner)))
        return rid

    def target(self, owner, rid):
        root = self.root(_relpath(owner))
        found = [n for n in _children(root) if n.getAttribute('Id') == rid]
        if len(found) != 1 or found[0].getAttribute('TargetMode') == 'External':
            raise ModelValidationError('Control requires an internal package relationship')
        return posixpath.normpath(posixpath.join(posixpath.dirname(owner), found[0].getAttribute('Target'))).lstrip('/')

    def content_type(self, path, value):
        _node(self.root('[Content_Types].xml'), C, 'Override', PartName='/'+path, ContentType=value)

    def finish(self):
        for path, doc in self.docs.items():
            self.parts[path] = doc.toxml(encoding='utf-8')
        return self.parts


def _control_cache(package, report_cache, hierarchy, timeline, cache_id, internal_id):
    path = package.alloc('xl/pivotCache/pivotCacheDefinition')
    root = minidom.parseString(package.parts[report_cache]).documentElement
    package.docs[path] = root.ownerDocument
    root.setAttribute('createdVersion', '3')
    # Dedicated control caches are driven by the slicer/timeline, not the
    # report refresh-on-open path. Native Excel omits these report flags.
    for attribute in ('refreshOnLoad', 'enableRefresh'):
        if root.hasAttribute(attribute):
            root.removeAttribute(attribute)
    root.setAttribute('xmlns:xr', XR)
    root.setAttributeNS(XR, 'xr:uid', '{'+str(uuid4()).upper()+'}')
    for tag in ('cacheFields', 'dimensions', 'measureGroups', 'maps'):
        for child in _children(root, S, tag):
            if tag == 'cacheFields':
                for node in list(child.childNodes): child.removeChild(node)
                child.setAttribute('count', '0')
            else: root.removeChild(child)
    for h in root.getElementsByTagNameNS(S, 'cacheHierarchy'):
        h.setAttribute('count', '2' if h.getAttribute('uniqueName') == hierarchy else '0')
        h.removeAttribute('oneField') if h.hasAttribute('oneField') else None
        for usage in _children(h, S, 'fieldsUsage'): h.removeChild(usage)
        if h.getAttribute('memberValueDatatype') == '7': h.setAttribute('time', '1')
    source = _children(root, S, 'cacheSource')[0]
    _extension(source, '{F057638F-6D5F-4e77-A914-E7F072B9BCA8}', X14, 'x14:sourceConnection').setAttribute('name', 'ThisWorkbookDataModel')
    ext = _extension(root, '{725AE2AE-9491-48be-B2B4-4EB974FC3084}', X14, 'x14:pivotCacheDefinition')
    ext.setAttribute('pivotCacheId', str(internal_id))
    if timeline:
        version = _extension(root, '{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}', X15, 'x15:pivotCacheIdVersion')
        version.setAttribute('cacheIdSupportedVersion', '6'); version.setAttribute('cacheIdCreatedVersion', '7')
    else: ext.setAttribute('slicerData', '1')
    wb = package.root('xl/workbook.xml')
    uri, ns, name = (('{A2CB5862-8E78-49c6-8D9D-AF26E26ADB89}', X15, 'x15:timelineCachePivotCaches') if timeline
                     else ('{876F7934-8845-4945-9796-88D515C7AA90}', X14, 'x14:pivotCaches'))
    rid = package.relation('xl/workbook.xml', R+'/pivotCacheDefinition', path)
    _node(_extension(wb, uri, ns, name), S, 'pivotCache', cacheId=cache_id, rid=rid)
    package.content_type(path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml')


def _drawing(package, sheet_path, spec, timeline):
    sheet = package.root(sheet_path)
    existing = _children(sheet, S, 'drawing')
    if existing:
        path = package.target(sheet_path, existing[0].getAttributeNS(R, 'id'))
        root = package.root(path)
    else:
        path = package.alloc('xl/drawings/drawing')
        root = package.root(path, XDR, 'xdr:wsDr')
        rid = package.relation(sheet_path, R+'/drawing', path)
        drawing = _node(sheet, S, 'drawing', rid=rid)
        later = next((n for n in _children(sheet) if n.localName in ('legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'pivotTableParts', 'extLst')), None)
        if later: sheet.insertBefore(drawing, later)
        package.content_type(path, 'application/vnd.openxmlformats-officedocument.drawing+xml')
    ids = [int(n.getAttribute('id')) for n in root.getElementsByTagNameNS(XDR, 'cNvPr')]
    match = re.fullmatch(r'([A-Za-z]+)([1-9][0-9]*)', spec.cell)
    col = 0
    for char in match[1].upper(): col = col*26+ord(char)-64
    anchor = _node(root, XDR, 'xdr:oneCellAnchor')
    origin = _node(anchor, XDR, 'xdr:from')
    for name, value in [('col', col-1), ('colOff', 0), ('row', int(match[2])-1), ('rowOff', 0)]:
        child = _node(origin, XDR, 'xdr:'+name); child.appendChild(child.ownerDocument.createTextNode(str(value)))
    _node(anchor, XDR, 'xdr:ext', cx=spec.width*9525, cy=spec.height*9525)
    alternate = _node(anchor, MC, 'mc:AlternateContent')
    uri, prefix, tag = (TSLE, 'tsle', 'timeslicer') if timeline else (SLE, 'sle', 'slicer')
    choice = _node(alternate, MC, 'mc:Choice', Requires=prefix if timeline else 'a14')
    choice.setAttribute('xmlns:'+prefix if timeline else 'xmlns:a14',
                        uri if timeline else 'http://schemas.microsoft.com/office/drawing/2010/main')
    frame = _node(choice, XDR, 'xdr:graphicFrame', macro='')
    nv = _node(frame, XDR, 'xdr:nvGraphicFramePr')
    _node(nv, XDR, 'xdr:cNvPr', id=max(ids+[0])+1, name=spec.name)
    _node(nv, XDR, 'xdr:cNvGraphicFramePr')
    transform = _node(frame, XDR, 'xdr:xfrm')
    _node(transform, A, 'a:off', x=0, y=0); _node(transform, A, 'a:ext', cx=0, cy=0)
    data = _node(_node(frame, A, 'a:graphic'), A, 'a:graphicData', uri=uri)
    _node(data, uri, prefix+':'+tag, name=spec.name)
    _node(alternate, MC, 'mc:Fallback')
    _node(anchor, XDR, 'xdr:clientData')


def append_controls(parts, slicers=(), timelines=(), *, model, workbook_name=None):
    """Atomically append real OLAP control parts to a serialized report package.

    Control sheets must already exist; the report package layer creates them.
    Model bytes and unrelated members are not changed. Initial selection is All.
    """
    if not slicers and not timelines: return
    package = _Package(parts)
    wb = package.root('xl/workbook.xml')
    sheets = {}
    pivots = {}
    for sheet in wb.getElementsByTagNameNS(S, 'sheet'):
        path = package.target('xl/workbook.xml', sheet.getAttributeNS(R, 'id'))
        sheets[sheet.getAttribute('name')] = (path, sheet.getAttribute('sheetId'))
        relpath = _relpath(path)
        if relpath not in parts: continue
        for rel in _children(package.root(relpath)):
            if rel.getAttribute('Type') != R+'/pivotTable': continue
            pivot_path = package.target(path, rel.getAttribute('Id'))
            pivot = package.root(pivot_path)
            cache_rel = next(n for n in _children(package.root(_relpath(pivot_path))) if n.getAttribute('Type') == R+'/pivotCacheDefinition')
            cache = package.target(pivot_path, cache_rel.getAttribute('Id'))
            pivots[pivot.getAttribute('name')] = (sheet.getAttribute('name'), sheet.getAttribute('sheetId'), cache)
    used_names = {n.getAttribute('name').casefold() for n in wb.getElementsByTagNameNS(S, 'definedName')}
    for path, payload in parts.items():
        if path.startswith(('xl/slicers/', 'xl/timelines/')) and path.endswith('.xml'):
            used_names.update(n.getAttribute('name').casefold() for n in _children(minidom.parseString(payload).documentElement))
    cache_ids = [int(n.getAttribute('cacheId')) for n in wb.getElementsByTagNameNS(S, 'pivotCache')]
    internal_ids = []
    for path, payload in parts.items():
        if path.startswith('xl/pivotCache/') and path.endswith('.xml'):
            internal_ids.extend(int(n.getAttribute('pivotCacheId')) for n in minidom.parseString(payload).getElementsByTagNameNS(X14, 'pivotCacheDefinition') if n.hasAttribute('pivotCacheId'))
    views = {}
    for timeline, specs in [(False, slicers), (True, timelines)]:
        ns = X15 if timeline else X14
        kind = 'timeline' if timeline else 'slicer'
        for spec in specs:
            if spec.name.casefold() in used_names: raise ModelValidationError('Control name already exists: '+spec.name)
            used_names.add(spec.name.casefold())
            if spec.pivot_table not in pivots: raise ModelValidationError('Unknown control PivotTable: '+spec.pivot_table)
            pivot_sheet, tabid, report_cache = pivots[spec.pivot_table]
            sheet_name = spec.sheet or pivot_sheet
            # Worksheet names are case-insensitive in Excel (and in report_package).
            sheet_name = next((n for n in sheets if n.casefold() == sheet_name.casefold()), sheet_name)
            if sheet_name not in sheets: raise ModelValidationError('Control worksheet does not exist: '+sheet_name)
            table, column = spec.field
            hierarchy = _mdx(table)+'.'+_mdx(column)
            report_root = package.root(report_cache)
            if not any(n.getAttribute('uniqueName') == hierarchy for n in report_root.getElementsByTagNameNS(S, 'cacheHierarchy')):
                raise ModelValidationError('Control field is absent from the report model')
            cache_name = kind.title()+'_'+re.sub(r'[^\w]', '_', spec.name)
            while cache_name.casefold() in used_names: cache_name += '_'
            used_names.add(cache_name.casefold())
            named = next(iter(_children(wb, S, 'definedNames')), None)
            if named is None:
                named = _node(wb, S, 'definedNames')
                later = next((n for n in _children(wb) if n.localName in ('calcPr','oleSize','customWorkbookViews','pivotCaches','smartTagPr','smartTagTypes','webPublishing','fileRecoveryPr','webPublishObjects','extLst')), None)
                if later: wb.insertBefore(named, later)
            definition = _node(named, S, 'definedName', name=cache_name)
            definition.appendChild(wb.ownerDocument.createTextNode('#N/A'))
            cache_id = max(cache_ids+[0])+1; cache_ids.append(cache_id)
            internal_id = max(internal_ids+[100000])+1; internal_ids.append(internal_id)
            _control_cache(package, report_cache, hierarchy, timeline, cache_id, internal_id)
            cachepath = package.alloc('xl/'+kind+'Caches/'+kind+'Cache')
            cache = package.root(cachepath, ns, kind+'CacheDefinition')
            cache.setAttribute('name', cache_name); cache.setAttribute('sourceName', hierarchy)
            _control_uid(cache)
            tables_node = _node(cache, ns, 'pivotTables')
            for target in dict.fromkeys((spec.pivot_table,)+tuple(getattr(spec, 'connected_pivots', ()))):
                if target not in pivots: raise ModelValidationError('Unknown connected PivotTable: '+target)
                _, target_id, target_cache = pivots[target]
                source_id = _children(package.root(report_cache), S, 'cacheSource')[0].getAttribute('connectionId')
                target_id_conn = _children(package.root(target_cache), S, 'cacheSource')[0].getAttribute('connectionId')
                if source_id != target_id_conn: raise ModelValidationError('Connected controls require the same model connection')
                _node(tables_node, ns, 'pivotTable', tabId=target_id, name=target)
            if timeline:
                import pandas as pd
                values = pd.Series(list(source_member_values(model, table, column))).dropna()
                if values.empty: raise ModelValidationError('Timeline needs at least one nonblank date')
                try:
                    start = datetime(int(values.min().year), 1, 1)
                    end = datetime(int(values.max().year)+1, 1, 1)
                except (AttributeError, ValueError, OverflowError) as exc:
                    raise ModelValidationError('Timeline requires supported date values before year 9999') from exc
                state = _node(cache, ns, 'state', minimalRefreshVersion=6, lastRefreshVersion=6, pivotCacheId=internal_id, filterType='unknown')
                _node(state, ns, 'bounds', startDate=start.isoformat(), endDate=end.isoformat())
            else:
                olap = _node(_node(cache, ns, 'data'), ns, 'olap', pivotCacheId=internal_id)
                levels = _node(olap, ns, 'levels', count=2)
                _node(levels, ns, 'level', uniqueName=hierarchy+'.[(All)]', sourceCaption='(All)', count=0)
                # These are member identities, not calculated report results.
                # An empty level leaves a real slicer with no clickable buttons
                # until Excel explicitly refreshes its dedicated control cache.
                members = {}
                for value in source_member_values(model, table, column):
                    key = member_unique_name(table, column, value)
                    members.setdefault(key, _caption(value))
                level = _node(levels, ns, 'level', uniqueName=hierarchy+'.'+_mdx(column), sourceCaption=column, count=len(members))
                if members:
                    member_range = _node(_node(level, ns, 'ranges'), ns, 'range', startItem=0)
                    for key, caption in members.items():
                        _node(member_range, ns, 'i', n=key, c=caption)
                _node(_node(olap, ns, 'selections', count=1), ns, 'selection', n=hierarchy+'.[All]')
            relation = ('http://schemas.microsoft.com/office/2011/relationships/timelineCache' if timeline else 'http://schemas.microsoft.com/office/2007/relationships/slicerCache')
            rid = package.relation('xl/workbook.xml', relation, cachepath)
            uri, group, entry = (('{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}', 'x15:timelineCacheRefs', 'x15:timelineCacheRef') if timeline
                                 else ('{BBE1A952-AA13-448e-AADC-164F8A28A991}', 'x14:slicerCaches', 'x14:slicerCache'))
            _node(_extension(wb, uri, ns, group), ns, entry, rid=rid)
            package.content_type(cachepath, 'application/vnd.ms-excel.'+kind+'Cache+xml')
            sheetpath = sheets[sheet_name][0]
            key = (sheetpath, timeline)
            if key not in views:
                relation = ('http://schemas.microsoft.com/office/2011/relationships/timeline' if timeline else 'http://schemas.microsoft.com/office/2007/relationships/slicer')
                old = [n for n in _children(package.root(_relpath(sheetpath), P, 'Relationships')) if n.getAttribute('Type') == relation]
                if len(old) > 1: raise ModelValidationError('Worksheet has multiple control view parts')
                if old:
                    viewpath = package.target(sheetpath, old[0].getAttribute('Id'))
                    views[key] = package.root(viewpath)
            if key not in views:
                viewpath = package.alloc('xl/'+kind+'s/'+kind)
                viewroot = package.root(viewpath, ns, kind+'s')
                views[key] = viewroot
                relation = ('http://schemas.microsoft.com/office/2011/relationships/timeline' if timeline else 'http://schemas.microsoft.com/office/2007/relationships/slicer')
                rid = package.relation(sheetpath, relation, viewpath)
                uri, group, entry = (('{7E03D99C-DC04-49d9-9315-930204A7B6E9}', 'x15:timelineRefs', 'x15:timelineRef') if timeline
                                     else ('{A8765BA9-456A-4dab-B4F3-ACF838C121DE}', 'x14:slicerList', 'x14:slicer'))
                _node(_extension(package.root(sheetpath), uri, ns, group), ns, entry, rid=rid)
                package.content_type(viewpath, 'application/vnd.ms-excel.'+kind+'+xml')
            attrs = dict(name=spec.name, cache=cache_name, caption=spec.caption or column)
            if timeline:
                level = {'years':0,'quarters':1,'months':2,'days':3}[spec.level]
                attrs.update(level=level, selectionLevel=level, scrollPosition=start.isoformat())
            else: attrs.update(level=1, rowHeight=257175, columnCount=spec.columns)
            _control_uid(_node(views[key], ns, kind, **attrs))
            _drawing(package, sheetpath, spec, timeline)
    result = package.finish()
    parts.clear(); parts.update(result)
