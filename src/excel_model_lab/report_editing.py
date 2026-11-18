"""Inspect and surgically edit existing OLAP reports without rebuilding caches.

Removing a report removes its live object, not its worksheet or displayed cells.
Unknown XML is retained; edits that require reinterpreting it are rejected.
"""
from dataclasses import dataclass, replace
from pathlib import Path
from tempfile import mkstemp
from xml.dom import minidom
from zipfile import ZipFile
import os
import posixpath
import re

from .errors import ModelValidationError
from .reports import PivotTableSpec, PivotChartSpec, _pivot_name, _text, _xml_text
from .report_filters import PivotMemberFilter, PivotSort

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
X = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'


@dataclass(frozen=True)
class ExistingReport:
    kind: str
    name: str
    sheet: str
    part: str
    spec: PivotTableSpec | PivotChartSpec | None
    unsupported: tuple[str, ...] = ()


@dataclass(frozen=True)
class ExistingControl:
    """Saved control bindings; presence does not certify Excel acceptance."""
    kind: str
    name: str
    sheet: str
    part: str
    cache_part: str
    cache_name: str
    caption: str
    source_name: str
    connected_pivots: tuple[tuple[str, str], ...]


def _nodes(node, namespace, tag):
    return list(node.getElementsByTagNameNS(namespace, tag))


def _value(node):
    return ''.join(c.data for c in node.childNodes if c.nodeType in (c.TEXT_NODE, c.CDATA_SECTION_NODE))


def _children(node, namespace, tag):
    return [n for n in node.childNodes if n.nodeType == n.ELEMENT_NODE
            and n.namespaceURI == namespace and n.localName == tag]


def _set_text(node, text):
    for child in list(node.childNodes):
        node.removeChild(child)
    node.appendChild(node.ownerDocument.createTextNode(text))


def _rels_path(part):
    return posixpath.dirname(part) + '/_rels/' + posixpath.basename(part) + '.rels'


def _target(owner, node):
    if node.getAttribute('TargetMode') == 'External':
        raise ModelValidationError('External report relationships are unsupported')
    path = posixpath.normpath(posixpath.join(posixpath.dirname(owner), node.getAttribute('Target'))).lstrip('/')
    if path.startswith('../'):
        raise ModelValidationError('Report relationship escapes the package')
    return path


class _Package:
    def __init__(self, source):
        with ZipFile(source) as archive:
            self.info = archive.infolist()
            self.comment = archive.comment
            if len({i.filename for i in self.info}) != len(self.info):
                raise ModelValidationError('Duplicate workbook package members')
            self.parts = {i.filename: archive.read(i) for i in self.info}
        self.docs = {}
        self.changed = set()
        self.deleted = set()

    def doc(self, path):
        if path not in self.docs:
            self.docs[path] = minidom.parseString(self.parts[path])
        return self.docs[path]

    def relationships(self, part):
        path = _rels_path(part)
        if path not in self.parts:
            return []
        return [n for n in self.doc(path).documentElement.childNodes if n.nodeType == n.ELEMENT_NODE]

    def related(self, part, rid):
        matches = [n for n in self.relationships(part) if n.getAttribute('Id') == rid]
        if len(matches) != 1:
            raise ModelValidationError('Report relationship is missing or ambiguous')
        result = _target(part, matches[0])
        if result not in self.parts:
            raise ModelValidationError('Report relationship target is missing')
        return result

    def inventory(self):
        result = []
        self.bindings = {}
        workbook = self.doc('xl/workbook.xml')
        for sheet in _nodes(workbook, S, 'sheet'):
            sheet_name = sheet.getAttribute('name')
            spath = self.related('xl/workbook.xml', sheet.getAttributeNS(R, 'id'))
            explicit = {n.getAttributeNS(R, 'id'): n for n in _nodes(self.doc(spath), S, 'pivotTablePart')}
            # Excel may omit pivotTableParts in sheet XML after saving, while
            # retaining its authoritative worksheet relationship.
            pivot_links = [(n.getAttribute('Id'), explicit.get(n.getAttribute('Id'), n))
                           for n in self.relationships(spath) if n.getAttribute('Type').endswith('/pivotTable')]
            for rid, part_node in pivot_links:
                ppath = self.related(spath, rid)
                pivot = self.doc(ppath).documentElement
                name = pivot.getAttribute('name')
                caches = [n for n in self.relationships(ppath) if n.getAttribute('Type').endswith('/pivotCacheDefinition')]
                if len(caches) != 1:
                    raise ModelValidationError('PivotTable needs one cache relationship')
                cpath = _target(ppath, caches[0])
                spec, unsupported = self.pivot_spec(pivot, self.doc(cpath), name, sheet_name)
                record = ExistingReport('pivot_table', name, sheet_name, ppath, spec, unsupported)
                result.append(record)
                self.bindings[ppath] = (spath, part_node, cpath)
            for drawing in _nodes(self.doc(spath), S, 'drawing'):
                dpath = self.related(spath, drawing.getAttributeNS(R, 'id'))
                for chart_node in _nodes(self.doc(dpath), C, 'chart'):
                    cpath = self.related(dpath, chart_node.getAttributeNS(R, 'id'))
                    chart = self.doc(cpath)
                    sources = _nodes(chart, C, 'pivotSource')
                    if not sources:
                        continue
                    anchor = chart_node
                    while anchor.parentNode is not self.doc(dpath).documentElement:
                        anchor = anchor.parentNode
                        if anchor is None:
                            raise ModelValidationError('Chart has no drawing anchor')
                    props = _nodes(anchor, X, 'cNvPr')
                    if len(props) != 1:
                        raise ModelValidationError('Chart name is ambiguous')
                    name = props[0].getAttribute('name')
                    spec, unsupported = self.chart_spec(chart, anchor, name, sheet_name)
                    record = ExistingReport('pivot_chart', name, sheet_name, cpath, spec, unsupported)
                    result.append(record)
                    self.bindings[cpath] = (dpath, anchor, chart_node, props[0])
        names = [r.name.casefold() for r in result]
        if len(names) != len(set(names)):
            raise ModelValidationError('Report names are ambiguous; use unique names before editing')
        return tuple(result)

    def pivot_spec(self, pivot, cache, name, sheet):
        try:
            sources = _nodes(cache, S, 'cacheSource')
            if len(sources) != 1 or sources[0].getAttribute('type') != 'external':
                return None, ('Non-OLAP PivotTable',)
            fields = _nodes(cache, S, 'cacheField')
            def field(index, measure=False):
                index = int(index)
                if index < 0 or index >= len(fields):
                    raise ValueError('Invalid cache field index')
                text = fields[index].getAttribute('name')
                components = re.findall(r'\[((?:[^\]]|\]\])*)\]', text)
                if '.'.join('['+c+']' for c in components) != text:
                    raise ValueError('Unsupported MDX field')
                components = [c.replace(']]', ']') for c in components]
                if measure and len(components) == 2 and components[0] == 'Measures':
                    return components[1]
                if not measure and len(components) == 3 and components[1] == components[2]:
                    return tuple(components[:2])
                raise ValueError('Unsupported hierarchy field')
            def axis(tag, item, attribute):
                parents = _nodes(pivot, S, tag)
                return tuple(field(n.getAttribute(attribute)) for p in parents for n in _nodes(p, S, item)
                             if n.getAttribute(attribute) != '-2')
            rows = axis('rowFields', 'field', 'x')
            columns = axis('colFields', 'field', 'x')
            filters = axis('pageFields', 'pageField', 'fld')
            measures = tuple(field(n.getAttribute('fld'), True) for n in _nodes(pivot, S, 'dataField'))
            pfields = _nodes(pivot, S, 'pivotField')
            hierarchies = _nodes(cache, S, 'cacheHierarchy')
            selections, sorts = [], []
            def scalar(index, value):
                from datetime import datetime
                hierarchy = int(fields[index].getAttribute('hierarchy'))
                dtype = hierarchies[hierarchy].getAttribute('memberValueDatatype')
                if dtype in ('2', '3', '16', '17', '18', '19', '20', '21'):
                    return int(value)
                if dtype in ('129', '130', ''):
                    return value
                if dtype == '11':
                    if value.lower() not in ('true','false','0','1'):
                        raise ValueError('Invalid boolean member')
                    return value.lower() in ('true','1')
                if dtype == '7':
                    return datetime.fromisoformat(value)
                raise ValueError('Unsupported selected member datatype')
            def shared_scalar(index, node):
                from datetime import datetime
                if node.localName == 'm':
                    return None
                if node.localName == 'b':
                    value = node.getAttribute('v')
                    if value not in ('0','1','true','false'):
                        raise ValueError('Invalid boolean cache member')
                    return value in ('1','true')
                if node.localName == 'd':
                    return datetime.fromisoformat(node.getAttribute('v'))
                return scalar(index,node.getAttribute('v'))
            for i, pfield in enumerate(pfields):
                if pfield.getAttribute('axis') not in ('axisRow', 'axisCol', 'axisPage'):
                    continue
                table, column = field(i)
                order = pfield.getAttribute('sortType')
                if order in ('ascending', 'descending'):
                    if _nodes(pfield, S, 'autoSortScope'):
                        raise ValueError('Measure-based sorting cannot be represented as label sorting')
                    sorts.append(PivotSort(table, column, order))
                shared = _children(fields[i], S, 'sharedItems')
                selected = [n for n in _nodes(pfield, S, 'item') if n.getAttribute('s') in ('1', 'true')]
                if selected:
                    if not shared:
                        raise ValueError('Selected members have no cache inventory')
                    values = [n for n in shared[0].childNodes if n.nodeType == n.ELEMENT_NODE]
                    selections.append(PivotMemberFilter(table, column, tuple(
                        shared_scalar(i, values[int(n.getAttribute('x'))]) for n in selected)))
            for page in _nodes(pivot, S, 'pageField'):
                name_value = page.getAttribute('name')
                if '.&[' in name_value or name_value.endswith('.&'):
                    index = int(page.getAttribute('fld'))
                    table, column = field(index)
                    prefix = '['+table.replace(']', ']]')+'].['+column.replace(']', ']]')+'].&'
                    member = name_value[len(prefix):] if name_value.startswith(prefix) else ''
                    if name_value == prefix:
                        selections.append(PivotMemberFilter(table,column,(None,)))
                        continue
                    if not re.fullmatch(r'\[(?:[^\]]|\]\])*\]', member):
                        raise ValueError('Unsupported page member selection')
                    selections.append(PivotMemberFilter(table, column, (scalar(index, member[1:-1].replace(']]', ']')),)))
            location = _nodes(pivot, S, 'location')[0].getAttribute('ref').split(':')[0]
            compact = pivot.getAttribute('compact') not in ('0', 'false')
            layout = 'compact' if compact else 'outline' if any(n.getAttribute('outline') not in ('0', 'false') for n in pfields) else 'tabular'
            spec = PivotTableSpec(name, sheet, location, rows, columns, filters, measures, layout,
                                  pivot.getAttribute('rowGrandTotals') not in ('0', 'false'),
                                  pivot.getAttribute('colGrandTotals') not in ('0', 'false'),
                                  tuple(selections), tuple(sorts))
            return spec, ()
        except (ValueError, IndexError, TypeError) as exc:
            return None, (str(exc),)

    def chart_spec(self, chart, anchor, name, sheet):
        try:
            source = _value(_nodes(_nodes(chart, C, 'pivotSource')[0], C, 'name')[0])
            pivot_name = source.rsplit('!', 1)[1]
            plot = _nodes(chart, C, 'plotArea')[0]
            kinds = [n for n in plot.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName.endswith('Chart')]
            if len(kinds) != 1 or kinds[0].localName not in ('barChart', 'lineChart'):
                return None, ('Unsupported or combined chart type',)
            kind = 'line' if kinds[0].localName == 'lineChart' else 'bar' if _nodes(kinds[0], C, 'barDir')[0].getAttribute('val') == 'bar' else 'column'
            origin = _nodes(anchor, X, 'from')[0]
            col = int(_value(_nodes(origin, X, 'col')[0])) + 1
            row = int(_value(_nodes(origin, X, 'row')[0])) + 1
            letters = ''
            while col:
                col, n = divmod(col-1, 26)
                letters = chr(65+n)+letters
            ext = [n for n in anchor.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName == 'ext']
            if ext:
                width, height = (round(int(ext[0].getAttribute(k))/9525) for k in ('cx', 'cy'))
            else:
                sizes = _nodes(anchor, A, 'ext')
                if not sizes:
                    return None, ('Chart anchor has no explicit extent',)
                width, height = (round(int(sizes[0].getAttribute(k))/9525) for k in ('cx', 'cy'))
            titles = _children(_nodes(chart, C, 'chart')[0], C, 'title')
            title = ''.join(_value(n) for n in _nodes(titles[0], A, 't')) if titles else ''
            chart_root = _nodes(chart, C, 'chart')[0]
            legend = _children(chart_root, C, 'legend')
            positions = {'r':'right', 'l':'left', 't':'top', 'b':'bottom', 'tr':'top_right'}
            position_nodes = _nodes(legend[0], C, 'legendPos') if legend else []
            position = positions.get(position_nodes[0].getAttribute('val'), 'right') if position_nodes else 'right'
            def axis_title(tag):
                axes = _children(plot, C, tag)
                titles = _children(axes[0], C, 'title') if axes else []
                return ''.join(_value(n) for n in _nodes(titles[0], A, 't')) if titles else ''
            def color(parent):
                props = _children(parent, C, 'spPr')
                fills = _children(props[0], A, 'solidFill') if props else []
                colors = _children(fills[0], A, 'srgbClr') if fills else []
                return colors[0].getAttribute('val') if colors else None
            styles = _children(chart.documentElement, C, 'style')
            values = _children(plot, C, 'valAx')
            formats = _children(values[0], C, 'numFmt') if values else []
            labels = _children(kinds[0], C, 'dLbls')
            flags = _children(labels[0], C, 'showVal') if labels else []
            return PivotChartSpec(name, pivot_name, sheet, letters+str(row), kind, width, height, title,
                style=int(styles[0].getAttribute('val')) if styles else 2,
                show_legend=bool(legend), legend_position=position,
                category_axis_title=axis_title('catAx'), value_axis_title=axis_title('valAx'),
                number_format=formats[0].getAttribute('formatCode') if formats else 'General',
                show_data_labels=bool(flags and flags[0].getAttribute('val') in ('1','true')),
                background_color=color(chart.documentElement), plot_background_color=color(plot)), ()
        except (ValueError, IndexError, TypeError) as exc:
            return None, (str(exc),)

    def unlink(self, owner, rid):
        path = _rels_path(owner)
        matches = [n for n in self.relationships(owner) if n.getAttribute('Id') == rid]
        if len(matches) != 1:
            raise ModelValidationError('Cannot unlink ambiguous report relationship')
        matches[0].parentNode.removeChild(matches[0])
        self.changed.add(path)

    def write(self, destination):
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        fd, name = mkstemp(prefix='.report-edit-', suffix='.xlsx', dir=destination.parent)
        os.close(fd)
        try:
            with ZipFile(name, 'w') as archive:
                archive.comment = self.comment
                for info in self.info:
                    if info.filename not in self.deleted:
                        archive.writestr(info, self.docs[info.filename].toxml(encoding='utf-8') if info.filename in self.changed else self.parts[info.filename])
            Path(name).replace(destination)
        finally:
            Path(name).unlink(missing_ok=True)
        return destination


def inspect_reports(source):
    """Return saved report definitions, with unsupported layouts identified."""
    return _Package(source).inventory()


def inspect_controls(source):
    """Inspect linked slicer/timeline view and cache metadata without editing."""
    return _inspect_controls(_Package(source))


def _inspect_controls(package):
    caches = {}
    for relation in package.relationships('xl/workbook.xml'):
        kind = relation.getAttribute('Type').rsplit('/',1)[-1]
        if kind not in ('slicerCache','timelineCache'):
            continue
        path = _target('xl/workbook.xml',relation)
        root = package.doc(path).documentElement
        key = (kind,root.getAttribute('name'))
        if key in caches:
            raise ModelValidationError('Ambiguous control cache name')
        caches[key] = (path,root)
    result = []
    for sheet in _nodes(package.doc('xl/workbook.xml'),S,'sheet'):
        spath=package.related('xl/workbook.xml',sheet.getAttributeNS(R,'id'))
        for relation in package.relationships(spath):
            kind=relation.getAttribute('Type').rsplit('/',1)[-1]
            if kind not in ('slicer','timeline'):
                continue
            path=_target(spath,relation)
            root=package.doc(path).documentElement
            for control in root.childNodes:
                if control.nodeType!=control.ELEMENT_NODE or control.localName!=kind:
                    continue
                cache_name=control.getAttribute('cache')
                cached=caches.get((kind+'Cache',cache_name))
                if cached is None:
                    raise ModelValidationError('Control references a missing workbook cache')
                cpath,cache=cached
                pivots=tuple((n.getAttribute('tabId'),n.getAttribute('name'))
                             for n in cache.getElementsByTagName('*') if n.localName=='pivotTable')
                result.append(ExistingControl(kind,control.getAttribute('name'),sheet.getAttribute('name'),
                    path,cpath,cache_name,control.getAttribute('caption'),cache.getAttribute('sourceName'),pivots))
    return tuple(result)


def _control_binding(package, record):
    views=package.doc(record.part)
    nodes=[n for n in views.getElementsByTagName('*') if n.localName==record.kind and n.getAttribute('name')==record.name]
    if len(nodes)!=1:
        raise ModelValidationError('Ambiguous control view')
    sheet=next(n for n in _nodes(package.doc('xl/workbook.xml'),S,'sheet') if n.getAttribute('name')==record.sheet)
    spath=package.related('xl/workbook.xml',sheet.getAttributeNS(R,'id'))
    matches=[]
    for relation in package.relationships(spath):
        if not relation.getAttribute('Type').endswith('/drawing'):
            continue
        dpath=_target(spath,relation)
        root=package.doc(dpath).documentElement
        for anchor in root.childNodes:
            if anchor.nodeType!=anchor.ELEMENT_NODE:
                continue
            props=[n for n in _nodes(anchor,X,'cNvPr') if n.getAttribute('name')==record.name]
            controls=[n for n in anchor.getElementsByTagName('*') if n.localName in ('slicer','timeSlicer','timeslicer') and n.getAttribute('name')==record.name]
            if props and controls:
                matches.append((dpath,anchor,props,controls))
    if len(matches)!=1:
        raise ModelValidationError('Control drawing anchor is missing or ambiguous')
    return nodes[0],spath,matches[0]


def _edit_control(package, record, changes, controls, reports):
    allowed={'name','caption','cell','width','height','columns','level','connected_pivots'}
    if set(changes)-allowed:
        raise ModelValidationError('Unsupported control edit')
    node,spath,(dpath,anchor,props,draw_controls)=_control_binding(package,record)
    if 'name' in changes:
        _text(changes['name'],'Control name')
        if len(changes['name'])>255:
            raise ModelValidationError('Control name is too long')
        node.setAttribute('name',changes['name'])
        for target in props+draw_controls:
            target.setAttribute('name',changes['name'])
        package.changed.update((record.part,dpath))
    if 'caption' in changes:
        _xml_text(changes['caption'],'Control caption')
        node.setAttribute('caption',changes['caption']);package.changed.add(record.part)
    if 'columns' in changes:
        if record.kind!='slicer' or type(changes['columns']) is not int or changes['columns']<1:
            raise ModelValidationError('Slicer columns must be a positive integer')
        node.setAttribute('columnCount',str(changes['columns']));package.changed.add(record.part)
    if 'level' in changes:
        levels={'years':0,'quarters':1,'months':2,'days':3}
        if record.kind!='timeline' or changes['level'] not in levels:
            raise ModelValidationError('Unsupported timeline level')
        node.setAttribute('level',str(levels[changes['level']]))
        # selectionLevel describes existing filter state; preserve it.
        package.changed.add(record.part)
    if set(changes)&{'cell','width','height'}:
        origins=_children(anchor,X,'from')
        extents=_children(anchor,X,'ext')
        if anchor.localName=='twoCellAnchor' and len(origins)==1:
            sizes=[n for n in _nodes(anchor,A,'ext') if n.hasAttribute('cx') and n.hasAttribute('cy')]
            frame=next((n for n in sizes if int(n.getAttribute('cx'))>0 and int(n.getAttribute('cy'))>0),None)
            if frame is None and not {'width','height'}<=set(changes):
                raise ModelValidationError('A two-cell control without stored extent requires explicit width and height')
            for key in ('width','height'):
                if key in changes and (type(changes[key]) is not int or changes[key]<=0):
                    raise ModelValidationError('Control dimensions must be positive integers')
            width=changes.get('width')
            height=changes.get('height')
            cx=str(width*9525) if width is not None else frame.getAttribute('cx')
            cy=str(height*9525) if height is not None else frame.getAttribute('cy')
            old_anchor=anchor
            qualified=anchor.prefix+':oneCellAnchor' if anchor.prefix else 'oneCellAnchor'
            anchor=anchor.ownerDocument.createElementNS(X,qualified)
            for attr in old_anchor.attributes.values():
                if attr.name!='editAs':anchor.setAttributeNS(attr.namespaceURI,attr.name,attr.value)
            for child in list(old_anchor.childNodes):
                if child.nodeType==child.ELEMENT_NODE and child.namespaceURI==X and child.localName=='to':
                    continue
                anchor.appendChild(child)
            ext=anchor.ownerDocument.createElementNS(X,anchor.prefix+':ext' if anchor.prefix else 'ext')
            ext.setAttribute('cx',cx);ext.setAttribute('cy',cy)
            anchor.insertBefore(ext,origins[0].nextSibling)
            old_anchor.parentNode.replaceChild(anchor,old_anchor)
            extents=[ext]
        if len(origins)!=1 or len(extents)!=1:
            raise ModelValidationError('Only one-cell control anchors support geometry edits')
        if 'cell' in changes:
            from .reports import _cell
            _cell(changes['cell'])
            match=re.fullmatch(r'([A-Za-z]+)([0-9]+)',changes['cell'])
            column=0
            for char in match[1].upper():column=column*26+ord(char)-64
            for tag,value in [('col',column-1),('row',int(match[2])-1),('colOff',0),('rowOff',0)]:
                values=_children(origins[0],X,tag)
                if len(values)!=1:raise ModelValidationError('Malformed control anchor')
                _set_text(values[0],str(value))
        for key,attr in [('width','cx'),('height','cy')]:
            if key in changes:
                if type(changes[key]) is not int or changes[key]<=0:
                    raise ModelValidationError('Control dimensions must be positive integers')
                extents[0].setAttribute(attr,str(changes[key]*9525))
        package.changed.add(dpath)
    if 'connected_pivots' in changes:
        targets=changes['connected_pivots']
        if isinstance(targets,(str,bytes)):
            raise ModelValidationError('connected_pivots must be a sequence')
        try:targets=tuple(targets)
        except TypeError as exc:raise ModelValidationError('connected_pivots must be a sequence') from exc
        if not targets or any(not isinstance(name,str) for name in targets) or len(set(targets))!=len(targets):
            raise ModelValidationError('Control connections need distinct PivotTable names')
        if sum(c.cache_part==record.cache_part for c in controls)>1:
            raise ModelValidationError('Cannot edit connections of a cache shared by multiple controls')
        pivots={r.name:r for r in reports if r.kind=='pivot_table'}
        if any(name not in pivots for name in targets):
            raise ModelValidationError('Unknown connected PivotTable')
        existing=[name for _,name in record.connected_pivots]
        def connection(name):
            if name not in pivots:raise ModelValidationError('Existing control has an unresolved PivotTable')
            cache=package.bindings[pivots[name].part][2]
            sources=_nodes(package.doc(cache),S,'cacheSource')
            if len(sources)!=1:raise ModelValidationError('Invalid PivotCache source')
            return sources[0].getAttribute('connectionId')
        connections={connection(name) for name in existing+list(targets)}
        if len(connections)!=1:
            raise ModelValidationError('Connected controls require the same model connection')
        cache=package.doc(record.cache_part).documentElement
        groups=[n for n in cache.childNodes if n.nodeType==n.ELEMENT_NODE and n.localName=='pivotTables']
        if len(groups)!=1:raise ModelValidationError('Ambiguous control connections')
        group=groups[0]
        for old in list(group.childNodes):group.removeChild(old)
        sheet_ids={n.getAttribute('name'):n.getAttribute('sheetId') for n in _nodes(package.doc('xl/workbook.xml'),S,'sheet')}
        for name in targets:
            entry=group.ownerDocument.createElementNS(group.namespaceURI,group.prefix+':pivotTable' if group.prefix else 'pivotTable')
            entry.setAttribute('tabId',sheet_ids[pivots[name].sheet]);entry.setAttribute('name',name)
            group.appendChild(entry)
        package.changed.add(record.cache_part)


def _detach_part(package,owner,target):
    """Detach known owning relationships and their XML extension references."""
    matches=[n for n in package.relationships(owner) if n.getAttribute('TargetMode')!='External' and _target(owner,n)==target]
    for relation in matches:
        rid=relation.getAttribute('Id')
        for node in list(package.doc(owner).getElementsByTagName('*')):
            if node.getAttributeNS(R,'id')==rid:
                parent=node.parentNode
                parent.removeChild(node);package.changed.add(owner)
                while parent.nodeType==parent.ELEMENT_NODE and parent.localName in (
                    'slicerList','timelineRefs','slicerCaches','timelineCacheRefs','ext','extLst'):
                    if any(n.nodeType==n.ELEMENT_NODE for n in parent.childNodes):break
                    grand=parent.parentNode
                    grand.removeChild(parent);parent=grand
        package.unlink(owner,rid)


def _remove_control(package,record,remaining):
    node,spath,(dpath,anchor,props,draw_controls)=_control_binding(package,record)
    node.parentNode.removeChild(node);package.changed.add(record.part)
    anchor.parentNode.removeChild(anchor);package.changed.add(dpath)
    if not any(c.part==record.part for c in remaining):
        _detach_part(package,spath,record.part)
        package.deleted.add(record.part)
    if not any(c.cache_part==record.cache_part for c in remaining):
        _detach_part(package,'xl/workbook.xml',record.cache_part)
        package.deleted.add(record.cache_part)


def _chart_format(package, record, changes):
    """Change requested formatting nodes, retaining cached series verbatim."""
    from .report_charts import chart_xml, validate_chart_format
    if record.spec is None:
        raise ModelValidationError('Chart formatting requires a supported chart type')
    try:
        spec = replace(record.spec, **changes)
        validate_chart_format(spec)
        template = minidom.parseString(chart_xml(spec, spec.pivot_table, record.sheet, 'report.xlsx'))
    except (ValueError, TypeError) as exc:
        raise ModelValidationError(str(exc)) from exc
    doc = package.doc(record.part)
    root, generated = doc.documentElement, template.documentElement
    chart, tchart = _nodes(doc, C, 'chart')[0], _nodes(template, C, 'chart')[0]
    plot, tplot = _children(chart, C, 'plotArea')[0], _children(tchart, C, 'plotArea')[0]
    def copy_child(parent, source, tag, order, namespace=C):
        old = _children(parent, namespace, tag)
        new = _children(source, namespace, tag)
        if len(old) > 1 or len(new) > 1:
            raise ModelValidationError('Ambiguous chart formatting node: '+tag)
        if old:
            parent.removeChild(old[0])
        if new:
            node = doc.importNode(new[0], True)
            node.setAttribute('xmlns:c', C)
            node.setAttribute('xmlns:a', A)
            rank = order.index(tag)
            anchor = next((n for n in parent.childNodes if n.nodeType == n.ELEMENT_NODE
                           and n.localName in order and order.index(n.localName) > rank), None)
            parent.insertBefore(node, anchor)
    root_order = ['date1904','lang','roundedCorners','AlternateContent','style','clrMapOvr','pivotSource','protection','chart','spPr','txPr','externalData','printSettings','userShapes','extLst']
    chart_order = ['title','autoTitleDeleted','pivotFmts','view3D','floor','sideWall','backWall','plotArea','legend','plotVisOnly','dispBlanksAs','showDLblsOverMax','extLst']
    plot_order = ['layout','barChart','lineChart','catAx','dateAx','serAx','valAx','dTable','spPr','extLst']
    if 'style' in changes:
        if any(n.localName == 'AlternateContent' and _nodes(n, C, 'style') for n in root.childNodes if n.nodeType == n.ELEMENT_NODE):
            raise ModelValidationError('Alternate-content chart styles are not supported for editing')
        copy_child(root, generated, 'style', root_order)
    if 'show_legend' in changes or 'legend_position' in changes:
        # A position-only change retains native legend formatting.
        existing = _children(chart, C, 'legend')
        target = _children(tchart, C, 'legend')
        if existing and target:
            copy_child(existing[0], target[0], 'legendPos', ['legendPos','legendEntry','layout','overlay','spPr','txPr','extLst'])
        else:
            copy_child(chart, tchart, 'legend', chart_order)
    axis_order = ['axId','scaling','delete','axPos','majorGridlines','minorGridlines','title','numFmt','majorTickMark','minorTickMark','tickLblPos','spPr','txPr','crossAx','crosses','crossesAt','auto','lblAlgn','lblOffset','tickLblSkip','tickMarkSkip','noMultiLvlLbl','crossBetween','majorUnit','minorUnit','dispUnits','extLst']
    for key, axis, tag in [('category_axis_title','catAx','title'),('value_axis_title','valAx','title'),('number_format','valAx','numFmt')]:
        if key in changes:
            axes, targets = _children(plot,C,axis), _children(tplot,C,axis)
            if len(axes) != 1 or len(targets) != 1:
                raise ModelValidationError('Unsupported chart axis configuration')
            copy_child(axes[0], targets[0], tag, axis_order)
    for key, parent, target in [('background_color',root,generated),('plot_background_color',plot,tplot)]:
        if key in changes:
            props = _children(parent,C,'spPr')
            targets = _children(target,C,'spPr')
            if props:
                # Replace only the fill, preserving borders/effects/geometry.
                fill_tags = {'solidFill','noFill','gradFill','blipFill','pattFill','grpFill'}
                for node in list(props[0].childNodes):
                    if node.nodeType == node.ELEMENT_NODE and node.namespaceURI == A and node.localName in fill_tags:
                        props[0].removeChild(node)
                if targets:
                    fills = _children(targets[0],A,'solidFill')
                    if fills:
                        node = doc.importNode(fills[0], True);node.setAttribute('xmlns:a', A)
                        anchor = next((n for n in props[0].childNodes if n.nodeType==n.ELEMENT_NODE and n.localName in ('ln','effectLst','effectDag','scene3d','sp3d','extLst')),None)
                        props[0].insertBefore(node,anchor)
            elif targets:
                copy_child(parent,target,'spPr',root_order if parent is root else plot_order)
    kinds = [n for n in plot.childNodes if n.nodeType==n.ELEMENT_NODE and n.localName.endswith('Chart')]
    tkinds = [n for n in tplot.childNodes if n.nodeType==n.ELEMENT_NODE and n.localName.endswith('Chart')]
    if len(kinds)!=1 or len(tkinds)!=1:
        raise ModelValidationError('Combined charts cannot be edited')
    kind, target = kinds[0], tkinds[0]
    if 'chart_type' in changes:
        if kind.localName == target.localName:
            if kind.localName == 'barChart':
                copy_child(kind,target,'barDir',['barDir','grouping','varyColors','ser','dLbls','gapWidth','overlap','serLines','axId','extLst'])
        else:
            # Native series caches are preserved. Unknown type-specific series
            # settings are rejected instead of being silently discarded.
            common = {'idx','order','tx','spPr','dPt','dLbls','trendline','errBars','cat','val','extLst'}
            series = _children(kind,C,'ser')
            for series_node in series:
                if any(n.nodeType==n.ELEMENT_NODE and n.localName not in common for n in series_node.childNodes):
                    raise ModelValidationError('Chart series has type-specific formatting that cannot be converted safely')
            replacement = doc.importNode(target, True)
            replacement.setAttribute('xmlns:c', C);replacement.setAttribute('xmlns:a', A)
            for n in _children(replacement,C,'axId'):
                replacement.removeChild(n)
            for n in _children(kind,C,'axId'):
                replacement.appendChild(n.cloneNode(True))
            anchor = next((n for n in replacement.childNodes if n.nodeType==n.ELEMENT_NODE and n.localName not in ('barDir','grouping','varyColors')),None)
            for n in series:
                replacement.insertBefore(n.cloneNode(True),anchor)
            plot.replaceChild(replacement,kind);kind=replacement
        for axis in ('catAx','valAx'):
            axes, targets = _children(plot,C,axis), _children(tplot,C,axis)
            if len(axes)==1 and len(targets)==1:
                copy_child(axes[0],targets[0],'axPos',axis_order)
    if 'show_data_labels' in changes:
        labels, tlabels = _children(kind,C,'dLbls'),_children(target,C,'dLbls')
        if labels and tlabels:
            copy_child(labels[0],tlabels[0],'showVal',['dLbl','delete','numFmt','spPr','txPr','dLblPos','showLegendKey','showVal','showCatName','showSerName','showPercent','showBubbleSize','separator','showLeaderLines','leaderLines','extLst'])
        else:
            copy_child(kind,target,'dLbls',['barDir','grouping','varyColors','ser','dLbls','dropLines','hiLowLines','upDownBars','marker','smooth','gapWidth','overlap','serLines','axId','extLst'])
    package.changed.add(record.part)


def _regenerate(package, record, spec, model):
    """Replace a supported, self-contained layout without touching other cells."""
    from .report_package import _definitions
    from .reports import WorkbookReport
    if model is None or record.spec is None:
        raise ModelValidationError('Axis edits require the current model and a supported saved report')
    if spec.sheet != record.sheet or spec.cell.upper() != record.spec.cell.upper():
        raise ModelValidationError('Moving an existing PivotTable is not supported')
    WorkbookReport(model)._validate((spec,), ())
    spath, part_node, cpath = package.bindings[record.part]
    if any(p.startswith(('xl/slicerCaches/', 'xl/timelineCaches/')) for p in package.parts):
        raise ModelValidationError('Cannot regenerate reports with unresolved slicer/timeline bindings')
    incoming = []
    for path in package.parts:
        if path.startswith('xl/pivotTables/') and path.endswith('.xml.rels'):
            owner = posixpath.join(posixpath.dirname(path).removesuffix('/_rels'), posixpath.basename(path)[:-5])
            incoming.extend(n for n in package.relationships(owner) if _target(owner, n) == cpath)
    if len(incoming) != 1:
        raise ModelValidationError('Cannot regenerate a shared PivotCache')
    pivot = package.doc(record.part).documentElement
    cache = package.doc(cpath)
    source = _nodes(cache, S, 'cacheSource')[0]
    payloads = _definitions(spec, model, int(pivot.getAttribute('cacheId')), int(source.getAttribute('connectionId')))
    previous_payloads = _definitions(record.spec, model, int(pivot.getAttribute('cacheId')), int(source.getAttribute('connectionId')))
    # Unknown fields/attributes must never disappear during regeneration.
    # Permit native cached row/column items and literal member inventories; they
    # are disposable calculated state and are rebuilt by Excel on refresh.
    disposable = {'rowItems', 'colItems', 'items', 'sharedItems', 'tupleCache'}
    for path, payload, previous in zip((cpath, record.part), payloads, previous_payloads):
        replacement = minidom.parseString(payload)
        allowed = {}
        for node in list(replacement.getElementsByTagName('*')) + list(minidom.parseString(previous).getElementsByTagName('*')):
            allowed.setdefault((node.namespaceURI, node.localName), set()).update(
                (a.namespaceURI, a.localName) for a in node.attributes.values())
        for node in package.doc(path).getElementsByTagName('*'):
            parent = node
            ignore = False
            while parent.nodeType == parent.ELEMENT_NODE:
                if parent.namespaceURI == S and parent.localName in disposable:
                    ignore = True
                    break
                parent = parent.parentNode
            if ignore:
                continue
            key = (node.namespaceURI, node.localName)
            attributes = {(a.namespaceURI, a.localName) for a in node.attributes.values()
                          if a.namespaceURI != 'http://www.w3.org/2000/xmlns/'}
            if key not in allowed or attributes - allowed[key]:
                raise ModelValidationError('Report contains metadata that cannot be safely regenerated: '+node.tagName)
        package.docs[path] = replacement
        package.changed.add(path)
    # Saved values in this object's old range are stale after changing axes.
    # Keep cells and styles but clear only their calculated values/formulas.
    location = _nodes(pivot, S, 'location')[0].getAttribute('ref')
    def coordinate(value):
        match = re.fullmatch(r'\$?([A-Za-z]+)\$?([1-9][0-9]*)', value)
        if not match:
            raise ModelValidationError('Unsupported PivotTable location')
        column = 0
        for char in match[1].upper():
            column = column*26+ord(char)-64
        return column, int(match[2])
    bounds = location.split(':')
    lo, hi = coordinate(bounds[0]), coordinate(bounds[-1])
    sheet = package.doc(spath)
    for cell in _nodes(sheet, S, 'c'):
        col, row = coordinate(cell.getAttribute('r'))
        if lo[0] <= col <= hi[0] and lo[1] <= row <= hi[1]:
            for child in list(cell.childNodes):
                if child.nodeType == child.ELEMENT_NODE and child.namespaceURI == S and child.localName in ('f', 'v', 'is'):
                    cell.removeChild(child)
            if cell.hasAttribute('t'):
                cell.removeAttribute('t')
            package.changed.add(spath)


def edit_reports(source, destination, *, pivot_updates=None, chart_updates=None, control_updates=None,
                 remove=(), workbook_name=None, model=None):
    """Atomically apply metadata edits/removals while retaining unknown XML.

    Pivot updates support name, layout, row_grand_totals, column_grand_totals.
    Chart updates support name, title, type and the public formatting options;
    cached series are retained. Full PivotTableSpec updates, or axis dictionaries,
    require model and reject metadata that cannot be preserved.
    Remove dependent charts explicitly together with their PivotTable.
    """
    package = _Package(source)
    reports = package.inventory()
    controls = _inspect_controls(package)
    records = reports + controls
    by_name = {record.name: record for record in records}
    if len({record.name.casefold() for record in records}) != len(records):
        raise ModelValidationError('Existing report/control names are ambiguous')
    pivot_updates, chart_updates = dict(pivot_updates or {}), dict(chart_updates or {})
    control_updates = dict(control_updates or {})
    remove = tuple(remove)
    requested = set(pivot_updates) | set(chart_updates) | set(control_updates) | set(remove)
    if requested - set(by_name):
        raise ModelValidationError('Unknown saved report: '+', '.join(sorted(requested-set(by_name))))
    if set(remove) & (set(pivot_updates) | set(chart_updates) | set(control_updates)):
        raise ModelValidationError('Cannot update and remove the same report')
    regenerate = {}
    surgical = {'name', 'layout', 'row_grand_totals', 'column_grand_totals'}
    for name, changes in tuple(pivot_updates.items()):
        record = by_name[name]
        if record.kind != 'pivot_table':
            raise ModelValidationError('Expected an existing PivotTable')
        if isinstance(changes, PivotTableSpec):
            spec = changes
        elif set(changes) - surgical:
            if record.spec is None:
                raise ModelValidationError('Saved report has an unsupported layout')
            try:
                spec = replace(record.spec, **changes)
            except TypeError as exc:
                raise ModelValidationError('Unsupported PivotTable edit') from exc
        else:
            continue
        regenerate[name] = spec
        pivot_updates[name] = {key: getattr(spec, key) for key in surgical}
    new_names = []
    for record in records:
        if record.name not in remove:
            changes = (pivot_updates if record.kind == 'pivot_table' else chart_updates if record.kind=='pivot_chart' else control_updates).get(record.name, {})
            changed_name = changes.get('name', record.name)
            _text(changed_name, 'Report name')
            new_names.append(changed_name.casefold())
    if len(new_names) != len(set(new_names)):
        raise ModelValidationError('Duplicate report name after editing')
    for name, spec in regenerate.items():
        _regenerate(package, by_name[name], spec, model)
    for name,changes in control_updates.items():
        record=by_name[name]
        if record.kind not in ('slicer','timeline'):
            raise ModelValidationError('Expected an existing slicer or timeline')
        _edit_control(package,record,changes,controls,reports)
    for control in controls:
        if control.name in remove:
            continue
        if any(name in remove for _,name in control.connected_pivots):
            raise ModelValidationError('Remove dependent controls before their PivotTables')
        # Native cache bindings identify PivotTables by name; keep them in sync.
        cache=package.doc(control.cache_part)
        for node in cache.getElementsByTagName('*'):
            if node.localName=='pivotTable' and node.getAttribute('name') in pivot_updates:
                new=pivot_updates[node.getAttribute('name')].get('name')
                if new:
                    node.setAttribute('name',new);package.changed.add(control.cache_part)
    for name, changes in pivot_updates.items():
        record = by_name[name]
        if record.kind != 'pivot_table' or set(changes)-{'name', 'layout', 'row_grand_totals', 'column_grand_totals'}:
            raise ModelValidationError('Unsupported PivotTable edit')
        pivot = package.doc(record.part).documentElement
        if 'name' in changes:
            _pivot_name(changes['name'])
            pivot.setAttribute('name', changes['name'])
        if 'layout' in changes:
            layout = changes['layout']
            if layout not in ('tabular', 'outline', 'compact'):
                raise ModelValidationError('Unsupported PivotTable layout')
            pivot.setAttribute('compact', str(int(layout == 'compact')))
            pivot.setAttribute('compactData', str(int(layout == 'compact')))
            for field in _nodes(pivot, S, 'pivotField'):
                field.setAttribute('compact', str(int(layout == 'compact')))
                field.setAttribute('outline', str(int(layout != 'tabular')))
        for key, attribute in (('row_grand_totals', 'rowGrandTotals'), ('column_grand_totals', 'colGrandTotals')):
            if key in changes:
                if type(changes[key]) is not bool:
                    raise ModelValidationError('Grand total options must be booleans')
                pivot.setAttribute(attribute, str(int(changes[key])))
        package.changed.add(record.part)
    for record in records:
        if record.kind != 'pivot_chart' or record.name in remove:
            continue
        chart = package.doc(record.part)
        sources = _nodes(chart, C, 'pivotSource')
        source_name = _nodes(sources[0], C, 'name')[0]
        old_source = _value(source_name)
        old_pivot = old_source.rsplit('!', 1)[-1]
        if old_pivot in remove:
            raise ModelValidationError('Remove dependent PivotCharts together with their PivotTable')
        renamed = pivot_updates.get(old_pivot, {}).get('name')
        if renamed or workbook_name is not None:
            new_source = old_source.rsplit('!', 1)[0]+'!'+(renamed or old_pivot)
            if workbook_name is not None:
                _text(workbook_name, 'Workbook filename')
                if any(c in workbook_name for c in '[]\\/') or not new_source.startswith('[') or ']' not in new_source:
                    raise ModelValidationError('Unsupported chart workbook source')
                new_source = '['+workbook_name+']'+new_source.split(']', 1)[1]
            _set_text(source_name, new_source)
            package.changed.add(record.part)
    for name, changes in chart_updates.items():
        record = by_name[name]
        formatting = {'style','show_legend','legend_position','category_axis_title','value_axis_title',
                      'number_format','show_data_labels','background_color','plot_background_color','chart_type'}
        if record.kind != 'pivot_chart' or set(changes)-({'name', 'title'} | formatting):
            raise ModelValidationError('Unsupported PivotChart edit')
        if set(changes) & formatting:
            _chart_format(package, record, {key:value for key,value in changes.items() if key in formatting})
        dpath, anchor, chart_node, props = package.bindings[record.part]
        if 'name' in changes:
            _text(changes['name'], 'Chart name')
            if len(changes['name']) > 255:
                raise ModelValidationError('Chart name is too long')
            props.setAttribute('name', changes['name'])
            package.changed.add(dpath)
        if 'title' in changes:
            _xml_text(changes['title'], 'Chart title')
            doc = package.doc(record.part)
            chart_root = _nodes(doc, C, 'chart')[0]
            titles = _children(chart_root, C, 'title')
            if not titles and changes['title']:
                from .report_charts import chart_xml
                template = minidom.parseString(chart_xml({'title': changes['title']}, 'Pivot', 'Report', 'report.xlsx'))
                title = doc.importNode(_children(_nodes(template, C, 'chart')[0], C, 'title')[0], True)
                # Imported prefixes must remain declared in a native document.
                title.setAttribute('xmlns:c', C)
                title.setAttribute('xmlns:a', A)
                chart_root.insertBefore(title, chart_root.firstChild)
                titles = [title]
            if not titles and not changes['title']:
                continue
            # A literal one-run title can be updated without flattening rich text.
            texts = _nodes(titles[0], A, 't') if len(titles) == 1 else []
            if len(texts) != 1 or _nodes(titles[0], C, 'strRef'):
                raise ModelValidationError('Only existing single-run literal chart titles can be edited')
            _set_text(texts[0], changes['title'])
            for flag in _children(chart_root, C, 'autoTitleDeleted'):
                flag.setAttribute('val', '0' if changes['title'] else '1')
            package.changed.add(record.part)
    for name in remove:
        record = by_name[name]
        if record.kind in ('slicer','timeline'):
            _remove_control(package,record,tuple(c for c in controls if c.name not in remove))
            continue
        if record.kind == 'pivot_table':
            linked_cache_parts={c.cache_part for c in controls}
            if any(p.startswith(('xl/slicerCaches/','xl/timelineCaches/')) and p.endswith('.xml') and p not in linked_cache_parts for p in package.parts):
                raise ModelValidationError('Cannot remove PivotTables with unresolved slicer or timeline bindings')
            spath, part_node, cpath = package.bindings[record.part]
            package.unlink(spath, part_node.getAttributeNS(R, 'id') or part_node.getAttribute('Id'))
            if part_node.localName == 'pivotTablePart':
                parent = part_node.parentNode
                parent.removeChild(part_node)
                remaining = [n for n in parent.childNodes if n.nodeType == n.ELEMENT_NODE]
                if remaining:
                    parent.setAttribute('count', str(len(remaining)))
                else:
                    parent.parentNode.removeChild(parent)
                package.changed.add(spath)
        else:
            dpath, anchor, chart_node, props = package.bindings[record.part]
            package.unlink(dpath, chart_node.getAttributeNS(R, 'id'))
            anchor.parentNode.removeChild(anchor)
            package.changed.add(dpath)
        package.deleted.add(record.part)
        if _rels_path(record.part) in package.parts:
            # Chart relationships may own embedded workbooks/styles/images. Keep
            # those data parts; removing only their owner cannot harm consumers.
            package.deleted.add(_rels_path(record.part))
    if package.deleted:
        types = package.doc('[Content_Types].xml')
        for node in list(types.documentElement.childNodes):
            if node.nodeType == node.ELEMENT_NODE and node.getAttribute('PartName').lstrip('/') in package.deleted:
                types.documentElement.removeChild(node)
        package.changed.add('[Content_Types].xml')
        # Reject any remaining unexpected incoming edge to a deleted part.
        for path in package.parts:
            if not path.endswith('.rels') or path in package.deleted:
                continue
            owner = posixpath.join(posixpath.dirname(path).removesuffix('/_rels'), posixpath.basename(path)[:-5])
            for node in package.doc(path).documentElement.childNodes:
                if node.nodeType == node.ELEMENT_NODE and node.getAttribute('TargetMode') != 'External' and _target(owner, node) in package.deleted:
                    raise ModelValidationError('Another workbook object still references the removed report')
    return package.write(destination)
