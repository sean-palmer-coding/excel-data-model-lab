"""OOXML report packaging. Model calculation remains the spreadsheet engine's job."""
from __future__ import annotations

import posixpath
import re
import shutil
import zipfile
from uuid import uuid4
from pathlib import Path
from xml.dom import minidom
from xml.etree import ElementTree as ET

from .errors import ModelValidationError
from .measures import list_script_measures

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'
X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
XR = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
ET.register_namespace('xr', XR)
ET.register_namespace('mc', MC)


def _e(parent, tag, **attrs):
    return ET.SubElement(parent, '{'+S+'}'+tag, {k: str(v) for k, v in attrs.items()})


def _xml(root):
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def _preserving_xml(root, original=None):
    """Retain namespace declarations used by QName-valued extension attributes."""
    data = _xml(root)
    if original is None:
        return data
    before = minidom.parseString(original)
    after = minidom.parseString(data)
    for node in before.getElementsByTagName('*'):
        for attr in node.attributes.values():
            if attr.name == 'xmlns' or attr.name.startswith('xmlns:'):
                if not after.documentElement.hasAttribute(attr.name):
                    after.documentElement.setAttribute(attr.name, attr.value)
    return after.toxml(encoding='utf-8')


def _mdx(name):
    return '[' + name.replace(']', ']]') + ']'


def _measures(model):
    result = []
    for f in model._image.files:
        if f.logical_path and f.logical_path.endswith('.scr.xml'):
            result.extend(list_script_measures(model._image.read_logical_file(f)))
    result.extend(model._pending_measures)
    return result


def _definitions(spec, model, cache_id, connection_id):
    """Write an unpopulated OLAP layout with explicit refresh-on-open."""
    axes = list(spec.rows) + list(spec.columns) + list(spec.filters)
    measures = {m.name: m for m in _measures(model)}
    fields = [(t, c, False) for t, c in axes] + [(measures[n].table, n, True) for n in spec.measures]
    row_indices=[fields.index((t,c,False)) for t,c in spec.rows]
    col_indices=[fields.index((t,c,False)) for t,c in spec.columns]
    filter_indices=[fields.index((t,c,False)) for t,c in spec.filters]
    measure_indices=[fields.index((measures[n].table,n,True)) for n in spec.measures]
    columns = list(model.columns)
    hierarchies = [(c.table, c.name, False) for c in columns]
    hierarchies.extend((m.table, m.name, True) for m in measures.values())
    # Keep references ordered and avoid duplicate hidden/helper entries.
    hierarchies = list(dict.fromkeys(hierarchies))
    hi = {h: i for i, h in enumerate(hierarchies)}
    tables = list(dict.fromkeys(c.table for c in columns))
    cache = ET.Element('{'+S+'}pivotCacheDefinition', {
        '{'+XR+'}uid':'{'+str(uuid4()).upper()+'}', '{'+MC+'}Ignorable':'xr',
        'saveData': '0', 'refreshOnLoad': '1', 'enableRefresh': '1',
        'createdVersion': '6', 'refreshedVersion': '8', 'minRefreshableVersion': '3',
        'recordCount': '0', 'supportSubquery': '1', 'supportAdvancedDrill': '1'})
    _e(cache, 'cacheSource', type='external', connectionId=connection_id)
    fs = _e(cache, 'cacheFields', count=len(fields))
    for t, c, measure in fields:
        name = '[Measures].'+_mdx(c) if measure else _mdx(t)+'.'+_mdx(c)+'.'+_mdx(c)
        _e(fs, 'cacheField', name=name, caption=c, numFmtId=0,
           hierarchy=hi[t, c, measure], level=32767 if measure else 1)
    hs = _e(cache, 'cacheHierarchies', count=len(hierarchies))
    types = { (c.table,c.name): c.data_type for c in columns }
    for h in hierarchies:
        t,c,measure = h
        name = '[Measures].'+_mdx(c) if measure else _mdx(t)+'.'+_mdx(c)
        attrs = dict(uniqueName=name,caption=c,displayFolder='',count=0)
        if measure:
            attrs.update(measure=1,measureGroup=t)
        else:
            typ = types.get((t,c),'')
            dtype = 130 if typ in ('string','object','String','Text','text') else 20
            if 'float' in typ.lower() or typ.lower() in ('decimal','double'): dtype=5
            if typ.lower() == 'currency': dtype=6
            if 'date' in typ.lower(): dtype=7
            if 'bool' in typ.lower(): dtype=11
            attrs.update(attribute=1,defaultMemberUniqueName=name+'.[All]',
                         allUniqueName=name+'.[All]',dimensionUniqueName=_mdx(t),
                         memberValueDatatype=dtype,unbalanced=0)
        node = _e(hs,'cacheHierarchy',**attrs)
        if h in fields:
            idx=fields.index(h)
            if measure: node.set('oneField','1')
            else: node.set('count','2')
            usage=_e(node,'fieldsUsage',count=1 if measure else 2)
            if not measure: _e(usage,'fieldUsage',x=-1)
            _e(usage,'fieldUsage',x=idx)
    _e(cache,'kpis',count=0)
    dims=_e(cache,'dimensions',count=len(tables)+1)
    _e(dims,'dimension',measure=1,name='Measures',uniqueName='[Measures]',caption='Measures')
    groups=_e(cache,'measureGroups',count=len(tables))
    for t in tables:
        _e(dims,'dimension',name=t,uniqueName=_mdx(t),caption=t)
        _e(groups,'measureGroup',name=t,caption=t)
    # OLAP dimension usage follows active model relationships.
    from .relationship_storage import stored_relationships
    rels=[r for _,r in stored_relationships(model)] + list(model._pending_relationships)
    maps=[]
    for i,t in enumerate(tables):
        reachable={t}
        while True:
            expanded=reachable | {r.to_table for r in rels if r.active and r.from_table in reachable}
            if expanded==reachable: break
            reachable=expanded
        maps.extend((i,tables.index(d)+1) for d in tables if d in reachable)
    mapping=_e(cache,'maps',count=len(maps))
    for g,d in maps: _e(mapping,'map',measureGroup=g,dimension=d)
    ext=_e(_e(cache,'extLst'),'ext',uri='{725AE2AE-9491-48be-B2B4-4EB974FC3084}')
    ET.SubElement(ext,'{'+X14+'}pivotCacheDefinition',{'supportSubqueryNonVisual':'1','supportSubqueryCalcMem':'1','supportAddCalcMems':'1'})

    pivot=ET.Element('{'+S+'}pivotTableDefinition',{
        '{'+XR+'}uid':'{'+str(uuid4()).upper()+'}', '{'+MC+'}Ignorable':'xr',
        'name':spec.name,'cacheId':str(cache_id),'dataCaption':'Values',
        'updatedVersion':'8','minRefreshableVersion':'3','createdVersion':'6',
        'useAutoFormatting':'1','compact':str(int(spec.layout=='compact')),'compactData':str(int(spec.layout=='compact')),'indent':'0',
        'rowGrandTotals':str(int(spec.row_grand_totals)),
        'colGrandTotals':str(int(spec.column_grand_totals))})
    location=_e(pivot,'location',ref=spec.cell.upper()+':'+spec.cell.upper(),firstHeaderRow=0,firstDataRow=1,firstDataCol=1)
    if spec.filters:
        location.set('rowPageCount',str(len(spec.filters)))
        location.set('colPageCount','1')
    pf=_e(pivot,'pivotFields',count=len(fields))
    nr,nc,nf=len(spec.rows),len(spec.columns),len(spec.filters)
    for i,(_,_,measure) in enumerate(fields):
        attrs=dict(compact=int(spec.layout=='compact'),outline=int(spec.layout!='tabular'),showAll=0)
        if i in measure_indices: attrs['dataField']=1
        elif i in row_indices+col_indices+filter_indices: attrs.update(axis='axisRow' if i in row_indices else 'axisCol' if i in col_indices else 'axisPage',
                           allDrilled=1,dataSourceSort=1,defaultAttributeDrillState=1)
        field_node=_e(pf,'pivotField',**attrs)
        if i in filter_indices:
            _e(_e(field_node,'items',count=1),'item',t='default')
    if nr:
        rows=_e(pivot,'rowFields',count=nr)
        for i in row_indices: _e(rows,'field',x=i)
    if nc or len(spec.measures)>1:
        cols=_e(pivot,'colFields',count=nc+int(len(spec.measures)>1))
        for i in col_indices: _e(cols,'field',x=i)
        if len(spec.measures)>1: _e(cols,'field',x=-2)
    if nf:
        pages=_e(pivot,'pageFields',count=nf)
        for i in filter_indices:
            t,c,_=fields[i]
            _e(pages,'pageField',fld=i,hier=hi[fields[i]],name=_mdx(t)+'.'+_mdx(c)+'.[All]',cap='All')
    df=_e(pivot,'dataFields',count=len(spec.measures))
    for i in measure_indices: _e(df,'dataField',fld=i,subtotal='count',baseField=0,baseItem=0)
    ph=_e(pivot,'pivotHierarchies',count=len(hierarchies))
    for _,_,measure in hierarchies:
        attrs={'dragToData':1}
        if measure: attrs.update(dragToRow=0,dragToCol=0,dragToPage=0)
        _e(ph,'pivotHierarchy',**attrs)
    _e(pivot,'pivotTableStyleInfo',name='PivotStyleLight16',showRowHeaders=1,showColHeaders=1,showRowStripes=0,showColStripes=0,showLastColumn=1)
    for axis,selected in [('row',spec.rows),('col',spec.columns)]:
        values_axis = axis == 'col' and len(spec.measures)>1
        if selected or values_axis:
            usage=_e(pivot,axis+'HierarchiesUsage',count=len(selected)+int(values_axis))
            for t,c in selected: _e(usage,axis+'HierarchyUsage',hierarchyUsage=hi[t,c,False])
            if values_axis: _e(usage,axis+'HierarchyUsage',hierarchyUsage=-2)
    ext=_e(_e(pivot,'extLst'),'ext',uri='{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}')
    ET.SubElement(ext,'{'+X14+'}pivotTableDefinition',{'calculatedMembersInFilters':'1','hideValuesRow':'1'})
    from .report_filters import apply_pivot_options
    apply_pivot_options(pivot, cache, spec, model)
    return _xml(cache),_xml(pivot)


def add_reports(source, destination, pivots, charts, *, model, workbook_name=None, slicers=(), timelines=()):
    """Append report sheets, retaining all unrelated source members byte-for-byte."""
    if not pivots and not charts and not slicers and not timelines:
        if Path(source).resolve() != Path(destination).resolve():
            shutil.copyfile(source, destination)
        return
    with zipfile.ZipFile(source) as z:
        parts={n:z.read(n) for n in z.namelist()}
        original_info=z.infolist(); comment=z.comment
    wb=minidom.parseString(parts['xl/workbook.xml'])
    wr=minidom.parseString(parts['xl/_rels/workbook.xml.rels'])
    ct=minidom.parseString(parts['[Content_Types].xml'])
    # minidom tracks namespace URIs but does not serialize declarations for us.
    for document, namespace in ((wb,S),(wr,P),(ct,C)):
        document.documentElement.setAttribute('xmlns',namespace)
    wb.documentElement.setAttribute('xmlns:r',R)
    sheets=wb.getElementsByTagNameNS(S,'sheets')[0]
    existing_sheets={n.getAttribute('name').casefold() for n in sheets.childNodes if n.nodeType==n.ELEMENT_NODE}
    report_names={ET.fromstring(v).get('name','').casefold() for k,v in parts.items() if re.fullmatch(r'xl/pivotTables/pivotTable\d+\.xml',k)}
    if any(p.name.casefold() in report_names for p in pivots):
        raise ModelValidationError('PivotTable name already exists in workbook')
    connrels=[n for n in wr.documentElement.childNodes if n.nodeType==n.ELEMENT_NODE and n.getAttribute('Type')==R+'/connections']
    if len(connrels)!=1: raise ModelValidationError('Workbook needs one Data Model connections relationship')
    target=connrels[0].getAttribute('Target')
    connpath=posixpath.normpath(posixpath.join('xl',target)).lstrip('/')
    conns=ET.fromstring(parts[connpath])
    matches=[n for n in conns if n.get('name','').casefold()=='thisworkbookdatamodel']
    if len(matches)!=1: raise ModelValidationError('Workbook needs a ThisWorkbookDataModel connection')
    connection_id=int(matches[0].get('id'))

    def alloc(prefix,suffix='.xml'):
        i=1
        while prefix+str(i)+suffix in parts: i+=1
        name=prefix+str(i)+suffix;parts[name]=b''
        return name

    def rel(doc,kind,target):
        ids={n.getAttribute('Id') for n in doc.documentElement.childNodes if n.nodeType==n.ELEMENT_NODE}
        i=1
        while 'rId'+str(i) in ids:i+=1
        node=doc.createElementNS(P,'Relationship')
        for k,v in [('Id','rId'+str(i)),('Type',R+'/'+kind),('Target',target)]:node.setAttribute(k,v)
        doc.documentElement.appendChild(node)
        return 'rId'+str(i)

    def types(path,kind):
        node=ct.createElementNS(C,'Override');node.setAttribute('PartName','/'+path)
        node.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.'+kind+'+xml')
        ct.documentElement.appendChild(node)

    def new_rels():return minidom.parseString('<Relationships xmlns="'+P+'"/>')

    def relpath(path):return posixpath.dirname(path)+'/_rels/'+posixpath.basename(path)+'.rels'

    new_sheets={}
    def sheet(name, allow_existing=False):
        key=name.casefold()
        if key in new_sheets:return new_sheets[key]
        if name.casefold() in existing_sheets:
            if not allow_existing:raise ModelValidationError('Reports require a new worksheet: '+name)
            matched=next(n for n in sheets.childNodes if n.nodeType==n.ELEMENT_NODE and n.getAttribute('name').casefold()==name.casefold())
            rid=matched.getAttributeNS(R,'id')
            target=next(n.getAttribute('Target') for n in wr.documentElement.childNodes if n.nodeType==n.ELEMENT_NODE and n.getAttribute('Id')==rid)
            path=posixpath.normpath(posixpath.join('xl',target)).lstrip('/')
            root=ET.fromstring(parts[path])
            doc=minidom.parseString(parts[relpath(path)]) if relpath(path) in parts else new_rels()
            new_sheets[key]=(path,root,doc)
            return new_sheets[key]
        existing_sheets.add(name.casefold())
        path=alloc('xl/worksheets/sheet')
        root=ET.Element('{'+S+'}worksheet');_e(root,'sheetData')
        doc=new_rels()
        rid=rel(wr,'worksheet',posixpath.relpath(path,'xl'))
        sid=max([int(n.getAttribute('sheetId')) for n in sheets.childNodes if n.nodeType==n.ELEMENT_NODE]+[0])+1
        n=wb.createElementNS(S,'sheet');n.setAttribute('name',name);n.setAttribute('sheetId',str(sid));n.setAttributeNS(R,'r:id',rid);sheets.appendChild(n)
        types(path,'spreadsheetml.worksheet')
        new_sheets[key]=(path,root,doc)
        return new_sheets[key]

    pcs=wb.getElementsByTagNameNS(S,'pivotCaches')
    if pcs:pcs=pcs[0]
    else:
        pcs=wb.createElementNS(S,'pivotCaches')
        # pivotCaches follows calcPr and precedes smartTag/web/ext collections.
        later={'smartTagPr','smartTagTypes','webPublishing','fileRecoveryPr','webPublishObjects','extLst'}
        anchor=next((n for n in wb.documentElement.childNodes if n.nodeType==n.ELEMENT_NODE and n.localName in later),None)
        wb.documentElement.insertBefore(pcs,anchor)
    # Slicer/timeline caches register their cacheIds under x14/x15 extension
    # lists; every workbook-level pivotCache shares one cacheId space.
    cache_ids={int(n.getAttribute('cacheId')) for n in wb.getElementsByTagNameNS(S,'pivotCache') if n.hasAttribute('cacheId')}
    for spec in pivots:
        path,ws,sr=sheet(spec.sheet)
        cachepath=alloc('xl/pivotCache/pivotCacheDefinition')
        cache_id=max(cache_ids|{0})+1;cache_ids.add(cache_id)
        rid=rel(wr,'pivotCacheDefinition',posixpath.relpath(cachepath,'xl'))
        node=wb.createElementNS(S,'pivotCache');node.setAttribute('cacheId',str(cache_id));node.setAttributeNS(R,'r:id',rid);pcs.appendChild(node)
        types(cachepath,'spreadsheetml.pivotCacheDefinition')
        pivotpath=alloc('xl/pivotTables/pivotTable')
        parts[cachepath],parts[pivotpath]=_definitions(spec,model,cache_id,connection_id)
        pr=new_rels();rel(pr,'pivotCacheDefinition',posixpath.relpath(cachepath,posixpath.dirname(pivotpath)))
        parts[relpath(pivotpath)]=pr.toxml(encoding='utf-8')
        pparts=ws.find('{'+S+'}pivotTableParts')
        if pparts is None:pparts=_e(ws,'pivotTableParts',count=0)
        rid=rel(sr,'pivotTable',posixpath.relpath(pivotpath,posixpath.dirname(path)))
        _e(pparts,'pivotTablePart',**{'{'+R+'}id':rid});pparts.set('count',str(len(pparts)))
        types(pivotpath,'spreadsheetml.pivotTable')

    if charts:
        from .report_editing import inspect_reports
        existing_pivots=tuple(r.spec for r in inspect_reports(source) if r.kind=='pivot_table' and r.spec is not None)
        _add_charts(parts,charts,existing_pivots+tuple(pivots),lambda name:sheet(name,True),alloc,types,rel,new_rels,relpath,workbook_name or Path(destination).name)
    for control in tuple(slicers)+tuple(timelines):
        if control.sheet is not None and control.sheet.casefold() not in existing_sheets:
            sheet(control.sheet)
    for path,root,sr in new_sheets.values():
        parts[path]=_preserving_xml(root,parts.get(path) or None);parts[relpath(path)]=sr.toxml(encoding='utf-8')
    parts['xl/workbook.xml']=wb.toxml(encoding='utf-8')
    parts['xl/_rels/workbook.xml.rels']=wr.toxml(encoding='utf-8')
    parts['[Content_Types].xml']=ct.toxml(encoding='utf-8')
    if slicers or timelines:
        from .report_slicers import append_controls
        append_controls(parts, slicers, timelines, model=model, workbook_name=workbook_name or Path(destination).name)
    with zipfile.ZipFile(destination,'w',compression=zipfile.ZIP_DEFLATED) as z:
        z.comment=comment
        for info in original_info:z.writestr(info,parts.pop(info.filename))
        for name,data in parts.items():z.writestr(name,data)


def _add_charts(parts,charts,pivots,sheet,alloc,types,rel,new_rels,relpath,workbook_name):
    from .report_charts import chart_xml
    X='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    A='http://schemas.openxmlformats.org/drawingml/2006/main'
    CH='http://schemas.openxmlformats.org/drawingml/2006/chart'
    pivots={p.name:p for p in pivots}
    drawings={}
    for spec in charts:
        pivot=pivots[spec.pivot_table]
        sheet_name=spec.sheet or pivot.sheet
        spath,ws,sr=sheet(sheet_name)
        if spath not in drawings:
            dn=ws.find('{'+S+'}drawing')
            if dn is not None:
                rid=dn.get('{'+R+'}id')
                target=next(n.getAttribute('Target') for n in sr.documentElement.childNodes if n.nodeType==n.ELEMENT_NODE and n.getAttribute('Id')==rid)
                dpath=posixpath.normpath(posixpath.join(posixpath.dirname(spath),target)).lstrip('/')
                drawing=ET.fromstring(parts[dpath])
                dr=minidom.parseString(parts[relpath(dpath)]) if relpath(dpath) in parts else new_rels()
            else:
                dpath=alloc('xl/drawings/drawing');dr=new_rels();drawing=ET.Element('{'+X+'}wsDr')
                rid=rel(sr,'drawing',posixpath.relpath(dpath,posixpath.dirname(spath)))
                dn=ET.Element('{'+S+'}drawing',{'{'+R+'}id':rid})
                # Worksheet drawing comes before legacy/control/pivot/ext sections.
                later={'legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','pivotTableParts','extLst'}
                index=next((i for i,n in enumerate(ws) if n.tag.rsplit('}',1)[-1] in later),len(ws))
                ws.insert(index,dn)
                types(dpath,'drawing')
            drawings[spath]=(dpath,drawing,dr)
        dpath,drawing,dr=drawings[spath]
        cpath=alloc('xl/charts/chart')
        parts[cpath]=chart_xml(spec,pivot.name,pivot.sheet,workbook_name)
        types(cpath,'drawingml.chart')
        rid=rel(dr,'chart',posixpath.relpath(cpath,posixpath.dirname(dpath)))
        match=re.fullmatch(r'([A-Za-z]+)([0-9]+)',spec.cell)
        col=0
        for ch in match[1].upper():col=col*26+ord(ch)-64
        anchor=ET.SubElement(drawing,'{'+X+'}oneCellAnchor')
        origin=ET.SubElement(anchor,'{'+X+'}from')
        for tag,value in [('col',col-1),('colOff',0),('row',int(match[2])-1),('rowOff',0)]:
            ET.SubElement(origin,'{'+X+'}'+tag).text=str(value)
        ET.SubElement(anchor,'{'+X+'}ext',{'cx':str(spec.width*9525),'cy':str(spec.height*9525)})
        frame=ET.SubElement(anchor,'{'+X+'}graphicFrame',{'macro':''})
        nv=ET.SubElement(frame,'{'+X+'}nvGraphicFramePr')
        shape_id=max([int(n.get('id','0')) for n in drawing.iter('{'+X+'}cNvPr')]+[0])+1
        ET.SubElement(nv,'{'+X+'}cNvPr',{'id':str(shape_id),'name':spec.name})
        ET.SubElement(nv,'{'+X+'}cNvGraphicFramePr')
        xf=ET.SubElement(frame,'{'+X+'}xfrm')
        ET.SubElement(xf,'{'+A+'}off',{'x':'0','y':'0'});ET.SubElement(xf,'{'+A+'}ext',{'cx':'0','cy':'0'})
        graphic=ET.SubElement(frame,'{'+A+'}graphic')
        data=ET.SubElement(graphic,'{'+A+'}graphicData',{'uri':CH})
        ET.SubElement(data,'{'+CH+'}chart',{'{'+R+'}id':rid})
        ET.SubElement(anchor,'{'+X+'}clientData')
    for dpath,drawing,dr in drawings.values():
        parts[dpath]=_preserving_xml(drawing,parts.get(dpath) or None);parts[relpath(dpath)]=dr.toxml(encoding='utf-8')
