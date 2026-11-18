"""Package integrity checks independent of Excel's calculation engine."""
from pathlib import Path
import posixpath
from xml.dom import minidom
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from excel_model_lab import WorkbookModel, ModelValidationError
from excel_model_lab.reports import WorkbookReport, PivotTableSpec, PivotChartSpec
from excel_model_lab.report_package import add_reports, S, R, P, C


FIXTURE = Path(__file__).parent / 'data' / 'native-authoring-base.xlsx'


def _model():
    model = WorkbookModel(FIXTURE)
    model.add_measure('Table1', 'TotalID', 'SUM(Table1[ID])')
    return model


def _pivot(name='Summary', sheet='Report'):
    return PivotTableSpec(name, sheet, rows=(('Table1', 'Label'),), measures=('TotalID',))


def _assert_links(archive):
    names = set(archive.namelist())
    assert len(names) == len(archive.namelist())
    for path in names:
        if not path.endswith('.rels'):
            continue
        root = ET.fromstring(archive.read(path))
        ids = [node.get('Id') for node in root]
        assert len(ids) == len(set(ids))
        owner = '' if path == '_rels/.rels' else posixpath.dirname(path).removesuffix('/_rels')
        for node in root:
            if node.get('TargetMode') == 'External':
                continue
            target = node.get('Target')
            resolved = posixpath.normpath(posixpath.join(owner, target)).lstrip('/')
            assert resolved in names, (path, target)
        if path != '_rels/.rels':
            source = posixpath.join(owner, posixpath.basename(path)[:-5])
            if source.endswith('.xml'):
                for node in ET.fromstring(archive.read(source)).iter():
                    for key, value in node.attrib.items():
                        if key.startswith('{'+R+'}'):
                            assert value in ids, (source, key, value)
    types = ET.fromstring(archive.read('[Content_Types].xml'))
    overrides = [node.get('PartName') for node in types if node.tag == '{'+C+'}Override']
    assert len(overrides) == len(set(overrides))
    defaults = {node.get('Extension') for node in types if node.tag == '{'+C+'}Default'}
    for name in names:
        if name == '[Content_Types].xml':
            continue
        assert '/'+name in overrides or name.rsplit('.', 1)[-1] in defaults, name


def test_writer_preserves_model_and_original_members(tmp_path):
    model = _model()
    source = model.save(tmp_path / 'source.xlsx')
    output = tmp_path / 'report.xlsx'
    add_reports(source, output, (_pivot(),), (), model=model)
    changed = {'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'}
    with ZipFile(source) as before, ZipFile(output) as after:
        assert before.read('xl/model/item.data') == after.read('xl/model/item.data')
        for item in before.infolist():
            if item.filename not in changed:
                assert before.read(item) == after.read(item.filename), item.filename
            rewritten = after.getinfo(item.filename)
            assert rewritten.date_time == item.date_time
            assert rewritten.compress_type == item.compress_type
        _assert_links(after)


def test_uses_actual_model_connection_id(tmp_path):
    model = _model()
    source = model.save(tmp_path / 'source.xlsx')
    altered = tmp_path / 'connection42.xlsx'
    with ZipFile(source) as before, ZipFile(altered, 'w') as after:
        for item in before.infolist():
            payload = before.read(item)
            if item.filename == 'xl/connections.xml':
                root = ET.fromstring(payload)
                target = next(node for node in root if node.get('name') == 'ThisWorkbookDataModel')
                target.set('id', '42')
                payload = ET.tostring(root)
            after.writestr(item, payload)
    output = tmp_path / 'report.xlsx'
    add_reports(altered, output, (_pivot(),), (), model=model)
    with ZipFile(output) as archive:
        caches = [name for name in archive.namelist() if name.startswith('xl/pivotCache/') and name.endswith('.xml')]
        assert len(caches) == 1
        assert ET.fromstring(archive.read(caches[0])).find('{'+S+'}cacheSource').get('connectionId') == '42'
        _assert_links(archive)


def test_multiple_pivots_charts_and_shared_chart_sheet(tmp_path):
    model = _model()
    report = WorkbookReport(model).add_pivot_table(_pivot()).add_pivot_table(_pivot('Second', 'Second Report'))
    report.add_pivot_chart(PivotChartSpec('FirstChart', 'Summary'))
    report.add_pivot_chart(PivotChartSpec('SecondChart', 'Second', sheet='Charts', chart_type='line'))
    report.add_pivot_chart(PivotChartSpec('ThirdChart', 'Summary', sheet='Charts', cell='G25', chart_type='bar'))
    output = report.save(tmp_path / 'report.xlsx')
    with ZipFile(output) as archive:
        _assert_links(archive)
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        names = [node.get('name') for node in workbook.find('{'+S+'}sheets')]
        assert all(names.count(name) == 1 for name in ('Report', 'Second Report', 'Charts'))
        caches = workbook.find('{'+S+'}pivotCaches')
        assert len(caches) == 2
        assert len({node.get('cacheId') for node in caches}) == 2
        pivots = [ET.fromstring(archive.read(path)) for path in archive.namelist()
                  if path.startswith('xl/pivotTables/') and path.endswith('.xml')]
        assert {node.get('name') for node in pivots} == {'Summary', 'Second'}
        assert {node.get('cacheId') for node in pivots} == {node.get('cacheId') for node in caches}
        assert len([path for path in archive.namelist() if path.startswith('xl/charts/chart') and path.endswith('.xml')]) == 3


def test_separate_caches_preserve_each_report_axis_indexes(tmp_path):
    model = _model()
    model.add_measure('Table1', 'CountID', 'COUNT(Table1[ID])')
    first = _pivot()
    second = PivotTableSpec('ByID', 'IDs', rows=(('Table1', 'ID'),),
                            columns=(('Table1', 'Label'),), measures=('CountID', 'TotalID'))
    output = WorkbookReport(model).add_pivot_table(first).add_pivot_table(second).save(tmp_path / 'union.xlsx')
    with ZipFile(output) as archive:
        _assert_links(archive)
        cache_names = [p for p in archive.namelist() if p.startswith('xl/pivotCache/') and p.endswith('.xml')]
        assert len(cache_names) == 2
        pivots = {p.get('name'): p for p in (ET.fromstring(archive.read(n)) for n in archive.namelist()
                  if n.startswith('xl/pivotTables/') and n.endswith('.xml'))}
        fields_by_name = {}
        for path in archive.namelist():
            if not path.startswith('xl/pivotTables/') or not path.endswith('.xml'):
                continue
            pivot = ET.fromstring(archive.read(path))
            relpath = posixpath.dirname(path)+'/_rels/'+posixpath.basename(path)+'.rels'
            target = ET.fromstring(archive.read(relpath))[0].get('Target')
            cachepath = posixpath.normpath(posixpath.join(posixpath.dirname(path), target))
            cache = ET.fromstring(archive.read(cachepath))
            fields_by_name[pivot.get('name')] = [f.get('name') for f in cache.find('{'+S+'}cacheFields')]
        for name, row in [('Summary', '[Table1].[Label].[Label]'), ('ByID', '[Table1].[ID].[ID]')]:
            pivot = pivots[name]
            field_names = fields_by_name[name]
            assert len(field_names) == len(set(field_names))
            assert len(pivot.find('{'+S+'}pivotFields')) == len(field_names)
            assert field_names[int(pivot.find('{'+S+'}rowFields')[0].get('x'))] == row
        second_pivot = pivots['ByID']
        field_names = fields_by_name['ByID']
        assert [field_names[int(f.get('fld'))] for f in second_pivot.find('{'+S+'}dataFields')] == [
            '[Measures].[CountID]', '[Measures].[TotalID]']
        assert [f.get('x') for f in second_pivot.find('{'+S+'}colFields')][-1] == '-2'
        assert second_pivot.find('{'+S+'}colHierarchiesUsage')[-1].get('hierarchyUsage') == '-2'
        first_measure = pivots['Summary'].find('{'+S+'}dataFields')[0]
        assert fields_by_name['Summary'][int(first_measure.get('fld'))] == '[Measures].[TotalID]'


@pytest.mark.parametrize('conflict', ['sheet', 'pivot'])
def test_existing_names_fail_atomically(tmp_path, conflict):
    initial = WorkbookReport(_model()).add_pivot_table(_pivot()).save(tmp_path / 'initial.xlsx')
    report = WorkbookReport(initial)
    target = tmp_path / 'destination.xlsx'
    target.write_bytes(b'keep this file')
    if conflict == 'sheet':
        report.add_pivot_table(_pivot('Other','Report'))
    before = report.list_reports()
    with pytest.raises(ModelValidationError):
        if conflict == 'pivot':
            report.add_pivot_table(_pivot('Summary','Other Report'))
        else:
            report.save(target)
    assert target.read_bytes() == b'keep this file'
    assert report.list_reports() == before
    assert not list(tmp_path.glob('.excelmodel-report-*'))


def test_model_rename_after_staging_rejected(tmp_path):
    model = _model()
    report = WorkbookReport(model).add_pivot_table(_pivot())
    model.rename_table('Table1', 'Renamed')
    with pytest.raises(ModelValidationError, match='field'):
        report.save(tmp_path / 'invalid.xlsx')
    assert not (tmp_path / 'invalid.xlsx').exists()


def test_empty_report_batch_preserves_all_package_parts(tmp_path):
    model = _model()
    source = model.save(tmp_path / 'source.xlsx')
    output = tmp_path / 'empty.xlsx'
    add_reports(source, output, (), (), model=model)
    with ZipFile(source) as before, ZipFile(output) as after:
        assert before.namelist() == after.namelist()
        for name in before.namelist():
            assert before.read(name) == after.read(name), name
        assert before.comment == after.comment


@pytest.mark.parametrize('layout,compact,outline', [
    ('tabular', '0', '0'), ('compact', '1', '1'), ('outline', '0', '1'),
])
def test_layout_axes_totals_and_multiple_measure_indexes(tmp_path, layout, compact, outline):
    model = _model()
    model.add_measure('Table1', 'CountID', 'COUNT(Table1[ID])')
    spec = PivotTableSpec('LayoutSummary', 'Layout', rows=(('Table1', 'Label'),),
                          filters=(('Table1', 'ID'),), measures=('TotalID', 'CountID'),
                          layout=layout, row_grand_totals=False, column_grand_totals=True)
    output = WorkbookReport(model).add_pivot_table(spec).save(tmp_path / 'layout.xlsx')
    with ZipFile(output) as archive:
        _assert_links(archive)
        pivot = ET.fromstring(archive.read('xl/pivotTables/pivotTable1.xml'))
        cache = ET.fromstring(archive.read('xl/pivotCache/pivotCacheDefinition1.xml'))
        assert pivot.get('compact') == compact
        assert pivot.get('compactData') == compact
        assert pivot.get('rowGrandTotals') == '0'
        assert pivot.get('colGrandTotals') == '1'
        fields = pivot.find('{'+S+'}pivotFields')
        assert [field.get('axis') for field in fields] == ['axisRow', 'axisPage', None, None]
        assert all(field.get('compact') == compact and field.get('outline') == outline for field in fields)
        assert [field.get('x') for field in pivot.find('{'+S+'}colFields')] == ['-2']
        assert [field.get('fld') for field in pivot.find('{'+S+'}dataFields')] == ['2', '3']
        page = pivot.find('{'+S+'}pageFields')[0]
        cache_fields = cache.find('{'+S+'}cacheFields')
        assert page.get('fld') == '1'
        assert page.get('hier') == cache_fields[1].get('hierarchy')
        # Native refresh needs an explicit All-member selection even before
        # cache members exist; otherwise a filtered chart can crash Excel.
        assert page.get('name') == '[Table1].[ID].[All]'
        assert page.get('cap') == 'All'
        items = fields[1].find('{'+S+'}items')
        assert items.get('count') == '1' and len(items) == 1
        assert items[0].get('t') == 'default'
        location = pivot.find('{'+S+'}location')
        assert location.get('rowPageCount') == '1'
        assert location.get('colPageCount') == '1'
        assert cache_fields[2].get('name') == '[Measures].[TotalID]'
        assert cache_fields[3].get('name') == '[Measures].[CountID]'


def test_chart_binding_uses_final_filename_and_exact_sheet_name(tmp_path):
    sheet = "Director's Sales & Costs"
    report = WorkbookReport(_model()).add_pivot_table(_pivot(sheet=sheet))
    report.add_pivot_chart(PivotChartSpec('SalesChart', 'Summary', sheet='Chart Display'))
    output = report.save(tmp_path / 'Final Sales & Costs.xlsx')
    with ZipFile(output) as archive:
        chart = ET.fromstring(archive.read('xl/charts/chart1.xml'))
        ns = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
        assert chart.find(ns+'pivotSource/'+ns+'name').text == f'[{output.name}]{sheet}!Summary'
        assert b'.excelmodel-report-' not in archive.read('xl/charts/chart1.xml')
        _assert_links(archive)


def test_saved_report_still_guards_model_field_removal(tmp_path):
    output = WorkbookReport(_model()).add_pivot_table(_pivot()).save(tmp_path / 'guard.xlsx')
    model = WorkbookModel(output)
    from excel_model_lab.abf import UnsupportedModelError
    with pytest.raises(UnsupportedModelError, match='pivotTable'):
        model.remove_measure('Table1', 'TotalID')


def test_appended_parts_remain_namespaced_with_prefixed_source_xml(tmp_path):
    model = _model()
    source = model.save(tmp_path / 'source.xlsx')
    prefixed = tmp_path / 'prefixed.xlsx'
    roots = {'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'}
    with ZipFile(source) as before, ZipFile(prefixed, 'w') as after:
        for item in before.infolist():
            payload = before.read(item)
            if item.filename in roots:
                # Namespace prefixes are arbitrary in OOXML; a default namespace
                # is not required on any of these source roots.
                doc = minidom.parseString(payload)
                namespace = doc.documentElement.namespaceURI
                for node in list(doc.getElementsByTagNameNS(namespace, '*')):
                    doc.renameNode(node, namespace, 'source:'+node.localName)
                doc.documentElement.setAttribute('xmlns:source', namespace)
                doc.documentElement.removeAttribute('xmlns')
                if item.filename == 'xl/workbook.xml':
                    for node in doc.getElementsByTagName('*'):
                        for attr in list(node.attributes.values()):
                            if attr.namespaceURI == R:
                                doc.renameNode(attr, R, 'link:'+attr.localName)
                    doc.documentElement.setAttribute('xmlns:link', R)
                    doc.documentElement.removeAttribute('xmlns:r')
                payload = doc.toxml(encoding='utf-8')
            after.writestr(item, payload)
    output = tmp_path / 'prefixed-report.xlsx'
    add_reports(prefixed, output, (_pivot(),), (), model=model)
    with ZipFile(output) as archive:
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        sheets = workbook.find('{'+S+'}sheets')
        assert all(node.tag == '{'+S+'}sheet' for node in sheets)
        caches = workbook.find('{'+S+'}pivotCaches')
        assert caches is not None
        assert all(node.tag == '{'+S+'}pivotCache' for node in caches)
        assert all(node.tag == '{'+P+'}Relationship'
                   for node in ET.fromstring(archive.read('xl/_rels/workbook.xml.rels')))
        assert all(node.tag in ('{'+C+'}Override', '{'+C+'}Default')
                   for node in ET.fromstring(archive.read('[Content_Types].xml')))
        _assert_links(archive)
