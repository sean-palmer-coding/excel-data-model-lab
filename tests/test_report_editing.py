from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import pytest

from excel_model_lab import WorkbookModel, WorkbookReport, PivotTableSpec, PivotChartSpec, ModelValidationError
from excel_model_lab.report_editing import inspect_reports, edit_reports, C

FIXTURE = Path(__file__).parent / 'data' / 'native-authoring-base.xlsx'


@pytest.fixture
def workbook(tmp_path):
    model = WorkbookModel(FIXTURE)
    model.add_measure('Table1', 'TotalID', 'SUM(Table1[ID])')
    report = WorkbookReport(model)
    report.add_pivot_table(PivotTableSpec('Summary', 'Report', rows=(('Table1', 'Label'),), measures=('TotalID',)))
    report.add_pivot_chart(PivotChartSpec('Chart', 'Summary', title='Original'))
    return report.save(tmp_path / 'original.xlsx')


def test_inspects_generated_specs(workbook):
    reports = inspect_reports(workbook)
    assert len(reports) == 2
    pivot, chart = reports
    assert not pivot.unsupported and not chart.unsupported
    assert pivot.spec.rows == (('Table1', 'Label'),)
    assert pivot.spec.measures == ('TotalID',)
    assert pivot.spec.layout == 'tabular'
    assert chart.spec.title == 'Original'
    assert chart.spec.pivot_table == 'Summary'


def test_surgical_edits_preserve_model_and_unrelated_parts(workbook, tmp_path):
    output = edit_reports(workbook, tmp_path/'edited.xlsx',
        pivot_updates={'Summary': {'name': 'Renamed', 'layout': 'compact', 'row_grand_totals': False}},
        chart_updates={'Chart': {'name': 'NewChart', 'title': 'Updated'}}, workbook_name='edited.xlsx')
    reports = inspect_reports(output)
    assert reports[0].spec.name == 'Renamed'
    assert reports[0].spec.layout == 'compact'
    assert reports[0].spec.row_grand_totals is False
    assert reports[1].spec.name == 'NewChart' and reports[1].spec.title == 'Updated'
    assert reports[1].spec.pivot_table == 'Renamed'
    with ZipFile(workbook) as before, ZipFile(output) as after:
        for name in before.namelist():
            if name not in {reports[0].part, reports[1].part, 'xl/drawings/drawing1.xml'}:
                assert before.read(name) == after.read(name), name
        source = ET.fromstring(after.read(reports[1].part)).find('{'+C+'}pivotSource/{'+C+'}name').text
        assert source == '[edited.xlsx]Report!Renamed'


def test_remove_dependencies_atomic_and_keep_sheet_data(workbook, tmp_path):
    target = tmp_path/'out.xlsx'
    target.write_bytes(b'unchanged')
    with pytest.raises(ModelValidationError, match='dependent'):
        edit_reports(workbook, target, remove=('Summary',))
    assert target.read_bytes() == b'unchanged'
    edit_reports(workbook, target, remove=('Chart', 'Summary'))
    assert inspect_reports(target) == ()
    with ZipFile(workbook) as before, ZipFile(target) as after:
        assert before.read('xl/model/item.data') == after.read('xl/model/item.data')
        assert before.read('xl/worksheets/sheet1.xml') == after.read('xl/worksheets/sheet1.xml')
        assert 'xl/worksheets/sheet2.xml' in after.namelist()


@pytest.mark.parametrize('kwargs', [
    {'pivot_updates': {'Missing': {'layout': 'compact'}}},
    {'pivot_updates': {'Summary': {'layout': 'invalid'}}},
    {'chart_updates': {'Chart': {'title': '\x00'}}},
    {'chart_updates': {'Chart': {'chart_type': 'pie'}}},
    {'pivot_updates': {'Summary': {'name': 'Chart'}}},
])
def test_invalid_edits_leave_destination_unchanged(workbook, tmp_path, kwargs):
    target = tmp_path/'unchanged.xlsx'
    target.write_bytes(b'original')
    with pytest.raises(ModelValidationError):
        edit_reports(workbook, target, **kwargs)
    assert target.read_bytes() == b'original'


def test_guarded_axis_regeneration(workbook, tmp_path):
    model = WorkbookModel(workbook)
    output = edit_reports(workbook, tmp_path/'axes.xlsx',
        pivot_updates={'Summary': {'rows': (), 'columns': (('Table1', 'ID'),)}}, model=model)
    pivot = inspect_reports(output)[0].spec
    assert pivot.rows == () and pivot.columns == (('Table1', 'ID'),)
    assert pivot.measures == ('TotalID',)
    with ZipFile(workbook) as before, ZipFile(output) as after:
        assert before.read('xl/model/item.data') == after.read('xl/model/item.data')
        assert before.read('xl/worksheets/sheet1.xml') == after.read('xl/worksheets/sheet1.xml')


def test_unknown_extension_preserved_by_surgical_edit_but_blocks_regeneration(workbook, tmp_path):
    altered = tmp_path/'custom.xlsx'
    part = inspect_reports(workbook)[0].part
    with ZipFile(workbook) as before, ZipFile(altered, 'w') as after:
        for info in before.infolist():
            payload = before.read(info)
            if info.filename == part:
                root = ET.fromstring(payload)
                root.set('{urn:vendor}custom', 'keep')
                payload = ET.tostring(root)
            after.writestr(info, payload)
    output = edit_reports(altered, tmp_path/'surgical.xlsx', pivot_updates={'Summary': {'row_grand_totals': False}})
    with ZipFile(output) as archive:
        assert ET.fromstring(archive.read(part)).get('{urn:vendor}custom') == 'keep'
    target = tmp_path/'rejected.xlsx'
    target.write_bytes(b'original')
    with pytest.raises(ModelValidationError, match='metadata'):
        edit_reports(altered, target, pivot_updates={'Summary': {'rows': (), 'columns': (('Table1', 'ID'),)}},
                     model=WorkbookModel(altered))
    assert target.read_bytes() == b'original'


def test_add_title_to_untitled_chart(tmp_path):
    model = WorkbookModel(FIXTURE)
    report = WorkbookReport(model).add_pivot_table(PivotTableSpec('Summary', 'Report', rows=(('Table1', 'ID'),)))
    report.add_pivot_chart(PivotChartSpec('Chart', 'Summary'))
    source = report.save(tmp_path/'no-title.xlsx')
    output = edit_reports(source, tmp_path/'title.xlsx', chart_updates={'Chart': {'title': 'Added & retained'}})
    assert inspect_reports(output)[1].spec.title == 'Added & retained'


def test_axis_change_requires_model_and_unknown_key_rejected(workbook, tmp_path):
    with pytest.raises(ModelValidationError, match='current model'):
        edit_reports(workbook, tmp_path/'missing.xlsx', pivot_updates={'Summary': {'rows': ()}})
    with pytest.raises(ModelValidationError, match='Unsupported'):
        edit_reports(workbook, tmp_path/'unknown.xlsx', pivot_updates={'Summary': {'invented': 1}})
    assert not (tmp_path/'missing.xlsx').exists()
    assert not (tmp_path/'unknown.xlsx').exists()


def test_native_relationship_only_pivot_is_inspected_and_removed(workbook, tmp_path):
    source = tmp_path/'relationship-only.xlsx'
    namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    with ZipFile(workbook) as before, ZipFile(source, 'w') as after:
        for item in before.infolist():
            payload = before.read(item)
            if item.filename.startswith('xl/worksheets/') and item.filename.endswith('.xml'):
                root = ET.fromstring(payload)
                for node in root.findall('{'+namespace+'}pivotTableParts'):
                    root.remove(node)
                payload = ET.tostring(root)
            after.writestr(item, payload)
    assert [r.name for r in inspect_reports(source)] == ['Summary', 'Chart']
    result = edit_reports(source, tmp_path/'removed.xlsx', remove=('Summary', 'Chart'))
    assert inspect_reports(result) == ()


def test_member_selections_and_sort_survive_inspection_and_regeneration(tmp_path):
    from excel_model_lab.report_filters import PivotMemberFilter, PivotSort
    model = WorkbookModel(FIXTURE)
    labels = model.get_table('Table1')['Label'].tolist()
    spec = PivotTableSpec('Summary', 'Report', rows=(('Table1','Label'),), filters=(('Table1','ID'),),
        member_filters=(PivotMemberFilter('Table1','Label',(labels[0],)), PivotMemberFilter('Table1','ID',(1,))),
        sorts=(PivotSort('Table1','Label','descending'),))
    source = WorkbookReport(model).add_pivot_table(spec).save(tmp_path/'selected.xlsx')
    parsed = inspect_reports(source)[0].spec
    assert parsed.member_filters == spec.member_filters
    assert parsed.sorts == spec.sorts
    output = edit_reports(source,tmp_path/'regenerated.xlsx',pivot_updates={'Summary':parsed},model=WorkbookModel(source))
    assert inspect_reports(output)[0].spec.member_filters == spec.member_filters
    assert inspect_reports(output)[0].spec.sorts == spec.sorts


def test_chart_formatting_surgical_roundtrip(workbook,tmp_path):
    options = dict(style=8,legend_position='bottom',category_axis_title='Categories',value_axis_title='Amount',
        number_format='#,##0.00',show_data_labels=True,background_color='ABCDEF',plot_background_color='123456',chart_type='bar')
    output=edit_reports(workbook,tmp_path/'format.xlsx',chart_updates={'Chart':options})
    chart=inspect_reports(output)[1].spec
    for key,value in options.items():
        assert getattr(chart,key)==value
    with ZipFile(workbook) as before,ZipFile(output) as after:
        changed=inspect_reports(output)[1].part
        for name in before.namelist():
            if name!=changed:
                assert before.read(name)==after.read(name)


def test_chart_type_conversion_preserves_cached_series(workbook,tmp_path):
    source=tmp_path/'cached.xlsx'
    part=inspect_reports(workbook)[1].part
    with ZipFile(workbook) as before,ZipFile(source,'w') as after:
        for item in before.infolist():
            payload=before.read(item)
            if item.filename==part:
                root=ET.fromstring(payload)
                chart=root.find('.//{'+C+'}barChart')
                series=ET.Element('{'+C+'}ser')
                ET.SubElement(series,'{'+C+'}idx',val='0')
                ET.SubElement(series,'{'+C+'}order',val='0')
                values=ET.SubElement(ET.SubElement(series,'{'+C+'}val'),'{'+C+'}numLit')
                ET.SubElement(values,'{'+C+'}ptCount',val='1')
                point=ET.SubElement(values,'{'+C+'}pt',idx='0')
                ET.SubElement(point,'{'+C+'}v').text='123'
                chart.insert(3,series)
                payload=ET.tostring(root)
            after.writestr(item,payload)
    output=edit_reports(source,tmp_path/'line.xlsx',chart_updates={'Chart':{'chart_type':'line'}})
    with ZipFile(output) as archive:
        root=ET.fromstring(archive.read(part))
        assert root.find('.//{'+C+'}lineChart/{'+C+'}ser/{'+C+'}val/{'+C+'}numLit/{'+C+'}pt/{'+C+'}v').text=='123'
    assert inspect_reports(output)[1].spec.chart_type=='line'


@pytest.mark.parametrize('axis', ['rows','filters'])
def test_typed_boolean_date_and_blank_selection_inspection(tmp_path,axis):
    from datetime import datetime
    import pandas as pd
    from excel_model_lab.report_filters import PivotMemberFilter
    model=WorkbookModel.create({'Data':pd.DataFrame({'Flag':[True,False,True],
        'When':pd.to_datetime(['2026-01-01','2026-01-02','2026-01-03']), 'Label':['A',None,'B']})})
    for column,value in [('Flag',True),('When',datetime(2026,1,1)),('Label',None)]:
        spec=PivotTableSpec('Summary','Report',**{axis:(('Data',column),)},
            member_filters=(PivotMemberFilter('Data',column,(value,)),))
        source=WorkbookReport(model).add_pivot_table(spec).save(tmp_path/(column+'.xlsx'))
        parsed=inspect_reports(source)[0]
        assert not parsed.unsupported
        assert parsed.spec.member_filters==spec.member_filters


def test_control_inspection_reads_view_cache_and_pivot_bindings(workbook,tmp_path):
    from excel_model_lab.reports import SlicerSpec
    from excel_model_lab.report_editing import inspect_controls
    source=WorkbookReport(workbook).add_slicer(SlicerSpec('Labels','Summary',('Table1','Label'),sheet='Report')).save(tmp_path/'slicer.xlsx')
    controls=inspect_controls(source)
    assert len(controls)==1
    control=controls[0]
    assert (control.kind,control.name,control.sheet)==('slicer','Labels','Report')
    assert control.source_name=='[Table1].[Label]'
    assert control.connected_pivots[0][1]=='Summary'


def test_existing_control_update_preserves_selection_cache(workbook,tmp_path):
    from excel_model_lab.reports import SlicerSpec
    from excel_model_lab.report_editing import inspect_controls
    source=WorkbookReport(workbook).add_slicer(SlicerSpec('Labels','Summary',('Table1','Label'),sheet='Report')).save(tmp_path/'slicer.xlsx')
    original=inspect_controls(source)[0]
    output=edit_reports(source,tmp_path/'updated.xlsx',control_updates={'Labels':{
        'name':'RenamedLabels','caption':'New caption','columns':2,'cell':'L5','width':240,'height':300}})
    control=inspect_controls(output)[0]
    assert control.name=='RenamedLabels' and control.caption=='New caption'
    with ZipFile(source) as before,ZipFile(output) as after:
        assert before.read(original.cache_part)==after.read(original.cache_part)
        assert before.read('xl/model/item.data')==after.read('xl/model/item.data')
        view=ET.fromstring(after.read(control.part))[0]
        assert view.get('columnCount')=='2'


def test_control_remove_preserves_other_report_and_rejects_pivot_dependency(workbook,tmp_path):
    from excel_model_lab.reports import SlicerSpec
    from excel_model_lab.report_editing import inspect_controls
    source=WorkbookReport(workbook).add_slicer(SlicerSpec('Labels','Summary',('Table1','Label'),sheet='Report')).save(tmp_path/'slicer.xlsx')
    target=tmp_path/'out.xlsx';target.write_bytes(b'original')
    with pytest.raises(ModelValidationError,match='dependent controls'):
        edit_reports(source,target,remove=('Summary','Chart'))
    assert target.read_bytes()==b'original'
    edit_reports(source,target,remove=('Labels',))
    assert inspect_controls(target)==()
    assert [r.name for r in inspect_reports(target)]==['Summary','Chart']
    with ZipFile(source) as before,ZipFile(target) as after:
        assert before.read('xl/model/item.data')==after.read('xl/model/item.data')


def test_control_connections_and_pivot_rename(workbook,tmp_path):
    from excel_model_lab.reports import SlicerSpec
    from excel_model_lab.report_editing import inspect_controls
    report=WorkbookReport(workbook).add_pivot_table(PivotTableSpec('Other','Other Report',rows=(('Table1','ID'),)))
    report.add_slicer(SlicerSpec('Labels','Summary',('Table1','Label'),sheet='Report'))
    source=report.save(tmp_path/'slicer.xlsx')
    output=edit_reports(source,tmp_path/'connected.xlsx',control_updates={'Labels':{'connected_pivots':('Summary','Other')}})
    assert [name for _,name in inspect_controls(output)[0].connected_pivots]==['Summary','Other']
    renamed=edit_reports(output,tmp_path/'renamed.xlsx',pivot_updates={'Summary':{'name':'Renamed'}})
    assert [name for _,name in inspect_controls(renamed)[0].connected_pivots]==['Renamed','Other']
    before=renamed.read_bytes()
    with pytest.raises(ModelValidationError,match='Unknown connected'):
        edit_reports(renamed,renamed,control_updates={'Labels':{'connected_pivots':('Missing',)}})
    assert renamed.read_bytes()==before


def test_two_cell_control_conversion_preserves_drawing_extensions(workbook,tmp_path):
    from excel_model_lab.reports import SlicerSpec
    from excel_model_lab.report_editing import X
    source=WorkbookReport(workbook).add_slicer(SlicerSpec('Labels','Summary',('Table1','Label'),sheet='Report')).save(tmp_path/'slicer.xlsx')
    altered=tmp_path/'two-cell.xlsx'
    control_drawing=None
    with ZipFile(source) as before,ZipFile(altered,'w') as after:
        for item in before.infolist():
            payload=before.read(item)
            if item.filename.startswith('xl/drawings/') and item.filename.endswith('.xml'):
                root=ET.fromstring(payload)
                for anchor in root:
                    if any(n.get('name')=='Labels' for n in anchor.iter()):
                        anchor.tag='{'+X+'}twoCellAnchor'
                        extent=anchor.find('{'+X+'}ext');anchor.remove(extent)
                        to=ET.Element('{'+X+'}to')
                        for tag,value in [('col',15),('colOff',0),('row',25),('rowOff',0)]:
                            ET.SubElement(to,'{'+X+'}'+tag).text=str(value)
                        anchor.insert(1,to)
                        ET.SubElement(anchor,'{urn:custom}metadata',keep='yes')
                        control_drawing=item.filename
                payload=ET.tostring(root)
            after.writestr(item,payload)
    with pytest.raises(ModelValidationError,match='explicit width and height'):
        edit_reports(altered,tmp_path/'invalid.xlsx',control_updates={'Labels':{'cell':'K3'}})
    output=edit_reports(altered,tmp_path/'one-cell.xlsx',control_updates={'Labels':{'cell':'K3','width':200,'height':300}})
    with ZipFile(output) as archive:
        root=ET.fromstring(archive.read(control_drawing))
        anchor=next(a for a in root if any(n.get('name')=='Labels' for n in a.iter()))
        assert anchor.tag=='{'+X+'}oneCellAnchor'
        assert anchor.find('{urn:custom}metadata').get('keep')=='yes'
        assert anchor.find('{'+X+'}ext').get('cx')==str(200*9525)
