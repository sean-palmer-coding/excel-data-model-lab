"""Public facade regressions for appending drawings and editing controls."""
import posixpath
from zipfile import ZipFile
from xml.dom import minidom
from xml.etree import ElementTree as ET
import pandas as pd
import pytest
from excel_model_lab import (WorkbookModel, WorkbookReport, PivotTableSpec,
    PivotChartSpec, SlicerSpec, TimelineSpec, ModelValidationError)
from excel_model_lab.report_editing import inspect_reports, inspect_controls

XDR='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
C='http://schemas.openxmlformats.org/drawingml/2006/chart'
MC='http://schemas.openxmlformats.org/markup-compatibility/2006'


def staged_report():
    model=WorkbookModel.create({'Sales':pd.DataFrame({'Category':['A','B'],
        'Amount':[10,20],'Day':pd.to_datetime(['2026-01-01','2026-01-02'])})})
    model.add_measure('Sales','Total','SUM(Sales[Amount])')
    return (WorkbookReport(model)
        .add_pivot_table(PivotTableSpec('Summary','Report',rows=(('Sales','Category'),),measures=('Total',)))
        .add_pivot_chart(PivotChartSpec('OriginalChart','Summary',title='Original'))
        .add_slicer(SlicerSpec('Categories','Summary',('Sales','Category')))
        .add_timeline(TimelineSpec('Dates','Summary',('Sales','Day'))))


@pytest.fixture
def controlled(tmp_path):
    return staged_report().save(tmp_path/'source.xlsx')


def assert_namespace_bindings(parts):
    for path,data in parts.items():
        if not path.endswith(('.xml','.rels')):continue
        doc=minidom.parseString(data)
        for node in doc.getElementsByTagName('*'):
            for prefix in node.getAttributeNS(MC,'Ignorable').split():
                current=node
                while current is not None and getattr(current,'nodeType',None)==current.ELEMENT_NODE:
                    if current.hasAttribute('xmlns:'+prefix):break
                    current=current.parentNode
                else:pytest.fail(f'{path}: undefined Ignorable prefix {prefix}')


def test_append_chart_existing_sheet_retains_controls_and_old_chart(controlled,tmp_path):
    output=(WorkbookReport(controlled)
        .add_pivot_chart(PivotChartSpec('AppendedChart','Summary',sheet='Report',cell='A25',chart_type='line',title='Added'))
        .save(tmp_path/'appended.xlsx'))
    reports={r.name:r for r in inspect_reports(output)}
    assert set(reports)=={'Summary','OriginalChart','AppendedChart'}
    assert reports['AppendedChart'].sheet=='Report'
    assert reports['AppendedChart'].spec.pivot_table=='Summary'
    assert {c.name for c in inspect_controls(output)}=={'Categories','Dates'}
    with ZipFile(controlled) as before,ZipFile(output) as after:
        parts={n:after.read(n) for n in after.namelist()}
        assert before.read('xl/model/item.data')==parts['xl/model/item.data']
        for control in inspect_controls(controlled):
            assert before.read(control.part)==parts[control.part]
            assert before.read(control.cache_part)==parts[control.cache_part]
        old_chart=ET.fromstring(before.read(reports['OriginalChart'].part))
        new_chart=ET.fromstring(parts[reports['OriginalChart'].part])
        old_chart.find('{'+C+'}pivotSource/{'+C+'}name').text=new_chart.find('{'+C+'}pivotSource/{'+C+'}name').text
        assert ET.tostring(old_chart)==ET.tostring(new_chart)
        drawings=[p for p in parts if p.startswith('xl/drawings/drawing') and p.endswith('.xml')]
        assert len(drawings)==1
        root=ET.fromstring(parts[drawings[0]])
        names=[n.get('name') for n in root.iter('{'+XDR+'}cNvPr')]
        assert set(names)=={'OriginalChart','AppendedChart','Categories','Dates'}
        ids=[n.get('id') for n in root.iter('{'+XDR+'}cNvPr')]
        assert len(set(ids))==len(ids)
        for p,data in parts.items():
            if not p.endswith('.rels'):continue
            owner='' if p=='_rels/.rels' else posixpath.dirname(p).removesuffix('/_rels')
            rels=ET.fromstring(data)
            assert len({r.get('Id') for r in rels})==len(rels)
            for r in rels:
                if r.get('TargetMode')!='External':
                    assert posixpath.normpath(posixpath.join(owner,r.get('Target'))).lstrip('/') in parts
        assert_namespace_bindings(parts)


def test_existing_control_update_wrappers_preserve_caches(controlled,tmp_path):
    report=WorkbookReport(controlled)
    assert report.update_slicer('Categories',caption='Choose category',columns=2,cell='L5') is report
    assert report.update_timeline('Dates',caption='Choose date',level='years',cell='L20') is report
    output=report.save(tmp_path/'updated.xlsx')
    controls={c.name:c for c in inspect_controls(output)}
    assert controls['Categories'].caption=='Choose category'
    assert controls['Dates'].caption=='Choose date'
    with ZipFile(controlled) as before,ZipFile(output) as after:
        for control in inspect_controls(controlled):assert before.read(control.cache_part)==after.read(control.cache_part)
        slicer=ET.fromstring(after.read(controls['Categories'].part))[0]
        timeline=ET.fromstring(after.read(controls['Dates'].part))[0]
        assert slicer.get('columnCount')=='2'
        assert timeline.get('level')=='0'


def test_staged_control_update_and_remove_wrappers_are_transactional(tmp_path):
    report=staged_report()
    assert report.update_slicer('Categories',caption='Updated',columns=2) is report
    assert report.update_timeline('Dates',caption='Updated dates',level='days') is report
    before=report.list_reports()
    with pytest.raises(ModelValidationError):report.update_slicer('Categories',columns=0)
    assert report.list_reports()==before
    with pytest.raises(ModelValidationError,match='dependent'):report.remove_report('Summary')
    assert report.list_reports()==before
    for name in ('Categories','Dates','OriginalChart','Summary'):
        assert report.remove_report(name) is report
    assert report.list_reports()==()


def test_existing_control_remove_wrappers_keep_other_objects(controlled,tmp_path):
    report=WorkbookReport(controlled)
    with pytest.raises(ModelValidationError,match='dependent'):report.remove_report('Summary')
    assert report.remove_report('Categories') is report
    output=report.save(tmp_path/'one-control.xlsx')
    assert [c.name for c in inspect_controls(output)]==['Dates']
    assert {r.name for r in inspect_reports(output)}=={'Summary','OriginalChart'}
    report=WorkbookReport(output)
    for name in ('Dates','OriginalChart','Summary'):report.remove_report(name)
    empty=report.save(tmp_path/'no-reports.xlsx')
    assert inspect_reports(empty)==() and inspect_controls(empty)==()


def test_append_charts_case_variant_existing_sheet_keeps_both(controlled,tmp_path):
    report=WorkbookReport(controlled)
    report.add_pivot_chart(PivotChartSpec('FirstAdded','Summary',sheet='Report',cell='A25'))
    report.add_pivot_chart(PivotChartSpec('SecondAdded','Summary',sheet='report',cell='H25'))
    output=report.save(tmp_path/'case-variant.xlsx')
    assert {r.name for r in inspect_reports(output)}=={'Summary','OriginalChart','FirstAdded','SecondAdded'}


def test_append_chart_and_slicer_share_existing_sheet_drawing(controlled,tmp_path):
    report=WorkbookReport(controlled)
    report.add_pivot_chart(PivotChartSpec('NewChart','Summary',sheet='Report',cell='A25'))
    report.add_slicer(SlicerSpec('AnotherCategory','Summary',('Sales','Category'),sheet='Report',cell='P3'))
    output=report.save(tmp_path/'chart-and-slicer.xlsx')
    assert {r.name for r in inspect_reports(output)}=={'Summary','OriginalChart','NewChart'}
    assert {r.name for r in inspect_controls(output)}=={'Categories','Dates','AnotherCategory'}
    with ZipFile(output) as z:
        parts={n:z.read(n) for n in z.namelist()}
        drawing=ET.fromstring(parts['xl/drawings/drawing1.xml'])
        assert len({n.get('id') for n in drawing.iter('{'+XDR+'}cNvPr')})==5
        assert_namespace_bindings(parts)
