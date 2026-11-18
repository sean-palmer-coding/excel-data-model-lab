from dataclasses import FrozenInstanceError
from pathlib import Path
import shutil
import sys
from types import SimpleNamespace

import pytest

from excel_model_lab import WorkbookModel, ModelValidationError
from excel_model_lab.reports import WorkbookReport, PivotTableSpec, PivotChartSpec


FIXTURE = Path(__file__).parent / 'data' / 'native-authoring-base.xlsx'


@pytest.fixture
def model():
    return WorkbookModel(FIXTURE)


def pivot(**changes):
    return PivotTableSpec(**dict({'name': 'Summary', 'sheet': 'Report',
                                 'rows': [('Table1', 'Label')]}, **changes))


def test_specs_freeze_nested_input_and_report_snapshot(model):
    rows = [['Table1', 'Label']]
    spec = pivot(rows=rows)
    rows[0][1] = 'ID'
    assert spec.rows == (('Table1', 'Label'),)
    with pytest.raises(FrozenInstanceError):
        spec.sheet = 'Other'
    reports = WorkbookReport(model).add_pivot_table(spec)
    snapshot = reports.list_reports()
    reports.add_pivot_chart(name='Chart1', pivot_table='Summary')
    assert snapshot == (spec,)
    assert len(reports.list_reports()) == 2


@pytest.mark.parametrize('changes', [
    {'rows': [('Missing', 'ID')]}, {'rows': [('Table1', 'Missing')]},
    {'columns': [('Table1', 'Label')]}, {'measures': ['Unknown']},
    {'sheet': 'Bad/Sheet'}, {'sheet': "'Bad"}, {'sheet': 'x' * 32},
    {'cell': 'XFE1'}, {'cell': 'A1048577'}, {'cell': 'A1:B2'},
    {'layout': 'unsupported'}, {'row_grand_totals': 1},
])
def test_invalid_pivot_does_not_stage_or_mutate_model(model, changes):
    reports = WorkbookReport(model)
    before = model._image.data
    with pytest.raises(ModelValidationError):
        reports.add_pivot_table(pivot(**changes))
    assert reports.list_reports() == ()
    assert model._image.data == before


def test_pending_measure_and_current_model_mutations(model, tmp_path, monkeypatch):
    model.add_measure('Table1', 'CountID', 'COUNT(Table1[ID])')
    reports = WorkbookReport(model).add_pivot_table(pivot(measures=['CountID']))
    model.remove_measure('Table1', 'CountID')
    monkeypatch.setitem(sys.modules, 'excel_model_lab.report_package', SimpleNamespace(add_reports=lambda *a, **k: None))
    target = tmp_path / 'unchanged.xlsx'
    target.write_bytes(b'original')
    with pytest.raises(ModelValidationError, match='measure'):
        reports.save(target)
    assert target.read_bytes() == b'original'
    assert len(reports.list_reports()) == 1


def test_current_deleted_column_revalidated_at_save(model, tmp_path, monkeypatch):
    reports = WorkbookReport(model).add_pivot_table(pivot())
    model.unbind_source('Table1').remove_column('Table1', 'Label')
    monkeypatch.setitem(sys.modules, 'excel_model_lab.report_package', SimpleNamespace(add_reports=lambda *a, **k: None))
    with pytest.raises(ModelValidationError, match='field'):
        reports.save(tmp_path / 'invalid.xlsx')
    assert not (tmp_path / 'invalid.xlsx').exists()


@pytest.mark.parametrize('changes', [
    {'pivot_table': 'Missing'}, {'name': 'summary'}, {'chart_type': 'pie'},
    {'width': 0}, {'height': True}, {'cell': 'A0'}, {'sheet': 'Bad?'},
])
def test_invalid_chart_keeps_prior_report(model, changes):
    reports = WorkbookReport(model).add_pivot_table(pivot())
    fields = dict({'name': 'Chart', 'pivot_table': 'Summary'}, **changes)
    with pytest.raises(ModelValidationError):
        reports.add_pivot_chart(**fields)
    assert reports.list_reports() == (pivot(),)


def test_duplicate_pivot_sheet_and_name_rejected(model):
    reports = WorkbookReport(model).add_pivot_table(pivot())
    for spec in (pivot(name='summary', sheet='Other'), pivot(name='Other', sheet='report')):
        with pytest.raises(ModelValidationError):
            reports.add_pivot_table(spec)
    assert len(reports.list_reports()) == 1


def test_writer_failure_is_atomic_and_cleans_staging(model, tmp_path, monkeypatch):
    reports = WorkbookReport(model).add_pivot_table(pivot())
    destination = tmp_path / 'out.xlsx'
    destination.write_bytes(b'original')

    def fail(source, output, pivots, charts, *, model, workbook_name, slicers=(), timelines=()):
        assert workbook_name == 'out.xlsx'
        assert source.exists() and source.parent.parent == tmp_path
        assert pivots == (pivot(),) and charts == ()
        output.write_bytes(b'partial')
        raise ModelValidationError('writer validation failed')

    monkeypatch.setitem(sys.modules, 'excel_model_lab.report_package', SimpleNamespace(add_reports=fail))
    with pytest.raises(ModelValidationError, match='writer'):
        reports.save(destination)
    assert destination.read_bytes() == b'original'
    assert list(tmp_path.iterdir()) == [destination]
    assert reports.list_reports() == (pivot(),)
    assert model.get_table('Table1').shape[0] == 3


def test_writer_receives_updated_model_and_bytes(model, tmp_path, monkeypatch):
    reports = WorkbookReport(model).add_pivot_table(pivot())
    model.replace_column('Table1', 'ID', [11, 22, 33])

    def copy(source, output, pivots, charts, *, model, workbook_name, slicers=(), timelines=()):
        assert workbook_name == 'report.xlsx'
        assert WorkbookModel(source).get_table('Table1')['ID'].tolist() == [11, 22, 33]
        shutil.copyfile(source, output)

    monkeypatch.setitem(sys.modules, 'excel_model_lab.report_package', SimpleNamespace(add_reports=copy))
    output = reports.save(tmp_path / 'report.xlsx')
    assert WorkbookModel(output).get_table('Table1')['ID'].tolist() == [11, 22, 33]
    assert reports.to_bytes()[:2] == b'PK'


def test_path_constructor_and_chart_separate_sheet():
    reports = WorkbookReport(FIXTURE).add_pivot_table(pivot())
    reports.add_pivot_chart(PivotChartSpec('Chart', 'Summary', sheet='Charts'))
    assert reports.list_reports()[1].sheet == 'Charts'


@pytest.mark.parametrize('name', ['A1', 'XFD1048576', 'R1C1', 'C', 'R', '1Report', 'Bad Name', 'Bad-Name', 'a\x00b', 'a\ud800b', 'a\ufffeb'])
def test_pivot_names_reject_unsupported_identifiers(model, name):
    report = WorkbookReport(model)
    with pytest.raises(ModelValidationError):
        report.add_pivot_table(pivot(name=name))
    assert report.list_reports() == ()


@pytest.mark.parametrize('bad', ['\x00', '\x08', '\udfff', '\ufffe', '\uffff'])
def test_chart_title_and_names_reject_invalid_xml(model, bad):
    report = WorkbookReport(model).add_pivot_table(pivot())
    for changes in ({'title': bad}, {'name': 'Chart'+bad}, {'sheet': 'Chart'+bad}):
        with pytest.raises(ModelValidationError):
            report.add_pivot_chart(**dict({'name': 'Chart', 'pivot_table': 'Summary'}, **changes))
    assert len(report.list_reports()) == 1


def test_field_only_pivot_and_multiline_chart_title_are_allowed(model):
    report = WorkbookReport(model).add_pivot_table(pivot(measures=()))
    report.add_pivot_chart(name='Sales chart', pivot_table='Summary', title='Sales\n& profit <2026>')
    assert len(report.list_reports()) == 2
    with pytest.raises(ModelValidationError, match='at least one'):
        WorkbookReport(model).add_pivot_table(pivot(rows=()))


def test_public_report_exports_and_invalid_measures():
    import excel_model_lab
    import excelmodel
    for module in (excel_model_lab, excelmodel):
        assert module.WorkbookReport is WorkbookReport
        assert module.PivotTableSpec is PivotTableSpec
        assert module.PivotChartSpec is PivotChartSpec
    with pytest.raises(ModelValidationError, match='sequence'):
        pivot(measures=None)


def test_staged_pivot_rename_updates_chart_and_controls(model):
    from excel_model_lab.reports import SlicerSpec
    report=WorkbookReport(model).add_pivot_table(pivot())
    report.add_pivot_chart(name='Chart',pivot_table='Summary')
    report.add_slicer(SlicerSpec('Labels','Summary',('Table1','Label')))
    report.update_pivot_table('Summary',name='Renamed')
    assert [s.name for s in report.list_reports()]==['Renamed','Chart','Labels']
    assert all(s.pivot_table=='Renamed' for s in report.list_reports()[1:])
    before=report.list_reports()
    with pytest.raises(ModelValidationError):
        report.update_pivot_table('Renamed',name='Chart')
    assert report.list_reports()==before


def test_control_field_rechecked_against_live_model_at_save(model,tmp_path):
    from excel_model_lab.reports import SlicerSpec
    report=WorkbookReport(model).add_pivot_table(pivot(rows=(('Table1','ID'),)))
    report.add_slicer(SlicerSpec('Labels','Summary',('Table1','Label')))
    model.unbind_source('Table1').remove_column('Table1','Label')
    target=tmp_path/'out.xlsx';target.write_bytes(b'original')
    with pytest.raises(ModelValidationError,match='control field'):
        report.save(target)
    assert target.read_bytes()==b'original'


def test_control_duplicates_and_shapes_are_atomic(model):
    from excel_model_lab.reports import SlicerSpec
    report=WorkbookReport(model).add_pivot_table(pivot())
    with pytest.raises(ModelValidationError,match='Duplicate connected'):
        report.add_slicer(SlicerSpec('Labels','Summary',('Table1','Label'),connected_pivots=('Summary',)))
    with pytest.raises(ModelValidationError,match='sequence'):
        SlicerSpec('Labels','Summary','ID')
    report.add_slicer(SlicerSpec('Labels','Summary',('Table1','Label')))
    before=report.list_reports()
    with pytest.raises(ModelValidationError,match='Duplicate'):
        report.add_pivot_chart(name='Labels',pivot_table='Summary')
    assert report.list_reports()==before


def test_existing_report_edit_validates_before_staging(model,tmp_path):
    source=WorkbookReport(model).add_pivot_table(pivot()).save(tmp_path/'source.xlsx')
    report=WorkbookReport(source)
    with pytest.raises(ModelValidationError):
        report.update_pivot_table('Summary',layout='invalid')
    assert report._pivot_updates=={}
    with pytest.raises(ModelValidationError):
        report.update_pivot_table('Summary',invented=True)
    assert report._pivot_updates=={}
    report.update_pivot_table('Summary',layout='compact')
    assert report._pivot_updates=={'Summary':{'layout':'compact'}}


def test_new_chart_can_reference_existing_pivot_and_tracks_rename(model,tmp_path):
    source=WorkbookReport(model).add_pivot_table(pivot()).save(tmp_path/'source.xlsx')
    report=WorkbookReport(source).add_pivot_chart(name='Chart',pivot_table='Summary',sheet='Charts')
    report.update_pivot_table('Summary',name='Renamed')
    assert report.list_reports()[0].pivot_table=='Renamed'
    with pytest.raises(ModelValidationError,match='dependent'):
        report.remove_report('Summary')
    assert report._remove==()


def test_existing_remove_rejects_dependency_immediately(model,tmp_path):
    source=WorkbookReport(model).add_pivot_table(pivot()).add_pivot_chart(name='Chart',pivot_table='Summary').save(tmp_path/'source.xlsx')
    report=WorkbookReport(source)
    with pytest.raises(ModelValidationError,match='dependent'):
        report.remove_report('Summary')
    assert report._remove==()
    report.remove_report('Chart').remove_report('Summary')
    assert report._remove==('Chart','Summary')
