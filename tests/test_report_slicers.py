from types import SimpleNamespace
from xml.etree import ElementTree as ET
from zipfile import ZipFile
import posixpath

import pandas as pd
import pytest

from excel_model_lab import WorkbookModel, ModelValidationError
from excel_model_lab.reports import WorkbookReport, PivotTableSpec, PivotChartSpec, SlicerSpec, TimelineSpec
from excel_model_lab.report_slicers import append_controls, S, R, X14, X15, XDR, A


def _base(tmp_path):
    model = WorkbookModel.create({'Sales': pd.DataFrame({
        'Category': ['Alpha', 'Beta', 'Gamma'], 'Amount': [10, 20, 30],
        'Day': pd.to_datetime(['2025-12-31', '2026-01-01', '2026-03-01'])})})
    model.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')
    report = WorkbookReport(model).add_pivot_table(PivotTableSpec(
        'Summary', 'Report', rows=(('Sales', 'Category'),), measures=('Total',)))
    report.add_pivot_table(PivotTableSpec('Second', 'Other', measures=('Total',)))
    report.add_pivot_chart(PivotChartSpec('Chart', 'Summary'))
    path = report.save(tmp_path / 'base.xlsx')
    with ZipFile(path) as z: parts = {n: z.read(n) for n in z.namelist()}
    return model, parts


def _spec(name='CategoryControl', timeline=False, **kwargs):
    attrs = dict(name=name, pivot_table='Summary', field=('Sales', 'Day' if timeline else 'Category'),
                 sheet=None, cell='J3', caption='', width=192, height=240,
                 columns=1, level='months', connected_pivots=())
    attrs.update(kwargs)
    return SimpleNamespace(**attrs)


def test_real_control_parts_separate_caches_and_preserve_chart(tmp_path):
    model, parts = _base(tmp_path)
    original = dict(parts)
    append_controls(parts, (_spec(connected_pivots=('Second',)),), (_spec('Dates', timeline=True),), model=model)
    assert parts['xl/model/item.data'] == original['xl/model/item.data']
    assert parts['xl/charts/chart1.xml'] == original['xl/charts/chart1.xml']
    for path, data in parts.items():
        if not path.endswith('.rels'): continue
        owner = '' if path == '_rels/.rels' else posixpath.dirname(path).removesuffix('/_rels')
        root = ET.fromstring(data)
        assert len({n.get('Id') for n in root}) == len(root)
        for rel in root:
            if rel.get('TargetMode') == 'External': continue
            target = posixpath.normpath(posixpath.join(owner, rel.get('Target'))).lstrip('/')
            assert target in parts, (path, target)
    slicer = ET.fromstring(parts['xl/slicerCaches/slicerCache1.xml'])
    olap = slicer.find('{'+X14+'}data/{'+X14+'}olap')
    assert slicer.get('sourceName') == '[Sales].[Category]'
    assert {n.get('name') for n in slicer.find('{'+X14+'}pivotTables')} == {'Summary', 'Second'}
    assert olap.find('{'+X14+'}selections')[0].get('n') == '[Sales].[Category].[All]'
    level = olap.find('{'+X14+'}levels')[1]
    assert level.get('count') == '3'
    assert [item.attrib for item in level.find('{'+X14+'}ranges')[0]] == [
        {'n': '[Sales].[Category].&['+name+']', 'c': name}
        for name in ('Alpha', 'Beta', 'Gamma')]
    dedicated = [ET.fromstring(data) for path, data in parts.items() if path.startswith('xl/pivotCache/') and path.endswith('.xml')]
    selected = [root for root in dedicated if any(n.get('pivotCacheId') == olap.get('pivotCacheId') for n in root.iter('{'+X14+'}pivotCacheDefinition'))]
    assert len(selected) == 1
    assert len(selected[0].find('{'+S+'}cacheFields')) == 0
    assert 'refreshOnLoad' not in selected[0].attrib
    assert 'enableRefresh' not in selected[0].attrib
    assert next(selected[0].iter('{'+X14+'}pivotCacheDefinition')).get('slicerData') == '1'
    timeline = ET.fromstring(parts['xl/timelineCaches/timelineCache1.xml'])
    bounds = timeline.find('{'+X15+'}state/{'+X15+'}bounds')
    assert bounds.attrib == {'startDate': '2025-01-01T00:00:00', 'endDate': '2027-01-01T00:00:00'}
    drawing = ET.fromstring(parts['xl/drawings/drawing1.xml'])
    assert len(drawing) == 3
    assert len(list(drawing.iter('{'+A+'}graphicData'))) == 3
    assert len({n.get('id') for n in drawing.iter('{'+XDR+'}cNvPr')}) == 3
    # These are case-sensitive extension URIs, not case-insensitive GUIDs.
    # Uppercasing them makes Excel silently discard controls or fail to open.
    workbook = ET.fromstring(parts['xl/workbook.xml'])
    workbook_uris = {n.get('uri') for n in workbook.iter('{'+S+'}ext')}
    assert '{BBE1A952-AA13-448e-AADC-164F8A28A991}' in workbook_uris
    assert '{A2CB5862-8E78-49c6-8D9D-AF26E26ADB89}' in workbook_uris
    sheet_uris = {n.get('uri') for path, data in parts.items()
                  if path.startswith('xl/worksheets/') and path.endswith('.xml')
                  for n in ET.fromstring(data).iter('{'+S+'}ext')}
    assert '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}' in sheet_uris
    assert '{7E03D99C-DC04-49d9-9315-930204A7B6E9}' in sheet_uris


def test_failed_control_batch_is_atomic(tmp_path):
    model, parts = _base(tmp_path)
    before = dict(parts)
    with pytest.raises(ModelValidationError, match='Unknown connected'):
        append_controls(parts, (_spec(connected_pivots=('Missing',)),), model=model)
    assert parts == before


def test_repeat_calls_merge_existing_slicer_views(tmp_path):
    model, parts = _base(tmp_path)
    append_controls(parts, (_spec(),), model=model)
    append_controls(parts, (_spec('SecondControl'),), model=model)
    views = [ET.fromstring(data) for path, data in parts.items() if path.startswith('xl/slicers/') and path.endswith('.xml')]
    assert len(views) == 1 and len(views[0]) == 2
    before = dict(parts)
    with pytest.raises(ModelValidationError, match='already exists'):
        append_controls(parts, (_spec(),), model=model)
    assert parts == before


def test_numeric_timeline_rejected_without_package_mutation(tmp_path):
    model, parts = _base(tmp_path)
    before = dict(parts)
    with pytest.raises(ModelValidationError, match='date'):
        append_controls(parts, (), (_spec('BadDate', timeline=True, field=('Sales', 'Amount')),), model=model)
    assert parts == before


def test_cache_names_disambiguate_sanitized_control_names(tmp_path):
    model, parts = _base(tmp_path)
    append_controls(parts, (_spec('A B'), _spec('A_B')), model=model)
    caches = [ET.fromstring(data).get('name') for path, data in parts.items()
              if path.startswith('xl/slicerCaches/') and path.endswith('.xml')]
    assert len(caches) == len(set(caches)) == 2
    assert set(caches) == {'Slicer_A_B', 'Slicer_A_B_'}


def test_public_control_validation_is_transactional(tmp_path):
    model, _ = _base(tmp_path)
    report = WorkbookReport(model).add_pivot_table(PivotTableSpec(
        'Summary', 'ControlReport', measures=('Total',)))
    report.add_slicer(SlicerSpec('Category', 'Summary', ('Sales', 'Category')))
    before = report.list_reports()
    for spec in [SlicerSpec('Category', 'Summary', ('Sales', 'Category')),
                 SlicerSpec('Other', 'Missing', ('Sales', 'Category')),
                 SlicerSpec('Other', 'Summary', ('Sales', 'Category'), connected_pivots=('Summary',))]:
        with pytest.raises(ModelValidationError):
            report.add_slicer(spec)
        assert report.list_reports() == before
    with pytest.raises(ModelValidationError, match='date'):
        report.add_timeline(TimelineSpec('Dates', 'Summary', ('Sales', 'Amount')))
    assert report.list_reports() == before


@pytest.mark.parametrize('values,expected', [
    (['A] & B', 'A] & B', None, None, ''], [
        ('[Sales].[Category].&[A]] & B]', 'A] & B'),
        ('[Sales].[Category].&', '(blank)'), ('[Sales].[Category].&[]', '')]),
    ([1, 2, 1, None], [('[Sales].[Category].&[1]', '1'),
                         ('[Sales].[Category].&[2]', '2'), ('[Sales].[Category].&', '(blank)')]),
    ([True, False, True], [('[Sales].[Category].&[True]', 'True'), ('[Sales].[Category].&[False]', 'False')]),
    ([pd.Timestamp('2026-01-02'), pd.NaT], [('[Sales].[Category].&[2026-01-02T00:00:00]', '2026-01-02T00:00:00'),
                                          ('[Sales].[Category].&', '(blank)')]),
])
def test_slicer_initial_members_are_distinct_escaped_and_include_blank(tmp_path, values, expected):
    _, parts = _base(tmp_path)
    model = SimpleNamespace(get_table=lambda _: pd.DataFrame({'Category': pd.Series(values, dtype=object)}))
    append_controls(parts, (_spec(),), model=model)
    root = ET.fromstring(parts['xl/slicerCaches/slicerCache1.xml'])
    level = root.find('{'+X14+'}data/{'+X14+'}olap/{'+X14+'}levels')[1]
    assert level.get('count') == str(len(expected))
    member_range = level.find('{'+X14+'}ranges')[0]
    assert member_range.get('startItem') == '0'
    assert [(item.get('n'), item.get('c')) for item in member_range] == expected


def test_unsupported_slicer_members_fail_without_package_mutation(tmp_path):
    _, parts = _base(tmp_path)
    before = dict(parts)
    model = SimpleNamespace(get_table=lambda _: pd.DataFrame({'Category': [1.25]}))
    with pytest.raises(ModelValidationError, match='members support'):
        append_controls(parts, (_spec(),), model=model)
    assert parts == before


def test_slicer_nullable_integer_readback_uses_integer_member_keys(tmp_path):
    _, parts = _base(tmp_path)
    model = SimpleNamespace(
        get_table=lambda _: pd.DataFrame({'Category': [1.0, float('nan'), 3.0]}),
        columns=[SimpleNamespace(table='Sales', name='Category', data_type='int64')])
    append_controls(parts, (_spec(),), model=model)
    root = ET.fromstring(parts['xl/slicerCaches/slicerCache1.xml'])
    assert [(n.get('n'), n.get('c')) for n in root.iter('{'+X14+'}i')] == [
        ('[Sales].[Category].&[1]', '1'), ('[Sales].[Category].&', '(blank)'),
        ('[Sales].[Category].&[3]', '3')]


@pytest.mark.parametrize('timeline', [False, True])
def test_unprocessed_calculated_slicer_inventory_rejected_atomically(tmp_path, timeline):
    model, parts = _base(tmp_path)
    model.add_calculated_column('Sales', 'Computed', '[Day]' if timeline else '[Amount] * 2',
                                data_type='datetime' if timeline else 'integer')
    # Build a report containing the new field in its model hierarchy metadata.
    report = WorkbookReport(model).add_pivot_table(PivotTableSpec(
        'Summary', 'Report', measures=('Total',)))
    path = report.save(tmp_path / 'calculated-base.xlsx')
    with ZipFile(path) as z: parts = {n: z.read(n) for n in z.namelist()}
    before = dict(parts)
    with pytest.raises(ModelValidationError, match='require Excel calculation'):
        spec = _spec(field=('Sales', 'Computed'), timeline=timeline)
        append_controls(parts, () if timeline else (spec,), (spec,) if timeline else (), model=model)
    assert parts == before
