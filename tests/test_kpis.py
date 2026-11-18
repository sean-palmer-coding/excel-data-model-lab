import pandas as pd
import pytest
from excel_model_lab import WorkbookModel
from excel_model_lab.measures import list_script_measures

def model():
    m = WorkbookModel.create({'Sales': pd.DataFrame({'Amount':[20,30,40]})})
    m.add_measure('Sales','Total Amount','SUM(Sales[Amount])',format_string='0')
    m.add_measure('Sales','Target Amount','100')
    return m

def test_kpi_roundtrip_update_remove(tmp_path):
    m = model().add_kpi('Sales','Total Amount',target_measure='Target Amount')
    p = m.save(tmp_path/'kpi.xlsx')
    m = WorkbookModel(p)
    assert m.list_kpis()[0]['lower'] == 0.4
    m.update_kpi('Sales','Total Amount',target_measure='Target Amount',lower=0.5,upper=0.9)
    assert m.list_kpis()[0]['upper'] == 0.9
    m.save(tmp_path/'updated.xlsx')
    m.remove_kpi('Sales','Total Amount').save(tmp_path/'removed.xlsx')
    assert m.list_kpis() == []
    f = next(f for f in m._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
    assert {v.name for v in list_script_measures(m._image.read_logical_file(f))} == {'Total Amount','Target Amount'}

def test_kpi_invalid_operations_are_atomic():
    m = model()
    before = m._image.data
    with pytest.raises(ValueError):
        m.add_kpi('Sales','Total Amount',target_measure='Target Amount',lower=0.9,upper=0.1)
    assert m._image.data == before
    assert len(m._pending_measures) == 2
    m.add_kpi('Sales','Total Amount',target_measure='Target Amount')
    before = m._image.data
    with pytest.raises(KeyError):
        m.update_kpi('Sales','Total Amount',target_measure='Missing')
    assert m._image.data == before

def test_kpi_removal_rejects_other_helper_dependencies():
    m = model().add_kpi('Sales','Total Amount',target_measure='Target Amount')
    m.add_measure('Sales','Dependent','[_Total Amount Goal] + 1')
    before = m._image.data
    with pytest.raises(ValueError, match='depends'):
        m.remove_kpi('Sales','Total Amount')
    assert m._image.data == before
    assert m._pending_measures[0].name == 'Dependent'


def test_fixed_target_descending_and_update_target_kind(tmp_path):
    from decimal import Decimal
    m = model().add_kpi('Sales', 'Total Amount', target_value=100, mode='descending', icon='Shapes')
    item, = m.list_kpis()
    assert item['target_value'] == Decimal(100)
    assert item['target_measure'] is None
    assert item['thresholds'] == (40, 80)
    assert item['threshold_type'] == 'absolute'
    assert item['mode'] == 'descending' and item['icon'] == 'Shapes'
    path = m.save(tmp_path / 'fixed.xlsx')
    m = WorkbookModel(path)
    assert m.list_kpis()[0] == item
    m.update_kpi('Sales', 'Total Amount', target_measure='Target Amount', thresholds=(0.5, 0.95))
    item, = m.list_kpis()
    assert item['target_value'] is None and item['target_measure'] == 'Target Amount'
    assert item['thresholds'] == (0.5, 0.95)
    m.remove_kpi('Sales', 'Total Amount')
    assert m.list_kpis() == []
    script = next(f for f in m._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
    assert b'GoalValue' not in m._image.read_logical_file(script)


@pytest.mark.parametrize('mode,icon,thresholds,expected_codes', [
    ('ascending', 'Traffic Light - Single', (40, 80), (-1, 0, 1)),
    ('descending', 'Road Signs', (40, 80), (1, 0, -1)),
    ('centered', 'Traffic Light', (34, 67, 133, 166), (-1, 0, 1, 0, -1)),
    ('centered_reversed', 'Shapes', (34, 67, 133, 166), (1, 0, -1, 0, 1)),
    ('ascending', 'Gauge - Ascending', (20, 40, 60, 80), (-2, -1, 0, 1, 2)),
    ('descending', 'Gauge - Ascending', (20, 40, 60, 80), (2, 1, 0, -1, -2)),
    ('centered', 'Gauge - Ascending', (20, 40, 60, 80, 120, 140, 160, 180), (-2, -1, 0, 1, 2, 1, 0, -1, -2)),
    ('centered_reversed', 'Gauge - Ascending', (20, 40, 60, 80, 120, 140, 160, 180), (2, 1, 0, -1, -2, -1, 0, 1, 2)),
])
def test_kpi_layout_metadata_and_status_bands(mode, icon, thresholds, expected_codes):
    import re
    m = model().add_kpi('Sales', 'Total Amount', target_value=100,
        mode=mode, icon=icon, thresholds=thresholds)
    item, = m.list_kpis()
    assert item['mode'] == mode and item['icon'] == icon
    assert item['thresholds'] == thresholds
    script = next(f for f in m._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
    measures = list_script_measures(m._image.read_logical_file(script))
    status = next(v.expression for v in measures if v.name == '_Total Amount Status')
    codes = tuple(map(int, re.findall(r'<[0-9.]+,(-?[0-9]+),', status)))
    last = int(re.search(r',(-?[0-9]+)\)+$', status).group(1))
    assert codes + (last,) == expected_codes
    assert 'ISBLANK' in status


def test_kpi_percentage_centered_and_structured_validation():
    from excel_model_lab.kpis import KpiValidationError
    from excel_model_lab.errors import ModelValidationError
    m = model().add_kpi('Sales', 'Total Amount', target_measure='Target Amount',
        thresholds=(0.34, 0.67, 1.33, 1.66), mode='centered')
    assert m.list_kpis()[0]['thresholds'] == (0.34, 0.67, 1.33, 1.66)
    before = m._image.data
    with pytest.raises(KpiValidationError) as error:
        m.update_kpi('Sales', 'Total Amount', target_value=100, icon='Unknown')
    assert isinstance(error.value, ModelValidationError)
    assert error.value.issues[0].code == 'unsupported_kpi_icon'
    assert error.value.issues[0].object_path == ('tables', 'Sales', 'kpis', 'Total Amount')
    assert m._image.data == before


@pytest.mark.parametrize('options', [
    {}, {'target_value': 100, 'target_measure': 'Target Amount'},
    {'target_value': float('nan')}, {'target_value': True},
    {'target_value': 100, 'mode': 'centered'},
    {'target_value': 100, 'thresholds': (80, 40)},
    {'target_value': 100, 'thresholds': (40, 40)},
    {'target_value': 100, 'thresholds': (20, 40, 60, 80)},
])
def test_kpi_invalid_generalized_options_do_not_mutate(options):
    m = model()
    before = m._image.data
    with pytest.raises(ValueError):
        m.add_kpi('Sales', 'Total Amount', **options)
    assert m._image.data == before and len(m._pending_measures) == 2
