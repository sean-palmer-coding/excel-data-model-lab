import pandas as pd
import pytest
from excel_model_lab import WorkbookModel
from excel_model_lab.measures import list_script_measures

def test_measure_update_remove_and_dependencies(tmp_path):
    m = WorkbookModel.create({'Sales':pd.DataFrame({'Amount':[20,30,40]})})
    m.add_measure('Sales','Total','SUM(Sales[Amount])')
    m.add_measure('Sales','Target','100')
    first = m.save(tmp_path/'first.xlsx')
    m = WorkbookModel(first)
    m.update_measure('Sales','Target',expression='120',format_string='0.00')
    updated = m.save(tmp_path/'updated.xlsx')
    m = WorkbookModel(updated)
    f = next(f for f in m._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
    assert next(v.expression for v in list_script_measures(m._image.read_logical_file(f)) if v.name == 'Target') == '120'
    m.add_kpi('Sales','Total',target_measure='Target')
    with pytest.raises(ValueError, match='depends'):
        m.remove_measure('Sales','Target')
    m.remove_kpi('Sales','Total').remove_measure('Sales','Target').save(tmp_path/'removed.xlsx')
    from test_measure_layout import layouts, measure_rows
    assert 'Measures\\Target' not in measure_rows(layouts(tmp_path/'removed.xlsx')['Diagrams'][1])

def test_measure_name_uniqueness_is_case_insensitive(tmp_path):
    m = WorkbookModel.create({'Sales':pd.DataFrame({'Amount':[20,30,40]})})
    m.add_measure('Sales','Total','1').add_measure('Sales','TOTAL','2')
    with pytest.raises(ValueError, match='already exists'):
        m.save(tmp_path/'duplicate.xlsx')
    assert not (tmp_path/'duplicate.xlsx').exists()


def test_measure_removal_ignores_literals_but_preserves_real_dependencies(tmp_path):
    m = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [20, 30]})})
    m.add_measure('Sales', 'A]B', '100')
    m.add_measure('Sales', 'Description', '\"[A]]B]\" /* [A]]B] */')
    path = m.save(tmp_path / 'literal.xlsx')
    m = WorkbookModel(path)
    m.remove_measure('Sales', 'A]B').save(tmp_path / 'removed-literal.xlsx')
    m = WorkbookModel(path)
    m.update_measure('Sales', 'Description', expression='[A]]B] + 1')
    with pytest.raises(ValueError, match='depends'):
        m.remove_measure('Sales', 'A]B')


def test_column_removal_checks_identifiers_in_pending_and_saved_measures(tmp_path):
    from excel_model_lab.abf import UnsupportedModelError
    m = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [20, 30], 'Unused': [1, 2]})})
    m.add_measure('Sales', 'Label', '\"[Unused]\" // [Unused]')
    m.remove_column('Sales', 'Unused')
    path = m.save(tmp_path / 'column-literal.xlsx')
    assert list(WorkbookModel(path).get_table('Sales').columns) == ['Amount']
    m = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [20, 30], 'Unused': [1, 2]})})
    m.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')
    with pytest.raises(UnsupportedModelError, match='calculation'):
        m.remove_column('Sales', 'Amount')
    path = m.save(tmp_path / 'column-reference.xlsx')
    with pytest.raises(UnsupportedModelError, match='calculation'):
        WorkbookModel(path).remove_column('Sales', 'Amount')
