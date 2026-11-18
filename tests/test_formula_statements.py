import pytest
from excel_model_lab.formula_statements import iter_measure_definitions, normalize_expression
from excel_model_lab.abf import UnsupportedModelError


def test_structural_reference_lexing_and_rewrite():
    from excel_model_lab.formula_statements import formula_references, rewrite_references
    text = '''SUM('Sa''les' /* comment */ [A]]B]) + COUNTROWS(Other) + [Measure]
        + LEN("Other[A]]B]") // Other[A]]B]
    '''
    refs = formula_references(text, table_names=["Sa'les", 'Other'])
    assert [(r.table, r.name, r.kind) for r in refs] == [("Sa'les", 'A]B', 'field'), (None, 'Other', 'table'), (None, 'Measure', 'field')]
    rewritten = rewrite_references(text, table_names=["Sa'les", 'Other'],
        table_renames={"Sa'les": 'New Sales', 'Other': "O'ther"},
        field_renames={("Sa'les", 'A]B'): 'Value]new', (None, 'Measure'): 'New Measure'})
    assert "SUM('New Sales' /* comment */ [Value]]new])" in rewritten
    assert "COUNTROWS('O''ther')" in rewritten
    assert 'LEN("Other[A]]B]") // Other[A]]B]' in rewritten
    assert '[New Measure]' in rewritten


def test_table_named_like_function_or_variable_is_not_rewritten():
    from excel_model_lab.formula_statements import rewrite_references
    text = 'VAR Sales = FILTER(Other, Other[Amount] > 1) RETURN COUNTROWS(Sales) + SUM(Other[Amount])'
    actual = rewrite_references(text, table_names=['Sales', 'Other', 'SUM'],
        table_renames={'Sales': 'Orders', 'Other': 'Details', 'SUM': 'Wrong'})
    assert 'VAR Sales' in actual and 'COUNTROWS(Sales)' in actual
    assert 'SUM(' in actual and "'Details'[Amount]" in actual


def test_literals_comments_and_escaped_names():
    text = '''-- CREATE MEASURE 'Fake'[Ignore]=1;
CREATE MEASURE 'Sa''les'[A]]B]="x;y" & ";" /* ; */;
CREATE MEASURE 'Sales'[Other]=2 // ; ignored
;'''
    definitions = list(iter_measure_definitions(text))
    assert len(definitions) == 2
    assert definitions[0].group('name') == 'A]]B'
    assert definitions[0].group('expression') == '"x;y" & ";" /* ; */'
    assert definitions[1].group('expression').strip() == '2 // ; ignored'


def test_expression_validation():
    assert normalize_expression('"a;b";') == '"a;b"'
    with pytest.raises(ValueError, match='multiple statements'):
        normalize_expression("1; CREATE MEASURE 'Sales'[Injected]=2;")
    with pytest.raises(UnsupportedModelError, match='Unterminated'):
        normalize_expression('"unfinished')


def test_edit_preserves_semicolon_literal(tmp_path):
    import pandas as pd
    from excel_model_lab import WorkbookModel
    from excel_model_lab.measures import list_script_measures
    m = WorkbookModel.create({'Sales': pd.DataFrame({'Amount':[1,2]})})
    m.add_measure('Sales','Text','"before;after"')
    path = m.save(tmp_path/'text.xlsx')
    m = WorkbookModel(path)
    m.update_measure('Sales','Text',expression='"new;value"')
    script = next(f for f in m._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
    result = list_script_measures(m._image.read_logical_file(script))
    assert next(x.expression for x in result if x.name == 'Text') == '"new;value"'
