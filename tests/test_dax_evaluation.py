import importlib.metadata
import math

import pandas as pd
import pytest

from excel_model_lab import WorkbookModel
from excel_model_lab.dax_evaluation import (
    DaxEvaluator, PbixMcpBackend, DaxBackendUnavailableError,
    DaxEvaluationError, DaxUnsupportedError, _scalar,
)


CASES = [
    ('Total', 'SUM(Sales[Amount])', 70),
    ('Rows', 'COUNTROWS(Sales)', 4),
    ('Count Amount', 'COUNT(Sales[Amount])', 3),
    ('Blank Amount', 'COUNTBLANK(Sales[Amount])', 1),
    ('Distinct Products', 'DISTINCTCOUNT(Sales[ProductID])', 3),
    ('Weighted', 'SUMX(Sales, Sales[Amount] * Sales[Units])', 210),
    ('Category A', 'CALCULATE([Total], Products[Category] = "A")', 50),
    ('Blank Scalar', 'BLANK()', None),
    ('Safe Divide', 'DIVIDE([Total], 0, -1)', -1),
    ('Clear Category', 'CALCULATE(CALCULATE([Total], ALL(Products)), Products[Category] = "A")', 70),
    ('Filtered Amount', 'CALCULATE([Total], FILTER(Sales, Sales[Amount] > 15))', 60),
    ('Related Weighted', 'SUMX(Sales, Sales[Amount] * RELATED(Products[Weight]))', 240),
    ('Category Max', 'MAXX(VALUES(Products[Category]), [Total])', 50),
    ('Blank Is Zero', 'IF(BLANK() = 0, 1, 0)', 1),
]


@pytest.fixture(scope='module')
def model():
    result = WorkbookModel.create({
        'Sales': pd.DataFrame({'ProductID': [1, 2, 1, 3],
                              'Amount': pd.array([10, 20, None, 40], dtype='Int64'),
                              'Units': [1, 2, 3, 4]}),
        'Products': pd.DataFrame({'ProductID': [1, 2, 3], 'Category': ['A', 'B', 'A'], 'Weight': [2, 3, 4]}),
    })
    result.add_relationship('Sales', 'ProductID', 'Products', 'ProductID')
    for name, expression, _ in CASES:
        result.add_measure('Sales', name, expression)
    return result


@pytest.fixture(scope='module')
def evaluator(model):
    pytest.importorskip('pbix_mcp')
    return DaxEvaluator(model)


@pytest.mark.parametrize('name,expression,expected', CASES)
def test_independent_and_native_oracle_cases(evaluator, name, expression, expected):
    answer = evaluator.evaluate(name)
    assert answer.value == expected
    assert answer.backend == 'pbix-mcp' and answer.backend_version == '0.9.153'
    assert answer.elapsed_seconds >= 0


def test_outer_filters_and_no_cross_call_cache_leak(evaluator):
    assert evaluator.evaluate('Total', filters={('Products', 'Category'): ['A']}).value == 50
    assert evaluator.evaluate('Total', filters={('Products', 'Category'): ['B']}).value == 20
    assert evaluator.evaluate('Total').value == 70
    assert evaluator.evaluate('Total', filters={('Products', 'Category'): []}).value is None
    assert evaluator.evaluate('Total', filters={('Products', 'Category'): ['absent']}).value is None
    assert evaluator.evaluate('Blank Scalar').value is None


def test_expression_home_table_and_case_insensitive_measure(evaluator):
    assert evaluator.evaluate('tOtAl').value == 70
    assert evaluator.evaluate_expression('[Total] + 1', table='Sales').value == 71
    with pytest.raises(ValueError, match='home table'):
        evaluator.evaluate_expression('1', table='Unknown')


@pytest.mark.parametrize('expression,exception', [
    ('NOT_A_DAX_FUNCTION(1)', DaxUnsupportedError),
    ('[Missing Measure] + 1', DaxEvaluationError),
    ('SUM(Sales[Missing])', DaxEvaluationError),
])
def test_backend_failure_never_becomes_blank(evaluator, expression, exception):
    with pytest.raises(exception):
        evaluator.evaluate_expression(expression, table='Sales')


def test_snapshot_does_not_follow_later_model_changes(tmp_path):
    pytest.importorskip('pbix_mcp')
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': [1, 2, 3]})})
    model.add_measure('T', 'Total', 'SUM(T[Value])')
    before = DaxEvaluator(model)
    model.replace_column('T', 'Value', [10, 20, 30])
    after = DaxEvaluator(model)
    assert before.evaluate('Total').value == 6
    assert after.evaluate('Total').value == 60
    saved = model.save(tmp_path / 'saved.xlsx')
    assert DaxEvaluator(saved).evaluate('Total').value == 60


def test_unprocessed_calculated_values_are_not_false_blanks():
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': [1, 2, 3]})})
    model.add_calculated_column('T', 'Doubled', '[Value] * 2', data_type='integer')
    with pytest.raises(DaxUnsupportedError, match='processed in Excel'):
        DaxEvaluator(model, backend=object())


@pytest.mark.parametrize('budget', [0, -1, math.inf, math.nan, True, '30'])
def test_invalid_budget_rejected(model, budget):
    with pytest.raises(ValueError, match='max_seconds'):
        DaxEvaluator(model, backend=object(), max_seconds=budget)


def test_missing_and_unverified_versions_fail_explicitly(monkeypatch):
    def missing(_):
        raise importlib.metadata.PackageNotFoundError('pbix-mcp')
    monkeypatch.setattr(importlib.metadata, 'version', missing)
    with pytest.raises(DaxBackendUnavailableError, match=r'\[dax\]'):
        PbixMcpBackend()
    monkeypatch.setattr(importlib.metadata, 'version', lambda _: '0.9.999')
    with pytest.raises(DaxBackendUnavailableError, match='require conformance'):
        PbixMcpBackend()


@pytest.mark.parametrize('filters', [{'Sales.Amount': [10]}, {('Sales', 'Missing'): [1]}, {('Sales', 'Amount'): '10'}])
def test_invalid_filter_definition(evaluator, filters):
    with pytest.raises(ValueError):
        evaluator.evaluate('Total', filters=filters)


@pytest.mark.parametrize('value', [float(2**53), float(-(2**53)), 1.5])
def test_unsafe_integer_readback_rejected(value):
    with pytest.raises(DaxUnsupportedError, match='exact safe integer'):
        _scalar(value, 'int64')


def test_nullable_integer_and_date_normalization():
    assert _scalar(float('nan'), 'int64') is None
    assert _scalar(2.0, 'int64') == 2
    assert isinstance(_scalar(2.0, 'int64'), int)
    assert _scalar(pd.Timestamp('2026-01-01')) == '2026-01-01T00:00:00'
    with pytest.raises(DaxUnsupportedError, match='whole-second'):
        _scalar(pd.Timestamp('2026-01-01', tz='UTC'))


def test_date_table_metadata_maps_to_backend():
    pytest.importorskip('pbix_mcp')
    model = WorkbookModel.create({'Calendar': pd.DataFrame({
        'Date': pd.date_range('2026-01-01', periods=3), 'Value': [10, 20, 30]})})
    model.mark_date_table('Calendar', 'Date')
    model.add_measure('Calendar', 'Total', 'SUM(Calendar[Value])')
    evaluator = DaxEvaluator(model)
    assert evaluator._snapshot.date_tables == {'Calendar': 'Date'}
    assert evaluator.evaluate('Total', filters={('Calendar', 'Date'): [pd.Timestamp('2026-01-02')]}).value == 20


def test_unmarked_dates_do_not_invent_relationships():
    pytest.importorskip('pbix_mcp')
    model = WorkbookModel.create({
        'Calendar': pd.DataFrame({'Date': pd.date_range('2026-01-01', periods=3)}),
        'Sales': pd.DataFrame({'Date': pd.date_range('2026-01-01', periods=3), 'Amount': [10, 20, 30]}),
    })
    model.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')
    model.add_measure('Sales', 'Nested', 'CALCULATE([Total], Calendar[Date] = DATE(2026, 1, 2))')
    evaluator = DaxEvaluator(model)
    assert evaluator.evaluate('Total', filters={('Calendar', 'Date'): [pd.Timestamp('2026-01-02')]}).value == 60
    assert evaluator.evaluate('Nested').value == 60
    model.mark_date_table('Calendar', 'Date')
    with pytest.raises(DaxUnsupportedError, match='Disconnected date-like'):
        DaxEvaluator(model, backend=object())


def test_published_backend_date_multihop_is_rejected():
    dates = pd.date_range('2026-01-01', periods=3)
    model = WorkbookModel.create({
        'Calendar': pd.DataFrame({'Date': dates}),
        'Orders': pd.DataFrame({'Date': dates, 'ID': [1, 2, 3]}),
        'Sales': pd.DataFrame({'OrderID': [1, 2, 3], 'Amount': [10, 20, 30]}),
    })
    model.add_relationship('Orders', 'Date', 'Calendar', 'Date')
    model.add_relationship('Sales', 'OrderID', 'Orders', 'ID')
    model.mark_date_table('Calendar', 'Date')
    with pytest.raises(DaxUnsupportedError, match='multihop'):
        DaxEvaluator(model, backend=object())


def test_circular_measures_error_instead_of_blank():
    pytest.importorskip('pbix_mcp')
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': [1, 2, 3]})})
    model.add_measure('T', 'Loop', '[Loop] + 1')
    evaluator = DaxEvaluator(model)
    with pytest.raises(DaxEvaluationError):
        evaluator.evaluate('Loop')


@pytest.mark.parametrize('expression', ['SUM(Missing[X])', 'COUNTROWS(Missing)',
    'SUM(Sales[Missing])', 'SUM(Sales[Amount], 2)', 'SUM('])
def test_invalid_bindings_and_argument_shapes_are_not_blanks(evaluator, expression):
    with pytest.raises(DaxEvaluationError):
        evaluator.evaluate_expression(expression, table='Sales')


def test_type_guards_prevent_backend_false_results():
    pytest.importorskip('pbix_mcp')
    model = WorkbookModel.create({'T': pd.DataFrame({
        'Text': ['A', 'B', 'A'], 'Flag': pd.array([True, None, False], dtype='boolean'),
        'Date': pd.date_range('2026-01-01', periods=3)})})
    evaluator = DaxEvaluator(model)
    for expression in ('SUM(T[Text])', 'SUM(T[Flag])', 'COUNT(T[Flag])',
                       'SUM(T[Date])', 'SUMX(T, T[Flag])', 'SUMX(T, T[Text])'):
        with pytest.raises(DaxEvaluationError):
            evaluator.evaluate_expression(expression, table='T')
    assert evaluator.evaluate_expression('DATE(2026, 1, 2)', table='T').value == pd.Timestamp('2026-01-02')


def test_unverified_function_and_unsafe_integer_results_rejected(evaluator):
    with pytest.raises(DaxUnsupportedError, match='outside the verified'):
        evaluator.evaluate_expression('MAX(Sales[Amount], "hello")', table='Sales')
    with pytest.raises(DaxUnsupportedError, match='overflow'):
        evaluator.evaluate_expression('4611686018427387904 + 4611686018427387904', table='Sales')
    with pytest.raises(DaxUnsupportedError, match='overflow'):
        evaluator.evaluate_expression('(4611686018427387904 * 2) - (4611686018427387904 * 2)', table='Sales')


def test_lexical_quotes_comments_and_schema_binding():
    pytest.importorskip('pbix_mcp')
    model = WorkbookModel.create({'Sales Table': pd.DataFrame({'Amount': [1, 2, 3]})})
    model.add_measure('Sales Table', 'Total', "SUM('Sales Table'[Amount])")
    evaluator = DaxEvaluator(model)
    assert evaluator.evaluate('Total').value == 6
    assert evaluator.evaluate_expression('IF(TRUE(), "SUM(Missing[X])", "other")', table='Sales Table').value == 'SUM(Missing[X])'


def test_backend_escaped_identifier_gap_is_explicit():
    model = WorkbookModel.create({"Sales O'Brien": pd.DataFrame({'Amount': [1, 2, 3]})})
    model.add_measure("Sales O'Brien", 'Total', "SUM('Sales O''Brien'[Amount])")
    with pytest.raises(DaxUnsupportedError, match='identifiers'):
        DaxEvaluator(model, backend=object())
