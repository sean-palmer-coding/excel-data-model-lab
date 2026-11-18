import pandas as pd
import pytest
from excel_model_lab import WorkbookModel, RelationshipValidationError


def test_preflight_reports_changes_without_mutation(tmp_path):
    model = WorkbookModel.create({'Lookup': pd.DataFrame({'ID': [1, 2]}),
        'Facts': pd.DataFrame({'ID': [1, 2]})})
    model.add_relationship('Facts', 'ID', 'Lookup', 'ID', cardinality='one-to-one')
    assert model.validate().valid
    model.replace_table('Facts', pd.DataFrame({'ID': [1, 1]}))
    before = dict(model._pending)
    image = model._image.data
    report = model.validate()
    assert not report.valid
    assert report.issues[0].code == 'invalid_relationship'
    assert report.issues[0].object_path == ('relationships', 'Facts', 'ID', 'Lookup', 'ID')
    assert 'Facts.ID' in report.issues[0].message
    assert 'rows 1 and 2' in report.issues[0].message
    with pytest.raises(RelationshipValidationError):
        report.raise_for_errors()
    assert model._pending == before
    assert model._image.data == image


def test_preflight_distinguishes_writer_limit():
    model = WorkbookModel.create({'Lookup': pd.DataFrame({'ID': [1, 2]}),
        'Facts': pd.DataFrame({'ID': [1, 2]})})
    model.add_relationship('Facts', 'ID', 'Lookup', 'ID')
    model.replace_table('Lookup', pd.DataFrame({'ID': pd.array([1, None], dtype='Int64')}))
    report = model.validate()
    assert not report.valid
    assert report.issues[0].code == 'unsupported_relationship'


def test_preflight_includes_pending_measure_conflicts():
    model = WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [1, 2]})})
    model.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')
    assert model.validate().valid
    model.add_measure('Sales', 'TOTAL', '1')
    pending = list(model._pending_measures)
    report = model.validate()
    assert not report.valid
    assert report.issues[0].code == 'invalid_measure'
    assert 'already exists' in report.issues[0].message
    assert model._pending_measures == pending
    from excel_model_lab import ModelValidationError
    with pytest.raises(ModelValidationError) as caught:
        report.raise_for_errors()
    assert not isinstance(caught.value, RelationshipValidationError)
    assert caught.value.issues == report.issues


def test_relationship_error_retains_legacy_and_general_exception_contract():
    from excel_model_lab import ModelValidationError
    from excel_model_lab.validation import ValidationIssue, ValidationReport
    issue = ValidationIssue('invalid_relationship', 'Duplicate key')
    with pytest.raises(RelationshipValidationError) as caught:
        ValidationReport((issue,)).raise_for_errors()
    assert isinstance(caught.value, (ModelValidationError, ValueError))
    assert caught.value.issues == (issue,)
