"""Regressions for defects found in the November 2026 code review."""
from zipfile import ZipFile
from xml.dom import minidom

import pandas as pd
import pytest

from excel_model_lab import (ExcelModelTemplate, WorkbookModel, WorkbookReport,
                             PivotTableSpec, SlicerSpec, TimelineSpec)
from excel_model_lab.hierarchy_storage import stored_hierarchies, resolve_hierarchy

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def _files(model, fragment):
    return sorted(f.logical_path.split('\\')[-1] for f in model._image.files
                  if f.logical_path and fragment in f.logical_path)


def test_removing_hierarchy_keeps_dotted_sibling_storage(tmp_path):
    model = ExcelModelTemplate.create({'T': pd.DataFrame({'Label': ['a', 'b'], 'ID': [1, 2]})})
    model.add_hierarchy('T', 'A', ['Label', 'ID'])
    model.add_hierarchy('T', 'A.B', ['ID'])
    model = ExcelModelTemplate(model.save(tmp_path / 'before.xlsx'))
    sibling_id, _ = resolve_hierarchy(model, 'T', 'A.B')
    assert '.' not in sibling_id
    sibling = _files(model, f'${sibling_id}.')
    assert len(sibling) == 5
    model.remove_hierarchy('T', 'A')
    after = ExcelModelTemplate(model.save(tmp_path / 'after.xlsx'))
    assert [h[1] for h in stored_hierarchies(after)] == ['A.B']
    assert _files(after, f'${sibling_id}.') == sibling
    assert not [n for n in _files(after, '$A.') if '$A.B.' not in n]


def test_removing_column_keeps_dotted_sibling_storage(tmp_path):
    model = ExcelModelTemplate.create({'T': pd.DataFrame({'No': [1, 2], 'No.': [3, 4], 'K': [5, 6]})})
    model = ExcelModelTemplate(model.save(tmp_path / 'before.xlsx'))
    model.remove_column('T', 'No')
    after = ExcelModelTemplate(model.save(tmp_path / 'after.xlsx'))
    assert [c.name for c in after.columns if c.table == 'T' and not c.name.startswith('__')] == ['No.', 'K']
    with after._open_reader() as reader:
        assert reader.get_table('T')['No.'].tolist() == [3, 4]


def test_added_pivot_cache_id_avoids_existing_control_caches(tmp_path):
    model = WorkbookModel.create({'Sales': pd.DataFrame({
        'Category': ['A', 'B'], 'Amount': [10, 20], 'Day': pd.to_datetime(['2026-01-01', '2026-01-02'])})})
    model.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')
    source = (WorkbookReport(model)
              .add_pivot_table(PivotTableSpec('Summary', 'Report', rows=(('Sales', 'Category'),), measures=('Total',)))
              .add_slicer(SlicerSpec('Categories', 'Summary', ('Sales', 'Category')))
              .add_timeline(TimelineSpec('Dates', 'Summary', ('Sales', 'Day')))
              .save(tmp_path / 'source.xlsx'))
    reopened = WorkbookModel(source)
    output = (WorkbookReport(reopened)
              .add_pivot_table(PivotTableSpec('Extra', 'Extra', measures=('Total',)))
              .save(tmp_path / 'extended.xlsx'))
    with ZipFile(output) as archive:
        workbook = minidom.parseString(archive.read('xl/workbook.xml'))
    ids = [n.getAttribute('cacheId') for n in workbook.getElementsByTagNameNS(S, 'pivotCache')]
    assert len(ids) >= 4
    assert len(ids) == len(set(ids))


def _integer_mutation_fixture(tmp_path):
    # Synthetic, non-null integer values exercise value encoding without the
    # private workbook excluded from source distributions and clean checkouts.
    return ExcelModelTemplate.create({
        'TheTable': pd.DataFrame({'K': list(range(2, 1001, 2))})
    }).save(tmp_path / 'integer-source.xlsx')


def test_set_integer_value_rejects_values_outside_existing_range(tmp_path):
    from excel_model_lab.mutate import set_integer_value
    fixture = _integer_mutation_fixture(tmp_path)
    for value in (2**53 + 1, -(2**40)):
        with pytest.raises(ValueError, match='existing range'):
            set_integer_value(fixture, tmp_path / 'out.xlsx', table='TheTable', column='K',
                              position=0, value=value)
    assert not (tmp_path / 'out.xlsx').exists()
    with pytest.raises(TypeError):
        set_integer_value(fixture, tmp_path / 'out.xlsx', table='TheTable', column='K',
                          position=0, value=11.0)


def _fixture_image():
    from pathlib import Path
    from excel_model_lab.abf import AbfImage
    return AbfImage.from_workbook(Path(__file__).parent / 'data' / 'native-authoring-base.xlsx')


def test_abf_rejects_negative_virtual_file_size():
    from dataclasses import replace
    from excel_model_lab.abf import InvalidModelError
    image = _fixture_image()
    image.files = [replace(image.files[0], size=-1)] + image.files[1:]
    with pytest.raises(InvalidModelError, match='Negative'):
        image.validate()


def test_abf_extract_refuses_colliding_names(tmp_path):
    from dataclasses import replace
    from excel_model_lab.abf import InvalidModelError
    image = _fixture_image()
    first = image.files[0]
    image.files = [replace(first, storage_path='a/b'), replace(first, storage_path='A_B')]
    with pytest.raises(InvalidModelError, match='same extracted filename'):
        image.extract(tmp_path / 'out')
    assert not (tmp_path / 'out').exists()


def test_abf_model_part_size_cap(monkeypatch):
    from excel_model_lab import abf
    monkeypatch.setattr(abf, 'MAX_MODEL_PART_BYTES', 1024)
    with pytest.raises(abf.InvalidModelError, match='limit'):
        _fixture_image()


def test_dax_parser_treats_operator_like_literals_as_operands():
    from excel_model_lab.dax_evaluation import DaxUnsupportedError
    from excel_model_lab.dax_validation import Parser
    for text in ('SUM(T[a]) "-" 1', '[A] [+] [B]'):
        with pytest.raises(DaxUnsupportedError):
            Parser(text).expression()
    assert Parser('[A] + [B]').expression().kind == 'binary'


def test_slicer_on_decimal_column_rejected_when_added():
    from excel_model_lab import ModelValidationError
    model = WorkbookModel.create({'S': pd.DataFrame({'a': ['x', 'y'], 'c': [1.5, 2.5]})})
    model.add_measure('S', 'T', 'SUM(S[c])')
    report = WorkbookReport(model).add_pivot_table(PivotTableSpec('P', 'R', measures=('T',)))
    with pytest.raises(ModelValidationError, match='not decimal'):
        report.add_slicer(SlicerSpec('X', 'P', ('S', 'c')))
    report.add_slicer(SlicerSpec('Y', 'P', ('S', 'a')))


def test_spec_relationship_endpoints_allow_dotted_table_names():
    from excel_model_lab.spec import _field_reference
    model = WorkbookModel.create({'Sales.2024': pd.DataFrame({'Key': [1]}), 'Sales': pd.DataFrame({'X': [1]})})
    assert _field_reference(model, 'Sales.2024.Key') == ('Sales.2024', 'Key')
    assert _field_reference(model, 'Sales.X') == ('Sales', 'X')
    assert _field_reference(model, ['Sales.2024', 'Key']) == ('Sales.2024', 'Key')


def test_spec_sort_only_column_entry_skips_configure(tmp_path):
    from excel_model_lab.spec import build_from_spec
    (tmp_path / 'm.csv').write_text('Month,MonthNo\nJan,1\nFeb,2\n', encoding='utf-8')
    spec = {'tables': {'M': {'source': 'm.csv'}},
            'calculated_columns': [{'table': 'M', 'name': 'Label', 'expression': '[Month]', 'data_type': 'text'}]}
    calls = []
    original = WorkbookModel.configure_column
    def spy(self, *args, **kwargs):
        calls.append(args)
        return original(self, *args, **kwargs)
    spec['tables']['M']['columns'] = {'Month': {'sort_by': 'MonthNo'}}
    WorkbookModel.configure_column = spy
    try:
        build_from_spec(None, tmp_path / 'out.xlsx', spec, base_directory=tmp_path)
    finally:
        WorkbookModel.configure_column = original
    assert calls == []


def _xml_texts(image, fragment):
    return [image.read_logical_file(f.storage_path).decode('utf-8', 'replace') for f in image.files
            if f.logical_path and f.logical_path.endswith('.xml') and fragment in f.logical_path]


def test_clone_table_keeps_captions_containing_the_table_id():
    from excel_model_lab.schema import clone_table
    model = ExcelModelTemplate.create({'Sales': pd.DataFrame({'Sales Amount': [1, 2], 'K': [3, 4]})})
    # create() assigns a generated physical ID; clone once to get a table
    # whose physical ID is exactly 'Sales', the shape the defect needs.
    image = clone_table(model._image, model._dimension_ids['Sales'], 'Sales')
    assert any('<ID>Sales</ID>' in t for t in _xml_texts(image, chr(92) + 'Sales.'))
    image = clone_table(image, 'Sales', 'CopyTable')
    texts = _xml_texts(image, 'CopyTable')
    assert texts
    assert any('<Name>Sales Amount</Name>' in t for t in texts)
    assert all('CopyTable Amount' not in t for t in texts)


def _sales_model():
    model = WorkbookModel.create({'Sales': pd.DataFrame({'Category': ['A', 'B'], 'Amount': [10, 20]})})
    model.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')
    return model


def test_slicer_sheet_name_matches_case_insensitively(tmp_path):
    report = (WorkbookReport(_sales_model())
              .add_pivot_table(PivotTableSpec('Summary', 'Report', rows=(('Sales', 'Category'),), measures=('Total',)))
              .add_slicer(SlicerSpec('Categories', 'Summary', ('Sales', 'Category'), sheet='report')))
    output = report.save(tmp_path / 'case.xlsx')
    with ZipFile(output) as archive:
        workbook = minidom.parseString(archive.read('xl/workbook.xml'))
    assert [n.getAttribute('name') for n in workbook.getElementsByTagNameNS(S, 'sheet')].count('Report') == 1


def test_lowercase_pivot_cell_is_written_uppercase(tmp_path):
    output = (WorkbookReport(_sales_model())
              .add_pivot_table(PivotTableSpec('Summary', 'Report', cell='c5', measures=('Total',)))
              .save(tmp_path / 'cell.xlsx'))
    with ZipFile(output) as archive:
        pivot = next(n for n in archive.namelist() if n.startswith('xl/pivotTables/') and n.endswith('.xml'))
        location = minidom.parseString(archive.read(pivot)).getElementsByTagNameNS(S, 'location')[0]
    assert location.getAttribute('ref') == 'C5:C5'


def test_set_integer_value_ignores_the_callers_decimal_context(tmp_path):
    from decimal import localcontext
    from pbixray import PBIXRay
    from excel_model_lab.mutate import set_integer_value
    fixture = _integer_mutation_fixture(tmp_path)
    with PBIXRay(fixture) as reader:
        original = reader.get_table('TheTable')['K']
    value = 123  # three significant digits; the column holds even values 2..1000
    assert original.min() <= value <= original.max() and value not in set(original)
    with localcontext() as ctx:
        ctx.prec = 2
        output = set_integer_value(fixture, tmp_path / 'out.xlsx', table='TheTable', column='K',
                                   position=0, value=value)
    with PBIXRay(output) as reader:
        actual = reader.get_table('TheTable')['K'].tolist()
    expected = original.tolist()
    expected[0] = value
    assert actual == expected


def test_remove_renamed_column_ignores_measures_on_a_reused_caption(tmp_path):
    from excel_model_lab.schema import _child, _elements, _text, remove_column
    model = ExcelModelTemplate.create({'T': pd.DataFrame({'A': [1, 2], 'B': [3, 4], 'C': [5, 6]})})
    model.add_measure('T', 'Total', 'SUM(T[A])')
    model = ExcelModelTemplate(model.save(tmp_path / 'm.xlsx'))  # measures are staged until save
    table_id = model._dimension_ids['T']
    ids = {caption: model._column_ids[('T', caption)] for caption in 'ABC'}
    # Simulate native renames A->X and B->A: physical IDs stay, captions move,
    # so T[A] in the measure now means physical column B.
    image, changes = model._image, {}
    for f in image.files:
        if not (f.logical_path and f.logical_path.endswith('.dim.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(f.storage_path))
        for attribute in _elements(doc, 'Attribute'):
            if attribute.parentNode.localName != 'Attributes':
                continue
            node = _child(attribute, 'Name')
            caption = {ids['A']: 'X', ids['B']: 'A'}.get(_text(_child(attribute, 'ID')))
            if caption:
                node.firstChild.data = caption
                changes[f.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    assert changes
    from excel_model_lab.abf import AbfImage
    image = AbfImage(image.assemble(changes))
    remove_column(image, table_id, ids['A'])  # caption X: unreferenced
    with pytest.raises(Exception, match='used by a calculation'):
        remove_column(image, table_id, ids['B'])  # caption A: referenced
