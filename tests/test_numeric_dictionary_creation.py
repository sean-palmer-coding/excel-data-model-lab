from decimal import Decimal

import pandas as pd
import pytest

from excel_model_lab import WorkbookModel
from excel_model_lab.model import _NumericDictionaryEncoding


@pytest.mark.parametrize('values', [
    pd.Series([-(1 << 63) + 1, 0, (1 << 63) - 2], dtype='int64'),
    pd.Series([-(1 << 63) + 1, None, (1 << 63) - 2], dtype='Int64'),
    pd.Series([Decimal('-922337203685477.5807'), Decimal('0.0001'), Decimal('922337203685477.5806')]),
    pd.Series([-1e100, 1e-100, 1e100]),
])
def test_wide_numeric_create_read_edit_and_stream(tmp_path, values):
    m = WorkbookModel.create({'Numbers': pd.DataFrame({'Value': values})})
    assert isinstance(m._encodings['Numbers', 'Value'], _NumericDictionaryEncoding)
    first = m.save(tmp_path / 'wide.xlsx')
    with WorkbookModel.open(first) as reader:
        actual = reader.get_table('Numbers')['Value'].tolist()
        chunks = pd.concat(list(reader.iter_table('Numbers', chunk_size=1)))['Value'].tolist()
    for observed in (actual, chunks):
        assert [pd.isna(v) for v in observed] == values.isna().tolist()
        assert [v for v in observed if not pd.isna(v)] == values.dropna().tolist()
    reordered = values.iloc[::-1].reset_index(drop=True)
    second = WorkbookModel(first).replace_table('Numbers', pd.DataFrame({'Value': reordered})).save(tmp_path / 'edited.xlsx')
    observed = WorkbookModel(second).get_table('Numbers')['Value'].tolist()
    assert [v for v in observed if not pd.isna(v)] == reordered.dropna().tolist()


@pytest.mark.parametrize('values', [
    pd.Series([1 << 63], dtype='uint64'),
    pd.Series([Decimal('922337203685477.5808')]),
    pd.Series([Decimal('0.00001')]),
    pd.Series([float('inf')]),
])
def test_invalid_numeric_ranges_rejected_before_save(values):
    with pytest.raises(ValueError):
        WorkbookModel.create({'Numbers': pd.DataFrame({'Value': values})})


def test_small_numeric_creation_retains_existing_value_encoding():
    model = WorkbookModel.create({'Numbers': pd.DataFrame({'Value': [1, 2, 3]})})
    assert not isinstance(model._encodings['Numbers', 'Value'], _NumericDictionaryEncoding)


@pytest.mark.parametrize('dtype', ['int64', 'Int64'])
def test_excel_integer_minimum_rejected_before_authoring(dtype):
    with pytest.raises(ValueError, match='blank numeric aggregates'):
        WorkbookModel.create({'T': pd.DataFrame({'Value': pd.Series([-2**63, 0, 1], dtype=dtype)})})


@pytest.mark.parametrize('source', [[1, 2, 3], [-2**63+1, 0, 2**63-2]])
def test_excel_integer_minimum_replacement_is_atomic(source):
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': source})})
    before, pending = model._image.data, dict(model._pending)
    with pytest.raises(ValueError, match='blank numeric aggregates'):
        model.replace_column('T', 'Value', [-2**63, 0, 1])
    assert model._image.data == before and model._pending == pending


@pytest.mark.parametrize('source', [[str(-2**63), '0', '1'], [1, 2, 3]])
def test_excel_integer_minimum_conversion_is_atomic(source):
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': source})})
    before, pending = model._image.data, dict(model._pending)
    with pytest.raises(ValueError, match='blank numeric aggregates'):
        model.change_column_type('T', 'Value', 'integer', values=[-2**63, 0, 1])
    assert model._image.data == before and model._pending == pending


def test_excel_currency_minimum_rejected_before_authoring():
    with pytest.raises(ValueError, match='blank MIN/MAX aggregates'):
        WorkbookModel.create({'T': pd.DataFrame({'Value': [Decimal('-922337203685477.5808')]})})


@pytest.mark.parametrize('source', [[Decimal('1'), Decimal('2')],
    [Decimal('-922337203685477.5807'), Decimal('922337203685477.5806')]])
def test_excel_currency_minimum_replacement_is_atomic(source):
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': source})})
    before, pending = model._image.data, dict(model._pending)
    with pytest.raises(ValueError, match='blank MIN/MAX aggregates'):
        model.replace_column('T', 'Value', [Decimal('-922337203685477.5808'), Decimal('0')])
    assert model._image.data == before and model._pending == pending


def test_excel_currency_minimum_conversion_is_atomic():
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': ['-922337203685477.5808', '0']})})
    before, pending = model._image.data, dict(model._pending)
    with pytest.raises(ValueError, match='blank MIN/MAX aggregates'):
        model.change_column_type('T', 'Value', 'currency')
    assert model._image.data == before and model._pending == pending


@pytest.mark.parametrize('maximum', [2**63-1, Decimal('922337203685477.5807')])
def test_excel_maximum_rejected_before_authoring(maximum):
    with pytest.raises(ValueError, match='blank MAX/SUM aggregates'):
        WorkbookModel.create({'T': pd.DataFrame({'Value': [maximum]})})


@pytest.mark.parametrize('source,maximum', [
    ([1, 2], 2**63-1), ([-2**63+1, 2**63-2], 2**63-1),
    ([Decimal('1'), Decimal('2')], Decimal('922337203685477.5807')),
    ([Decimal('-922337203685477.5807'), Decimal('922337203685477.5806')], Decimal('922337203685477.5807')),
])
def test_excel_maximum_replacement_is_atomic(source, maximum):
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': source})})
    before, pending = model._image.data, dict(model._pending)
    with pytest.raises(ValueError, match='blank MAX/SUM aggregates'):
        model.replace_column('T', 'Value', [maximum, 0])
    assert model._image.data == before and model._pending == pending


@pytest.mark.parametrize('source', [['1', '2'], [1, 2]])
@pytest.mark.parametrize('kind,maximum', [('integer', 2**63-1), ('currency', Decimal('922337203685477.5807'))])
def test_excel_maximum_conversion_is_atomic(source, kind, maximum):
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': source})})
    before, pending = model._image.data, dict(model._pending)
    with pytest.raises(ValueError, match='blank MAX/SUM aggregates'):
        model.change_column_type('T', 'Value', kind, values=[maximum, 0])
    assert model._image.data == before and model._pending == pending


@pytest.mark.parametrize('currency', [False, True])
def test_existing_signed64_endpoints_remain_readable(tmp_path, currency):
    from excel_model_lab.abf import AbfImage
    from excel_model_lab.dictionary import encode_numeric_dictionary
    low, high = -2**63+1, 2**63-2
    values = [Decimal(low) / 10000, Decimal(0), Decimal(high) / 10000] if currency else [low, 0, high]
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': values})})
    model = WorkbookModel(model.save(tmp_path / 'existing.xlsx'))
    # Model a pre-existing workbook, independently of authoring validation.
    encoding = model._encodings['T', 'Value']
    entries = [-2**63 if value == low else 2**63-1 if value == high else value
               for value in encoding.original_values]
    payload = encode_numeric_dictionary(model._image.read_logical_file(encoding.dictionary_path), entries)
    existing = WorkbookModel(model.source, _image=AbfImage(model._image.assemble({encoding.dictionary_path: payload})))
    expected = [Decimal(-2**63) / 10000, Decimal(0), Decimal(2**63-1) / 10000] if currency else [-2**63, 0, 2**63-1]
    assert existing.get_table('T').Value.tolist() == expected


def test_numeric_conversion_uses_numeric_hash_policy():
    import struct
    model = WorkbookModel.create({'Numbers': pd.DataFrame({'Value': [-1e100, 1e100]})})
    for file in model._image.files:
        if file.logical_path and file.logical_path.endswith('.dictionary'):
            payload = model._image.read_logical_file(file)
            if struct.unpack('<i', payload[:4])[0] == 1:
                assert struct.unpack('<6i', payload[4:28]) == (-1, 16, 64, 3, -1, -1)


def test_existing_value_column_promotes_without_losing_metadata(tmp_path):
    model = WorkbookModel.create({'Numbers': pd.DataFrame({'ID': [1, 2, 3], 'Value': [1, 2, 3]})})
    model.add_measure('Numbers', 'Total', 'SUM(Numbers[Value])')
    original = model.save(tmp_path / 'original.xlsx')
    updated = WorkbookModel(original).replace_column('Numbers', 'Value', [-2**63 + 1, 0, 2**63 - 2])
    assert isinstance(updated._encodings['Numbers', 'Value'], _NumericDictionaryEncoding)
    output = updated.save(tmp_path / 'promoted.xlsx')
    reopened = WorkbookModel(output)
    assert reopened.get_table('Numbers')['Value'].tolist() == [-2**63 + 1, 0, 2**63 - 2]
    assert reopened.get_table('Numbers')['ID'].tolist() == [1, 2, 3]
    from excel_model_lab.measures import list_script_measures
    assert any(m.name == 'Total' for file in reopened._image.files
               if file.logical_path and file.logical_path.endswith('.xml')
               for m in list_script_measures(reopened._image.read_logical_file(file)))


def test_numeric_promotion_batch_failure_is_atomic(tmp_path):
    model = WorkbookModel.create({'Numbers': pd.DataFrame({'Value': [1, 2, 3], 'Other': [4, 5, 6]})})
    before = model.get_table('Numbers')
    with pytest.raises(ValueError):
        model.replace_columns('Numbers', {'Value': [-2**63 + 1, 0, 2**63 - 2], 'Other': [1]})
    assert model.get_table('Numbers').equals(before)
    assert not isinstance(model._encodings['Numbers', 'Value'], _NumericDictionaryEncoding)


def test_promotion_preserves_linked_source_and_worksheet_parts(tmp_path):
    from pathlib import Path
    from zipfile import ZipFile
    source = Path(__file__).parent / 'data/native-authoring-base.xlsx'
    model = WorkbookModel(source)
    before = model.list_native_sources()
    output = model.replace_column('Table1', 'ID', [-2**63 + 1, 0, 2**63 - 2]).save(tmp_path / 'linked-promoted.xlsx')
    reopened = WorkbookModel(output)
    assert reopened.list_native_sources() == before
    with ZipFile(source) as first, ZipFile(output) as second:
        for name in ('xl/worksheets/sheet1.xml', 'xl/tables/table1.xml'):
            assert first.read(name) == second.read(name)


def test_promotion_preserves_date_perspective_hierarchy_and_calculated_column(tmp_path):
    dates = pd.date_range('2026-01-01', periods=3)
    model = WorkbookModel.create({'Calendar': pd.DataFrame({'Date': dates, 'Value': [1, 2, 3]})})
    model.mark_date_table('Calendar', 'Date')
    model.add_hierarchy('Calendar', 'Dates', ['Date'])
    model.add_calculated_column('Calendar', 'Computed', '[Value]+1', data_type='integer')
    model.create_perspective('View', columns=[('Calendar', 'Value')], hierarchies=[('Calendar', 'Dates')])
    model.replace_column('Calendar', 'Value', [-2**63 + 1, 0, 2**63 - 2])
    output = model.save(tmp_path / 'metadata-promoted.xlsx')
    reopened = WorkbookModel(output)
    assert reopened.list_perspectives() == ('View',)
    assert any(edge.reason == 'date table key' for edge in reopened.dependency_graph().edges)
    calculated, = reopened.list_calculated_columns()
    assert calculated.expression == '[Value]+1' and calculated.processing_state == 'Unprocessed'
    frame = reopened.get_table('Calendar')
    assert frame.Computed.isna().all()
    assert frame.Date.tolist() == dates.tolist()


def test_promotion_preserves_relationship_when_both_keys_change(tmp_path):
    from excel_model_lab.relationship_storage import stored_relationships
    model = WorkbookModel.create({'Keys': pd.DataFrame({'ID': [1, 2, 3]}),
                                  'Facts': pd.DataFrame({'ID': [1, 2, 1]})})
    model.add_relationship('Facts', 'ID', 'Keys', 'ID')
    model = WorkbookModel(model.save(tmp_path / 'related-before.xlsx'))
    wide = [-2**63 + 1, 0, 2**63 - 2]
    model.replace_tables({'Keys': pd.DataFrame({'ID': wide}),
                          'Facts': pd.DataFrame({'ID': [wide[0], wide[1], wide[0]]})})
    reopened = WorkbookModel(model.save(tmp_path / 'related-after.xlsx'))
    assert len(stored_relationships(reopened)) == 1
    assert reopened.get_table('Keys').ID.tolist() == wide
    assert reopened.get_table('Facts').ID.tolist() == [wide[0], wide[1], wide[0]]


@pytest.mark.parametrize('source,target,expected', [
    (['1', None, '-2'], 'integer', [1, None, -2]),
    (['1.25', None, '-2.5'], 'decimal', [1.25, None, -2.5]),
    (['1.2500', None, '-2.5000'], 'currency', [Decimal('1.25'), None, Decimal('-2.5')]),
    (['true', None, '0'], 'boolean', [True, None, False]),
    (['2026-01-01', None, '2026-01-02'], 'datetime', [pd.Timestamp('2026-01-01'), None, pd.Timestamp('2026-01-02')]),
    ([-2**63 + 1, 0, 2**63 - 2], 'text', [str(-2**63+1), '0', str(2**63-2)]),
])
def test_dictionary_type_conversion_roundtrip(tmp_path, source, target, expected):
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': source})})
    model.change_column_type('T', 'Value', target)
    actual = WorkbookModel(model.save(tmp_path / f'{target}.xlsx')).get_table('T').Value.tolist()
    assert [pd.isna(v) for v in actual] == [v is None for v in expected]
    assert [v for v in actual if not pd.isna(v)] == [v for v in expected if v is not None]


def test_dictionary_conversion_precision_and_atomic_failure():
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': ['1.25', '2.5']})})
    before, pending = model._image.data, dict(model._pending)
    with pytest.raises(ValueError, match='not exactly representable'):
        model.change_column_type('T', 'Value', 'integer')
    assert model._image.data == before and model._pending == pending
    with pytest.raises(ValueError, match='not exactly representable'):
        model.change_column_type('T', 'Value', 'decimal', decimal_places=1)
    assert model._image.data == before and model._pending == pending


def test_dictionary_conversion_all_null_currency(tmp_path):
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': pd.Series([None, None], dtype='string')})})
    model.change_column_type('T', 'Value', 'currency')
    reopened = WorkbookModel(model.save(tmp_path / 'null-currency.xlsx'))
    assert reopened.get_table('T').Value.isna().all()
    from excel_model_lab.relationship_storage import relationship_key_type
    assert relationship_key_type(reopened, 'T', 'Value') == 'Currency'


def test_dictionary_conversion_relationship_and_date_guards():
    model = WorkbookModel.create({'K': pd.DataFrame({'ID': ['a', 'b']}),
                                  'F': pd.DataFrame({'ID': ['a', 'a']})})
    model.add_relationship('F', 'ID', 'K', 'ID')
    with pytest.raises(ValueError, match='relationships'):
        model.change_column_type('K', 'ID', 'integer', values=[1, 2])
    calendar = WorkbookModel.create({'D': pd.DataFrame({'Date': pd.date_range('2026-01-01', periods=2)})})
    calendar.mark_date_table('D', 'Date')
    with pytest.raises(ValueError, match='Unmark'):
        calendar.change_column_type('D', 'Date', 'text')


@pytest.mark.parametrize('values', [[1, 2, 3], ['1', '2', '3'], [-2**63+1, 0, 2**63-2]])
def test_conversion_preserves_calculated_columns_and_perspective(tmp_path, values):
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': values})})
    model.add_calculated_column('T', 'Computed', '[Value]+1', data_type='integer')
    model.create_perspective('View', columns=[('T', 'Value')])
    model.change_column_type('T', 'Value', 'decimal', values=[1.25, 2.5, 3.75], decimal_places=2)
    reopened = WorkbookModel(model.save(tmp_path / 'converted-metadata.xlsx'))
    assert reopened.get_table('T').Value.tolist() == [1.25, 2.5, 3.75]
    assert reopened.get_table('T').Computed.isna().all()
    assert reopened.list_perspectives() == ('View',)
    calculated, = reopened.list_calculated_columns()
    assert calculated.expression == '[Value]+1'
    assert calculated.processing_state == 'Unprocessed'
    with pytest.raises(ValueError, match='not exactly representable'):
        reopened.replace_column('T', 'Value', [1.251, 2.5, 3.75])


@pytest.mark.parametrize('values', [[1, 2, 3], ['1', '2', '3'], [-2**63+1, 0, 2**63-2]])
def test_conversion_resolves_visible_caption_to_physical_id(tmp_path, values):
    from xml.dom import minidom
    from excel_model_lab.abf import AbfImage
    from excel_model_lab.schema import _elements, _child, _text
    model = WorkbookModel.create({'T': pd.DataFrame({'Physical': values})})
    changes = {}
    for file in model._image.files:
        if not file.logical_path or not file.logical_path.endswith('.dim.xml'):
            continue
        doc = minidom.parseString(model._image.read_logical_file(file))
        for attribute in _elements(doc, 'Attribute'):
            ids = [n for n in attribute.childNodes if getattr(n, 'localName', None) == 'ID']
            if ids and _text(ids[0]) == 'Physical':
                _child(attribute, 'Name').firstChild.data = 'Visible'
                changes[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    model = WorkbookModel(model.source, _image=AbfImage(model._image.assemble(changes)))
    assert model.get_table('T').columns.tolist() == ['Visible']
    model.change_column_type('T', 'Visible', 'decimal', decimal_places=2, values=[1.25, 2.5, 3.75])
    reopened = WorkbookModel(model.save(tmp_path / 'caption-converted.xlsx'))
    assert reopened.get_table('T').Visible.tolist() == [1.25, 2.5, 3.75]
