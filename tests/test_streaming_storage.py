"""Storage edge cases and bounded-memory segment writing."""
from io import BytesIO
import struct
import tracemalloc

import pytest

from excel_model_lab.vertipaq import encode_pure_bitpacked_ids, write_pure_bitpacked_ids


def unpack(blob, width, minimum):
    entries = struct.unpack_from('<Q', blob)[0]
    if not entries:
        assert blob == bytes(16)
        return []
    rows = struct.unpack_from('<I', blob, 12)[0]
    start = 16 + entries * 8
    per = 64 // width
    return [minimum + ((struct.unpack_from('<Q', blob, start + (i // per) * 8)[0]
                        >> ((i % per) * width)) & ((1 << width) - 1))
            for i in range(rows)]


@pytest.mark.parametrize('width', [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 16, 21, 32])
@pytest.mark.parametrize('count', [0, 1, 2, 3, 63, 64, 65, 1024])
def test_streaming_boundaries(width, count):
    maximum = min((1 << width) - 1, 0xFFFFFFFF - 3)
    values = [3 + (i % (maximum + 1)) for i in range(count)]
    output = BytesIO()
    size = write_pure_bitpacked_ids(output, iter(values), record_count=count,
                                   bit_width=width, min_data_id=3)
    blob = output.getvalue()
    assert size == len(blob)
    assert unpack(blob, width, 3) == values
    assert blob == encode_pure_bitpacked_ids(values, bit_width=width, min_data_id=3)


@pytest.mark.parametrize('ids', [[2] * 100, [3], [9000] * 65])
def test_constant_and_null_streams(ids):
    assert unpack(encode_pure_bitpacked_ids(ids, bit_width=1, min_data_id=ids[0]),
                  1, ids[0]) == ids


@pytest.mark.parametrize('ids,count', [([], 1), ([3, 4], 1), ([None], 0)])
def test_incorrect_counts_rejected(ids, count):
    with pytest.raises(ValueError, match='data IDs'):
        write_pure_bitpacked_ids(BytesIO(), iter(ids), record_count=count,
                                bit_width=2, min_data_id=3)


def test_fractional_ids_are_not_silently_truncated():
    with pytest.raises(TypeError):
        encode_pure_bitpacked_ids([3.5], bit_width=2, min_data_id=3)


def test_unsigned_overflow_rejected():
    with pytest.raises(ValueError, match='data ID'):
        encode_pure_bitpacked_ids([1 << 32], bit_width=32, min_data_id=3)


def test_streaming_million_rows_has_bounded_scratch_space():
    class Discard:
        size = 0
        def write(self, value):
            self.size += len(value)
    sink = Discard()
    tracemalloc.start()
    try:
        size = write_pure_bitpacked_ids(sink, range(3, 1_000_003), record_count=1_000_000,
                                       bit_width=21, min_data_id=3)
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    assert size == sink.size
    assert peak < 300_000


@pytest.mark.parametrize('values', [[], [' '], ['a', 'b'], ['a', '\U0001f600', 'caf\u00e9'],
                                   [f'item-{i}' for i in range(1000)]])
def test_streaming_string_dictionary_matches_independent_decoder(values):
    from pathlib import Path
    from excel_model_lab import WorkbookModel
    from excel_model_lab.dictionary import write_uncompressed_string_dictionary, encode_uncompressed_string_dictionary
    from excel_model_lab.reader_adapter import decode_dictionary
    model = WorkbookModel(Path(__file__).parents[1] / 'src/excel_model_lab/data/authoring-seed.xlsx')
    template = model._image.read_logical_file(model._encodings['Table1', 'Label'].dictionary_path)
    output = BytesIO()
    size = write_uncompressed_string_dictionary(output, template, values)
    assert size == len(output.getvalue())
    assert output.getvalue() == encode_uncompressed_string_dictionary(template, values)
    with model._open_reader() as reader:
        decoded = decode_dictionary(reader, output.getvalue(), 3)
    assert [decoded.values[i] for i in sorted(decoded.values)] == values


def test_numeric_vector_extrema_and_invalid_values():
    from excel_model_lab.dictionary import encode_numeric_dictionary
    from pbixray.column_data.dictionary import ColumnDataDictionary
    template = struct.pack('<i6iQI', 0, *([0] * 6), 0, 8)
    values = [-(1 << 63), (1 << 63) - 1]
    blob = encode_numeric_dictionary(template, values)
    assert blob[-16:] == struct.pack('<qq', *values)
    assert int(ColumnDataDictionary.from_io(BytesIO(blob)).dictionary_type) == 0
    for invalid in [1 << 63, 10 ** 1000, 1.5, float('nan'), float('inf')]:
        with pytest.raises(ValueError):
            encode_numeric_dictionary(template, [invalid])


@pytest.mark.parametrize('page_size', [60, 70, 128, 1024])
def test_multiple_string_pages_handles_and_unicode(page_size):
    from pathlib import Path
    from excel_model_lab import WorkbookModel
    from excel_model_lab.dictionary import write_uncompressed_string_dictionary
    from excel_model_lab.reader_adapter import decode_dictionary
    from pbixray.column_data.dictionary import ColumnDataDictionary
    model = WorkbookModel(Path(__file__).parents[1] / 'src/excel_model_lab/data/authoring-seed.xlsx')
    template = model._image.read_logical_file(model._encodings['Table1', 'Label'].dictionary_path)
    values = [' ', '\U0001f600', 'x' * 700] + [f'{i}-caf\u00e9' for i in range(50)]
    output = BytesIO()
    assert write_uncompressed_string_dictionary(output, template, values, page_size_bytes=page_size) == len(output.getvalue())
    parsed = ColumnDataDictionary.from_io(BytesIO(output.getvalue()))
    pages = parsed.data.dictionary_pages
    assert len(pages) > 1
    assert not parsed.data.page_layout_information.f_store_compressed
    handles = parsed.data.dictionary_record_handles_vector_info.vector_of_record_handle_structures
    for index, value in enumerate(values):
        offset, page_id = struct.unpack_from('<II', handles, index * 8)
        page = pages[page_id]
        assert page.page_start_index <= index < page.page_start_index + page.page_string_count
        assert offset >= 0
    with model._open_reader() as reader:
        decoded = decode_dictionary(reader, output.getvalue(), 3)
    assert [decoded.values[i] for i in sorted(decoded.values)] == values


def test_paged_model_roundtrip_preserves_nulls_duplicates_and_id_growth(tmp_path, monkeypatch):
    import pandas as pd
    from functools import partial
    from excel_model_lab import WorkbookModel
    from excel_model_lab import model as model_module
    from excel_model_lab.dictionary import encode_uncompressed_string_dictionary
    monkeypatch.setattr(model_module, 'encode_uncompressed_string_dictionary',
                        partial(encode_uncompressed_string_dictionary, page_size_bytes=128))
    values = [' '] + [f'{i}-\U0001f600-\u6771\u4eac' for i in range(40)] + [None, '0-\U0001f600-\u6771\u4eac', ' ']
    model = WorkbookModel.create({'Words': pd.DataFrame({'Label': values})})
    first = model.save(tmp_path / 'pages.xlsx')
    reopened = WorkbookModel(first)
    actual = reopened.get_table('Words').Label.tolist()
    assert [pd.isna(v) for v in actual] == [v is None for v in values]
    assert [v for v in actual if not pd.isna(v)] == [v for v in values if v is not None]
    changed = ['new-\u00e9'] + values[1:]
    output = reopened.replace_column('Words', 'Label', changed).save(tmp_path / 'grown-pages.xlsx')
    again = WorkbookModel(output)
    assert set(again._encodings['Words', 'Label'].original_values) == {v for v in changed if v is not None}
    assert again.get_table('Words').Label.iloc[0] == 'new-\u00e9'


def test_paged_dictionary_rejects_duplicate_values_before_output():
    from pathlib import Path
    from excel_model_lab import WorkbookModel
    from excel_model_lab.dictionary import write_uncompressed_string_dictionary
    model = WorkbookModel(Path(__file__).parents[1] / 'src/excel_model_lab/data/authoring-seed.xlsx')
    template = model._image.read_logical_file(model._encodings['Table1', 'Label'].dictionary_path)
    output = BytesIO()
    with pytest.raises(ValueError, match='unique'):
        write_uncompressed_string_dictionary(output, template, ['\U0001f600', 'x'*100, '\U0001f600'], page_size_bytes=64)
    assert output.getvalue() == b''
