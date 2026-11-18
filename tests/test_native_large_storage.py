"""Byte-level references captured from an Excel-created synthetic ID1..10000 table.

The fixtures contain only numeric storage streams, without workbook metadata.
"""
from pathlib import Path

from excel_model_lab.hash_index import encode_numeric_hash_index
from excel_model_lab.vertipaq import encode_pure_bitpacked_ids

DATA = Path(__file__).parent / "data"


def test_large_hash_matches_excel_including_overflow_buckets():
    native = (DATA / "native-10000.hidx").read_bytes()
    assert encode_numeric_hash_index(native, list(range(3, 10003))) == native


def test_packed_word_boundary_matches_excel():
    native = (DATA / "native-10000.idf").read_bytes()
    assert encode_pure_bitpacked_ids(
        range(3, 10003), bit_width=16, min_data_id=3
    ) == native
