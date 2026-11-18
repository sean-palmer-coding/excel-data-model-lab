from pathlib import Path
import os

import pytest
from pbixray import PBIXRay

from excelmodel import WorkbookModel
from excel_model_lab.model import ModelValidationError
from excel_model_lab.package import repack_workbook

FIXTURE = Path(__file__).parent / "data" / "null_data_id.xlsx"


@pytest.mark.parametrize('data_type', ['int64', 'string'])
def test_native_empty_column_without_dictionary_is_inspectable(data_type):
    metadata = {'DataType': data_type, 'Dictionary': None, 'HIDX': None}
    reason = WorkbookModel._unsupported_reason(metadata, [{'records': 0, 'bit_width': 1}])
    assert 'native empty column' in reason


def test_numeric_dictionary_growth_rebuilds_member_order(tmp_path):
    from excel_model_lab.hierarchy_storage import intrinsic_order

    model = WorkbookModel(FIXTURE.with_name('supplier-quality.xlsx'))
    frame = model.get_table('Defect Type')
    frame['Defect Type ID'] = [-17.5, 3.0, 1.0]
    output = model.replace_table('Defect Type', frame).save(tmp_path / 'grown.xlsx')
    reopened = WorkbookModel(output)
    assert reopened.get_table('Defect Type')['Defect Type ID'].tolist() == [-17.5, 3.0, 1.0]
    # IDs may compact, but member order must follow the actual numeric values.
    order = intrinsic_order(reopened, 'Defect Type', 'Defect Type ID')
    entries = reopened._encodings['Defect Type', 'Defect Type ID'].original_values
    assert [entries[i - 3] for i in sorted(order, key=order.get) if i >= 3] == [-17.5, 1.0, 3.0]
    # A second save must retain the new member and its ordering.
    again = reopened.save(tmp_path / 'again.xlsx')
    assert WorkbookModel(again).get_table('Defect Type').equals(reopened.get_table('Defect Type'))


def test_supplier_resize_compacts_dictionaries_without_losing_values(tmp_path):
    import pandas as pd
    from excel_model_lab.abf import AbfImage

    fixture = FIXTURE.with_name('supplier-quality.xlsx')
    with PBIXRay(fixture) as reader:
        tables = {name: reader.get_table(name).iloc[:count].reset_index(drop=True)
                  for name, count in [('Plant', 10), ('Defect Type', 2)]}
    output = WorkbookModel(fixture).replace_tables(tables).save(tmp_path / 'supplier.xlsx')
    reopened = WorkbookModel(output)
    for name, expected in tables.items():
        for column in expected:
            encoding = reopened._encodings.get((name, column))
            if hasattr(encoding, 'original_values'):
                assert len(encoding.original_values) == expected[column].nunique(dropna=True)
    with PBIXRay(output) as reader:
        for name, expected in tables.items():
            pd.testing.assert_frame_equal(reader.get_table(name).reset_index(drop=True), expected)


@pytest.mark.parametrize('required, expected', [(11, 12), (13, 16), (17, 21), (22, 32), (32, 32)])
def test_bit_width_rounds_to_supported_codec(required, expected):
    from excel_model_lab.vertipaq import select_bit_width
    assert select_bit_width((1 << required) - 1) == expected


def test_combined_resize_uses_supported_compression_classes(tmp_path):
    from excel_model_lab.abf import AbfImage
    from excel_model_lab.vertipaq import SUPPORTED_BIT_WIDTHS
    import re

    with PBIXRay(FIXTURE) as reader:
        frame = reader.get_table('TheTable')[['A', 'N', 'C', 'S', 'K']].iloc[40:80].reset_index(drop=True)
    frame.loc[0, 'S'] = 'New text test'
    output = WorkbookModel(FIXTURE).replace_table('TheTable', frame).save(tmp_path / 'combined.xlsx')
    image = AbfImage.from_workbook(output)
    widths = []
    for f in image.files:
        if f.logical_path and f.logical_path.endswith('.tbl.xml'):
            widths.extend(int(w) for w in re.findall(rb'XMRENoSplitCompressionInfo&lt;(\d+)', image.read_logical_file(f)))
    assert 12 in widths
    assert set(widths) <= set(SUPPORTED_BIT_WIDTHS)


def test_subset_resize_rebuilds_dictionary_members(tmp_path):
    from excel_model_lab.abf import AbfImage
    import pandas as pd

    with PBIXRay(FIXTURE) as reader:
        expected = reader.get_table('TheTable')[['A', 'N', 'C', 'S', 'K']].iloc[40:80].reset_index(drop=True)
    output = WorkbookModel(FIXTURE).replace_table('TheTable', expected).save(tmp_path / 'subset.xlsx')
    reopened = WorkbookModel(output)
    assert set(reopened._encodings['TheTable', 'S'].original_values) == set(expected.S.dropna())
    with PBIXRay(output) as reader:
        pd.testing.assert_frame_equal(reader.get_table('TheTable')[list(expected)].reset_index(drop=True), expected)


def test_large_metadata_uses_excel_frame_size(tmp_path):
    import struct
    from excel_model_lab.abf import AbfImage

    before = AbfImage.from_workbook(FIXTURE)
    f = next(f for f in before.files if f.logical_path and
             f.logical_path.endswith('.tbl.xml') and 'H$' not in f.logical_path)
    logical = before.read_logical_file(f)
    after = AbfImage(before.assemble({f.storage_path: logical}))
    assert after.read_logical_file(f.storage_path) == logical
    payload = after.read_file(f.storage_path)
    if after.error_code:
        payload = payload[:-4]
    offset = 0
    while offset < len(payload):
        size, stored = struct.unpack_from('<HH', payload, offset)
        assert 0 < size <= 4096
        offset += 4 + stored
    assert offset == len(payload)


@pytest.mark.parametrize("existing", [False, True])
def test_failed_validation_does_not_commit_output(tmp_path, monkeypatch, existing):
    output = tmp_path / "report.xlsx"
    if existing:
        output.write_bytes(b"existing report")
    model = WorkbookModel.from_template(FIXTURE)

    def reject(path):
        raise ModelValidationError("deliberate validation failure")

    monkeypatch.setattr(model, "_validate_output", reject)
    with pytest.raises(ModelValidationError, match="deliberate"):
        model.save(output)
    if existing:
        assert output.read_bytes() == b"existing report"
    else:
        assert not output.exists()
    assert not list(tmp_path.glob(".excelmodel-*"))


def test_save_can_replace_source_after_validation(tmp_path):
    output = tmp_path / "report.xlsx"
    output.write_bytes(FIXTURE.read_bytes())
    with PBIXRay(output) as decoded:
        values = decoded.get_table("TheTable")["K"].tolist()
    values[0] = 11
    WorkbookModel.from_template(output).replace_column("TheTable", "K", values).save(output)
    with PBIXRay(output) as decoded:
        assert decoded.get_table("TheTable")["K"].tolist() == values


def test_low_level_repack_can_replace_source(tmp_path):
    output = tmp_path / "report.xlsx"
    output.write_bytes(FIXTURE.read_bytes())
    repack_workbook(output, output)
    with PBIXRay(output) as decoded:
        assert len(decoded.get_table("TheTable")) == 500


@pytest.mark.skipif(os.name != "nt", reason="Windows inherited file access regression")
@pytest.mark.parametrize("writer", ["save", "repack"])
def test_output_inherits_normal_folder_access(tmp_path, writer):
    import ctypes
    from ctypes import wintypes

    # Compare DACL bytes with a normal file created in the same folder. Moving a
    # file from TemporaryDirectory used to retain its private owner-only DACL.
    security = ctypes.WinDLL("advapi32", use_last_error=True)
    get_security = security.GetFileSecurityW
    get_security.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, ctypes.c_void_p,
                             wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
    get_security.restype = wintypes.BOOL
    get_dacl = security.GetSecurityDescriptorDacl
    get_dacl.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.BOOL),
                         ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(wintypes.BOOL)]
    get_dacl.restype = wintypes.BOOL

    def dacl_bytes(path):
        needed = wintypes.DWORD()
        get_security(str(path), 4, None, 0, ctypes.byref(needed))
        assert needed.value
        buffer = ctypes.create_string_buffer(needed.value)
        assert get_security(str(path), 4, buffer, len(buffer), ctypes.byref(needed))
        present, defaulted = wintypes.BOOL(), wintypes.BOOL()
        acl = ctypes.c_void_p()
        assert get_dacl(buffer, ctypes.byref(present), ctypes.byref(acl), ctypes.byref(defaulted))
        assert present.value and acl.value
        length = ctypes.c_ushort.from_address(acl.value + 2).value
        return ctypes.string_at(acl.value, length)

    reference = tmp_path / "normal-file.txt"
    reference.write_text("reference", encoding="utf-8")
    output = tmp_path / "model.xlsx"
    if writer == "save":
        WorkbookModel.from_template(FIXTURE).save(output)
    else:
        repack_workbook(FIXTURE, output)
    assert dacl_bytes(output) == dacl_bytes(reference)
