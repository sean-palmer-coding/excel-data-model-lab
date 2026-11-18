from pathlib import Path

import pytest

from excel_model_lab.abf import AbfImage, VirtualFileAddition, UnsupportedModelError


FIXTURE = Path(__file__).parent / 'data' / 'null_data_id.xlsx'


def addition(image, *, name='NEWFILE', data=b'new column payload'):
    node = next(n for n in image.backup_log_root.findall(
        'FileGroups/FileGroup/FileList/BackupFile') if n.findtext('Path').endswith('.idf'))
    path = node.findtext('Path').rsplit('\\', 1)[0] + '\\new-column.idf'
    return VirtualFileAddition(name, path, node.findtext('StoragePath'), data, b'\x00' * 4)


def test_add_remove_preserves_existing_payloads_and_source():
    source = AbfImage.from_workbook(FIXTURE)
    original = source.data
    new = addition(source, data=bytes(range(256)) * 100)
    result = AbfImage(source.assemble(additions=[new]))
    assert result.read_logical_file(new.storage_path) == new.data
    assert result.file_count == source.file_count + 1
    for entry in source.files:
        if entry.storage_path != 'LOG':
            assert result.read_file(entry.storage_path) == source.read_file(entry)
    restored = AbfImage(result.assemble(removals=[new.storage_path]))
    assert restored.file_count == source.file_count
    assert {x.storage_path for x in restored.files} == {x.storage_path for x in source.files}
    assert source.data == original


def test_duplicate_additions_rejected():
    source = AbfImage.from_workbook(FIXTURE)
    new = addition(source)
    with pytest.raises(ValueError, match='Duplicate'):
        source.assemble(additions=[new, new])


def test_unknown_removal_and_reserved_files_rejected():
    source = AbfImage.from_workbook(FIXTURE)
    with pytest.raises(KeyError):
        source.assemble(removals=['missing'])
    with pytest.raises(UnsupportedModelError):
        source.assemble(removals=['LOG'])
    name = source.files[1].storage_path
    with pytest.raises(ValueError, match='replace and remove'):
        source.assemble({name: b'x'}, removals=[name])
