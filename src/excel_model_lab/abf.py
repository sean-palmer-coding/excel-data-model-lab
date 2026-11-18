from __future__ import annotations

import hashlib
import json
import re
import struct
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from xpress8 import Xpress8

from .package import MODEL_PART

STREAM_SIGNATURE = b"\xff\xfe" + "STREAM_STORAGE_SIGNATURE_)!@#$%^&*(".encode("utf-16le")
HEADER_PAGE_SIZE = 4096
# Excel workbooks cap at 2 GB of model; leave headroom, but bound crafted input.
MAX_MODEL_PART_BYTES = 4 * 1024 ** 3


class UnsupportedModelError(ValueError):
    """Raised when an item.data dialect is outside the current lab scope."""


class InvalidModelError(ValueError):
    """Raised when the ABF layout is internally inconsistent."""


@dataclass(frozen=True)
class VirtualFile:
    storage_path: str
    offset: int
    size: int
    delete: bool
    created_timestamp: int
    access_timestamp: int
    modified_timestamp: int
    logical_path: str | None = None
    logical_size: int | None = None

    @property
    def end(self) -> int:
        return self.offset + self.size


@dataclass(frozen=True)
class VirtualFileAddition:
    """Logical file to add beside an existing file in its backup file group.

    This is a container primitive, not a model-schema operation. Callers must
    update the corresponding semantic metadata in the same assembly.
    """

    storage_path: str
    logical_path: str
    group_anchor: str
    data: bytes
    trailer: bytes | None = None
    target_group: str | None = None


@dataclass(frozen=True)
class VirtualFileGroupAddition:
    """Clone a dimension/measure-group/partition backup group under a new ID."""

    key: str
    group_anchor: str
    object_id: str
    name: str
    persist_location_path: str
    instance_id: str
    object_version: str | None = None
    object_class: str | None = None


def _decode_utf16_xml(data: bytes) -> ET.Element:
    try:
        return ET.fromstring(data.decode("utf-16").rstrip("\x00"))
    except (UnicodeDecodeError, ET.ParseError) as exc:
        raise InvalidModelError(f"Invalid UTF-16 model XML: {exc}") from exc


def _integer(root: ET.Element, name: str) -> int:
    value = root.findtext(name)
    if value is None:
        raise InvalidModelError(f"Missing {name} in ABF header")
    return int(value)


def _boolean(root: ET.Element, name: str) -> bool:
    return root.findtext(name, "false").lower() == "true"


def _classify(path: str | None, storage_path: str) -> str:
    name = (path or storage_path).replace("\\", "/").lower()
    if storage_path == "PARTITIONS":
        return "partitions"
    if storage_path == "LOG":
        return "backup_log"
    if name.endswith(".dictionary"):
        return "dictionary"
    if name.endswith(".idf"):
        return "column_segment"
    if name.endswith(".hidx"):
        return "hash_index"
    if name.endswith(".tbl.xml"):
        return "table_metadata"
    if name.endswith(".dim.xml"):
        return "dimension_metadata"
    if name.endswith(".prt.xml"):
        return "partition_metadata"
    if name.endswith(".det.xml"):
        return "database_metadata"
    if name.endswith(".scr.xml"):
        return "dax_script"
    if name.endswith(".xml"):
        return "xml_metadata"
    if name.endswith(".bin"):
        return "binary_metadata"
    return "unknown"


class AbfImage:
    """Parsed view of the uncompressed ABF stream stored in an XLSX model part.

    The first writer boundary deliberately supports only STREAM_STORAGE images.
    Xpress-compressed outer streams will be added after the clone writer is stable.
    """

    def __init__(self, data: bytes):
        self.data = bytes(data)
        if not self.data.startswith(STREAM_SIGNATURE):
            raise UnsupportedModelError("Only uncompressed STREAM_STORAGE item.data is supported")
        if len(self.data) < HEADER_PAGE_SIZE:
            raise InvalidModelError("ABF stream is shorter than its fixed header page")
        self.header_root = _decode_utf16_xml(self.data[len(STREAM_SIGNATURE):HEADER_PAGE_SIZE])
        self.virtual_directory_offset = _integer(self.header_root, "m_cbOffsetHeader")
        self.virtual_directory_size = _integer(self.header_root, "DataSize")
        self.data_offset = _integer(self.header_root, "m_cbOffsetData")
        self.file_count = _integer(self.header_root, "Files")
        self.apply_compression = _boolean(self.header_root, "ApplyCompression")
        self.error_code = _boolean(self.header_root, "ErrorCode")
        vd_end = self.virtual_directory_offset + self.virtual_directory_size
        if self.data_offset != HEADER_PAGE_SIZE:
            raise UnsupportedModelError(f"Unexpected data offset {self.data_offset}")
        if not (HEADER_PAGE_SIZE <= self.virtual_directory_offset <= vd_end <= len(self.data)):
            raise InvalidModelError("Virtual directory extent lies outside the ABF stream")
        self.virtual_directory_bytes = self.data[self.virtual_directory_offset:vd_end]
        self.virtual_directory_root = _decode_utf16_xml(self.virtual_directory_bytes)
        self.files = self._parse_files()
        self.backup_log_root: ET.Element | None = None
        self._attach_logical_paths()
        self.validate()

    @classmethod
    def from_workbook(cls, workbook: str | Path) -> "AbfImage":
        with zipfile.ZipFile(workbook) as package:
            try:
                info = package.getinfo(MODEL_PART)
            except KeyError as exc:
                raise UnsupportedModelError(f"Workbook has no {MODEL_PART}") from exc
            # zipfile never inflates past the declared size, so checking it
            # first bounds memory use for crafted (zip-bomb) workbooks.
            if info.file_size > MAX_MODEL_PART_BYTES:
                raise InvalidModelError(
                    f"{MODEL_PART} declares {info.file_size} bytes, above the "
                    f"{MAX_MODEL_PART_BYTES}-byte limit (excel_model_lab.abf.MAX_MODEL_PART_BYTES)")
            return cls(package.read(info))

    def _parse_files(self) -> list[VirtualFile]:
        files = []
        for node in self.virtual_directory_root.findall("BackupFile"):
            files.append(
                VirtualFile(
                    storage_path=node.findtext("Path", ""),
                    size=int(node.findtext("Size", "0")),
                    offset=int(node.findtext("m_cbOffsetHeader", "0")),
                    delete=node.findtext("Delete", "false").lower() == "true",
                    created_timestamp=int(node.findtext("CreatedTimestamp", "0")),
                    access_timestamp=int(node.findtext("Access", "0")),
                    modified_timestamp=int(node.findtext("LastWriteTime", "0")),
                )
            )
        return files

    def _attach_logical_paths(self) -> None:
        log = next((f for f in self.files if f.storage_path == "LOG"), None)
        if log is None:
            raise InvalidModelError("Virtual directory has no LOG file")
        raw = self.read_file(log)
        if self.error_code and len(raw) >= 4:
            raw = raw[:-4]
        root = _decode_utf16_xml(raw)
        self.backup_log_root = root
        mapping: dict[str, tuple[str, int]] = {}
        groups = root.findall("FileGroups/FileGroup")
        persist_root = ""
        if len(groups) > 1:
            persist_root = (groups[1].findtext("PersistLocationPath") or "") + "\\"
        for node in root.findall("FileGroups/FileGroup/FileList/BackupFile"):
            storage = node.findtext("StoragePath", "")
            path = node.findtext("Path", "")
            if persist_root and path.startswith(persist_root):
                path = path[len(persist_root):]
            mapping[storage] = (path, int(node.findtext("Size", "0")))
        self.files = [
            VirtualFile(
                **{**f.__dict__, "logical_path": mapping.get(f.storage_path, (None, None))[0],
                   "logical_size": mapping.get(f.storage_path, (None, None))[1]}
            )
            for f in self.files
        ]

    def validate(self) -> None:
        if len(self.files) != self.file_count:
            raise InvalidModelError(
                f"Header declares {self.file_count} files but directory contains {len(self.files)}"
            )
        ordered = sorted(self.files, key=lambda f: f.offset)
        cursor = self.data_offset
        for item in ordered:
            if item.offset < 0 or item.size < 0:
                raise InvalidModelError(f"Negative offset or size for virtual file {item.storage_path}")
            if item.offset < cursor:
                raise InvalidModelError(f"Overlapping virtual file {item.storage_path}")
            if item.end > self.virtual_directory_offset:
                raise InvalidModelError(f"Virtual file {item.storage_path} overlaps directory")
            cursor = item.end

    def read_file(self, item: VirtualFile | str) -> bytes:
        if isinstance(item, str):
            matches = [f for f in self.files if f.storage_path == item]
            if len(matches) != 1:
                raise KeyError(item)
            item = matches[0]
        return self.data[item.offset:item.end]

    def read_logical_file(self, item: VirtualFile | str) -> bytes:
        """Return a virtual file after removing the trailer and Xpress8 wrapper."""
        if isinstance(item, str):
            item = next((f for f in self.files if f.storage_path == item), None)
            if item is None:
                raise KeyError(item)
        raw = self.read_file(item)
        payload = raw[:-4] if self.error_code else raw
        if item.storage_path in {"PARTITIONS", "LOG"}:
            return payload
        if self.apply_compression:
            return Xpress8().decompress_chunked(payload)
        return payload

    def inventory(self) -> dict:
        rows = []
        for item in sorted(self.files, key=lambda f: f.offset):
            stored = self.read_file(item)
            rows.append(
                {
                    "storage_path": item.storage_path,
                    "logical_path": item.logical_path,
                    "kind": _classify(item.logical_path, item.storage_path),
                    "offset": item.offset,
                    "stored_size": item.size,
                    "logical_size": item.logical_size,
                    "sha256": hashlib.sha256(stored).hexdigest(),
                    "deleted": item.delete,
                }
            )
        return {
            "stream_size": len(self.data),
            "data_offset": self.data_offset,
            "virtual_directory_offset": self.virtual_directory_offset,
            "virtual_directory_size": self.virtual_directory_size,
            "file_count": self.file_count,
            "apply_compression": self.apply_compression,
            "error_code": self.error_code,
            "files": rows,
        }

    def clone(self) -> bytes:
        """Reassemble the supported ABF image from its parsed regions.

        This intentionally constructs a new buffer instead of returning the input.
        Gaps are zero-filled; nonzero gaps are rejected because their semantics are
        not yet understood.
        """
        output = bytearray(len(self.data))
        output[:HEADER_PAGE_SIZE] = self.data[:HEADER_PAGE_SIZE]
        cursor = self.data_offset
        for item in sorted(self.files, key=lambda f: f.offset):
            gap = self.data[cursor:item.offset]
            if any(gap):
                raise UnsupportedModelError(
                    f"Nonzero unclassified gap at {cursor}:{item.offset}"
                )
            output[item.offset:item.end] = self.read_file(item)
            cursor = item.end
        gap = self.data[cursor:self.virtual_directory_offset]
        if any(gap):
            raise UnsupportedModelError(
                f"Nonzero unclassified gap at {cursor}:{self.virtual_directory_offset}"
            )
        vd_end = self.virtual_directory_offset + self.virtual_directory_size
        output[self.virtual_directory_offset:vd_end] = self.virtual_directory_bytes
        trailer = self.data[vd_end:]
        if any(trailer):
            raise UnsupportedModelError(f"Nonzero unclassified trailer at {vd_end}")
        return bytes(output)

    @staticmethod
    def _xml_bytes(root: ET.Element, *, bom: bool = False) -> bytes:
        text = ET.tostring(root, encoding="unicode", short_empty_elements=True)
        return text.encode("utf-16" if bom else "utf-16le")

    @staticmethod
    def _align(value: int, boundary: int = HEADER_PAGE_SIZE) -> int:
        return (value + boundary - 1) // boundary * boundary

    def assemble(self, replacements: dict[str, bytes] | None = None, *,
                 additions: Iterable[VirtualFileAddition] = (),
                 groups: Iterable[VirtualFileGroupAddition] = (),
                 removals: Iterable[str] = ()) -> bytes:
        """Build a new ABF stream, optionally replacing logical virtual-file bytes.

        Replacements are uncompressed logical bytes keyed by storage path. The
        current implementation preserves each file's four-byte trailer while its
        meaning is still under investigation, then regenerates BackupLog sizes,
        virtual-directory offsets, and the fixed header page.
        """
        replacements = dict(replacements or {})
        additions = tuple(additions)
        removals = set(removals)
        known = {f.storage_path for f in self.files}
        unknown = (set(replacements) | removals) - known
        if unknown:
            raise KeyError(f"Unknown virtual files: {sorted(unknown)}")
        if {"PARTITIONS", "LOG"} & (set(replacements) | removals):
            raise UnsupportedModelError("Direct replacement/removal of PARTITIONS or LOG is not supported")
        if set(replacements) & removals:
            raise ValueError("Cannot replace and remove the same virtual file")
        if self.backup_log_root is None:
            raise InvalidModelError("BackupLog was not parsed")

        log_root = ET.fromstring(ET.tostring(self.backup_log_root, encoding="utf-8"))
        vd_root = ET.fromstring(ET.tostring(self.virtual_directory_root, encoding="utf-8"))
        owners = {node.findtext('StoragePath'): (file_list, node)
                  for file_list in log_root.findall('FileGroups/FileGroup/FileList')
                  for node in file_list.findall('BackupFile')}
        new_groups = {}
        group_container = log_root.find('FileGroups')
        for addition in groups:
            if not addition.key or addition.key in new_groups:
                raise ValueError('Duplicate or empty file-group key')
            source_groups = [g for g in group_container.findall('FileGroup')
                             if any(n.findtext('StoragePath') == addition.group_anchor
                                    for n in g.findall('FileList/BackupFile'))]
            if len(source_groups) != 1 or source_groups[0].findtext('Class') not in ('100006', '100016', '100021', '100054'):
                raise UnsupportedModelError('Only dimension, measure-group and partition groups can be cloned')
            source_group = source_groups[0]
            target_class = addition.object_class or source_group.findtext('Class')
            if addition.object_class is not None and target_class not in ('100003', '100053', '100054'):
                raise UnsupportedModelError('Only data-source, data-source-view and perspective file groups may override the source class')
            if not addition.object_id or any(g.findtext('Class') == target_class
                    and g.findtext('PersistLocationPath') == addition.persist_location_path
                    and (target_class not in ('100003', '100053', '100054') or g.findtext('ID') == addition.object_id)
                    for g in group_container.findall('FileGroup')):
                raise ValueError('Duplicate or empty file-group identity')
            group = ET.fromstring(ET.tostring(source_group, encoding='utf-8'))
            group.find('Class').text = target_class
            for field, value in (('ID', addition.object_id), ('Name', addition.name),
                                 ('PersistLocationPath', addition.persist_location_path),
                                 ('ObjectID', addition.instance_id)):
                group.find(field).text = value
            group.find('FileList').clear()
            if addition.object_version is not None:
                group.find('ObjectVersion').text = addition.object_version
            group_container.append(group)
            new_groups[addition.key] = group
        existing_paths = {node.findtext('Path', '').replace('/', '\\').casefold()
                          for _, node in owners.values()}
        items = [f for f in self.files if f.storage_path not in removals]
        added = {}
        for item in additions:
            if self.error_code and (item.trailer is None or len(item.trailer) != 4):
                raise UnsupportedModelError("Adding files requires an explicit four-byte trailer for this dialect")
            if not item.storage_path or item.storage_path in known:
                raise ValueError(f"Duplicate or empty storage path: {item.storage_path!r}")
            logical = item.logical_path.replace('/', '\\')
            if not logical or logical.casefold() in existing_paths:
                raise ValueError(f"Duplicate or empty logical path: {logical!r}")
            if item.group_anchor not in owners or item.group_anchor in removals:
                raise ValueError(f"Invalid file-group anchor: {item.group_anchor!r}")
            file_list, anchor = owners[item.group_anchor]
            anchor_path = anchor.findtext('Path', '')
            # Require a fully qualified backup path in the anchor's directory.
            # This avoids ambiguous relative paths when rebuilding the log.
            if item.target_group is not None:
                if item.target_group not in new_groups:
                    raise ValueError('Unknown target file group')
                group = new_groups[item.target_group]
                base = group.findtext('PersistLocationPath')
                parent = base.rsplit('\\', 1)[0]
                if not (logical.startswith(base + '\\') or (
                    logical.rsplit('\\', 1)[0] == parent
                    and logical.split('\\')[-1].startswith(group.findtext('ID') + '.')
                )):
                    raise ValueError('New virtual file is outside its target group')
                file_list = group.find('FileList')
            elif logical.rsplit('\\', 1)[0] != anchor_path.rsplit('\\', 1)[0]:
                raise ValueError("New virtual file must be beside its group anchor")
            node = ET.fromstring(ET.tostring(anchor, encoding='utf-8'))
            node.find('Path').text = logical
            node.find('StoragePath').text = item.storage_path
            node.find('Size').text = str(len(item.data))
            file_list.append(node)
            directory = ET.SubElement(vd_root, 'BackupFile')
            for key, value in [('Path', item.storage_path), ('Size', '0'),
                               ('m_cbOffsetHeader', '0'), ('Delete', 'false'),
                               ('CreatedTimestamp', '0'), ('Access', '0'), ('LastWriteTime', '0')]:
                ET.SubElement(directory, key).text = value
            items.append(VirtualFile(item.storage_path, len(self.data) + len(added),
                                     0, False, 0, 0, 0, logical, len(item.data)))
            replacements[item.storage_path] = bytes(item.data)
            added[item.storage_path] = item
            known.add(item.storage_path)
            existing_paths.add(logical.casefold())
        for name in removals:
            if name not in owners:
                raise ValueError(f"Cannot remove an unclassified virtual file: {name}")
            owner, node = owners[name]
            owner.remove(node)
        file_groups = log_root.find('FileGroups')
        for group in list(file_groups):
            if group.findtext('Class') in ('100003', '100006', '100016', '100021', '100053', '100054') and not list(group.find('FileList')):
                file_groups.remove(group)
        for node in list(vd_root):
            if node.findtext('Path') in removals:
                vd_root.remove(node)
        # Excel and independent readers locate the backup log as the last
        # directory entry. Newly added data files must precede it.
        log_directory = next(n for n in vd_root if n.findtext('Path') == 'LOG')
        vd_root.remove(log_directory)
        vd_root.append(log_directory)

        stored: dict[str, bytes] = {}
        logical_sizes: dict[str, int] = {}
        for item in items:
            if item.storage_path not in replacements:
                stored[item.storage_path] = self.read_file(item)
                if item.logical_size is not None:
                    logical_sizes[item.storage_path] = item.logical_size
                continue
            logical = bytes(replacements[item.storage_path])
            logical_sizes[item.storage_path] = len(logical)
            old = (self.read_file(item) if item.storage_path not in added
                   else added[item.storage_path].trailer or b'')
            trailer = old[-4:] if self.error_code else b""
            if self.apply_compression:
                chunks = []
                # Match the 4 KiB frames emitted by the Excel fixtures. The
                # Python decoder accepts larger frames that desktop Excel may
                # reject, so decoder round trips alone cannot justify 32 KiB.
                for offset in range(0, len(logical), 0x1000):
                    chunk = logical[offset:offset + 0x1000]
                    try:
                        compressed = Xpress8().compress(chunk)
                    except ValueError:
                        # Chunk framing explicitly permits stored/raw payloads
                        # when compression cannot reduce high-entropy data.
                        compressed = chunk
                    if len(compressed) >= len(chunk):
                        compressed = chunk
                    if len(compressed) > 0xFFFF:
                        raise UnsupportedModelError(
                            "Xpress8 produced a chunk larger than its 16-bit frame"
                        )
                    chunks.append(struct.pack("<HH", len(chunk), len(compressed)) + compressed)
                payload = b"".join(chunks)
            else:
                payload = logical
            stored[item.storage_path] = payload + trailer

        for node in log_root.findall("FileGroups/FileGroup/FileList/BackupFile"):
            storage_path = node.findtext("StoragePath", "")
            if storage_path in logical_sizes:
                size_node = node.find("Size")
                if size_node is not None:
                    size_node.text = str(logical_sizes[storage_path])
        log_item = next(f for f in self.files if f.storage_path == "LOG")
        old_log = self.read_file(log_item)
        log_trailer = old_log[-4:] if self.error_code else b""
        stored["LOG"] = self._xml_bytes(log_root, bom=True) + log_trailer

        vd_nodes = {node.findtext("Path", ""): node for node in vd_root.findall("BackupFile")}
        cursor = self.data_offset
        extents: dict[str, tuple[int, int]] = {}
        for item in sorted(items, key=lambda f: (f.storage_path == 'LOG', f.offset)):
            data = stored[item.storage_path]
            extents[item.storage_path] = (cursor, len(data))
            cursor += len(data)
        vd_offset = self._align(cursor)
        for storage_path, (offset, size) in extents.items():
            node = vd_nodes[storage_path]
            node.find("Size").text = str(size)
            node.find("m_cbOffsetHeader").text = str(offset)
        vd_bytes = self._xml_bytes(vd_root)

        header_root = ET.fromstring(ET.tostring(self.header_root, encoding="utf-8"))
        header_root.find("m_cbOffsetHeader").text = str(vd_offset)
        header_root.find("DataSize").text = str(len(vd_bytes))
        header_root.find("Files").text = str(len(items))
        header_xml = self._xml_bytes(header_root)
        header_capacity = HEADER_PAGE_SIZE - len(STREAM_SIGNATURE)
        if len(header_xml) > header_capacity:
            raise UnsupportedModelError("Regenerated ABF header exceeds its fixed page")

        total_size = self._align(vd_offset + len(vd_bytes))
        output = bytearray(total_size)
        output[:len(STREAM_SIGNATURE)] = STREAM_SIGNATURE
        output[len(STREAM_SIGNATURE):len(STREAM_SIGNATURE) + len(header_xml)] = header_xml
        for storage_path, (offset, size) in extents.items():
            output[offset:offset + size] = stored[storage_path]
        output[vd_offset:vd_offset + len(vd_bytes)] = vd_bytes
        rebuilt = bytes(output)
        AbfImage(rebuilt)
        return rebuilt

    def extract(self, destination: str | Path) -> Path:
        destination = Path(destination)
        planned, owners = [], {}
        for item in self.files:
            safe = re.sub(r"[^A-Za-z0-9_.-]", "_", item.storage_path)
            if safe.strip(".") == "" or safe.casefold() == "inventory.json":
                safe = "_" + safe
            # Distinct storage paths can sanitize (or case-fold on Windows)
            # to one name; refuse rather than silently overwrite an entry.
            if safe.casefold() in owners:
                raise InvalidModelError(
                    f"Virtual files {owners[safe.casefold()]!r} and {item.storage_path!r} "
                    "map to the same extracted filename")
            owners[safe.casefold()] = item.storage_path
            planned.append((item, safe))
        destination.mkdir(parents=True, exist_ok=True)
        for item, safe in planned:
            (destination / safe).write_bytes(self.read_file(item))
        (destination / "inventory.json").write_text(
            json.dumps(self.inventory(), indent=2) + "\n", encoding="utf-8"
        )
        return destination


def diff_inventories(left: AbfImage, right: AbfImage) -> dict:
    left_rows = {f["storage_path"]: f for f in left.inventory()["files"]}
    right_rows = {f["storage_path"]: f for f in right.inventory()["files"]}
    rows = []
    for name in sorted(set(left_rows) | set(right_rows)):
        a, b = left_rows.get(name), right_rows.get(name)
        if a == b:
            status = "unchanged"
        elif a is None:
            status = "added"
        elif b is None:
            status = "removed"
        else:
            status = "changed"
        rows.append({"storage_path": name, "status": status, "left": a, "right": b})
    return {
        "left_stream_size": len(left.data),
        "right_stream_size": len(right.data),
        "summary": {status: sum(r["status"] == status for r in rows)
                    for status in ("unchanged", "changed", "added", "removed")},
        "files": rows,
    }


def replace_model_part(source: str | Path, destination: str | Path, model: bytes) -> Path:
    from .package import repack_workbook

    return repack_workbook(source, destination, replacement_model=model)
