from __future__ import annotations

from xml.etree import ElementTree as ET

from .abf import InvalidModelError

IMBI = "http://schemas.microsoft.com/analysisservices/imbi"
XSI = "http://www.w3.org/2001/XMLSchema-instance"


def _member_object(parent: ET.Element, name: str) -> ET.Element | None:
    ns = {"i": IMBI}
    for member in parent.findall("i:Members/i:Member", ns):
        if member.findtext("i:Name", namespaces=ns) == name:
            return member.find("i:XMObject", ns)
    return None


def _set(properties: ET.Element | None, name: str, value: object) -> None:
    if properties is None:
        return
    node = properties.find(f"{{{IMBI}}}{name}")
    if node is not None:
        if isinstance(value, bool):
            node.text = "true" if value else "false"
        else:
            node.text = str(value)


def update_string_column_metadata(
    xml: bytes,
    *,
    column: str,
    data_ids: list[int],
    distinct_count: int,
) -> bytes:
    """Update the metadata duplicated around a string column dictionary."""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as exc:
        raise InvalidModelError(f"Invalid table metadata XML: {exc}") from exc
    target = next(
        (
            node for node in root.iter(f"{{{IMBI}}}XMObject")
            if node.attrib.get("class") == "XMRawColumn"
            and node.attrib.get("name") == column
        ),
        None,
    )
    if target is None:
        raise InvalidModelError(f"Table metadata has no XMRawColumn named {column}")

    has_nulls = 2 in data_ids
    ordinary = [value for value in data_ids if value != 2]
    min_id = min(ordinary, default=2)
    max_id = max(ordinary, default=2)
    stats = _member_object(target, "ColumnStats")
    props = stats.find(f"{{{IMBI}}}Properties") if stats is not None else None
    for name, value in {
        "DistinctStates": distinct_count + int(has_nulls),
        "MinDataID": min_id,
        "MaxDataID": max_id,
        "RowCount": len(data_ids),
        "HasNulls": has_nulls,
    }.items():
        _set(props, name, value)

    hierarchy = _member_object(target, "IntrinsicHierarchy")
    hierarchy_props = hierarchy.find(f"{{{IMBI}}}Properties") if hierarchy is not None else None
    _set(hierarchy_props, "DistinctDataIDs", distinct_count + int(has_nulls))

    for node in target.iter(f"{{{IMBI}}}XMObject"):
        cls = node.attrib.get("class", "")
        node_props = node.find(f"{{{IMBI}}}Properties")
        if cls == "XMColumnSegmentStats":
            _set(node_props, "MinDataID", min_id)
            _set(node_props, "MaxDataID", max_id)
            _set(node_props, "RowCount", len(data_ids))
            _set(node_props, "HasNulls", has_nulls)
        elif cls.startswith("XMHashDataDictionary<"):
            _set(node_props, "LastId", distinct_count + 2)
            _set(node_props, "Nullable", has_nulls)

    ET.register_namespace("", IMBI)
    ET.register_namespace("xsi", XSI)
    rendered = ET.tostring(root, encoding="utf-8", short_empty_elements=True)
    # xsi:type values use this lexical prefix, which ElementTree cannot infer.
    if b"xmlns:xsd=" not in rendered:
        rendered = rendered.replace(
            b">",
            b' xmlns:xsd="http://www.w3.org/2001/XMLSchema">',
            1,
        )
    return rendered


def update_column_segment_metadata(
    xml: bytes,
    *,
    column: str,
    data_ids: list[int],
    bit_width: int,
    min_data_id: int,
    dictionary_count: int | None = None,
) -> bytes:
    """Synchronize a column rebuilt as one pure bit-packed segment."""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as exc:
        raise InvalidModelError(f"Invalid table metadata XML: {exc}") from exc
    target = next(
        (
            node for node in root.iter(f"{{{IMBI}}}XMObject")
            if node.attrib.get("class") == "XMRawColumn"
            and node.attrib.get("name") == column
        ),
        None,
    )
    if target is None:
        raise InvalidModelError(f"Table metadata has no XMRawColumn named {column}")
    has_nulls = 2 in data_ids
    ordinary = [value for value in data_ids if value != 2]
    low = min(ordinary, default=2)
    high = max(ordinary, default=2)
    distinct = (dictionary_count if dictionary_count is not None else len(set(ordinary))) + int(has_nulls)

    stats = _member_object(target, "ColumnStats")
    props = stats.find(f"{{{IMBI}}}Properties") if stats is not None else None
    for name, value in {
        "DistinctStates": distinct,
        "MinDataID": low,
        "MaxDataID": high,
        "OriginalMinSegmentDataID": min_data_id,
        "RowCount": len(data_ids),
        "HasNulls": has_nulls,
        "RLERuns": 0,
        "OthersRLERuns": 0,
    }.items():
        _set(props, name, value)
    hierarchy = _member_object(target, "IntrinsicHierarchy")
    hierarchy_props = hierarchy.find(f"{{{IMBI}}}Properties") if hierarchy is not None else None
    _set(hierarchy_props, "DistinctDataIDs", distinct)

    collections = target.findall(f"{{{IMBI}}}Collections/{{{IMBI}}}Collection")
    segments_collection = next(
        (item for item in collections if item.findtext(f"{{{IMBI}}}Name") == "Segments"),
        None,
    )
    if segments_collection is None:
        raise InvalidModelError(f"Column {column} has no Segments collection")
    outer = segments_collection.find(f"{{{IMBI}}}XMObject")
    if outer is None:
        raise InvalidModelError(f"Column {column} has no segment object")
    _set(outer.find(f"{{{IMBI}}}Properties"), "Records", len(data_ids))
    sub = _member_object(outer, "SubSegment")
    if sub is None:
        raise InvalidModelError(f"Column {column} segment has no SubSegment")
    _set(sub.find(f"{{{IMBI}}}Properties"), "Records", len(data_ids))

    compression = _member_object(outer, "CompressionInfo")
    if compression is None or not compression.attrib.get("class", "").startswith("XMHybridRLECompressionInfo"):
        raise InvalidModelError("Rebuilt IDF requires a hybrid RLE segment wrapper")
    compression.attrib["class"] = f"XMHybridRLECompressionInfo<class XMRENoSplitCompressionInfo<{bit_width}>>"
    for codec in (_member_object(compression, "SubCompression"), _member_object(sub, "CompressionInfo")):
        if codec is None:
            raise InvalidModelError("Rebuilt segment has no bit-packed compression metadata")
        codec.attrib["class"] = f"XMRENoSplitCompressionInfo<{bit_width}>"
        _set(codec.find(f"{{{IMBI}}}Properties"), "Min", min_data_id)
    rle = _member_object(compression, "RLECompression")
    if rle is None:
        raise InvalidModelError("Rebuilt segment has no RLE compression metadata")
    rle_props = rle.find(f"{{{IMBI}}}Properties")
    _set(rle_props, "StorageAllocSize", 32 if data_ids else 0)
    _set(rle_props, "StorageUsedSize", 4 if data_ids else 0)
    _set(rle_props, "SegmentNeedsResizing", False)
    # The emitted stream consists of one bit-packed run plus allocation padding.
    _set(props, "OthersRLERuns", 1 if data_ids else 0)
    segment_stats = _member_object(outer, "ColumnSegmentStats")
    seg_props = segment_stats.find(f"{{{IMBI}}}Properties") if segment_stats is not None else None
    for name, value in {
        "MinDataID": low,
        "MaxDataID": high,
        "OriginalMinSegmentDataID": min_data_id,
        "RowCount": len(data_ids),
        "HasNulls": has_nulls,
        "RLERuns": 0,
        "OthersRLERuns": 1 if data_ids else 0,
    }.items():
        _set(seg_props, name, value)
    for node in target.iter(f"{{{IMBI}}}XMObject"):
        if node.attrib.get("class") == "XMRawColumnPartitionDataObject":
            _set(node.find(f"{{{IMBI}}}Properties"), "SegmentCount", 1)
        if dictionary_count is not None and node.attrib.get("class", "").startswith("XMHashDataDictionary"):
            node_props = node.find(f"{{{IMBI}}}Properties")
            _set(node_props, "LastId", dictionary_count + 2)
            _set(node_props, "Nullable", has_nulls)

    segment_map = _member_object(root, "SegmentMap")
    if segment_map is not None:
        maps = [node for node in segment_map.iter(f"{{{IMBI}}}XMObject")
                if node.attrib.get("class") == "XMSegment1Map"]
        if len(maps) == 1:
            _set(maps[0].find(f"{{{IMBI}}}Properties"), "Records", len(data_ids))
        else:
            raise InvalidModelError("Table resize requires a single partition segment map")

    ET.register_namespace("", IMBI)
    ET.register_namespace("xsi", XSI)
    rendered = ET.tostring(root, encoding="utf-8", short_empty_elements=True)
    if b"xmlns:xsd=" not in rendered:
        rendered = rendered.replace(
            b">", b' xmlns:xsd="http://www.w3.org/2001/XMLSchema">', 1
        )
    return rendered


def update_constant_column_row_count(xml: bytes, *, column: str, row_count: int) -> bytes:
    """Update record counts for an engine-managed constant/empty column."""
    root = ET.fromstring(xml)
    target = next(
        (node for node in root.iter(f"{{{IMBI}}}XMObject")
         if node.attrib.get("class") == "XMRawColumn" and node.attrib.get("name") == column),
        None,
    )
    if target is None:
        raise InvalidModelError(f"Table metadata has no XMRawColumn named {column}")
    stats = _member_object(target, "ColumnStats")
    stats_props = stats.find(f"{{{IMBI}}}Properties") if stats is not None else None
    _set(stats_props, "RowCount", row_count)
    _set(stats_props, "DistinctStates", row_count)
    _set(stats_props, "MaxDataID", row_count + 2)
    hierarchy = _member_object(target, "IntrinsicHierarchy")
    if hierarchy is not None:
        _set(hierarchy.find(f"{{{IMBI}}}Properties"), "DistinctDataIDs", row_count)
    for collection in target.findall(f"{{{IMBI}}}Collections/{{{IMBI}}}Collection"):
        if collection.findtext(f"{{{IMBI}}}Name") != "Segments":
            continue
        for segment in collection.findall(f"{{{IMBI}}}XMObject"):
            _set(segment.find(f"{{{IMBI}}}Properties"), "Records", row_count)
            sub = _member_object(segment, "SubSegment")
            if sub is not None:
                _set(sub.find(f"{{{IMBI}}}Properties"), "Records", row_count)
            segment_stats = _member_object(segment, "ColumnSegmentStats")
            _set(
                segment_stats.find(f"{{{IMBI}}}Properties") if segment_stats is not None else None,
                "RowCount",
                row_count,
            )
            if segment_stats is not None:
                _set(segment_stats.find(f"{{{IMBI}}}Properties"), "MaxDataID", row_count + 2)
    ET.register_namespace("", IMBI)
    ET.register_namespace("xsi", XSI)
    rendered = ET.tostring(root, encoding="utf-8", short_empty_elements=True)
    if b"xmlns:xsd=" not in rendered:
        rendered = rendered.replace(b">", b' xmlns:xsd="http://www.w3.org/2001/XMLSchema">', 1)
    return rendered
