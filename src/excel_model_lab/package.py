from __future__ import annotations

import hashlib
import json
import os
import posixpath
import zipfile
from pathlib import Path
from tempfile import mkstemp
from typing import Any

MODEL_PART = "xl/model/item.data"


def _new_model_parts(original, model):
    """Wire a model into an existing workbook while retaining its other parts."""
    from xml.dom import minidom
    names = set(original.namelist())
    if MODEL_PART in names:
        return {MODEL_PART: model}
    rel_namespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
    type_namespace = 'http://schemas.openxmlformats.org/package/2006/content-types'
    sheet_namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    relpath = 'xl/_rels/workbook.xml.rels'
    rels = minidom.parseString(original.read(relpath)) if relpath in names else minidom.parseString(f'<Relationships xmlns="{rel_namespace}"/>')
    def relationship(kind, target):
        existing = list(rels.documentElement.childNodes)
        matches = [n for n in existing if n.nodeType == n.ELEMENT_NODE and n.getAttribute('Type') == kind]
        if matches:
            if len(matches) != 1 or matches[0].getAttribute('TargetMode') == 'External':
                raise ValueError('Workbook has an invalid internal model or connections relationship')
            resolved = posixpath.normpath(posixpath.join('xl', matches[0].getAttribute('Target'))).lstrip('/')
            if not resolved.startswith('xl/'):
                raise ValueError('Workbook relationship points outside the workbook parts')
            return resolved
        ids = {n.getAttribute('Id') for n in existing if n.nodeType == n.ELEMENT_NODE}
        number = 1
        while f'rId{number}' in ids:
            number += 1
        node = rels.createElementNS(rel_namespace, 'Relationship')
        for key, value in [('Id', f'rId{number}'), ('Type', kind), ('Target', target)]:
            node.setAttribute(key, value)
        rels.documentElement.appendChild(node)
        return 'xl/' + target
    # Excel's native embedded model uses powerPivotData. A plausible-looking
    # office/2007 "model" relationship is ignored: Model.Initialize then creates
    # an empty model and PivotTables report that their connection was deleted.
    if any(n.nodeType == n.ELEMENT_NODE and n.getAttribute('Type') ==
           'http://schemas.microsoft.com/office/2007/relationships/model'
           for n in rels.documentElement.childNodes):
        raise ValueError('Workbook has an unsupported legacy model relationship')
    model_path = relationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/powerPivotData', 'model/item.data')
    if model_path != MODEL_PART:
        raise ValueError('Workbook has a conflicting Data Model relationship')
    connections_path = relationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections', 'connections.xml')
    connections = minidom.parseString(original.read(connections_path)) if connections_path in names else minidom.parseString(f'<connections xmlns="{sheet_namespace}"/>')
    current = [n for n in connections.documentElement.childNodes if n.nodeType == n.ELEMENT_NODE]
    if any(n.getAttribute('name').casefold() == 'thisworkbookdatamodel' for n in current):
        raise ValueError('Workbook has a conflicting Data Model connection')
    with zipfile.ZipFile(Path(__file__).parent / 'data' / 'authoring-seed.xlsx') as seed:
        seed_connections = minidom.parseString(seed.read('xl/connections.xml'))
    connection = connections.importNode(seed_connections.documentElement.firstChild, True)
    connection.setAttribute('id', str(max([int(n.getAttribute('id')) for n in current] + [0]) + 1))
    connections.documentElement.appendChild(connection)
    types = minidom.parseString(original.read('[Content_Types].xml'))
    for part, mime in [('/' + MODEL_PART, 'application/vnd.openxmlformats-officedocument.model+data'),
                       ('/' + connections_path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml')]:
        if not any(n.nodeType == n.ELEMENT_NODE and n.getAttribute('PartName') == part for n in types.documentElement.childNodes):
            node = types.createElementNS(type_namespace, 'Override')
            node.setAttribute('PartName', part)
            node.setAttribute('ContentType', mime)
            types.documentElement.appendChild(node)
    return {MODEL_PART: model, relpath: rels.toxml(encoding='utf-8'),
            connections_path: connections.toxml(encoding='utf-8'), '[Content_Types].xml': types.toxml(encoding='utf-8')}


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def inspect_workbook(path: str | Path) -> dict[str, Any]:
    """Return a deterministic inventory of an XLSX package and its model part."""
    source = Path(path)
    with zipfile.ZipFile(source) as package:
        members = []
        for item in package.infolist():
            data = package.read(item.filename)
            members.append(
                {
                    "name": item.filename,
                    "size": len(data),
                    "compressed_size": item.compress_size,
                    "compression": item.compress_type,
                    "sha256": _sha256(data),
                }
            )
        model = next((member for member in members if member["name"] == MODEL_PART), None)
    return {
        "path": str(source),
        "size": source.stat().st_size,
        "has_data_model": model is not None,
        "model": model,
        "members": members,
    }


def semantic_snapshot(path: str | Path) -> dict[str, Any]:
    """Decode the model with PBIXRay and return JSON-safe semantic metadata."""
    from pbixray import PBIXRay

    with PBIXRay(str(path)) as model:
        tables = sorted(str(name) for name in model.tables)
        schema = model.schema.fillna("").to_dict(orient="records")
        relationships = model.relationships.fillna("").to_dict(orient="records")
        measures = model.dax_measures.fillna("").to_dict(orient="records")
        row_counts = {name: len(model.get_table(name)) for name in tables}
        return {
            "tables": tables,
            "schema": schema,
            "relationships": relationships,
            "measures": measures,
            "row_counts": row_counts,
            "model_size": model.size,
        }


def repack_workbook(
    source: str | Path,
    destination: str | Path,
    *,
    replacement_model: bytes | None = None,
) -> Path:
    """Repack an XLSX, preserving member metadata and optionally replacing item.data."""
    source = Path(source)
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    # A private temporary directory gives its children a private Windows ACL.
    # Stage directly in the output folder so the committed file inherits access.
    descriptor, staged_name = mkstemp(prefix=".excelmodel-pack-", suffix=".xlsx", dir=destination.parent)
    os.close(descriptor)
    staged = Path(staged_name)
    try:
        with zipfile.ZipFile(source, "r") as original, zipfile.ZipFile(staged, "w") as output:
            output.comment = original.comment
            replacements = _new_model_parts(original, replacement_model) if replacement_model is not None else {}
            if replacement_model is not None:
                from .structural_edit import renamed_table_package_parts
                replacements.update(renamed_table_package_parts(original, replacements, replacement_model))
                from .measure_layout import layout_parts
                replacements.update(layout_parts(original, replacements, replacement_model))
                from .native_sources import binding_package_parts
                replacements.update(binding_package_parts(original, replacements, replacement_model))
                if MODEL_PART in original.namelist() and original.read(MODEL_PART) != replacement_model:
                    from .cube_cache import invalidated_cube_cache_parts
                    replacements.update(invalidated_cube_cache_parts(original, replacements))
            for item in original.infolist():
                data = original.read(item.filename)
                if item.filename in replacements:
                    data = replacements.pop(item.filename)
                output.writestr(item, data)
            for name, data in replacements.items():
                output.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
        staged.replace(destination)
    finally:
        staged.unlink(missing_ok=True)
    return destination


def write_json(value: dict[str, Any], destination: str | Path) -> None:
    Path(destination).write_text(json.dumps(value, indent=2, default=str) + "\n", encoding="utf-8")
