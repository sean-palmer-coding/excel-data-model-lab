from __future__ import annotations

from .collation import text_key

import uuid
from dataclasses import dataclass
from xml.etree import ElementTree as ET

from .abf import InvalidModelError
from .errors import RelationshipValidationError

ENGINE = "http://schemas.microsoft.com/analysisservices/2003/engine"
DDL300 = "http://schemas.microsoft.com/analysisservices/2011/engine/300"
DDL300_300 = "http://schemas.microsoft.com/analysisservices/2011/engine/300/300"
XSI = "http://www.w3.org/2001/XMLSchema-instance"


def validate_relationship_keys(relationship, foreign, primary):
    import pandas as pd
    from .abf import UnsupportedModelError
    ends = [(relationship.to_table, relationship.to_column, primary)]
    if relationship.cardinality == 'one-to-one':
        ends.append((relationship.from_table, relationship.from_column, foreign))
    for table, column, values in ends:
        seen = {}
        for row, value in enumerate(values, 1):
            if pd.isna(value) or value == '':
                raise UnsupportedModelError(
                    f'{table}.{column}: blank key at row {row}; blank lookup keys '
                    'are not yet supported by this writer')
            key = text_key(value) if isinstance(value, str) else value
            if key in seen:
                raise RelationshipValidationError(
                    f'{table}.{column} must contain unique values for '
                    f'{relationship.cardinality}; duplicate key {value!r} '
                    f'in rows {seen[key]} and {row}')
            seen[key] = row


@dataclass(frozen=True)
class Relationship:
    from_table: str
    from_column: str
    to_table: str
    to_column: str
    active: bool = True
    cardinality: str = "many-to-one"


def add_relationships_to_dimension(
    xml: bytes,
    *,
    from_dimension_id: str,
    relationships: list[tuple[Relationship, str]],
) -> bytes:
    """Append many-to-one relationship objects to a dimension definition."""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as exc:
        raise InvalidModelError(f"Invalid dimension XML: {exc}") from exc
    container = next(
        (node for node in root.iter() if node.tag.split("}")[-1] == "Relationships"),
        None,
    )
    if container is None:
        dimension = next(
            (node for node in root.iter() if node.tag.split("}")[-1] == "Dimension"),
            None,
        )
        if dimension is None:
            raise InvalidModelError("Dimension definition has no Dimension object")
        container = ET.SubElement(dimension, f"{{{DDL300_300}}}Relationships")

    for relationship, to_dimension_id in relationships:
        node = ET.SubElement(container, f"{{{DDL300_300}}}Relationship")
        ET.SubElement(node, f"{{{ENGINE}}}ID").text = str(uuid.uuid4())
        ET.SubElement(node, f"{{{ENGINE}}}Visible").text = (
            "true" if relationship.active else "false"
        )
        ET.SubElement(node, f"{{{DDL300}}}ProcessingState").text = "Processed"
        for end_name, multiplicity, dimension_id, attribute in (
            ("FromRelationshipEnd", "One" if relationship.cardinality == "one-to-one" else "Many", from_dimension_id, relationship.from_column),
            ("ToRelationshipEnd", "One", to_dimension_id, relationship.to_column),
        ):
            end = ET.SubElement(node, f"{{{DDL300_300}}}{end_name}")
            ET.SubElement(end, f"{{{ENGINE}}}Role")
            ET.SubElement(end, f"{{{DDL300_300}}}Multiplicity").text = multiplicity
            ET.SubElement(end, f"{{{ENGINE}}}DimensionID").text = dimension_id
            attributes = ET.SubElement(end, f"{{{ENGINE}}}Attributes")
            attr = ET.SubElement(attributes, f"{{{ENGINE}}}Attribute")
            ET.SubElement(attr, f"{{{ENGINE}}}AttributeID").text = attribute

    ET.register_namespace("", ENGINE)
    ET.register_namespace("ddl300", DDL300)
    ET.register_namespace("ddl300_300", DDL300_300)
    ET.register_namespace("xsi", XSI)
    rendered = ET.tostring(root, encoding="utf-8", short_empty_elements=True)
    if b"xmlns:xsd=" not in rendered:
        rendered = rendered.replace(
            b">", b' xmlns:xsd="http://www.w3.org/2001/XMLSchema">', 1
        )
    # Preserve namespace declarations used inside attribute values (valuens,
    # xsi:type). ElementTree cannot see those references and drops them.
    from xml.dom import minidom
    from .schema import _elements
    original = minidom.parseString(xml)
    generated = minidom.parseString(rendered)
    containers = _elements(original, 'Relationships')
    if containers:
        target = containers[0]
    else:
        target = original.createElementNS(DDL300_300, 'ddl300_300:Relationships')
        dimension = _elements(original, 'Dimension')[0]
        following = next((node for node in dimension.childNodes
                          if getattr(node, 'localName', None) == 'StructureVersion'), None)
        dimension.insertBefore(target, following)
    original.documentElement.setAttribute('xmlns:ddl300', DDL300)
    original.documentElement.setAttribute('xmlns:ddl300_300', DDL300_300)
    if relationships:
        for node in _elements(generated, 'Relationship')[-len(relationships):]:
            target.appendChild(original.importNode(node, True))
    return original.documentElement.toxml(encoding='utf-8')


def validate_relationship_collation(model, relationship):
    from .collation import model_string_order
    for table, column in ((relationship.from_table, relationship.from_column),
                          (relationship.to_table, relationship.to_column)):
        values = model.get_table(table)[column].tolist()
        strings = list(dict.fromkeys(v for v in values if isinstance(v, str)))
        if strings:
            model_string_order(model._image, model._dimension_ids[table], strings)
