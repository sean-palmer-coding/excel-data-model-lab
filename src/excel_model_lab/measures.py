from __future__ import annotations

from .formula_statements import iter_measure_definitions, normalize_expression


from dataclasses import dataclass
from xml.etree import ElementTree as ET

from .abf import InvalidModelError

ENGINE = "http://schemas.microsoft.com/analysisservices/2003/engine"
XSI = "http://www.w3.org/2001/XMLSchema-instance"
@dataclass(frozen=True)
class Measure:
    table: str
    name: str
    expression: str
    format_string: str = ""
    display_folder: str = ""
    description: str = ""


def list_script_measures(xml: bytes) -> list[Measure]:
    root = ET.fromstring(xml)
    result = []
    for text in root.iter(f"{{{ENGINE}}}Text"):
        for match in iter_measure_definitions(text.text or ""):
            result.append(
                Measure(
                    table=match.group("table").replace("''", "'"),
                    name=match.group("name").replace("]]", "]"),
                    expression=match.group("expression").strip(),
                )
            )
    return result


def add_measures_to_script(xml: bytes, measures: list[Measure]) -> bytes:
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as exc:
        raise InvalidModelError(f"Invalid MDX script XML: {exc}") from exc
    commands = root.find(f".//{{{ENGINE}}}Commands")
    properties = root.find(f".//{{{ENGINE}}}CalculationProperties")
    if commands is None or properties is None:
        raise InvalidModelError("MDX script lacks Commands or CalculationProperties")
    existing = {item.name.casefold() for item in list_script_measures(xml)}
    for measure in measures:
        if measure.name.casefold() in existing:
            raise ValueError(f"Measure already exists: {measure.table}[{measure.name}]")
        existing.add(measure.name.casefold())
        table = measure.table.replace("'", "''")
        name = measure.name.replace("]", "]]" )
        command = ET.SubElement(commands, f"{{{ENGINE}}}Command")
        text = ET.SubElement(command, f"{{{ENGINE}}}Text")
        text.text = (
            "----------------------------------------------------------\n"
            "-- PowerPivot measures command (do not modify manually) --\n"
            "----------------------------------------------------------\n\n\n"
            f"CREATE MEASURE '{table}'[{name}]={normalize_expression(measure.expression)}\n;\n"
        )
        annotations = ET.SubElement(command, f"{{{ENGINE}}}Annotations")
        for key, value in (("FullName", measure.name), ("Table", measure.table)):
            annotation = ET.SubElement(annotations, f"{{{ENGINE}}}Annotation")
            ET.SubElement(annotation, f"{{{ENGINE}}}Name").text = key
            ET.SubElement(annotation, f"{{{ENGINE}}}Value").text = value

        prop = ET.SubElement(properties, f"{{{ENGINE}}}CalculationProperty")
        fields = {
            "CalculationReference": f"[{name}]",
            "CalculationType": "Member",
            "Description": measure.description,
            "FormatString": "'" + measure.format_string.replace("'", "''") + "'",
            "ForeColor": "",
            "BackColor": "",
            "FontName": "",
            "FontSize": "",
            "FontFlags": "",
            "NonEmptyBehavior": "",
            "AssociatedMeasureGroupID": "",
            "DisplayFolder": measure.display_folder,
            "Language": "",
        }
        for key, value in fields.items():
            ET.SubElement(prop, f"{{{ENGINE}}}{key}").text = value
        annotations = ET.SubElement(prop, f"{{{ENGINE}}}Annotations")
        for key, value in (("Type", "User"), ("IsPrivate", "False")):
            annotation = ET.SubElement(annotations, f"{{{ENGINE}}}Annotation")
            ET.SubElement(annotation, f"{{{ENGINE}}}Name").text = key
            ET.SubElement(annotation, f"{{{ENGINE}}}Value").text = value
        ET.SubElement(prop, f"{{{ENGINE}}}Translations")

    ET.register_namespace("", ENGINE)
    ET.register_namespace("xsi", XSI)
    rendered = ET.tostring(root, encoding="utf-8", short_empty_elements=True)
    if b"xmlns:xsd=" not in rendered:
        rendered = rendered.replace(
            b">", b' xmlns:xsd="http://www.w3.org/2001/XMLSchema">', 1
        )
    return rendered
