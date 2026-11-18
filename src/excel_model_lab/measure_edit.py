"""Edit independent stored measures without evaluating expressions."""
from .formula_statements import iter_measure_definitions, normalize_expression, referenced_identifiers
from xml.etree import ElementTree as ET
from .measures import ENGINE
from .kpis import annotations, bracket, render

NS = {'e': ENGINE}

def edit_measure(xml, table, name, *, remove=False, expression=None,
                 format_string=None, description=None, display_folder=None):
    root = ET.fromstring(xml)
    commands = root.find('.//e:Commands', NS)
    matches = [c for c in commands if annotations(c).get('Table') == table
               and annotations(c).get('FullName') == name]
    if len(matches) != 1:
        raise KeyError('Measure not found: ' + table + bracket(name))
    command = matches[0]
    text = command.find('e:Text', NS)
    definitions = list(iter_measure_definitions(text.text or ''))
    if len(definitions) != 1 or 'CREATE KPI' in (text.text or '').upper():
        raise ValueError('Remove the KPI before editing its base measure')
    properties = root.find('.//e:CalculationProperties', NS)
    prop = next(p for p in properties if p.findtext('e:CalculationReference', namespaces=NS) == bracket(name))
    if remove:
        reference = name.casefold()
        for other in commands:
            if other is not command and reference in {
                identifier.casefold() for identifier in
                referenced_identifiers(other.findtext('e:Text', '', NS))
            }:
                raise ValueError('Another calculation depends on measure ' + name)
        commands.remove(command)
        properties.remove(prop)
    else:
        if expression is not None:
            if not isinstance(expression, str) or not expression.strip():
                raise ValueError('Measure expression must be nonempty')
            match = definitions[0]
            text.text = text.text[:match.start('expression')] + normalize_expression(expression) + '\n' + text.text[match.end('expression'):]
        for key, value in [('FormatString', format_string), ('Description', description), ('DisplayFolder', display_folder)]:
            if value is None:
                continue
            if not isinstance(value, str):
                raise ValueError('Measure metadata must be text')
            node = prop.find('e:' + key, NS)
            if node is None:
                node = ET.SubElement(prop, '{' + ENGINE + '}' + key)
            node.text = "'" + value.replace("'", "''") + "'" if key == 'FormatString' else value
            if key == 'FormatString':
                ann = prop.find('e:Annotations', NS)
                for item in list(ann) if ann is not None else []:
                    if item.findtext('e:Name', namespaces=NS) == 'Format':
                        ann.remove(item)
    return render(root)
