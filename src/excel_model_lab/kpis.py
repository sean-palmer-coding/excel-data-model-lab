"""Power Pivot KPI metadata and status definitions (no DAX evaluation)."""
from .formula_statements import iter_measure_definitions, normalize_expression
from copy import deepcopy
from pathlib import Path
import re
from decimal import Decimal, InvalidOperation
from .errors import ModelValidationError
from xml.etree import ElementTree as ET
from .measures import ENGINE, list_script_measures

NS = {'e': ENGINE}
def tag(name):
    return '{' + ENGINE + '}' + name

def annotations(element):
    return {a.findtext(tag('Name')): a.findtext(tag('Value'), '')
            for a in element.findall('e:Annotations/e:Annotation', NS)}

def bracket(name):
    return '[' + name.replace(']', ']]') + ']'

def quoted(name):
    return "'" + name.replace("'", "''") + "'"

class KpiValidationError(ModelValidationError, ValueError):
    """Invalid KPI configuration, with an object-scoped validation issue."""


_ICONS = ('Traffic Light - Single', 'Road Signs', 'Traffic Light', 'Shapes', 'Gauge - Ascending')
_MODES = {'ascending': 'Linear', 'descending': 'LinearReversed',
          'centered': 'Centered', 'centered_reversed': 'CenteredReversed'}


def _error(table, measure, message, code='invalid_kpi'):
    from .validation import ValidationIssue
    raise KpiValidationError(message, issues=(ValidationIssue(code, message,
        object_path=('tables', table, 'kpis', measure)),))


def _number(value, table, measure):
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        _error(table, measure, 'KPI values must be finite numbers')
    try:
        result = Decimal(str(value))
    except InvalidOperation:
        _error(table, measure, 'KPI values must be finite numbers')
    if not result.is_finite():
        _error(table, measure, 'KPI values must be finite numbers')
    return result


def _literal(value):
    return format(value, 'f')


def list_kpis(xml):
    """Read native KPI metadata, retaining legacy lower/upper result keys.

    Percentage thresholds are fractions (0.4 means 40 percent); absolute
    thresholds are raw measure values. Native hidden helper names are preserved.
    """
    root = ET.fromstring(xml)
    result = []
    for prop in root.findall('.//e:CalculationProperty', NS):
        data = annotations(prop)
        if data.get('Type') != 'Kpi':
            continue
        ref = prop.findtext(tag('CalculationReference'))
        name = ref[1:-1].replace(']]', ']')
        command = next((c for c in root.findall('.//e:Command', NS)
                       if annotations(c).get('FullName') == name), None)
        if command is None:
            _error('', name, 'KPI has no annotated measure command', 'unsupported_kpi_layout')
        scale = 100 if data.get('KpiThresholdType') == 'Percentage' else 1
        count = int(data.get('KpiThresholdCount', '2'))
        thresholds = tuple(float(data['KpiThreshold_' + str(i)]) / scale for i in range(count))
        graphic = re.search(r"STATUS_GRAPHIC\s*=\s*'((?:[^']|'')*)'", command.findtext(tag('Text'), ''), re.I)
        result.append(dict(table=annotations(command)['Table'], measure=name,
            target_measure=data.get('GoalMeasure'), target_table=data.get('GoalMeasureTableName'),
            target_value=Decimal(data['GoalValue']) if data.get('GoalType') == 'StaticValue' else None,
            threshold_type=data.get('KpiThresholdType', '').lower(), thresholds=thresholds,
            lower=thresholds[0] if thresholds else None, upper=thresholds[1] if len(thresholds) > 1 else None,
            mode=next((key for key, value in _MODES.items() if value == data.get('KpiStatusType')), data.get('KpiStatusType')),
            icon=graphic.group(1).replace("''", "'") if graphic else None))
    return result


def render(root):
    ET.register_namespace('', ENGINE)
    return ET.tostring(root, encoding='utf-8')

def add_kpi(xml, table, measure, *, target_measure=None, target_table=None,
            target_value=None, lower=0.4, upper=0.8, thresholds=None,
            mode='ascending', icon='Traffic Light - Single', threshold_type=None):
    """Create a KPI with a measure or fixed target and explicit status bands.

    Measure targets default to percentage fractions; fixed targets to absolute
    values. Explicit thresholds must be increasing. Three-state icons use two
    linear/four centered thresholds; the five-state gauge uses four/eight.
    Without explicit thresholds, the legacy lower/upper fractions are used for
    linear three-state KPIs (scaled by a fixed target when absolute).
    """
    if (target_measure is None) == (target_value is None):
        _error(table, measure, 'Specify exactly one of target_measure or target_value')
    if target_measure is not None and (not isinstance(target_measure, str) or not target_measure.strip()):
        _error(table, measure, 'target_measure must be a nonempty string')
    if target_value is not None and target_table is not None:
        _error(table, measure, 'target_table only applies to measure targets')
    if mode not in _MODES:
        _error(table, measure, 'Unsupported KPI mode: ' + str(mode), 'unsupported_kpi_mode')
    if icon not in _ICONS:
        _error(table, measure, 'Unsupported KPI icon: ' + str(icon), 'unsupported_kpi_icon')
    target_table = target_table or table
    kind = threshold_type or ('percentage' if target_measure is not None else 'absolute')
    if kind not in ('percentage', 'absolute'):
        _error(table, measure, 'threshold_type must be percentage or absolute')
    target = _number(target_value, table, measure) if target_value is not None else None
    centered = mode.startswith('centered')
    count = (4 if icon == 'Gauge - Ascending' else 2) * (2 if centered else 1)
    if thresholds is None:
        if count != 2:
            _error(table, measure, 'Centered and gauge KPIs require explicit thresholds')
        low, high = _number(lower, table, measure), _number(upper, table, measure)
        if not Decimal(0) <= low < high <= Decimal(1):
            _error(table, measure, 'KPI thresholds must satisfy 0 <= lower < upper <= 1')
        thresholds = (low, high)
        if kind == 'absolute':
            if target is None:
                _error(table, measure, 'Absolute measure-target KPIs require explicit thresholds')
            thresholds = tuple(t * target for t in thresholds)
    try:
        thresholds = tuple(_number(t, table, measure) for t in thresholds)
    except TypeError:
        _error(table, measure, 'thresholds must be a sequence of numbers')
    if len(thresholds) != count or any(a >= b for a, b in zip(thresholds, thresholds[1:])):
        _error(table, measure, f'This KPI needs {count} strictly increasing thresholds')
    existing = list_script_measures(xml)
    required = [(table, measure)] + ([(target_table, target_measure)] if target_measure is not None else [])
    for key in required:
        if key not in {(m.table, m.name) for m in existing}:
            raise KeyError('Measure not found: ' + str(key))
    if any(k['measure'].casefold() == measure.casefold() for k in list_kpis(xml)):
        _error(table, measure, 'KPI already exists: ' + measure)
    goal, status = '_' + measure + ' Goal', '_' + measure + ' Status'
    if any(m.name.casefold() in {goal.casefold(), status.casefold()} for m in existing):
        _error(table, measure, 'KPI helper measure name already exists')
    root = ET.fromstring(xml)
    commands = root.findall('.//e:Command', NS)
    matches = [c for c in commands if annotations(c).get('FullName') == measure and annotations(c).get('Table') == table]
    if len(matches) != 1:
        _error(table, measure, 'KPI requires a native annotated measure command', 'unsupported_kpi_layout')
    command = matches[0]
    container = root.find('.//e:Commands', NS)
    container.remove(command)
    container.append(command)
    text = command.find(tag('Text'))
    if len(list(iter_measure_definitions(text.text or ''))) != 1:
        _error(table, measure, 'Measure command already contains dependent definitions', 'unsupported_kpi_layout')
    value = quoted(table) + bracket(measure)
    goal_ref = quoted(table) + bracket(goal)
    compared = value + '/' + goal_ref if kind == 'percentage' else value
    max_status = 2 if icon == 'Gauge - Ascending' else 1
    statuses = list(range(-max_status, max_status + 1))
    if centered:
        statuses += statuses[-2::-1]
    if mode in ('descending', 'centered_reversed'):
        statuses = [-status for status in statuses]
    status_expression = str(statuses[-1])
    for threshold, status_code in reversed(list(zip(thresholds, statuses))):
        status_expression = f'IF({compared}<{_literal(threshold)},{status_code},{status_expression})'
    status_expression = f'IF(ISBLANK({compared}),BLANK(),{status_expression})'
    goal_expression = quoted(target_table) + bracket(target_measure) if target_measure is not None else _literal(target)
    text.text = (text.text or '').rstrip() + '\n' + (
        f'CREATE MEASURE {quoted(table)}{bracket(goal)} = {goal_expression};\n'
        f'CREATE MEASURE {quoted(table)}{bracket(status)} = {status_expression};\n'
        f'CREATE KPI CURRENTCUBE.{bracket(measure)} AS Measures.{bracket(measure)}, '
        f'ASSOCIATED_MEASURE_GROUP = {quoted(table)}, GOAL = Measures.{bracket(goal)}, '
        f"STATUS = Measures.{bracket(status)}, STATUS_GRAPHIC = '{icon}';\n")
    properties = root.find('.//e:CalculationProperties', NS)
    base = next(p for p in properties if p.findtext(tag('CalculationReference')) == bracket(measure))
    native = ET.parse(Path(__file__).parent / 'data' / 'kpi-properties.xml').getroot()
    names = {'[Total Amount]': bracket(measure), '[_Total Amount Goal]': bracket(goal),
             '[_Total Amount Status]': bracket(status), 'KPIs.[Total Amount]': 'KPIs.' + bracket(measure)}
    updates = {'MainObjectName': measure, 'GoalType': 'Measure' if target_measure is not None else 'StaticValue',
        'KpiStatusType': _MODES[mode], 'KpiThresholdType': kind.title(),
        'KpiThresholdOrdering': 'Ascending', 'KpiThresholdCount': str(count)}
    updates.update({'KpiThreshold_' + str(i): _literal(t * (100 if kind == 'percentage' else 1))
                    for i, t in enumerate(thresholds)})
    if target_measure is not None:
        updates.update(GoalMeasure=target_measure, GoalMeasureTableName=target_table)
    else:
        updates['GoalValue'] = _literal(target)
    for index, prototype in enumerate(native):
        prop = deepcopy(prototype)
        prop.find(tag('CalculationReference')).text = names[prop.findtext(tag('CalculationReference'))]
        for annotation in prop.findall('e:Annotations/e:Annotation', NS):
            key = annotation.findtext(tag('Name'))
            if key in updates:
                annotation.find(tag('Value')).text = updates[key]
        if index == 0:
            # Keep the base measure's formatting, description and display folder.
            dest = base.find(tag('Annotations'))
            old = annotations(base)
            for annotation in prop.findall('e:Annotations/e:Annotation', NS):
                key = annotation.findtext(tag('Name'))
                if key in ('Format', 'IsPrivate'):
                    continue
                for item in list(dest):
                    if item.findtext(tag('Name')) == key:
                        dest.remove(item)
                dest.append(annotation)
        else:
            properties.append(prop)
    # Add fields absent from the three-state template and remove stale target
    # metadata when changing between measure and fixed targets.
    dest = base.find(tag('Annotations'))
    for annotation in list(dest):
        key = annotation.findtext(tag('Name'))
        if key in ('GoalMeasure', 'GoalMeasureTableName', 'GoalValue') or key.startswith('KpiThreshold_'):
            dest.remove(annotation)
    for key, value in updates.items():
        if key == 'MainObjectName':
            continue
        old = next((a for a in dest if a.findtext(tag('Name')) == key), None)
        if old is None:
            old = ET.SubElement(dest, tag('Annotation'))
            ET.SubElement(old, tag('Name')).text = key
            ET.SubElement(old, tag('Value'))
        old.find(tag('Value')).text = value
    return render(root)

def remove_kpi(xml, table, measure):
    root = ET.fromstring(xml)
    item = next((k for k in list_kpis(xml) if (k['table'], k['measure']) == (table, measure)), None)
    if item is None:
        raise KeyError('KPI not found: ' + measure)
    command = next(c for c in root.findall('.//e:Command', NS)
                   if annotations(c).get('FullName') == measure and annotations(c).get('Table') == table)
    helpers = (bracket('_' + measure + ' Goal').casefold(), bracket('_' + measure + ' Status').casefold())
    for other in root.findall('.//e:Command', NS):
        if other is not command and any(h in (other.findtext(tag('Text')) or '').casefold() for h in helpers):
            _error(table, measure, 'Another calculation depends on KPI helper measures', 'dependent_kpi_helper')
    text = command.find(tag('Text'))
    matches = list(iter_measure_definitions(text.text or ''))
    if len(matches) != 3 or matches[0].group('name').replace(']]', ']') != measure:
        _error(table, measure, 'Unsupported KPI command layout', 'unsupported_kpi_layout')
    # Only our/native three-state layout is removed; retain the base measure.
    if {m.group('name').replace(']]', ']') for m in matches[1:]} != {'_' + measure + ' Goal', '_' + measure + ' Status'}:
        _error(table, measure, 'Unsupported KPI helper names', 'unsupported_kpi_layout')
    text.text = text.text[:matches[0].end()] + '\n'
    properties = root.find('.//e:CalculationProperties', NS)
    for prop in list(properties):
        data = annotations(prop)
        if data.get('MainObjectName') == measure and data.get('Type') in ('SupportMdx', 'SupportKpi'):
            properties.remove(prop)
        elif prop.findtext(tag('CalculationReference')) == bracket(measure):
            for annotation in list(prop.find(tag('Annotations'))):
                key = annotation.findtext(tag('Name'))
                if key == 'Type':
                    annotation.find(tag('Value')).text = 'User'
                elif key.startswith('Kpi') or key in ('GoalType', 'GoalValue', 'GoalMeasure', 'GoalMeasureTableName', 'ValueDescription', 'GoalDescription', 'StatusDescription', 'TrendDescription'):
                    prop.find(tag('Annotations')).remove(annotation)
    return render(root)
