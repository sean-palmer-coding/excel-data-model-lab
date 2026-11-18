"""Validated OLAP member selections and label sorting for generated reports.

Member labels come from the current model, not calculated result caches.
"""
from dataclasses import dataclass
from datetime import date, datetime, time
from numbers import Integral
import math
from xml.etree import ElementTree as ET
import pandas as pd

from .errors import ModelValidationError

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


@dataclass(frozen=True)
class PivotMemberFilter:
    table: str
    column: str
    members: tuple

    def __post_init__(self):
        if isinstance(self.members, (str, bytes)):
            raise ModelValidationError('members must be a sequence of scalar values')
        try:
            object.__setattr__(self, 'members', tuple(self.members))
        except TypeError as exc:
            raise ModelValidationError('members must be a sequence of scalar values') from exc


@dataclass(frozen=True)
class PivotSort:
    table: str
    column: str
    order: str = 'ascending'


def _mdx(value):
    return '[' + value.replace(']', ']]') + ']'


def _scalar(value):
    if value is None or value is pd.NA or value is pd.NaT:
        return ('blank', None)
    if pd.api.types.is_bool(value):
        return ('boolean', bool(value))
    if isinstance(value, (datetime, date)):
        stamp = pd.Timestamp(datetime.combine(value, time()) if not isinstance(value, datetime) else value)
        if stamp.tzinfo is not None or stamp.microsecond or stamp.nanosecond:
            raise ModelValidationError('Date members require timezone-naive whole-second precision')
        return ('datetime', stamp.isoformat())
    if isinstance(value, str):
        if any(not (ord(c) in (9, 10, 13) or 32 <= ord(c) <= 0xD7FF or
                    0xE000 <= ord(c) <= 0xFFFD or 0x10000 <= ord(c) <= 0x10FFFF) for c in value):
            raise ModelValidationError('Filter member contains invalid XML text')
        return ('text', value)
    if isinstance(value, Integral) and not isinstance(value, bool):
        return ('integer', int(value))
    if isinstance(value, float) and pd.isna(value):
        return ('blank', None)
    raise ModelValidationError('Filter members support text, integer, boolean, date/datetime, and None values')


def member_unique_name(table, column, value):
    """Return an escaped attribute member key, never executable MDX text."""
    kind, key = _scalar(value)
    return _mdx(table) + '.' + _mdx(column) + '.&' + ('' if kind == 'blank' else _mdx(str(key)))


def _caption(value):
    kind, key = _scalar(value)
    return '(blank)' if kind == 'blank' else str(key)


def source_member_values(model, table, column, source=None):
    """Restore typed source keys when nullable reader columns use floats.

    This does not coerce requested filter values or enable Double members.
    Large float integers cannot establish the original exact model key.
    """
    calculated = getattr(model, 'list_calculated_columns', lambda: ())()
    table_id = getattr(model, '_dimension_ids', {}).get(table, table)
    for definition in calculated:
        if definition.table_id == table_id and definition.name == column:
            # DAX can depend on other tables, measures and relationships. Until
            # Excel recalculates, staged changes cannot certify cached members.
            pending = any(getattr(model, name, None) for name in
                          ('_pending', '_pending_measures', '_pending_relationships', '_resized_tables'))
            if definition.processing_state != 'Processed' or pending:
                raise ModelValidationError(
                    f'{table}.{column}: calculated member values require Excel calculation and save/reopen before report filtering or slicer authoring')
    metadata = next((c for c in getattr(model, 'columns', ())
                     if c.table == table and c.name == column), None)
    data_type = metadata.data_type.lower() if metadata is not None else ''
    values = model.get_table(table)[column] if source is None else source
    for value in values:
        if isinstance(value, float) and not pd.isna(value):
            if data_type in ('int64', 'bigint', 'integer'):
                if not math.isfinite(value) or not value.is_integer() or abs(value) >= 2**53:
                    raise ModelValidationError('Integer member readback is not an exact safe integer')
                value = int(value)
            elif data_type in ('bool', 'boolean'):
                if value not in (0.0, 1.0):
                    raise ModelValidationError('Boolean member readback must be zero or one')
                value = bool(value)
        elif data_type in ('bool', 'boolean') and isinstance(value, Integral):
            if value not in (0, 1):
                raise ModelValidationError('Boolean member readback must be zero or one')
            value = bool(value)
        yield value


def validate_pivot_options(spec, model):
    """Check selections against current staged model data before any writes."""
    axes = tuple(spec.rows) + tuple(spec.columns) + tuple(spec.filters)
    filters = getattr(spec, 'member_filters', ())
    sorts = getattr(spec, 'sorts', ())
    frames = {}
    seen = set()
    for option in filters:
        if not isinstance(option, PivotMemberFilter):
            raise ModelValidationError('member_filters must contain PivotMemberFilter objects')
        field = (option.table, option.column)
        if field not in axes:
            raise ModelValidationError('A member filter requires a field on a report axis')
        if field in seen:
            raise ModelValidationError('Duplicate member filter field')
        seen.add(field)
        if not option.members:
            raise ModelValidationError('A member filter must select at least one member')
        keys = [_scalar(value) for value in option.members]
        if len(set(keys)) != len(keys):
            raise ModelValidationError('Duplicate selected member')
        if option.table not in frames:
            frames[option.table] = model.get_table(option.table)
        try:
            source = frames[option.table][option.column]
        except KeyError as exc:
            raise ModelValidationError('Unknown member filter field') from exc
        actual = set()
        for value in source_member_values(model, option.table, option.column, source):
            try:
                actual.add(_scalar(value))
            except ModelValidationError:
                # Unsupported or missing source values are irrelevant unless
                # selected; selected unsupported values already failed above.
                continue
        if any(key not in actual for key in keys):
            raise ModelValidationError('Selected member is absent from the current model column')
    seen = set()
    for option in sorts:
        if not isinstance(option, PivotSort):
            raise ModelValidationError('sorts must contain PivotSort objects')
        field = (option.table, option.column)
        if field not in tuple(spec.rows) + tuple(spec.columns):
            raise ModelValidationError('Label sorting requires a row or column field')
        if field in seen:
            raise ModelValidationError('Duplicate sort field')
        seen.add(field)
        if option.order not in ('ascending', 'descending'):
            raise ModelValidationError('Sort order must be ascending or descending')


def _child(parent, name, **attributes):
    return ET.SubElement(parent, '{' + S + '}' + name,
                         {key: str(value) for key, value in attributes.items()})


def apply_pivot_options(pivot, cache, spec, model):
    """Apply options to freshly built definition roots, before serialization."""
    validate_pivot_options(spec, model)
    axes = list(spec.rows) + list(spec.columns) + list(spec.filters)
    pivot_fields = pivot.find('{' + S + '}pivotFields')
    cache_fields = cache.find('{' + S + '}cacheFields')
    for option in getattr(spec, 'sorts', ()):
        pivot_fields[axes.index((option.table, option.column))].set('sortType', option.order)
    for option in getattr(spec, 'member_filters', ()):
        index = axes.index((option.table, option.column))
        field = pivot_fields[index]
        if (option.table, option.column) in spec.filters and len(option.members) > 1:
            # OLAP page filters use the hierarchy switch, not the similarly
            # named pivotField switch. Otherwise pageField.name=[All] wins.
            hierarchy = int(cache_fields[index].get('hierarchy'))
            pivot.find('{' + S + '}pivotHierarchies')[hierarchy].set('multipleItemSelectionAllowed', '1')
        if (option.table, option.column) in spec.filters and len(option.members) == 1:
            page = next(p for p in pivot.find('{' + S + '}pageFields') if p.get('fld') == str(index))
            value = option.members[0]
            page.set('name', member_unique_name(option.table, option.column, value))
            page.set('cap', _caption(value))
            continue
        # Native Excel stores selected OLAP captions in sharedItems and marks
        # their pivot items selected. These are actual source member labels,
        # not generated aggregate records or worksheet calculation results.
        cf = cache_fields[index]
        for old in list(cf):
            if old.tag == '{' + S + '}sharedItems':
                cf.remove(old)
        shared = _child(cf, 'sharedItems', count=len(option.members))
        kinds = {_scalar(v)[0] for v in option.members}
        # Match Excel's type flags: boolean-only caches retain the schema's
        # defaults (including containsString), rather than synthesizing a flag
        # combination absent from the native boolean reference.
        if 'blank' in kinds:
            shared.set('containsBlank', '1')
        if 'integer' in kinds:
            shared.set('containsString', '0')
            shared.set('containsNumber', '1')
            shared.set('containsInteger', '1')
        if kinds == {'datetime'}:
            shared.set('containsString', '0')
            shared.set('containsDate', '1')
            shared.set('containsSemiMixedTypes', '0')
            shared.set('containsNonDate', '0')
        for value in option.members:
            kind, key = _scalar(value)
            if kind == 'blank':
                _child(shared, 'm')
            else:
                _child(shared, {'text':'s', 'integer':'n', 'boolean':'b', 'datetime':'d'}[kind],
                       v=str(int(key)) if kind == 'boolean' else str(key))
        for old in list(field):
            if old.tag == '{' + S + '}items':
                field.remove(old)
        items = ET.Element('{' + S + '}items', {'count': str(len(option.members) + 1)})
        field.insert(0, items)
        for i in range(len(option.members)):
            _child(items, 'item', s=1, x=i)
        _child(items, 'item', t='default')
