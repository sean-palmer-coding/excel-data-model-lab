"""Numeric dictionary authoring for values outside the arithmetic ID range."""
from decimal import Decimal
import math
import operator
import struct
from xml.dom import minidom

import pandas as pd

from .abf import AbfImage, UnsupportedModelError
from .schema import _elements, _child, _text


def storage_number(value, semantic):
    """Convert a logical number without losing int64 or four-place currency."""
    if semantic == 'Currency':
        exact = Decimal(str(value)) * 10000
        if not exact.is_finite() or exact != exact.to_integral_value():
            raise ValueError('Currency requires finite values with at most four decimal places')
        number = int(exact)
    elif semantic == 'Double':
        number = float(value)
        if not math.isfinite(number):
            raise ValueError('Numeric values must be finite')
        if number != value:
            raise ValueError('Numeric dictionary conversion would lose precision')
        return number
    else:
        try:
            number = operator.index(value)
        except TypeError:
            try:
                number = int(value)
            except (TypeError, ValueError, OverflowError) as exc:
                raise ValueError('Value is not exactly representable as an integer') from exc
            if number != value:
                raise ValueError('Value is not exactly representable as an integer')
    if not -(1 << 63) <= number < (1 << 63):
        raise ValueError('Numeric value exceeds its signed 64-bit storage range')
    if semantic == 'Currency' and number == -(1 << 63):
        raise ValueError('Currency -922337203685477.5808 cannot be authored: Excel Power Pivot '
                         'loads it but returns blank MIN/MAX aggregates; minimum supported '
                         'currency is -922337203685477.5807')
    if semantic == 'BigInt' and number == -(1 << 63):
        raise ValueError('Integer -9223372036854775808 cannot be authored: Excel Power Pivot '
                         'loads it but returns blank numeric aggregates; minimum supported '
                         'integer is -9223372036854775807')
    if semantic == 'BigInt' and number == (1 << 63) - 1:
        raise ValueError('Integer 9223372036854775807 cannot be authored: Excel Power Pivot '
                         'loads it but returns blank MAX/SUM aggregates; maximum supported '
                         'integer is 9223372036854775806')
    if semantic == 'Currency' and number == (1 << 63) - 1:
        raise ValueError('Currency 922337203685477.5807 cannot be authored: Excel Power Pivot '
                         'loads it but returns blank MAX/SUM aggregates; maximum supported '
                         'currency is 922337203685477.5806')
    return number


def needs_numeric_dictionary(values, data_type):
    """Choose dictionary storage when value encoding cannot preserve the range."""
    semantic = {'integer': 'BigInt', 'currency': 'Currency', 'decimal': 'Double'}[data_type]
    actual = [storage_number(value, semantic) for value in values if not pd.isna(value)]
    if not actual:
        return False
    if data_type == 'decimal':
        decimals = [Decimal(str(value)) for value in actual]
        places = max(max(0, -value.normalize().as_tuple().exponent) for value in decimals)
        if places > 15:
            return True
        scaled = [value * (10 ** places) for value in decimals]
    else:
        scaled = actual
    # Large bases also pass through floating metadata in upstream readers.
    return (max(scaled) - min(scaled) > 0x7ffffffc or
            min(scaled) < -(1 << 53) or max(scaled) > (1 << 53))


def convert_string_seed_to_numeric(image, table_id, column, data_type):
    """Convert a new three-member string seed to an exact numeric dictionary.

    The seed's two-way intrinsic hierarchy is already dictionary-compatible;
    its stable IDs 3,4,5 become numeric 1,2,3. No existing user values may be
    passed to this creation-only helper.
    """
    types = {'integer': ('BigInt', 20, 0, 'XM_Long'),
             'currency': ('Currency', 6, 0, 'XM_Long'),
             'decimal': ('Double', 5, 1, 'XM_Real')}
    semantic, dbtype, xmtype, physical = types[data_type]
    replacements = {}
    found = set()
    dictionary_name = None
    table_directory = None
    for file in image.files:
        if not file.logical_path:
            continue
        filename = file.logical_path.rsplit('\\', 1)[-1]
        if not filename.endswith(('.dim.xml', '.det.xml', '.tbl.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        changed = False
        if filename.endswith('.det.xml'):
            for dimension in _elements(doc, 'Dimension'):
                ids = [n for n in dimension.childNodes if getattr(n, 'localName', None) == 'CubeDimensionID']
                if not ids or _text(ids[0]) != table_id:
                    continue
                for attribute in _elements(dimension, 'Attribute'):
                    if _text(_child(attribute, 'AttributeID')) == column:
                        for key in _elements(attribute, 'KeyColumn'):
                            _child(key, 'DataType').firstChild.data = semantic
                            changed = True
        elif filename.startswith(table_id + '.') and filename.endswith('.dim.xml'):
            for attribute in _elements(doc, 'Attribute'):
                ids = [n for n in attribute.childNodes if getattr(n, 'localName', None) == 'ID']
                if ids and _text(ids[0]) == column:
                    for key in _elements(attribute, 'KeyColumn'):
                        _child(key, 'DataType').firstChild.data = semantic
                    changed = True
                    found.add('dimension')
        elif filename.startswith(table_id + '.') and filename.endswith('.tbl.xml'):
            for raw in _elements(doc, 'XMObject'):
                if raw.getAttribute('class') != 'XMRawColumn' or raw.getAttribute('name') != column:
                    continue
                dictionaries = [n for n in _elements(raw, 'XMObject')
                                if n.getAttribute('class').startswith('XMHashDataDictionary<')]
                if len(dictionaries) != 1 or 'XM_String' not in dictionaries[0].getAttribute('class'):
                    raise UnsupportedModelError('Numeric creation requires a fresh string dictionary seed')
                dictionary = dictionaries[0]
                dictionary.setAttribute('class', f'XMHashDataDictionary<{physical}>')
                if physical == 'XM_Long':
                    # Payload vectors are signed int64. Native nullable lookups
                    # require this flag even when nonnullable models load without it.
                    flag = doc.createElement('OperatingOn32')
                    flag.setAttribute('xsi:type', 'xsd:boolean')
                    flag.appendChild(doc.createTextNode('false'))
                    _child(dictionary, 'Properties').appendChild(flag)
                for flag in _elements(dictionary, 'DictionaryFlags'):
                    flag.parentNode.removeChild(flag)
                dictionary_name = dictionary.getAttribute('name')
                table_directory = file.logical_path.rsplit('\\', 1)[0]
                # Native numeric dictionaries use Usage=3. Usage=0 happens to
                # load some cardinalities, but Excel rejects two distinct large
                # integers even with byte-identical dictionary/lookup streams.
                for tag, value in [('DBType', dbtype), ('XMType', xmtype), ('Usage', 3), ('EncodingHint', 1)]:
                    for node in _elements(raw, tag):
                        node.firstChild.data = str(value)
                changed = True
                found.add('storage')
        if changed:
            replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    if found != {'dimension', 'storage'}:
        raise UnsupportedModelError('Cannot resolve numeric creation column')
    dictionary_file = next(f for f in image.files if f.logical_path == table_directory + '\\' + dictionary_name)
    # Vector entries are raw signed int64 or IEEE binary64, as specified by
    # MS-XLDM numeric vectors.
    # Numeric dictionaries use the native numeric hash policy. String hash
    # policy fields are not interchangeable even though readers ignore them.
    header = struct.pack('<6i', -1, 16, 64, 3, -1, -1)
    initial = [10000, 20000, 30000] if data_type == 'currency' else [1, 2, 3]
    payload = (struct.pack('<i', xmtype) + header + struct.pack('<QI', 3, 8) +
               struct.pack('<3' + ('d' if xmtype else 'q'), *initial))
    replacements[dictionary_file.storage_path] = payload
    return AbfImage(image.assemble(replacements))


def converted_values(values, data_type, decimal_places):
    """Explicit, lossless stored-column conversions; blanks remain blanks."""
    import numbers
    result = []
    for value in values:
        if pd.isna(value):
            result.append(pd.NaT if data_type == 'datetime' else None)
            continue
        if data_type == 'text':
            result.append(str(value))
        elif data_type == 'integer':
            number = Decimal(value) if isinstance(value, str) else value
            result.append(storage_number(number, 'BigInt'))
        elif data_type == 'currency':
            number = Decimal(str(value))
            storage_number(number, 'Currency')
            result.append(number)
        elif data_type == 'decimal':
            number = float(value)
            if not math.isfinite(number):
                raise ValueError('Numeric values must be finite')
            if isinstance(value, (str, Decimal)):
                if Decimal(str(number)) != Decimal(str(value)):
                    raise ValueError('Value is not exactly representable as a decimal')
            elif number != value:
                raise ValueError('Value is not exactly representable as a decimal')
            scaled = Decimal(str(number)) * (10 ** decimal_places)
            if scaled != scaled.to_integral_value():
                raise ValueError('Value is not exactly representable at the declared decimal precision')
            result.append(number)
        elif data_type == 'boolean':
            if isinstance(value, str):
                mapping = {'true': True, 'false': False, '0': False, '1': True}
                if value.casefold() not in mapping:
                    raise ValueError('Boolean conversion accepts only true, false, 0, 1 and blanks')
                value = mapping[value.casefold()]
            if value not in (True, False, 0, 1):
                raise ValueError('Boolean conversion accepts only true, false, 0, 1 and blanks')
            result.append(bool(value))
        else:
            if isinstance(value, numbers.Number):
                raise ValueError('Datetime conversion requires dates, not numeric serials')
            value = pd.Timestamp(value)
            if value.tzinfo is not None:
                raise ValueError('Datetime values must be timezone-naive')
            if value.nanosecond or value.microsecond:
                raise UnsupportedModelError('Datetime authoring currently supports whole-second precision')
            result.append(value)
    dtype = {'integer': 'Int64', 'decimal': 'Float64', 'boolean': 'boolean',
             'text': 'string', 'datetime': 'datetime64[ns]', 'currency': object}[data_type]
    return pd.Series(result, dtype=dtype)


def change_semantic_metadata(image, table_id, column, semantic):
    """Update ASSL key types after replacing a stored column's physical codec."""
    replacements = {}
    physical = None
    for file in image.files:
        if not file.logical_path or not file.logical_path.endswith('.dim.xml'):
            continue
        filename = file.logical_path.rsplit('\\', 1)[-1]
        if not filename.startswith(table_id + '.'):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        for attribute in _elements(doc, 'Attribute'):
            names = [n for n in attribute.childNodes if getattr(n, 'localName', None) == 'Name']
            if names and _text(names[0]) == column:
                physical = _text(_child(attribute, 'ID'))
                for key in _elements(attribute, 'KeyColumn'):
                    _child(key, 'DataType').firstChild.data = semantic
                    _child(key, 'DataSize').firstChild.data = '0'
        replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    if physical is None:
        raise KeyError(column)
    for file in image.files:
        if not file.logical_path or not file.logical_path.endswith('.det.xml'):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        changed = False
        for dimension in _elements(doc, 'Dimension'):
            ids = [n for n in dimension.childNodes if getattr(n, 'localName', None) == 'CubeDimensionID']
            if not ids or _text(ids[0]) != table_id:
                continue
            for attribute in _elements(dimension, 'Attribute'):
                if _text(_child(attribute, 'AttributeID')) == physical:
                    for key in _elements(attribute, 'KeyColumn'):
                        _child(key, 'DataType').firstChild.data = semantic
                        _child(key, 'DataSize').firstChild.data = '0'
                        changed = True
        if changed:
            replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements))


def retype_dictionary_column(model, table, column, data_type, decimal_places, values):
    """Build a candidate first so conversion failures leave the caller intact."""
    from .empty_storage import populate_empty_table
    frame = model.get_table(table)
    calculated = {c.name for c in model.list_calculated_columns()
                  if c.table_id == model._dimension_ids[table]}
    frame = frame.drop(columns=list(calculated), errors='ignore')
    replacement = list(frame[column] if values is None else values)
    if len(replacement) != len(frame):
        raise ValueError('Type conversion must preserve the table row count')
    encoding = model._encodings[table, column]
    if not encoding.nullable and any(pd.isna(value) for value in replacement):
        raise ValueError('Blanks are invalid for a non-nullable column')
    frame[column] = converted_values(replacement, data_type, decimal_places)
    image = populate_empty_table(model, table, frame, type_overrides={column: data_type})
    semantic = {'integer': 'BigInt', 'decimal': 'Double', 'currency': 'Currency',
                'boolean': 'Boolean', 'datetime': 'Date', 'text': 'WChar'}[data_type]
    image = change_semantic_metadata(image, model._dimension_ids[table], column, semantic)
    candidate = type(model)(model.source, _image=image)
    if data_type == 'decimal':
        from .model import _ValueEncoding
        if isinstance(candidate._encodings[table, column], _ValueEncoding):
            from .schema import change_numeric_type
            image = change_numeric_type(image, model._dimension_ids[table], column,
                                        'decimal', 10 ** decimal_places)
            candidate = type(model)(model.source, _image=image)
    candidate._pending = {key: value for key, value in model._pending.items() if key[0] != table}
    candidate._pending_measures = list(model._pending_measures)
    candidate._pending_relationships = list(model._pending_relationships)
    candidate._resized_tables = dict(model._resized_tables)
    candidate.replace_table(table, frame)
    model.__dict__.update(candidate.__dict__)
    return model
