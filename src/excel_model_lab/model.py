from __future__ import annotations

from .reader_adapter import storage_schema, segment_metadata, decode_dictionary


from collections.abc import Mapping, Sequence
from contextlib import contextmanager
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
import os
import struct
from tempfile import TemporaryDirectory, mkstemp
from typing import Any
import pandas as pd

from .abf import AbfImage, UnsupportedModelError, replace_model_part
from .errors import ModelValidationError
from .dependencies import report_dependency_parts
from .dictionary import encode_numeric_dictionary, encode_uncompressed_string_dictionary
from .metadata import (
    update_column_segment_metadata,
    update_constant_column_row_count,
    update_string_column_metadata,
)
from .hash_index import encode_numeric_hash_index
from .measures import Measure, add_measures_to_script, list_script_measures
from .relationships import Relationship, add_relationships_to_dimension
from .vertipaq import encode_idf_preserving_layout, encode_pure_bitpacked_ids, select_bit_width


@dataclass(frozen=True)
class ColumnInfo:
    table: str
    name: str
    data_type: str
    row_count: int
    nullable: bool
    supported: bool
    reason: str | None = None


@dataclass(frozen=True)
class _ValueEncoding:
    storage_path: str
    hash_index_path: str
    metadata_path: str
    segments: tuple[dict, ...]
    base_id: int
    magnitude: float
    row_count: int
    nullable: bool
    logical_scale: int
    semantic_type: str

    def encode(self, values: Sequence[Any], *, allow_resize: bool = False) -> list[int]:
        if not allow_resize and len(values) != self.row_count:
            raise ValueError(
                f"Column has {self.row_count} rows but received {len(values)} values"
            )
        result: list[int] = []
        for position, value in enumerate(values):
            import pandas as pd

            if bool(pd.isna(value)):
                if not self.nullable:
                    raise ValueError(
                        f"Null at position {position} is invalid for a non-nullable column"
                    )
                result.append(2)  # XM_DATA_ID_NULL
                continue
            try:
                if self.semantic_type == "Date":
                    import pandas as pd

                    timestamp = pd.Timestamp(value)
                    epoch = pd.Timestamp("1899-12-30")
                    if timestamp.tzinfo is not None:
                        raise ValueError('Datetime values must be timezone-naive')
                    from fractions import Fraction
                    ticks = (timestamp - epoch).value
                    exact = Fraction(ticks, 86400_000_000_000) * Fraction(str(self.magnitude)) - self.base_id
                    if exact.denominator != 1:
                        raise ValueError('Datetime is not exactly representable at the stored precision')
                    result.append(exact.numerator)
                    continue
                else:
                    if self.semantic_type == 'Boolean' and value not in (True, False, 0, 1):
                        raise ValueError('Boolean values must be True, False, 0 or 1')
                    logical_value = Decimal(str(int(value) if isinstance(value, bool) else value)) * self.logical_scale
                if not logical_value.is_finite():
                    raise ValueError('Numeric values must be finite')
                encoded_decimal = (
                    logical_value * Decimal(str(self.magnitude)) - self.base_id
                )
            except (InvalidOperation, TypeError, ValueError) as exc:
                raise ValueError(f"Value {value!r} at position {position} is not numeric") from exc
            encoded = int(encoded_decimal.to_integral_value())
            if encoded_decimal != encoded:
                raise ValueError(
                    f"Value {value!r} at position {position} is not exactly representable"
                )
            result.append(encoded)
        return result


@dataclass(frozen=True)
class _StringEncoding:
    storage_path: str
    dictionary_path: str | None
    segments: tuple[dict, ...]
    row_count: int
    nullable: bool
    cardinality: int
    metadata_path: str
    original_values: tuple[str, ...]

    def encode(
        self, values: Sequence[Any], *, allow_resize: bool = False, compact: bool = False
    ) -> tuple[list[int], list[str]]:
        import pandas as pd

        if not allow_resize and len(values) != self.row_count:
            raise ValueError(
                f"Column has {self.row_count} rows but received {len(values)} values"
            )
        # Validation retains IDs; save compacts members and rebuilds indexes.
        dictionary: list[str] = [] if compact else list(self.original_values)
        ids_by_value: dict[str, int] = {value: index + 3 for index, value in enumerate(dictionary)}
        data_ids: list[int] = []
        for position, value in enumerate(values):
            if bool(pd.isna(value)):
                if not self.nullable:
                    raise ValueError(
                        f"Null at position {position} is invalid for a non-nullable column"
                    )
                data_ids.append(2)
                continue
            text = str(value)
            if text == '':
                raise UnsupportedModelError(
                    f'Empty stored text at position {position} is not supported by the Excel model writer; '
                    'choose None explicitly for BLANK or supply nonempty text')
            if text not in ids_by_value:
                ids_by_value[text] = len(dictionary) + 3
                dictionary.append(text)
            data_ids.append(ids_by_value[text])
        return data_ids, dictionary


@dataclass(frozen=True)
class _NumericDictionaryEncoding:
    storage_path: str
    dictionary_path: str | None
    segments: tuple[dict, ...]
    row_count: int
    nullable: bool
    metadata_path: str
    original_values: tuple[int | float, ...]
    semantic_type: str = 'Double'

    def encode(
        self, values: Sequence[Any], *, allow_resize: bool = False, compact: bool = False
    ) -> tuple[list[int], list[int | float]]:
        import pandas as pd

        if not allow_resize and len(values) != self.row_count:
            raise ValueError(
                f"Column has {self.row_count} rows but received {len(values)} values"
            )
        dictionary: list[int | float] = [] if compact else list(self.original_values)
        ids_by_value: dict[int | float, int] = {value: index + 3 for index, value in enumerate(dictionary)}
        data_ids: list[int] = []
        for position, value in enumerate(values):
            if bool(pd.isna(value)):
                if not self.nullable:
                    raise ValueError(
                        f"Null at position {position} is invalid for a non-nullable column"
                    )
                data_ids.append(2)
                continue
            from .numeric_storage import storage_number
            numeric = storage_number(value.item() if hasattr(value, "item") else value,
                                     self.semantic_type)
            if numeric not in ids_by_value:
                ids_by_value[numeric] = len(dictionary) + 3
                dictionary.append(numeric)
            data_ids.append(ids_by_value[numeric])
        return data_ids, dictionary


class ExcelModelTemplate:
    """Headless editor for a supported, model-enabled XLSX template.

    Mutations are accumulated in memory and emitted atomically by :meth:`save`.
    Complete table replacement may change row counts; the template continues to
    supply model schema and workbook presentation objects.
    """

    def __init__(self, source: str | Path, *, _image: AbfImage | None = None):
        from pbixray import PBIXRay

        self.source = Path(source)
        self._image = _image if _image is not None else AbfImage.from_workbook(self.source)
        self._encodings: dict[
            tuple[str, str], _ValueEncoding | _StringEncoding | _NumericDictionaryEncoding
        ] = {}
        self._columns: list[ColumnInfo] = []
        self._pending: dict[tuple[str, str], list[Any]] = {}
        self._pending_measures: list[Measure] = []
        self._pending_relationships: list[Relationship] = []
        self._resized_tables: dict[str, int] = {}
        self._dimension_ids: dict[str, str] = {}
        self._dimension_paths: dict[str, str] = {}
        self._column_ids: dict[tuple[str, str], str] = {}

        # HasNulls describes existing data, not whether a stored column can
        # accept a blank. Read the semantic key's null-processing contract.
        from .schema import _elements, _child, _text
        from xml.dom import minidom
        nullable_keys = set()
        calculated_keys = set()
        for item in self._image.files:
            if not item.logical_path or not item.logical_path.endswith('.dim.xml'):
                continue
            doc = minidom.parseString(self._image.read_logical_file(item))
            for dimension in _elements(doc, 'Dimension'):
                dimension_id = _text(_child(dimension, 'ID'))
                dimension_name = _text(_child(dimension, 'Name'))
                for attribute in _child(dimension, 'Attributes').childNodes:
                    if attribute.nodeType != attribute.ELEMENT_NODE or attribute.localName != 'Attribute':
                        continue
                    self._column_ids[dimension_name, _text(_child(attribute, 'Name'))] = _text(_child(attribute, 'ID'))
                    if _elements(attribute, 'Expression'):
                        calculated_keys.add((dimension_id, _text(_child(attribute, 'Name'))))
                    if _text(_child(attribute, 'Usage')) != 'Regular':
                        continue
                    keys = _elements(attribute, 'KeyColumn')
                    if len(keys) == 1 and _text(_child(keys[0], 'NullProcessing')) == 'Preserve':
                        nullable_keys.add((dimension_id, _text(_child(attribute, 'Name'))))

        with self._open_reader() as model:
            schema = storage_schema(model)
            for _, metadata in schema.iterrows():
                table = str(metadata["TableName"])
                dimension_id = str(metadata["DimensionID"])
                self._dimension_ids[table] = dimension_id
                column = str(metadata["ColumnName"])
                segments = segment_metadata(model, metadata)
                row_count = sum(int(segment.get("records", 0)) for segment in segments)
                nullable = (dimension_id, column) in nullable_keys or any(bool(segment.get("has_nulls")) for segment in segments)
                reason = ('Calculated columns require the calculated-column API'
                          if (dimension_id, column) in calculated_keys else
                          self._unsupported_reason(metadata, segments))
                self._columns.append(
                    ColumnInfo(
                        table=table,
                        name=column,
                        data_type=str(metadata["DataType"]),
                        row_count=row_count,
                        nullable=nullable,
                        supported=reason is None,
                        reason=reason,
                    )
                )
                if reason is None:
                    storage_path = self._resolve_storage_path(str(metadata["IDF"]))
                    if metadata["DataType"] == "string":
                        dictionary_path = (self._resolve_storage_path(str(metadata["Dictionary"]))
                                           if pd.notna(metadata["Dictionary"]) else None)
                        decoded_values = (decode_dictionary(model,
                            self._image.read_logical_file(dictionary_path), 3).values
                            if dictionary_path else {})
                        self._encodings[(table, column)] = _StringEncoding(
                            storage_path=storage_path,
                            dictionary_path=dictionary_path,
                            segments=tuple(dict(item) for item in segments),
                            row_count=row_count,
                            nullable=nullable,
                            cardinality=len(decoded_values),
                            original_values=tuple(decoded_values[index] for index in sorted(decoded_values)),
                            metadata_path=self._resolve_table_metadata(
                                str(metadata["DimensionID"])
                            ),
                        )
                    elif pd.notna(metadata["Dictionary"]) or (
                            pd.isna(metadata["HIDX"]) and segments and all(
                                bool(item.get('has_nulls')) and int(item.get('max_data_id', 2)) == 2
                                for item in segments)):
                        dictionary_path = (self._resolve_storage_path(str(metadata["Dictionary"]))
                                           if pd.notna(metadata["Dictionary"]) else None)
                        decoded_values = (decode_dictionary(model,
                            self._image.read_logical_file(dictionary_path), 3).values
                            if dictionary_path else {})
                        self._encodings[(table, column)] = _NumericDictionaryEncoding(
                            storage_path=storage_path,
                            dictionary_path=dictionary_path,
                            segments=tuple(dict(item) for item in segments),
                            row_count=row_count,
                            nullable=nullable,
                            metadata_path=self._resolve_table_metadata(dimension_id),
                            original_values=tuple(decoded_values[index] for index in sorted(decoded_values)),
                            semantic_type=('Currency' if metadata.get('SemanticType') == 'Currency' else
                                           'Double' if metadata['DataType'] == 'float64' else 'BigInt'),
                        )
                    else:
                        self._encodings[(table, column)] = _ValueEncoding(
                            storage_path=storage_path,
                            hash_index_path=self._resolve_storage_path(str(metadata["HIDX"])),
                            metadata_path=self._resolve_table_metadata(
                                str(metadata["DimensionID"])
                            ),
                            segments=tuple(dict(item) for item in segments),
                            base_id=int(metadata["BaseId"]),
                            magnitude=float(metadata["Magnitude"]),
                            row_count=row_count,
                            nullable=nullable,
                            logical_scale=(10000 if metadata.get("SemanticType") == "Currency" else 1),
                            semantic_type=('Boolean' if metadata['DataType'] == 'bool' else
                                           ('BigInt' if metadata['DataType'] == 'int64' else 'Double')
                                           if metadata.get('SemanticType') in (None, 'Other') else str(metadata['SemanticType'])),
                        )

            for table, dimension_id in self._dimension_ids.items():
                matches = []
                for item in self._image.files:
                    if not item.logical_path:
                        continue
                    name = item.logical_path.replace("/", "\\").rsplit("\\", 1)[-1]
                    if name.startswith(dimension_id + ".") and name.endswith(".dim.xml"):
                        matches.append(item.storage_path)
                if len(matches) == 1:
                    self._dimension_paths[table] = matches[0]

    @classmethod
    def open(cls, source, *, read_only=True, on_disk=False, temp_dir=None):
        """Open a saved model for inspection, or an XLSX model for editing.

        Read-only mode retains PBIXRay's format coverage and streaming options.
        Editing uses this writer's narrower Excel model compatibility contract.
        """
        if not isinstance(read_only, bool):
            raise TypeError('read_only must be a boolean')
        if read_only:
            from .reader import ModelReader
            return ModelReader(source, on_disk=on_disk, temp_dir=temp_dir)
        if on_disk or temp_dir is not None:
            raise ValueError('on_disk and temp_dir are available only in read-only mode')
        return cls(source)

    @classmethod
    def from_template(cls, source: str | Path) -> "ExcelModelTemplate":
        return cls(source)

    @classmethod
    def create(cls, tables: Mapping[str, Any], *, workbook: str | Path | None = None) -> "ExcelModelTemplate":
        """Create a model from dataframes without a caller-supplied workbook.

        Tables may be empty when their dataframe columns have explicit types.
        Excel is not required at runtime. The bundled seed contains synthetic
        encoding structures; all seed columns and rows are removed from output.
        """
        import uuid
        from pandas.api.types import is_integer_dtype, is_string_dtype, is_float_dtype, is_bool_dtype, is_datetime64_any_dtype

        if not isinstance(tables, Mapping) or not tables:
            raise ValueError('Provide at least one named dataframe')
        if any(not isinstance(name, str) or not name.strip() for name in tables):
            raise ValueError('Every table needs a nonempty name')
        if len({name.casefold() for name in tables}) != len(tables):
            raise ValueError('Table names must be unique ignoring case')
        seed = Path(__file__).parent / 'data' / 'authoring-seed.xlsx'
        model = cls(seed)
        integer = '_SeedInteger_' + uuid.uuid4().hex
        text = '_SeedText_' + uuid.uuid4().hex
        model.rename_column('Table1', 'ID', integer)
        model.rename_column('Table1', 'Label', text)
        captions = []
        for index, (name, data) in enumerate(tables.items(), 1):
            data = cls._as_dataframe(data)
            if not isinstance(data, pd.DataFrame) or not len(data.columns):
                raise ValueError('New tables require a dataframe with typed columns')
            if any(not isinstance(c, str) for c in data.columns) or len({c.casefold() for c in data.columns}) != len(data.columns):
                raise ValueError('Column names must be unique strings ignoring case')
            physical_name = '_ModelTable_' + uuid.uuid4().hex
            model.add_table(physical_name, pd.DataFrame({integer: [1, 2, 3], text: ['Alpha', 'Beta', 'Gamma']}), like='Table1')
            for column in data.columns:
                values = data[column]
                new_type = None
                if is_bool_dtype(values.dtype):
                    prototype, new_type = integer, 'boolean'
                elif is_datetime64_any_dtype(values.dtype):
                    prototype, new_type = integer, 'datetime'
                elif len(values.dropna()) and all(isinstance(v, Decimal) for v in values.dropna()):
                    prototype, new_type = integer, 'currency'
                elif is_integer_dtype(values.dtype) or is_float_dtype(values.dtype):
                    prototype = integer
                    if is_float_dtype(values.dtype):
                        new_type = 'decimal'
                elif is_string_dtype(values.dtype) and all(isinstance(v, str) for v in values.dropna()):
                    prototype = text
                else:
                    raise UnsupportedModelError(f'New column type is not yet supported: {values.dtype}')
                numeric_kind = new_type or ('integer' if prototype == integer else None)
                if numeric_kind in {'integer', 'decimal', 'currency'}:
                    from .numeric_storage import needs_numeric_dictionary, convert_string_seed_to_numeric
                    if needs_numeric_dictionary(values, numeric_kind):
                        model.add_column(physical_name, column, ['Alpha', 'Beta', 'Gamma'], like=text)
                        image = convert_string_seed_to_numeric(model._image,
                            model._dimension_ids[physical_name], model._column_ids[(physical_name, column)], numeric_kind)
                        candidate = cls(model.source, _image=image)
                        candidate._pending = dict(model._pending)
                        candidate._pending[(physical_name, column)] = ([Decimal(1), Decimal(2), Decimal(3)]
                            if numeric_kind == 'currency' else [1, 2, 3])
                        candidate._pending_measures = list(model._pending_measures)
                        candidate._pending_relationships = list(model._pending_relationships)
                        candidate._resized_tables = dict(model._resized_tables)
                        model = candidate
                        continue
                if new_type:
                    model.add_column(physical_name, column, [1, 2, 3], like=prototype)
                    initial_values = [values.iloc[0] if len(values) else None] * 3
                    if new_type == 'datetime':
                        timestamps = [pd.Timestamp(v) for v in values.dropna()]
                        if any(v.tzinfo is not None for v in timestamps):
                            raise ValueError('Datetime values must be timezone-naive')
                        if any(v.microsecond or v.nanosecond for v in timestamps):
                            raise UnsupportedModelError('Datetime authoring currently supports whole-second precision')
                        if any(v != v.normalize() for v in timestamps):
                            initial_values = [pd.Timestamp('1899-12-30 00:00:01')] * 3
                    model.change_column_type(physical_name, column, new_type,
                        decimal_places=cls._decimal_places(values) if new_type == 'decimal' else 0,
                        values=initial_values)
                else:
                    model.add_column(physical_name, column, [values.iloc[0] if len(values) else None] * 3, like=prototype)
            model.remove_column(physical_name, integer).remove_column(physical_name, text)
            model.replace_table(physical_name, data)
            captions.append((physical_name, name))
        model.remove_table('Table1')
        for physical_name, name in captions:
            model.rename_table(physical_name, name)
        if workbook is not None:
            from zipfile import ZipFile
            with ZipFile(workbook) as archive:
                if 'xl/model/item.data' in archive.namelist():
                    raise ValueError('Workbook already contains a model; open it and add tables instead')
                if 'xl/workbook.xml' not in archive.namelist():
                    raise ValueError('Not an Excel workbook')
            model.source = Path(workbook)
        return model

    def _open_reader(self):
        """Decode the current schema, including unsaved structural edits."""
        from io import BytesIO
        from zipfile import ZipFile, ZIP_STORED
        from pbixray import PBIXRay

        envelope = BytesIO()
        with ZipFile(envelope, 'w', compression=ZIP_STORED) as archive:
            archive.writestr('xl/model/item.data', self._image.data)
        envelope.seek(0)
        reader = PBIXRay(envelope)
        self._normalize_reader_types(reader)
        return reader

    @staticmethod
    def _normalize_reader_types(reader):
        from .reader_adapter import normalize_reader_types
        normalize_reader_types(reader)

    @staticmethod
    def _decimal_places(values) -> int:
        places = 0
        for value in values:
            if pd.isna(value):
                continue
            number = Decimal(str(value))
            if not number.is_finite():
                raise ValueError('Numeric values must be finite')
            places = max(places, -number.normalize().as_tuple().exponent)
        if places > 15:
            raise UnsupportedModelError('Decimal values require more than 15 decimal places')
        return places

    def change_column_type(self, table: str, column: str, data_type: str, *,
                           decimal_places: int = 6, values: Sequence[Any] | None = None) -> "ExcelModelTemplate":
        """Convert a stored column without rounding or discarding blanks.

        ``decimal_places`` determines the stored decimal precision (0–15).
        Conversions that lose fractional values fail without changing the model.
        """
        from .schema import change_numeric_type
        if data_type not in ('integer', 'decimal', 'datetime', 'currency', 'boolean', 'text'):
            raise UnsupportedModelError('Unknown stored data type')
        if isinstance(decimal_places, bool) or not isinstance(decimal_places, int) or not 0 <= decimal_places <= 15:
            raise ValueError('decimal_places must be an integer from 0 to 15')
        encoding = self._encodings.get((table, column))
        if not isinstance(encoding, (_ValueEncoding, _NumericDictionaryEncoding, _StringEncoding)):
            raise UnsupportedModelError('Conversion requires a supported stored column')
        from .relationship_storage import stored_relationships, relationship_key_type
        semantic = {'integer': 'BigInt', 'decimal': 'Double', 'datetime': 'Date', 'currency': 'Currency', 'boolean': 'Boolean', 'text': 'WChar'}[data_type]
        if semantic != relationship_key_type(self, table, column) and any(
                (table, column) in ((r.from_table, r.from_column), (r.to_table, r.to_column))
                for r in [item for _, item in stored_relationships(self)] + self._pending_relationships):
            raise ValueError('Remove relationships using this column before changing its data type')
        if semantic != 'Date' and any(edge.reason == 'date table key' and
                edge.target == ('tables', table, 'columns', column) for edge in self.dependency_graph().edges):
            raise ValueError('Unmark the date table before changing its key data type')
        if not isinstance(encoding, _ValueEncoding) or data_type == 'text':
            from .numeric_storage import retype_dictionary_column
            return retype_dictionary_column(self, table, column, data_type, decimal_places, values)
        with self._open_reader() as reader:
            data = reader.get_table(table)
        calculated = {item.name for item in self.list_calculated_columns()
                      if item.table_id == self._dimension_ids[table]}
        data = data.drop(columns=list(calculated), errors='ignore')
        columns = {name: self._pending.get((table, name), data[name].tolist()) for name in data.columns}
        if values is not None:
            replacement = list(values)
            if len(replacement) != len(columns[column]):
                raise ValueError('Type conversion must preserve the table row count')
            columns[column] = replacement
        if data_type == 'boolean':
            if any(not pd.isna(v) and v not in (True, False, 0, 1) for v in columns[column]):
                raise ValueError('Boolean conversion accepts only True, False, 0, 1 and blanks')
            columns[column] = [None if pd.isna(v) else bool(v) for v in columns[column]]
        if data_type == 'datetime':
            import numbers
            if any(isinstance(v, numbers.Number) and not pd.isna(v) for v in columns[column]):
                raise ValueError('Datetime conversion requires dates, not numeric serials')
            columns[column] = [pd.NaT if pd.isna(v) else pd.Timestamp(v) for v in columns[column]]
        magnitude = 10 ** decimal_places if data_type == 'decimal' else 1
        if data_type == 'datetime':
            from math import gcd
            day_ticks = 86400_000_000_000
            resolution = day_ticks
            for value in columns[column]:
                if pd.isna(value):
                    continue
                if value.tzinfo is not None:
                    raise ValueError('Datetime values must be timezone-naive')
                if value.nanosecond or value.microsecond:
                    raise UnsupportedModelError('Datetime authoring currently supports whole-second precision')
                resolution = gcd(resolution, (value - pd.Timestamp('1899-12-30')).value)
            magnitude = day_ticks // resolution
        image = change_numeric_type(self._image, self._dimension_ids[table], column, data_type, magnitude)
        candidate = type(self)(self.source, _image=image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        candidate.replace_table(table, pd.DataFrame(columns))
        self.__dict__.update(candidate.__dict__)
        return self

    def add_column(self, table: str, name: str, values: Sequence[Any], *,
                   like: str | None = None) -> "ExcelModelTemplate":
        """Add a stored column using an existing column's type and encoding.

        ``like`` identifies a column in the same table. Values must satisfy its
        type and nullability rules. This operation does not create DAX columns.
        Changes remain in memory until save; invalid input leaves this object
        untouched. The original worksheet source is not refreshed or changed.
        """
        from .schema import clone_column

        materialized = list(values)
        if like is None:
            from .schema import clone_table
            from .column_import import import_column
            from io import BytesIO
            from zipfile import ZipFile
            if table not in self._dimension_ids:
                raise KeyError(table)
            if not materialized and getattr(values, 'dtype', None) is None:
                raise ValueError('Empty new columns require an explicitly typed pandas Series or array')
            original_count = next(c.row_count for c in self._columns if c.table == table)
            expected_count = self._resized_tables.get(table, original_count)
            if len(materialized) != expected_count:
                raise ValueError(f'New column needs {expected_count} values, received {len(materialized)}')
            data = pd.DataFrame({name: pd.Series(materialized, dtype=getattr(values, 'dtype', None))})
            donor = type(self).create({'ColumnSource': data})
            donor_id = donor._dimension_ids['ColumnSource']
            donor.rename_table('ColumnSource', donor_id)
            target_id = self._dimension_ids[table]
            donor_image = clone_table(donor._image, donor_id, target_id)
            prepared = type(self)(donor.source, _image=donor_image)
            # Materialize a column with the target's current physical row count;
            # the caller's pending resize is applied with the remaining columns.
            prepared.replace_table(target_id, pd.DataFrame({name: [materialized[0] if materialized else None] * original_count}))
            with ZipFile(BytesIO(prepared.to_bytes())) as archive:
                donor_image = AbfImage(archive.read('xl/model/item.data'))
            image = import_column(self._image, donor_image, target_id, donor._column_ids[(donor_id, name)])
        else:
            if (table, like) not in self._encodings:
                raise UnsupportedModelError(f'No writable source column: {table}.{like}')
            image = clone_column(self._image, self._dimension_ids[table],
                                 self._column_ids.get((table, like), like), name)
        candidate = type(self)(self.source, _image=image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        with candidate._open_reader() as reader:
            data = reader.get_table(table)
        expected_count = self._resized_tables.get(table, len(data))
        if len(materialized) != expected_count:
            raise ValueError(f'New column needs {expected_count} values, received {len(materialized)}')
        columns = {column: candidate._pending.get((table, column), data[column].tolist())
                   for column in data.columns if column != name}
        columns[name] = materialized
        candidate.replace_table(table, pd.DataFrame(columns))
        self.__dict__.update(candidate.__dict__)
        return self

    def remove_column(self, table: str, name: str) -> "ExcelModelTemplate":
        """Remove an independent stored column without cascading dependencies."""
        from zipfile import ZipFile
        from .schema import remove_column

        if (table, name) not in self._encodings:
            raise UnsupportedModelError(f'No writable column: {table}.{name}')
        self.dependency_graph().assert_removable(('tables', table, 'columns', name))
        with ZipFile(self.source) as archive:
            if any(path.startswith(('xl/pivotTables/', 'xl/pivotCache/', 'xl/slicers/'))
                   for path in archive.namelist()):
                raise UnsupportedModelError('Column removal from workbooks with PivotTables or slicers is not yet supported')
        if any((table, name) in ((r.from_table, r.from_column), (r.to_table, r.to_column))
               for r in self._pending_relationships):
            raise UnsupportedModelError('Column is used by a pending relationship')
        # Schema storage is keyed by physical attribute ID, which can differ
        # from the caption after a native rename in Power Pivot.
        image = remove_column(self._image, self._dimension_ids[table],
                              self._column_ids.get((table, name), name))
        candidate = type(self)(self.source, _image=image)
        candidate._pending = {key: value for key, value in self._pending.items() if key != (table, name)}
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        self.__dict__.update(candidate.__dict__)
        return self

    def rename_table(self, table: str, name: str) -> "ExcelModelTemplate":
        """Rename a table and its DAX references, preserving physical identity."""
        from zipfile import ZipFile
        from .schema import rename_table

        if table not in self._dimension_ids:
            raise KeyError(table)
        if any(t.casefold() == name.casefold() for t in self._dimension_ids if t != table):
            raise ValueError(f'Table already exists: {name}')
        with ZipFile(self.source) as archive:
            if report_dependency_parts(archive):
                raise UnsupportedModelError('Table rename with presentation dependencies is not yet supported')
        image = rename_table(self._image, self._dimension_ids[table], name)
        from .structural_edit import rebind_table_references
        image = rebind_table_references(image, table, name, table_names=self.tables)
        candidate = type(self)(self.source, _image=image)
        candidate._pending = {(name if t == table else t, c): v for (t, c), v in self._pending.items()}
        candidate._resized_tables = {name if t == table else t: n for t, n in self._resized_tables.items()}
        from dataclasses import replace
        from .formula_statements import rewrite_references
        candidate._pending_measures = [replace(m, table=name if m.table == table else m.table,
            expression=rewrite_references(m.expression, table_names=self.tables, table_renames={table: name}))
            for m in self._pending_measures]
        candidate._pending_relationships = [replace(r, from_table=name if r.from_table == table else r.from_table,
            to_table=name if r.to_table == table else r.to_table) for r in self._pending_relationships]
        self.__dict__.update(candidate.__dict__)
        return self

    def rename_column(self, table: str, column: str, name: str) -> "ExcelModelTemplate":
        """Rename a stored column and unambiguous DAX references.

        Relationships, sorts, hierarchies, perspectives, source schema and report
        dependencies remain restrictive. The column moves to the end of its
        table's column order because its storage is rebuilt under the new name.
        """
        candidate = type(self)(self.source, _image=self._image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        with candidate._open_reader() as reader:
            values = candidate._pending.get((table, column), reader.get_table(table)[column].tolist())
        candidate.add_column(table, name, values, like=column)
        from .structural_edit import rebind_column_references, rebind_column_expression
        from dataclasses import replace
        graph = candidate.dependency_graph()
        measure_names = [obj.path[3] for obj in graph.objects
                         if len(obj.path) == 4 and obj.path[2] == 'measures']
        candidate._image = rebind_column_references(candidate._image, table, column, name,
            table_names=candidate.tables, measure_names=measure_names)
        candidate._pending_measures = [replace(measure, expression=rebind_column_expression(
            measure.expression, measure.table, table, column, name,
            table_names=candidate.tables, measure_names=measure_names)) for measure in candidate._pending_measures]
        candidate.remove_column(table, column)
        self.__dict__.update(candidate.__dict__)
        return self

    def configure_column(self, table: str, column: str, *, hidden: bool | None = None,
                         description: str | None = None, format_string: str | None = None) -> "ExcelModelTemplate":
        """Set field visibility, description, and a display format string."""
        from .schema import configure_column
        if (table, column) not in self._encodings:
            raise KeyError((table, column))
        image = configure_column(self._image, self._dimension_ids[table],
            self._column_ids.get((table, column), column), hidden=hidden, description=description, format_string=format_string)
        self._image = image
        return self

    def remove_table(self, table: str) -> "ExcelModelTemplate":
        """Remove an independent table and its data; never cascade dependencies."""
        from zipfile import ZipFile
        from .schema import remove_table

        if table not in self._dimension_ids:
            raise KeyError(table)
        self.dependency_graph().assert_removable(('tables', table))
        with ZipFile(self.source) as archive:
            if report_dependency_parts(archive):
                raise UnsupportedModelError('Table removal with presentation dependencies is not yet supported')
        if any(table in (r.from_table, r.to_table) for r in self._pending_relationships):
            raise UnsupportedModelError('Table is used by a pending relationship')
        image = remove_table(self._image, self._dimension_ids[table])
        candidate = type(self)(self.source, _image=image)
        candidate._pending = {key: value for key, value in self._pending.items() if key[0] != table}
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = {t: n for t, n in self._resized_tables.items() if t != table}
        self.__dict__.update(candidate.__dict__)
        return self

    def add_table(self, name: str, dataframe: Any, *, like: str | None = None) -> "ExcelModelTemplate":
        """Add a dataframe as a new table, optionally copying a table schema."""
        from .schema import clone_table, import_table

        if not isinstance(name, str) or not name.strip():
            raise ValueError('A table needs a nonempty name')
        if name.casefold() in {t.casefold() for t in self._dimension_ids}:
            raise ValueError(f'Table already exists: {name}')
        if like is None:
            donor = type(self).create({name: dataframe})
            image = import_table(self._image, donor._image, donor._dimension_ids[name])
            candidate = type(self)(self.source, _image=image)
            candidate._pending = {**self._pending, **donor._pending}
            candidate._pending_measures = list(self._pending_measures)
            candidate._pending_relationships = list(self._pending_relationships)
            candidate._resized_tables = {**self._resized_tables, **donor._resized_tables}
            self.__dict__.update(candidate.__dict__)
            return self

        if like not in self._dimension_ids:
            raise KeyError(like)
        image = clone_table(self._image, self._dimension_ids[like], name)
        candidate = type(self)(self.source, _image=image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        candidate.replace_table(name, dataframe)
        self.__dict__.update(candidate.__dict__)
        return self

    @staticmethod
    def _as_dataframe(value: Any) -> Any:
        if hasattr(value, "to_pandas") and not hasattr(value, "columns"):
            return value.to_pandas()
        return value

    @staticmethod
    def _unsupported_reason(metadata: Any, segments: Sequence[Mapping[str, Any]]) -> str | None:
        if segments and not any(int(item.get('records', 0)) for item in segments):
            return "native empty column requires complete-table replacement to build dictionary/index"
        is_string_dictionary = metadata["DataType"] == "string" and (
            pd.notna(metadata["Dictionary"]) or (segments and all(
                bool(item.get('has_nulls')) and int(item.get('max_data_id', 2)) == 2
                for item in segments)))
        is_numeric = metadata["DataType"] in {"int64", "float64", "datetime64[ns]", "bool"}
        if not (is_string_dictionary or is_numeric):
            return "only value-encoded numeric columns are currently writable"
        if not segments:
            return "column has no writable segments"
        if any(not 1 <= int(item.get("bit_width", 0)) <= 63 for item in segments):
            return "constant or unknown-width columns are not yet writable"
        return None

    def _resolve_storage_path(self, idf_name: str) -> str:
        matches = [
            item.storage_path
            for item in self._image.files
            if item.logical_path
            and item.logical_path.replace("/", "\\").endswith("\\" + idf_name)
        ]
        if len(matches) != 1:
            raise KeyError(f"Could not uniquely resolve physical storage for {idf_name}")
        return matches[0]

    def _resolve_table_metadata(self, dimension_id: str) -> str:
        matches = []
        for item in self._image.files:
            if not item.logical_path:
                continue
            name = item.logical_path.replace("/", "\\").rsplit("\\", 1)[-1]
            if name.startswith(dimension_id + ".") and name.endswith(".tbl.xml"):
                matches.append(item.storage_path)
        if len(matches) != 1:
            raise KeyError(
                f"Could not uniquely resolve table metadata for {dimension_id}"
            )
        return matches[0]

    @property
    def columns(self) -> tuple[ColumnInfo, ...]:
        return tuple(self._columns)

    def dependency_graph(self):
        """Return an immutable snapshot of model and workbook dependencies.

        DAX references are resolved lexically against model names. Ambiguous and
        virtual-column references remain explicit; this does not evaluate DAX.
        """
        from .dependencies import build_dependency_graph
        return build_dependency_graph(self)

    @contextmanager
    def transaction(self):
        """Stage a group of in-memory edits and commit only on normal exit.

        Use the yielded model for every edit. An exception discards its changes.
        Call ``save`` after exiting the context: files explicitly saved inside a
        transaction are external side effects and cannot be rolled back.
        """
        from copy import deepcopy
        candidate = type(self)(self.source, _image=self._image)
        candidate._pending = deepcopy(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        yield candidate
        self.__dict__.update(candidate.__dict__)

    def _materialized_candidate(self):
        """Build an isolated candidate with all pending data already encoded."""
        if self._pending or self._pending_measures or self._pending_relationships or self._resized_tables:
            with TemporaryDirectory(prefix='excelmodel-stage-') as directory:
                path = self.save(Path(directory) / 'model.xlsx')
                image = AbfImage.from_workbook(path)
        else:
            image = self._image
        return type(self)(self.source, _image=image)

    def list_calculated_columns(self):
        """Inspect definitions and processing state; this does not evaluate DAX."""
        from .calculated_columns import list_calculated_columns
        return list_calculated_columns(self._image)

    def add_calculated_column(self, table, name, expression, *, data_type):
        """Add a typed DAX column with BLANK cache pending calculation in Excel."""
        from .calculated_columns import add_calculated_column
        candidate = self._materialized_candidate()
        image = add_calculated_column(candidate._image, candidate._dimension_ids[table],
                                      name, expression, data_type=data_type)
        result = type(self)(self.source, _image=image)
        self.__dict__.update(result.__dict__)
        return self

    def update_calculated_column(self, table, name, expression, *, data_type):
        """Change a DAX column definition and invalidate all calculated caches."""
        from .calculated_columns import update_calculated_column, invalidate_calculated_columns
        candidate = self._materialized_candidate()
        image = update_calculated_column(candidate._image, candidate._dimension_ids[table],
                                         name, expression, data_type=data_type)
        image = invalidate_calculated_columns(image)
        result = type(self)(self.source, _image=image)
        self.__dict__.update(result.__dict__)
        return self

    def remove_calculated_column(self, table, name):
        """Remove a DAX column only when no model or report object uses it."""
        from .calculated_columns import remove_calculated_column
        self.dependency_graph().assert_removable(('tables', table, 'columns', name))
        candidate = self._materialized_candidate()
        image = remove_calculated_column(candidate._image, candidate._dimension_ids[table], name)
        result = type(self)(self.source, _image=image)
        self.__dict__.update(result.__dict__)
        return self

    def _validate_date_table_key(self, table, column):
        from .date_tables import validate_date_table_key
        return validate_date_table_key(self, table, column)

    def mark_date_table(self, table, column):
        """Mark a stored date key after checking unique, contiguous whole days."""
        from .date_tables import configure_date_table
        self._validate_date_table_key(table, column)
        candidate = self._materialized_candidate()
        from .schema import _elements, _child, _text
        from xml.dom import minidom
        document = minidom.parseString(candidate._image.read_logical_file(candidate._dimension_paths[table]))
        attribute = next(a for a in _elements(document, 'Attribute')
                         if (_text(_child(a, 'Name')) or _text(_child(a, 'ID'))) == column)
        image = configure_date_table(candidate._image, candidate._dimension_ids[table], _text(_child(attribute, 'ID')))
        self.__dict__.update(type(self)(self.source, _image=image).__dict__)
        return self

    def unmark_date_table(self, table):
        """Remove the date-table designation while preserving its data."""
        from .date_tables import configure_date_table, _error
        if table not in self._dimension_ids:
            _error(table, None, 'Date table does not exist', lookup=True)
        candidate = self._materialized_candidate()
        image = configure_date_table(candidate._image, candidate._dimension_ids[table], None)
        self.__dict__.update(type(self)(self.source, _image=image).__dict__)
        return self

    def list_perspectives(self):
        """List named selections of model fields without changing the model."""
        from .perspectives import list_perspectives
        return list_perspectives(self._image)

    def create_perspective(self, name, *, columns=(), measures=(), hierarchies=(), description=''):
        """Create a perspective from explicit (table, object name) selections."""
        return self._set_perspective(name, columns=columns, measures=measures,
                                     hierarchies=hierarchies, description=description, replace=False)

    def update_perspective(self, name, *, columns=(), measures=(), hierarchies=(), description=''):
        """Replace all selections in an existing perspective atomically."""
        return self._set_perspective(name, columns=columns, measures=measures,
                                     hierarchies=hierarchies, description=description, replace=True)

    def _set_perspective(self, name, *, columns, measures, hierarchies, description, replace):
        from .perspectives import set_perspective
        if replace:
            canonical = next((n for n in self.list_perspectives() if n.casefold() == name.casefold()), name)
            self.dependency_graph().assert_removable(('perspectives', canonical))
        candidate = self._materialized_candidate()
        image = set_perspective(candidate._image, name, columns=columns, measures=measures,
                                hierarchies=hierarchies, description=description, replace=replace)
        self.__dict__.update(type(self)(self.source, _image=image).__dict__)
        return self

    def remove_perspective(self, name):
        """Remove a perspective, rejecting report bindings that may use it."""
        from .perspectives import remove_perspective
        canonical = next((n for n in self.list_perspectives() if n.casefold() == name.casefold()), name)
        self.dependency_graph().assert_removable(('perspectives', canonical))
        candidate = self._materialized_candidate()
        image = remove_perspective(candidate._image, name)
        self.__dict__.update(type(self)(self.source, _image=image).__dict__)
        return self

    def rename_perspective(self, name, new_name):
        """Rename a perspective while retaining its identity and selections."""
        from .perspectives import rename_perspective
        canonical = next((n for n in self.list_perspectives() if n.casefold() == name.casefold()), name)
        self.dependency_graph().assert_removable(('perspectives', canonical))
        candidate = self._materialized_candidate()
        image = rename_perspective(candidate._image, name, new_name)
        self.__dict__.update(type(self)(self.source, _image=image).__dict__)
        return self

    def replace_column(
        self,
        table: str,
        column: str,
        values: Sequence[Any],
    ) -> "ExcelModelTemplate":
        key = (table, column)
        encoding = self._encodings.get(key)
        if encoding is None:
            info = next(
                (item for item in self._columns if (item.table, item.name) == key), None
            )
            if info is None:
                raise KeyError(f"Model column {table}.{column} does not exist")
            raise UnsupportedModelError(f"Cannot write {table}.{column}: {info.reason}")
        materialized = list(values)
        expected_count = self._resized_tables.get(table, encoding.row_count)
        if len(materialized) != expected_count:
            raise ValueError(f'Column has {expected_count} rows but received {len(materialized)} values')
        if self._needs_numeric_storage_rebuild(encoding, materialized):
            calculated = {c.name for c in self.list_calculated_columns()
                          if c.table_id == self._dimension_ids[table]}
            frame = self.get_table(table).drop(columns=list(calculated), errors='ignore')
            frame[column] = pd.Series(materialized, dtype=object)
            return self.replace_table(table, frame)
        encoding.encode(materialized, allow_resize=table in self._resized_tables)
        from .calculated_columns import list_calculated_columns, invalidate_calculated_columns
        if list_calculated_columns(self._image):
            self._image = invalidate_calculated_columns(self._image, changed_tables={self._dimension_ids[table]})
        self._pending[key] = materialized
        return self

    @staticmethod
    def _needs_numeric_storage_rebuild(encoding, values):
        if isinstance(encoding, (_StringEncoding, _NumericDictionaryEncoding)):
            return encoding.dictionary_path is None and any(not pd.isna(value) for value in values)
        if not isinstance(encoding, _ValueEncoding):
            return False
        kind = {'BigInt': 'integer', 'Double': 'decimal', 'Currency': 'currency'}.get(encoding.semantic_type)
        if kind is None:
            return False
        # Preserve the existing nullability contract before choosing a codec.
        if not encoding.nullable and any(pd.isna(value) for value in values):
            return False
        try:
            encoding.encode(values, allow_resize=True)
        except ValueError:
            # Promotion expands the numeric range, never its declared precision.
            # The ordinary encoder below reports the original validation error.
            return False
        from .numeric_storage import needs_numeric_dictionary
        return needs_numeric_dictionary(values, kind)

    def replace_columns(
        self,
        table: str,
        columns: Mapping[str, Sequence[Any]],
    ) -> "ExcelModelTemplate":
        staged = dict(self._pending)
        staged_state = dict(self.__dict__)
        staged_image = self._image
        try:
            for column, values in columns.items():
                self.replace_column(table, column, values)
        except Exception:
            self.__dict__.update(staged_state)
            self._pending = staged
            self._image = staged_image
            raise
        return self

    def replace_dataframe(
        self,
        table: str,
        dataframe: Any,
        *,
        columns: Sequence[str] | None = None,
    ) -> "ExcelModelTemplate":
        dataframe = self._as_dataframe(dataframe)
        selected = list(columns) if columns is not None else list(dataframe.columns)
        missing = [column for column in selected if column not in dataframe.columns]
        if missing:
            raise KeyError(f"DataFrame is missing columns: {', '.join(missing)}")
        return self.replace_columns(
            table,
            {column: dataframe[column].tolist() for column in selected},
        )

    def replace_table(self, table: str, dataframe: Any) -> "ExcelModelTemplate":
        """Replace every writable user column, allowing the row count to change."""
        dataframe = self._as_dataframe(dataframe)
        from .calculated_columns import list_calculated_columns, invalidate_calculated_columns
        calculated = list_calculated_columns(self._image)
        calculated_names = {c.name for c in calculated if c.table_id == self._dimension_ids.get(table)}
        expected = [
            item.name for item in self._columns
            if item.table == table and not item.name.startswith("__XL_") and item.name not in calculated_names
        ]
        if not expected:
            raise KeyError(f"Model table {table} does not exist")
        unsupported = [
            item.name for item in self._columns
            if item.table == table and not item.name.startswith("__XL_") and not item.supported and item.name not in calculated_names
        ]
        missing = [column for column in expected if column not in dataframe.columns]
        extra = [column for column in dataframe.columns if column not in expected]
        if missing or extra:
            details = []
            if missing:
                details.append("missing " + ", ".join(missing))
            if extra:
                details.append("unexpected " + ", ".join(extra))
            raise ValueError("DataFrame schema mismatch: " + "; ".join(details))
        if calculated_names:
            original_count = next(c.row_count for c in self._columns if c.table == table and c.name in expected)
            if (len(dataframe) == original_count and table not in self._resized_tables and
                    not any(self._needs_numeric_storage_rebuild(self._encodings.get((table, name)), dataframe[name].tolist())
                            for name in expected)):
                return self.replace_columns(table, {name: dataframe[name].tolist() for name in expected})
        if unsupported:
            columns = [c for c in self._columns if c.table == table and c.name in unsupported]
            if all(c.row_count == 0 and (c.reason or '').startswith('native empty column') for c in columns):
                if not len(dataframe):
                    return self
                from .empty_storage import populate_empty_table
                image = populate_empty_table(self, table, dataframe)
                candidate = type(self)(self.source, _image=image)
                candidate._pending = dict(self._pending)
                candidate._pending_measures = list(self._pending_measures)
                candidate._pending_relationships = list(self._pending_relationships)
                candidate._resized_tables = dict(self._resized_tables)
                candidate.replace_table(table, dataframe)
                self.__dict__.update(candidate.__dict__)
                return self
            raise UnsupportedModelError(
                f"Table has columns that cannot be rebuilt: {', '.join(unsupported)}"
            )
        row_count = len(dataframe)
        if any(self._needs_numeric_storage_rebuild(self._encodings[table, name], dataframe[name].tolist())
               for name in expected):
            from .empty_storage import populate_empty_table
            image = populate_empty_table(self, table, dataframe)
            candidate = type(self)(self.source, _image=image)
            candidate._pending = {key: value for key, value in self._pending.items() if key[0] != table}
            candidate._pending_measures = list(self._pending_measures)
            candidate._pending_relationships = list(self._pending_relationships)
            candidate._resized_tables = dict(self._resized_tables)
            candidate.replace_table(table, dataframe)
            self.__dict__.update(candidate.__dict__)
            return self
        staged = dict(self._pending)
        try:
            for column in expected:
                key = (table, column)
                values = dataframe[column].tolist()
                self._encodings[key].encode(values, allow_resize=True)
                self._pending[key] = values
        except Exception:
            self._pending = staged
            raise
        if calculated:
            try:
                image = invalidate_calculated_columns(self._image, changed_tables={self._dimension_ids[table]})
            except Exception:
                self._pending = staged
                raise
            self._image = image
        self._resized_tables[table] = row_count
        return self

    def replace_tables(self, tables: Mapping[str, Any]) -> "ExcelModelTemplate":
        """Stage complete replacements for multiple model tables."""
        staged_state = dict(self.__dict__)
        staged_pending = dict(self._pending)
        staged_resized = dict(self._resized_tables)
        try:
            for table, dataframe in tables.items():
                self.replace_table(table, dataframe)
        except Exception:
            self.__dict__.update(staged_state)
            self._pending = staged_pending
            self._resized_tables = staged_resized
            raise
        return self

    @property
    def tables(self) -> tuple[str, ...]:
        """Table names in model order, including unsaved schema edits."""
        return tuple(self._dimension_ids)

    def get_table(self, table: str) -> pd.DataFrame:
        """Return stored data with staged edits applied, without running DAX."""
        if table not in self._dimension_ids:
            raise KeyError(table)
        with self._open_reader() as reader:
            frame = reader.get_table(table)
        values = {column: list(self._pending.get((table, column), frame[column].tolist()))
                  for column in frame.columns if not column.startswith('__XL_')}
        if table in self._resized_tables:
            for column in self.list_calculated_columns():
                if column.table_id == self._dimension_ids[table]:
                    values[column.name] = [None] * self._resized_tables[table]
        for column, items in values.items():
            encoding = self._encodings.get((table, column))
            if (isinstance(encoding, _NumericDictionaryEncoding) and
                    encoding.semantic_type == 'BigInt' and any(pd.isna(v) for v in items)):
                values[column] = pd.array(items, dtype='Int64')
        return pd.DataFrame(values)

    def refresh_sources(self, sources: Mapping[str, Any]) -> 'ExcelModelTemplate':
        """Read explicit sources and stage all replacements atomically.

        Sources implement read() -> DataFrame. Nothing is written until save().
        This does not invoke Excel refresh or stored external connections.
        """
        unknown = set(sources) - set(self.tables)
        if unknown:
            raise KeyError(f'Unknown refresh tables: {sorted(unknown)}')
        frames = {table: source.read() for table, source in sources.items()}
        if any(not isinstance(frame, pd.DataFrame) for frame in frames.values()):
            raise TypeError('Each refresh source must return a DataFrame')
        return self.replace_tables(frames)

    def list_native_sources(self):
        """Inspect Excel refresh bindings independently from Python sources."""
        from .native_sources import list_native_sources
        return list_native_sources(self)

    def bind_source(self, table, source):
        """Bind an existing worksheet table for deliberate refresh in Excel.

        The cached model data is unchanged and refresh-on-open is disabled.
        Worksheet column names/order must match the model's stored columns.
        """
        from .native_sources import LinkedTableSource, OleDbSource, bind_linked_table
        if isinstance(source, LinkedTableSource):
            image = bind_linked_table(self, table, source)
        elif isinstance(source, OleDbSource):
            from .ole_db import bind_ole_db
            image = bind_ole_db(self, table, source)
        else:
            raise TypeError('Expected LinkedTableSource or OleDbSource; Python refresh sources use refresh_sources()')
        self._image = image
        return self

    def unbind_source(self, table):
        """Convert a bound table to a snapshot without changing cached data."""
        from .native_sources import store_binding
        from .ole_db import detach_table_binding
        image = detach_table_binding(self._image, self._dimension_ids[table])
        self._image = store_binding(image, self._dimension_ids[table], {'kind': 'snapshot'})
        return self

    def _edit_kpi(self, operation):
        script = next(f for f in self._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
        xml = self._image.read_logical_file(script)
        if self._pending_measures:
            xml = add_measures_to_script(xml, self._pending_measures)
        xml = operation(xml)
        image = AbfImage(self._image.assemble({script.storage_path: xml}))
        candidate = type(self)(self.source, _image=image)
        candidate._pending = dict(self._pending)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        self.__dict__.update(candidate.__dict__)
        return self

    def update_measure(self, table, name, *, expression=None, format_string=None,
                       description=None, display_folder=None):
        """Edit an independent measure; expressions are evaluated by Excel."""
        from .measure_edit import edit_measure
        return self._edit_kpi(lambda xml: edit_measure(xml, table, name,
            expression=expression, format_string=format_string,
            description=description, display_folder=display_folder))

    def remove_measure(self, table, name):
        """Remove an independent measure, rejecting calculation dependencies."""
        from .measure_edit import edit_measure
        self.dependency_graph().assert_removable(('tables', table, 'measures', name))
        from zipfile import ZipFile
        with ZipFile(self.source) as archive:
            if any(n.startswith(('xl/pivotTables/', 'xl/pivotCache/')) for n in archive.namelist()):
                raise UnsupportedModelError('Cannot remove measures from a workbook with PivotTable dependencies')
        return self._edit_kpi(lambda xml: edit_measure(xml, table, name, remove=True))

    def list_kpis(self):
        from .kpis import list_kpis
        script = next(f for f in self._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
        return list_kpis(self._image.read_logical_file(script))

    def rename_measure(self, table, name, new_name):
        """Rename an unambiguous measure and DAX uses atomically.

        KPI, perspective and workbook-report bindings must be removed first.
        Dependent calculated caches are invalidated for Excel to process.
        """
        from .dependencies import DependencyError
        from .validation import ValidationIssue
        from .structural_edit import rename_measure_references
        if not isinstance(new_name, str) or not new_name.strip() or new_name != new_name.strip() or any(ord(c) < 32 for c in new_name):
            raise ValueError('Measure name must be nonempty, trimmed text without control characters')
        graph = self.dependency_graph()
        target = next((obj.path for obj in graph.objects if len(obj.path) == 4
                       and obj.path[2] == 'measures' and obj.path[1].casefold() == table.casefold()
                       and obj.path[3].casefold() == name.casefold()), None)
        if target is None:
            raise KeyError(('tables', table, 'measures', name))
        table, name = target[1], target[3]
        conflicts = [obj.path for obj in graph.objects if len(obj.path) == 4
                     and obj.path[2] in ('measures', 'columns') and obj.path != target
                     and obj.path[3].casefold() == new_name.casefold()]
        if conflicts:
            raise DependencyError('New measure name conflicts with a model field', issues=(ValidationIssue(
                'ambiguous_reference', 'New measure name conflicts with a model field',
                object_path=target, dependencies=tuple(sorted(conflicts))),))
        dependents = graph.dependents(target)
        blockers = {edge.source for edge in dependents if not (
            len(edge.source) == 4 and edge.source[2] in ('measures', 'columns')
            and edge.reason == 'DAX field reference')}
        blockers.update(ref.source for ref in graph.unresolved if ref.name.casefold() == name.casefold()
                        and (ref.table is None or ref.table.casefold() == table.casefold()))
        if blockers:
            raise DependencyError('Measure rename has unsupported or ambiguous dependencies', issues=(ValidationIssue(
                'dependent_objects', 'Measure rename has unsupported or ambiguous dependencies',
                object_path=target, dependencies=tuple(sorted(blockers))),))
        candidate = self._materialized_candidate()
        image = rename_measure_references(candidate._image, table, name, new_name, table_names=candidate.tables)
        if any(len(edge.source) == 4 and edge.source[2] == 'columns' for edge in dependents):
            from .calculated_columns import invalidate_calculated_columns
            image = invalidate_calculated_columns(image)
        result = type(self)(self.source, _image=image)
        self.__dict__.update(result.__dict__)
        return self

    def add_kpi(self, table, measure, *, target_measure=None, target_table=None,
                target_value=None, lower=0.4, upper=0.8, thresholds=None,
                mode='ascending', icon='Traffic Light - Single', threshold_type=None):
        """Add a KPI with a fixed or measure target and explicit status rules."""
        from .kpis import add_kpi
        return self._edit_kpi(lambda xml: add_kpi(xml, table, measure,
            target_measure=target_measure, target_table=target_table, target_value=target_value,
            lower=lower, upper=upper, thresholds=thresholds, mode=mode, icon=icon,
            threshold_type=threshold_type))

    def remove_kpi(self, table, measure):
        """Remove KPI helpers/settings while retaining its base measure."""
        from .kpis import remove_kpi
        self.dependency_graph().assert_removable(('kpis', measure))
        return self._edit_kpi(lambda xml: remove_kpi(xml, table, measure))

    def update_kpi(self, table, measure, *, target_measure=None, target_table=None,
                   target_value=None, lower=0.4, upper=0.8, thresholds=None,
                   mode='ascending', icon='Traffic Light - Single', threshold_type=None):
        """Atomically replace a supported KPI's target and thresholds."""
        from .kpis import add_kpi, remove_kpi
        return self._edit_kpi(lambda xml: add_kpi(remove_kpi(xml, table, measure), table, measure,
            target_measure=target_measure, target_table=target_table, target_value=target_value,
            lower=lower, upper=upper, thresholds=thresholds, mode=mode, icon=icon,
            threshold_type=threshold_type))

    def add_measure(
        self,
        table: str,
        name: str,
        expression: str,
        *,
        format_string: str = "",
        display_folder: str = "",
        description: str = "",
    ) -> "ExcelModelTemplate":
        if table not in {item.table for item in self._columns}:
            raise KeyError(f"Model table {table} does not exist")
        if not name.strip() or not expression.strip():
            raise ValueError("Measure name and expression are required")
        from .formula_statements import normalize_expression
        expression = normalize_expression(expression)
        self._pending_measures.append(
            Measure(table, name, expression, format_string, display_folder, description)
        )
        return self

    def set_sort_column(self, table: str, column: str, by: str | None) -> 'ExcelModelTemplate':
        """Order a column by another column; pass None to restore its own order."""
        from .sorting import stored_sorts, set_sort_metadata, rebuild_sort, validate_sort_values
        calculated = {c.name: c for c in self.list_calculated_columns()
                      if c.table_id == self._dimension_ids.get(table)}
        for name in [column] + ([by] if by is not None else []):
            if (table, name) not in self._encodings and name not in calculated:
                raise KeyError(f'Model column {table}.{name} does not exist')
        if by == column:
            by = None
        if by is not None:
            source_path = ('tables', table, 'columns', column)
            key_path = ('tables', table, 'columns', by)
            if any(edge.target == source_path for edge in
                   self.dependency_graph().dependencies(key_path, transitive=True)):
                raise ValueError('Sort key depends on the column being sorted; this would create a circular dependency')
        with self._open_reader() as reader:
            frame = reader.get_table(table)
        values = self._pending.get((table, column), frame[column].tolist())
        keys = self._pending.get((table, by), frame[by].tolist()) if by else values
        if column not in calculated and by not in calculated:
            validate_sort_values(values, keys, table, column)
        graph = {c: b for t, c, b in stored_sorts(self) if t == table and c != column}
        target, visited = by, {column}
        while target is not None:
            if target in visited:
                raise ValueError('Sort columns cannot form a circular dependency')
            visited.add(target)
            target = graph.get(target)
        image = set_sort_metadata(self, table, column, by)
        if column in calculated or by in calculated:
            from .calculated_columns import invalidate_calculated_columns
            image = invalidate_calculated_columns(image)
        candidate = type(self)(self.source, _image=image)
        if column not in calculated and by not in calculated and (by is None or not any((table, n) in self._pending for n in (column, by))):
            image = rebuild_sort(candidate, table, column, by)
            from .hierarchy_storage import stored_hierarchies, rebuild_hierarchy_storage
            for hierarchy_table, hierarchy_name, levels in stored_hierarchies(self):
                if hierarchy_table == table and column in levels:
                    current = type(self)(self.source, _image=image)
                    image = rebuild_hierarchy_storage(current, table, hierarchy_name, levels)
        candidate = type(self)(self.source, _image=image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        self.__dict__.update(candidate.__dict__)
        return self

    def add_hierarchy(self, table: str, name: str, levels) -> 'ExcelModelTemplate':
        """Create an ordered hierarchy from existing columns in one table."""
        from .hierarchy_storage import add_hierarchy_storage, stored_hierarchies
        levels = list(levels)
        if not isinstance(name, str) or not name.strip() or any(c in name for c in '$/\\:*?"<>|'):
            raise ValueError('Hierarchy name must be nonempty and contain no storage path delimiters')
        if not levels or len(set(levels)) != len(levels):
            raise ValueError('Hierarchy requires distinct ordered columns')
        calculated = {c.name for c in self.list_calculated_columns()
                      if c.table_id == self._dimension_ids.get(table)}
        for column in levels:
            if (table, column) not in self._encodings and column not in calculated:
                raise KeyError(f'Model column {table}.{column} does not exist')
        if any(t == table and n.casefold() == name.casefold() for t, n, _ in stored_hierarchies(self)):
            raise ValueError(f'Hierarchy {table}.{name} already exists')
        image = add_hierarchy_storage(self, table, name, levels)
        candidate = type(self)(self.source, _image=image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        self.__dict__.update(candidate.__dict__)
        return self

    def remove_hierarchy(self, table: str, name: str) -> 'ExcelModelTemplate':
        """Remove a hierarchy while retaining its source columns."""
        from .hierarchy_storage import remove_hierarchy_storage, stored_hierarchies
        if not any(t == table and n == name for t, n, _ in stored_hierarchies(self)):
            raise KeyError(f'Hierarchy {table}.{name} does not exist')
        self.dependency_graph().assert_removable(('tables', table, 'hierarchies', name))
        from zipfile import ZipFile
        with ZipFile(self.source) as archive:
            if any(path.startswith(('xl/pivotTables/', 'xl/pivotCache/', 'xl/slicers/'))
                   for path in archive.namelist()):
                raise UnsupportedModelError('Hierarchy removal from workbooks with PivotTables or slicers requires presentation rebinding')
        image = remove_hierarchy_storage(self, table, name)
        candidate = type(self)(self.source, _image=image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        self.__dict__.update(candidate.__dict__)
        return self

    def configure_hierarchy(self, table: str, name: str, *, hidden: bool | None = None,
                            description: str | None = None, level_names: dict | None = None) -> 'ExcelModelTemplate':
        """Set hierarchy visibility, description and source-column level captions."""
        from .hierarchy_storage import configure_hierarchy_storage
        image = configure_hierarchy_storage(self, table, name, hidden=hidden,
                                            description=description, level_names=level_names)
        candidate = type(self)(self.source, _image=image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        self.__dict__.update(candidate.__dict__)
        return self

    def update_hierarchy(self, table: str, name: str, *, levels=None,
                         new_name: str | None = None) -> 'ExcelModelTemplate':
        """Rename or replace/reorder hierarchy levels atomically.

        Existing PivotTables and slicers require presentation rebinding and are
        rejected before mutation. Metadata for retained source levels is kept.
        """
        from .hierarchy_storage import stored_hierarchies, resolve_hierarchy
        from .schema import _elements, _child, _text
        from xml.dom import minidom
        existing = next((columns for t, n, columns in stored_hierarchies(self)
                         if t == table and n == name), None)
        if existing is None:
            raise KeyError(f'Hierarchy {table}.{name} does not exist')
        hierarchy_id, hierarchy_caption = resolve_hierarchy(self, table, name)
        levels = existing if levels is None else list(levels)
        new_name = hierarchy_caption if new_name is None else new_name
        doc = minidom.parseString(self._image.read_logical_file(self._dimension_paths[table]))
        old = next(h for h in _elements(doc, 'Hierarchy') if _text(_child(h, 'ID')) == hierarchy_id)
        column_captions = {physical: caption for (owner, caption), physical in self._column_ids.items()
                           if owner == table}
        captions = {column_captions.get(_text(_child(level, 'SourceAttributeID')),
                                        _text(_child(level, 'SourceAttributeID'))): _text(_child(level, 'Name'))
                    for level in _elements(old, 'Level')}
        description = _text(_child(old, 'Description'))
        cube = next(f for f in self._image.files if f.logical_path and f.logical_path.endswith('.cub.xml'))
        doc = minidom.parseString(self._image.read_logical_file(cube))
        dimension = next(n for n in _elements(doc, 'Dimension')
                         if any(getattr(c, 'localName', None) == 'DimensionID' and
                                _text(c) == self._dimension_ids[table] for c in n.childNodes))
        old = next(h for h in _elements(dimension, 'Hierarchy') if _text(_child(h, 'HierarchyID')) == hierarchy_id)
        hidden = _text(_child(old, 'Visible')).lower() == 'false'
        candidate = type(self)(self.source, _image=self._image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        candidate.remove_hierarchy(table, name)
        candidate.add_hierarchy(table, new_name, levels)
        candidate.configure_hierarchy(table, new_name, hidden=hidden, description=description,
                                      level_names={c: captions[c] for c in levels if c in captions})
        self.__dict__.update(candidate.__dict__)
        return self

    def configure_table(self, table: str, *, hidden: bool | None = None,
                        description: str | None = None) -> 'ExcelModelTemplate':
        """Set a table's field-list visibility and description."""
        from .schema import configure_table
        image = configure_table(self._image, self._dimension_ids[table], hidden=hidden, description=description)
        candidate = type(self)(self.source, _image=image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        self.__dict__.update(candidate.__dict__)
        return self

    def add_relationship(
        self,
        from_table: str,
        from_column: str,
        to_table: str,
        to_column: str,
        *,
        active: bool = True,
        cardinality: str = "many-to-one",
    ) -> "ExcelModelTemplate":
        from .relationships import RelationshipValidationError, validate_relationship_keys, validate_relationship_collation
        if cardinality not in ('many-to-one', 'one-to-many', 'one-to-one'):
            raise RelationshipValidationError(
                'cardinality must be many-to-one, one-to-many or one-to-one; '
                'Excel does not support direct many-to-many relationships')
        if cardinality == 'one-to-many':
            from_table, from_column, to_table, to_column = to_table, to_column, from_table, from_column
            cardinality = 'many-to-one'
        if not isinstance(active, bool):
            raise TypeError('active must be a boolean')
        for table, column in ((from_table, from_column), (to_table, to_column)):
            if not any(item.table == table and item.name == column for item in self._columns):
                raise KeyError(f"Model column {table}.{column} does not exist")
        if from_table not in self._dimension_paths:
            raise UnsupportedModelError(
                f"Could not resolve relationship metadata for {from_table}"
            )
        from .relationship_storage import stored_relationships, relationship_key_type
        if relationship_key_type(self, from_table, from_column) != relationship_key_type(self, to_table, to_column):
            raise RelationshipValidationError('Relationship columns must have the same data type')
        proposed = Relationship(from_table, from_column, to_table, to_column, active, cardinality)
        existing = [r for _, r in stored_relationships(self)] + self._pending_relationships
        if from_table == to_table:
            raise RelationshipValidationError('Self relationships are not supported')
        for relation in existing:
            if (relation.from_table, relation.from_column, relation.to_table, relation.to_column) == (from_table, from_column, to_table, to_column):
                raise RelationshipValidationError('Relationship already exists')
        if active:
            # Multiple active paths make filter propagation ambiguous.
            graph = {}
            for relation in existing:
                if relation.active:
                    graph.setdefault(relation.from_table, set()).add(relation.to_table)
                    graph.setdefault(relation.to_table, set()).add(relation.from_table)
            frontier, visited = [from_table], set()
            while frontier:
                node = frontier.pop()
                if node == to_table:
                    raise RelationshipValidationError('Relationship creates an ambiguous active filter path')
                if node not in visited:
                    visited.add(node)
                    frontier.extend(graph.get(node, set()) - visited)
        validate_relationship_collation(self, proposed)
        from .relationship_storage import validate_cached_relationship_keys
        validate_cached_relationship_keys(self, proposed)
        self._pending_relationships.append(proposed)
        return self

    def remove_relationship(self, from_table: str, from_column: str,
                            to_table: str, to_column: str) -> 'ExcelModelTemplate':
        """Remove one relationship, retaining the tables and their values."""
        from .relationship_storage import stored_relationships, remove_relationship_storage, synchronize_indirect_bindings
        key = (from_table, from_column, to_table, to_column)
        def matches(relation):
            return (relation.from_table, relation.from_column, relation.to_table, relation.to_column) == key
        pending = [r for r in self._pending_relationships if not matches(r)]
        stored = [(rid, r) for rid, r in stored_relationships(self) if matches(r)]
        if not stored and len(pending) == len(self._pending_relationships):
            raise KeyError('Relationship does not exist')
        image = self._image
        for rid, relation in stored:
            image = remove_relationship_storage(type(self)(self.source, _image=image), rid)
        if stored:
            image = synchronize_indirect_bindings(type(self)(self.source, _image=image))
        candidate = type(self)(self.source, _image=image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = pending
        candidate._resized_tables = dict(self._resized_tables)
        self.__dict__.update(candidate.__dict__)
        return self

    def set_relationship_active(self, from_table: str, from_column: str,
                                to_table: str, to_column: str, *, active: bool) -> 'ExcelModelTemplate':
        """Change activation atomically; retain the original if validation fails."""
        candidate = type(self)(self.source, _image=self._image)
        candidate._pending = dict(self._pending)
        candidate._pending_measures = list(self._pending_measures)
        candidate._pending_relationships = list(self._pending_relationships)
        candidate._resized_tables = dict(self._resized_tables)
        from .relationship_storage import stored_relationships
        relationships = [r for _, r in stored_relationships(candidate)] + candidate._pending_relationships
        original = next((r for r in relationships if
            (r.from_table, r.from_column, r.to_table, r.to_column) ==
            (from_table, from_column, to_table, to_column)), None)
        candidate.remove_relationship(from_table, from_column, to_table, to_column)
        candidate.add_relationship(from_table, from_column, to_table, to_column,
            active=active, cardinality=original.cardinality)
        self.__dict__.update(candidate.__dict__)
        return self

    def validate(self):
        """Return a non-mutating relationship preflight and column support report.

        Does not execute DAX, validate all workbook dependencies, or replace
        save-time validation and native Excel compatibility testing.
        """
        from .validation import validate_model
        return validate_model(self)

    def save(self, destination: str | Path, *, validate: bool = True) -> Path:
        from .preserve import preserve_table
        from .relationships import validate_relationship_keys
        from .relationship_storage import stored_relationships
        from .schema import _elements, _child, _text
        from xml.dom import minidom
        for table, path in self._dimension_paths.items():
            doc = minidom.parseString(self._image.read_logical_file(path))
            dimension = _elements(doc, 'Dimension')[0]
            if _text(_child(dimension, 'Type')) == 'Time':
                key = next(attr for attr in _elements(_child(dimension, 'Attributes'), 'Attribute')
                           if _text(_child(attr, 'Usage')) == 'Key')
                self._validate_date_table_key(table, _text(_child(key, 'Name')) or _text(_child(key, 'ID')))
        for relationship in [r for _, r in stored_relationships(self)] + self._pending_relationships:
            if relationship in self._pending_relationships or any(
                (table, column) in self._pending for table, column in (
                    (relationship.from_table, relationship.from_column),
                    (relationship.to_table, relationship.to_column))):
                from .relationship_storage import validate_cached_relationship_keys
                validate_cached_relationship_keys(self, relationship)

        replacements: dict[str, bytes] = {}
        preserved_tables = set()
        calculated_tables = {column.table_id for column in self.list_calculated_columns()}
        for table in self._resized_tables:
            if self._dimension_ids[table] in calculated_tables:
                continue
            preserved = preserve_table(self, table)
            if preserved is not None:
                replacements.update(preserved)
                preserved_tables.add(table)
        metadata_updates: dict[str, list[tuple[str, list[int], int]]] = {}
        original_tables = {}
        for key, values in self._pending.items():
            if key[0] in preserved_tables:
                continue
            physical_column = self._column_ids[key]
            encoding = self._encodings[key]
            template_idf = self._image.read_logical_file(encoding.storage_path)
            resized = key[0] in self._resized_tables
            if isinstance(encoding, _StringEncoding):
                data_ids, dictionary = encoding.encode(values, allow_resize=resized, compact=True)
                if dictionary:
                    replacements[encoding.dictionary_path] = encode_uncompressed_string_dictionary(
                        self._image.read_logical_file(encoding.dictionary_path), dictionary
                    )
                if tuple(dictionary) != encoding.original_values or (2 in data_ids) != any(
                    bool(s.get('has_nulls')) for s in encoding.segments
                ):
                    from .lookup import string_lookup_replacements
                    replacements.update(string_lookup_replacements(self._image, self._dimension_ids[key[0]],
                        physical_column, dictionary, has_nulls=2 in data_ids))
            elif isinstance(encoding, _NumericDictionaryEncoding):
                data_ids, dictionary = encoding.encode(values, allow_resize=resized, compact=True)
                if encoding.dictionary_path is not None:
                    replacements[encoding.dictionary_path] = encode_numeric_dictionary(
                        self._image.read_logical_file(encoding.dictionary_path), dictionary
                    )
                if tuple(dictionary) != encoding.original_values or (2 in data_ids) != any(
                    bool(s.get('has_nulls')) for s in encoding.segments
                ):
                    from .lookup import dictionary_lookup_replacements
                    replacements.update(dictionary_lookup_replacements(
                        self._image, self._dimension_ids[key[0]], physical_column, dictionary,
                        has_nulls=2 in data_ids, ordered=sorted(range(len(dictionary)), key=dictionary.__getitem__)))
            else:
                data_ids = encoding.encode(values, allow_resize=resized)
                ordinary_ids = [i for v, i in zip(values, data_ids) if not pd.isna(v)]
                rebased = bool(ordinary_ids and (min(ordinary_ids) < 3 or max(ordinary_ids) > 0x7fffffff))
                if rebased:
                    from xml.dom import minidom
                    from .schema import _elements, _child
                    offset = min(ordinary_ids) - 3
                    if max(ordinary_ids) - min(ordinary_ids) > 0x7ffffffc:
                        raise UnsupportedModelError('Numeric range exceeds the value-encoding capacity')
                    data_ids = [2 if pd.isna(v) else i - offset for v, i in zip(values, data_ids)]
                    metadata = minidom.parseString(replacements.get(encoding.metadata_path,
                        self._image.read_logical_file(encoding.metadata_path)))
                    raw = next(n for n in _elements(metadata, 'XMObject')
                        if n.getAttribute('class') == 'XMRawColumn' and n.getAttribute('name') == physical_column)
                    bases = _elements(raw, 'BaseId')
                    if len(bases) != 1:
                        raise UnsupportedModelError('Cannot uniquely resolve the numeric value base')
                    bases[0].firstChild.data = str(encoding.base_id + offset)
                    replacements[encoding.metadata_path] = metadata.documentElement.toxml(encoding='utf-8')
                    resized = True
                if key[0] not in original_tables:
                    from pbixray import PBIXRay
                    with self._open_reader() as reader:
                        original_tables[key[0]] = reader.get_table(key[0])
                existing_ids = (set() if encoding.semantic_type == 'Boolean' else
                    set(encoding.encode(original_tables[key[0]][key[1]].tolist())))
                # HIDX keys also address retained hierarchy structures. A subset
                # rewrite must not renumber them by rebuilding a smaller index.
                if rebased or not set(data_ids).issubset(existing_ids):
                    from .lookup import numeric_lookup_replacements
                    replacements[encoding.hash_index_path] = encode_numeric_hash_index(
                        self._image.read_logical_file(encoding.hash_index_path), data_ids
                    )
                    replacements.update(numeric_lookup_replacements(
                        self._image, self._dimension_ids[key[0]], physical_column, data_ids
                    ))
            if not resized:
                offset = 0
                for segment in encoding.segments:
                    count = int(segment['records'])
                    chunk = data_ids[offset:offset + count]
                    offset += count
                    minimum = 2 if segment.get('has_nulls') else int(segment['min_data_id'])
                    width = int(segment['bit_width'])
                    if ((2 in chunk) != bool(segment.get('has_nulls'))
                            or any(not 0 <= value - minimum < 1 << width for value in chunk)):
                        resized = True
                        break
            if resized:
                min_data_id = 2 if 2 in data_ids else min(data_ids, default=3)
                span = max(data_ids, default=min_data_id) - min_data_id
                bit_width = select_bit_width(span)
                replacements[encoding.storage_path] = encode_pure_bitpacked_ids(
                    data_ids, bit_width=bit_width, min_data_id=min_data_id
                )
                metadata_xml = replacements.get(
                    encoding.metadata_path,
                    self._image.read_logical_file(encoding.metadata_path),
                )
                replacements[encoding.metadata_path] = update_column_segment_metadata(
                    metadata_xml,
                    column=physical_column,
                    data_ids=data_ids,
                    bit_width=bit_width,
                    min_data_id=min_data_id,
                    dictionary_count=(
                        len(dictionary)
                        if isinstance(encoding, (_StringEncoding, _NumericDictionaryEncoding))
                        else None
                    ),
                )
            else:
                if isinstance(encoding, _StringEncoding) and len(encoding.segments) == 1:
                    segment = encoding.segments[0]
                    minimum = 2 if segment.get('has_nulls') else int(segment['min_data_id'])
                    width = int(segment['bit_width'])
                    if any(not 0 <= value - minimum < 1 << width for value in data_ids):
                        minimum = min(data_ids, default=2)
                        width = select_bit_width(max(data_ids, default=minimum) - minimum)
                        replacements[encoding.storage_path] = encode_pure_bitpacked_ids(
                            data_ids, bit_width=width, min_data_id=minimum
                        )
                        replacements[encoding.metadata_path] = update_column_segment_metadata(
                            replacements.get(encoding.metadata_path, self._image.read_logical_file(encoding.metadata_path)),
                            column=physical_column, data_ids=data_ids, bit_width=width,
                            min_data_id=minimum, dictionary_count=len(dictionary),
                        )
                        continue
                if isinstance(encoding, (_StringEncoding, _NumericDictionaryEncoding)):
                    metadata_updates.setdefault(encoding.metadata_path, []).append(
                        (physical_column, data_ids, len(dictionary))
                    )
                replacements[encoding.storage_path] = encode_idf_preserving_layout(
                    template_idf,
                    data_ids=data_ids,
                    segments=encoding.segments,
                )
        for storage_path, updates in metadata_updates.items():
            metadata_xml = replacements.get(storage_path, self._image.read_logical_file(storage_path))
            for column, data_ids, cardinality in updates:
                metadata_xml = update_string_column_metadata(
                    metadata_xml,
                    column=column,
                    data_ids=data_ids,
                    distinct_count=cardinality,
                )
            replacements[storage_path] = metadata_xml
        for table, row_count in self._resized_tables.items():
            if table in preserved_tables:
                continue
            encoding = next(
                value for key, value in self._encodings.items() if key[0] == table
            )
            metadata_xml = replacements.get(
                encoding.metadata_path,
                self._image.read_logical_file(encoding.metadata_path),
            )
            from xml.etree import ElementTree as ET
            from .metadata import IMBI
            engine_names = [node.get('name') for node in ET.fromstring(metadata_xml).iter(f'{{{IMBI}}}XMObject')
                            if node.get('class') == 'XMRawColumn'
                            and node.get('name') in ('RowNumber', '__XL_RowNumber')
                            and (table, node.get('name')) not in self._encodings]
            if len(engine_names) != 1:
                raise UnsupportedModelError('Cannot uniquely resolve engine row-number column')
            for engine_name in engine_names:
                if engine_name:
                    metadata_xml = update_constant_column_row_count(
                        metadata_xml, column=engine_name, row_count=row_count
                    )
                    row_files = [file for file in self._image.files
                                 if file.logical_path and file.logical_path.replace("\\", "/").endswith(
                                     f"/{self._dimension_ids[table]}.{engine_name}.0.idf")]
                    if not row_files:
                        row_files = [file for file in self._image.files if file.logical_path
                                     and file.logical_path.replace("\\", "/").split("/")[-1].endswith(
                                         f".{self._dimension_ids[table]}.{engine_name}.0.idf")]
                    if len(row_files) != 1:
                        raw = next(node for node in ET.fromstring(metadata_xml).iter(f'{{{IMBI}}}XMObject')
                                   if node.get('class') == 'XMRawColumn' and node.get('name') == engine_name)
                        filename = next(node.get('name') for node in raw.iter(f'{{{IMBI}}}XMObject')
                                        if node.get('class') == 'XMRawColumnPartitionDataObject')
                        row_files = [file for file in self._image.files if file.logical_path
                                     and file.logical_path.replace('\\', '/').split('/')[-1] == filename]
                        if len(row_files) != 1:
                            raise UnsupportedModelError("Cannot resolve engine row-number stream")
                    row_file = row_files[0]
                    row_data = bytearray(self._image.read_logical_file(row_file))
                    if len(row_data) < 16 or struct.unpack_from("<I", row_data, 8)[0] != 0xFFFFFFFF:
                        raise UnsupportedModelError("Unsupported engine row-number stream")
                    struct.pack_into("<I", row_data, 12, row_count)
                    replacements[row_file.storage_path] = bytes(row_data)
            replacements[encoding.metadata_path] = metadata_xml
        if self._pending_measures:
            scripts = [
                item for item in self._image.files
                if item.logical_path and item.logical_path.endswith(".scr.xml")
            ]
            if len(scripts) != 1:
                raise UnsupportedModelError(
                    f"Expected one model script, found {len(scripts)}"
                )
            script = scripts[0]
            replacements[script.storage_path] = add_measures_to_script(
                self._image.read_logical_file(script), self._pending_measures
            )
        # Native all-null hash columns have an unnamed dictionary and no payload.
        # A named zero-member dictionary loads offline but Excel rejects the model.
        removals = []
        for key, values in self._pending.items():
            encoding = self._encodings[key]
            if isinstance(encoding, (_StringEncoding, _NumericDictionaryEncoding)) and not any(not pd.isna(v) for v in values):
                from .null_dictionary import remove_empty_dictionary
                remove_empty_dictionary(self, key, encoding, replacements, removals)
        model = self._image.assemble(replacements, removals=removals) if replacements or removals else self._image.clone()
        if self._pending_relationships or self._pending:
            from .relationship_storage import add_relationship_storage, stored_relationships
            image = AbfImage(model)
            for relationship_id, relationship in stored_relationships(self):
                if ((relationship.from_table, relationship.from_column) in self._pending
                        or relationship.to_table in self._resized_tables
                        or (relationship.to_table, relationship.to_column) in self._pending):
                    current = type(self)(self.source, _image=image)
                    image = add_relationship_storage(current, relationship, existing_id=relationship_id)
            for relationship in self._pending_relationships:
                current = type(self)(self.source, _image=image)
                image = add_relationship_storage(current, relationship)
            from .relationship_storage import synchronize_indirect_bindings
            image = synchronize_indirect_bindings(type(self)(self.source, _image=image))
            model = image.data
        from .relationship_storage import synchronize_unknown_members
        from .sorting import stored_sorts, rebuild_sort
        image = synchronize_unknown_members(type(self)(self.source, _image=AbfImage(model)))
        reordered = set()
        calculated_sort_columns = {(table, c.name) for table, identity in self._dimension_ids.items()
                                   for c in self.list_calculated_columns() if c.table_id == identity}
        for table, column, by in stored_sorts(self):
            if (table, column) in calculated_sort_columns or (table, by) in calculated_sort_columns:
                continue  # Excel rebuilds calculated-key order after processing.
            if (table, column) in self._pending or (table, by) in self._pending:
                current = type(self)(self.source, _image=image)
                image = rebuild_sort(current, table, column, by)
                reordered.add((table, column))
        if self._pending:
            from .hierarchy_storage import stored_hierarchies, rebuild_hierarchy_storage
            deferred_sorts = {(table, column) for table, column, by in stored_sorts(self)
                              if (table, by) in calculated_sort_columns}
            for table, name, levels in stored_hierarchies(self):
                if any((table, column) in calculated_sort_columns | deferred_sorts for column in levels):
                    continue
                if any((table, column) in self._pending or (table, column) in reordered for column in levels):
                    current = type(self)(self.source, _image=image)
                    image = rebuild_hierarchy_storage(current, table, name, levels)
        from .empty_storage import empty_table_storage
        for table, count in self._resized_tables.items():
            if count == 0:
                image = empty_table_storage(image, self._dimension_ids[table])
        if calculated_tables:
            from .calculated_columns import resize_calculated_columns
            for table, count in self._resized_tables.items():
                if self._dimension_ids[table] in calculated_tables:
                    image = resize_calculated_columns(image, self._dimension_ids[table], count)
        model = image.data
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        # Validate beside the destination, then commit on the same filesystem.
        descriptor, staged_name = mkstemp(prefix=".excelmodel-", suffix=".xlsx", dir=destination.parent)
        os.close(descriptor)
        output = Path(staged_name)
        try:
            replace_model_part(self.source, output, model)
            if validate:
                self._validate_output(output)
            output.replace(destination)
        finally:
            output.unlink(missing_ok=True)
        return destination

    def to_bytes(self, *, validate: bool = True) -> bytes:
        """Serialize to bytes for web responses without choosing an output path."""
        with TemporaryDirectory(prefix="excelmodel-") as directory:
            output = self.save(Path(directory) / "model.xlsx", validate=validate)
            return output.read_bytes()

    def _validate_output(self, output: Path) -> None:
        import pandas as pd
        from pbixray import PBIXRay

        try:
            with self._open_reader() as before, PBIXRay(output) as after:
                self._normalize_reader_types(after)
                if not before.dax_measures.equals(after.dax_measures):
                    raise ModelValidationError("DAX measures changed during model rewrite")
                before_relationships = {
                    tuple(row.get(name) for name in (
                        "FromTableName", "FromColumnName", "ToTableName", "ToColumnName",
                        "IsActive", "Cardinality",
                    ))
                    for _, row in before.relationships.iterrows()
                }
                after_relationships = {
                    tuple(row.get(name) for name in (
                        "FromTableName", "FromColumnName", "ToTableName", "ToColumnName",
                        "IsActive", "Cardinality",
                    ))
                    for _, row in after.relationships.iterrows()
                }
                if not before_relationships.issubset(after_relationships):
                    raise ModelValidationError("Existing relationships changed during model rewrite")
                for (table, column), expected_values in self._pending.items():
                    actual = after.get_table(table)[column].reset_index(drop=True)
                    encoding = self._encodings[(table, column)]
                    exact_integer = (isinstance(encoding, _NumericDictionaryEncoding) and
                                     encoding.semantic_type == 'BigInt')
                    expected = pd.Series(expected_values, name=column,
                                         dtype=object if exact_integer else None)
                    if isinstance(encoding, (_ValueEncoding, _NumericDictionaryEncoding)) and encoding.semantic_type == 'Currency':
                        # PBIXRay casts its decoded Decimal series to float64.
                        # Compare decimal text to avoid binary float artifacts.
                        actual = actual.map(lambda v: v if pd.isna(v) else Decimal(str(v)))
                        expected = expected.map(lambda v: v if pd.isna(v) else Decimal(str(v)))
                    if len(actual) != len(expected):
                        raise ModelValidationError(
                            f"{table}.{column} decoded to {len(actual)} rows; "
                            f"expected {len(expected)}"
                        )
                    equal = actual.isna() & expected.isna()
                    present = actual.notna() & expected.notna()
                    equal.loc[present] = actual.loc[present].eq(expected.loc[present]).fillna(False)
                    if not bool(equal.all()):
                        position = int(equal[~equal].index[0])
                        raise ModelValidationError(
                            f"{table}.{column} differs at position {position}: "
                            f"decoded {actual.iloc[position]!r}, expected {expected.iloc[position]!r}"
                        )
                if self._pending_measures:
                    image = AbfImage.from_workbook(output)
                    script = next(
                        item for item in image.files
                        if item.logical_path and item.logical_path.endswith(".scr.xml")
                    )
                    actual_measures = {
                        (item.table, item.name, item.expression)
                        for item in list_script_measures(image.read_logical_file(script))
                    }
                    for item in self._pending_measures:
                        if (item.table, item.name, item.expression.rstrip(";")) not in actual_measures:
                            raise ModelValidationError(
                                f"Generated measure was not found: {item.table}[{item.name}]"
                            )
                if self._pending_relationships:
                    actual_relationships = {
                        (
                            row["FromTableName"], row["FromColumnName"],
                            row["ToTableName"], row["ToColumnName"], bool(row["IsActive"]),
                        )
                        for _, row in after.relationships.iterrows()
                    }
                    for item in self._pending_relationships:
                        expected = (
                            item.from_table, item.from_column,
                            item.to_table, item.to_column, item.active,
                        )
                        if expected not in actual_relationships:
                            raise ModelValidationError(
                                f"Generated relationship was not found: {expected}"
                            )
        except ModelValidationError:
            raise
        except Exception as exc:
            raise ModelValidationError(f"Generated workbook could not be decoded: {exc}") from exc


WorkbookModel = ExcelModelTemplate


def replace_dataframe(
    source: str | Path,
    destination: str | Path,
    *,
    table: str,
    dataframe: Any,
    columns: Sequence[str] | None = None,
    validate: bool = True,
) -> Path:
    """Convenience wrapper for one-shot DataFrame replacement."""
    return (
        ExcelModelTemplate(source)
        .replace_dataframe(table, dataframe, columns=columns)
        .save(destination, validate=validate)
    )
