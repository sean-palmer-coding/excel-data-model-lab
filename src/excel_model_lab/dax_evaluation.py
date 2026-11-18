"""Optional, read-only DAX evaluation over an isolated model snapshot.

This adapter does not calculate or write Excel caches. The external evaluator is
loaded only when explicitly requested, and its failures are never returned as BLANK.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from importlib import import_module, metadata, resources
from hashlib import sha256
from math import isfinite
from numbers import Integral, Real
from pathlib import Path
import re
from time import perf_counter
from typing import Any, Mapping, Protocol

import pandas as pd

PBIX_MCP_VERSION = '0.9.153'
PBIX_MCP_WHEEL_SHA256 = '25688dda8fe1fb98c78e2dcb09bc489821c3c2b4353d62be47670638b4646fd8'
_BACKEND_FILES = {
    '__init__.py': '493fb70130eb6510f613195d043fb6d441f4ea8f455dc4e4b0b40a31cead68c6',
    'errors.py': 'cd43a551138e5c27a8a7c7712adb6a6a6fd203f161370ad26fe6ad1ce479cf55',
    'dax/__init__.py': '6e693304fd69de263f854a5d542c0af4813709685827dc99f8420a1773ad2209',
    'dax/_case_data.py': 'e2273349bb5cdb234f9c4a8e89b9595db8370f2b0ca7d228baec41c5ea14810f',
    'dax/_collation_data.py': '6d263e90bc0d39231c615fb3741711d5663d258bf95fe47e06c9fb6f98ba5ad0',
    'dax/calc_tables.py': '9fef7409abeabe6f89c1b0494f75d268d7cbf236617e3e0c06a015193bb612dc',
    'dax/collation.py': 'ed39526c67e34d59cb24d24f1618dddfb7b323c650565bd1a600f18257440c61',
    'dax/engine.py': 'eb303ce5dee6513136e9d43c41e476ef25a0465d1fd79b77e5549bf25b8e618b',
    'dax/function_catalog.py': 'a0aec692822f4cb41271ed516e4540f8f8efdee489aa994a09fc4c28c10908be',
    'dax/se_scan.py': 'ef65dd70021558bec75d519d839e53f7ad28ce4965da4d5bdcec046c87ce3c21',
}


class DaxEvaluationError(ValueError):
    """The expression could not be evaluated; this is not a BLANK result."""


class DaxUnsupportedError(DaxEvaluationError):
    """Unsupported backend, model semantics, expression or data type."""


class DaxBackendUnavailableError(ImportError):
    """The exact optional evaluation dependency is not installed."""


@dataclass(frozen=True)
class DaxEvaluation:
    value: Any
    elapsed_seconds: float
    backend: str
    backend_version: str


@dataclass(frozen=True)
class DaxModelSnapshot:
    """Adapter input; mappings are owned privately by one evaluator instance."""
    tables: dict
    measures: dict
    measure_tables: dict
    relationships: tuple
    date_tables: dict
    column_types: dict


class DaxBackend(Protocol):
    """Backend implementations must raise on unsupported/error results."""
    name: str
    version: str

    def evaluate(self, snapshot: DaxModelSnapshot, measure: str, filters: dict,
                 *, max_seconds: float) -> DaxEvaluation: ...


def _scalar(value, data_type=''):
    if value is None or value is pd.NA or value is pd.NaT:
        return None
    if isinstance(value, (str, bool)):
        return value
    if pd.api.types.is_bool(value):
        return bool(value)
    if isinstance(value, (datetime, date)):
        stamp = pd.Timestamp(value)
        if stamp.tzinfo is not None or stamp.microsecond or stamp.nanosecond:
            raise DaxUnsupportedError('Dates require timezone-naive whole-second values')
        return stamp.isoformat()
    if isinstance(value, Decimal):
        raise DaxUnsupportedError('Currency/Decimal evaluation is not yet verified by this adapter')
    if isinstance(value, Integral):
        return int(value)
    if isinstance(value, Real):
        value = float(value)
        if pd.isna(value):
            return None
        if not isfinite(value):
            raise DaxUnsupportedError('Infinite numeric source values are unsupported')
        if data_type.lower() in ('int64', 'integer', 'bigint'):
            if not value.is_integer() or abs(value) >= 2**53:
                raise DaxUnsupportedError('Integer readback is not an exact safe integer')
            return int(value)
        return value
    raise DaxUnsupportedError(f'Unsupported source scalar: {type(value).__name__}')


def _snapshot(model):
    from .excel_metadata import ExcelMetadata
    from .measures import list_script_measures
    from .relationship_storage import stored_relationships

    pending = any(getattr(model, key) for key in
                  ('_pending', '_pending_measures', '_pending_relationships', '_resized_tables'))
    for calculated in model.list_calculated_columns():
        if calculated.processing_state != 'Processed' or pending:
            raise DaxUnsupportedError(
                'Calculated columns must be processed in Excel and saved/reopened before evaluation; '
                'staged model changes invalidate their cached values')
    types = {(c.table, c.name): c.data_type for c in model.columns}
    tables = {}
    for name in model.tables:
        frame = model.get_table(name)
        columns = list(frame.columns)
        if any(not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_ ]*', label) for label in [name, *columns]):
            raise DaxUnsupportedError('Verified evaluator identifiers use ASCII letters, digits, spaces and underscores, starting with a letter or underscore')
        typed = []
        for column in columns:
            data_type = types.get((name, column), '')
            if data_type.lower() == 'currency':
                raise DaxUnsupportedError('Currency evaluation is not yet verified by this adapter')
            typed.append([_scalar(v, data_type) for v in frame[column]])
        tables[name] = {'columns': columns, 'rows': [list(row) for row in zip(*typed)]}
    measures, owners = {}, {}
    definitions = []
    for item in model._image.files:
        if item.logical_path and item.logical_path.endswith('.scr.xml'):
            definitions.extend(list_script_measures(model._image.read_logical_file(item)))
    definitions.extend(model._pending_measures)
    folded = set()
    for item in definitions:
        if not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_ ]*', item.name):
            raise DaxUnsupportedError('Verified measure identifiers use ASCII letters, digits, spaces and underscores')
        if item.name.casefold() in folded:
            raise DaxUnsupportedError('Ambiguous case-insensitive measure names')
        folded.add(item.name.casefold())
        measures[item.name] = item.expression.rstrip(';')
        owners[item.name] = item.table
    relationships = []
    for item in [r for _, r in stored_relationships(model)] + list(model._pending_relationships):
        relationships.append({'FromTable': item.from_table, 'FromColumn': item.from_column,
                              'ToTable': item.to_table, 'ToColumn': item.to_column,
                              'IsActive': int(item.active), 'CrossFilteringBehavior': 1,
                              'FromCardinality': 1 if item.cardinality == 'one-to-one' else 2,
                              'ToCardinality': 1})
    excel = ExcelMetadata.from_files((f.logical_path, model._image.read_logical_file(f))
                                    for f in model._image.files if f.logical_path and f.logical_path.endswith('.xml'))
    dates = {}
    for row in excel.date_tables.itertuples(index=False):
        keys = row.date_key_columns
        if len(keys) != 1:
            raise DaxUnsupportedError('Date-table evaluation requires one explicit key')
        match = excel.columns[(excel.columns.table_id == row.id) & (excel.columns.id == keys[0])]
        if len(match) != 1:
            raise DaxUnsupportedError('Unable to resolve the model date-table key')
        dates[row.name] = match.iloc[0]['name']
    # The pinned published engine predates its upstream date-multihop fix.
    # It also has a by-column-name fallback for disconnected date sources.
    # Refuse those shapes rather than substitute a guessed propagation path.
    if len(dates) > 1:
        raise DaxUnsupportedError('Multiple marked date tables are not verified by this adapter')
    if dates:
        date_table = next(iter(dates))
        distances = {date_table: 0}
        changed = True
        while changed:
            changed = False
            for rel in relationships:
                source, target = rel['ToTable'], rel['FromTable']
                if source in distances and target not in distances:
                    distances[target] = distances[source] + 1
                    changed = True
        if any(depth > 1 for depth in distances.values()):
            raise DaxUnsupportedError('The pinned backend does not support multihop date-table propagation')
        fallback = {'order date', 'date', 'orderdate', 'transaction date'}
        for name, data in tables.items():
            if name not in distances and fallback.intersection(c.casefold() for c in data['columns']):
                raise DaxUnsupportedError('Disconnected date-like tables could trigger backend heuristic filtering')
    return DaxModelSnapshot(tables, measures, owners, tuple(relationships), dates, types)


class PbixMcpBackend:
    """Version-pinned optional Python evaluator; no downloads or server startup."""
    name = 'pbix-mcp'
    version = PBIX_MCP_VERSION

    def __init__(self):
        try:
            installed = metadata.version('pbix-mcp')
        except metadata.PackageNotFoundError as exc:
            raise DaxBackendUnavailableError(
                'Install excel-data-model-lab[dax] to enable the optional DAX evaluator') from exc
        if installed != self.version:
            raise DaxBackendUnavailableError(
                f'Expected pbix-mcp=={self.version}; found {installed}. Backend upgrades require conformance verification.')
        package = resources.files('pbix_mcp')
        for path, expected in _BACKEND_FILES.items():
            payload = package.joinpath(*path.split('/')).read_bytes().replace(b'\r\n', b'\n')
            if sha256(payload).hexdigest() != expected:
                raise DaxBackendUnavailableError(
                    f'Unverified pbix-mcp runtime file: {path}; install the published {self.version} wheel')
        self._module = import_module('pbix_mcp.dax.engine')

    def evaluate(self, snapshot, measure, filters, *, max_seconds):
        from .dax_validation import validate_measure
        started = perf_counter()
        try:
            result_type = validate_measure(snapshot, measure)
        except (IndexError, RecursionError) as exc:
            raise DaxUnsupportedError('Incomplete or excessively nested DAX expression') from exc
        date_table = next(iter(snapshot.date_tables), '__excel_model_lab_no_marked_date_table')
        while date_table not in snapshot.date_tables and date_table in snapshot.tables:
            date_table += '_'
        context = self._module.DAXContext(snapshot.tables, snapshot.measures,
                                          date_table=date_table,
                                          date_column=snapshot.date_tables.get(date_table),
                                          filter_context=filters,
                                          relationships=list(snapshot.relationships))
        context.measure_tables = dict(snapshot.measure_tables)
        context.model_columns = {name: set(data['columns']) for name, data in snapshot.tables.items()}
        context.date_tables = dict(snapshot.date_tables)
        engine = self._module.DAXEngine()
        engine._max_eval_seconds = max_seconds
        try:
            value = engine.evaluate_measure(measure, context)
        except Exception as exc:
            raise DaxEvaluationError(f'{measure}: backend evaluation failed: {exc}') from exc
        if engine.timed_out:
            raise DaxEvaluationError(f'{measure}: evaluation exceeded the backend time budget')
        if engine.unsupported_functions:
            raise DaxUnsupportedError(f'{measure}: unsupported DAX: ' + ', '.join(sorted(engine.unsupported_functions)))
        if engine.eval_errors:
            raise DaxEvaluationError(f'{measure}: ' + '; '.join(str(v) for v in engine.eval_errors.values()))
        if isinstance(value, (dict, list, tuple, set)):
            raise DaxUnsupportedError('Only scalar measure results are exposed by this adapter')
        if isinstance(value, float) and not isfinite(value):
            raise DaxEvaluationError('The backend returned a non-finite result')
        if isinstance(value, Integral) and not -(2**63) <= value < 2**63:
            raise DaxUnsupportedError('Integer evaluation overflow is outside the verified signed Int64 range')
        if value is not None and result_type == 'datetime':
            try:
                value = pd.Timestamp(value).to_pydatetime()
            except (TypeError, ValueError) as exc:
                raise DaxEvaluationError('Invalid date result returned by the backend') from exc
        return DaxEvaluation(value, perf_counter() - started, self.name, self.version)


class DaxEvaluator:
    """Read-only scalar evaluation over model data captured at construction.

    Recreate this object after changing the source model. No result is written
    back to the workbook or advertised as an Excel-calculated cache.
    """
    def __init__(self, model, *, backend: DaxBackend | None = None, max_seconds=30.0):
        from .model import WorkbookModel
        if isinstance(max_seconds, bool) or not isinstance(max_seconds, Real) or not isfinite(max_seconds) or max_seconds <= 0:
            raise ValueError('max_seconds must be finite and positive')
        if isinstance(model, (str, Path)):
            model = WorkbookModel(model)
        if not isinstance(model, WorkbookModel):
            raise TypeError('Expected a WorkbookModel or workbook path')
        self._snapshot = _snapshot(model)
        self._backend = backend if backend is not None else PbixMcpBackend()
        self.max_seconds = float(max_seconds)

    @property
    def measures(self):
        return tuple(self._snapshot.measures)

    def _filters(self, filters):
        result = {}
        if filters is None:
            return result
        if not isinstance(filters, Mapping):
            raise TypeError('filters must map (table, column) pairs to selected value sequences')
        for field, values in filters.items():
            if not isinstance(field, tuple) or len(field) != 2:
                raise ValueError('Filter keys must be (table, column) pairs')
            table, column = field
            if table not in self._snapshot.tables or column not in self._snapshot.tables[table]['columns']:
                raise ValueError(f'Unknown filter field: {field!r}')
            if isinstance(values, (str, bytes)) or not isinstance(values, (list, tuple, set, frozenset)):
                raise ValueError('Selected filter values must be a list, tuple or set')
            result[f'{table}.{column}'] = [_scalar(v) for v in values]
        return result

    def evaluate(self, measure, *, filters=None):
        names = {name.casefold(): name for name in self.measures}
        if not isinstance(measure, str) or measure.casefold() not in names:
            raise ValueError(f'Unknown measure: {measure!r}')
        return self._backend.evaluate(self._snapshot, names[measure.casefold()], self._filters(filters),
                                      max_seconds=self.max_seconds)

    def evaluate_expression(self, expression, *, table, filters=None):
        if not isinstance(expression, str) or not expression.strip():
            raise ValueError('A nonempty scalar DAX expression is required')
        if table not in self._snapshot.tables:
            raise ValueError(f'Unknown expression home table: {table!r}')
        name = '__excel_model_lab_expression'
        while name.casefold() in {n.casefold() for n in self.measures}:
            name += '_'
        snapshot = DaxModelSnapshot(self._snapshot.tables,
            {**self._snapshot.measures, name: expression},
            {**self._snapshot.measure_tables, name: table},
            self._snapshot.relationships, self._snapshot.date_tables, self._snapshot.column_types)
        return self._backend.evaluate(snapshot, name, self._filters(filters), max_seconds=self.max_seconds)
