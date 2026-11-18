"""Transactional report authoring over a live, independently editable model."""
from dataclasses import dataclass, replace
from pathlib import Path
from tempfile import TemporaryDirectory, mkstemp
import os
import re

from .errors import ModelValidationError
from .model import WorkbookModel
from .report_filters import PivotMemberFilter, PivotSort, validate_pivot_options


@dataclass(frozen=True)
class PivotTableSpec:
    name: str
    sheet: str
    cell: str = 'A3'
    rows: tuple = ()
    columns: tuple = ()
    filters: tuple = ()
    measures: tuple = ()
    layout: str = 'tabular'
    row_grand_totals: bool = True
    column_grand_totals: bool = True
    member_filters: tuple = ()
    sorts: tuple = ()

    def __post_init__(self):
        for option in ('member_filters', 'sorts'):
            try:
                object.__setattr__(self, option, tuple(getattr(self, option)))
            except TypeError as exc:
                raise ModelValidationError(f'{option} must be a sequence') from exc
        for axis in ('rows', 'columns', 'filters'):
            try:
                values = tuple(tuple(field) if not isinstance(field, str) else field
                               for field in getattr(self, axis))
            except TypeError as exc:
                raise ModelValidationError('Report axes must contain (table, column) pairs') from exc
            object.__setattr__(self, axis, values)
        if isinstance(self.measures, str):
            raise ModelValidationError('Measures must be a sequence of names')
        try:
            object.__setattr__(self, 'measures', tuple(self.measures))
        except TypeError as exc:
            raise ModelValidationError('Measures must be a sequence of names') from exc


@dataclass(frozen=True)
class PivotChartSpec:
    name: str
    pivot_table: str
    sheet: str | None = None
    cell: str = 'G3'
    chart_type: str = 'column'
    width: int = 640
    height: int = 360
    title: str = ''
    style: int = 2
    show_legend: bool = True
    legend_position: str = 'right'
    category_axis_title: str = ''
    value_axis_title: str = ''
    number_format: str = 'General'
    show_data_labels: bool = False
    background_color: str | None = None
    plot_background_color: str | None = None


@dataclass(frozen=True)
class SlicerSpec:
    name: str
    pivot_table: str
    field: tuple
    sheet: str | None = None
    cell: str = 'J3'
    caption: str = ''
    width: int = 192
    height: int = 240
    columns: int = 1
    connected_pivots: tuple = ()

    def __post_init__(self):
        _control_sequences(self)


@dataclass(frozen=True)
class TimelineSpec:
    name: str
    pivot_table: str
    field: tuple
    sheet: str | None = None
    cell: str = 'J20'
    caption: str = ''
    width: int = 480
    height: int = 140
    level: str = 'months'
    connected_pivots: tuple = ()

    def __post_init__(self):
        _control_sequences(self)


def _control_sequences(spec):
    for key in ('field', 'connected_pivots'):
        value = getattr(spec, key)
        if isinstance(value, (str, bytes)):
            raise ModelValidationError(key+' must be a sequence')
        try:
            object.__setattr__(spec, key, tuple(value))
        except TypeError as exc:
            raise ModelValidationError(key+' must be a sequence') from exc


def _xml_text(value, label):
    if not isinstance(value, str) or any(
        not (ord(c) in (9, 10, 13) or 32 <= ord(c) <= 0xD7FF
             or 0xE000 <= ord(c) <= 0xFFFD or 0x10000 <= ord(c) <= 0x10FFFF)
        for c in value
    ):
        raise ModelValidationError(f'{label} must be valid XML text')


def _text(value, label):
    _xml_text(value, label)
    if not value.strip() or any(ord(c) < 32 for c in value):
        raise ModelValidationError(f'{label} must be nonempty text without control characters')


def _sheet(value):
    _text(value, 'Sheet name')
    if len(value) > 31 or re.search(r'[\[\]:*?/\\]', value) or value.startswith("'") or value.endswith("'"):
        raise ModelValidationError('Invalid worksheet name')


def _cell(value):
    match = re.fullmatch(r'([A-Za-z]{1,3})([1-9][0-9]{0,6})', value) if isinstance(value, str) else None
    if not match:
        raise ModelValidationError('Cell must be a single A1 address')
    column = 0
    for char in match[1].upper():
        column = column * 26 + ord(char) - 64
    if column > 16384 or int(match[2]) > 1048576:
        raise ModelValidationError('Cell is outside Excel worksheet bounds')


def _pivot_name(value):
    _text(value, 'PivotTable name')
    if (len(value) > 255 or not (value[0].isalpha() or value[0] == '_')
            or any(not (c.isalnum() or c in '_.') for c in value)
            or value.casefold() in ('r', 'c') or re.fullmatch(r'(?i)R[1-9][0-9]*C[1-9][0-9]*', value)):
        raise ModelValidationError('PivotTable name must be an identifier, not a cell reference')
    try:
        _cell(value)
    except ModelValidationError:
        pass
    else:
        raise ModelValidationError('PivotTable name cannot be a cell reference')


class WorkbookReport:
    """Stage reports without altering the model or executing its calculations.

    The supplied model remains live: its current fields are checked again at
    save time. New reports are written to new sheets by the package writer.
    """

    def __init__(self, model_or_path):
        self.model = model_or_path if isinstance(model_or_path, WorkbookModel) else WorkbookModel(model_or_path)
        self._pivots = ()
        self._charts = ()
        self._slicers = ()
        self._timelines = ()
        self._pivot_updates = {}
        self._chart_updates = {}
        self._control_updates = {}
        self._remove = ()

    def _validate(self, pivots, charts):
        from .measures import list_script_measures
        fields = {(column.table, column.name) for column in self.model.columns}
        measures = list(self.model._pending_measures)
        for item in self.model._image.files:
            if item.logical_path and item.logical_path.endswith('.scr.xml'):
                measures.extend(list_script_measures(self.model._image.read_logical_file(item)))
        measure_names = {measure.name for measure in measures}
        names = set()
        sheets = set()
        for pivot in pivots:
            _pivot_name(pivot.name)
            if pivot.name.casefold() in names:
                raise ModelValidationError('Duplicate report name')
            names.add(pivot.name.casefold())
            _sheet(pivot.sheet)
            _cell(pivot.cell)
            if pivot.sheet.casefold() in sheets:
                raise ModelValidationError('Each PivotTable requires a distinct new sheet')
            sheets.add(pivot.sheet.casefold())
            if pivot.layout not in ('tabular', 'compact', 'outline'):
                raise ModelValidationError('Unsupported PivotTable layout')
            if type(pivot.row_grand_totals) is not bool or type(pivot.column_grand_totals) is not bool:
                raise ModelValidationError('Grand total options must be booleans')
            seen = set()
            if not (pivot.rows or pivot.columns or pivot.filters or pivot.measures):
                raise ModelValidationError('A PivotTable needs at least one field or measure')
            for field in pivot.rows + pivot.columns + pivot.filters:
                if not isinstance(field, tuple) or len(field) != 2 or not all(isinstance(v, str) for v in field):
                    raise ModelValidationError('Report axes must contain (table, column) pairs')
                if field not in fields or field[1].startswith('__XL_'):
                    raise ModelValidationError(f'Unknown report field: {field!r}')
                for value in field:
                    _text(value, 'Report field name')
                if field in seen:
                    raise ModelValidationError('A report field cannot occur on multiple axes or twice on an axis')
                seen.add(field)
            if any(not isinstance(name, str) or name not in measure_names for name in pivot.measures):
                raise ModelValidationError('Unknown report measure')
            for name in pivot.measures:
                _text(name, 'Report measure name')
            if len(set(pivot.measures)) != len(pivot.measures):
                raise ModelValidationError('Duplicate report measure')
            try:
                validate_pivot_options(pivot, self.model)
            except (TypeError, AttributeError) as exc:
                raise ModelValidationError('Malformed member filter or sort option') from exc
        existing = self._existing_specs()
        existing_names = {r.name.casefold() for r in self.inspect_reports() if r.name not in self._remove}
        for updates in (self._pivot_updates, self._chart_updates):
            for old, changes in updates.items():
                if old not in self._remove and 'name' in changes:
                    _text(changes['name'], 'Report name')
                    existing_names.discard(old.casefold())
                    existing_names.add(changes['name'].casefold())
        by_name = {pivot.name: pivot for pivot in existing if isinstance(pivot, PivotTableSpec)}
        by_name.update({pivot.name: pivot for pivot in pivots})
        for chart in charts:
            _text(chart.name, 'Chart name')
            if len(chart.name) > 255:
                raise ModelValidationError('Chart names cannot exceed 255 characters')
            if chart.name.casefold() in names:
                raise ModelValidationError('Duplicate report name')
            if chart.name.casefold() in existing_names:
                raise ModelValidationError('Duplicate existing report name')
            names.add(chart.name.casefold())
            if not isinstance(chart.pivot_table, str) or chart.pivot_table not in by_name:
                raise ModelValidationError('Unknown chart PivotTable')
            target = by_name[chart.pivot_table]
            if any(field not in fields for field in target.rows+target.columns+target.filters) or any(name not in measure_names for name in target.measures):
                raise ModelValidationError('Chart PivotTable refers to a missing model field or measure')
            validate_pivot_options(target,self.model)
            sheet = chart.sheet if chart.sheet is not None else by_name[chart.pivot_table].sheet
            _sheet(sheet)
            _cell(chart.cell)
            if chart.chart_type not in ('column', 'bar', 'line'):
                raise ModelValidationError('Unsupported PivotChart type')
            if any(type(value) is not int or value <= 0 for value in (chart.width, chart.height)):
                raise ModelValidationError('Chart dimensions must be positive integers')
            _xml_text(chart.title, 'Chart title')
            from .report_charts import validate_chart_format
            try:
                validate_chart_format(chart)
            except ValueError as exc:
                raise ModelValidationError(str(exc)) from exc

    def add_pivot_table(self, spec=None, **kwargs):
        if spec is not None and kwargs:
            raise TypeError('Provide a spec or keyword fields, not both')
        spec = PivotTableSpec(**kwargs) if spec is None else spec
        if not isinstance(spec, PivotTableSpec):
            raise TypeError('Expected PivotTableSpec')
        candidate = self._pivots + (spec,)
        staged = self._candidate()
        staged._pivots = candidate
        staged._validate_all()
        self._pivots = candidate
        return self

    def add_pivot_chart(self, spec=None, **kwargs):
        if spec is not None and kwargs:
            raise TypeError('Provide a spec or keyword fields, not both')
        spec = PivotChartSpec(**kwargs) if spec is None else spec
        if not isinstance(spec, PivotChartSpec):
            raise TypeError('Expected PivotChartSpec')
        candidate = self._charts + (spec,)
        staged = self._candidate()
        staged._charts = candidate
        staged._validate_all()
        self._charts = candidate
        return self

    def list_reports(self):
        """Return immutable staged specifications; no report caches are implied."""
        return self._pivots + self._charts + self._slicers + self._timelines

    def inspect_reports(self):
        """Inspect report objects stored in the source workbook."""
        from .report_editing import inspect_reports
        return inspect_reports(self.model.source)

    def inspect_controls(self):
        from .report_editing import inspect_controls
        return inspect_controls(self.model.source)

    def _update_control(self, name, changes, kind):
        attr = '_slicers' if kind == 'slicer' else '_timelines'
        for i, spec in enumerate(getattr(self, attr)):
            if spec.name == name:
                candidate = self._candidate()
                items = getattr(self, attr)
                setattr(candidate, attr, items[:i]+(replace(spec, **changes),)+items[i+1:])
                candidate._validate_all()
                self.__dict__.update(candidate.__dict__)
                return self
        records = {r.name:r for r in self.inspect_controls()}
        if name not in records or records[name].kind != kind or name in self._remove:
            raise ModelValidationError('Unknown existing control: '+name)
        candidate = self._candidate()
        candidate._control_updates[name] = {**candidate._control_updates.get(name,{}),**changes}
        candidate._validate_all()
        candidate._preflight_existing()
        self.__dict__.update(candidate.__dict__)
        return self

    def update_slicer(self, name, /, **changes):
        return self._update_control(name, changes, 'slicer')

    def update_timeline(self, name, /, **changes):
        return self._update_control(name, changes, 'timeline')

    def _existing_specs(self):
        result = []
        for record in self.inspect_reports():
            if record.name in self._remove or record.spec is None:
                continue
            updates = self._pivot_updates if record.kind == 'pivot_table' else self._chart_updates
            try:
                result.append(replace(record.spec, **updates.get(record.name, {})))
            except TypeError as exc:
                raise ModelValidationError('Unsupported existing report edit') from exc
        return tuple(result)

    def _candidate(self):
        from copy import copy
        candidate = copy(self)
        candidate._pivot_updates = {k:dict(v) for k,v in self._pivot_updates.items()}
        candidate._chart_updates = {k:dict(v) for k,v in self._chart_updates.items()}
        candidate._control_updates = {k:dict(v) for k,v in self._control_updates.items()}
        return candidate

    def _validate_all(self):
        self._validate(self._pivots, self._charts)
        controls = self._slicers+self._timelines
        names = [s.name.casefold() for s in self.list_reports()]
        names += [s.name.casefold() for s in self._existing_specs()]
        names += [self._control_updates.get(s.name,{}).get('name',s.name).casefold()
                  for s in self.inspect_controls() if s.name not in self._remove]
        if len(names) != len(set(names)):
            raise ModelValidationError('Duplicate report name')
        for spec in controls:
            self._check_control(spec, '_timelines' if isinstance(spec,TimelineSpec) else '_slicers')

    def _rename_dependents(self, old, new):
        if old == new:
            return
        self._charts = tuple(replace(s,pivot_table=new) if s.pivot_table==old else s for s in self._charts)
        for attr in ('_slicers','_timelines'):
            setattr(self,attr,tuple(replace(s,pivot_table=new if s.pivot_table==old else s.pivot_table,
                connected_pivots=tuple(new if n==old else n for n in s.connected_pivots)) for s in getattr(self,attr)))

    def update_pivot_table(self, name, /, **changes):
        for i, spec in enumerate(self._pivots):
            if spec.name == name:
                candidate = self._candidate()
                candidate._pivots = self._pivots[:i] + (replace(spec, **changes),) + self._pivots[i+1:]
                candidate._rename_dependents(name,changes.get('name',name))
                candidate._validate_all()
                self.__dict__.update(candidate.__dict__)
                return self
        self._stage_edit(name, changes, 'pivot_table')
        return self

    def update_pivot_chart(self, name, /, **changes):
        for i, spec in enumerate(self._charts):
            if spec.name == name:
                updated = self._charts[:i] + (replace(spec, **changes),) + self._charts[i+1:]
                candidate = self._candidate()
                candidate._charts = updated
                candidate._validate_all()
                self._charts = updated
                return self
        self._stage_edit(name, changes, 'pivot_chart')
        return self

    def _stage_edit(self, name, changes, kind):
        records = {r.name:r for r in self.inspect_reports()}
        if name not in records or records[name].kind != kind or name in self._remove:
            raise ModelValidationError('Unknown existing report: '+name)
        candidate = self._candidate()
        target = candidate._pivot_updates if kind == 'pivot_table' else candidate._chart_updates
        target[name] = {**target.get(name, {}), **changes}
        if kind == 'pivot_table':
            old = self._pivot_updates.get(name,{}).get('name',name)
            candidate._rename_dependents(old,changes.get('name',old))
        candidate._validate_all()
        candidate._preflight_existing()
        self.__dict__.update(candidate.__dict__)

    def _preflight_existing(self):
        if self._pivot_updates or self._chart_updates or self._control_updates or self._remove:
            from .report_editing import edit_reports
            with TemporaryDirectory(prefix='excelmodel-report-check-') as directory:
                edit_reports(self.model.source,Path(directory)/'checked.xlsx',pivot_updates=self._pivot_updates,
                    chart_updates=self._chart_updates,control_updates=self._control_updates,remove=self._remove,model=self.model)

    def remove_report(self, name):
        all_staged = self.list_reports()
        if any(s.name == name for s in all_staged):
            if any(getattr(s, 'pivot_table', None) == name or name in getattr(s, 'connected_pivots', ()) for s in all_staged):
                raise ModelValidationError('Remove dependent charts and controls first')
            for attr in ('_pivots','_charts','_slicers','_timelines'):
                setattr(self,attr,tuple(s for s in getattr(self,attr) if s.name != name))
        elif any(r.name == name for r in self.inspect_reports()+self.inspect_controls()):
            candidate = self._candidate()
            effective = self._pivot_updates.get(name,{}).get('name',name)
            if any(getattr(s,'pivot_table',None)==effective or effective in getattr(s,'connected_pivots',()) for s in all_staged):
                raise ModelValidationError('Remove dependent charts and controls first')
            candidate._remove = tuple(dict.fromkeys(self._remove+(name,)))
            candidate._pivot_updates.pop(name,None); candidate._chart_updates.pop(name,None)
            candidate._control_updates.pop(name,None)
            candidate._validate_all()
            candidate._preflight_existing()
            self.__dict__.update(candidate.__dict__)
        else:
            raise ModelValidationError('Unknown report: '+name)
        return self

    def _check_control(self, spec, kind):
        expected = TimelineSpec if kind == '_timelines' else SlicerSpec
        if not isinstance(spec, expected):
            raise TypeError('Expected '+expected.__name__)
        _text(spec.name, 'Control name'); _cell(spec.cell)
        _xml_text(spec.caption, 'Control caption')
        if spec.sheet is not None: _sheet(spec.sheet)
        if len(spec.field) != 2 or not all(isinstance(v,str) for v in spec.field):
            raise ModelValidationError('Control field must be a (table, column) pair')
        fields={(c.table,c.name):c for c in self.model.columns}
        if spec.field not in fields or spec.field[1].startswith('__XL_'):
            raise ModelValidationError('Unknown control field')
        if any(type(v) is not int or v <= 0 for v in (spec.width,spec.height)):
            raise ModelValidationError('Control dimensions must be positive integer pixels')
        if kind == '_slicers' and (type(spec.columns) is not int or spec.columns < 1):
            raise ModelValidationError('Slicer columns must be a positive integer')
        if kind == '_slicers' and any(t in str(fields[spec.field].data_type).lower()
                                      for t in ('float', 'double', 'decimal', 'currency')):
            # Member keys are only emitted for text/integer/boolean/date values;
            # fail here instead of after every other report change is staged.
            raise ModelValidationError('Slicers support text, integer, boolean and date columns, not decimal values')
        if kind == '_timelines':
            if spec.level not in ('years','quarters','months','days'):
                raise ModelValidationError('Invalid timeline level')
            if 'date' not in fields[spec.field].data_type.lower():
                raise ModelValidationError('Timeline requires a date column')
        pivots={s.name for s in self._pivots}
        pivots.update(s.name for s in self._existing_specs() if isinstance(s,PivotTableSpec))
        if not isinstance(spec.pivot_table,str) or any(not isinstance(n,str) for n in spec.connected_pivots):
            raise ModelValidationError('Connected PivotTable names must be strings')
        connections=(spec.pivot_table,)+spec.connected_pivots
        if len(connections)!=len(set(connections)):
            raise ModelValidationError('Duplicate connected PivotTable')
        if any(n not in pivots for n in (spec.pivot_table,)+spec.connected_pivots):
            raise ModelValidationError('Unknown connected PivotTable')

    def _add_control(self, spec, kind):
        expected = TimelineSpec if kind == '_timelines' else SlicerSpec
        if not isinstance(spec,expected):
            raise TypeError('Expected '+expected.__name__)
        candidate=self._candidate()
        setattr(candidate,kind,getattr(self,kind)+(spec,))
        candidate._validate_all()
        self.__dict__.update(candidate.__dict__)
        return self

    def add_slicer(self, spec=None, **kwargs):
        if spec is not None and kwargs: raise TypeError('Provide a spec or keyword fields, not both')
        return self._add_control(SlicerSpec(**kwargs) if spec is None else spec, '_slicers')

    def add_timeline(self, spec=None, **kwargs):
        if spec is not None and kwargs: raise TypeError('Provide a spec or keyword fields, not both')
        return self._add_control(TimelineSpec(**kwargs) if spec is None else spec, '_timelines')

    def save(self, destination, *, validate=True):
        from .report_package import add_reports
        self._validate_all()
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        descriptor, name = mkstemp(prefix='.excelmodel-report-', suffix='.xlsx', dir=destination.parent)
        os.close(descriptor)
        staged = Path(name)
        try:
            with TemporaryDirectory(prefix='.excelmodel-report-source-', dir=destination.parent) as directory:
                source = self.model.save(Path(directory) / 'model.xlsx', validate=validate)
                if self._pivot_updates or self._chart_updates or self._control_updates or self._remove:
                    from .report_editing import edit_reports
                    source = edit_reports(source, Path(directory)/'edited.xlsx',
                        pivot_updates=self._pivot_updates, chart_updates=self._chart_updates,
                        control_updates=self._control_updates,remove=self._remove, model=self.model, workbook_name=destination.name)
                add_reports(source, staged, self._pivots, self._charts, model=self.model,
                            workbook_name=destination.name, slicers=self._slicers, timelines=self._timelines)
            staged.replace(destination)
        finally:
            staged.unlink(missing_ok=True)
        return destination

    def to_bytes(self, *, validate=True):
        with TemporaryDirectory(prefix='excelmodel-report-') as directory:
            return self.save(Path(directory) / 'report.xlsx', validate=validate).read_bytes()
