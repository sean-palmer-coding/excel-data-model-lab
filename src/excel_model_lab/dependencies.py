"""Inspectable model dependencies, without evaluating or fully parsing DAX.

Paths use public names and explicit object kinds. A graph is an immutable snapshot
of a model revision; physical table identities are recorded separately so clients
can correlate objects after display-name changes. Unknown DAX names are reported,
never guessed to be safe for structural edits.
"""
from dataclasses import dataclass
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from .abf import UnsupportedModelError
from .errors import ModelValidationError
from .formula_statements import formula_references


class DependencyError(ModelValidationError, UnsupportedModelError):
    """A destructive edit would invalidate known or unresolved dependencies."""


@dataclass(frozen=True)
class ModelObject:
    path: tuple[str, ...]
    identity: tuple[str, ...]


@dataclass(frozen=True)
class Dependency:
    source: tuple[str, ...]
    target: tuple[str, ...]
    reason: str


@dataclass(frozen=True)
class UnresolvedReference:
    source: tuple[str, ...]
    table: str | None
    name: str
    reason: str


@dataclass(frozen=True)
class DependencyGraph:
    objects: tuple[ModelObject, ...]
    edges: tuple[Dependency, ...]
    unresolved: tuple[UnresolvedReference, ...] = ()

    def dependencies(self, path, *, transitive=False):
        return self._walk(tuple(path), reverse=False, transitive=transitive)

    def dependents(self, path, *, transitive=False):
        return self._walk(tuple(path), reverse=True, transitive=transitive)

    def _walk(self, path, *, reverse, transitive):
        found, frontier, visited = set(), [path], {path}
        while frontier:
            current = frontier.pop()
            for edge in self.edges:
                origin, destination = (edge.target, edge.source) if reverse else (edge.source, edge.target)
                if origin != current:
                    continue
                found.add(edge)
                if transitive and destination not in visited:
                    visited.add(destination)
                    frontier.append(destination)
        return tuple(sorted(found, key=lambda e: (e.source, e.target, e.reason)))

    def assert_removable(self, path):
        """Reject dependencies with structured issue paths; never mutate state."""
        from .validation import ValidationIssue
        path = tuple(path)
        if path not in {obj.path for obj in self.objects}:
            raise KeyError(path)
        targets = {path}
        if len(path) == 2 and path[0] == 'tables':
            targets.update(obj.path for obj in self.objects
                           if obj.path[:2] == path and len(obj.path) == 4 and obj.path[2] == 'columns')
        blockers = {edge.source for edge in self.edges
                    if edge.target in targets and edge.source not in targets}
        # A name cannot be safely removed when an unresolved field could bind to
        # it. Unrelated unresolved fields do not prevent independent operations.
        for ref in self.unresolved:
            if ref.source in targets:
                continue
            if (len(path) == 2 and (ref.table is None or ref.table.casefold() == path[1].casefold())) or (
                    len(path) == 4 and ref.name.casefold() == path[3].casefold()
                    and (ref.table is None or ref.table.casefold() == path[1].casefold())):
                blockers.add(ref.source)
        if blockers:
            names = ', '.join('/'.join(source) for source in sorted(blockers))
            issue = ValidationIssue('dependent_objects', f'Cannot remove {"/".join(path)}; a calculation or model object depends on it: {names}',
                                    object_path=path, dependencies=tuple(sorted(blockers)))
            raise DependencyError(issue.message, issues=(issue,))


def report_dependency_parts(archive):
    """Conservative report bindings, including worksheet/named CUBE formulas."""
    import re
    from .formula_statements import protected_spans
    result = []
    pattern = re.compile(r'(?i)\bCUBE(?:VALUE|MEMBER|MEMBERPROPERTY|SET|SETCOUNT|RANKEDMEMBER|KPIMEMBER)\s*\(')
    for path in archive.namelist():
        if not path.endswith('.xml'):
            continue
        if path.startswith(('xl/pivotTables/', 'xl/pivotCache/', 'xl/slicers/',
                            'xl/slicerCaches/', 'xl/timelines/', 'xl/timelineCaches/')):
            result.append(path)
        elif path.startswith('xl/worksheets/') or path == 'xl/workbook.xml':
            with archive.open(path) as stream:
                for _, node in ET.iterparse(stream, events=('end',)):
                    if node.tag.split('}')[-1] in ('f', 'definedName') and node.text:
                        spans = protected_spans(node.text)
                        if any(not any(start <= match.start() < end for start, end in spans)
                               for match in pattern.finditer(node.text)):
                            result.append(path)
                            break
                    node.clear()
    return tuple(result)


def build_dependency_graph(model):
    """Inspect stored metadata and staged measures/relationships in one snapshot."""
    from .measures import list_script_measures
    from .relationship_storage import stored_relationships
    from .hierarchy_storage import stored_hierarchies
    from .sorting import stored_sorts
    ns = '{http://schemas.microsoft.com/analysisservices/2003/engine}'
    objects, edges, unresolved = {}, set(), []
    tables = {name.casefold(): name for name in model._dimension_ids}
    physical = {identity: name for name, identity in model._dimension_ids.items()}
    columns, measures, expressions = {}, {}, []
    physical_columns, physical_hierarchies = {}, {}

    def add(path, identity=None):
        objects[path] = ModelObject(path, identity or path)
        return path

    def edge(source, target, reason):
        if target in objects:
            edges.add(Dependency(source, target, reason))

    for table, identity in model._dimension_ids.items():
        add(('tables', table), ('tables', identity))
        root = ET.fromstring(model._image.read_logical_file(model._dimension_paths[table]))
        dimension = root.find('.//' + ns + 'ObjectDefinition/' + ns + 'Dimension')
        if dimension is None:
            dimension = root.find('.//' + ns + 'Dimension')
        if dimension is None:
            raise UnsupportedModelError('Cannot inspect dimension dependencies')
        for attr in dimension.findall(ns + 'Attributes/' + ns + 'Attribute'):
            column = attr.findtext(ns + 'Name') or attr.findtext(ns + 'ID')
            column_id = attr.findtext(ns + 'ID')
            if not column or attr.findtext(ns + 'Type') == 'RowNumber' or column_id.startswith('RowNumber-'):
                continue
            path = add(('tables', table, 'columns', column), ('tables', identity, 'columns', column_id))
            columns[(table.casefold(), column.casefold())] = path
            physical_columns[(identity, column_id)] = path
            edge(path, ('tables', table), 'contained column')
            for expression in attr.findall('.//' + ns + 'Expression'):
                if expression.text:
                    expressions.append((path, table, expression.text))
            if dimension.findtext(ns + 'Type') == 'Time' and attr.findtext(ns + 'Usage') == 'Key':
                marker = add(('tables', table, 'date_table'))
                edge(marker, path, 'date table key')
        for hierarchy in dimension.findall(ns + 'Hierarchies/' + ns + 'Hierarchy'):
            name = hierarchy.findtext(ns + 'Name') or hierarchy.findtext(ns + 'ID')
            hierarchy_id = hierarchy.findtext(ns + 'ID')
            path = add(('tables', table, 'hierarchies', name),
                       ('tables', identity, 'hierarchies', hierarchy_id))
            physical_hierarchies[(identity, hierarchy_id)] = path

    scripts = [model._image.read_logical_file(f) for f in model._image.files
               if f.logical_path and f.logical_path.endswith('.scr.xml')]
    for measure in [m for xml in scripts for m in list_script_measures(xml)] + list(model._pending_measures):
        path = add(('tables', measure.table, 'measures', measure.name))
        measures.setdefault(measure.name.casefold(), []).append(path)
        edge(path, ('tables', measure.table), 'measure home table')
        expressions.append((path, measure.table, measure.expression))

    for source, home, expression in expressions:
        for ref in formula_references(expression, table_names=tables):
            if ref.kind == 'table':
                canonical = tables.get(ref.name.casefold())
                if canonical:
                    edge(source, ('tables', canonical), 'DAX table reference')
                else:
                    unresolved.append(UnresolvedReference(source, ref.name, ref.name, 'unknown table'))
                continue
            matching_measures = measures.get(ref.name.casefold(), [])
            if ref.table is not None:
                candidate = columns.get((ref.table.casefold(), ref.name.casefold()))
                candidates = ([candidate] if candidate else []) + [p for p in matching_measures if p[1].casefold() == ref.table.casefold()]
            else:
                candidate = columns.get((home.casefold(), ref.name.casefold()))
                candidates = matching_measures + ([candidate] if candidate else [])
            candidates = set(candidates)
            if candidates:
                for target in candidates:
                    edge(source, target, 'DAX field reference' if len(candidates) == 1 else 'ambiguous DAX field reference')
            else:
                unresolved.append(UnresolvedReference(source, ref.table, ref.name, 'unresolved field or virtual column'))

    for identifier, relation in stored_relationships(model):
        path = add(('relationships', str(identifier)))
        for table, column in ((relation.from_table, relation.from_column), (relation.to_table, relation.to_column)):
            edge(path, columns.get((table.casefold(), column.casefold())), 'relationship key')
    for index, relation in enumerate(model._pending_relationships):
        path = add(('relationships', 'pending', str(index)))
        for table, column in ((relation.from_table, relation.from_column), (relation.to_table, relation.to_column)):
            edge(path, columns.get((table.casefold(), column.casefold())), 'relationship key')
    for table, name, levels in stored_hierarchies(model):
        path = ('tables', table, 'hierarchies', name)
        if path not in objects:
            add(path)
        edge(path, ('tables', table), 'hierarchy home table')
        for level in levels:
            edge(path, columns.get((table.casefold(), level.casefold())), 'hierarchy level')
    for table, column, by in stored_sorts(model):
        source = columns.get((table.casefold(), column.casefold()))
        target = columns.get((table.casefold(), by.casefold()))
        if source and target:
            edge(source, target, 'sort by column')

    # KPI targets are metadata references even when no visible DAX field refers
    # to them. Do not rely on helper measure naming conventions.
    for xml in scripts:
        root = ET.fromstring(xml)
        for prop in root.findall('.//' + ns + 'CalculationProperty'):
            annotations = {a.findtext(ns + 'Name'): a.findtext(ns + 'Value', '')
                           for a in prop.findall(ns + 'Annotations/' + ns + 'Annotation')}
            if annotations.get('Type') != 'Kpi':
                continue
            reference = prop.findtext(ns + 'CalculationReference', '')
            name = reference[1:-1].replace(']]', ']')
            path = add(('kpis', name))
            for measure_name in (name, annotations.get('GoalMeasure')):
                if measure_name:
                    for target in measures.get(measure_name.casefold(), []):
                        edge(path, target, 'KPI measure')

    for file in model._image.files:
        if not file.logical_path or not file.logical_path.endswith('.persp.xml'):
            continue
        root = ET.fromstring(model._image.read_logical_file(file))
        perspective = root.find('.//' + ns + 'Perspective')
        if perspective is None:
            continue
        perspective_id = perspective.findtext(ns + 'ID')
        path = add(('perspectives', perspective.findtext(ns + 'Name') or perspective_id),
                   ('perspectives', perspective_id))
        for dimension in perspective.findall(ns + 'Dimensions/' + ns + 'Dimension'):
            table_id = dimension.findtext(ns + 'CubeDimensionID')
            table = physical.get(table_id)
            if table:
                edge(path, ('tables', table), 'perspective table')
                for attribute in dimension.findall('.//' + ns + 'AttributeID'):
                    edge(path, physical_columns.get((table_id, attribute.text)), 'perspective column')
                for hierarchy in dimension.findall('.//' + ns + 'HierarchyID'):
                    edge(path, physical_hierarchies.get((table_id, hierarchy.text)), 'perspective hierarchy')
        for calculation in perspective.findall('.//' + ns + 'Calculation/' + ns + 'Name'):
            for ref in formula_references(calculation.text or ''):
                for target in measures.get(ref.name.casefold(), []):
                    edge(path, target, 'perspective measure')

    # Excel stores perspective hierarchy selections separately in cube
    # annotations rather than the perspective's Hierarchies element.
    for file in model._image.files:
        if not file.logical_path or not file.logical_path.endswith('.cub.xml'):
            continue
        root = ET.fromstring(model._image.read_logical_file(file))
        for annotation in root.findall('.//' + ns + 'Annotation'):
            if annotation.findtext(ns + 'Name') != 'PerspectiveEmptyHierarchyNames':
                continue
            payload = annotation.findtext(ns + 'Value', '')
            try:
                selections = ET.fromstring(payload)
            except ET.ParseError as exc:
                raise UnsupportedModelError('Cannot inspect perspective hierarchy bindings') from exc
            for perspective in selections.findall('Perspective'):
                source = ('perspectives', perspective.get('Name'))
                for selection in perspective.findall('Tables/Table'):
                    table = tables.get(selection.get('Name', '').casefold())
                    names = selection.get('EmptyHierarchyNames', '').split(',')
                    for name in filter(None, names):
                        target = ('tables', table, 'hierarchies', name)
                        if target in objects:
                            edge(source, target, 'perspective hierarchy')
                        elif table:
                            # Unresolved native selections remain restrictive.
                            for path in tuple(objects):
                                if path[:3] == ('tables', table, 'hierarchies'):
                                    edge(source, path, 'unresolved perspective hierarchy selection')

    # Report bindings are intentionally conservative until their expressions can
    # be rebound. Expose concrete package parts instead of hiding the limitation.
    from .native_sources import list_native_sources
    calculated_paths = {source for source, _, _ in expressions if len(source) == 4 and source[2] == 'columns'}
    for binding in list_native_sources(model):
        if binding.kind != 'snapshot':
            path = add(('sources', binding.table_id))
            edge(path, ('tables', binding.table), 'Excel refresh source binding')
            for column in tuple(objects):
                if column[:3] == ('tables', binding.table, 'columns') and column not in calculated_paths:
                    edge(path, column, 'Excel refresh source schema')
    with ZipFile(model.source) as archive:
        report_parts = report_dependency_parts(archive)
    model_objects = tuple(objects)
    for part in report_parts:
        path = add(('workbook_parts', part))
        for target in model_objects:
            edge(path, target, 'unresolved workbook report binding')
    return DependencyGraph(tuple(sorted(objects.values(), key=lambda obj: obj.path)),
                           tuple(sorted(edges, key=lambda e: (e.source, e.target, e.reason))),
                           tuple(unresolved))
