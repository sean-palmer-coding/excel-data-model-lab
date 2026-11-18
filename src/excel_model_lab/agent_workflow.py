"""Bounded agent operations and recoverable desktop/file handoffs.

This module is not an agent runtime and never evaluates supplied code. Desktop
operations require an exclusive interaction window: Excel does not expose an
atomic compare-and-swap operation for unsaved user changes.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import shutil
from itertools import islice
from tempfile import mkdtemp
from typing import Any, Callable, Iterable
from uuid import uuid4
from zipfile import ZipFile, BadZipFile


class WorkflowConflictError(RuntimeError):
    """The workbook no longer matches the reviewed plan, or is already leased."""


class WorkflowExecutionError(RuntimeError):
    """A workflow failed; checkpoint/recovery paths identify preserved files."""

    def __init__(self, message, *, checkpoint=None, completed=(), stage=None,
                 published=False, recovery_paths=(), recovery_error=None):
        super().__init__(message)
        self.checkpoint = checkpoint
        self.completed = tuple(completed)
        self.stage = stage
        self.published = published
        self.recovery_paths = tuple(Path(p) for p in recovery_paths)
        self.recovery_error = recovery_error


def _hash(path):
    path = Path(path)
    before = path.stat()
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    after = path.stat()
    if (before.st_size, before.st_mtime_ns, before.st_ino) != (after.st_size, after.st_mtime_ns, after.st_ino):
        raise WorkflowConflictError('Workbook changed while reading its fingerprint')
    return digest.hexdigest()


def _json(value):
    return json.dumps(value, ensure_ascii=False, allow_nan=False, sort_keys=True,
                      separators=(',', ':'))


@dataclass(frozen=True)
class AgentOperation:
    """A JSON-serializable operation; ``arguments`` returns a defensive copy.

    Supported names: read_range, write_range, write_formula, calculate,
    update_pivots, save. Refresh and arbitrary attribute access are excluded.
    """
    name: str
    payload: str = '{}'

    @classmethod
    def create(cls, name: str, **arguments):
        operation = cls(name, _json(arguments))
        operation._validate()
        return operation

    @property
    def arguments(self):
        return json.loads(self.payload)

    def _validate(self, max_cells=100000):
        required = {
            'read_range': {'sheet', 'address'},
            'write_range': {'sheet', 'address', 'values'},
            'write_formula': {'sheet', 'address', 'formula'},
            'calculate': set(), 'update_pivots': set(), 'save': set(),
        }
        if self.name not in required:
            raise ValueError('Unsupported agent operation: ' + str(self.name))
        if not isinstance(self.payload, str) or len(self.payload) > 10_000_000:
            raise ValueError('Operation payload exceeds the 10 MB character limit')
        values = self.arguments
        if not isinstance(values, dict) or set(values) != required[self.name]:
            raise ValueError('Unexpected or missing operation arguments')
        if _json(values) != self.payload:
            raise ValueError('Operation payload must use canonical JSON; use AgentOperation.create')
        for key in ('sheet', 'address', 'formula'):
            if key in values and (not isinstance(values[key], str) or not values[key]):
                raise ValueError(key + ' must be a nonempty string')
        if 'address' in values:
            from .desktop_session import _bounds
            rows, columns = _bounds(values['address'], max_cells)
            if self.name == 'write_range':
                matrix = values['values']
                if not isinstance(matrix, list) or len(matrix) != rows or any(
                        not isinstance(row, list) or len(row) != columns for row in matrix):
                    raise ValueError('values must be an exact two-dimensional range matrix')
                for row in matrix:
                    for value in row:
                        if value is not None and not isinstance(value, (str, int, float, bool)):
                            raise ValueError('Cell values must be JSON scalars')
                        if isinstance(value, str) and len(value.encode('utf-16-le')) // 2 > 32767:
                            raise ValueError('Cell text exceeds Excel limits')
            if self.name == 'write_formula':
                if (rows, columns) != (1, 1) or not values['formula'].startswith('='):
                    raise ValueError('write_formula requires one cell and a formula beginning with =')
                if len(values['formula'].encode('utf-16-le')) // 2 > 8192:
                    raise ValueError('Formula exceeds Excel limits')
        return self

    def to_dict(self):
        return {'name': self.name, 'arguments': self.arguments}


@dataclass(frozen=True)
class WorkflowPlan:
    id: str
    path: Path
    pid: int
    disk_sha256: str
    operations: tuple[AgentOperation, ...]
    description: str
    kind: str = 'operations'

    def to_dict(self):
        return {'id': self.id, 'path': str(self.path), 'pid': self.pid,
                'disk_sha256': self.disk_sha256, 'kind': self.kind,
                'description': self.description,
                'operations': [op.to_dict() for op in self.operations]}


@dataclass(frozen=True)
class WorkbookCheckpoint:
    directory: Path
    workbook: Path
    manifest: Path
    sha256: str

    def verify(self):
        """Verify both the persisted manifest and the checkpoint's workbook."""
        record = json.loads(self.manifest.read_text(encoding='utf-8'))
        if record.get('sha256') != self.sha256 or _hash(self.workbook) != self.sha256:
            raise WorkflowConflictError('Checkpoint integrity check failed')
        return True


@dataclass(frozen=True)
class WorkflowResult:
    plan_id: str
    checkpoint: WorkbookCheckpoint | None
    results: tuple[Any, ...]
    published: bool = False


class _Lease:
    """Cooperative lock, intentionally never removing a pre-existing lock."""
    def __init__(self, path):
        self.path = path.with_name('.' + path.name + '.excelmodel.lock')
        self.token = uuid4().hex

    def __enter__(self):
        try:
            with self.path.open('x', encoding='utf-8') as file:
                file.write(_json({'token': self.token, 'pid': os.getpid()}))
        except FileExistsError as exc:
            raise WorkflowConflictError('Workbook has an existing workflow lease: ' + str(self.path)) from exc
        return self

    def __exit__(self, *args):
        try:
            record = json.loads(self.path.read_text(encoding='utf-8'))
            if record.get('token') == self.token:
                self.path.unlink()
        except (FileNotFoundError, ValueError):
            pass


class AgentWorkflow:
    """Inspect, plan and apply a bounded workflow to one DesktopSession.

    Plans are single-use and bound to this object and the exact saved workbook.
    Do not interact with Excel concurrently while applying a plan. A cooperative
    lease coordinates this library's workflows, not arbitrary applications.
    """
    def __init__(self, session, *, checkpoint_directory=None, max_operations=100):
        if isinstance(max_operations, bool) or not isinstance(max_operations, int) or max_operations < 1:
            raise ValueError('max_operations must be a positive integer')
        self.session = session
        self.checkpoint_directory = Path(checkpoint_directory).resolve() if checkpoint_directory is not None else None
        self.max_operations = max_operations
        self._plans = {}

    def inspect(self):
        state = self.session.snapshot()
        return {'path': str(state.path), 'pid': state.pid, 'saved': state.saved,
                'read_only': state.read_only, 'disk_sha256': state.disk_sha256,
                'supported_operations': ['read_range', 'write_range', 'write_formula',
                                         'calculate', 'update_pivots', 'save']}

    def _clean(self):
        state = self.session.snapshot()
        if not state.saved:
            raise WorkflowConflictError('Workbook has unsaved changes; save explicitly and create a fresh plan')
        if Path(state.path).suffix.lower() != '.xlsx':
            raise ValueError('Agent workflows support .xlsx workbooks only')
        return state

    def plan(self, operations: Iterable[AgentOperation], *, description=''):
        operations = tuple(islice(iter(operations), self.max_operations + 1))
        if not operations or len(operations) > self.max_operations:
            raise ValueError('Plan must contain between 1 and max_operations operations')
        for operation in operations:
            if not isinstance(operation, AgentOperation):
                raise TypeError('Operations must be AgentOperation instances')
            operation._validate(getattr(self.session, '_max_cells', 100000))
        return self._plan(operations, description, 'operations')

    def plan_edit(self, *, description):
        """Plan a trusted Python callback handoff; callback code is never serialized."""
        if not isinstance(description, str) or not description.strip():
            raise ValueError('A file-edit plan requires a description')
        return self._plan((), description, 'file_edit')

    def _plan(self, operations, description, kind):
        if not isinstance(description, str):
            raise TypeError('description must be text')
        state = self._clean()
        if state.read_only and (kind == 'file_edit' or any(op.name != 'read_range' for op in operations)):
            raise WorkflowConflictError('Cannot plan mutations for a read-only workbook')
        plan = WorkflowPlan(uuid4().hex, Path(state.path).resolve(), state.pid,
                            state.disk_sha256, operations, description, kind)
        self._plans[plan.id] = plan
        return plan

    def _claim(self, plan, kind):
        if not isinstance(plan, WorkflowPlan) or self._plans.get(plan.id) != plan or plan.kind != kind:
            raise WorkflowConflictError('Plan is unknown, modified, consumed, or has the wrong operation kind')
        del self._plans[plan.id]
        state = self._clean()
        if (Path(state.path).resolve(), state.pid, state.disk_sha256) != (plan.path, plan.pid, plan.disk_sha256):
            raise WorkflowConflictError('Workbook changed since the plan was created')
        if state.read_only and (kind == 'file_edit' or any(op.name != 'read_range' for op in plan.operations)):
            raise WorkflowConflictError('Workbook became read-only after planning')
        return state

    def _checkpoint(self, plan):
        root = self.checkpoint_directory or plan.path.parent / '.excelmodel-checkpoints'
        root.mkdir(parents=True, exist_ok=True)
        directory = Path(mkdtemp(prefix='checkpoint-', dir=root))
        workbook = directory / plan.path.name
        shutil.copy2(plan.path, workbook)
        if _hash(workbook) != plan.disk_sha256 or _hash(plan.path) != plan.disk_sha256:
            raise WorkflowConflictError('Workbook changed while creating checkpoint; preserved copy at ' + str(workbook))
        manifest = directory / 'checkpoint.json'
        record = {'schema_version': 1, 'created_utc': datetime.now(timezone.utc).isoformat(),
                  'source': str(plan.path), 'sha256': plan.disk_sha256, 'plan': plan.to_dict()}
        with manifest.open('x', encoding='utf-8') as stream:
            stream.write(json.dumps(record, indent=2, ensure_ascii=False) + '\n')
            stream.flush()
            os.fsync(stream.fileno())
        return WorkbookCheckpoint(directory, workbook, manifest, plan.disk_sha256)

    def apply(self, plan: WorkflowPlan):
        """Apply live operations. Failure reports completed operations; no fake rollback.

        A mutating plan creates a disk checkpoint before its first operation.
        Writes remain unsaved unless a save operation is explicitly included.
        """
        self._claim(plan, 'operations')
        completed = []
        checkpoint = None
        with _Lease(plan.path):
            self._assert_unchanged(plan)
            mutating = any(op.name != 'read_range' for op in plan.operations)
            if mutating:
                checkpoint = self._checkpoint(plan)
            current_hash = plan.disk_sha256
            expected_saved = True
            try:
                for operation in plan.operations:
                    state = self.session.snapshot()
                    if Path(state.path).resolve() != plan.path or state.pid != plan.pid or state.disk_sha256 != current_hash:
                        raise WorkflowConflictError('Workbook identity or disk contents changed during operation batch')
                    if state.saved != expected_saved:
                        raise WorkflowConflictError('Workbook saved/unsaved state changed between operations')
                    if operation.name == 'write_formula':
                        arguments = operation.arguments
                        result = self.session.write_range(arguments['sheet'], arguments['address'],
                                                          [[arguments['formula']]], formulas=True)
                    else:
                        result = getattr(self.session, operation.name)(**operation.arguments)
                    completed.append(result)
                    after = self.session.snapshot()
                    if operation.name == 'save':
                        current_hash = after.disk_sha256
                    elif after.disk_sha256 != current_hash:
                        raise WorkflowConflictError('Workbook file changed during a desktop operation')
                    if operation.name != 'read_range':
                        expected_saved = after.saved
            except Exception as exc:
                raise WorkflowExecutionError('Live workflow failed; completed operations may have changed Excel',
                                             checkpoint=checkpoint, completed=completed,
                                             stage='operations', recovery_paths=(() if checkpoint is None else (checkpoint.workbook,))) from exc
        return WorkflowResult(plan.id, checkpoint, tuple(completed))

    def _assert_unchanged(self, plan):
        state = self._clean()
        if (Path(state.path).resolve(), state.pid, state.disk_sha256) != (plan.path, plan.pid, plan.disk_sha256):
            raise WorkflowConflictError('Workbook changed before execution')
        if state.read_only and (plan.kind == 'file_edit' or any(op.name != 'read_range' for op in plan.operations)):
            raise WorkflowConflictError('Workbook became read-only before execution')

    def apply_edit(self, plan: WorkflowPlan, edit: Callable[[Path, Path], Any], *, validate=None):
        """Close, stage, publish and reopen a saved workbook with recovery evidence.

        ``edit(input_path, output_path)`` must create a complete output workbook.
        It must only modify the supplied output. The callback is trusted caller
        Python, never a string supplied to eval/exec. A reopen failure preserves
        the published file and checkpoint rather than blindly rolling back.
        """
        if not callable(edit) or (validate is not None and not callable(validate)):
            raise TypeError('edit and validate must be callable')
        self._claim(plan, 'file_edit')
        checkpoint = None
        closed = False
        published = False
        stage = 'checkpoint'
        staging = None
        with _Lease(plan.path):
            self._assert_unchanged(plan)
            checkpoint = self._checkpoint(plan)
            try:
                self._assert_unchanged(plan)
                stage = 'close'
                self.session.close()
                closed = True
                if _hash(plan.path) != plan.disk_sha256:
                    raise WorkflowConflictError('Workbook changed while closing')
                staging = Path(mkdtemp(prefix='.excelmodel-edit-', dir=plan.path.parent))
                input_dir = staging / 'input'
                output_dir = staging / 'output'
                input_dir.mkdir()
                output_dir.mkdir()
                source = input_dir / plan.path.name
                output = output_dir / plan.path.name
                shutil.copy2(checkpoint.workbook, source)
                stage = 'edit'
                edit(source, output)
                checkpoint.verify()
                if _hash(source) != plan.disk_sha256:
                    raise WorkflowConflictError('Edit callback modified its read-only input copy')
                stage = 'validate'
                _validate_workbook(output)
                if validate is not None:
                    validation_hash = _hash(output)
                    validate(output)
                    if _hash(output) != validation_hash:
                        raise WorkflowConflictError('Validation callback modified the candidate workbook')
                # Preserve the completed candidate even if publication/reopen fails.
                candidate = checkpoint.directory / ('candidate-' + plan.path.name)
                shutil.copy2(output, candidate)
                stage = 'publish'
                self.session.assert_closed()
                if _hash(plan.path) != plan.disk_sha256:
                    raise WorkflowConflictError('Workbook changed during file editing; refusing to replace it')
                shutil.copymode(plan.path, output)
                os.replace(output, plan.path)
                published = True
                stage = 'reopen'
                self.session.reopen()
                closed = False
            except Exception as exc:
                recovery_error = None
                if closed and not published:
                    try:
                        self.session.reopen()
                        closed = False
                    except Exception as recovery:
                        recovery_error = str(recovery)
                paths = [checkpoint.workbook]
                if staging is not None:
                    paths.append(staging)
                if published:
                    paths.append(plan.path)
                raise WorkflowExecutionError('File handoff failed at ' + stage + '; preserved files are available for recovery',
                                             checkpoint=checkpoint, stage=stage,
                                             published=published, recovery_paths=paths,
                                             recovery_error=recovery_error) from exc
            else:
                # A cleanup error must not masquerade as a failed publication.
                # Checkpoint/candidate are intentionally retained regardless.
                shutil.rmtree(staging, ignore_errors=True)
        return WorkflowResult(plan.id, checkpoint, (), published=True)


def _validate_workbook(path):
    if not path.is_file() or path.is_symlink():
        raise ValueError('Edit callback must create a regular workbook file')
    try:
        with ZipFile(path) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)):
                raise ValueError('Workbook contains duplicate package parts')
            if not {'[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml'} <= set(names):
                raise ValueError('Output is missing required workbook package parts')
            if archive.testzip() is not None:
                raise ValueError('Output workbook has a corrupt package member')
            from xml.etree import ElementTree
            for name in ('[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml'):
                ElementTree.fromstring(archive.read(name))
    except BadZipFile as exc:
        raise ValueError('Output is not a valid workbook ZIP package') from exc
