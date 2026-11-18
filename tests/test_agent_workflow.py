from dataclasses import dataclass, replace
import hashlib
import json
from pathlib import Path
import shutil
from zipfile import ZipFile

import pytest

from excel_model_lab.agent_workflow import (
    AgentOperation, AgentWorkflow, WorkflowConflictError, WorkflowExecutionError,
)


def workbook(path, marker='original'):
    with ZipFile(path, 'w') as archive:
        archive.writestr('[Content_Types].xml', '<Types/>')
        archive.writestr('_rels/.rels', '<Relationships/>')
        archive.writestr('xl/workbook.xml', '<workbook/>')
        archive.writestr('marker', marker)
    return path


def marker(path):
    with ZipFile(path) as archive:
        return archive.read('marker').decode()


@dataclass(frozen=True)
class Snapshot:
    path: Path
    pid: int
    saved: bool
    read_only: bool
    disk_sha256: str


class Session:
    def __init__(self, path):
        self.path = path.resolve()
        self.pid = 123
        self.saved = True
        self.read_only = False
        self.closed = False
        self.calls = []
        self.reopen_error = False

    def snapshot(self):
        if self.closed:
            raise RuntimeError('closed')
        return Snapshot(self.path, self.pid, self.saved, self.read_only,
                        hashlib.sha256(self.path.read_bytes()).hexdigest())

    def close(self):
        if not self.saved:
            raise RuntimeError('dirty')
        self.closed = True
        self.calls.append('close')

    def reopen(self):
        self.calls.append('reopen')
        if self.reopen_error:
            raise RuntimeError('Excel cannot open')
        self.closed = False
        self.saved = True
        return self

    def assert_closed(self):
        if not self.closed:
            raise RuntimeError('Workbook was reopened by another user')

    def read_range(self, sheet, address):
        self.calls.append(('read', sheet, address))
        return [[1]]

    def write_range(self, sheet, address, values, *, formulas=False):
        self.calls.append(('write', sheet, address, values, formulas))
        self.saved = False

    def calculate(self):
        self.calls.append('calculate')

    def update_pivots(self):
        self.calls.append('update_pivots')

    def save(self):
        workbook(self.path, 'saved live edits')
        self.saved = True
        self.calls.append('save')


@pytest.fixture
def setup(tmp_path):
    session = Session(workbook(tmp_path / 'model.xlsx'))
    return session, AgentWorkflow(session)


def test_inspect_and_read_plan_has_no_checkpoint_or_optional_dependency(setup):
    session, workflow = setup
    assert workflow.inspect()['saved']
    plan = workflow.plan([AgentOperation.create('read_range', sheet='Sheet1', address='A1')])
    assert json.loads(json.dumps(plan.to_dict()))['operations'][0]['name'] == 'read_range'
    result = workflow.apply(plan)
    assert result.results == ([[1]],)
    assert result.checkpoint is None
    assert session.saved
    with pytest.raises(WorkflowConflictError, match='consumed'):
        workflow.apply(plan)


def test_live_batch_checkpoint_and_explicit_save(setup):
    session, workflow = setup
    plan = workflow.plan([
        AgentOperation.create('write_range', sheet='Sheet1', address='A1', values=[['=literal']]),
        AgentOperation.create('write_formula', sheet='Sheet1', address='B1', formula='=SUM(A2:A3)'),
        AgentOperation.create('calculate'), AgentOperation.create('update_pivots'),
        AgentOperation.create('save'), AgentOperation.create('read_range', sheet='Sheet1', address='B1'),
    ])
    result = workflow.apply(plan)
    assert result.checkpoint.verify()
    assert marker(result.checkpoint.workbook) == 'original'
    assert marker(session.path) == 'saved live edits'
    assert session.calls[0][-1] is False
    assert session.calls[1][-1] is True
    assert session.saved
    assert not list(session.path.parent.glob('*.lock'))


def test_writes_remain_unsaved_and_new_plans_refuse_dirty_state(setup):
    session, workflow = setup
    workflow.apply(workflow.plan([AgentOperation.create('write_range', sheet='S', address='A1', values=[[5]])]))
    assert not session.saved
    with pytest.raises(WorkflowConflictError, match='unsaved'):
        workflow.plan_edit(description='change model')
    assert marker(session.path) == 'original'


@pytest.mark.parametrize('change', ['dirty', 'file', 'pid', 'path'])
def test_stale_plan_refused_before_mutation(setup, change):
    session, workflow = setup
    plan = workflow.plan_edit(description='edit')
    if change == 'dirty': session.saved = False
    elif change == 'file': workbook(session.path, 'user save')
    elif change == 'pid': session.pid = 99
    else: session.path = workbook(session.path.with_name('another.xlsx')).resolve()
    with pytest.raises(WorkflowConflictError):
        workflow.apply_edit(plan, lambda source, output: shutil.copy2(source, output))
    assert not session.calls


def test_single_use_plan_cannot_be_modified_or_shared_between_workflows(setup):
    session, workflow = setup
    plan = workflow.plan_edit(description='edit')
    with pytest.raises(WorkflowConflictError):
        workflow.apply_edit(replace(plan, description='different'), lambda s, o: None)
    with pytest.raises(WorkflowConflictError):
        AgentWorkflow(session).apply_edit(plan, lambda s, o: None)


def test_existing_lease_is_not_removed(setup):
    session, workflow = setup
    plan = workflow.plan_edit(description='edit')
    lease = session.path.with_name('.' + session.path.name + '.excelmodel.lock')
    lease.write_text('another writer')
    with pytest.raises(WorkflowConflictError, match='lease'):
        workflow.apply_edit(plan, lambda s, o: None)
    assert lease.read_text() == 'another writer'
    assert not session.calls


def test_handoff_closes_stages_publishes_and_reopens(setup):
    session, workflow = setup
    before = session.path.read_bytes()
    def edit(source, output):
        assert session.closed
        assert source.name == output.name == session.path.name
        assert source.read_bytes() == before
        assert session.path.read_bytes() == before
        workbook(output, 'modified')
    result = workflow.apply_edit(workflow.plan_edit(description='Modify a stored column'), edit)
    assert result.published
    assert marker(session.path) == 'modified'
    assert result.checkpoint.workbook.read_bytes() == before
    assert marker(result.checkpoint.directory / ('candidate-' + session.path.name)) == 'modified'
    assert session.calls == ['close', 'reopen']
    assert not session.closed
    assert not list(session.path.parent.glob('.excelmodel-edit-*'))


@pytest.mark.parametrize('failure', ['callback', 'invalid', 'input_changed', 'validator'])
def test_failure_preserves_original_and_reopens_with_recovery_paths(setup, failure):
    session, workflow = setup
    before = session.path.read_bytes()
    def edit(source, output):
        if failure == 'callback': raise RuntimeError('callback stopped')
        if failure == 'invalid': output.write_bytes(b'not xlsx')
        else: workbook(output, 'modified')
        if failure == 'input_changed': workbook(source, 'unauthorized input change')
    def validate(path):
        if failure == 'validator': raise RuntimeError('invalid model')
    with pytest.raises(WorkflowExecutionError) as error:
        workflow.apply_edit(workflow.plan_edit(description='edit'), edit, validate=validate)
    assert session.path.read_bytes() == before
    assert not session.closed
    assert not error.value.published
    assert error.value.checkpoint.verify()
    assert all(path.exists() for path in error.value.recovery_paths)


def test_concurrent_external_save_is_preserved_instead_of_replaced(setup):
    session, workflow = setup
    def edit(source, output):
        workbook(output, 'agent candidate')
        workbook(session.path, 'external user save')
    with pytest.raises(WorkflowExecutionError) as error:
        workflow.apply_edit(workflow.plan_edit(description='edit'), edit)
    assert isinstance(error.value.__cause__, WorkflowConflictError)
    assert marker(session.path) == 'external user save'
    assert marker(error.value.checkpoint.workbook) == 'original'
    assert marker(error.value.checkpoint.directory / ('candidate-' + session.path.name)) == 'agent candidate'
    assert not session.closed


def test_reopen_failure_preserves_published_file_and_checkpoint_without_blind_restore(setup):
    session, workflow = setup
    session.reopen_error = True
    with pytest.raises(WorkflowExecutionError) as error:
        workflow.apply_edit(workflow.plan_edit(description='edit'), lambda s, o: workbook(o, 'published'))
    assert error.value.stage == 'reopen'
    assert error.value.published
    assert marker(session.path) == 'published'
    assert marker(error.value.checkpoint.workbook) == 'original'
    assert session.path in error.value.recovery_paths
    assert session.calls == ['close', 'reopen']


def test_callback_and_recovery_failure_report_both(setup):
    session, workflow = setup
    session.reopen_error = True
    def fail(source, output): raise RuntimeError('bad callback')
    with pytest.raises(WorkflowExecutionError) as error:
        workflow.apply_edit(workflow.plan_edit(description='edit'), fail)
    assert 'bad callback' in str(error.value.__cause__)
    assert error.value.recovery_error == 'Excel cannot open'
    assert marker(session.path) == 'original'


def test_partial_live_failure_reports_checkpoint_and_completed_operations(setup):
    session, workflow = setup
    def fail(): raise RuntimeError('calculation failure')
    session.calculate = fail
    plan = workflow.plan([AgentOperation.create('write_range', sheet='S', address='A1', values=[[5]]),
                          AgentOperation.create('calculate')])
    with pytest.raises(WorkflowExecutionError) as error:
        workflow.apply(plan)
    assert error.value.completed == (None,)
    assert not session.saved
    assert error.value.checkpoint.verify()


def test_read_only_mutation_rejected_and_operations_cannot_eval(setup):
    session, workflow = setup
    session.read_only = True
    with pytest.raises(WorkflowConflictError, match='read-only'):
        workflow.plan_edit(description='edit')
    with pytest.raises(WorkflowConflictError, match='read-only'):
        workflow.plan([AgentOperation.create('save')])
    for name in ('eval', 'exec', '__getattribute__', 'refresh', 'quit'):
        with pytest.raises(ValueError, match='Unsupported'):
            AgentOperation.create(name)


def test_operation_values_are_defensive_copies_and_bounds_are_enforced(setup):
    session, workflow = setup
    values = [[1]]
    operation = AgentOperation.create('write_range', sheet='S', address='A1', values=values)
    values[0][0] = 2
    operation.arguments['values'][0][0] = 3
    assert operation.arguments['values'] == [[1]]
    with pytest.raises(ValueError): workflow.plan([])
    with pytest.raises(ValueError): AgentWorkflow(session, max_operations=0)
    with pytest.raises(ValueError): AgentOperation.create('calculate', arbitrary=True)
    with pytest.raises(ValueError): AgentOperation.create('write_range', sheet='S', address='A1', values=[[float('nan')]])


def test_checkpoint_tampering_is_detected(setup):
    session, workflow = setup
    result = workflow.apply(workflow.plan([AgentOperation.create('calculate')]))
    workbook(result.checkpoint.workbook, 'tampered')
    with pytest.raises(WorkflowConflictError, match='integrity'):
        result.checkpoint.verify()


@pytest.mark.parametrize('kwargs', [
    {'name':'write_range','sheet':'S','address':'A1:B1','values':[[1]]},
    {'name':'write_range','sheet':'S','address':'A1','values':[[{'bad':1}]]},
    {'name':'write_formula','sheet':'S','address':'A1:B1','formula':'=1'},
    {'name':'write_formula','sheet':'S','address':'A1','formula':'not a formula'},
    {'name':'read_range','sheet':'S','address':'A:A'},
    {'name':'read_range','sheet':'S','address':'XFE1'},
])
def test_invalid_operation_rejected_before_any_execution(setup, kwargs):
    session, workflow = setup
    with pytest.raises((ValueError, RuntimeError)):
        AgentOperation.create(**kwargs)
    assert not session.calls


def test_plan_iterator_is_bounded(setup):
    session, workflow = setup
    count = 0
    def forever():
        nonlocal count
        while True:
            count += 1
            yield AgentOperation.create('calculate')
    with pytest.raises(ValueError):
        workflow.plan(forever())
    assert count == 101


def test_user_edit_between_read_and_write_detected(setup):
    session, workflow = setup
    def read(sheet, address):
        session.saved = False
        return [[1]]
    session.read_range = read
    plan = workflow.plan([AgentOperation.create('read_range', sheet='S', address='A1'),
                          AgentOperation.create('write_range', sheet='S', address='A1', values=[[9]])])
    with pytest.raises(WorkflowExecutionError) as error:
        workflow.apply(plan)
    assert isinstance(error.value.__cause__, WorkflowConflictError)
    assert not session.calls


def test_reopened_workbook_during_callback_blocks_publish(setup):
    session, workflow = setup
    def edit(source, output):
        workbook(output, 'candidate')
        session.closed = False
    with pytest.raises(WorkflowExecutionError) as error:
        workflow.apply_edit(workflow.plan_edit(description='edit'), edit)
    assert not error.value.published
    assert marker(session.path) == 'original'


def test_validation_callback_may_not_modify_output(setup):
    session, workflow = setup
    with pytest.raises(WorkflowExecutionError) as error:
        workflow.apply_edit(workflow.plan_edit(description='edit'),
                            lambda s,o: workbook(o,'candidate'),
                            validate=lambda o:o.write_bytes(b'corrupt'))
    assert isinstance(error.value.__cause__, WorkflowConflictError)
    assert marker(session.path) == 'original'
