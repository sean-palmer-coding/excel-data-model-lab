from pathlib import Path
from types import SimpleNamespace
from zipfile import ZipFile
import threading
import pytest
import excel_model_lab.desktop_session as ds


class Range:
    def __init__(self, book):
        self.book = book
        self._value = [[42]]
        self.formula = None
    def options(self, **kwargs):
        assert kwargs == {'ndim': 2}
        return self
    @property
    def value(self):
        return self._value
    @value.setter
    def value(self, value):
        self._value = value
        self.book.api.Saved = False


class Sheet:
    def __init__(self, book):
        self.target = Range(book)
        self.calculated = 0
        self.api = SimpleNamespace(Calculate=self.calculate)
    def range(self, address):
        return self.target
    def calculate(self):
        self.calculated += 1


class Sheets:
    def __init__(self, book):
        self.sheet = Sheet(book)
    def __getitem__(self, name):
        if name != 'Data':
            raise KeyError(name)
        return self.sheet
    def __iter__(self):
        return iter([self.sheet])


class Book:
    def __init__(self, app, path):
        self.app = app
        self.fullname = str(path)
        self.name = Path(path).name
        self.refreshed = 0
        self.api = SimpleNamespace(Saved=True, ReadOnly=False, RefreshAll=self.refresh)
        self.sheets = Sheets(self)
    def save(self):
        self.api.Saved = True
    def close(self):
        self.app.books.items.remove(self)
    def refresh(self):
        self.refreshed += 1


class Books:
    def __init__(self, app):
        self.app = app
        self.items = []
        self.open_options = []
    def __iter__(self):
        return iter(self.items)
    def __len__(self):
        return len(self.items)
    def open(self, path, **kwargs):
        self.open_options.append(kwargs)
        assert self.app.api.AutomationSecurity == 3
        assert self.app.api.EnableEvents is False
        book = Book(self.app, path)
        book.api.ReadOnly = kwargs['read_only']
        self.items.append(book)
        return book


class App:
    def __init__(self, pid=51):
        self.pid = pid
        self.books = Books(self)
        self.api = SimpleNamespace(AutomationSecurity=1, EnableEvents=True)
        self.quitted = False
    def quit(self):
        self.quitted = True


@pytest.fixture
def env(tmp_path, monkeypatch):
    path = tmp_path/'source.xlsx'
    with ZipFile(path, 'w') as z:
        z.writestr('xl/connections.xml', '<connections/>')
    apps = {}
    def new_app(**kwargs):
        assert kwargs['add_book'] is False
        app = App(max(apps, default=50)+1)
        apps[app.pid] = app
        return app
    xw = SimpleNamespace(apps=apps, App=new_app)
    monkeypatch.setattr(ds, '_load_xlwings', lambda: xw)
    return path, apps, xw


def test_open_avoids_active_book_and_restores_security(env):
    path, apps, _ = env
    s = ds.DesktopSession.open(path)
    assert s.snapshot() == ds.DesktopSnapshot(path, 51, True, False, ds._digest(path))
    assert apps[51].api.AutomationSecurity == 1
    assert apps[51].api.EnableEvents is True
    assert apps[51].books.open_options[0]['update_links'] is False
    assert s.read_range('Data','A1') == [[42]]
    s.calculate()
    assert s._book.sheets.sheet.calculated == 1
    assert s._book.refreshed == 0
    s.refresh()
    assert s._book.refreshed == 1


def test_attach_exact_path_and_pid(env, tmp_path):
    path, apps, _ = env
    s = ds.DesktopSession.open(path)
    other = App(52); apps[52] = other
    other.books.items.append(Book(other, tmp_path/'other'/'source.xlsx'))
    assert ds.DesktopSession.attach(path).pid == 51
    other.books.items.append(Book(other,path))
    with pytest.raises(ds.DesktopConflictError, match='exactly one'):
        ds.DesktopSession.attach(path)
    assert ds.DesktopSession.attach(path,pid=52).pid == 52
    with pytest.raises(ds.DesktopConflictError, match='already open'):
        ds.DesktopSession.open(path)


def test_dirty_close_requires_explicit_decision_and_attached_never_quits(env):
    path, apps, _ = env
    ds.DesktopSession.open(path)
    s = ds.DesktopSession.attach(path)
    s.write_range('Data','A1',[[12]])
    with pytest.raises(ds.DesktopConflictError, match='Unsaved'):
        s.close()
    assert len(apps[51].books) == 1
    s.close(save=True)
    with pytest.raises(ds.DesktopSessionError, match='attached'):
        s.quit()
    assert not apps[51].quitted


def test_reopen_and_owned_empty_quit(env, tmp_path):
    path, apps, _ = env
    s = ds.DesktopSession.open(path)
    s.close()
    assert s.reopen() is s
    assert len(apps[51].books.open_options) == 2
    s.close()
    foreign = Book(apps[51], tmp_path/'foreign.xlsx')
    apps[51].books.items.append(foreign)
    with pytest.raises(ds.DesktopConflictError, match='other workbooks'):
        s.quit()
    apps[51].books.items.remove(foreign)
    s.quit()
    assert apps[51].quitted


@pytest.mark.parametrize('address', ['A:A','1:1','A0','A1048577','XFE1','B2:A1','[other]Sheet!A1','Data!A1','A1,B2','name'])
def test_reject_unbounded_or_invalid_ranges_before_access(env,address):
    s = ds.DesktopSession.open(env[0])
    with pytest.raises(ds.DesktopSessionError):
        s.read_range('Data',address)


def test_matrix_size_bounds_literal_strings_and_formulas(env):
    s = ds.DesktopSession.open(env[0],max_cells=2)
    with pytest.raises(ds.DesktopSessionError, match='max_cells'):
        s.read_range('Data','A1:C1')
    with pytest.raises(ds.DesktopSessionError, match='matrix'):
        s.write_range('Data','A1:B1',[[1]])
    s.write_range('Data','A1:B1',[['=1+1','@something']])
    assert s._book.sheets.sheet.target.value == [["'=1+1","'@something"]]
    s.write_range('Data','A1',[['=1+1']],formulas=True)
    assert s._book.sheets.sheet.target.formula == [['=1+1']]
    with pytest.raises(ds.DesktopSessionError, match='Formula'):
        s.write_range('Data','A1',[[2]],formulas=True)


def test_read_only_and_save_as_detection(env):
    s = ds.DesktopSession.open(env[0],read_only=True)
    with pytest.raises(ds.DesktopSessionError,match='read-only'):
        s.write_range('Data','A1',[[1]])
    with pytest.raises(ds.DesktopSessionError,match='read-only'):
        s.save()
    s._book.fullname += '.changed'
    with pytest.raises(ds.DesktopConflictError,match='identity'):
        s.snapshot()


def test_fingerprint_changes_for_external_edits(env):
    s = ds.DesktopSession.open(env[0]); before = s.snapshot()
    with env[0].open('ab') as f:
        f.write(b'changed')
    assert s.snapshot().disk_sha256 != before.disk_sha256


def test_preflight_rejects_refresh_before_process_start(env):
    path, apps, _ = env
    with ZipFile(path,'w') as z:
        z.writestr('xl/connections.xml','<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection refreshOnLoad="1"/></connections>')
    with pytest.raises(ds.DesktopSessionError,match='refresh-on-open'):
        ds.DesktopSession.open(path)
    assert apps == {}


def test_failed_open_does_not_discard_partially_opened_book(env,monkeypatch):
    path,apps,_ = env
    original = Books.open
    def fail(self,*args,**kwargs):
        original(self,*args,**kwargs)
        raise RuntimeError('late COM failure')
    monkeypatch.setattr(Books,'open',fail)
    with pytest.raises(ds.DesktopSessionError,match='could not open'):
        ds.DesktopSession.open(path)
    assert not apps[51].quitted
    assert apps[51].api.AutomationSecurity == 1


def test_wrong_thread_is_rejected(env):
    s = ds.DesktopSession.open(env[0]); errors=[]
    def run():
        try: s.snapshot()
        except Exception as exc: errors.append(exc)
    thread = threading.Thread(target=run);thread.start();thread.join()
    assert isinstance(errors[0], ds.DesktopSessionError)
    assert 'thread' in str(errors[0])


def test_alternate_connection_part_and_periodic_refresh_rejected(env):
    path, _, _ = env
    with ZipFile(path, 'w') as z:
        z.writestr('xl/connections2.xml', '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection interval="5"/></connections>')
    with pytest.raises(ds.DesktopSessionError, match='refresh-on-open'):
        ds.DesktopSession.open(path)


def test_embedded_pivot_initialization_allowed_but_external_cache_rejected(env):
    path, _, _ = env
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    def fixture(cid):
        with ZipFile(path, 'w') as z:
            z.writestr('xl/connections2.xml', f'<connections xmlns="{ns}"><connection id="7" name="ThisWorkbookDataModel"><dbPr connection="Data Model Connection"/></connection></connections>')
            z.writestr('xl/pivotCache/pivotCacheDefinition1.xml', f'<pivotCacheDefinition xmlns="{ns}" refreshOnLoad="1"><cacheSource connectionId="{cid}"/></pivotCacheDefinition>')
    fixture(7)
    ds._preflight(path)
    fixture(8)
    with pytest.raises(ds.DesktopSessionError, match='refresh-on-open'):
        ds._preflight(path)


def test_assert_closed_detects_reopened_file(env):
    path, apps, _ = env
    s = ds.DesktopSession.open(path)
    with pytest.raises(ds.DesktopConflictError, match='must be closed'):
        s.assert_closed()
    s.close(); s.assert_closed()
    other=App(52); apps[52]=other
    other.books.items.append(Book(other,path))
    with pytest.raises(ds.DesktopConflictError, match='reopened'):
        s.assert_closed()
    with pytest.raises(ds.DesktopConflictError, match='reopened'):
        s.reopen()


def test_update_pivots_skips_non_model_reports(env):
    s = ds.DesktopSession.open(env[0]); updates=[]
    def pivot(olap,name):
        cache=SimpleNamespace(OLAP=olap, WorkbookConnection=SimpleNamespace(Name=name))
        return SimpleNamespace(PivotCache=lambda:cache,Update=lambda:updates.append(name))
    items=[pivot(False,'worksheet'),pivot(True,'External'),pivot(True,'ThisWorkbookDataModel')]
    s._book.api.Model=SimpleNamespace(Initialize=lambda:None)
    s._book.sheets.sheet.api.PivotTables=lambda:SimpleNamespace(Count=3,Item=lambda i:items[i-1])
    assert s.update_pivots()==1
    assert updates==['ThisWorkbookDataModel']


def test_model_initialization_retries_only_busy_exception(env, monkeypatch):
    s = ds.DesktopSession.open(env[0]); calls=[]
    busy=RuntimeError(-2147352567,'Exception occurred',(0,'Excel','busy',None,0,-2146777998))
    def initialize():
        calls.append(1)
        if len(calls)==1:raise busy
    s._book.api.Model=SimpleNamespace(Initialize=initialize)
    s._book.sheets.sheet.api.PivotTables=lambda:SimpleNamespace(Count=0)
    monkeypatch.setattr(ds.time,'sleep',lambda seconds:None)
    assert s.update_pivots()==0
    assert len(calls)==2


def test_model_initialization_persistent_busy_has_bounded_retry(env, monkeypatch):
    s = ds.DesktopSession.open(env[0]); calls=[]
    busy=RuntimeError(-2147352567,'Exception occurred',(0,'Excel','busy',None,0,-2146777998))
    def initialize():
        calls.append(1)
        raise busy
    s._book.api.Model=SimpleNamespace(Initialize=initialize)
    ticks=iter([0,0,1,6])
    monkeypatch.setattr(ds.time,'monotonic',lambda:next(ticks))
    monkeypatch.setattr(ds.time,'sleep',lambda seconds:None)
    with pytest.raises(ds.DesktopSessionError) as caught:
        s.update_pivots()
    assert caught.value.__cause__ is busy
    assert len(calls)==3


def test_model_initialization_unrelated_error_not_retried(env, monkeypatch):
    s = ds.DesktopSession.open(env[0]); calls=[]
    failure=RuntimeError('invalid model')
    def initialize():
        calls.append(1)
        raise failure
    s._book.api.Model=SimpleNamespace(Initialize=initialize)
    monkeypatch.setattr(ds.time,'sleep',lambda seconds:pytest.fail('unexpected retry'))
    with pytest.raises(ds.DesktopSessionError) as caught:
        s.update_pivots()
    assert caught.value.__cause__ is failure
    assert len(calls)==1
