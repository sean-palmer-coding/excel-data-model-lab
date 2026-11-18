"""Optional Windows desktop bridge. No Excel process is started on import."""
from contextlib import contextmanager
from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import importlib
import os
import re
import sys
import time
import threading
from xml.etree import ElementTree as ET
from zipfile import ZipFile, BadZipFile


class DesktopSessionError(RuntimeError):
    """A desktop operation could not be completed safely."""


class DesktopConflictError(DesktopSessionError):
    """Workbook identity, saved state or file contents changed."""


@dataclass(frozen=True)
class DesktopSnapshot:
    path: Path
    pid: int
    saved: bool
    read_only: bool
    disk_sha256: str


def _path(path):
    result = Path(path).expanduser().resolve(strict=True)
    if not result.is_file() or result.suffix.lower() != '.xlsx':
        raise DesktopSessionError('Desktop sessions require an existing .xlsx file')
    return result


def _key(path):
    return os.path.normcase(str(Path(path).resolve()))


def _digest(path):
    try:
        before = path.stat()
        digest = sha256()
        with path.open('rb') as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b''):
                digest.update(chunk)
        after = path.stat()
        if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
            raise DesktopConflictError('Workbook changed while reading its fingerprint')
        return digest.hexdigest()
    except OSError as exc:
        raise DesktopSessionError('Cannot fingerprint the workbook file') from exc


def _preflight(path):
    # update_links=False does not suppress refresh-on-open connections.
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    try:
        with ZipFile(path) as archive:
            roots = [ET.fromstring(archive.read(name)) for name in archive.namelist()
                     if name.endswith('.xml')]
            connections = {}
            for root in roots:
                if root.tag == ns + 'connections':
                    for connection in root.findall(ns + 'connection'):
                        cid = connection.get('id')
                        db = connection.find(ns + 'dbPr')
                        embedded = (connection.get('name', '').casefold() == 'thisworkbookdatamodel'
                                    and db is not None and db.get('connection') == 'Data Model Connection')
                        # Duplicate IDs across roots cannot relax the check.
                        connections[cid] = connections.get(cid, True) and embedded
            for root in roots:
                for element in root.iter():
                    if element.tag not in (ns + 'connection', ns + 'queryTable', ns + 'pivotCacheDefinition'):
                        continue
                    automatic = (any(element.get(attr, '').lower() in ('1', 'true')
                                     for attr in ('refreshOnLoad', 'refreshOnOpen'))
                                 or element.get('interval', '0') not in ('', '0'))
                    if not automatic:
                        continue
                    if element.tag == ns + 'pivotCacheDefinition':
                        source = element.find(ns + 'cacheSource')
                        if source is not None and connections.get(source.get('connectionId'), False):
                            continue  # Embedded report layout initialization, not external import refresh.
                    raise DesktopSessionError('Disable external refresh-on-open before opening a managed desktop session')
    except (BadZipFile, ET.ParseError) as exc:
        raise DesktopSessionError('Workbook is not a valid XLSX package') from exc


def _load_xlwings():
    if sys.platform != 'win32':
        raise DesktopSessionError('Desktop sessions currently require Windows Excel')
    try:
        return importlib.import_module('xlwings')
    except ImportError as exc:
        raise DesktopSessionError('Install excel-data-model-lab[desktop] and Microsoft Excel for desktop sessions') from exc


def _matches(xw, path, pid=None):
    found = []
    try:
        pids = list(xw.apps.keys())
        if pid is not None:
            if isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
                raise DesktopSessionError('pid must be a positive integer')
            pids = [p for p in pids if p == pid]
        for app_pid in pids:
            app = xw.apps[app_pid]
            for book in app.books:
                if _key(book.fullname) == _key(path):
                    found.append((app, book))
    except DesktopSessionError:
        raise
    except Exception as exc:
        raise DesktopSessionError('Could not enumerate Excel workbooks; close busy dialogs and retry') from exc
    return found


def _bounds(address, limit):
    if not isinstance(address, str):
        raise DesktopSessionError('Range must be a bounded A1 rectangle')
    match = re.fullmatch(r'\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?', address)
    if not match:
        raise DesktopSessionError('Range must be a bounded A1 rectangle without sheet or external references')
    def col(text):
        value = 0
        for char in text.upper():
            value = value * 26 + ord(char) - 64
        return value
    c1, r1 = col(match[1]), int(match[2])
    c2, r2 = col(match[3] or match[1]), int(match[4] or match[2])
    if not (1 <= c1 <= c2 <= 16384 and 1 <= r1 <= r2 <= 1048576):
        raise DesktopSessionError('Range is outside Excel bounds or reversed')
    if (r2-r1+1)*(c2-c1+1) > limit:
        raise DesktopSessionError('Range exceeds the session max_cells limit')
    return r2-r1+1, c2-c1+1


@contextmanager
def _safe_open(app):
    api = app.api
    security, events = api.AutomationSecurity, api.EnableEvents
    try:
        api.AutomationSecurity = 3
        api.EnableEvents = False
        yield
    finally:
        try:
            api.EnableEvents = events
        finally:
            api.AutomationSecurity = security


def _ready(book, timeout=30):
    deadline = time.monotonic() + timeout
    while True:
        try:
            if book.name and book.fullname:
                return
        except Exception:
            pass
        if time.monotonic() >= deadline:
            raise DesktopSessionError('Excel did not finish opening the workbook within 30 seconds')
        time.sleep(0.1)


def _initialize_model(book, timeout=5):
    # Excel rejects this call while automatic CUBE calculation initializes the
    # same model. Retry only the observed busy HRESULT, never report mutations.
    deadline = time.monotonic() + timeout
    while True:
        try:
            book.api.Model.Initialize()
            return
        except Exception as exc:
            codes = [getattr(exc, 'hresult', None)]
            args = getattr(exc, 'args', ())
            if args:
                codes.append(args[0])
            if len(args) > 2 and isinstance(args[2], (tuple, list)) and len(args[2]) > 5:
                codes.append(args[2][5])
            busy = any(isinstance(code, int) and (code & 0xffffffff) == 0x800AC472 for code in codes)
            if not busy or time.monotonic() >= deadline:
                raise
            time.sleep(0.2)


class DesktopSession:
    """A single exact-path workbook, used from the thread that created it.

    Closing never quits Excel. ``quit`` only releases an owned, empty app.
    No context-manager cleanup discards unsaved work implicitly.
    """
    def __init__(self, app, book, path, *, owned=False, max_cells=100000):
        if isinstance(max_cells, bool) or not isinstance(max_cells, int) or max_cells < 1:
            raise DesktopSessionError('max_cells must be a positive integer')
        self._thread_id = threading.get_ident()
        self._app, self._book, self._path = app, book, path
        self._owned, self._max_cells = owned, max_cells
        self._pid = int(app.pid)
        self._read_only = bool(book.api.ReadOnly)
        self._closed = False
        self._quit = False

    @classmethod
    def attach(cls, path, *, pid=None, max_cells=100000):
        path = _path(path)
        found = _matches(_load_xlwings(), path, pid)
        if len(found) != 1:
            raise DesktopConflictError('Expected exactly one matching open workbook; specify its PID if ambiguous')
        return cls(*found[0], path, max_cells=max_cells)

    @classmethod
    def open(cls, path, *, visible=False, read_only=False, max_cells=100000):
        path = _path(path)
        _preflight(path)
        xw = _load_xlwings()
        if _matches(xw, path):
            raise DesktopConflictError('Workbook is already open; attach explicitly instead')
        if isinstance(max_cells, bool) or not isinstance(max_cells, int) or max_cells < 1:
            raise DesktopSessionError('max_cells must be a positive integer')
        app = xw.App(visible=visible, add_book=False)
        try:
            with _safe_open(app):
                book = app.books.open(str(path), update_links=False, read_only=read_only,
                                      ignore_read_only_recommended=True, notify=False, add_to_mru=False)
            _ready(book)
            return cls(app, book, path, owned=True, max_cells=max_cells)
        except Exception as exc:
            # Never kill the process or discard a workbook whose opening partially succeeded.
            try:
                if len(app.books) == 0:
                    app.quit()
            except Exception:
                pass
            if isinstance(exc, DesktopSessionError):
                raise
            raise DesktopSessionError(f'Excel could not open the workbook; inspect instance PID {app.pid} before retrying') from exc

    @property
    def path(self):
        return self._path

    @property
    def pid(self):
        return self._pid

    def _check_thread(self):
        if threading.get_ident() != self._thread_id:
            raise DesktopSessionError('Use the desktop session only from its creating thread')

    def _check(self, *, writable=False):
        self._check_thread()
        if self._closed or self._quit:
            raise DesktopSessionError('Workbook session is closed')
        try:
            if _key(self._book.fullname) != _key(self._path) or int(self._app.pid) != self._pid:
                raise DesktopConflictError('Workbook identity changed (possibly Save As)')
            if writable and bool(self._book.api.ReadOnly):
                raise DesktopSessionError('Workbook is read-only')
        except DesktopSessionError:
            raise
        except Exception as exc:
            raise DesktopSessionError('Workbook is unavailable or Excel is busy') from exc

    def snapshot(self):
        self._check()
        try:
            return DesktopSnapshot(self._path, self._pid, bool(self._book.api.Saved),
                                   bool(self._book.api.ReadOnly), _digest(self._path))
        except DesktopSessionError:
            raise
        except Exception as exc:
            raise DesktopSessionError('Cannot inspect workbook state') from exc

    def _range(self, sheet, address, *, writable=False):
        self._check(writable=writable)
        if not isinstance(sheet, str) or not sheet:
            raise DesktopSessionError('sheet must be an exact worksheet name')
        shape = _bounds(address, self._max_cells)
        try:
            return self._book.sheets[sheet].range(address), shape
        except Exception as exc:
            raise DesktopSessionError('Cannot access the requested worksheet range') from exc

    def read_range(self, sheet, address):
        """Return a two-dimensional list, including for single cells."""
        target, _ = self._range(sheet, address)
        try:
            return target.options(ndim=2).value
        except Exception as exc:
            raise DesktopSessionError('Could not read the requested range') from exc

    def write_range(self, sheet, address, values, *, formulas=False):
        """Write an exact rectangular matrix. Values beginning '=' stay literal."""
        target, shape = self._range(sheet, address, writable=True)
        if not isinstance(values, (list, tuple)) or len(values) != shape[0] or any(
                not isinstance(row, (list, tuple)) or len(row) != shape[1] for row in values):
            raise DesktopSessionError('values must be a two-dimensional matrix matching the range exactly')
        matrix = [list(row) for row in values]
        if formulas:
            if any(not isinstance(value, str) or not value.startswith('=') for row in matrix for value in row):
                raise DesktopSessionError('Formula writes require strings beginning with = in every cell')
        else:
            # Excel interprets leading =,+,-,@ as formulas in some formats/providers.
            matrix = [["'" + v if isinstance(v, str) and v.startswith(('=', '+', '-', '@', "'")) else v
                       for v in row] for row in matrix]
        try:
            if formulas:
                target.formula = matrix
            else:
                target.value = matrix
        except Exception as exc:
            raise DesktopSessionError('Range write failed; Excel may have partially applied it') from exc

    def calculate(self):
        """Calculate this workbook's sheets; CUBE formulas may contact their endpoint."""
        self._check()
        try:
            for sheet in self._book.sheets:
                sheet.api.Calculate()
        except Exception as exc:
            raise DesktopSessionError('Workbook calculation failed') from exc

    def update_pivots(self):
        """Initialize the embedded model and update report layouts, without RefreshAll."""
        self._check(writable=True)
        try:
            _initialize_model(self._book)
            count = 0
            for sheet in self._book.sheets:
                pivots = sheet.api.PivotTables()
                for index in range(1, pivots.Count + 1):
                    pivot = pivots.Item(index)
                    # Only embedded-model reports are part of this operation.
                    cache = pivot.PivotCache()
                    if not cache.OLAP:
                        continue
                    connection = cache.WorkbookConnection
                    if str(connection.Name).casefold() == 'thisworkbookdatamodel':
                        pivot.Update()
                        count += 1
            return count
        except Exception as exc:
            raise DesktopSessionError('Model report update failed; some layouts may already have updated') from exc

    def refresh(self):
        """Explicitly request RefreshAll; may authenticate/contact external sources.

        This starts Excel refresh and does not certify asynchronous completion.
        """
        self._check(writable=True)
        try:
            self._book.api.RefreshAll()
        except Exception as exc:
            raise DesktopSessionError('Excel refresh failed or requires interaction') from exc

    def save(self):
        self._check(writable=True)
        try:
            self._book.save()
            result = self.snapshot()
            if not result.saved:
                raise DesktopConflictError('Excel did not finish saving the workbook')
            return result
        except DesktopSessionError:
            raise
        except Exception as exc:
            raise DesktopSessionError('Workbook save failed') from exc

    def close(self, *, save=False, discard=False):
        if save and discard:
            raise DesktopSessionError('Choose save or discard, not both')
        current = self.snapshot()
        if save:
            self.save()
        elif not current.saved and not discard:
            raise DesktopConflictError('Unsaved workbook changes require explicit save=True or discard=True')
        try:
            self._book.close()
            self._closed = True
        except Exception as exc:
            raise DesktopSessionError('Workbook close failed') from exc

    def assert_closed(self):
        """Reject publication if this session or another instance has the file open.

        This is an optimistic check, not an OS lock against a subsequent open.
        """
        self._check_thread()
        if not self._closed:
            raise DesktopConflictError('Workbook must be closed before file-level publication')
        if _matches(_load_xlwings(), self._path):
            raise DesktopConflictError('Workbook has been reopened in Excel; abort file-level publication')

    def reopen(self):
        self._check_thread()
        if not self._closed or self._quit:
            raise DesktopSessionError('Reopen requires a closed workbook and a live app')
        _preflight(self._path)
        self.assert_closed()
        try:
            with _safe_open(self._app):
                self._book = self._app.books.open(str(self._path), update_links=False,
                    read_only=self._read_only, ignore_read_only_recommended=True, notify=False, add_to_mru=False)
            _ready(self._book)
            self._closed = False
            self._check()
            return self
        except DesktopSessionError:
            raise
        except Exception as exc:
            raise DesktopSessionError('Reopen failed; the file remains available on disk') from exc

    def quit(self):
        """Quit only an app created by this session and containing no workbooks."""
        self._check_thread()
        if not self._owned:
            raise DesktopSessionError('Cannot quit an attached Excel instance')
        if not self._closed:
            raise DesktopSessionError('Close the workbook explicitly before quitting')
        try:
            if len(self._app.books):
                raise DesktopConflictError('Excel now contains other workbooks; leave the app running')
            self._app.quit()
            self._quit = True
        except DesktopSessionError:
            raise
        except Exception as exc:
            raise DesktopSessionError('Could not quit the owned Excel instance') from exc
