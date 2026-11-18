"""Compose an XlsxWriter workbook with an embedded model and native reports."""
from pathlib import Path
from tempfile import TemporaryDirectory
import os

from .model import WorkbookModel
from .reports import WorkbookReport
from .cube_formulas import CubeFormulas


def attach_model(workbook, destination, tables, *, configure_model=None, configure_reports=None):
    """Attach a new model/reports to a finalized workbook, preserving its parts.

    The source must be closed and must not already contain an embedded model.
    Existing model workbooks should be edited with WorkbookModel instead.
    """
    if any(fn is not None and not callable(fn) for fn in (configure_model, configure_reports)):
        raise TypeError('Configuration callbacks must be callable')
    if Path(workbook).suffix.lower() != '.xlsx':
        raise ValueError('Composition source must be an .xlsx workbook')
    destination = Path(destination)
    if destination.suffix.lower() != '.xlsx':
        raise ValueError('Composition output must be an .xlsx workbook')
    with TemporaryDirectory(prefix='.excelmodel-attach-', dir=destination.parent) as staging:
        model = WorkbookModel.create(tables, workbook=workbook)
        if configure_model is not None:
            configure_model(model)
        report = WorkbookReport(model)
        if configure_reports is not None:
            configure_reports(report)
        completed = report.save(Path(staging) / destination.name)
        os.replace(completed, destination)
    return destination


def compose_workbook(destination, tables, write_worksheets, *, configure_model=None,
                     configure_reports=None, xlsxwriter_options=None):
    """Atomically create a complete workbook without an Excel installation.

    ``configure_model(model)`` runs first. ``write_worksheets(writer, cubes)``
    authors worksheets with the ordinary XlsxWriter API and validated CUBE
    helpers. ``configure_reports(report)`` then adds model-backed reports.
    Callbacks mutate the passed objects; return values are ignored. They must
    not close the writer or retain it for later use. Existing output is replaced
    only after the entire operation succeeds. Existing workbooks are not opened
    or edited by XlsxWriter. Optional import keeps core authoring independent.
    """
    if not callable(write_worksheets):
        raise TypeError('write_worksheets must be callable')
    if any(fn is not None and not callable(fn) for fn in (configure_model, configure_reports)):
        raise TypeError('Configuration callbacks must be callable')
    try:
        import xlsxwriter
    except ImportError as exc:
        raise ImportError('Workbook composition requires XlsxWriter; install excel-model-lab[workbooks]') from exc
    destination = Path(destination)
    if destination.suffix.lower() != '.xlsx':
        raise ValueError('Composition output must be an .xlsx workbook')
    options = {'strings_to_formulas': False, 'strings_to_urls': False}
    if xlsxwriter_options is not None:
        options.update(dict(xlsxwriter_options))
    # Staging beside the destination makes the final replacement atomic.
    with TemporaryDirectory(prefix='.excelmodel-compose-', dir=destination.parent) as staging:
        staging = Path(staging)
        model = WorkbookModel.create(tables)
        if configure_model is not None:
            configure_model(model)
        cubes = CubeFormulas(model)
        worksheet_path = staging / 'worksheets.xlsx'
        with xlsxwriter.Workbook(worksheet_path, options) as writer:
            write_worksheets(writer, cubes)
        # This is the same supported envelope used by create(workbook=...).
        # The model is privately owned here; no caller's open model is rebound.
        model.source = worksheet_path
        report = WorkbookReport(model)
        if configure_reports is not None:
            configure_reports(report)
        completed = report.save(staging / destination.name)
        os.replace(completed, destination)
    return destination
