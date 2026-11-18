"""Opt-in live SSAS/Excel acceptance against a NEW synthetic test database.

Requires Windows Excel, pythonnet, xlwings, XlsxWriter, TOM and ADOMD.NET.
Creates and retains the named test database (ending with total 120) and workbook.
Refuses an existing database or output directory; never uses real source data.
"""
import argparse
import hashlib
import importlib.metadata
import json
from pathlib import Path
import sys
import time
from zipfile import ZipFile


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--server', required=True)
    parser.add_argument('--database', required=True)
    parser.add_argument('--tom-directory', type=Path, required=True)
    parser.add_argument('--adomd-assembly', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if sys.platform != 'win32':
        parser.error('Native acceptance requires Windows Excel')
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    import clr
    for name in ('Microsoft.AnalysisServices.Core', 'Microsoft.AnalysisServices.Tabular'):
        clr.AddReference(str((args.tom_directory / (name + '.dll')).resolve()))
    from Microsoft.AnalysisServices.Tabular import Server
    from Microsoft.AnalysisServices import XmlaError
    import xlsxwriter
    from excel_model_lab import (ExternalModelConnection, AdomdQueryClient,
                                 ExternalQueryError, add_external_connection,
                                 inspect_external_connections, DesktopSession)
    connection = ExternalModelConnection('Live Sales', args.server, args.database)
    tom = Server()
    tom.Connect('Data Source="' + args.server.replace('"', '""') + '";Integrated Security=SSPI;')
    record = {'server': args.server, 'database': args.database, 'synthetic_data_only': True,
              'ssas_version': str(tom.Version), 'server_mode': str(tom.ServerMode),
              'xlwings': importlib.metadata.version('xlwings'),
              'pythonnet': importlib.metadata.version('pythonnet'), 'cases': {}}
    session = None

    def execute(command):
        results = tom.Execute(json.dumps(command))
        errors = [str(m) for result in results for m in result.Messages if isinstance(m, XmlaError)]
        if errors:
            raise RuntimeError('Synthetic model operation failed: ' + '; '.join(errors))

    def source(last):
        return {'type': 'calculated', 'expression':
                'DATATABLE("Amount", INTEGER, "Label", STRING, {{10,"Alpha"},{20,"Beta"},{' + str(last) + ',"Gamma"}})'}

    def wait_value(expected):
        deadline = time.monotonic() + 30
        while True:
            actual = session.read_range('Live Sales', 'B3')
            if actual == [[expected]]:
                return
            if time.monotonic() >= deadline:
                raise AssertionError(f'Expected CUBE total {expected}, observed {actual}')
            time.sleep(0.2)

    try:
        if any(db.Name == args.database for db in tom.Databases):
            raise ValueError('Refusing to overwrite an existing database')
        model = {'name': args.database, 'compatibilityLevel': 1600, 'model': {
            'culture': 'en-US', 'tables': [{'name': 'Sales', 'columns': [
                {'name': 'Amount', 'dataType': 'int64', 'type': 'calculatedTableColumn', 'sourceColumn': '[Amount]'},
                {'name': 'Label', 'dataType': 'string', 'type': 'calculatedTableColumn', 'sourceColumn': '[Label]'}],
                'partitions': [{'name': 'Synthetic', 'mode': 'import', 'source': source(30)}],
                'measures': [{'name': 'Total Amount', 'expression': 'SUM(Sales[Amount])', 'formatString': '0'}]}]}}
        execute({'create': {'database': model}})
        execute({'refresh': {'type': 'full', 'objects': [{'database': args.database}]}})
        client = AdomdQueryClient(connection, assembly_path=args.adomd_assembly.resolve())
        for label, query, expected in (
            ('total', 'EVALUATE ROW("Total",[Total Amount])', ((60,),)),
            ('labels', 'EVALUATE SELECTCOLUMNS(Sales,"Label",Sales[Label],"Amount",Sales[Amount]) ORDER BY [Amount]',
             (('Alpha', 10), ('Beta', 20), ('Gamma', 30))),
            ('blank', 'EVALUATE ROW("Blank",BLANK())', ((None,),)),
            ('empty', 'EVALUATE FILTER(Sales,FALSE())', ())):
            got = client.query(query)
            assert got.rows == expected, (label, got)
            record['cases'][label] = {'columns': got.columns, 'rows': got.rows}
        for label, query, options in (
            ('row_limit', 'EVALUATE Sales', {'max_rows': 2}),
            ('multiple_results', 'EVALUATE ROW("x",1) EVALUATE ROW("y",2)', {}),
            ('invalid_dax', 'EVALUATE ThisTableDoesNotExist', {})):
            try:
                client.query(query, **options)
            except ExternalQueryError as exc:
                record['cases'][label] = {'rejected': True, 'message': str(exc),
                                           'has_provider_cause': exc.__cause__ is not None}
            else:
                raise AssertionError(label + ' was not rejected')
        assert client.query('EVALUATE ROW("Total",[Total Amount])').rows == ((60,),)
        record['cases']['reuse_after_errors'] = True
        big = client.query('EVALUATE UNION(ROW("Exact",9007199254740993),ROW("Exact",BLANK()))')
        assert big.rows == ((9007199254740993,), (None,))
        frame = big.to_dataframe()
        assert frame.iloc[0, 0] == 9007199254740993 and frame.iloc[1, 0] is None
        record['cases']['exact_nullable_integer_dataframe'] = True
        baseline = output / 'worksheet-source.xlsx'
        with xlsxwriter.Workbook(baseline) as writer:
            sheet = writer.add_worksheet('Live Sales')
            sheet.write('A3', 'Total Amount')
            sheet.write_formula('B3', '=CUBEVALUE("Live Sales","[Measures].[Total Amount]")', None, '')
        workbook = output / 'live-sales.xlsx'
        add_external_connection(baseline, workbook, connection)
        assert len(inspect_external_connections(workbook)) == 1
        with ZipFile(workbook) as archive:
            assert 'xl/model/item.data' not in archive.namelist()
        session = DesktopSession.open(workbook)
        record['excel_version'] = str(session._app.api.Version)
        record['excel_build'] = str(session._app.api.Build)
        native = session._book.api.Connections.Item('Live Sales')
        assert native.Type == 1 and native.OLEDBConnection.OLAP
        session.calculate()
        wait_value(60)
        session.save()
        session.close()
        session.reopen()
        wait_value(60)
        record['cases']['initial_excel_save_reopen'] = 60
        before = hashlib.sha256(workbook.read_bytes()).hexdigest()
        execute({'alter': {'object': {'database': args.database, 'table': 'Sales', 'partition': 'Synthetic'},
                          'partition': {'name': 'Synthetic', 'mode': 'import', 'source': source(90)}}})
        execute({'refresh': {'type': 'full', 'objects': [{'database': args.database}]}})
        assert client.query('EVALUATE ROW("Total",[Total Amount])').rows == ((120,),)
        assert hashlib.sha256(workbook.read_bytes()).hexdigest() == before
        record['cases']['server_change_without_workbook_rewrite'] = True
        record['cases']['excel_before_refresh'] = session.read_range('Live Sales', 'B3')[0][0]
        session.refresh()
        session.calculate()
        wait_value(120)
        session.save()
        session.close()
        session.reopen()
        wait_value(120)
        record['cases']['refreshed_excel_save_reopen'] = 120
        session.close(discard=True)
        session.quit()
        session = None
        record['passed'] = True
    finally:
        if session is not None:
            # This script owns the synthetic test workbook and application only.
            try:
                session.close(discard=True)
                session.quit()
            except Exception as exc:
                record['cleanup_error'] = str(exc)
        tom.Disconnect()
        (output / 'verification.json').write_text(json.dumps(record, indent=2, default=str) + '\n', encoding='utf-8')
    print(json.dumps(record, indent=2, default=str))


if __name__ == '__main__':
    main()
