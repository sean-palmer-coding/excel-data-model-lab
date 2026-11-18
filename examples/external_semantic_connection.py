"""Author external CUBE formulas; a deployed Analysis Services model is required.

Expected test model: Sales[Amount] = 10, 20, 30 and [Total Amount] = SUM(...).
This script does not connect, authenticate, refresh or fabricate result caches.
"""
import argparse
from pathlib import Path
from tempfile import TemporaryDirectory
import xlsxwriter
from excel_model_lab.external_connections import ExternalModelConnection, add_external_connection


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', type=Path)
    parser.add_argument('--server', default=r'localhost\ASTabular')
    parser.add_argument('--database', default='Sales')
    parser.add_argument('--cube', default='Model')
    args = parser.parse_args()
    with TemporaryDirectory() as directory:
        source = Path(directory)/'source.xlsx'
        with xlsxwriter.Workbook(source) as workbook:
            sheet = workbook.add_worksheet('Live Sales')
            sheet.set_column('A:A', 28)
            sheet.set_column('B:B', 20)
            sheet.write('A1', 'External semantic model')
            sheet.write('A3', 'Total Amount')
            # A string placeholder deliberately avoids inventing numeric data.
            sheet.write_formula('B3', '=CUBEVALUE("Live Sales","[Measures].[Total Amount]")', None, '')
            sheet.write('A5', 'Requires live authentication and calculation in Excel')
        add_external_connection(source, args.output, ExternalModelConnection(
            'Live Sales', args.server, args.database, cube=args.cube))
    print(args.output)


if __name__ == '__main__':
    main()
