import sqlite3

import pandas as pd
import pytest

from excel_model_lab import WorkbookModel, CsvSource, ParquetSource, SqlSource


def test_csv_sql_refresh_and_save(tmp_path):
    csv = tmp_path / 'values.csv'
    csv.write_text('ID,Label\n7,Seven\n8,Eight\n', encoding='utf-8')
    model = WorkbookModel.create({'CSV': pd.DataFrame({'ID':[1], 'Label':['One']}),
                                  'SQL': pd.DataFrame({'ID':[1], 'Amount':[1]})})
    with sqlite3.connect(':memory:') as db:
        db.execute('create table sales (id integer, amount integer)')
        db.executemany('insert into sales values (?,?)', [(1,10),(2,20),(3,30)])
        model.refresh_sources({'CSV': CsvSource(csv),
            'SQL': SqlSource(db, 'select id as ID, amount as Amount from sales where id >= ?', params=(2,))})
    assert model.get_table('SQL').to_dict('list') == {'ID':[2,3], 'Amount':[20,30]}
    copy = model.get_table('SQL')
    copy.loc[0,'Amount'] = 999
    assert model.get_table('SQL').loc[0,'Amount'] == 20
    output = model.save(tmp_path / 'refreshed.xlsx')
    reopened = WorkbookModel(output)
    assert reopened.tables == ('CSV','SQL')
    assert reopened.get_table('CSV').to_dict('list') == {'ID':[7,8], 'Label':['Seven','Eight']}


def test_refresh_failures_do_not_partially_stage(tmp_path):
    csv = tmp_path / 'good.csv'
    csv.write_text('ID\n9\n', encoding='utf-8')
    model = WorkbookModel.create({'A':pd.DataFrame({'ID':[1]}),'B':pd.DataFrame({'ID':[2]})})
    before = dict(model._pending)
    with pytest.raises(FileNotFoundError):
        model.refresh_sources({'A':CsvSource(csv),'B':CsvSource(tmp_path/'missing.csv')})
    assert model._pending == before
    csv.write_text('Wrong\n9\n', encoding='utf-8')
    with pytest.raises(ValueError, match='schema mismatch'):
        model.refresh_sources({'A':CsvSource(csv)})
    assert model._pending == before


def test_parquet_refresh(tmp_path):
    pytest.importorskip('pyarrow')
    source = tmp_path / 'data.parquet'
    pd.DataFrame({'ID':[4,5]}).to_parquet(source)
    model = WorkbookModel.create({'Data':pd.DataFrame({'ID':[1]})})
    model.refresh_sources({'Data':ParquetSource(source)})
    assert model.get_table('Data')['ID'].tolist() == [4,5]
