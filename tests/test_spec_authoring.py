import json
import subprocess
import sys

import pandas as pd
import pytest

from excel_model_lab import WorkbookModel
from excel_model_lab.spec import build_from_spec, load_spec
from excel_model_lab.relationship_storage import stored_relationships
from excel_model_lab.sorting import stored_sorts
from excel_model_lab.hierarchy_storage import stored_hierarchies


def test_create_cli_csv_model_and_extend(tmp_path):
    pd.DataFrame({'ID':[1,2,3], 'Label':['Alpha','Beta','Gamma'], 'Rank':[3,2,1]}).to_csv(tmp_path/'products.csv', index=False)
    spec = {'tables': {'Products': {'source':'products.csv', 'columns':{'Label':{'sort_by':'Rank'}},
        'hierarchies':[{'name':'Catalog','levels':['Label','ID'], 'level_names':{'ID':'Product number'}}]}}}
    spec_path = tmp_path/'model.json'
    spec_path.write_text(json.dumps(spec))
    output = tmp_path/'created.xlsx'
    subprocess.run([sys.executable, '-m', 'excel_model_lab.cli', 'create', str(output), str(spec_path)], check=True, capture_output=True, text=True)
    model = WorkbookModel(output)
    assert stored_sorts(model) == [('Products','Label','Rank')]
    assert stored_hierarchies(model) == [('Products','Catalog',['Label','ID'])]
    pd.DataFrame({'ID':[1,2,3,1], 'Amount':[10,20,30,40]}).to_csv(tmp_path/'sales.csv', index=False)
    extended = build_from_spec(output, tmp_path/'extended.xlsx', {'tables':{'Sales':'sales.csv'},
        'relationships':[{'from':'Sales.ID','to':'Products.ID'}]}, base_directory=tmp_path)
    model = WorkbookModel(extended)
    with model._open_reader() as reader:
        assert reader.get_table('Sales')['Amount'].sum() == 100
    assert len(stored_relationships(model)) == 1


def test_spec_invalid_active_preserves_destination(tmp_path):
    pd.DataFrame({'ID':[1,2,3]}).to_csv(tmp_path/'data.csv', index=False)
    output=tmp_path/'out.xlsx'
    output.write_bytes(b'keep')
    spec={'tables':{'A':'data.csv','B':'data.csv'}, 'relationships':[{'from':'A.ID','to':'B.ID','active':'false'}]}
    with pytest.raises(TypeError, match='boolean'):
        build_from_spec(None, output, spec, base_directory=tmp_path)
    assert output.read_bytes() == b'keep'
    invalid=tmp_path/'invalid.json'
    invalid.write_text('[]')
    with pytest.raises(ValueError, match='mapping'):
        load_spec(invalid)

def test_parquet_model_preserves_nullable_types(tmp_path):
    pytest.importorskip('pyarrow')
    frame = pd.DataFrame({'ID':pd.Series([1,None,3],dtype='Int64'), 'Amount':[1.25,None,3.5],
                          'Flag':pd.Series([True,None,False],dtype='boolean')})
    frame.to_parquet(tmp_path/'data.parquet', index=False)
    output=build_from_spec(None,tmp_path/'parquet.xlsx',{'tables':{'T':'data.parquet'}},base_directory=tmp_path)
    model=WorkbookModel(output)
    with model._open_reader() as reader:
        actual=reader.get_table('T')
    assert actual.iloc[1].isna().all()
    assert actual['ID'].dropna().tolist() == [1,3]
    assert actual['Amount'].dropna().tolist() == [1.25,3.5]
    assert actual['Flag'].dropna().tolist() == [True,False]

def test_spec_new_metadata_and_ordered_edits(tmp_path):
    pd.DataFrame({'Amount':[10,20,30]}).to_csv(tmp_path/'sales.csv', index=False)
    spec = {'tables': {'Sales': 'sales.csv'},
        'measures': [{'table':'Sales','name':'Total','expression':'SUM(Sales[Amount])'}],
        'calculated_columns': [{'table':'Sales','name':'Double','expression':'Sales[Amount]*2','data_type':'integer'}],
        'kpis': [{'table':'Sales','measure':'Total','target_value':100,'thresholds':[40,80]}],
        'perspectives': [{'name':'Finance','columns':[['Sales','Amount']], 'measures':[['Sales','Total']]}],
        'operations': [{'op':'rename_table','table':'Sales','name':'Sales Data'},
                       {'op':'rename_perspective','name':'Finance','new_name':'Finance View'}]}
    output = build_from_spec(None, tmp_path/'model.xlsx', spec, base_directory=tmp_path)
    m = WorkbookModel(output)
    assert m.get_table('Sales Data')['Amount'].tolist() == [10,20,30]
    assert m.list_calculated_columns()[0].expression == "'Sales Data'[Amount]*2"
    assert m.list_perspectives() == ('Finance View',)
    graph = json.loads(subprocess.run([sys.executable,'-m','excel_model_lab.cli','dependencies',str(output)],
        check=True,capture_output=True,text=True).stdout)
    assert any(edge['reason'] == 'KPI measure' for edge in graph['edges'])
    report = subprocess.run([sys.executable,'-m','excel_model_lab.cli','validate',str(output)],capture_output=True,text=True)
    parsed = json.loads(report.stdout)
    assert report.returncode == (0 if parsed['valid'] else 1)
    assert 'issues' in parsed


def test_spec_metadata_only_template_and_binding_operations(tmp_path):
    from pathlib import Path
    source = Path(__file__).parent/'data'/'native-authoring-base.xlsx'
    output = build_from_spec(source,tmp_path/'edited.xlsx',{
        'tables':{'Table1':{'description':'Metadata only'}},
        'operations':[{'op':'bind_source','table':'Table1','source':{'kind':'linked_table','table_name':'Table1'}}]})
    assert WorkbookModel(output).list_native_sources()[0].kind == 'linked_table'
    for spec in ({'operations':[{'op':'save','destination':'unwanted.xlsx'}]}, {'unknown':True}):
        with pytest.raises(ValueError):
            build_from_spec(source,tmp_path/'edited.xlsx',spec)
        assert WorkbookModel(output).get_table('Table1')['ID'].tolist() == [1,2,3]
