import pandas as pd
from xml.dom import minidom
from excel_model_lab import WorkbookModel, ModelReader
from excel_model_lab.schema import _elements, _child, _text

def test_stored_label_calculated_sort_defers_order_and_dependent_hierarchy(tmp_path):
    m=WorkbookModel.create({'Sales':pd.DataFrame({'Amount':[40,10,30], 'Band':['High','Low','High']})})
    m.add_calculated_column('Sales','Rank','IF([Amount]<20,1,2)',data_type='integer')
    m.add_hierarchy('Sales','Detail',['Band','Amount'])
    m.set_sort_column('Sales','Band','Rank')
    first=m.save(tmp_path/'sorted.xlsx')
    m=WorkbookModel(first)
    dimension=minidom.parseString(m._image.read_logical_file(m._dimension_paths['Sales']))
    band=next(a for a in _elements(dimension,'Attribute') if _text(_child(a,'ID'))=='Band')
    assert _text(_child(band,'ProcessingState'))=='Unprocessed'
    assert _text(_child(band,'AttributeHierarchyProcessingState'))=='Unprocessed'
    assert _text(_child(band,'OrderByAttributeID'))=='Rank'
    hierarchy=next(h for h in _elements(dimension,'Hierarchy') if _text(_child(h,'ID'))=='Detail')
    assert _text(_child(hierarchy,'ProcessingState'))=='Unprocessed'
    assert not any(f.logical_path and 'U$' in f.logical_path and '$Detail.' in f.logical_path for f in m._image.files)
    m.replace_column('Sales','Amount',[5,10,15])
    m.replace_column('Sales','Band',['A','B','C'])
    second=m.save(tmp_path/'changed.xlsx')
    with ModelReader(second) as reader:
        frame=reader.get_table('Sales')
        assert frame.Amount.tolist()==[5,10,15]
        assert frame.Band.tolist()==['A','B','C']
        assert frame.Rank.isna().all()
