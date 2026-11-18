import struct
import pandas as pd
import pytest
from excel_model_lab import WorkbookModel, ModelReader
from excel_model_lab.schema import _elements, _child, _text
from xml.dom import minidom

@pytest.mark.parametrize('dtype,expression', [('integer','[ID]*2'),('text','FORMAT([ID],"0")'),('boolean','[ID]>1')])
def test_empty_calculated_cache_has_no_phantom_blank_run_and_repopulates(tmp_path,dtype,expression):
    model=WorkbookModel.create({'T':pd.DataFrame({'ID':[1,2]})})
    model.add_calculated_column('T','Calculated',expression,data_type=dtype)
    model.replace_table('T',pd.DataFrame({'ID':pd.Series([],dtype='int64')}))
    path=model.save(tmp_path/'empty.xlsx')
    reopened=WorkbookModel(path)
    with ModelReader(path) as reader:
        assert reader.get_table('T').empty
        assert 'Calculated' in reader.get_table('T')
    image=reopened._image
    stream=next(f for f in image.files if f.logical_path and '.Calculated.' in f.logical_path and f.logical_path.endswith('.idf'))
    assert image.read_logical_file(stream)==struct.pack('<Q',16)+bytes(128)+struct.pack('<QQ',1,0)
    rawfile=next(f for f in image.files if f.logical_path and f.logical_path.endswith('.tbl.xml'))
    doc=minidom.parseString(image.read_logical_file(rawfile))
    raw=next(n for n in _elements(doc,'XMObject') if n.getAttribute('class')=='XMRawColumn' and n.getAttribute('name')=='Calculated')
    assert all(_text(n)=='0' for tag in ['RLERuns','StorageUsedSize'] for n in _elements(raw,tag))
    assert all(_text(n)=='false' for n in _elements(raw,'HasNulls'))
    reopened.replace_table('T',pd.DataFrame({'ID':[7,8]}))
    populated=reopened.save(tmp_path/'populated.xlsx')
    with ModelReader(populated) as reader:
        frame=reader.get_table('T')
        assert frame.ID.tolist()==[7,8]
        assert frame.Calculated.isna().all()
