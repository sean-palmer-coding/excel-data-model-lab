"""Reject native-incompatible stored empty text without silently changing it."""
from io import BytesIO
from pathlib import Path
import pandas as pd
import pytest
from excel_model_lab import WorkbookModel, ModelReader
from excel_model_lab.abf import UnsupportedModelError
from excel_model_lab.dictionary import write_uncompressed_string_dictionary


@pytest.mark.parametrize('values',[['A','',None],['A','','C'],['']])
def test_create_rejects_empty_stored_text(values):
    with pytest.raises(UnsupportedModelError,match='Empty stored text.*None explicitly'):
        WorkbookModel.create({'T':pd.DataFrame({'Text':values})})


def test_replace_empty_text_is_atomic_and_explicit_null_remains_supported(tmp_path):
    model=WorkbookModel.create({'T':pd.DataFrame({'Text':['A','B','C']})})
    original=model.get_table('T').copy()
    with pytest.raises(UnsupportedModelError,match='position 1'):
        model.replace_column('T','Text',['A','','C'])
    pd.testing.assert_frame_equal(model.get_table('T'),original)
    output=model.replace_column('T','Text',['A',None,'C']).save(tmp_path/'null.xlsx')
    with ModelReader(output) as reader:
        values=reader.get_table('T')['Text'].tolist()
        assert values[0]=='A' and pd.isna(values[1]) and values[2]=='C'


def test_dictionary_empty_string_rejected_before_writing_bytes():
    seed=WorkbookModel(Path(__file__).parents[1]/'src/excel_model_lab/data/authoring-seed.xlsx')
    template=seed._image.read_logical_file(seed._encodings['Table1','Label'].dictionary_path)
    sink=BytesIO(b'keep');sink.seek(4)
    with pytest.raises(UnsupportedModelError,match='dictionary position 1'):
        write_uncompressed_string_dictionary(sink,template,['A',''])
    assert sink.getvalue()==b'keep'


def test_empty_string_dax_expression_is_not_rejected(tmp_path):
    model=WorkbookModel.create({'T':pd.DataFrame({'Text':['A','B','C']})})
    model.add_measure('T','Empty Text','""')
    model.add_calculated_column('T','Empty Calculated','""',data_type='text')
    assert model.save(tmp_path/'dax-empty.xlsx').is_file()
