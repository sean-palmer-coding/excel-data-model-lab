from pathlib import Path
import hashlib
import pandas as pd
import pytest
from pbixray import PBIXRay
from excel_model_lab import WorkbookModel, ModelReader

FIXTURE = Path(__file__).parent/'data'/'native-authoring-base.xlsx'


def test_reader_does_not_require_writer_support(monkeypatch):
    def forbidden(*args, **kwargs):
        raise AssertionError('Reader initialized the writer')
    monkeypatch.setattr(WorkbookModel, '__init__', forbidden)
    before = hashlib.sha256(FIXTURE.read_bytes()).digest()
    with WorkbookModel.open(FIXTURE) as reader, PBIXRay(FIXTURE) as reference:
        assert isinstance(reader, ModelReader)
        assert list(reader.tables) == list(reference.tables)
        for name in ('schema', 'relationships', 'dax_measures', 'statistics', 'mashup_queries'):
            pd.testing.assert_frame_equal(getattr(reader, name), getattr(reference, name))
        table = reader.tables[0]
        pd.testing.assert_frame_equal(reader.get_table(table), reference.get_table(table))
    assert hashlib.sha256(FIXTURE.read_bytes()).digest() == before


def test_stream_projection_and_resource_lifecycle(tmp_path):
    with WorkbookModel.open(FIXTURE, on_disk=True, temp_dir=tmp_path) as reader:
        table = reader.tables[0]
        whole = reader.get_table(table)
        column = whole.columns[0]
        chunks = list(reader.iter_table(table, columns=[column], chunk_size=2,
                                       strings_as_categorical=False))
        assert all(len(chunk) <= 2 for chunk in chunks)
        pd.testing.assert_frame_equal(pd.concat(chunks).reset_index(drop=True),
                                      whole[[column]].reset_index(drop=True))
        pending = reader.iter_table(table, chunk_size=1)
        next(pending)
    reader.close()
    with pytest.raises(RuntimeError, match='closed'):
        next(pending)
    with pytest.raises(RuntimeError, match='closed'):
        _ = reader.schema
    assert not list(tmp_path.iterdir())


def test_reader_exposes_upstream_metadata_contract():
    properties = {name for name, value in vars(PBIXRay).items() if isinstance(value, property)}
    assert properties <= set(vars(ModelReader))


def test_explicit_edit_mode():
    assert isinstance(WorkbookModel.open(FIXTURE, read_only=False), WorkbookModel)
    with pytest.raises(ValueError, match='read-only'):
        WorkbookModel.open(FIXTURE, read_only=False, on_disk=True)


def test_reader_dates_and_nullable_booleans_match_streaming(tmp_path):
    dates = pd.to_datetime(['2026-01-01 12:34:56', None, '2026-02-03 01:02:03'])
    frame = pd.DataFrame({'When': dates, 'Flag': pd.Series([True, None, False], dtype='boolean')})
    output = WorkbookModel.create({'Events': frame}).save(tmp_path / 'dates.xlsx')
    with ModelReader(output, on_disk=True, temp_dir=tmp_path) as reader:
        whole = reader.get_table('Events')
        chunks = list(reader.iter_table('Events', chunk_size=1, strings_as_categorical=False))
        streamed = pd.concat(chunks, ignore_index=True)
        pd.testing.assert_frame_equal(whole, streamed)
        assert list(whole.When.dropna()) == list(dates.dropna())
        assert whole.Flag.iloc[0] == True
        assert pd.isna(whole.Flag.iloc[1])
        assert whole.Flag.iloc[2] == False


def test_reader_initialization_failure_closes_disk_resources(tmp_path, monkeypatch):
    from excel_model_lab import reader_adapter
    def fail(reader):
        raise ValueError('normalization failed')
    monkeypatch.setattr(reader_adapter, 'normalize_reader_types', fail)
    with pytest.raises(ValueError, match='normalization failed'):
        ModelReader(FIXTURE, on_disk=True, temp_dir=tmp_path)
    assert not list(tmp_path.iterdir())

def test_processed_calculated_inferred_types_override_empty_key_schema(monkeypatch):
    from types import SimpleNamespace
    import pandas as pd
    from excel_model_lab import reader_adapter
    schema=pd.DataFrame([
        {'TableName':'T','ColumnName':name,'SemanticType':'Empty','PandasDataType':'object','Dictionary':'dictionary','Magnitude':1}
        for name in ('Date','Flag','Integer','Money','Pending')])
    attrs=''.join(f'<Attribute><ID>{name}</ID><Name>{name}</Name><InferredDatatype>{kind}</InferredDatatype><ProcessingState>{state}</ProcessingState><KeyColumns><KeyColumn><DataType>Empty</DataType><Source><Expression>1</Expression></Source></KeyColumn></KeyColumns></Attribute>'
                  for name,kind,state in [('Date','Date','Processed'),('Flag','Boolean','Processed'),('Integer','BigInt','Processed'),('Money','Currency','Processed'),('Pending','Date','Unprocessed')])
    payload=f'<Load><ObjectDefinition><Dimension><ID>T</ID><Name>T</Name><Attributes>{attrs}</Attributes></Dimension></ObjectDefinition></Load>'.encode()
    decoder=SimpleNamespace(_meta=SimpleNamespace(schema_df=schema),_codes_to_series=lambda *args:None)
    reader=SimpleNamespace(_vertipaq_decoder=decoder)
    monkeypatch.setattr(reader_adapter,'excel_metadata_files',lambda reader:[('T.dim.xml',payload)])
    reader_adapter.normalize_reader_types(reader)
    assert schema.SemanticType.tolist()==['Date','Boolean','BigInt','Currency','Empty']
    assert schema.PandasDataType.tolist()==['datetime64[ns]','boolean','Int64','object','object']
