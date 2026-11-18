from pathlib import Path
import pytest
from excel_model_lab import ModelReader
from excel_model_lab.excel_metadata import ExcelMetadata


def test_excel_assl_metadata_snapshot():
    dim = b'''<Load xmlns="urn:test"><ObjectDefinition><Dimension>
      <ID>calendar-id</ID><Name>Calendar</Name><Type>Time</Type><Description>Dates</Description>
      <Attributes>
        <Attribute><ID>row</ID><Type>RowNumber</Type></Attribute>
        <Attribute><ID>d</ID><Name>Date</Name><Usage>Key</Usage>
          <KeyColumns><KeyColumn><DataType>Date</DataType></KeyColumn></KeyColumns></Attribute>
        <Attribute><ID>y</ID><Name>Year</Name><InferredDatatype>BigInt</InferredDatatype>
          <ProcessingState>Unprocessed</ProcessingState><AttributeHierarchyVisible>false</AttributeHierarchyVisible>
          <AttributeHierarchyDisplayFolder>Calendar fields</AttributeHierarchyDisplayFolder>
          <OrderByAttributeID>d</OrderByAttributeID><FormatString>0</FormatString>
          <KeyColumns><KeyColumn><DataType>Empty</DataType><Source><Expression>YEAR([Date])</Expression></Source></KeyColumn></KeyColumns>
          <Annotations><Annotation><Name>Format</Name><Value><Format Format="General"/></Value></Annotation></Annotations>
        </Attribute>
      </Attributes><Hierarchies><Hierarchy><ID>h</ID><Name>Calendar</Name><Levels>
        <Level><ID>l</ID><Name>Year</Name><SourceAttributeID>y</SourceAttributeID></Level>
      </Levels></Hierarchy></Hierarchies>
    </Dimension></ObjectDefinition></Load>'''
    cube = b'''<Load><ObjectDefinition><Cube><Dimensions><Dimension><ID>calendar-id</ID><Visible>false</Visible></Dimension></Dimensions>
      <Annotations><Annotation><Name>PerspectiveEmptyHierarchyNames</Name><Value>&lt;Perspectives&gt;&lt;Perspective Name="Summary"&gt;&lt;Tables&gt;&lt;Table Name="Calendar" EmptyHierarchyNames="Calendar"/&gt;&lt;/Tables&gt;&lt;/Perspective&gt;&lt;/Perspectives&gt;</Value></Annotation></Annotations>
    </Cube></ObjectDefinition></Load>'''
    perspective = b'''<Load><ObjectDefinition><Perspective><ID>p</ID><Name>Summary</Name>
       <Dimensions><Dimension><CubeDimensionID>calendar-id</CubeDimensionID><Attributes><Attribute><AttributeID>y</AttributeID></Attribute></Attributes>
       <Hierarchies><Hierarchy><HierarchyID>h</HierarchyID></Hierarchy></Hierarchies></Dimension></Dimensions>
       <Calculations><Calculation><Name>[Measures].[Total]</Name><Type>Member</Type></Calculation></Calculations>
       <MeasureGroups><MeasureGroup><MeasureGroupID>sales-id</MeasureGroupID></MeasureGroup></MeasureGroups>
    </Perspective></ObjectDefinition></Load>'''
    m = ExcelMetadata.from_files([('a.dim.xml', dim), ('a.persp.xml', perspective), ('a.cub.xml', cube)])
    assert m.available
    assert m.date_tables.iloc[0]['date_key_columns'] == ('d',)
    assert m.tables.iloc[0]['hidden']
    assert len(m.columns) == 2
    calc = m.calculated_columns.iloc[0]
    assert (calc['expression'], calc['data_type'], calc['processing_state']) == ('YEAR([Date])', 'BigInt', 'Unprocessed')
    assert calc['display_folder'] == 'Calendar fields'
    assert calc['sort_by_column_id'] == 'd'
    assert 'General' in calc['annotations']['Format']
    assert m.levels.iloc[0]['column_id'] == 'y'
    assert m.perspectives.iloc[0]['measure_group_ids'] == ('sales-id',)
    assert m.perspective_columns.iloc[0]['column_id'] == 'y'
    assert m.perspective_measures.iloc[0]['name'] == '[Measures].[Total]'
    assert m.perspective_hierarchies.iloc[0]['hierarchy_id'] == 'h'
    assert m.perspective_hierarchy_annotations.iloc[0]['empty_hierarchy_names'] == 'Calendar'


def test_empty_metadata_is_explicitly_unavailable():
    m = ExcelMetadata.from_files([])
    assert not m.available
    assert m.date_tables.empty and m.calculated_columns.empty
    assert list(m.date_tables.columns) == list(m.tables.columns)


def test_reader_reuses_loaded_model_and_enforces_lifecycle(tmp_path, monkeypatch):
    import excel_model_lab.reader as reader_module
    fixture = Path(__file__).parent / 'data' / 'native-authoring-base.xlsx'
    with ModelReader(fixture, on_disk=True, temp_dir=tmp_path) as reader:
        def forbidden(*args, **kwargs):
            raise AssertionError('Metadata must not reopen the model')
        monkeypatch.setattr(reader_module, 'PBIXRay', forbidden)
        m = reader.excel_metadata
        assert m.available and m.tables.iloc[0]['name'] == 'Table1'
        m.tables.loc[0, 'name'] = 'Changed snapshot'
        assert reader.excel_metadata.tables.iloc[0]['name'] == 'Table1'
        assert list(reader.tables) == ['Table1']
    with pytest.raises(RuntimeError, match='closed'):
        _ = reader.excel_metadata
    assert not list(tmp_path.iterdir())
