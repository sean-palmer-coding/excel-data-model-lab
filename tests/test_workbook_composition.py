from zipfile import ZipFile
from xml.etree import ElementTree as ET
import pandas as pd
import pytest

pytest.importorskip('xlsxwriter')
from excel_model_lab.composition import compose_workbook, attach_model
from excel_model_lab import ModelReader
from excel_model_lab.reports import PivotTableSpec, PivotChartSpec

S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def test_composition_preserves_worksheet_assets_and_existing_drawings(tmp_path):
    original = {}
    def worksheets(writer, cubes):
        sheet = writer.add_worksheet('Dashboard')
        inputs = writer.add_worksheet('Inputs')
        inputs.add_table('A1:B3', {'name': 'InputTable', 'columns': [{'header':'Name'}, {'header':'Amount'}],
                                  'data': [['=literal', 10], ['Beta', 20]]})
        inputs.conditional_format('B2:B3', {'type':'data_bar'})
        writer.define_name('Amounts', '=Inputs!$B$2:$B$3')
        fmt = writer.add_format({'bold':True, 'num_format':'$#,##0.00'})
        sheet.write_formula('B2', '=SUM(Amounts)', fmt, '#N/A')
        for target in (sheet, inputs):
            chart = writer.add_chart({'type':'column'})
            chart.add_series({'values':'=Inputs!$B$2:$B$3'})
            target.insert_chart('D3', chart)
        cubes.write(sheet, 'B3', cubes.value(cubes.measure('Total')))
    def reports(report):
        with ZipFile(report.model.source) as z:
            original.update({n:z.read(n) for n in z.namelist()})
        report.add_pivot_table(PivotTableSpec('Summary','Report',rows=(('Sales','Name'),),measures=('Total',)))
        report.add_pivot_chart(PivotChartSpec('PivotChart','Summary',sheet='Dashboard',cell='D20'))
    path = compose_workbook(tmp_path/'dashboard.xlsx', {'Sales':pd.DataFrame({'Name':['A','B'],'Amount':[10,20]})},
        worksheets, configure_model=lambda m:m.add_measure('Sales','Total','SUM(Sales[Amount])'), configure_reports=reports)
    with ZipFile(path) as z:
        assert z.testzip() is None
        relationships = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        model_links = [r for r in relationships if r.get('Target') == 'model/item.data']
        assert len(model_links) == 1
        assert model_links[0].get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/powerPivotData'
        for name in ('xl/styles.xml','xl/sharedStrings.xml','xl/tables/table1.xml','xl/charts/chart1.xml','xl/charts/chart2.xml','xl/worksheets/sheet2.xml','xl/drawings/drawing2.xml'):
            assert z.read(name) == original[name], name
        workbook = ET.fromstring(z.read('xl/workbook.xml'))
        assert workbook.find(S+'definedNames')[0].text == 'Inputs!$B$2:$B$3'
        assert len([n for n in z.namelist() if n.startswith('xl/charts/chart') and n.endswith('.xml')]) == 3
        assert '[dashboard.xlsx]' in z.read('xl/charts/chart3.xml').decode()
        sheet = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        assert sheet.find(f'.//{S}c[@r="B3"]/{S}v').text == '#N/A'
        assert '=literal' in z.read('xl/sharedStrings.xml').decode()
        inputs = ET.fromstring(z.read('xl/worksheets/sheet2.xml'))
        assert inputs.find(f'.//{S}c[@r="A2"]/{S}f') is None
    with ModelReader(path) as reader:
        assert reader.get_table('Sales')['Amount'].sum() == 30


@pytest.mark.parametrize('failure_stage', ['worksheets','model','reports'])
def test_callback_failure_is_atomic(tmp_path, failure_stage):
    destination = tmp_path/'keep.xlsx'
    destination.write_bytes(b'original')
    def fail(*args): raise RuntimeError('stop')
    callbacks = dict(write_worksheets=lambda w,c:w.add_worksheet('Base'),
                     configure_model=None,configure_reports=None)
    callbacks[{'worksheets':'write_worksheets','model':'configure_model','reports':'configure_reports'}[failure_stage]] = fail
    with pytest.raises(RuntimeError,match='stop'):
        compose_workbook(destination, {'T':pd.DataFrame({'ID':[1]})}, **callbacks)
    assert destination.read_bytes() == b'original'
    assert not list(tmp_path.glob('.excelmodel-compose-*'))


def test_attach_preserves_existing_connections_and_allocates_model_id(tmp_path):
    import xlsxwriter
    source = tmp_path/'base.xlsx'
    with xlsxwriter.Workbook(source) as w:
        w.add_worksheet('Input').write('A1','Keep me')
    with ZipFile(source) as z:
        parts = {n:z.read(n) for n in z.namelist()}
    parts['xl/connections.xml'] = b'<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="17" name="Existing" type="5" refreshedVersion="6"><dbPr connection="Provider=Example;" command="existing"/><extLst><ext uri="vendor-preserve"/></extLst></connection></connections>'
    rns = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    rels = ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
    ET.SubElement(rels,rns+'Relationship',Id='rId99',Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections',Target='connections.xml')
    parts['xl/_rels/workbook.xml.rels'] = ET.tostring(rels)
    with ZipFile(source,'w') as z:
        for n,data in parts.items(): z.writestr(n,data)
    path = attach_model(source,tmp_path/'attached.xlsx',{'T':pd.DataFrame({'ID':[1,2]})},
        configure_model=lambda m:m.add_measure('T','Total','SUM(T[ID])'),
        configure_reports=lambda r:r.add_pivot_table(PivotTableSpec('Summary','Report',measures=('Total',))))
    with ZipFile(path) as z:
        connections = ET.fromstring(z.read('xl/connections.xml'))
        assert [(c.get('id'),c.get('name')) for c in connections] == [('17','Existing'),('18','ThisWorkbookDataModel')]
        assert connections[0].find(f'{S}extLst/{S}ext').get('uri') == 'vendor-preserve'
        assert z.read('xl/worksheets/sheet1.xml') == parts['xl/worksheets/sheet1.xml']
        cache = ET.fromstring(z.read('xl/pivotCache/pivotCacheDefinition1.xml'))
        assert cache.find(S+'cacheSource').get('connectionId') == '18'


def test_attach_rejects_existing_model_without_overwrite(tmp_path):
    source = compose_workbook(tmp_path/'base.xlsx',{'T':pd.DataFrame({'ID':[1]})},lambda w,c:w.add_worksheet('Input'))
    before = source.read_bytes()
    with pytest.raises(ValueError,match='already contains a model'):
        attach_model(source,source,{'T':pd.DataFrame({'ID':[2]})})
    assert source.read_bytes() == before


def test_attach_rejects_reserved_connection_name_atomically(tmp_path):
    import xlsxwriter
    source = tmp_path/'base.xlsx'
    with xlsxwriter.Workbook(source) as w:
        w.add_worksheet('Input')
    with ZipFile(source) as z:
        parts = {n:z.read(n) for n in z.namelist()}
    parts['xl/connections.xml'] = b'<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="1" name="THISWORKBOOKDATAMODEL"/></connections>'
    with ZipFile(source,'w') as z:
        for n,data in parts.items(): z.writestr(n,data)
    destination = tmp_path/'keep.xlsx'
    destination.write_bytes(b'keep')
    with pytest.raises(ValueError,match='conflicting Data Model connection'):
        attach_model(source,destination,{'T':pd.DataFrame({'ID':[1]})})
    assert destination.read_bytes() == b'keep'


@pytest.mark.parametrize('kind,target', [
    ('http://schemas.microsoft.com/office/2007/relationships/model', 'model/item.data'),
    ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/powerPivotData', 'model/other.data'),
])
def test_attach_rejects_invalid_model_relationships_atomically(tmp_path, kind, target):
    import xlsxwriter
    source = tmp_path/'base.xlsx'
    with xlsxwriter.Workbook(source) as w:
        w.add_worksheet('Input')
    with ZipFile(source) as z:
        parts = {n:z.read(n) for n in z.namelist()}
    rns = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    rels = ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
    ET.SubElement(rels,rns+'Relationship',Id='rId99',Type=kind,Target=target)
    parts['xl/_rels/workbook.xml.rels'] = ET.tostring(rels)
    with ZipFile(source,'w') as z:
        for n,data in parts.items(): z.writestr(n,data)
    destination = tmp_path/'keep.xlsx'
    destination.write_bytes(b'keep')
    with pytest.raises(ValueError,match='model relationship|Data Model relationship'):
        attach_model(source,destination,{'T':pd.DataFrame({'ID':[1]})})
    assert destination.read_bytes() == b'keep'
