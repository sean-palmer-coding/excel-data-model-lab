import pandas as pd
import pytest

from excel_model_lab import WorkbookModel
from excel_model_lab.dependencies import build_dependency_graph, DependencyError
from excel_model_lab.errors import ModelValidationError


def model():
    return WorkbookModel.create({'Sales': pd.DataFrame({'Amount': [10, 20], 'Unused': [1, 2]}),
                                 'Other': pd.DataFrame({'Amount': [30, 40], 'ID': [1, 2]})})


def test_qualified_references_do_not_block_unrelated_same_named_column():
    m = model()
    m.add_measure('Sales', 'Total', 'SUM(Other[Amount]) + LEN("Sales[Amount]") /* Sales[Amount] */')
    graph = build_dependency_graph(m)
    graph.assert_removable(('tables', 'Sales', 'columns', 'Amount'))
    with pytest.raises(DependencyError) as error:
        graph.assert_removable(('tables', 'Other', 'columns', 'Amount'))
    assert isinstance(error.value, ModelValidationError)
    assert error.value.issues[0].object_path == ('tables', 'Other', 'columns', 'Amount')
    assert error.value.issues[0].dependencies == (('tables', 'Sales', 'measures', 'Total'),)


def test_transitive_measure_dependencies_and_table_references():
    m = model()
    m.add_measure('Sales', 'First', 'SUM(Sales[Amount])')
    m.add_measure('Other', 'Second', '[First] + COUNTROWS(Sales)')
    graph = build_dependency_graph(m)
    dependents = graph.dependents(('tables', 'Sales', 'columns', 'Amount'), transitive=True)
    assert {edge.source[-1] for edge in dependents} == {'First', 'Second'}
    with pytest.raises(DependencyError):
        graph.assert_removable(('tables', 'Sales'))


def test_hierarchy_sort_and_staged_relationships_are_dependencies():
    m = model()
    m.add_hierarchy('Sales', 'Details', ['Amount', 'Unused'])
    m.set_sort_column('Other', 'Amount', 'ID')
    m.add_relationship('Sales', 'Unused', 'Other', 'ID')
    graph = build_dependency_graph(m)
    assert 'hierarchy level' in {edge.reason for edge in graph.dependents(('tables', 'Sales', 'columns', 'Amount'))}
    reasons = {edge.reason for edge in graph.dependents(('tables', 'Other', 'columns', 'ID'))}
    assert {'sort by column', 'relationship key'} <= reasons


def test_graph_is_snapshot_and_unresolved_matching_fields_are_restricted():
    m = model()
    original = build_dependency_graph(m)
    m.add_measure('Sales', 'Mystery', '[Missing]')
    graph = build_dependency_graph(m)
    assert not original.unresolved
    assert graph.unresolved[0].name == 'Missing'
    graph.assert_removable(('tables', 'Sales', 'columns', 'Unused'))


def test_failed_removal_leaves_model_unchanged():
    m = model()
    m.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')
    with pytest.raises(DependencyError):
        m.remove_column('Sales', 'Amount')
    assert m.get_table('Sales')['Amount'].tolist() == [10, 20]
    assert len(m._pending_measures) == 1


def test_remove_independent_table_preserves_other_pending_calculations(tmp_path):
    m = model()
    m.add_measure('Sales', 'Total', 'SUM(Sales[Amount])')
    m.remove_table('Other')
    path = m.save(tmp_path / 'independent.xlsx')
    graph = WorkbookModel(path).dependency_graph()
    assert ('tables', 'Sales', 'measures', 'Total') in {obj.path for obj in graph.objects}


def test_remove_column_qualified_other_table_measure_after_save(tmp_path):
    m = model()
    m.add_measure('Sales', 'Total', 'SUM(Other[Amount])')
    path = m.save(tmp_path / 'qualified.xlsx')
    m = WorkbookModel(path)
    m.remove_column('Sales', 'Amount')
    assert list(m.get_table('Sales')) == ['Unused']


def test_transaction_commits_together_or_rolls_back():
    m = model()
    with pytest.raises(KeyError):
        with m.transaction() as edit:
            edit.replace_column('Sales', 'Amount', [100, 200])
            edit.remove_table('Missing')
    assert m.get_table('Sales')['Amount'].tolist() == [10, 20]
    with m.transaction() as edit:
        edit.replace_column('Sales', 'Amount', [100, 200])
        edit.remove_column('Sales', 'Unused')
    assert m.get_table('Sales').to_dict('list') == {'Amount': [100, 200]}


def test_table_rename_rebinds_stored_and_pending_dax_and_preserves_identity(tmp_path):
    from excel_model_lab.measures import list_script_measures
    m = model()
    m.add_measure('Sales', 'Total', 'SUM(Sales[Amount]) + LEN("Sales[Amount]") /* Sales */')
    m.add_measure('Other', 'Rows', 'COUNTROWS(Sales)')
    m.add_measure('Sales', 'Target', '100')
    path = m.save(tmp_path / 'before-rename.xlsx')
    m = WorkbookModel(path)
    m.add_kpi('Sales', 'Total', target_measure='Target')
    m.add_hierarchy('Sales', 'Details', ['Unused'])
    m.create_perspective('Finance', columns=[('Sales', 'Amount')],
                         measures=[('Sales', 'Total')], hierarchies=[('Sales', 'Details')])
    identity = m._dimension_ids['Sales']
    m.add_measure('Other', 'Pending', 'SUM(Sales[Unused])')
    m.rename_table('Sales', "New 'Sales'")
    assert m._dimension_ids["New 'Sales'"] == identity
    output = m.save(tmp_path / 'renamed.xlsx')
    reopened = WorkbookModel(output)
    script = next(f for f in reopened._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
    measures = {measure.name: measure for measure in list_script_measures(reopened._image.read_logical_file(script))}
    assert measures['Total'].table == "New 'Sales'"
    assert measures['Total'].expression == "SUM('New ''Sales'''[Amount]) + LEN(\"Sales[Amount]\") /* Sales */"
    assert measures['Rows'].expression == "COUNTROWS('New ''Sales''')"
    assert measures['Pending'].expression == "SUM('New ''Sales'''[Unused])"
    assert reopened.list_kpis()[0]['target_table'] == "New 'Sales'"
    assert ('perspectives', 'Finance') in {edge.source for edge in reopened.dependency_graph().dependents(('tables', "New 'Sales'", 'hierarchies', 'Details'))}


def test_column_rename_rebinds_only_resolved_references(tmp_path):
    from excel_model_lab.measures import list_script_measures
    m = model()
    m.add_measure('Sales', 'Total', 'SUM(Sales[Amount]) + SUM([Amount]) + LEN("[Amount]")')
    m.add_measure('Other', 'OtherTotal', 'SUM(Other[Amount])')
    m = WorkbookModel(m.save(tmp_path / 'column-before.xlsx'))
    m.rename_column('Sales', 'Amount', 'New Amount')
    image = WorkbookModel(m.save(tmp_path / 'column-renamed.xlsx'))._image
    script = next(file for file in image.files if file.logical_path and file.logical_path.endswith('.scr.xml'))
    measures = {m.name: m.expression for m in list_script_measures(image.read_logical_file(script))}
    assert measures['Total'] == 'SUM(Sales[New Amount]) + SUM([New Amount]) + LEN("[Amount]")'
    assert measures['OtherTotal'] == 'SUM(Other[Amount])'


def test_column_rename_restricts_non_dax_dependencies_atomically():
    m = model()
    m.add_hierarchy('Sales', 'Detail', ['Amount'])
    before = m._image.data
    with pytest.raises(DependencyError):
        m.rename_column('Sales', 'Amount', 'New Amount')
    assert m._image.data == before
    assert list(m.get_table('Sales')) == ['Amount', 'Unused']

def test_perspective_dependency_uses_physical_column_identity():
    from xml.dom import minidom
    from excel_model_lab.abf import AbfImage
    from excel_model_lab.schema import _elements, _child, _text
    m = model()
    m.create_perspective('Finance', columns=[('Sales', 'Amount')])
    file = m._dimension_paths['Sales']
    document = minidom.parseString(m._image.read_logical_file(file))
    attribute = next(a for a in _elements(document, 'Attribute') if _text(_child(a, 'ID')) == 'Amount')
    caption = _child(attribute, 'Name')
    caption.firstChild.data = 'Display Amount'
    storage = file
    m._image = AbfImage(m._image.assemble({storage: document.toxml(encoding='utf-8')}))
    graph = m.dependency_graph()
    assert any(edge.source == ('perspectives', 'Finance') for edge in
               graph.dependents(('tables', 'Sales', 'columns', 'Display Amount')))

@pytest.mark.parametrize('operation', ['remove', 'rename', 'update'])
@pytest.mark.parametrize('part', ['xl/pivotTables/pivotTable1.xml', 'xl/slicerCaches/slicerCache1.xml', 'xl/timelineCaches/timelineCache1.xml'])
def test_perspective_report_guard_uses_canonical_name(tmp_path, operation, part):
    from zipfile import ZipFile
    m = model()
    m.create_perspective('Finance', columns=[('Sales', 'Amount')])
    base = m.save(tmp_path / 'base.xlsx')
    source = tmp_path / 'report.xlsx'
    with ZipFile(base) as before, ZipFile(source, 'w') as after:
        for item in before.infolist():
            after.writestr(item, before.read(item))
        after.writestr(part, '<reportBinding/>')
    m = WorkbookModel(source)
    with pytest.raises(DependencyError) as failure:
        if operation == 'remove':
            m.remove_perspective('finance')
        elif operation == 'rename':
            m.rename_perspective('finance', 'Other')
        else:
            m.update_perspective('finance', columns=[])
    assert failure.value.issues[0].object_path == ('perspectives', 'Finance')
    assert m.list_perspectives() == ('Finance',)

@pytest.mark.parametrize('formula,blocked', [('CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total]")',True), ('"CUBEVALUE("',False)])
def test_cube_formula_report_guard_ignores_literal_text(tmp_path, formula, blocked):
    from zipfile import ZipFile
    from xml.etree import ElementTree as ET
    from excel_model_lab.abf import UnsupportedModelError
    m = model()
    base = m.save(tmp_path/'base.xlsx')
    source = tmp_path/'cube.xlsx'
    with ZipFile(base) as before, ZipFile(source,'w') as after:
        for item in before.infolist():
            payload = before.read(item)
            if item.filename == 'xl/worksheets/sheet1.xml':
                sheet = ET.fromstring(payload)
                ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
                row = ET.SubElement(sheet.find(ns+'sheetData'),ns+'row',{'r':'99'})
                cell = ET.SubElement(row,ns+'c',{'r':'A99'})
                ET.SubElement(cell,ns+'f').text = formula
                payload = ET.tostring(sheet)
            after.writestr(item,payload)
    m = WorkbookModel(source)
    report_edges = [e for e in m.dependency_graph().edges if e.source[0]=='workbook_parts']
    assert bool(report_edges) == blocked
    if blocked:
        with pytest.raises(UnsupportedModelError):
            m.rename_table('Sales','Renamed')
        with pytest.raises(DependencyError):
            m.remove_column('Sales','Unused')
    else:
        m.rename_table('Sales','Renamed')

def test_kpi_removal_rejects_cube_report_binding(tmp_path):
    from zipfile import ZipFile
    m = model()
    m.add_measure('Sales','Total','SUM(Sales[Amount])')
    m.add_kpi('Sales','Total',target_value=100)
    base = m.save(tmp_path/'base.xlsx')
    source = tmp_path/'report.xlsx'
    with ZipFile(base) as before, ZipFile(source,'w') as after:
        for item in before.infolist():
            after.writestr(item,before.read(item))
        after.writestr('xl/worksheets/kpi-report.xml', '<worksheet><f>CUBEKPIMEMBER("ThisWorkbookDataModel","Total",1)</f></worksheet>')
    m = WorkbookModel(source)
    with pytest.raises(DependencyError):
        m.remove_kpi('Sales','Total')
    assert m.list_kpis()[0]['measure'] == 'Total'

def test_measure_rename_rebinds_stored_pending_and_calculated_references(tmp_path):
    from excel_model_lab.measures import list_script_measures
    m = model()
    m.add_measure('Sales','Total','SUM(Sales[Amount])')
    m.add_measure('Other','Saved','[Total] + Sales[Total] + LEN("[Total]") /* [Total] */')
    m.add_calculated_column('Sales','Calc','[Total] + 1',data_type='integer')
    first = m.save(tmp_path/'before.xlsx')
    m = WorkbookModel(first)
    m.add_measure('Sales','Pending','[Total] + 2')
    m.rename_measure('sales','total','Net Total')
    output = m.save(tmp_path/'after.xlsx')
    reopened = WorkbookModel(output)
    script = next(f for f in reopened._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
    measures = {measure.name:measure.expression for measure in list_script_measures(reopened._image.read_logical_file(script))}
    assert 'Total' not in measures
    assert measures['Net Total'] == 'SUM(Sales[Amount])'
    assert measures['Saved'] == '[Net Total] + Sales[Net Total] + LEN("[Total]") /* [Total] */'
    assert measures['Pending'] == '[Net Total] + 2'
    assert reopened.list_calculated_columns()[0].expression == '[Net Total] + 1'
    assert reopened.list_calculated_columns()[0].processing_state == 'Unprocessed'
    assert reopened.get_table('Sales')['Amount'].tolist() == [10,20]


@pytest.mark.parametrize('dependency', ['kpi','perspective','ambiguous','virtual','conflict','report'])
def test_measure_rename_rejects_unsupported_dependencies_atomically(dependency, tmp_path):
    m = model()
    m.add_measure('Sales','Total','SUM(Sales[Amount])')
    if dependency == 'kpi':
        m.add_kpi('Sales','Total',target_value=100)
    elif dependency == 'perspective':
        m.create_perspective('Finance',measures=[('Sales','Total')])
    elif dependency == 'ambiguous':
        m.add_column('Other','Total',[1,2])
        m.add_measure('Other','Uses','[Total]')
    elif dependency == 'virtual':
        m.add_measure('Other','Uses','SUMX(ADDCOLUMNS(Sales,"Total",1),[Total])')
    elif dependency == 'report':
        from zipfile import ZipFile
        base = m.save(tmp_path/'base.xlsx')
        source = tmp_path/'report.xlsx'
        with ZipFile(base) as before, ZipFile(source,'w') as after:
            for item in before.infolist():
                after.writestr(item,before.read(item))
            after.writestr('xl/worksheets/cube.xml', '<worksheet><f>CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total]")</f></worksheet>')
        m = WorkbookModel(source)
    else:
        m.add_measure('Other','Renamed','1')
    before, pending = m._image.data, list(m._pending_measures)
    with pytest.raises(DependencyError) as failure:
        m.rename_measure('Sales','Total','Renamed')
    assert failure.value.issues[0].object_path == ('tables','Sales','measures','Total')
    assert failure.value.issues[0].dependencies
    assert m._image.data == before and m._pending_measures == pending

def test_measure_rename_preserves_other_table_column_and_metadata(tmp_path):
    from excel_model_lab.measures import list_script_measures
    from xml.etree import ElementTree as ET
    m = model()
    m.add_column('Other','Total',[1,2])
    m.add_measure('Sales','Total','SUM(Sales[Amount])',format_string='0.00',description='Keep description',display_folder='Finance')
    m.add_measure('Other','Only Column','SUM(Other[Total])')
    m.rename_measure('Sales','Total','Net]Total')
    output = m.save(tmp_path/'renamed.xlsx')
    reopened = WorkbookModel(output)
    script = next(f for f in reopened._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
    data = reopened._image.read_logical_file(script)
    values = {m.name:m.expression for m in list_script_measures(data)}
    assert values['Only Column'] == 'SUM(Other[Total])'
    assert values['Net]Total'] == 'SUM(Sales[Amount])'
    ns='{http://schemas.microsoft.com/analysisservices/2003/engine}'
    prop=next(p for p in ET.fromstring(data).iter(ns+'CalculationProperty') if p.findtext(ns+'CalculationReference')=='[Net]]Total]')
    assert prop.findtext(ns+'FormatString') == "'0.00'"
    assert prop.findtext(ns+'Description') == 'Keep description'
    assert prop.findtext(ns+'DisplayFolder') == 'Finance'

def test_measure_rename_rejects_unknown_script_statement(tmp_path):
    from xml.etree import ElementTree as ET
    from excel_model_lab.abf import AbfImage
    m = model()
    m.add_measure('Sales','Total','SUM(Sales[Amount])')
    output=m.save(tmp_path/'base.xlsx')
    m=WorkbookModel(output)
    script=next(f for f in m._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
    ns='{http://schemas.microsoft.com/analysisservices/2003/engine}'
    root=ET.fromstring(m._image.read_logical_file(script))
    command=ET.SubElement(root.find('.//'+ns+'Commands'),ns+'Command')
    ET.SubElement(command,ns+'Text').text='CREATE MEMBER CURRENTCUBE.[Other] AS [Total];'
    m._image=AbfImage(m._image.assemble({script.storage_path:ET.tostring(root)}))
    before=m._image.data
    with pytest.raises(DependencyError) as failure:
        m.rename_measure('Sales','Total','Renamed')
    assert failure.value.issues[0].code=='unsupported_script_reference'
    assert m._image.data==before

def test_measure_rename_handles_cdata_script_text(tmp_path):
    from xml.dom import minidom
    from excel_model_lab.abf import AbfImage
    from excel_model_lab.schema import _elements
    from excel_model_lab.measures import list_script_measures
    m=model().add_measure('Sales','Total','SUM(Sales[Amount])')
    m=WorkbookModel(m.save(tmp_path/'base.xlsx'))
    script=next(f for f in m._image.files if f.logical_path and f.logical_path.endswith('.scr.xml'))
    doc=minidom.parseString(m._image.read_logical_file(script))
    for node in _elements(doc,'Text'):
        text=''.join(c.data for c in node.childNodes if c.nodeType==c.TEXT_NODE)
        while node.firstChild:
            node.removeChild(node.firstChild)
        node.appendChild(doc.createCDATASection(text))
    m._image=AbfImage(m._image.assemble({script.storage_path:doc.toxml(encoding='utf-8')}))
    m.rename_measure('Sales','Total','Renamed')
    assert [x.name for x in list_script_measures(m._image.read_logical_file(script.storage_path))]==['Renamed']
