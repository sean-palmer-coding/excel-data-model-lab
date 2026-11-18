from pathlib import Path
from xml.dom import minidom
from zipfile import ZipFile
import pandas as pd
import pytest

from excel_model_lab import ExcelModelTemplate


@pytest.mark.parametrize('custom_connections', [False, True])
def test_insert_model_into_ordinary_package(tmp_path, custom_connections):
    seed = Path(__file__).parents[1] / 'src/excel_model_lab/data/authoring-seed.xlsx'
    source = tmp_path / 'ordinary.xlsx'
    with ZipFile(seed) as original, ZipFile(source, 'w') as package:
        for item in original.infolist():
            if item.filename in ('xl/model/item.data', 'xl/connections.xml'):
                continue
            data = original.read(item.filename)
            if item.filename == 'xl/_rels/workbook.xml.rels':
                doc = minidom.parseString(data)
                for node in list(doc.documentElement.childNodes):
                    if node.nodeType != node.ELEMENT_NODE:
                        continue
                    kind = node.getAttribute('Type').rsplit('/', 1)[-1]
                    if kind == 'model' or (kind == 'connections' and not custom_connections):
                        node.parentNode.removeChild(node)
                    elif kind == 'connections':
                        node.setAttribute('Target', 'custom-links.xml')
                data = doc.toxml(encoding='utf-8')
            elif item.filename == '[Content_Types].xml':
                doc = minidom.parseString(data)
                for node in list(doc.documentElement.childNodes):
                    if node.nodeType == node.ELEMENT_NODE and node.getAttribute('PartName') in ('/xl/model/item.data', '/xl/connections.xml'):
                        node.parentNode.removeChild(node)
                data = doc.toxml(encoding='utf-8')
            package.writestr(item, data)
        if custom_connections:
            package.writestr('xl/custom-links.xml', b'<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="7" name="Existing connection"/></connections>')
    model = ExcelModelTemplate.create({'Inventory': pd.DataFrame({'ID': [10, 20, 30]})}, workbook=source)
    output = model.save(tmp_path / 'inserted.xlsx')
    with ZipFile(source) as before, ZipFile(output) as after:
        assert after.read('xl/worksheets/sheet1.xml') == before.read('xl/worksheets/sheet1.xml')
        path = 'xl/custom-links.xml' if custom_connections else 'xl/connections.xml'
        doc = minidom.parseString(after.read(path))
        names = [n.getAttribute('name') for n in doc.getElementsByTagName('connection')]
        assert 'ThisWorkbookDataModel' in names
        if custom_connections:
            assert 'Existing connection' in names
            assert 'xl/connections.xml' not in after.namelist()
        assert after.testzip() is None
