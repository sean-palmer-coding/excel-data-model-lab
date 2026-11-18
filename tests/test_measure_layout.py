from zipfile import ZipFile
import posixpath
from xml.dom import minidom
import pandas as pd
from excel_model_lab import WorkbookModel
from excel_model_lab.schema import _elements, _text

def layouts(path):
    result = {}
    with ZipFile(path) as z:
        rels = minidom.parseString(z.read('xl/_rels/workbook.xml.rels'))
        for r in _elements(rels, 'Relationship'):
            if r.getAttribute('Type').endswith('/customXml'):
                name = posixpath.normpath(posixpath.join('xl', r.getAttribute('Target')))
                outer = minidom.parseString(z.read(name))
                result[outer.documentElement.namespaceURI.rsplit('/',1)[-1]] = (z.read(name),
                    minidom.parseString(''.join(n.data for n in _elements(outer, 'CustomContent')[0].childNodes
                                               if n.nodeType in (n.TEXT_NODE, n.CDATA_SECTION_NODE))))
    return result

def measure_rows(doc):
    result = {}
    for states in _elements(doc, 'ViewStates'):
        for entry in states.childNodes:
            keys = _elements(entry, 'Key')
            rows = _elements(entry, 'Row')
            if keys and _text(keys[-1]).startswith('Measures\\') and '\\TagInfo\\' not in _text(keys[-1]):
                result[_text(keys[-1])] = int(_text(rows[0])) if rows else 0
    return result

def test_measure_layout_links_positions_and_preservation(tmp_path):
    m = WorkbookModel.create({'Sales':pd.DataFrame({'Amount':[20,30,40]})})
    m.add_measure('Sales','Total','SUM(Sales[Amount])')
    m.add_measure('Sales','Target','100')
    m.add_kpi('Sales','Total',target_measure='Target')
    first = m.save(tmp_path/'first.xlsx')
    initial = layouts(first)
    assert set(initial) == {'Diagrams','MeasureGridState'}
    rows = measure_rows(initial['Diagrams'][1])
    assert rows == {'Measures\\Total':0,'Measures\\Target':1}
    second = WorkbookModel(first).save(tmp_path/'second.xlsx')
    assert {k:v[0] for k,v in layouts(second).items()} == {k:v[0] for k,v in initial.items()}
    m = WorkbookModel(second).add_measure('Sales','Third','3')
    third = m.save(tmp_path/'third.xlsx')
    assert measure_rows(layouts(third)['Diagrams'][1]) == {**rows,'Measures\\Third':2}
