"""Cloned tables must not share a globally scoped engine count-measure name."""
from xml.dom import minidom

import pandas as pd

from excel_model_lab import WorkbookModel
from excel_model_lab.schema import _child, _elements, _text


def test_created_tables_have_distinct_engine_count_measure_names(tmp_path):
    model = WorkbookModel.create({
        'Sales': pd.DataFrame({'Amount': [10, 20, 30]}),
        'Products': pd.DataFrame({'Label': ['Tent', 'Boot', 'Pack']}),
        'Customers': pd.DataFrame({'Customer': ['A', 'B', 'C']}),
    })
    saved = WorkbookModel(model.save(tmp_path / 'multi-table.xlsx'))
    names = []
    for file in saved._image.files:
        if not file.logical_path or not file.logical_path.endswith('.det.xml'):
            continue
        document = minidom.parseString(saved._image.read_logical_file(file))
        group = _elements(document, 'MeasureGroup')[0]
        table_id = _text(_child(group, 'ID'))
        measure = _elements(group, 'Measure')[0]
        name = _text(_child(measure, 'Name'))
        assert name == '__XL_Count ' + table_id
        names.append(name)
    assert len(names) == 3
    assert len(set(names)) == 3
