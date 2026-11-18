"""OOXML PivotChart parts linked to a real PivotTable.

Chart series are intentionally absent until Excel refreshes the PivotTable.
This writer neither evaluates the model nor invents cached chart values.
"""
from collections.abc import Mapping
import re
from xml.dom import minidom

CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIVOT_OPTIONS_NS = 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart'
LEGEND_POSITIONS = {'right': 'r', 'left': 'l', 'top': 't', 'bottom': 'b', 'top_right': 'tr'}


def validate_chart_format(spec):
    """Validate formatting independently of drawing/package allocation."""
    style = _value(spec, 'style', 2)
    if type(style) is not int or not 1 <= style <= 48:
        raise ValueError('Chart style must be an integer from 1 to 48')
    for name in ('show_legend', 'show_data_labels'):
        if type(_value(spec, name, name == 'show_legend')) is not bool:
            raise ValueError(f'{name} must be a boolean')
    position = _value(spec, 'legend_position', 'right')
    if not isinstance(position, str) or position not in LEGEND_POSITIONS:
        raise ValueError('Unsupported legend_position')
    for name in ('category_axis_title', 'value_axis_title'):
        _xml_text(_value(spec, name, ''), name, empty=True)
    _xml_text(_value(spec, 'number_format', 'General'), 'number_format')
    for name in ('background_color', 'plot_background_color'):
        color = _value(spec, name, None)
        if color is not None and (not isinstance(color, str) or not re.fullmatch(r'[0-9a-fA-F]{6}', color)):
            raise ValueError(f'{name} must be a six-digit RGB color or None')


def _value(spec, name, default):
    return spec.get(name, default) if isinstance(spec, Mapping) else getattr(spec, name, default)


def _xml_text(value, label, *, empty=False):
    if not isinstance(value, str) or (not empty and not value.strip()):
        raise ValueError(f'{label} must be a nonempty string' if not empty else f'{label} must be a string')
    if any(not (ord(c) in (9, 10, 13) or 0x20 <= ord(c) <= 0xD7FF or
                0xE000 <= ord(c) <= 0xFFFD or 0x10000 <= ord(c) <= 0x10FFFF) for c in value):
        raise ValueError(f'{label} contains a character XML cannot represent')
    return value


def _pivot_source_name(pivot_name, pivot_sheet, workbook_name):
    """Office's name binding is a literal object path, not a cell formula.

    MS-OI29500 section 2.1.1503 specifies [filename]sheet!pivottable. Adding
    formula-style quotes around a sheet name would change this object name.
    """
    pivot_name = _xml_text(pivot_name, 'PivotTable name')
    pivot_sheet = _xml_text(pivot_sheet, 'PivotTable sheet')
    workbook_name = _xml_text(workbook_name, 'Workbook filename')
    for value, label, forbidden in (
        (pivot_name, 'PivotTable name', '[]!'),
        (pivot_sheet, 'PivotTable sheet', '[]:\\/*?'),
        (workbook_name, 'Workbook filename', '[]\\/'),
    ):
        if any(ord(c) < 32 or c in forbidden for c in value):
            raise ValueError(f'{label} cannot be represented in a PivotChart source name')
    if len(pivot_sheet) > 31 or pivot_sheet.startswith("'") or pivot_sheet.endswith("'"):
        raise ValueError('PivotTable sheet is not a valid Excel worksheet name')
    return f'[{workbook_name}]{pivot_sheet}!{pivot_name}'


def chart_xml(spec, pivot_name, pivot_sheet, workbook_name):
    """Return a PivotChart part for a column, bar, or line chart.

    ``spec`` may be a mapping or an object with ``chart_type`` and ``title``
    attributes. Position, size, sheet and object names belong to the drawing
    and package layer. ``workbook_name`` must be the output workbook basename.
    The associated PivotTable must exist in that workbook when Excel opens it.
    """
    kind = _value(spec, 'chart_type', 'column')
    if kind not in ('column', 'bar', 'line'):
        raise ValueError('chart_type must be column, bar, or line')
    title = _xml_text(_value(spec, 'title', ''), 'Chart title', empty=True)
    validate_chart_format(spec)
    source_name = _pivot_source_name(pivot_name, pivot_sheet, workbook_name)
    doc = minidom.Document()
    root = doc.createElementNS(CHART_NS, 'c:chartSpace')
    root.setAttribute('xmlns:c', CHART_NS)
    root.setAttribute('xmlns:a', DRAWING_NS)
    root.setAttribute('xmlns:c14', PIVOT_OPTIONS_NS)
    doc.appendChild(root)

    def node(parent, name, text=None, **attributes):
        prefix = name.split(':', 1)[0] if ':' in name else 'c'
        qualified = name if ':' in name else 'c:' + name
        namespace = {'c': CHART_NS, 'a': DRAWING_NS, 'c14': PIVOT_OPTIONS_NS}[prefix]
        child = doc.createElementNS(namespace, qualified)
        for key, value in attributes.items():
            child.setAttribute(key, str(value))
        if text is not None:
            child.appendChild(doc.createTextNode(text))
        parent.appendChild(child)
        return child

    node(root, 'date1904', val=0)
    node(root, 'lang', val='en-US')
    node(root, 'roundedCorners', val=0)
    def add_title(parent, text):
        title_node = node(parent, 'title')
        rich = node(node(title_node, 'tx'), 'rich')
        node(rich, 'a:bodyPr')
        node(rich, 'a:lstStyle')
        run = node(node(rich, 'a:p'), 'a:r')
        node(run, 'a:rPr', lang='en-US')
        node(run, 'a:t', text)
        node(title_node, 'overlay', val=0)

    def fill(parent, color):
        if color is not None:
            node(node(node(parent, 'spPr'), 'a:solidFill'), 'a:srgbClr', val=color.upper())

    node(root, 'style', val=_value(spec, 'style', 2))
    source = node(root, 'pivotSource')
    node(source, 'name', source_name)
    node(source, 'fmtId', val=0)
    chart = node(root, 'chart')
    if title:
        add_title(chart, title)
    node(chart, 'autoTitleDeleted', val=0 if title else 1)
    plot = node(chart, 'plotArea')
    node(plot, 'layout')
    chart_type = node(plot, 'lineChart' if kind == 'line' else 'barChart')
    if kind != 'line':
        node(chart_type, 'barDir', val='bar' if kind == 'bar' else 'col')
    node(chart_type, 'grouping', val='standard' if kind == 'line' else 'clustered')
    node(chart_type, 'varyColors', val=0)
    # CT_BarChart/CT_LineChart permit zero series. Excel derives the series from
    # pivotSource after refreshing, so no fabricated point/cache data is stored.
    if _value(spec, 'show_data_labels', False):
        labels = node(chart_type, 'dLbls')
        node(labels, 'numFmt', formatCode=_value(spec, 'number_format', 'General'), sourceLinked=0)
        for name, value in (('showLegendKey', 0), ('showVal', 1), ('showCatName', 0),
                            ('showSerName', 0), ('showPercent', 0), ('showBubbleSize', 0)):
            node(labels, name, val=value)
    if kind == 'line':
        node(chart_type, 'marker', val=0)
        node(chart_type, 'smooth', val=0)
    else:
        node(chart_type, 'gapWidth', val=150)
        node(chart_type, 'overlap', val=0)
    category_id, value_id = 1, 2
    node(chart_type, 'axId', val=category_id)
    node(chart_type, 'axId', val=value_id)
    for category, identity, cross in ((True, category_id, value_id), (False, value_id, category_id)):
        axis = node(plot, 'catAx' if category else 'valAx')
        node(axis, 'axId', val=identity)
        node(node(axis, 'scaling'), 'orientation', val='minMax')
        node(axis, 'delete', val=0)
        position = ('l' if category else 'b') if kind == 'bar' else ('b' if category else 'l')
        node(axis, 'axPos', val=position)
        if not category:
            node(axis, 'majorGridlines')
        axis_title = _value(spec, 'category_axis_title' if category else 'value_axis_title', '')
        if axis_title:
            add_title(axis, axis_title)
        number_format = 'General' if category else _value(spec, 'number_format', 'General')
        node(axis, 'numFmt', formatCode=number_format, sourceLinked=int(number_format == 'General'))
        node(axis, 'majorTickMark', val='none')
        node(axis, 'minorTickMark', val='none')
        node(axis, 'tickLblPos', val='nextTo')
        node(axis, 'crossAx', val=cross)
        node(axis, 'crosses', val='autoZero')
        if category:
            node(axis, 'auto', val=1)
            node(axis, 'lblAlgn', val='ctr')
            node(axis, 'lblOffset', val=100)
        else:
            node(axis, 'crossBetween', val='between')
    fill(plot, _value(spec, 'plot_background_color', None))
    if _value(spec, 'show_legend', True):
        legend = node(chart, 'legend')
        node(legend, 'legendPos', val=LEGEND_POSITIONS[_value(spec, 'legend_position', 'right')])
        node(legend, 'layout')
        node(legend, 'overlay', val=0)
    node(chart, 'plotVisOnly', val=1)
    node(chart, 'dispBlanksAs', val='gap')
    node(chart, 'showDLblsOverMax', val=0)
    fill(root, _value(spec, 'background_color', None))
    extension = node(node(root, 'extLst'), 'ext', uri='{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}')
    options = node(extension, 'c14:pivotOptions')
    for name in ('dropZoneFilter', 'dropZoneCategories', 'dropZoneData', 'dropZonesVisible'):
        node(options, 'c14:' + name, val=1)
    return doc.toxml(encoding='utf-8')
