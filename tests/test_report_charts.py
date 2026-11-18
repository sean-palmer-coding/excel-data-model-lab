from types import SimpleNamespace
from xml.etree import ElementTree as ET
import pytest
from excel_model_lab.report_charts import chart_xml, CHART_NS, DRAWING_NS

NS = {'c': CHART_NS, 'a': DRAWING_NS}

@pytest.mark.parametrize('kind,element,direction,positions', [
    ('column','barChart','col',('b','l')),
    ('bar','barChart','bar',('l','b')),
    ('line','lineChart',None,('b','l')),
])
def test_pivot_chart_source_axes_and_no_invented_cache(kind,element,direction,positions):
    root=ET.fromstring(chart_xml({'chart_type':kind},'SalesPivot','Report','book.xlsx'))
    assert root.findtext('c:pivotSource/c:name',namespaces=NS)=='[book.xlsx]Report!SalesPivot'
    assert root.find('c:pivotSource/c:fmtId',NS).get('val')=='0'
    plot=root.find('c:chart/c:plotArea',NS)
    chart=plot.find('c:'+element,NS)
    assert chart is not None
    if direction:assert chart.find('c:barDir',NS).get('val')==direction
    axes=[plot.find('c:catAx',NS),plot.find('c:valAx',NS)]
    ids=[axis.find('c:axId',NS).get('val') for axis in axes]
    assert [n.get('val') for n in chart.findall('c:axId',NS)]==ids
    assert [axis.find('c:crossAx',NS).get('val') for axis in axes]==ids[::-1]
    assert tuple(axis.find('c:axPos',NS).get('val') for axis in axes)==positions
    assert not root.findall('.//c:ser',NS)
    assert not root.findall('.//c:numCache',NS)
    assert not root.findall('.//c:strCache',NS)
    assert not root.findall('.//c:pt',NS)
    assert not root.findall('.//c:f',NS)


def test_chart_dataclass_shape_and_literal_pivot_name_escaping():
    spec=SimpleNamespace(chart_type='line',title='Revenue < net & tax > 東京',
                         width=640,height=360,cell='G3',name='Trend',sheet=None)
    root=ET.fromstring(chart_xml(spec,'SalesPivot',"Owner's Report",'Sales & Costs.xlsx'))
    assert root.findtext('c:pivotSource/c:name',namespaces=NS)=="[Sales & Costs.xlsx]Owner's Report!SalesPivot"
    assert root.findtext('c:chart/c:title/c:tx/c:rich/a:p/a:r/a:t',namespaces=NS)==spec.title
    assert root.find('c:chart/c:autoTitleDeleted',NS).get('val')=='0'
    assert root.find('c:chart/c:plotArea/c:lineChart',NS) is not None


def test_empty_title_has_no_title_object_and_default_column_type():
    root=ET.fromstring(chart_xml({},'SalesPivot','Report','book.xlsx'))
    assert root.find('c:chart/c:title',NS) is None
    assert root.find('c:chart/c:autoTitleDeleted',NS).get('val')=='1'
    assert root.find('c:chart/c:plotArea/c:barChart/c:barDir',NS).get('val')=='col'


@pytest.mark.parametrize('spec,pivot,sheet,book,match', [
    ({'chart_type':'pie'},'P','S','book.xlsx','chart_type'),
    ({'title':'Bad\x00title'},'P','S','book.xlsx','XML'),
    ({'title':None},'P','S','book.xlsx','title'),
    ({},'','S','book.xlsx','PivotTable name'),
    ({},'P!other','S','book.xlsx','source name'),
    ({},'P','Bad[Sheet]','book.xlsx','source name'),
    ({},'P',"'Bad",'book.xlsx','worksheet name'),
    ({},'P','X'*32,'book.xlsx','worksheet name'),
    ({},'P','S','C:/book.xlsx','source name'),
    ({},'P','S','[book].xlsx','source name'),
])
def test_chart_rejects_invalid_type_xml_and_source_bindings(spec,pivot,sheet,book,match):
    with pytest.raises(ValueError,match=match):chart_xml(spec,pivot,sheet,book)


@pytest.mark.parametrize('kind', ['column', 'bar', 'line'])
def test_formatting_preserves_dynamic_series_and_schema_order(kind):
    root = ET.fromstring(chart_xml(dict(chart_type=kind, style=48,
        legend_position='bottom', category_axis_title='Region & category',
        value_axis_title='Amount < USD', number_format='$#,##0.00',
        show_data_labels=True, background_color='aabbcc', plot_background_color='123456'),
        'P', 'S', 'book.xlsx'))
    assert root.find('c:style', NS).get('val') == '48'
    assert root.find('c:chart/c:legend/c:legendPos', NS).get('val') == 'b'
    plot = root.find('c:chart/c:plotArea', NS)
    for tag, title in [('catAx', 'Region & category'), ('valAx', 'Amount < USD')]:
        axis = plot.find('c:' + tag, NS)
        assert axis.findtext('c:title/c:tx/c:rich/a:p/a:r/a:t', namespaces=NS) == title
        children = [x.tag.split('}')[1] for x in axis]
        assert children.index('axPos') < children.index('title') < children.index('numFmt')
    fmt = plot.find('c:valAx/c:numFmt', NS)
    assert fmt.attrib == {'formatCode': '$#,##0.00', 'sourceLinked': '0'}
    assert plot.find('c:catAx/c:numFmt', NS).get('formatCode') == 'General'
    series_type = plot.find('c:lineChart' if kind == 'line' else 'c:barChart', NS)
    labels = series_type.find('c:dLbls', NS)
    assert labels.find('c:showVal', NS).get('val') == '1'
    assert labels.find('c:numFmt', NS).get('formatCode') == '$#,##0.00'
    children = [x.tag.split('}')[1] for x in series_type]
    assert children.index('varyColors') < children.index('dLbls') < children.index('axId')
    assert root.find('c:spPr/a:solidFill/a:srgbClr', NS).get('val') == 'AABBCC'
    assert plot.find('c:spPr/a:solidFill/a:srgbClr', NS).get('val') == '123456'
    assert not root.findall('.//c:ser', NS)
    assert not root.findall('.//c:numCache', NS)


@pytest.mark.parametrize('position,expected', [('right','r'),('left','l'),('top','t'),('bottom','b'),('top_right','tr')])
def test_legend_positions(position, expected):
    root = ET.fromstring(chart_xml({'legend_position':position}, 'P', 'S', 'book.xlsx'))
    assert root.find('c:chart/c:legend/c:legendPos', NS).get('val') == expected


def test_legend_disabled_and_default_labels_absent():
    root = ET.fromstring(chart_xml({'show_legend':False}, 'P', 'S', 'book.xlsx'))
    assert root.find('c:chart/c:legend', NS) is None
    assert root.find('.//c:dLbls', NS) is None


@pytest.mark.parametrize('name,value', [
    ('style',0),('style',49),('style',True),('style',2.0),
    ('show_legend',1),('show_data_labels','yes'),
    ('legend_position','middle'),('legend_position',[]),
    ('category_axis_title',None),('value_axis_title','bad\x00'),
    ('number_format',''),('number_format',None),
    ('background_color','#123456'),('background_color','GGGGGG'),
    ('plot_background_color',123456),('plot_background_color','ABC'),
])
def test_formatting_rejects_invalid_values(name,value):
    with pytest.raises(ValueError):
        chart_xml({name:value}, 'P', 'S', 'book.xlsx')
