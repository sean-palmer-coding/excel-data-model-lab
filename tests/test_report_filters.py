from types import SimpleNamespace
from xml.etree import ElementTree as ET
import pandas as pd
import pytest
from excel_model_lab.report_filters import (PivotMemberFilter, PivotSort,
    validate_pivot_options, apply_pivot_options, member_unique_name, S)
from excel_model_lab.errors import ModelValidationError
from excel_model_lab.report_filters import source_member_values


def test_nullable_integer_source_float_readback_matches_integer_selection():
    spec, model = fixture_options((PivotMemberFilter('Sales', 'ID', (1, 3)),))
    model.get_table = lambda _: pd.DataFrame({'ID': [1.0, float('nan'), 3.0]})
    model.columns = [SimpleNamespace(table='Sales', name='ID', data_type='int64')]
    validate_pivot_options(spec, model)
    spec.member_filters = (PivotMemberFilter('Sales', 'ID', (1.0,)),)
    with pytest.raises(ModelValidationError, match='members support'):
        validate_pivot_options(spec, model)


@pytest.mark.parametrize('value', [1.5, float('inf'), float(2**53), -float(2**53), float(2**53 + 2)])
def test_unsafe_integer_source_float_readback_rejected(value):
    model = SimpleNamespace(columns=[SimpleNamespace(table='T', name='C', data_type='int64')])
    with pytest.raises(ModelValidationError, match='exact safe integer'):
        list(source_member_values(model, 'T', 'C', [value]))


def test_source_normalization_respects_double_and_boolean_metadata():
    model = SimpleNamespace(columns=[SimpleNamespace(table='T', name='C', data_type='double')])
    assert list(source_member_values(model, 'T', 'C', [1.0])) == [1.0]
    assert isinstance(list(source_member_values(model, 'T', 'C', [1.0]))[0], float)
    model.columns[0].data_type = 'bool'
    assert list(source_member_values(model, 'T', 'C', [1.0, 0.0])) == [True, False]


@pytest.mark.parametrize('state,pending', [('Unprocessed', {}), ('Processed', {('Other', 'ID'): [2]})])
def test_calculated_member_inventory_rejects_pending_or_stale_cache(state, pending):
    spec, model = fixture_options((PivotMemberFilter('Sales', 'ID', (None,)),))
    model.list_calculated_columns = lambda: [SimpleNamespace(table_id='Sales', name='ID', processing_state=state)]
    model._pending = pending
    with pytest.raises(ModelValidationError, match='require Excel calculation'):
        validate_pivot_options(spec, model)


def test_processed_calculated_member_inventory_remains_supported():
    model = SimpleNamespace(list_calculated_columns=lambda: [
        SimpleNamespace(table_id='T', name='C', processing_state='Processed')])
    assert list(source_member_values(model, 'T', 'C', [2, 4, 6])) == [2, 4, 6]


def fixture_options(filters=(), sorts=()):
    spec = SimpleNamespace(rows=(('Sales','Category'),), columns=(('Sales','ID'),),
        filters=(('Sales','Region'),), member_filters=filters, sorts=sorts)
    frame = pd.DataFrame({'Category':['Alpha','Beta','Gamma'], 'ID':[1,2,3], 'Region':['East','West','East']})
    return spec, SimpleNamespace(get_table=lambda table: frame)


def definitions():
    pivot = ET.fromstring(f'<pivotTableDefinition xmlns="{S}"><pivotFields count="3"><pivotField/><pivotField/><pivotField><items count="1"><item t="default"/></items></pivotField></pivotFields><pageFields><pageField fld="2" name="[Sales].[Region].[All]" cap="All"/></pageFields></pivotTableDefinition>')
    cache = ET.fromstring(f'<pivotCacheDefinition xmlns="{S}" recordCount="0"><cacheFields><cacheField/><cacheField/><cacheField/></cacheFields></pivotCacheDefinition>')
    return pivot,cache


def test_single_page_member_and_label_sort_without_result_cache():
    spec,model=fixture_options((PivotMemberFilter('Sales','Region',['East']),), (PivotSort('Sales','Category','descending'),))
    pivot,cache=definitions()
    apply_pivot_options(pivot,cache,spec,model)
    assert pivot.find(f'{{{S}}}pivotFields')[0].get('sortType')=='descending'
    page=pivot.find(f'{{{S}}}pageFields')[0]
    assert page.get('name')=='[Sales].[Region].&[East]'
    assert page.get('cap')=='East'
    assert cache.get('recordCount')=='0'
    assert not cache.findall(f'.//{{{S}}}sharedItems')


@pytest.mark.parametrize('column,members,index', [('Category',('Alpha','Gamma'),0),('ID',(1,3),1)])
def test_selected_row_column_members_are_real_source_labels(column,members,index):
    spec,model=fixture_options((PivotMemberFilter('Sales',column,members),))
    pivot,cache=definitions()
    apply_pivot_options(pivot,cache,spec,model)
    items=pivot.find(f'{{{S}}}pivotFields')[index].find(f'{{{S}}}items')
    assert [item.attrib for item in items]==[{'s':'1','x':'0'},{'s':'1','x':'1'},{'t':'default'}]
    labels=cache.find(f'{{{S}}}cacheFields')[index].find(f'{{{S}}}sharedItems')
    assert [n.get('v') for n in labels]==[str(v) for v in members]
    assert cache.get('recordCount')=='0'


def test_member_key_escaping_and_frozen_spec():
    assert member_unique_name('Ta]ble','Co]l','A]&[B')=='[Ta]]ble].[Co]]l].&[A]]&[B]'
    spec=PivotMemberFilter('T','C',['A'])
    assert spec.members==('A',)
    with pytest.raises(AttributeError):spec.members=('B',)


@pytest.mark.parametrize('option',[
    PivotMemberFilter('Sales','Category',()),
    PivotMemberFilter('Sales','Category',('Missing',)),
    PivotMemberFilter('Sales','Category',('Alpha','Alpha')),
    PivotMemberFilter('Sales','ID',('1',)),
    PivotMemberFilter('Sales','ID',(True,)),
    PivotMemberFilter('Sales','ID',(1.0,)),
    PivotMemberFilter('Sales','Category',('bad\x00',)),
    PivotMemberFilter('Sales','Absent',('Alpha',)),
])
def test_invalid_member_selection_fails_before_mutation(option):
    spec,model=fixture_options((option,))
    pivot,cache=definitions();before=(ET.tostring(pivot),ET.tostring(cache))
    with pytest.raises(ModelValidationError):apply_pivot_options(pivot,cache,spec,model)
    assert (ET.tostring(pivot),ET.tostring(cache))==before


@pytest.mark.parametrize('sorts',[
    (PivotSort('Sales','Region'),),
    (PivotSort('Sales','Category','manual'),),
    (PivotSort('Sales','Category'),PivotSort('Sales','Category')),
    ('bad',),
])
def test_invalid_sorts(sorts):
    spec,model=fixture_options(sorts=sorts)
    with pytest.raises(ModelValidationError):validate_pivot_options(spec,model)


def test_duplicate_filter_and_bare_member_string_rejected():
    option=PivotMemberFilter('Sales','Category',('Alpha',))
    spec,model=fixture_options((option,option))
    with pytest.raises(ModelValidationError):validate_pivot_options(spec,model)
    with pytest.raises(ModelValidationError):PivotMemberFilter('Sales','Category','Alpha')


def test_validation_uses_current_model_values():
    spec,model=fixture_options((PivotMemberFilter('Sales','Category',('Alpha',)),))
    validate_pivot_options(spec,model)
    model.get_table=lambda _:pd.DataFrame({'Category':['Gamma']})
    with pytest.raises(ModelValidationError,match='absent'):validate_pivot_options(spec,model)


def test_multiple_page_selection_enables_olap_hierarchy():
    spec,model=fixture_options((PivotMemberFilter('Sales','Region',('East','West')),))
    pivot,cache=definitions()
    cache.find(f'{{{S}}}cacheFields')[2].set('hierarchy','1')
    hierarchies=ET.SubElement(pivot,f'{{{S}}}pivotHierarchies')
    for _ in range(3):ET.SubElement(hierarchies,f'{{{S}}}pivotHierarchy')
    apply_pivot_options(pivot,cache,spec,model)
    assert hierarchies[1].get('multipleItemSelectionAllowed')=='1'
    assert hierarchies[0].get('multipleItemSelectionAllowed') is None
    assert pivot.find(f'{{{S}}}pivotFields')[2].get('multipleItemSelectionAllowed') is None


@pytest.mark.parametrize('value,key',[(True,'[True]'),(False,'[False]'),(None,''),(pd.NA,''),(pd.NaT,''),(pd.Timestamp('2026-01-02'),'[2026-01-02T00:00:00]')])
def test_native_typed_member_keys(value,key):
    assert member_unique_name('Typed','Field',value)=='[Typed].[Field].&'+key


@pytest.mark.parametrize('values,tags,lexical',[
    ([True,False],['b','b'],['1','0']),
    ([pd.Timestamp('2026-01-01'),pd.Timestamp('2026-01-02')],['d','d'],['2026-01-01T00:00:00','2026-01-02T00:00:00']),
    ([None,1],['m','n'],[None,'1']),
    (['', 'Text'],['s','s'],['','Text']),
])
def test_native_typed_shared_items(values,tags,lexical):
    spec=SimpleNamespace(rows=(('Typed','Field'),),columns=(),filters=(),sorts=(),
        member_filters=(PivotMemberFilter('Typed','Field',values),))
    dtype='Int64' if values==[None,1] else None
    frame=pd.DataFrame({'Field':pd.Series(values,dtype=dtype)})
    model=SimpleNamespace(get_table=lambda _:frame)
    pivot,cache=definitions()
    apply_pivot_options(pivot,cache,spec,model)
    shared=cache.find(f'{{{S}}}cacheFields')[0].find(f'{{{S}}}sharedItems')
    assert [n.tag.split('}')[1] for n in shared]==tags
    assert [n.get('v') for n in shared]==lexical


@pytest.mark.parametrize('value',[pd.Timestamp('2026-01-01',tz='UTC'),pd.Timestamp('2026-01-01 00:00:00.001')])
def test_unverified_datetime_precision_rejected(value):
    with pytest.raises(ModelValidationError,match='whole-second'):member_unique_name('T','C',value)


def test_boolean_shared_items_preserve_native_default_type_flags():
    spec=SimpleNamespace(rows=(('Typed','Flag'),),columns=(),filters=(),sorts=(),
        member_filters=(PivotMemberFilter('Typed','Flag',(True,)),))
    model=SimpleNamespace(get_table=lambda _:pd.DataFrame({'Flag':[True,False]}))
    pivot,cache=definitions()
    apply_pivot_options(pivot,cache,spec,model)
    shared=cache.find(f'{{{S}}}cacheFields')[0].find(f'{{{S}}}sharedItems')
    assert shared.attrib=={'count':'1'}
    assert shared[0].tag==f'{{{S}}}b' and shared[0].attrib=={'v':'1'}
