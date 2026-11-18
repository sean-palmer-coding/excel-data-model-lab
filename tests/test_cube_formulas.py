from types import SimpleNamespace
from datetime import datetime
import pandas as pd
import pytest

from excel_model_lab.cube_formulas import CubeFormulas, CubeCell
from excel_model_lab.errors import ModelValidationError


@pytest.fixture
def cubes():
    frame = pd.DataFrame({'Category': ['A]B', 'Beta', '=SUM(A1)', None],
                          'ID': pd.array([1, 2, 3, None], dtype='Int64')})
    model = SimpleNamespace(columns=[SimpleNamespace(table='Sales', name=n, data_type=str(frame[n].dtype)) for n in frame],
        get_table=lambda _: frame, _pending_measures=[SimpleNamespace(name='Total "Sales]')],
        _image=SimpleNamespace(files=[]))
    return CubeFormulas(model)


def test_mdx_and_excel_strings_escape_independently(cubes):
    member = cubes.member('Sales', 'Category', 'A]B', caption='A "quoted" member')
    assert str(member) == '=CUBEMEMBER("ThisWorkbookDataModel","[Sales].[Category].&[A]]B]","A ""quoted"" member")'
    measure = cubes.measure('Total "Sales]')
    assert '[Measures].[Total ""Sales]]]' in str(measure)
    formula = cubes.value(measure, member)
    assert str(formula).startswith('=CUBEVALUE("ThisWorkbookDataModel",CUBEMEMBER(')
    assert str(cubes.member('Sales', 'Category', '=SUM(A1)')).count('=') == 2


def test_sets_and_references_are_typed(cubes):
    a, b = cubes.member('Sales', 'Category', 'A]B'), cubes.member('Sales', 'Category', 'Beta')
    selected = cubes.set([a, b], caption='Two')
    assert '{[Sales].[Category].&[A]]B],[Sales].[Category].&[Beta]}' in str(selected)
    assert str(cubes.set_count(CubeCell('$B$7', 'set', "O'Brien"))) == "=CUBESETCOUNT('O''Brien'!$B$7)"
    assert str(cubes.ranked_member(CubeCell('B7', 'set'), 2)).endswith(',B7,2)')
    with pytest.raises(ModelValidationError): cubes.value('=SUM(A1)')
    with pytest.raises(ModelValidationError): cubes.set([a, a])
    with pytest.raises(ModelValidationError): cubes.set([a, cubes.member('Sales', 'ID', 1)])
    with pytest.raises(ModelValidationError): cubes.set_count(a)
    with pytest.raises(ModelValidationError): cubes.ranked_member(selected, True)
    with pytest.raises(ModelValidationError): cubes.ranked_member(selected, 0)
    with pytest.raises(ModelValidationError,match='same hierarchy'): cubes.value(a,b)
    with pytest.raises(ModelValidationError): cubes.set(None)


@pytest.mark.parametrize('address', ['A0', 'XFE1', 'A1048577', 'A1:B2', 'A1+1', 'Sheet!A1'])
def test_bad_cell_reference_rejected(address):
    with pytest.raises(ModelValidationError): CubeCell(address)


def test_missing_members_and_long_inline_arguments_rejected(cubes):
    with pytest.raises(ModelValidationError, match='Unknown model field'): cubes.member('Sales', 'Missing', 1)
    with pytest.raises(ModelValidationError, match='absent'): cubes.member('Sales', 'ID', 9)
    with pytest.raises(ModelValidationError, match='Unknown model measure'): cubes.measure('Missing')
    with pytest.raises(ModelValidationError, match='255'): cubes.member('Sales', 'ID', 1, caption='X'*256)
    with pytest.raises(ModelValidationError, match='255'): cubes.member('Sales', 'ID', 1, caption='😀'*128)
    with pytest.raises(ModelValidationError): cubes.member('Sales', 'ID', 1, caption='bad\x00')
    assert '[Sales].[ID].&"' in str(cubes.member('Sales', 'ID', None))


def test_write_uses_unavailable_cache_not_zero(cubes):
    calls = []
    worksheet = SimpleNamespace(write_formula=lambda *args: calls.append(args) or 0)
    cubes.write(worksheet, 'B3', cubes.measure('Total "Sales]'))
    assert calls[0][-1] == '#N/A'


def test_member_property_explicitly_rejected_for_embedded_model(cubes):
    with pytest.raises(ModelValidationError, match='not supported for the embedded Excel Data Model'):
        cubes.member_property(cubes.member('Sales', 'Category', 'Beta'), 'MEMBER_CAPTION')
