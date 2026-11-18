import os
import pandas as pd
import pytest
from excel_model_lab import WorkbookModel
from excel_model_lab.abf import UnsupportedModelError
from excel_model_lab.collation import string_order


def test_equivalent_spellings_are_rejected():
    with pytest.raises(UnsupportedModelError, match='compare equal'):
        string_order(['Alpha','ALPHA'])
    with pytest.raises(UnsupportedModelError, match='NUL'):
        string_order(['a\x00b'])


@pytest.mark.skipif(os.name != 'nt', reason='Windows NLS backend')
def test_unicode_public_roundtrip(tmp_path):
    values=['caf\u00e9','\u6771\u4eac','\u00c5ngstr\u00f6m','na\u00efve','R\u00e9sum\u00e9','\U0001f600']
    m=WorkbookModel.create({'Unicode':pd.DataFrame({'ID':range(1,7),'Label':values,'Amount':range(1,7)})})
    m.add_hierarchy('Unicode','Labels',['Label','ID'])
    output=m.save(tmp_path/'unicode.xlsx')
    assert WorkbookModel(output).get_table('Unicode')['Label'].tolist()==values
    with pytest.raises(UnsupportedModelError, match='compare equal'):
        string_order(['caf\u00e9','cafe\u0301'])


def test_nonwindows_unicode_uses_pinned_weights(monkeypatch):
    import types
    import excel_model_lab.collation as collation
    monkeypatch.setattr(collation, 'os', types.SimpleNamespace(name='posix'))
    collation.text_key.cache_clear()
    assert string_order(['Beta','Alpha']) == [1,0]
    assert string_order(['\u6771\u4eac', 'caf\u00e9']) == [1, 0]
    with pytest.raises(UnsupportedModelError, match=r'U\+E000'):
        string_order(['\ue000'])
    collation.text_key.cache_clear()


@pytest.mark.skipif(os.name != 'nt', reason='Windows NLS backend')
def test_equivalent_relationship_keys_match_across_tables(tmp_path):
    from excel_model_lab.collation import text_key
    from excel_model_lab.relationship_storage import stored_relationships
    assert text_key('caf\u00e9') == text_key('CAFE\u0301')
    m=WorkbookModel.create({'Lookup':pd.DataFrame({'Key':['caf\u00e9','\u6771\u4eac']}),
        'Facts':pd.DataFrame({'Key':['CAFE\u0301','\u6771\u4eac'], 'Amount':[10,20]})})
    m.add_relationship('Facts','Key','Lookup','Key')
    out=m.save(tmp_path/'related-unicode.xlsx')
    reopened=WorkbookModel(out)
    assert reopened.validate().valid
    assert len(stored_relationships(reopened)) == 1

