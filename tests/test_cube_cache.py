from io import BytesIO
from xml.dom import minidom
from zipfile import ZipFile
import pytest
from excel_model_lab.cube_cache import invalidated_cube_cache_parts, S, R, CONNECTION_REL

T='http://schemas.openxmlformats.org/package/2006/content-types'
MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'


def package(*, target='custom/connections2.xml', duplicate=False):
    data=BytesIO()
    with ZipFile(data,'w') as z:
        z.writestr('xl/workbook.xml',f'<s:workbook xmlns:s="{S}"><s:sheets/><s:pivotCaches/></s:workbook>')
        z.writestr('xl/_rels/workbook.xml.rels',f'<Relationships xmlns="{R}"><Relationship Id="r1" Type="{CONNECTION_REL}" Target="{target}"/></Relationships>')
        extra='<connection id="7" name="duplicate"/>' if duplicate else ''
        z.writestr('xl/custom/connections2.xml',f'<connections xmlns="{S}"><connection id="7" name="ThisWorkbookDataModel"><dbPr connection="Data Model Connection"/></connection><connection id="8" name="Remote"><dbPr connection="MSOLAP;Data Source=server"/></connection>{extra}</connections>')
        z.writestr('[Content_Types].xml',f'<Types xmlns="{T}">'+''.join(f'<Override PartName="/xl/custom/cache{i}.xml" ContentType="{MIME}"/>' for i in range(1,4))+'</Types>')
        for i,cid in [(1,7),(2,8),(3,7)]:
            payload='<s:tupleCache><s:entries count="1"><s:n v="100"/></s:entries><s:queryCache count="1"/></s:tupleCache>' if i!=3 else ''
            z.writestr(f'xl/custom/cache{i}.xml',f'<s:pivotCacheDefinition xmlns:s="{S}" xmlns:v="urn:vendor" tupleCache="1" v:custom="keep"><s:cacheSource type="external" connectionId="{cid}"/><s:cacheFields count="0"/>{payload}<s:extLst><s:ext uri="vendor"><v:payload/></s:ext></s:extLst></s:pivotCacheDefinition>')
    data.seek(0)
    return ZipFile(data)


def test_only_embedded_tuple_payload_removed_preserving_metadata_and_external():
    with package() as z:
        before=z.read('xl/custom/cache2.xml')
        changed=invalidated_cube_cache_parts(z,{})
        assert set(changed)=={'xl/custom/cache1.xml','xl/workbook.xml'}
        doc=minidom.parseString(changed['xl/custom/cache1.xml'])
        assert not doc.getElementsByTagNameNS(S,'tupleCache')
        assert doc.documentElement.getAttribute('tupleCache')=='1'
        assert doc.documentElement.getAttributeNS('urn:vendor','custom')=='keep'
        assert doc.getElementsByTagNameNS('urn:vendor','payload')
        assert doc.getElementsByTagNameNS(S,'cacheFields')
        assert z.read('xl/custom/cache2.xml')==before
        workbook=minidom.parseString(changed['xl/workbook.xml'])
        calc=workbook.getElementsByTagNameNS(S,'calcPr')[0]
        assert calc.getAttribute('calcId')=='0'
        assert calc.getAttribute('forceFullCalc')=='1'
        assert [n.localName for n in workbook.documentElement.childNodes if n.nodeType==n.ELEMENT_NODE]==['sheets','calcPr','pivotCaches']


def test_staged_parts_used_and_idempotent():
    with package() as z:
        changed=invalidated_cube_cache_parts(z,{})
        assert invalidated_cube_cache_parts(z,changed)=={}


def test_absolute_connection_target():
    with package(target='/xl/custom/connections2.xml') as z:
        assert set(invalidated_cube_cache_parts(z,{}))=={'xl/custom/cache1.xml','xl/workbook.xml'}


def test_duplicate_connections_fail_explicitly():
    with package(duplicate=True) as z:
        with pytest.raises(ValueError,match='unique'):
            invalidated_cube_cache_parts(z,{})


def test_unresolvable_connection_target_fails():
    with package(target='../../elsewhere.xml') as z:
        with pytest.raises(ValueError,match='resolve'):
            invalidated_cube_cache_parts(z,{})


def test_repack_invalidates_only_when_model_bytes_change(tmp_path, monkeypatch):
    from excel_model_lab.package import repack_workbook, MODEL_PART
    import excel_model_lab.structural_edit as structural
    import excel_model_lab.measure_layout as layout
    import excel_model_lab.native_sources as native
    monkeypatch.setattr(structural,'renamed_table_package_parts',lambda *args:{})
    monkeypatch.setattr(layout,'layout_parts',lambda *args:{})
    monkeypatch.setattr(native,'binding_package_parts',lambda *args:{})
    source=tmp_path/'source.xlsx'
    with package() as template, ZipFile(source,'w') as z:
        for info in template.infolist():z.writestr(info,template.read(info.filename))
        z.writestr(MODEL_PART,b'old model bytes')
    same=tmp_path/'same.xlsx'; changed=tmp_path/'changed.xlsx'
    repack_workbook(source,same,replacement_model=b'old model bytes')
    repack_workbook(source,changed,replacement_model=b'new model bytes')
    with ZipFile(source) as old, ZipFile(same) as equal, ZipFile(changed) as new:
        assert equal.read('xl/custom/cache1.xml')==old.read('xl/custom/cache1.xml')
        assert new.read('xl/custom/cache1.xml')!=old.read('xl/custom/cache1.xml')
        assert new.read('xl/custom/cache2.xml')==old.read('xl/custom/cache2.xml')
        assert new.read(MODEL_PART)==b'new model bytes'


def test_manual_calculation_mode_retained():
    with package() as z:
        staged={'xl/workbook.xml':f'<workbook xmlns="{S}"><calcPr calcMode="manual" calcId="123"/></workbook>'.encode()}
        updates=invalidated_cube_cache_parts(z,staged)
        calc=minidom.parseString(updates['xl/workbook.xml']).getElementsByTagNameNS(S,'calcPr')[0]
        assert calc.getAttribute('calcMode')=='manual'
        assert calc.getAttribute('fullCalcOnLoad')=='1'
