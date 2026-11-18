from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
source=Path('outputs/excel-model-lab/tests/data/native-authoring-base.xlsx')
output=Path('outputs/excel-model-lab/src/excel_model_lab/data/authoring-seed.xlsx')
s='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
p='http://schemas.openxmlformats.org/package/2006/relationships'
parts={
'[Content_Types].xml': '''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/><Override PartName="/xl/model/item.data" ContentType="application/vnd.openxmlformats-officedocument.model+data"/></Types>''',
'_rels/.rels':f'<Relationships xmlns="{p}"><Relationship Id="rId1" Type="{r}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
'xl/workbook.xml':f'<workbook xmlns="{s}" xmlns:r="{r}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
'xl/_rels/workbook.xml.rels':f'<Relationships xmlns="{p}"><Relationship Id="rId1" Type="{r}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{r}/powerPivotData" Target="model/item.data"/><Relationship Id="rId3" Type="{r}/connections" Target="connections.xml"/></Relationships>',
'xl/worksheets/sheet1.xml':f'<worksheet xmlns="{s}"><sheetData/></worksheet>',
'xl/connections.xml':f'<connections xmlns="{s}"><connection id="1" name="ThisWorkbookDataModel" description="Data Model" type="5" refreshedVersion="8" minRefreshableVersion="5"><dbPr connection="Data Model Connection" command="Model" commandType="1"/><olapPr sendLocale="1" rowDrillCount="1000"/><extLst><ext uri="{{DE250136-89BD-433C-8126-D09CA5730AF9}}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:connection id="" model="1"/></ext></extLst></connection></connections>'}
with ZipFile(source) as z: parts['xl/model/item.data']=z.read('xl/model/item.data')
with ZipFile(output,'w',compression=ZIP_DEFLATED) as z:
    for name,data in parts.items(): z.writestr(name,data)
print(output)

