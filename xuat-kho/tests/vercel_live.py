"""Live deployment smoke check; run with the deployed https URL argument."""
import io
import json
from pathlib import Path
import sys
import urllib.request
from zipfile import ZipFile
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
def run(base):
    base=base.rstrip('/')
    source=(ROOT/'samples/data-mau.xlsx').read_bytes()
    def request(path,payload=None):
        boundary='remote-test-boundary-12901'
        chunks=[b'--'+boundary.encode()+b'\r\nContent-Disposition: form-data; name="data"; filename="data-mau.xlsx"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n'+source+b'\r\n']
        if payload is not None:
            chunks.append(b'--'+boundary.encode()+b'\r\nContent-Disposition: form-data; name="payload"\r\n\r\n'+json.dumps(payload,ensure_ascii=False).encode()+b'\r\n')
        chunks.append(b'--'+boundary.encode()+b'--\r\n')
        req=urllib.request.Request(base+path,b''.join(chunks),{'Content-Type':'multipart/form-data; boundary='+boundary})
        try:
            with urllib.request.urlopen(req,timeout=60) as response:return response.read()
        except urllib.error.HTTPError as error:
            raise RuntimeError(f'{path}: {error.code} {error.read().decode(errors="replace")}') from None
    data=json.loads(request('/api/analyze'))
    assert len(data['sheets'])==4
    assert len([sheet for sheet in data['sheets'] if sheet['selected']])==3
    print('Live recognize: 4 sheets, 3 selected.')
    payload={'token':data['token'],'sheets':[sheet for sheet in data['sheets'] if sheet['selected']],'overrides':{},'format':'xlsx'}
    result=request('/api/export',payload)
    wb=load_workbook(io.BytesIO(result));assert len(wb.worksheets)==3
    assert [sheet['H27'].value for sheet in wb]==[593000,432000,544000]
    print('Live XLSX export: 3 receipts, total1,569,000 VND.')
    payload['format']='zip'
    result=request('/api/export',payload)
    with ZipFile(io.BytesIO(result)) as archive:assert len(archive.namelist())==3
    print('Live ZIP export: 3 separate files.')
    payload={'overrides':{'Trung tâm':{'header':5,'header_span':1,'mapping':{'name':'B','quantity':'D','price':'E','amount':'F','unit':'C','note':'G'}}}}
    remapped=json.loads(request('/api/remap',payload));assert remapped['sheets'][0]['total']==593000
    print('Live manual recognition/remap passed.')

if __name__=='__main__':run(sys.argv[1])
