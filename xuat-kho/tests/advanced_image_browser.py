"""Check region selection, customer-specific exports, edits and theme retention.

--real additionally reads the supplied dense screenshot using bundled Tesseract.
It checks structural completeness and review blockers, not perfect OCR accuracy.
"""
import argparse,io,json,sys,tempfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parent))
from test_multi_customer_image import document
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright,expect
ROOT=Path(__file__).resolve().parents[1]

def run(base,real=False):
 with tempfile.TemporaryDirectory() as directory,sync_playwright() as p:
  browser=p.chromium.launch(headless=True,executable_path='/usr/bin/chromium',args=['--no-sandbox'])
  page=browser.new_page(viewport={'width':1440,'height':1050},locale='vi-VN');errors=[];page.on('pageerror',lambda e:errors.append(str(e)))
  page.goto(base);page.locator('#theme-normal').click();page.reload();assert page.evaluate('document.documentElement.dataset.theme')=='normal'
  page.locator('#theme-linh').click()
  regions=page.evaluate('''()=>{let c=document.createElement('canvas');c.width=800;c.height=600;let x=c.getContext('2d');x.fillStyle='white';x.fillRect(0,0,800,600);x.fillStyle='black';x.font='24px Arial';x.fillText('Khach hang ABC',200,200);x.fillText('Gao 2 10000 20000',200,250);const region=ImageRegions.detect(c);const result=ImageRegions.prepare(c);return {region,width:result.canvas.width,height:result.canvas.height};}''')
  assert regions['region']['width']<.6 and regions['region']['height']<.3,regions
  assert .20<regions['region']['x']<.30,regions
  page.locator('#mode-images').click();page.locator('#image-files').set_input_files(str(ROOT/'samples/anh-nhieu-khach-hang.jpg'))
  page.locator('[data-crop-image="0"]').click();page.locator('#image-region-dialog').wait_for(state='visible')
  page.locator('[data-region-field="x"]').fill('10');page.locator('[data-region-field="y"]').fill('20');page.locator('[data-region-field="width"]').fill('80');page.locator('[data-region-field="height"]').fill('60');page.locator('[data-region-apply]').click();expect(page.locator('.thumbnail-caption')).to_contain_text('Vùng đã chọn')
  page.locator('[data-crop-image="0"]').click();page.locator('[data-region-auto]').click();expect(page.locator('.thumbnail-caption')).to_contain_text('Tự bỏ lề trắng')
  page.locator('[data-full-image="0"]').click();expect(page.locator('.thumbnail-caption')).to_contain_text('Toàn bộ ảnh')
  fixture=document()
  page.evaluate('''doc=>{window.Tesseract={createWorker:async()=>({setParameters:async()=>{},recognize:async()=>({data:{text:doc.text,tsv:'',confidence:95}}),terminate:async()=>{}})};ImageRegions.readTable=async()=>doc.table_data;}''',fixture)
  page.locator('#analyze').click();page.wait_for_function("busy === ''")
  assert page.evaluate('sheets.length')==2
  assert page.locator('[data-customer-select]').count()==2
  page.locator('[data-customer-select="1"]').check();assert page.evaluate('sheets[1]._ocrConfirmed && sheets[1].selected && !sheets[0].selected')
  page.locator('[data-item="0"][data-item-field="note"]').fill('Giữ chỉnh sửa khách hàng B')
  page.locator('[data-tab="source"]').click();page.locator('#ocr-text-draft').fill('Bản nháp chưa gửi khách B')
  page.locator('[data-open-sheet="0"]').click();page.locator('[data-tab="source"]').click()
  corrected='Khách hàng: Khách A đã sửa\nNgày xuất: 02/10/2026\nTên hàng | ĐVT | Số lượng | Đơn giá | Thành tiền\nGạo tẻ | kg | 4 | 10000 | 40000'
  page.locator('#ocr-text-draft').fill(corrected);page.locator('[data-reparse-ocr]').click();page.wait_for_function("busy === ''")
  assert page.evaluate('sheets[0].total')==40000
  assert page.evaluate('sheets[1].items[0].note')=='Giữ chỉnh sửa khách hàng B'
  assert page.evaluate('sheets[1]._ocrDraft')=='Bản nháp chưa gửi khách B'
  assert page.evaluate('sheets[1].selected && sheets[1]._ocrConfirmed && !sheets[0].selected')
  page.locator('#theme-normal').click();page.locator('#theme-linh').click()
  assert page.evaluate('sheets[1].items[0].note')=='Giữ chỉnh sửa khách hàng B'
  with page.expect_download() as pending:page.locator('#export').click()
  output=Path(directory)/'customer.xlsx';pending.value.save_as(output);wb=load_workbook(output);assert len(wb.worksheets)==1 and wb.active['H27'].value==40000
  page.wait_for_function("busy === ''")
  for theme in ['normal','linh']:
   page.locator('#theme-'+theme).click()
   for width in [1440,768,390,320]:
    page.set_viewport_size({'width':width,'height':900})
    assert page.evaluate('document.documentElement.scrollWidth <= innerWidth'),(theme,width)
   page.screenshot(path=str(ROOT/f'samples/giao-dien-{theme}-mobile-v4.2.png'),full_page=True)
   page.set_viewport_size({'width':1440,'height':1050});page.screenshot(path=str(ROOT/f'samples/giao-dien-{theme}-v4.2.png'),full_page=True)
  if real:
   page.reload();page.locator('#mode-images').click();page.locator('#image-files').set_input_files(str(ROOT/'samples/anh-nhieu-khach-hang.jpg'));page.locator('#analyze').click();page.wait_for_function("busy === ''",timeout=240000)
   result=page.evaluate('sheets');assert len(result)==5, page.locator('#status').inner_text()
   assert all(s['date']=='2026-10-02' for s in result)
   assert all(len([r for r in s['image_info']['source_rows'] if r.get('type')=='goods'])==17 for s in result)
   assert all(not s['selected'] and s['requires_review'] for s in result)
   assert all(s['errors'] for s in result), 'Dense source has uncertain numbers; these must not silently become valid.'
   page.locator('[data-customer-select="0"]').check();page.locator('#export').click();page.wait_for_function("busy === ''");expect(page.locator('#status')).to_have_class('status error')
   page.screenshot(path=str(ROOT/'samples/anh-mau-5-khach-v4.2.png'),full_page=True)
   print('Real Tesseract: 5 customers × 17 source rows, date preserved, uncertain export blocked.')
  assert not errors,errors
  browser.close()
 print('Advanced browser checks passed: crop/reset, blank margins, customer-only export, sibling draft/edit retention, theme persistence and 320–1440px layouts.')

if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('--base',default='http://127.0.0.1:8095');parser.add_argument('--real',action='store_true');args=parser.parse_args();run(args.base,args.real)
