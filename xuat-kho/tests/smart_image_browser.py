"""Read tilted/dim synthetic invoices through real local OCR and export them."""
import argparse
import base64
from pathlib import Path
import tempfile
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

ROOT=Path(__file__).resolve().parents[1]
EXPECTED=[('Gạo tẻ',2,20000,40000),('Thịt lợn',3,80000,240000),('Rau cải',5,12000,60000)]

def run(base):
    with tempfile.TemporaryDirectory() as directory, sync_playwright() as p:
        browser=p.chromium.launch(headless=True,executable_path='/usr/bin/chromium',args=['--no-sandbox'])
        page=browser.new_page(viewport={'width':1440,'height':1050},locale='vi-VN')
        errors=[];page.on('pageerror',lambda e:errors.append(str(e)))
        source=base64.b64encode((ROOT/'samples/anh-minh-hoa.png').read_bytes()).decode()
        for angle,dim in [(3,False),(-4,True)]:
            page.goto(base)
            result=page.evaluate('''async ({source,angle,dim})=>{
              const image=new Image();image.src='data:image/png;base64,'+source;await image.decode();
              const c=document.createElement('canvas');c.width=image.width+150;c.height=image.height+150;
              const x=c.getContext('2d');x.fillStyle='white';x.fillRect(0,0,c.width,c.height);
              x.translate(c.width/2,c.height/2);x.rotate(angle*Math.PI/180);x.drawImage(image,-image.width/2,-image.height/2);
              if(dim){const pixels=x.getImageData(0,0,c.width,c.height);for(let i=0;i<pixels.data.length;i+=4)for(let j=0;j<3;j++)pixels.data[i+j]=pixels.data[i+j]*.55+60;x.putImageData(pixels,0,0);}
              const prepared=ImageRegions.prepare(c),quality=ImageRegions.inspect(prepared.canvas);
              prepared.canvas.width=prepared.canvas.height=1;
              return {image:c.toDataURL('image/png'),quality};
            }''',{'source':source,'angle':angle,'dim':dim})
            assert abs(result['quality']['angle']+angle)<.3,result
            assert result['quality']['contrast']==dim,result
            file=Path(directory)/f'nghieng-{angle}.png';file.write_bytes(base64.b64decode(result['image'].split(',',1)[1]))
            page.locator('#mode-images').click()
            assert page.locator('#ocr-enhance').is_checked()
            page.locator('#image-files').set_input_files(str(file))
            with page.expect_response(lambda response:response.url.endswith('/api/image-analyze'),timeout=240000) as pending:
                page.locator('#analyze').click()
            response=pending.value;assert response.status==200,response.text()
            page.wait_for_function("busy === ''")
            sheet=response.json()['sheets'][0]
            actual=[(item['name'],item['quantity'],item['price'],item['amount']) for item in sheet['items']]
            assert actual==EXPECTED,(angle,dim,actual,sheet['errors'])
            assert sheet['date']=='2026-10-01' and sheet['customer']=='Bếp ăn Minh An'
            assert not sheet['selected'] and sheet['requires_review']
            assert 'Đã chỉnh nghiêng' in page.locator('.image-reference').inner_text()
            if dim:assert 'tăng tương phản' in page.locator('.image-reference').inner_text()
            page.locator('[data-image-confirm]').check()
            with page.expect_download() as download:page.locator('#export').click()
            output=Path(directory)/'output.xlsx';download.value.save_as(output)
            assert load_workbook(output).active['H27'].value==340000
            print(f'Real OCR passed: tilt {angle}°, dim={dim}, exact 3 goods and exported total 340000.')
        page.goto(base)
        # A dense shared-price table must retain all customer groups after deskew.
        dense=base64.b64encode((ROOT/'samples/anh-nhieu-khach-hang.jpg').read_bytes()).decode()
        encoded=page.evaluate('''async source=>{
          const image=new Image();image.src='data:image/jpeg;base64,'+source;await image.decode();
          const c=document.createElement('canvas');c.width=image.width+100;c.height=image.height+100;
          const x=c.getContext('2d');x.fillStyle='white';x.fillRect(0,0,c.width,c.height);
          x.translate(c.width/2,c.height/2);x.rotate(3*Math.PI/180);x.drawImage(image,-image.width/2,-image.height/2);
          return c.toDataURL('image/png').split(',')[1];
        }''',dense)
        file=Path(directory)/'nhieu-khach-nghieng.png';file.write_bytes(base64.b64decode(encoded))
        page.locator('#mode-images').click();page.locator('#image-files').set_input_files(str(file))
        with page.expect_response(lambda response:response.url.endswith('/api/image-analyze'),timeout=240000) as pending:
            page.locator('#analyze').click()
        response=pending.value;assert response.status==200,response.text()
        page.wait_for_function("busy === ''")
        sheets=response.json()['sheets']
        assert len(sheets)==5,[(s['customer'],len(s['items'])) for s in sheets]
        assert all(len([r for r in s['image_info']['source_rows'] if r.get('type')=='goods'])==17 for s in sheets)
        assert all(s['requires_review'] and not s['selected'] and s['errors'] for s in sheets)
        assert all(any('tiêu đề Tổng' in w for w in s['image_info']['warnings']) for s in sheets)
        print('Real dense tilted OCR passed: 5 customers × 17 source rows; inferred headers disclosed, uncertain data blocked.')
        # White/blank canvases and a straight photo should not receive guessed skew.
        assert page.evaluate('''()=>{const c=document.createElement('canvas');c.width=900;c.height=600;const x=c.getContext('2d');x.fillStyle='white';x.fillRect(0,0,900,600);return ImageRegions.enhance(c).canvas===c;}''')
        assert not errors,errors
        browser.close()

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--base',default='http://127.0.0.1:8095');run(parser.parse_args().base)
