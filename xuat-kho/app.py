from __future__ import annotations
import argparse
import copy
import io
import json
import secrets
import threading
import time
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from zipfile import ZipFile, BadZipFile
from engine import analyze, export, validate_template

ROOT=Path(__file__).parent
SESSIONS={}
LOCK=threading.Lock()
MAX_UPLOAD=25*1024*1024


def check_excel(data):
    try:
        with ZipFile(io.BytesIO(data)) as z:
            if sum(i.file_size for i in z.infolist())>100*1024*1024: raise ValueError('Excel giải nén quá lớn (giới hạn 100 MB).')
            if '[Content_Types].xml' not in z.namelist(): raise ValueError('File không phải Excel .xlsx.')
    except BadZipFile: raise ValueError('Chỉ hỗ trợ file .xlsx. Với .xls, hãy lưu lại dạng .xlsx trong Excel.')


class Handler(BaseHTTPRequestHandler):
    def respond(self,status,data,content_type='application/json; charset=utf-8',filename=None):
        if isinstance(data,(dict,list)): data=json.dumps(data,ensure_ascii=False).encode()
        if isinstance(data,str): data=data.encode()
        self.send_response(status)
        self.send_header('Content-Type',content_type)
        self.send_header('Content-Length',str(len(data)))
        self.send_header('Cache-Control','no-store')
        self.send_header('X-Content-Type-Options','nosniff')
        if filename: self.send_header('Content-Disposition',f'attachment; filename="{filename}"')
        self.end_headers(); self.wfile.write(data)

    def do_GET(self):
        files={'/':('static/index.html','text/html; charset=utf-8'),'/app.js':('static/app.js','application/javascript; charset=utf-8'),'/style.css':('static/style.css','text/css; charset=utf-8'),'/template':('templates/Template.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        path=self.path.split('?')[0]
        if path not in files: return self.respond(404,{'error':'Không tìm thấy trang.'})
        filename,content_type=files[path]
        self.respond(200,(ROOT/filename).read_bytes(),content_type,'Template.xlsx' if path=='/template' else None)

    def do_POST(self):
        try:
            length=int(self.headers.get('Content-Length',0))
            if length<=0 or length>MAX_UPLOAD: raise ValueError('Dung lượng tải lên phải từ 1 byte đến 25 MB.')
            body=self.rfile.read(length)
            now=time.time()
            with LOCK:
                for key in list(SESSIONS):
                    if now-SESSIONS[key]['time']>3600: del SESSIONS[key]
            if self.path=='/api/analyze':
                message=BytesParser(policy=default).parsebytes(('Content-Type: '+self.headers.get('Content-Type','')+'\r\nMIME-Version: 1.0\r\n\r\n').encode()+body)
                if not message.is_multipart(): raise ValueError('Hãy chọn file Excel dữ liệu.')
                parts={p.get_param('name',header='content-disposition'):p.get_payload(decode=True) for p in message.iter_parts()}
                source=parts.get('data')
                if not source: raise ValueError('Chưa chọn file Excel dữ liệu.')
                template=parts.get('template') or (ROOT/'templates/Template.xlsx').read_bytes()
                check_excel(source); check_excel(template); validate_template(template)
                sheets=analyze(source)
                token=secrets.token_urlsafe(24)
                with LOCK:
                    if len(SESSIONS)>=20: raise ValueError('Có quá nhiều phiên đang mở. Hãy khởi động lại ứng dụng hoặc chờ phiên cũ hết hạn.')
                    SESSIONS[token]={'source':source,'template':template,'sheets':sheets,'time':now}
                return self.respond(200,{'token':token,'sheets':sheets})
            payload=json.loads(body)
            with LOCK: session=SESSIONS.get(payload.get('token'))
            if not session: raise ValueError('Phiên làm việc hết hạn. Hãy tải lại file.')
            if self.path=='/api/remap':
                sheets=analyze(session['source'],payload.get('overrides',{}))
                with LOCK: session['sheets']=sheets; session['time']=now
                return self.respond(200,{'sheets':sheets})
            if self.path=='/api/export':
                sheets=copy.deepcopy(session['sheets'])
                settings={s['sheet']:s for s in payload.get('sheets',[])}
                for sheet in sheets:
                    settings_for_sheet=settings.get(sheet['sheet'],{})
                    for field in ['selected','customer','recipient','date','address','voucher','content']:
                        if field in settings_for_sheet: sheet[field]=settings_for_sheet[field]
                data=export(session['template'],sheets,bool(payload.get('exclude_zero')))
                return self.respond(200,data,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','Phieu-xuat-kho.xlsx')
            self.respond(404,{'error':'Không tìm thấy chức năng.'})
        except Exception as exc:
            self.respond(400,{'error':str(exc) or 'Không thể xử lý file. Hãy kiểm tra Excel.'})

    def log_message(self,fmt,*args):
        # Never log worksheet contents or request bodies.
        print(fmt % args)

def serve(host='127.0.0.1', port=8080, open_browser=False):
    server = ThreadingHTTPServer((host, port), Handler)
    url = f'http://{host}:{server.server_port}'
    print(f'Ứng dụng đã sẵn sàng: {url}', flush=True)
    print('Giữ cửa sổ này mở khi sử dụng. Nhấn Ctrl+C để dừng.', flush=True)
    if open_browser:
        import webbrowser
        threading.Timer(0.5, webbrowser.open, args=(url,)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()

if __name__=='__main__':
    parser=argparse.ArgumentParser(description='Xuất kho từ Excel nhiều sheet')
    parser.add_argument('--host',default='127.0.0.1'); parser.add_argument('--port',type=int,default=8080)
    args=parser.parse_args()
    serve(args.host, args.port)
