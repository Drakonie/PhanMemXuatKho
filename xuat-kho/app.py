from __future__ import annotations
import argparse
import base64
import copy
import io
import json
import math
import mimetypes
import re
import secrets
import threading
import time
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from zipfile import ZipFile, BadZipFile, ZIP_DEFLATED
from engine import analyze, export, validate_template, number
from image_tables import validate_table

ROOT = Path(__file__).resolve().parent
VERSION = '4.3.0'
SESSIONS = {}
LOCK = threading.RLock()
MAX_UPLOAD = 25 * 1024 * 1024
SESSION_TTL = 3600
EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
IMAGE_CAPABILITIES = {'capabilities': {'image_input': True}, 'ocr_engine': 'tesseract-browser'}


def check_excel(data):
    try:
        with ZipFile(io.BytesIO(data)) as archive:
            entries = archive.infolist()
            if sum(item.file_size for item in entries) > 100 * 1024 * 1024:
                raise ValueError('Excel giải nén quá lớn (giới hạn 100 MB).')
            if '[Content_Types].xml' not in archive.namelist() or 'xl/workbook.xml' not in archive.namelist():
                raise ValueError('File không phải Excel .xlsx.')
    except BadZipFile:
        raise ValueError('Chỉ hỗ trợ file .xlsx. Với .xls, hãy lưu lại dạng .xlsx trong Excel.')


def clean_sessions():
    now = time.time()
    with LOCK:
        for token in list(SESSIONS):
            if now - SESSIONS[token]['time'] > SESSION_TTL:
                del SESSIONS[token]


def valid_text(value, label, maximum=1000):
    if value is None:
        return ''
    if not isinstance(value, str):
        raise ValueError(f'{label} phải là văn bản.')
    value = value.strip()
    if len(value) > maximum or any(ord(char) < 32 and char not in '\n\t\r' for char in value):
        raise ValueError(f'{label} quá dài hoặc chứa ký tự không hợp lệ.')
    return value


def read_ocr_documents(data):
    """Validate browser OCR results before constructing the source workbook."""
    try:
        documents = json.loads(data or b'')
    except (ValueError, TypeError, UnicodeDecodeError):
        raise ValueError('Kết quả đọc ảnh không hợp lệ. Hãy chọn ảnh và nhận diện lại.') from None
    if not isinstance(documents, list) or not 1 <= len(documents) <= 10:
        raise ValueError('Mỗi lần nhận diện cần từ 1 đến 10 ảnh.')
    validated = []
    for index, document in enumerate(documents, 1):
        if not isinstance(document, dict):
            raise ValueError(f'Kết quả ảnh {index} không hợp lệ.')
        filename = valid_text(document.get('filename', f'Anh-{index}.png'), 'Tên ảnh', 180)
        filename = Path(filename.replace('\\', '/')).name or f'Anh-{index}.png'
        text = valid_text(document.get('text', ''), f'Ảnh {index}: văn bản OCR', 1_000_000)
        tsv = valid_text(document.get('tsv', ''), f'Ảnh {index}: tọa độ OCR', 6_000_000)
        if not text and not tsv:
            raise ValueError(f'Ảnh {index}: chưa có kết quả OCR. Hãy nhận diện lại ảnh rõ hơn.')
        if tsv.count('\n') > 55_000:
            raise ValueError(f'Ảnh {index}: quá nhiều từ OCR; hãy chia ảnh thành các phần nhỏ hơn.')
        words = sum(line.split('\t', 1)[0] == '5' for line in tsv.splitlines())
        if words > 50_000:
            raise ValueError(f'Ảnh {index}: vượt 50.000 từ OCR.')
        dimensions = {}
        for key in ('width', 'height'):
            value = document.get(key)
            if value is not None:
                if type(value) not in (int, float) or not math.isfinite(value) or not 0 < value <= 32768:
                    raise ValueError(f'Ảnh {index}: kích thước không hợp lệ.')
                dimensions[key] = value
        rotation = document.get('rotation', 0)
        if type(rotation) not in (int, float) or not math.isfinite(rotation) or not -360 <= rotation <= 360:
            raise ValueError(f'Ảnh {index}: góc xoay không hợp lệ.')
        language = valid_text(document.get('language', 'vie+eng'), 'Ngôn ngữ OCR', 40)
        customer_edits = document.get('customer_edits', {})
        if not isinstance(customer_edits, dict) or len(customer_edits) > 30:
            raise ValueError(f'Ảnh {index}: bản sửa theo khách hàng không hợp lệ (tối đa 30 nhóm).')
        clean_edits = {}
        for group, content in customer_edits.items():
            if not isinstance(group, str) or not group.isascii() or not group.isdigit() or not 1 <= int(group) <= 30:
                raise ValueError(f'Ảnh {index}: chỉ số nhóm khách hàng không hợp lệ.')
            key = str(int(group))
            if key in clean_edits:
                raise ValueError(f'Ảnh {index}: nhóm khách hàng bị trùng.')
            clean_edits[key] = valid_text(content, f'Ảnh {index}: bản sửa khách hàng {key}', 1_000_000)
        if sum(len(content) for content in clean_edits.values()) > 1_000_000:
            raise ValueError(f'Ảnh {index}: tổng bản sửa theo khách hàng vượt 1 triệu ký tự.')
        validated.append({'filename': filename, 'text': text, 'tsv': tsv, **dimensions,
                          'rotation': rotation, 'language': language,
                          'edited_text': document.get('edited_text') is True,
                          'customer_edits': clean_edits,
                          'table_data': validate_table(document.get('table_data'))})
    return validated


def decorate_image_sheets(sheets, report):
    """Keep OCR evidence separate from the engine's table recognition score."""
    image_info = report.get('image_info', {})
    if not isinstance(image_info, dict):
        image_info = {}
    document_info = {entry.get('sheet'): entry for entry in report.get('documents', []) if isinstance(entry, dict)}
    for sheet in sheets:
        info = image_info.get(sheet['sheet']) or document_info.get(sheet['sheet'], {})
        sheet.update(source_type='image', sourceType='image', image_info=info, requires_review=True,
                     selected=False, confidence_kind='table-structure', ocr_confidence=info.get('ocr_confidence'))
        if not info.get('metadata', {}).get('recipient'):
            sheet['recipient'] = ''
        warning = 'Dữ liệu được đọc từ ảnh. Hãy đối chiếu tên hàng, số lượng, đơn giá và thành tiền với ảnh trước khi chọn xuất.'
        for message in [warning, *info.get('warnings', [])]:
            if message and message not in sheet['warnings']:
                sheet['warnings'].append(message)
        for issue in info.get('issues', []):
            if not isinstance(issue, dict) or not issue.get('message'):
                continue
            if issue not in sheet['issues']:
                sheet['issues'].append(issue)
            category = 'errors' if issue.get('level') == 'error' else 'warnings'
            if issue['message'] not in sheet[category]:
                sheet[category].append(issue['message'])
        note = 'Điểm nhận diện bảng và độ tin cậy OCR được hiển thị riêng; OCR có thể đọc sai dấu hoặc chữ số.'
        if note not in sheet['detection_notes']:
            sheet['detection_notes'].append(note)
    return sheets


def read_image_overrides(data):
    if not data:
        return {}
    try:
        overrides = json.loads(data)
    except (ValueError, TypeError, UnicodeDecodeError):
        raise ValueError('Cấu hình cột ảnh không hợp lệ.') from None
    if not isinstance(overrides, dict) or len(overrides) > 10:
        raise ValueError('Cấu hình cột ảnh không hợp lệ.')
    validated = {}
    fields = {'name', 'quantity', 'price', 'amount', 'unit', 'code', 'note'}
    for name, entry in overrides.items():
        if not isinstance(name, str) or not name or not isinstance(entry, dict):
            raise ValueError('Cấu hình sheet ảnh không hợp lệ.')
        header, span, mapping = entry.get('header'), entry.get('header_span', 1), entry.get('mapping')
        if type(header) is not int or not 1 <= header <= 100000 or type(span) is not int or span not in (1, 2, 3):
            raise ValueError('Dòng tiêu đề ảnh không hợp lệ.')
        if not isinstance(mapping, dict) or set(mapping) - fields:
            raise ValueError('Cột nhận diện ảnh không hợp lệ.')
        if any(value not in (None, '') and (not isinstance(value, str) or not re.fullmatch(r'[A-Z]{1,3}', value)) for value in mapping.values()):
            raise ValueError('Cột nhận diện ảnh không hợp lệ.')
        validated[name] = {'header': header, 'header_span': span, 'mapping': mapping}
    return validated


def analyze_images(parts):
    documents = read_ocr_documents(parts.get('ocr_documents'))
    from image_import import build_image_workbook
    template = parts.get('template') or (ROOT / 'templates/Template.xlsx').read_bytes()
    check_excel(template)
    validate_template(template)
    source, report = build_image_workbook(documents)
    check_excel(source)
    overrides = read_image_overrides(parts.get('image_overrides'))
    sheets = decorate_image_sheets(analyze(source, overrides), report)
    if set(overrides) - {sheet['sheet'] for sheet in sheets}:
        raise ValueError('Cấu hình cột có sheet không thuộc các ảnh đã chọn.')
    return source, template, sheets, {
        'filename': 'Du-lieu-tu-anh.xlsx', 'source_filename': 'Du-lieu-tu-anh.xlsx',
        'source_type': 'image', 'sourceType': 'image', 'source_base64': base64.b64encode(source).decode('ascii'),
        'image_report': report, 'image_info': report.get('image_info', {}),
        'ocr_warnings': report.get('warnings', []),
        'overrides': overrides,
    }


def apply_settings(original, settings):
    """Validate every edited row; edits cannot omit unreadable source rows."""
    sheets = copy.deepcopy(original)
    if not isinstance(settings, list):
        raise ValueError('Danh sách phiếu không hợp lệ.')
    entries = {}
    for entry in settings:
        if not isinstance(entry, dict) or not isinstance(entry.get('sheet'), str):
            raise ValueError('Thông tin sheet không hợp lệ.')
        if entry['sheet'] in entries:
            raise ValueError('Danh sách sheet bị trùng.')
        entries[entry['sheet']] = entry
    known = {sheet['sheet'] for sheet in sheets}
    if set(entries) - known:
        raise ValueError('Có sheet không thuộc file đã nhận diện.')
    for sheet in sheets:
        entry = entries.get(sheet['sheet'], {})
        sheet['selected'] = entry.get('selected', False) is True
        if not sheet['selected']:
            continue
        for key in ['customer', 'recipient', 'date', 'address', 'voucher', 'content']:
            if key in entry:
                sheet[key] = valid_text(entry[key], key)
        if not sheet.get('customer', '').strip():
            raise ValueError(f"Sheet {sheet['sheet']}: chưa nhập tên khách hàng.")
        try:
            datetime.strptime(sheet.get('date', ''), '%Y-%m-%d')
        except (ValueError, TypeError):
            raise ValueError(f"Sheet {sheet['sheet']}: ngày xuất không hợp lệ.")
        if 'items' not in entry:
            # Original engine errors remain authoritative without explicit edits.
            continue
        edited = entry['items']
        if not isinstance(edited, list) or len(edited) > 100000:
            raise ValueError('Danh sách hàng không hợp lệ hoặc vượt 100.000 dòng.')
        originals = {item['row']: item for item in sheet['items']}
        ids = [item.get('row') for item in edited if isinstance(item, dict)]
        if len(ids) != len(edited) or any(type(row) is not int for row in ids) or len(set(ids)) != len(ids) or set(ids) != set(originals):
            raise ValueError(f"Sheet {sheet['sheet']}: không được bỏ hoặc thêm dòng nguồn khi chỉnh dữ liệu.")
        exclusions = entry.get('excluded_rows', [])
        if not isinstance(exclusions, list) or any(type(row) is not int for row in exclusions) or len(set(exclusions)) != len(exclusions) or set(exclusions) - set(originals):
            raise ValueError(f"Sheet {sheet['sheet']}: danh sách dòng bỏ chọn không hợp lệ.")
        excluded = set(exclusions)
        unresolved = []
        if sheet.get('issues'):
            unresolved = [issue['message'] for issue in sheet['issues'] if issue.get('level') == 'error' and not issue.get('recoverable')]
        else:
            for error in sheet.get('errors', []):
                match = re.search(r'Dòng\s+(\d+)', error)
                if not match or int(match[1]) not in originals:
                    unresolved.append(error)
        if unresolved:
            raise ValueError(f"Sheet {sheet['sheet']}: {unresolved[0]} Hãy chỉnh cột nhận diện hoặc sửa file nguồn rồi tải lại.")
        validated = []
        for item in edited:
            row = item['row']
            if row in excluded:
                continue
            result = {'row': row}
            for field in ['name', 'unit', 'code', 'note']:
                result[field] = valid_text(item.get(field, originals[row].get(field, '')), f'Dòng {row}: {field}')
            if not result['name']:
                raise ValueError(f'Dòng {row}: chưa nhập tên hàng.')
            for field, label in [('quantity', 'số lượng'), ('price', 'đơn giá'), ('amount', 'thành tiền')]:
                parsed = number(item.get(field))
                if parsed is None or parsed < 0 or parsed >= Decimal('1e18'):
                    raise ValueError(f'Dòng {row}: {label} phải là số không âm nhỏ hơn 10¹⁸.')
                result[field] = parsed
            calculated = (result['quantity'] * result['price']).quantize(Decimal(1), rounding=ROUND_HALF_UP)
            rounded_amount = result['amount'].quantize(Decimal(1), rounding=ROUND_HALF_UP)
            if abs(calculated - rounded_amount) > 1:
                raise ValueError(f'Dòng {row}: thành tiền khác số lượng × đơn giá. Hãy sửa số lượng, đơn giá hoặc thành tiền.')
            result.update(quantity=float(result['quantity']), price=float(result['price']), amount=int(rounded_amount))
            validated.append(result)
        sheet['items'] = validated
        sheet['total'] = sum(item['amount'] for item in validated)
        sheet['errors'] = []
        sheet['issues'] = [issue for issue in sheet.get('issues', []) if issue.get('level') != 'error']
    return sheets


def zip_name(name, used):
    cleaned = re.sub(r'[\\/:*?"<>|\x00-\x1f]', '-', name).strip(' .')[:100] or 'Phieu-xuat-kho'
    if cleaned.split('.', 1)[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *[f'COM{i}' for i in range(1, 10)], *[f'LPT{i}' for i in range(1, 10)]}:
        cleaned = 'Phieu-' + cleaned
    candidate, index = cleaned, 2
    while candidate.casefold() in used:
        candidate = f'{cleaned}-{index}'
        index += 1
    used.add(candidate.casefold())
    return candidate + '.xlsx'


def create_download(template, sheets, exclude_zero=False, format='xlsx'):
    if format == 'xlsx':
        return export(template, sheets, exclude_zero), EXCEL_TYPE, 'Phieu-xuat-kho.xlsx'
    if format != 'zip':
        raise ValueError('Định dạng xuất không hợp lệ.')
    chosen = [sheet for sheet in sheets if sheet.get('selected')]
    if not chosen:
        raise ValueError('Chọn ít nhất một sheet để xuất.')
    stream = io.BytesIO()
    used = set()
    with ZipFile(stream, 'w', ZIP_DEFLATED) as archive:
        for sheet in chosen:
            archive.writestr(zip_name(sheet['sheet'], used), export(template, [sheet], exclude_zero))
    return stream.getvalue(), 'application/zip', 'Phieu-xuat-kho.zip'


class Handler(BaseHTTPRequestHandler):
    def respond(self, status, data, content_type='application/json; charset=utf-8', filename=None):
        if isinstance(data, (dict, list)):
            data = json.dumps(data, ensure_ascii=False, allow_nan=False).encode()
        if isinstance(data, str):
            data = data.encode()
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        if filename:
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        path = self.path.split('?')[0]
        if path == '/api/health':
            return self.respond(200, {'app': 'xuat-kho', 'status': 'ok', 'version': VERSION, **IMAGE_CAPABILITIES})
        if path == '/favicon.ico':
            return self.respond(200, '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect width="64" height="64" rx="16" fill="#186851"/><text x="32" y="42" text-anchor="middle" fill="white" font-size="30" font-family="Arial">HA</text></svg>', 'image/svg+xml')
        asset_folder = next((folder for folder in ('ocr', 'theme') if path.startswith(f'/{folder}/')), None)
        if asset_folder:
            # Serve bundled assets only; a route must never escape its own folder.
            asset_root = (ROOT / 'static' / asset_folder).resolve()
            asset = (asset_root / path[len(asset_folder)+2:]).resolve()
            if not asset.is_relative_to(asset_root) or not asset.is_file():
                return self.respond(404, {'error': 'Không tìm thấy tài nguyên OCR.'})
            content_type = {'js': 'application/javascript', 'wasm': 'application/wasm', 'gz': 'application/gzip'}.get(asset.suffix.lstrip('.')) or mimetypes.guess_type(str(asset))[0] or 'application/octet-stream'
            return self.respond(200, asset.read_bytes(), content_type)
        files = {'/': ('static/index.html', 'text/html; charset=utf-8'), '/app.js': ('static/app.js', 'application/javascript; charset=utf-8'), '/image-import.js': ('static/image-import.js', 'application/javascript; charset=utf-8'), '/style.css': ('static/style.css', 'text/css; charset=utf-8'), '/theme.css': ('static/theme.css', 'text/css; charset=utf-8'), '/template': ('templates/Template.xlsx', EXCEL_TYPE)}
        files.update({f'/{name}.js': (f'static/{name}.js', 'application/javascript; charset=utf-8') for name in ('image-regions', 'image-enhance', 'theme-switch')})
        if path not in files:
            return self.respond(404, {'error': 'Không tìm thấy trang.'})
        filename, content_type = files[path]
        if not (ROOT / filename).is_file():
            return self.respond(404, {'error': 'Không tìm thấy tài nguyên.'})
        self.respond(200, (ROOT / filename).read_bytes(), content_type, 'Template.xlsx' if path == '/template' else None)

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            if length <= 0 or length > MAX_UPLOAD:
                raise ValueError('Dung lượng tải lên phải từ 1 byte đến 25 MB.')
            body = self.rfile.read(length)
            clean_sessions()
            if self.path in ('/api/analyze', '/api/image-analyze'):
                message = BytesParser(policy=default).parsebytes(('Content-Type: ' + self.headers.get('Content-Type', '') + '\r\nMIME-Version: 1.0\r\n\r\n').encode() + body)
                if not message.is_multipart():
                    raise ValueError('Hãy chọn file Excel dữ liệu.')
                parts = {}
                source_name = 'Du-lieu.xlsx'
                for part in message.iter_parts():
                    key = part.get_param('name', header='content-disposition')
                    if key in parts or part.is_multipart():
                        raise ValueError('Các trường tải lên không hợp lệ hoặc bị trùng.')
                    parts[key] = part.get_payload(decode=True)
                    if key == 'data':
                        source_name = Path((part.get_filename() or source_name).replace('\\', '/')).name[:180]
                image_response = {}
                if self.path == '/api/image-analyze':
                    source, template, sheets, image_response = analyze_images(parts)
                    source_name = image_response['filename']
                else:
                    source = parts.get('data')
                    if not source:
                        raise ValueError('Chưa chọn file Excel dữ liệu.')
                    template = parts.get('template') or (ROOT / 'templates/Template.xlsx').read_bytes()
                    check_excel(source)
                    check_excel(template)
                    validate_template(template)
                    sheets = analyze(source)
                token = secrets.token_urlsafe(24)
                previous = (parts.get('replace_token') or b'').decode(errors='ignore')
                with LOCK:
                    if previous in SESSIONS:
                        del SESSIONS[previous]
                    if len(SESSIONS) >= 20:
                        raise ValueError('Có quá nhiều phiên đang mở. Hãy khởi động lại ứng dụng hoặc chờ phiên cũ hết hạn.')
                    SESSIONS[token] = {'source': source, 'template': template, 'sheets': sheets, 'overrides': image_response.get('overrides', {}), 'filename': source_name, 'time': time.time(),
                                       'image_report': image_response.get('image_report')}
                return self.respond(200, {'token': token, 'sheets': sheets, 'filename': source_name, 'version': VERSION, **image_response})
            payload = json.loads(body)
            if not isinstance(payload, dict):
                raise ValueError('Yêu cầu không hợp lệ.')
            with LOCK:
                session = SESSIONS.get(payload.get('token'))
                if session:
                    session = copy.copy(session)
            if not session:
                raise ValueError('Phiên làm việc hết hạn. Hãy tải lại file.')
            if self.path == '/api/remap':
                overrides = payload.get('overrides', {})
                if not isinstance(overrides, dict):
                    raise ValueError('Cấu hình cột không hợp lệ.')
                merged = {**session['overrides'], **overrides}
                for key in list(merged):
                    if merged[key] is None:
                        del merged[key]
                sheets = analyze(session['source'], merged)
                if session.get('image_report'):
                    decorate_image_sheets(sheets, session['image_report'])
                with LOCK:
                    if payload['token'] not in SESSIONS:
                        raise ValueError('Phiên làm việc hết hạn. Hãy tải lại file.')
                    SESSIONS[payload['token']].update(sheets=sheets, overrides=merged, time=time.time())
                return self.respond(200, {'sheets': sheets})
            if self.path == '/api/export':
                sheets = apply_settings(session['sheets'], payload.get('sheets', []))
                data, content_type, filename = create_download(session['template'], sheets, payload.get('exclude_zero') is True, payload.get('format', 'xlsx'))
                with LOCK:
                    if payload['token'] in SESSIONS:
                        SESSIONS[payload['token']]['time'] = time.time()
                return self.respond(200, data, content_type, filename)
            self.respond(404, {'error': 'Không tìm thấy chức năng.'})
        except Exception as exc:
            self.respond(400, {'error': str(exc) or 'Không thể xử lý file. Hãy kiểm tra Excel.'})

    def log_message(self, fmt, *args):
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


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Xuất kho từ Excel nhiều sheet')
    parser.add_argument('--host', default='127.0.0.1')
    parser.add_argument('--port', type=int, default=8080)
    args = parser.parse_args()
    serve(args.host, args.port)
