"""Stateless Vercel endpoints; every request carries its own Excel sources.

Vercel may use a different process for each invocation. Source workbooks and
manual mappings therefore stay in the browser and are resent on remap/export.
The desktop server in app.py continues to use its existing session transport.
"""
from __future__ import annotations

import io
import json
import secrets
import sys
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app import EXCEL_TYPE, VERSION, IMAGE_CAPABILITIES, analyze_images, apply_settings, check_excel, create_download
from engine import analyze, validate_template

# Keep a margin below Vercel's 4.5 MB request/response limit, including multipart
# boundaries and JSON metadata. These limits are advertised to the browser.
MAX_UPLOAD = 4 * 1024 * 1024
MAX_RESPONSE = 4 * 1024 * 1024
TRANSPORT = {
    'transport': 'stateless',
    'max_upload_bytes': MAX_UPLOAD,
    'max_response_bytes': MAX_RESPONSE,
}
MAPPING_FIELDS = {'name', 'quantity', 'price', 'amount', 'unit', 'code', 'note'}
UPLOAD_ERROR = 'Dung lượng yêu cầu vượt 4 MB của bản trực tuyến. Hãy chia workbook hoặc giảm dung lượng dữ liệu và template.'


class UploadTooLarge(ValueError):
    pass


def read_request_body(stream, headers) -> bytes:
    """Support Vercel's chunked proxy without waiting for socket EOF.

    The BaseHTTPRequestHandler runtime forwards raw HTTP framing, unlike its
    WSGI adapter. Local clients usually send Content-Length; Fluid Compute may
    forward Transfer-Encoding: chunked even for small browser uploads.
    """
    length_header = headers.get('Content-Length')
    transfer = (headers.get('Transfer-Encoding') or '').strip().lower()
    if transfer:
        if transfer != 'chunked' or length_header is not None:
            raise ValueError('Cách truyền file không hợp lệ. Hãy thử tải lại.')
        chunks, total = [], 0
        while True:
            line = stream.readline(8193)
            if len(line) > 8192 or not line.endswith(b'\r\n'):
                raise ValueError('File tải lên chưa đầy đủ. Hãy thử lại.')
            token = line[:-2].split(b';', 1)[0].strip()
            if not token or any(char not in b'0123456789abcdefABCDEF' for char in token):
                raise ValueError('Cách truyền file không hợp lệ. Hãy thử tải lại.')
            size = int(token, 16)
            if size == 0:
                trailer_bytes = 0
                while True:
                    trailer = stream.readline(8193)
                    trailer_bytes += len(trailer)
                    if len(trailer) > 8192 or trailer_bytes > 16384 or not trailer.endswith(b'\r\n'):
                        raise ValueError('Phần kết thúc file tải lên không hợp lệ.')
                    if trailer == b'\r\n':
                        return b''.join(chunks)
            total += size
            if total > MAX_UPLOAD:
                raise UploadTooLarge(UPLOAD_ERROR)
            data = stream.read(size)
            if len(data) != size or stream.read(2) != b'\r\n':
                raise ValueError('File tải lên chưa đầy đủ. Hãy thử lại.')
            chunks.append(data)
    if length_header is not None:
        try:
            length = int(length_header)
        except (ValueError, TypeError):
            raise ValueError('Dung lượng file tải lên không hợp lệ.') from None
        if length < 0:
            raise ValueError('Dung lượng file tải lên không hợp lệ.')
        if length > MAX_UPLOAD:
            raise UploadTooLarge(UPLOAD_ERROR)
        body = stream.read(length)
        if len(body) != length:
            raise ValueError('File tải lên chưa đầy đủ. Hãy thử lại.')
        return body
    if isinstance(stream, io.BytesIO):
        # Some adapters supply an already buffered body without HTTP framing.
        # Never read until EOF on a live socket: the client waits for a reply.
        body = stream.read(MAX_UPLOAD + 1)
        if len(body) > MAX_UPLOAD:
            raise UploadTooLarge(UPLOAD_ERROR)
        return body
    return b''


def read_parts(body: bytes, content_type: str, image_input=False) -> tuple[dict, str]:
    """Read multipart files without cgi or writes to a shared filesystem."""
    message = BytesParser(policy=default).parsebytes(
        ('Content-Type: ' + content_type + '\r\nMIME-Version: 1.0\r\n\r\n').encode()
        + body
    )
    if message.get_content_type() != 'multipart/form-data' or not message.is_multipart() or message.defects:
        raise ValueError('Hãy tải lại file Excel dữ liệu để tiếp tục.')
    parts = {}
    source_name = 'Du-lieu.xlsx'
    for part in message.iter_parts():
        name = part.get_param('name', header='content-disposition')
        if name not in {'data', 'template', 'payload', 'replace_token', 'ocr_documents', 'image_overrides'} or name in parts or part.is_multipart():
            raise ValueError('Các trường tải lên không hợp lệ hoặc bị trùng.')
        parts[name] = part.get_payload(decode=True)
        if name == 'data':
            source_name = Path((part.get_filename() or source_name).replace('\\', '/')).name[:180]
    if image_input and not parts.get('ocr_documents'):
        raise ValueError('Chưa có kết quả đọc ảnh. Hãy chọn ảnh và nhận diện lại.')
    if not image_input and not parts.get('data'):
        raise ValueError('Chưa chọn file Excel dữ liệu.')
    return parts, source_name


def read_payload(parts: dict) -> dict:
    try:
        payload = json.loads(parts.get('payload', b''))
    except (ValueError, TypeError, UnicodeDecodeError):
        raise ValueError('Thông tin phiếu không hợp lệ. Hãy nhận diện lại file.') from None
    if not isinstance(payload, dict):
        raise ValueError('Thông tin phiếu phải là một đối tượng.')
    return payload


def read_overrides(payload: dict) -> dict:
    overrides = payload.get('overrides', {})
    if not isinstance(overrides, dict):
        raise ValueError('Cấu hình cột không hợp lệ.')
    validated = {}
    for name, entry in overrides.items():
        if not isinstance(name, str) or not name:
            raise ValueError('Tên sheet trong cấu hình cột không hợp lệ.')
        # A null entry restores automatic recognition of this sheet.
        if entry is None:
            continue
        if not isinstance(entry, dict) or type(entry.get('header')) is not int or entry['header'] < 1:
            raise ValueError('Dòng tiêu đề phải là số nguyên dương.')
        span = entry.get('header_span', 1)
        if type(span) is not int or span not in (1, 2, 3):
            raise ValueError('Số dòng tiêu đề phải từ 1 đến 3.')
        mapping = entry.get('mapping')
        if not isinstance(mapping, dict) or set(mapping) - MAPPING_FIELDS or any(not isinstance(column, str) for column in mapping.values()):
            raise ValueError('Cấu hình cột không hợp lệ.')
        validated[name] = {**entry, 'header_span': span, 'mapping': mapping}
    return validated


class handler(BaseHTTPRequestHandler):
    """Vercel's Python runtime discovers the lowercase handler class."""

    def respond(self, status, data, content_type='application/json; charset=utf-8', filename=None):
        if isinstance(data, (dict, list)):
            data = json.dumps(data, ensure_ascii=False, allow_nan=False).encode()
        if isinstance(data, str):
            data = data.encode()
        if len(data) > MAX_RESPONSE:
            status = 413
            filename = None
            content_type = 'application/json; charset=utf-8'
            data = json.dumps({
                'error': 'Kết quả vượt giới hạn 4 MB của bản trực tuyến. Hãy chia file dữ liệu hoặc chọn ít sheet hơn để xuất.'
            }, ensure_ascii=False).encode()
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
        path = self.path.split('?', 1)[0]
        if path == '/api/health':
            return self.respond(200, {'app': 'xuat-kho', 'status': 'ok', 'version': VERSION, 'hosting': 'vercel', **TRANSPORT, **IMAGE_CAPABILITIES})
        if path == '/template':
            return self.respond(200, (ROOT / 'templates/Template.xlsx').read_bytes(), EXCEL_TYPE, 'Template.xlsx')
        if path == '/favicon.ico':
            return self.respond(200, '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect width="64" height="64" rx="16" fill="#186851"/><text x="32" y="42" text-anchor="middle" fill="white" font-size="30" font-family="Arial">HA</text></svg>', 'image/svg+xml')
        self.respond(404, {'error': 'Không tìm thấy trang.'})

    def do_POST(self):
        path = self.path.split('?', 1)[0]
        if path not in {'/api/analyze', '/api/image-analyze', '/api/remap', '/api/export'}:
            return self.respond(404, {'error': 'Không tìm thấy chức năng.'})
        try:
            body = read_request_body(self.rfile, self.headers)
            if not body:
                raise ValueError('Chưa có file Excel dữ liệu.')
            parts, source_name = read_parts(body, self.headers.get('Content-Type', ''), path == '/api/image-analyze')
            if path == '/api/image-analyze':
                _, _, sheets, image_response = analyze_images(parts)
                return self.respond(200, {'token': secrets.token_urlsafe(24), 'sheets': sheets,
                                          'version': VERSION, **image_response, **TRANSPORT})
            source = parts['data']
            template = parts.get('template') or (ROOT / 'templates/Template.xlsx').read_bytes()
            check_excel(source)
            check_excel(template)
            validate_template(template)
            payload = {} if path == '/api/analyze' else read_payload(parts)
            overrides = read_overrides(payload)
            sheets = analyze(source, overrides)
            if set(overrides) - {sheet['sheet'] for sheet in sheets}:
                raise ValueError('Cấu hình cột có sheet không thuộc file dữ liệu.')
            if path == '/api/analyze':
                return self.respond(200, {
                    'token': secrets.token_urlsafe(24), 'sheets': sheets,
                    'filename': source_name, 'version': VERSION, **TRANSPORT,
                })
            if path == '/api/remap':
                return self.respond(200, {'sheets': sheets, **TRANSPORT})
            sheets = apply_settings(sheets, payload.get('sheets', []))
            data, content_type, filename = create_download(
                template, sheets, payload.get('exclude_zero') is True, payload.get('format', 'xlsx')
            )
            return self.respond(200, data, content_type, filename)
        except UploadTooLarge as exc:
            self.respond(413, {'error': str(exc)})
        except Exception as exc:
            self.respond(400, {'error': str(exc) or 'Không thể xử lý file. Hãy kiểm tra Excel.'})

    def log_message(self, fmt, *args):
        # Workbook contents and filenames are never written to deployment logs.
        pass
