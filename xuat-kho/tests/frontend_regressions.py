"""Focused browser regressions for edits, parsing, cancellation and preview races.

Uses the real import/export API with a deterministic OCR worker. The separate
image_browser_smoke.py exercises bundled Tesseract and real Vietnamese OCR.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import tempfile

from openpyxl import load_workbook
from playwright.sync_api import expect, sync_playwright


ROOT = Path(__file__).resolve().parents[1]
OCR_TEXT = '''Khách hàng: Bếp ăn Minh An
Ngày xuất: 01/10/2026
Tên hàng | ĐVT | Số lượng | Đơn giá | Thành tiền
Gạo tẻ | kg | 2 | 20.000 | 40.000
Thịt lợn | kg | 3 | 80.000 | 240.000
Rau cải | kg | 5 | 12.000 | 60.000
'''


def idle(page):
    page.wait_for_function("busy === ''")


def begin_ocr(page):
    page.evaluate("""() => {
      window.__jobSettled = false;
      window.__job = ImageImport.analyze().finally(() => window.__jobSettled = true);
    }""")


def run(base='http://127.0.0.1:8096'):
    with tempfile.TemporaryDirectory() as directory, sync_playwright() as playwright:
        temporary = Path(directory)
        browser = playwright.chromium.launch(headless=True, executable_path='/usr/bin/chromium', args=['--no-sandbox'])
        page = browser.new_page(viewport={'width': 1440, 'height': 1100}, locale='vi-VN')
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.goto(base)
        page.locator('#data-file').set_input_files(str(ROOT / 'samples/data-mau.xlsx'))
        page.locator('#analyze').click()
        page.locator('#workspace').wait_for(state='visible')
        idle(page)

        # Decimal and grouping choices must not silently change the invoice.
        expected_numbers = {
            '1,234.50': 1234.5, '1.234,50': 1234.5, '0.001': 0.001,
            '0,001': 0.001, '1,234': 1.234, '1.234': 1234,
            '1 234,50': 1234.5, '1.2.3,4': None, '1 2': None,
        }
        actual = page.evaluate('(values) => values.map(value => numericValue(value))', list(expected_numbers))
        assert actual == list(expected_numbers.values()), actual
        page.locator('[data-item="0"][data-item-field="quantity"]').fill('1')
        page.locator('[data-item="0"][data-item-field="price"]').fill('1,234.50')
        assert page.evaluate('current().items[0].price') == 1234.5
        assert page.evaluate('current().items[0].amount') == 1235
        page.locator('#deselect-all').click()
        page.locator('[data-sheet-select="0"]').check()
        with page.expect_download() as pending:
            page.locator('#export').click()
        output = temporary / 'decimal.xlsx'
        pending.value.save_as(output)
        sheet = load_workbook(output).active
        assert sheet['G12'].value == 1234.5 and sheet['H12'].value == 1235
        idle(page)

        quantity = page.locator('[data-item="0"][data-item-field="quantity"]')
        quantity.fill('1.2.3,4')
        assert page.evaluate('current().items[0].quantity') == 1
        page.locator('[data-tab=source]').click()
        page.locator('[data-tab=items]').click()
        expect(page.locator('[data-item="0"][data-item-field="quantity"]')).to_have_value('1.2.3,4')
        page.locator('#export').click()
        expect(page.locator('#status')).to_have_class('status error')
        page.locator('[data-action=undo]').click()

        # An unsuccessful upload preserves edited data and the working token.
        page.locator('[data-item="0"][data-item-field="note"]').fill('Giữ chỉnh sửa sau lỗi tải lên')
        old_token = page.evaluate('token')
        bad = temporary / 'broken.xlsx'
        bad.write_bytes(b'not an Excel workbook')
        page.locator('#data-file').set_input_files(str(bad))
        page.locator('#analyze').click()
        expect(page.locator('#status')).to_have_class('status error')
        idle(page)
        assert page.evaluate('token') == old_token
        expect(page.locator('[data-item="0"][data-item-field="note"]')).to_have_value('Giữ chỉnh sửa sau lỗi tải lên')
        expect(page.locator('#export')).to_be_enabled()

        # Use real image analysis and Excel generation with deterministic text.
        page.evaluate("""(text) => {
          window.__ocrText = text;
          window.__holdRecognition = false;
          window.__holdTermination = true;
          window.__canvas = null;
          window.Tesseract = {createWorker: async () => ({
            setParameters: async () => {},
            recognize: async canvas => {
              window.__canvas = canvas;
              if (window.__holdRecognition) return new Promise(() => {});
              return {data: {text: window.__ocrText, confidence: 99}};
            },
            terminate: async () => {
              if (window.__holdTermination) await new Promise(resolve => window.__finishTermination = resolve);
            }
          })};
        }""", OCR_TEXT)
        second = temporary / 'anh-thu-hai.png'
        second.write_bytes((ROOT / 'samples/anh-minh-hoa.png').read_bytes())
        page.locator('#mode-images').click()
        page.locator('#image-files').set_input_files([str(ROOT / 'samples/anh-minh-hoa.png'), str(second)])
        begin_ocr(page)
        page.wait_for_function("typeof window.__finishTermination === 'function'")
        assert page.evaluate('busy') == 'image'
        # Newly rendered fields remain disabled until cleanup completes.
        expect(page.locator('[data-field=customer]')).to_be_disabled()
        expect(page.locator('[data-item="0"][data-item-field="quantity"]')).to_be_disabled()
        page.evaluate('window.__holdTermination = false; window.__finishTermination()')
        idle(page)
        expect(page.locator('[data-field=customer]')).to_be_enabled()
        expect(page.locator('#data-file')).to_be_disabled()
        assert page.locator('.selected:checked').count() == 0

        # Pending text, invalid number drafts, exclusions and review on B survive
        # a text correction and a column remap on A.
        page.locator('[data-open-sheet="1"]').click()
        page.locator('[data-field=customer]').fill('Khách đã sửa B')
        page.locator('[data-image-confirm]').check()
        page.locator('[data-item="0"][data-item-field="quantity"]').fill('chưa sửa xong')
        page.locator('[data-row-toggle="1"]').click()
        page.locator('[data-tab=source]').click()
        draft = page.locator('#ocr-text-draft').input_value() + '\nGhi chú đang sửa B'
        page.locator('#ocr-text-draft').fill(draft)
        page.locator('[data-open-sheet="0"]').click()
        page.locator('[data-tab=source]').click()
        page.locator('[data-reparse-ocr]').click()
        page.wait_for_function("document.getElementById('status').textContent.includes('Đã đọc lại văn bản')")
        idle(page)
        assert page.locator('.selected:checked').count() == 1
        assert page.evaluate('sheets[1]._ocrDraft') == draft
        assert page.evaluate('sheets[1].customer') == 'Khách đã sửa B'
        assert page.evaluate('sheets[1]._excluded.size') == 1
        assert page.evaluate('sheets[1].items[0]._draft.quantity') == 'chưa sửa xong'
        page.locator('[data-tab=mapping]').click()
        page.locator('[data-action=remap]').click()
        page.wait_for_function("document.getElementById('status').textContent.includes('Đã đọc lại sheet')")
        idle(page)
        page.locator('[data-open-sheet="1"]').click()
        expect(page.locator('[data-item="0"][data-item-field="quantity"]')).to_have_value('chưa sửa xong')
        page.locator('[data-tab=source]').click()
        expect(page.locator('#ocr-text-draft')).to_have_value(draft)

        # A worker that never settles recognize() must not hold the cancelled job
        # or its large canvas alive. Starting another OCR job must still work.
        old_token = page.evaluate('token')
        page.evaluate('window.__holdRecognition = true; window.__canvas = null')
        begin_ocr(page)
        page.wait_for_function('window.__canvas !== null')
        page.locator('#ocr-cancel').click()
        page.wait_for_function('window.__jobSettled === true')
        assert page.evaluate('[window.__canvas.width, window.__canvas.height]') == [1, 1]
        assert page.evaluate('token') == old_token
        assert page.evaluate('sheets[1]._ocrDraft') == draft
        idle(page)
        page.evaluate('window.__holdRecognition = false')
        begin_ocr(page)
        page.wait_for_function('window.__jobSettled === true')
        assert page.locator('.sheet-card').count() == 2
        assert page.locator('.selected:checked').count() == 0

        # Make rotated decode asynchronous and count preview URLs. A stale
        # preview cannot reopen a closed dialog or leak superseded URLs.
        page.evaluate("""() => {
          window.__originalDecode = Image.prototype.decode;
          window.__decodeReleases = [];
          Image.prototype.decode = function() {
            return new Promise(resolve => window.__decodeReleases.push(resolve))
              .then(() => window.__originalDecode.call(this));
          };
          window.__previewUrls = new Set();
          const create = URL.createObjectURL.bind(URL), revoke = URL.revokeObjectURL.bind(URL);
          URL.createObjectURL = blob => {const url = create(blob); window.__previewUrls.add(url); return url;};
          URL.revokeObjectURL = url => {window.__previewUrls.delete(url); revoke(url);};
        }""")
        page.locator('[data-rotate-image="0"]').click()
        page.locator('[data-preview-image="0"]').click()
        page.evaluate("document.getElementById('image-zoom-close').click()")
        page.evaluate('window.__decodeReleases.shift()()')
        page.wait_for_function('window.__decodeReleases.length === 0')
        # Let the native decode and canvas callbacks finish before checking.
        page.evaluate('() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))')
        assert not page.locator('#image-zoom').is_visible()
        assert page.evaluate('window.__previewUrls.size') == 0
        page.locator('[data-preview-image="0"]').click()
        page.locator('[data-preview-image="0"]').click()
        assert page.evaluate('window.__decodeReleases.length') == 2
        page.evaluate('window.__decodeReleases.splice(0).forEach(release => release())')
        expect(page.locator('#image-zoom')).to_be_visible()
        assert page.evaluate('window.__previewUrls.size') == 1
        page.locator('#image-zoom-close').click()
        assert page.evaluate('window.__previewUrls.size') == 0

        # Clearing chosen photos keeps the open review usable. Only successful
        # replacement by Excel releases the source photos that are no longer used.
        source_urls = [page.locator('.source-image-button img').get_attribute('src')]
        page.locator('[data-open-sheet="1"]').click()
        source_urls.append(page.locator('.source-image-button img').get_attribute('src'))
        page.evaluate("""() => {
          window.__releasedSourceUrls = [];
          const revoke = URL.revokeObjectURL.bind(URL);
          URL.revokeObjectURL = url => {window.__releasedSourceUrls.push(url); revoke(url);};
        }""")
        page.locator('#clear-images').click()
        assert page.evaluate('window.__releasedSourceUrls.length') == 0
        assert page.locator('.image-thumbnail').count() == 0
        page.locator('#mode-excel').click()
        page.locator('#data-file').set_input_files(str(bad))
        page.locator('#analyze').click()
        expect(page.locator('#status')).to_have_class('status error')
        idle(page)
        assert page.evaluate('window.__releasedSourceUrls.length') == 0
        assert page.locator('.source-image-button img').get_attribute('src') == source_urls[1]
        page.locator('#data-file').set_input_files(str(ROOT / 'samples/data-mau.xlsx'))
        page.locator('#analyze').click()
        page.wait_for_function("document.getElementById('source-filename').textContent.includes('data-mau')")
        idle(page)
        assert set(page.evaluate('window.__releasedSourceUrls')) == set(source_urls)
        assert not errors, errors
        browser.close()
        print('Frontend regressions passed: decimal export, invalid drafts, failed upload preservation, OCR draft/edit preservation, busy controls, cancellation cleanup/retry, zoom races and source-photo cleanup.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--base', default='http://127.0.0.1:8096')
    run(parser.parse_args().base)
