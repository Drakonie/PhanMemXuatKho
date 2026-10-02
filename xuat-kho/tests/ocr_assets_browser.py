"""Exercise both bundled OCR cores with real Vietnamese OCR and no remote assets.

Requires developer tools Playwright and Chromium, not additional app packages.
Starts and stops its own local app by default, or accepts ``--base URL``.
"""
from __future__ import annotations

import argparse
import base64
import os
from pathlib import Path
import socket
import subprocess
import sys
import tempfile
import time
from urllib.parse import urlsplit
from urllib.request import urlopen

from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / 'samples/anh-minh-hoa.png'
CORES = ('tesseract-core-lstm.wasm.js', 'tesseract-core-simd-lstm.wasm.js')
EXPECTED = [
    ('Gạo tẻ', 'kg', 2, 20000, 40000),
    ('Thịt lợn', 'kg', 3, 80000, 240000),
    ('Rau cải', 'kg', 5, 12000, 60000),
]


def check_core(browser, base, core):
    origin = urlsplit(base)
    external, failures, page_errors, requests, responses = [], [], [], [], []
    context = browser.new_context(service_workers='block')

    def local_only(route):
        url = urlsplit(route.request.url)
        if url.scheme in ('http', 'https') and (url.scheme, url.netloc) != (origin.scheme, origin.netloc):
            external.append(route.request.url)
            route.abort()
        else:
            route.continue_()

    context.route('**/*', local_only)
    context.on('request', lambda request: requests.append(urlsplit(request.url).path))
    context.on('response', lambda response: responses.append((urlsplit(response.url).path, response.status)))
    context.on('requestfailed', lambda request: failures.append((request.url, request.failure)))
    page = context.new_page()
    page.on('pageerror', lambda error: page_errors.append(str(error)))
    try:
        page.goto(base)
        page.add_script_tag(url=base.rstrip('/') + '/ocr/tesseract.min.js')
        document = page.evaluate('''async ({core, image}) => {
            const original = document.createElement('canvas');
            const photo = new Image();
            photo.src = 'data:image/png;base64,' + image;
            await photo.decode();
            original.width = photo.naturalWidth;
            original.height = photo.naturalHeight;
            original.getContext('2d').drawImage(photo, 0, 0);
            const prepared = ImageRegions.prepare(original, {mode:'auto'});
            let worker;
            try {
                worker = await Tesseract.createWorker('vie+eng', 1, {
                    workerPath:'/ocr/worker.min.js',
                    corePath:'/ocr/core/' + core,
                    langPath:'/ocr/lang',
                    workerBlobURL:false,
                    gzip:true,
                    cacheMethod:'none'
                });
                await worker.setParameters({
                    tessedit_pageseg_mode:'6', preserve_interword_spaces:'1'
                });
                const {data} = await worker.recognize(prepared.canvas, {}, {text:true, tsv:true});
                return {
                    filename:'anh-minh-hoa.png', text:data.text, tsv:data.tsv,
                    width:prepared.canvas.width, height:prepared.canvas.height,
                    language:'vie+eng', ocr_confidence:data.confidence
                };
            } finally {
                if (worker) await worker.terminate();
                original.width = original.height = 1;
                prepared.canvas.width = prepared.canvas.height = 1;
            }
        }''', {'core': core, 'image': base64.b64encode(FIXTURE.read_bytes()).decode('ascii')})
        data = page.evaluate('''async document => {
            const form = new FormData();
            form.append('ocr_documents', JSON.stringify([document]));
            const response = await fetch('/api/image-analyze', {method:'POST', body:form});
            if (!response.ok) throw new Error(await response.text());
            return await response.json();
        }''', document)
        assert len(data['sheets']) == 1, data['sheets']
        sheet = data['sheets'][0]
        actual = [(item['name'], item['unit'], item['quantity'], item['price'], item['amount']) for item in sheet['items']]
        assert actual == EXPECTED, (core, actual, document['text'])
        assert sheet['customer'] == 'Bếp ăn Minh An' and sheet['date'] == '2026-10-01', sheet
        assert sheet['total'] == 340000 and not sheet['errors'], sheet
        assert sheet['requires_review'] is True and sheet['selected'] is False
        assets = [path for path in requests if path.startswith('/ocr/')]
        assert f'/ocr/core/{core}' in assets, (core, assets)
        assert set(path for path in assets if path.startswith('/ocr/core/')) == {f'/ocr/core/{core}'}, assets
        assert '/ocr/lang/vie.traineddata.gz' in assets and '/ocr/lang/eng.traineddata.gz' in assets, assets
        assert not any(path.endswith('.wasm') for path in assets), assets
        assert all(status == 200 for path, status in responses if path.startswith('/ocr/')), responses
        assert not external and not failures and not page_errors, (external, failures, page_errors)
        print(f'Passed {core}: fresh local vie+eng models, 3 exact goods, total 340000; no standalone WASM or remote requests.')
    finally:
        context.close()


def run(base=None):
    process = None
    with tempfile.TemporaryFile(mode='w+b') as server_log:
        try:
            if base is None:
                with socket.socket() as listener:
                    listener.bind(('127.0.0.1', 0))
                    port = listener.getsockname()[1]
                base = f'http://127.0.0.1:{port}'
                process = subprocess.Popen([sys.executable, str(ROOT / 'app.py'), '--port', str(port)], cwd=ROOT, stdout=server_log, stderr=subprocess.STDOUT)
                deadline = time.monotonic() + 20
                while True:
                    try:
                        with urlopen(base, timeout=1) as response:
                            if response.status == 200:
                                break
                    except OSError:
                        if process.poll() is not None or time.monotonic() >= deadline:
                            server_log.seek(0)
                            raise RuntimeError(server_log.read().decode('utf-8', errors='replace'))
                        time.sleep(0.1)
            with sync_playwright() as playwright:
                launch = {'headless': True, 'args': ['--no-sandbox']}
                executable = os.environ.get('CHROMIUM_EXECUTABLE', '/usr/bin/chromium')
                if Path(executable).is_file():
                    launch['executable_path'] = executable
                browser = playwright.chromium.launch(**launch)
                try:
                    for core in CORES:
                        check_core(browser, base, core)
                finally:
                    browser.close()
        finally:
            if process is not None:
                process.terminate()
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--base', help='Existing app URL; starts its own local server when omitted.')
    run(parser.parse_args().base)
