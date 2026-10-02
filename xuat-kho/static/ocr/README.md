# Bundled browser OCR

Tesseract.js 6.0.1 and tesseract.js-core 6.0.0, distributed under Apache-2.0.
Vietnamese and English fast LSTM traineddata originate from https://github.com/tesseract-ocr/tessdata_fast and use Apache-2.0.

All runtime assets are served from this application. No OCR CDN or third-party processing endpoint is configured. OCR runs in a browser Web Worker; recognized text and word positions are sent to the application API to construct an Excel source.

The adjacent license files cover runtime bundles and language models. asset-manifest.json records sizes and SHA-256 hashes of the shipped assets.

The application uses LSTM-only OCR (OEM 1). The package includes the official
`tesseract-core-simd-lstm.wasm.js` and `tesseract-core-lstm.wasm.js` bundles;
Tesseract selects the SIMD variant when supported and otherwise uses the second
bundle. Each already embeds its complete WASM binary. Standalone `.wasm` copies
and legacy OCR cores are therefore omitted. Both original language models and
the retained core bundles are unchanged; their hashes remain in the manifest.
