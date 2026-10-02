'use strict';
window.ImageImport = (() => {
  let inputMode = 'excel', chosen = [], sourceImages = new Map(), documents = [], worker = null, cancelled = false, scriptPromise = null;
  let serial = 0, jobId = 0, zoomId = 0, zoomUrl = null, cancelPending = null;
  const MAX_IMAGES = 10;
  const MAX_IMAGE_BYTES = 20 * 1024 * 1024;
  function mode() {return inputMode;}
  function switchMode(value) {
    if (busy) return;
    inputMode = value;
    $('mode-excel').classList.toggle('active',value === 'excel');$('mode-images').classList.toggle('active',value === 'images');
    $('mode-excel').setAttribute('aria-pressed',String(value === 'excel'));$('mode-images').setAttribute('aria-pressed',String(value === 'images'));
    $('data-drop').hidden = value !== 'excel';$('data-file').disabled = value !== 'excel';$('data-file').required = value === 'excel';
    $('image-input').hidden = value !== 'images';
    $('analyze').innerHTML = value === 'images' ? 'Đọc ảnh <span aria-hidden="true">→</span>' : `Đọc${sheets.length ? ' lại' : ''} file Excel <span aria-hidden="true">→</span>`;
    updateUploadHelp(value);
  }
  function canRevoke(url) {return ![...sourceImages.values()].some(image => image.url === url);}
  function clearResults() {
    sourceImages.forEach(image => {if (!chosen.some(candidate => candidate.url === image.url)) URL.revokeObjectURL(image.url);});
    sourceImages = new Map();documents = [];
  }
  function thumbnails() {
    $('image-thumbnails').innerHTML = chosen.map((image,index) => `<div class="image-thumbnail"><button class="thumbnail-preview" type="button" data-preview-image="${index}" aria-label="Phóng to ảnh ${esc(image.file.name)}"><img src="${image.url}" style="transform:rotate(${image.rotation}deg)" alt="${esc(image.file.name)}"></button><div class="thumbnail-caption"><strong title="${esc(image.file.name)}">${esc(image.file.name)}</strong><span>${number(image.file.size / 1024)} KB${image.rotation ? ` · Xoay ${image.rotation}°` : ''} · ${image.region?.mode === 'manual' ? 'Vùng đã chọn' : image.region?.mode === 'full' ? 'Toàn bộ ảnh' : 'Tự bỏ lề trắng'}</span></div><div class="thumbnail-actions"><button type="button" data-crop-image="${index}">Chọn vùng đọc</button><button type="button" data-full-image="${index}">Toàn bộ ảnh</button><button type="button" data-rotate-image="${index}" title="Xoay ảnh 90°" aria-label="Xoay ảnh ${esc(image.file.name)}">↻ Xoay</button><button type="button" data-remove-image="${index}" title="Bỏ ảnh" aria-label="Bỏ ảnh ${esc(image.file.name)}">×</button></div></div>`).join('');
    $('clear-images').hidden = !chosen.length;
  }
  function addFiles(files) {
    if (busy) return;
    const candidates = [...files];
    if (chosen.length + candidates.length > MAX_IMAGES) {notify('Mỗi lần nhận diện tối đa 10 ảnh. Bỏ bớt ảnh trong danh sách hoặc chọn ít ảnh hơn.',true);return;}
    const invalid = candidates.find(file => !['image/jpeg','image/png','image/webp'].includes(file.type) && !/\.(jpe?g|png|webp)$/i.test(file.name));
    if (invalid) {notify(`Ảnh “${invalid.name}” chưa được hỗ trợ. Chọn JPG, PNG hoặc WebP.`,true);return;}
    if (candidates.some(file => !file.size || file.size > MAX_IMAGE_BYTES)) {notify('Mỗi ảnh cần có dung lượng từ 1 byte đến 20 MB. Hãy giảm kích thước ảnh quá lớn.',true);return;}
    candidates.forEach(file => {if (!chosen.some(image => image.file.name === file.name && image.file.size === file.size && image.file.lastModified === file.lastModified)) chosen.push({file,url:URL.createObjectURL(file),rotation:0,id:++serial});});
    thumbnails();
  }
  async function zoom(url,name,rotation=0) {
    const request = ++zoomId;
    if (zoomUrl) {URL.revokeObjectURL(zoomUrl);zoomUrl = null;}
    if (rotation) {
      let canvas;
      try {
        canvas = await imageCanvas({url,file:{name},rotation});
        if (request !== zoomId) return;
        const blob = await new Promise(resolve => canvas.toBlob(resolve,'image/jpeg',.92));
        if (request !== zoomId) return;
        if (blob) {zoomUrl = URL.createObjectURL(blob);url = zoomUrl;}
      } catch (_) {} finally {if (canvas) {canvas.width=1;canvas.height=1;}}
    }
    if (request !== zoomId) return;
    const dialog = $('image-zoom');$('image-zoom-content').src = url;$('image-zoom-content').style.transform = '';$('image-zoom-name').textContent = name;
    if (typeof dialog.showModal === 'function') dialog.showModal();else {dialog.setAttribute('open','');dialog.classList.add('fallback-open');}
  }
  function releaseZoom() {zoomId++;$('image-zoom-content').removeAttribute('src');if (zoomUrl) {URL.revokeObjectURL(zoomUrl);zoomUrl = null;}}
  function closeZoom() {releaseZoom();const dialog = $('image-zoom');if (typeof dialog.close === 'function') dialog.close();else dialog.removeAttribute('open');dialog.classList.remove('fallback-open');}
  function progress(value,title,description='') {
    const rounded = Math.min(100,Math.max(0,Math.round(value)));
    $('ocr-progress').hidden = false;$('ocr-progress-title').textContent = title;$('ocr-progress-percent').textContent = `${rounded}%`;$('ocr-progress-bar').value = rounded;$('ocr-progress-description').textContent = description;
  }
  function loadTesseract() {
    if (window.Tesseract) return Promise.resolve(window.Tesseract);
    if (!scriptPromise) scriptPromise = new Promise((resolve,reject) => {
      const script = document.createElement('script');script.src = '/ocr/tesseract.min.js';script.async = true;
      script.onload = () => {if (window.Tesseract) resolve(window.Tesseract);else {script.remove();scriptPromise = null;reject(new Error('Chưa tải được bộ đọc ảnh. Hãy tải lại trang và thử lại.'));}};
      script.onerror = () => {script.remove();scriptPromise = null;reject(new Error('Không tải được bộ đọc ảnh. Kiểm tra kết nối và thử lại.'));};document.head.appendChild(script);
    });
    return scriptPromise;
  }
  async function imageCanvas(image) {
    const element = new Image();element.src = image.url;await element.decode();
    const width = element.naturalWidth,height = element.naturalHeight;
    if (!width || !height || width * height > 120000000) throw new Error(`Ảnh “${image.file.name}” không đọc được hoặc kích thước quá lớn.`);
    const scale = Math.min(1,3400 / Math.max(width,height));
    const rotated = image.rotation % 180 !== 0;
    const canvas = document.createElement('canvas');canvas.width = Math.max(1,Math.round((rotated ? height : width)*scale));canvas.height = Math.max(1,Math.round((rotated ? width : height)*scale));
    const context = canvas.getContext('2d');context.fillStyle = '#fff';context.fillRect(0,0,canvas.width,canvas.height);context.translate(canvas.width/2,canvas.height/2);context.rotate(image.rotation*Math.PI/180);context.drawImage(element,-width*scale/2,-height*scale/2,width*scale,height*scale);
    return canvas;
  }
  function documentNames(images) {
    const used = new Set();
    return images.map(image => {let name = image.file.name;let count = 2;while (used.has(name)) name = `${count++}-${image.file.name}`;used.add(name);return name;});
  }
  function decodeSource(encoded) {
    if (!encoded || typeof encoded !== 'string') throw new Error('Máy chủ chưa trả về bảng Excel từ ảnh. Hãy nhận diện lại.');
    const binary = atob(encoded);const bytes = new Uint8Array(binary.length);for (let index=0;index<binary.length;index++) bytes[index] = binary.charCodeAt(index);
    return new File([bytes],'Du-lieu-tu-anh.xlsx',{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }
  function imageFor(sheet) {
    const info = sheet.image_info || {};
    return sourceImages.get(info.filename || info.image_name || info.source_filename || sheet.source_filename) || [...sourceImages.values()][info.index ?? 0] || null;
  }
  function docFor(sheet) {
    const info = sheet.image_info || {};
    return documents.find(doc => doc.filename === (info.filename || info.image_name || info.source_filename || sheet.source_filename)) || documents[info.index ?? activeIndex] || null;
  }
  function reference(sheet,compact=false) {
    const image = imageFor(sheet);
    if (!image) return '<div class="empty-state">Ảnh nguồn chưa có trên trình duyệt này.</div>';
    const info = sheet.image_info || {}, score = info.ocr_confidence ?? info.confidence;
    return `<div class="image-reference ${compact ? 'compact' : ''}"><div class="image-reference-heading"><strong>Ảnh gốc</strong>${typeof score === 'number' ? `<span title="Điểm đọc ký tự do bộ OCR trả về, không đảm bảo số tiền chính xác.">OCR ${Math.round(score)}/100</span>` : ''}</div><button type="button" class="source-image-button" data-zoom-source aria-label="Phóng to ảnh gốc ${esc(image.filename || image.file.name)}"><img src="${image.url}" style="transform:rotate(${image.rotation}deg)" alt="Ảnh nguồn ${esc(image.filename || image.file.name)}"><span>Phóng to để đối chiếu</span></button><p>${esc(image.filename || image.file.name)}${image.preparation?.length ? `<br>${image.preparation.map(esc).join('<br>')}` : ''}</p></div>`;
  }
  function itemsLayout(table,sheet) {return `<div class="image-review-layout"><aside class="image-review-reference">${reference(sheet,true)}</aside><div class="image-items-pane">${table}</div></div>`;}
  function structuredText(sheet,doc) {
    const info = sheet.image_info || {};
    if (info.corrected_text) return info.corrected_text;
    if (info.edited_text) return doc?.text || info.ocr_text || '';
    const rows = info.source_rows || [];
    if (!rows.some(row => row.cells)) return doc?.text || info.ocr_text || '';
    const labels = {customer:'Khách hàng',date:'Ngày xuất',address:'Địa chỉ',voucher:'Số phiếu',recipient:'Người nhận hàng'};
    const fields = ['name','code','unit','quantity','price','amount','note'];
    const quote = value => {const text = String(value ?? '');return /[|\t"\r\n]/.test(text) ? `"${text.replaceAll('"','""')}"` : text;};
    const lines = Object.entries(info.metadata || {}).filter(([key]) => labels[key]).map(([key,value]) => `${labels[key]}: ${value}`);
    const header = 'Tên hàng | Mã hàng | ĐVT | Số lượng | Đơn giá | Thành tiền | Ghi chú';
    lines.push(header);
    rows.forEach(row => {
      if (row.type === 'metadata') {Object.entries(row.metadata || {}).forEach(([key,value]) => {if (labels[key]) lines.push(`${labels[key]}: ${value}`);});lines.push(header);}
      else if (row.cells) lines.push(fields.map(field => quote(row.cells[field])).join(' | '));
    });
    return lines.join('\n');
  }
  function sourceMarkup(sheet) {
    const doc = docFor(sheet), text = sheet._ocrDraft ?? structuredText(sheet,doc);
    return `<div class="image-source-layout">${reference(sheet)}<div class="ocr-text-editor"><h4>Văn bản nhận diện</h4><p>Sửa chữ hoặc số đọc sai rồi đọc lại. Bảng chữ đã tách theo cột để dễ sửa; mỗi mặt hàng một dòng. Dùng dấu | hoặc ký tự Tab và giữ dòng tên cột.</p><label for="ocr-text-draft">Bảng chữ để sửa tên hàng và các số<textarea id="ocr-text-draft" rows="15" spellcheck="false">${esc(text)}</textarea></label><div class="ocr-editor-actions"><span>Đọc lại sẽ thay dữ liệu mặt hàng và yêu cầu xác nhận lại ảnh này.</span><button type="button" class="button-secondary" data-reparse-ocr>Đọc lại văn bản đã sửa</button></div><details class="ocr-raw-details"><summary>Văn bản OCR gốc và vị trí chữ</summary><pre>${esc(doc?.text || '')}</pre><pre>${esc(doc?.tsv || '')}</pre></details></div></div>`;
  }
  function acceptResult(data,sourceFile,template,newDocuments,newSources,target=null) {
    if (!Array.isArray(data.sheets) || !data.sheets.length) throw new Error('Chưa tìm được bảng hàng trong ảnh. Thử ảnh rõ hơn hoặc chỉnh văn bản nhận diện.');
    const previous = sheets;
    const nextSheets = data.sheets.map((source,index) => {
      const info = source.image_info || {};
      source = {...source,source_type:'image',image_info:{...info,filename:info.filename || newDocuments[index]?.filename,index:info.index ?? index}};
      const fresh = prepareSheet(source), old = previous.find(sheet => sheet.sheet === source.sheet);
      if (target && old) {
        metadata.forEach(field => {if (source.sheet !== target || old._metadataTouched.has(field)) fresh[field] = old[field];});fresh._metadataTouched = old._metadataTouched;
        if (source.sheet !== target && old._ocrDraft !== undefined) fresh._ocrDraft = old._ocrDraft;
        if (source.sheet !== target && fresh.items.map(item => item.row).join(',') === old.items.map(item => item.row).join(',')) {fresh.items = old.items;fresh._originalItems = old._originalItems;fresh._excluded = old._excluded;fresh.selected = old.selected;fresh._ocrConfirmed = old._ocrConfirmed;fresh._selectionTouched = old._selectionTouched;}
      }
      return fresh;
    });
    configureRuntime(data);token = data.token || '';sheets = nextSheets;filename = data.filename || `${newDocuments.length} ảnh bảng hàng`;
    analyzedFiles = {data:sourceFile,template};overrides = data.overrides || {};documents = newDocuments;const oldSources = sourceImages;sourceImages = newSources;
    oldSources.forEach(image => {if (![...newSources.values(),...chosen].some(next => next.url === image.url)) URL.revokeObjectURL(image.url);});
    activeIndex = target ? Math.max(0,sheets.findIndex(sheet => sheet.sheet === target)) : 0;activeTab = 'items';page = 1;itemQuery = '';$('sheet-search').value = '';renderAll();
  }
  async function sendDocuments(newDocuments,template,newSources,target=null) {
    const form = new FormData();form.append('ocr_documents',JSON.stringify(newDocuments));if (template) form.append('template',template);if (token && transport !== 'stateless') form.append('replace_token',token);
    if (target) form.append('image_overrides',JSON.stringify(Object.fromEntries(Object.entries(overrides).filter(([name]) => name !== target))));
    const data = await (await api('/api/image-analyze',form)).json();
    if (((data.transport || transport) !== 'stateless' && !data.token)) throw new Error('Máy chủ chưa tạo được phiên kiểm tra ảnh.');
    const sourceFile = decodeSource(data.source_base64);
    acceptResult(data,sourceFile,template,newDocuments,newSources,target);
  }
  async function analyze() {
    if (busy) return;
    if (!chosen.length) {notify('Chọn ít nhất một ảnh bảng hàng hoặc chụp ảnh mới.',true);return;}
    const images = chosen.map(image => ({...image})), template = $('template-file').files[0] || null;
    if (template && !/\.xlsx$/i.test(template.name)) {notify('Template cần là file .xlsx.',true);return;}
    const job = ++jobId;const isCancelled = () => cancelled || job !== jobId;let ownWorker = null;
    let cancelJob;
    const cancellation = new Promise((resolve,reject) => {cancelJob = () => reject(new Error('Đã hủy đọc ảnh.'));});
    const waitForJob = promise => Promise.race([promise,cancellation]);
    cancelPending = cancelJob;
    cancelled = false;setBusy('image');$('ocr-cancel').disabled = false;progress(0,'Đang chuẩn bị bộ đọc ảnh…','Lần đầu cần tải bộ đọc tiếng Việt và tiếng Anh. Những lần sau sẽ nhanh hơn.');notify('Đang đọc chữ và các cột số từ ảnh trên trình duyệt…',false,true);
    let currentImage = 0;
    try {
      await waitForJob(runtimeReady);const tesseract = await waitForJob(loadTesseract());if (isCancelled()) return;
      const creatingWorker = tesseract.createWorker('vie+eng',1,{workerPath:'/ocr/worker.min.js',corePath:'/ocr/core',langPath:'/ocr/lang',workerBlobURL:false,gzip:true,logger:message => {
        if (isCancelled()) return;
        if (message.status === 'recognizing text') progress(10+85*(currentImage+(message.progress || 0))/images.length,`Đang đọc ảnh ${currentImage+1}/${images.length}`,images[currentImage]?.file.name || '');
        else if (message.status?.includes('loading') || message.status?.includes('initializing')) progress(5,'Đang tải và khởi tạo bộ đọc ảnh…','Bộ đọc tiếng Việt + tiếng Anh; không cần API key.');
      }}).then(async created => {if (isCancelled()) {try {await created.terminate();} catch (_) {}}return created;});
      ownWorker = await waitForJob(creatingWorker);
      if (isCancelled()) return;
      worker = ownWorker;
      await waitForJob(ownWorker.setParameters({tessedit_pageseg_mode:$('ocr-layout').value || '6',preserve_interword_spaces:'1'}));
      const names = documentNames(images), newDocuments = [],newSources = new Map();
      for (currentImage=0;currentImage<images.length;currentImage++) {
        if (isCancelled()) return;
        const image = images[currentImage];progress(10+85*currentImage/images.length,`Đang đọc ảnh ${currentImage+1}/${images.length}`,image.file.name);
        const original = await waitForJob(imageCanvas(image).then(created => {if (isCancelled()) {created.width=1;created.height=1;}return created;}));
        const prepared = window.ImageRegions ? ImageRegions.prepare(original,image.region || {mode:'auto'}) : {canvas:original};
        const improved = $('ocr-enhance')?.checked && ImageRegions.enhance ? ImageRegions.enhance(prepared.canvas) : {canvas:prepared.canvas,changes:[]};
        const canvas = improved.canvas;
        try {
          if (isCancelled()) return;
          await waitForJob(ownWorker.setParameters({tessedit_pageseg_mode:$('ocr-layout').value || '6',tessedit_char_whitelist:'',preserve_interword_spaces:'1'}));
          const result = await waitForJob(ownWorker.recognize(canvas,{}, {text:true,tsv:true,blocks:true}));if (isCancelled()) return;
          const data = result.data;
          let parameterMode = null;
          const tableData = window.ImageRegions ? await ImageRegions.readTable(canvas,async (cell,numeric) => {
            if (parameterMode !== numeric) {await waitForJob(ownWorker.setParameters({tessedit_pageseg_mode:'7',tessedit_char_whitelist:numeric?'0123456789.,-()':''}));parameterMode=numeric;}
            return (await waitForJob(ownWorker.recognize(cell,{}, {text:true,tsv:true}))).data;
          },(row,count)=>progress(10+85*(currentImage+(row+1)/Math.max(1,count))/images.length,`Đọc từng ô · ảnh ${currentImage+1}/${images.length}`,`Dòng ${row+1}/${count} · ${image.file.name}`)) : null;
          newDocuments.push({filename:names[currentImage],text:data.text || '',tsv:data.tsv || '',width:canvas.width,height:canvas.height,rotation:image.rotation,language:'vie+eng',ocr_confidence:Number(data.confidence || 0),...(tableData ? {table_data:tableData} : {})});
          newSources.set(names[currentImage],{...image,filename:names[currentImage],preparation:improved.changes});
        } finally {canvas.width = 1;canvas.height = 1;prepared.canvas.width=prepared.canvas.height=1;original.width = 1;original.height = 1;}
      }
      $('ocr-cancel').disabled = true;progress(97,'Đang dựng bảng hàng để kiểm tra…','Tách tên hàng, số lượng, đơn giá và thành tiền từ văn bản.');
      await sendDocuments(newDocuments,template,newSources);if (isCancelled()) return;
      notify(`Đã đọc ${images.length} ảnh. Đối chiếu ảnh gốc, sửa dữ liệu và xác nhận từng phiếu trước khi xuất.`);progress(100,'Đã nhận diện xong','Các phiếu từ ảnh chưa được chọn xuất cho đến khi bạn xác nhận.');
    } catch (error) {if (!isCancelled()) notify(error.message || 'Chưa đọc được ảnh. Thử ảnh rõ hơn hoặc tải lại trang.',true);} finally {
      if (cancelPending === cancelJob) cancelPending = null;
      if (ownWorker) {try {await ownWorker.terminate();} catch (_) {}if (worker === ownWorker) worker = null;}
      if (job === jobId) {setBusy();switchMode(inputMode);$('ocr-progress').hidden = true;}
    }
  }
  async function cancel() {if (busy !== 'image' || $('ocr-cancel').disabled) return;cancelled = true;jobId++;if (cancelPending) cancelPending();const previousWorker = worker;worker = null;setBusy();switchMode(inputMode);$('ocr-progress').hidden = true;notify('Đã hủy đọc ảnh. Kết quả kiểm tra đang mở vẫn được giữ.');if (previousWorker) {try {await previousWorker.terminate();} catch (_) {}}}
  async function reparse() {
    if (busy) return;
    const sheet = current(), original = docFor(sheet);if (!original) {notify('Chưa tìm được văn bản ảnh gốc để đọc lại.',true);return;}
    const text = $('ocr-text-draft').value.trim();if (!text) {notify('Nhập dòng tiêu đề và ít nhất một dòng hàng để đọc lại.',true);return;}
    const group = sheet.image_info?.customer_group;
    const newDocuments = documents.map(doc => doc.filename !== original.filename ? {...doc} : group ? {...doc,customer_edits:{...(doc.customer_edits || {}),[String(group.index)]:text}} : {...doc,text,tsv:'',edited_text:true});
    setBusy('image-reparse');notify('Đang đọc lại văn bản ảnh đã sửa…',false,true);
    try {await sendDocuments(newDocuments,analyzedFiles?.template || null,sourceImages,sheet.sheet);notify('Đã đọc lại văn bản. Đối chiếu mặt hàng và xác nhận lại phiếu từ ảnh trước khi xuất.');} catch (error) {notify(error.message,true);} finally {setBusy();switchMode(inputMode);}
  }
  let cropRequest = 0;
  async function cropImage(index) {
    if (busy || !chosen[index]) return;
    const image = chosen[index], request = ++cropRequest;
    let dialog = $('image-region-dialog');
    if (!dialog) {
      dialog = document.createElement('dialog');dialog.id = 'image-region-dialog';dialog.className = 'image-region-dialog';
      dialog.innerHTML = '<div class="image-zoom-toolbar"><strong>Chọn vùng có dữ liệu</strong><button type="button" data-region-close class="button-ghost">Đóng ×</button></div><p class="region-help">Kéo một khung trên ảnh, hoặc nhập tỷ lệ bên dưới. Giữ tên hàng, cột số và tiêu đề khách hàng trong vùng đọc.</p><div class="region-preview"><img alt="Ảnh để chọn vùng đọc" draggable="false"><div class="region-box"></div></div><div class="region-fields">'+['x','y','width','height'].map((name,i)=>`<label>${['Trái','Trên','Rộng','Cao'][i]} (%)<input type="number" data-region-field="${name}" min="${i>1 ? 1 : 0}" max="100" step="0.1"></label>`).join('')+'</div><div class="region-actions"><button type="button" data-region-auto class="button-secondary">Tự bỏ lề trắng</button><button type="button" data-region-apply class="button-primary">Dùng vùng đã chọn</button></div>';
      document.body.appendChild(dialog);
    }
    const canvas = await imageCanvas(image);
    try {if (request !== cropRequest || !chosen.includes(image)) return;dialog.querySelector('img').src = canvas.toDataURL('image/jpeg',.95);} finally {canvas.width=1;canvas.height=1;}
    let region = image.region?.mode === 'manual' ? {...image.region} : {x:0,y:0,width:1,height:1};
    const preview = dialog.querySelector('.region-preview'), box = dialog.querySelector('.region-box');
    const inputs = [...dialog.querySelectorAll('[data-region-field]')];
    function draw(fill=true) {
      region = ImageRegions.normalized(region);box.style.left=`${region.x*100}%`;box.style.top=`${region.y*100}%`;box.style.width=`${region.width*100}%`;box.style.height=`${region.height*100}%`;
      if(fill)inputs.forEach(input=>input.value=String(Math.round(region[input.dataset.regionField]*1000)/10));
    }
    const close = () => {cropRequest++;dialog.close();dialog.querySelector('img').removeAttribute('src');};
    dialog.querySelector('[data-region-close]').onclick=close;
    dialog.oncancel=()=>{cropRequest++;dialog.querySelector('img').removeAttribute('src');};
    dialog.querySelector('[data-region-auto]').onclick=()=>{if(chosen.includes(image)){image.region={mode:'auto'};thumbnails();}close();};
    dialog.querySelector('[data-region-apply]').onclick=()=>{if(chosen.includes(image)){image.region={...region,mode:'manual'};thumbnails();notify('Đã chọn vùng đọc. Bấm nhận diện để đọc lại; dữ liệu đang mở được giữ đến khi đọc thành công.');}close();};
    inputs.forEach(input=>input.oninput=()=>{region[input.dataset.regionField]=Number(input.value)/100;draw(false);});
    let anchor=null;
    const point = event => {const rect=preview.getBoundingClientRect();return {x:Math.min(1,Math.max(0,(event.clientX-rect.left)/rect.width)),y:Math.min(1,Math.max(0,(event.clientY-rect.top)/rect.height))};};
    preview.onpointerdown=event=>{anchor=point(event);preview.setPointerCapture(event.pointerId);event.preventDefault();};
    preview.onpointermove=event=>{if(!anchor)return;const end=point(event);region={x:Math.min(anchor.x,end.x),y:Math.min(anchor.y,end.y),width:Math.max(.01,Math.abs(end.x-anchor.x)),height:Math.max(.01,Math.abs(end.y-anchor.y))};draw();};
    preview.onpointerup=()=>{anchor=null;};preview.onpointercancel=()=>{anchor=null;};
    draw();dialog.showModal();
  }
  $('mode-excel').addEventListener('click',() => switchMode('excel'));$('mode-images').addEventListener('click',() => switchMode('images'));
  $('image-files').addEventListener('change',event => {addFiles(event.target.files);event.target.value='';});$('camera-file').addEventListener('change',event => {addFiles(event.target.files);event.target.value='';});
  $('clear-images').addEventListener('click',() => {if (busy) return;chosen.forEach(image => {if (canRevoke(image.url)) URL.revokeObjectURL(image.url);});chosen=[];thumbnails();});
  ['dragenter','dragover'].forEach(type => $('image-drop').addEventListener(type,event => {event.preventDefault();if (!busy) $('image-drop').classList.add('dragging');}));
  ['dragleave','drop'].forEach(type => $('image-drop').addEventListener(type,event => {event.preventDefault();$('image-drop').classList.remove('dragging');}));
  $('image-drop').addEventListener('drop',event => addFiles(event.dataTransfer?.files || []));
  $('image-thumbnails').addEventListener('click',event => {
    if (busy) return;const button = event.target.closest('button');if (!button) return;
    if (button.dataset.previewImage !== undefined) {const image = chosen[Number(button.dataset.previewImage)];zoom(image.url,image.file.name,image.rotation);}
    if (button.dataset.cropImage !== undefined) cropImage(Number(button.dataset.cropImage));
    if (button.dataset.fullImage !== undefined) {chosen[Number(button.dataset.fullImage)].region={mode:'full'};thumbnails();}
    if (button.dataset.rotateImage !== undefined) {const image = chosen[Number(button.dataset.rotateImage)];image.rotation=(image.rotation+90)%360;image.region={mode:'auto'};thumbnails();}
    if (button.dataset.removeImage !== undefined) {const [image] = chosen.splice(Number(button.dataset.removeImage),1);if (canRevoke(image.url)) URL.revokeObjectURL(image.url);thumbnails();}
  });
  $('ocr-cancel').addEventListener('click',cancel);$('image-zoom-close').addEventListener('click',closeZoom);
  $('image-zoom').addEventListener('close',() => {if (!$('image-zoom').open) releaseZoom();});
  $('image-zoom').addEventListener('click',event => {if (event.target === $('image-zoom')) closeZoom();});
  $('sheet-detail').addEventListener('click',event => {const button = event.target.closest('button');if (!button) return;if (button.hasAttribute('data-zoom-source')) {const image = imageFor(current());if (image) zoom(image.url,image.filename || image.file.name,image.rotation);}if (button.hasAttribute('data-reparse-ocr')) reparse();});
  $('sheet-detail').addEventListener('input',event => {if (event.target.id === 'ocr-text-draft' && !busy) current()._ocrDraft=event.target.value;});
  return {mode,analyze,itemsLayout,sourceMarkup,clearResults};
})();
