'use strict';
const $ = id => document.getElementById(id);
const fields = {name:'Tên hàng', quantity:'Số lượng', price:'Đơn giá', amount:'Thành tiền', unit:'ĐVT', code:'Mã hàng', note:'Ghi chú'};
const metadata = ['customer','recipient','date','address','voucher'];
const formatter = new Intl.NumberFormat('vi-VN', {maximumFractionDigits: 6});
const moneyFormatter = new Intl.NumberFormat('vi-VN', {maximumFractionDigits: 2});
const money = value => moneyFormatter.format(Number.isFinite(Number(value)) ? Number(value) : 0);
const number = value => formatter.format(Number.isFinite(Number(value)) ? Number(value) : 0);
const esc = value => String(value ?? '').replace(/[&<>"']/g, character => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[character]));
const clone = value => JSON.parse(JSON.stringify(value));
let token = '', sheets = [], activeIndex = 0, activeTab = 'items', itemQuery = '', page = 1, busy = '', filename = '';
let transport = /^(localhost|127\.0\.0\.1|\[?::1\]?)$/.test(location.hostname) ? 'session' : 'stateless';
let maxUploadBytes = (transport === 'stateless' ? 4 : 25) * 1024 * 1024;
let analyzedFiles = null, overrides = {};
const PAGE_SIZE = 50;
const EXPORT_TIMEOUT_MS = 90000;
let exportTask = null, generatedDownloadUrl = '';

function updateUploadHelp(mode = window.ImageImport?.mode() || 'excel') {
  $('upload-help').innerHTML = mode === 'images'
    ? '<span aria-hidden="true">✦</span> Chọn ảnh ở trên, rồi bấm “Đọc ảnh”. Tối đa 10 ảnh, mỗi ảnh tối đa 20 MB.'
    : `<span aria-hidden="true">✦</span> Chọn file ở trên, rồi bấm “Đọc file Excel”. Tổng file tối đa ${number(maxUploadBytes / 1024 / 1024)} MB.`;
}
function configureRuntime(data = {}) {
  if (data.transport === 'stateless' || data.transport === 'session') transport = data.transport;
  if (Number.isFinite(data.max_upload_bytes) && data.max_upload_bytes > 0) maxUploadBytes = data.max_upload_bytes;
  else maxUploadBytes = (transport === 'stateless' ? 4 : 25) * 1024 * 1024;
  $('connection-label').textContent = transport === 'stateless' ? 'Ứng dụng trực tuyến' : 'Ứng dụng cục bộ';
  updateUploadHelp();
  $('session-note').textContent = transport === 'stateless' ? 'Tải lại trang sẽ xóa phiên làm việc. Hãy tải file xuất trước khi đóng trang.' : 'Giữ bố cục template · Tự thêm dòng khi nhiều mặt hàng';
}
async function initializeRuntime() {
  configureRuntime();
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(),3000);
  try {
    const response = await fetch('/api/health',{signal:controller.signal});
    if (response.ok) {
      const data = await response.json();
      configureRuntime({...data,transport:data.transport || 'session'});
    }
  } catch (_) {} finally {clearTimeout(timeout);}
}
function requestFormData(path, body) {
  if (transport !== 'stateless' || body instanceof FormData || !['/api/remap','/api/export'].includes(path)) return body;
  if (!analyzedFiles?.data) throw new Error('Chọn lại file Excel và bấm nhận diện trước khi tiếp tục.');
  const form = new FormData();
  form.append('data',analyzedFiles.data);
  if (analyzedFiles.template) form.append('template',analyzedFiles.template);
  form.append('payload',JSON.stringify({...body,overrides:body.overrides || overrides}));
  return form;
}
function validateRequestSize(body) {
  if (transport !== 'stateless' || !(body instanceof FormData)) return;
  // Leave room for the multipart boundaries, field headers, and file names.
  let bytes = 1024;
  for (const [key,value] of body.entries()) bytes += 1024 + new Blob([key]).size + (typeof value === 'string' ? new Blob([value]).size : value.size);
  if (bytes > maxUploadBytes) throw new Error(`Tổng file và nội dung chỉnh sửa vượt giới hạn ${number(maxUploadBytes / 1024 / 1024)} MB. Hãy chia workbook hoặc giảm dung lượng file.`);
}

function notify(message, error = false, waiting = false) {
  const element = $('status');
  element.hidden = !message;
  element.textContent = message;
  element.className = `status${error ? ' error' : ''}${waiting ? ' busy' : ''}`;
  element.setAttribute('role', error ? 'alert' : 'status');
}
function setBusy(kind = '') {
  busy = kind;
  document.querySelectorAll('#upload button:not(#ocr-cancel), #upload input, #upload select, #workspace button:not(#export-cancel), #workspace input, #workspace select, #workspace textarea').forEach(element => {
    if (kind) {
      if (!element.hasAttribute('data-was-disabled')) element.dataset.wasDisabled = element.disabled ? '1' : '0';
      element.disabled = true;
    } else if (element.hasAttribute('data-was-disabled')) {
      element.disabled = element.dataset.wasDisabled === '1';
      delete element.dataset.wasDisabled;
    }
  });
  $('workspace').setAttribute('aria-busy', kind ? 'true' : 'false');
  $('upload').setAttribute('aria-busy', kind ? 'true' : 'false');
  $('analyze').innerHTML = ['analyze','image'].includes(kind) ? 'Đang đọc dữ liệu… <span class="spinner-button" aria-hidden="true"></span>' : window.ImageImport?.mode() === 'images' ? 'Đọc ảnh <span aria-hidden="true">→</span>' : `Đọc${sheets.length ? ' lại' : ''} file Excel <span aria-hidden="true">→</span>`;
  $('export').innerHTML = kind === 'export' ? 'Đang tạo file… <span class="spinner-button" aria-hidden="true"></span>' : '3. Tạo và tải file <span aria-hidden="true">↓</span>';
  if (!kind) {if (sheets.length) renderNavigation();syncTotals();}
}
async function api(path, body, options = {}) {
  body = requestFormData(path,body);
  validateRequestSize(body);
  let response;
  try {
    response = await fetch(path, {method:'POST', body: body instanceof FormData ? body : JSON.stringify(body), headers: body instanceof FormData ? {} : {'Content-Type':'application/json'}, signal:options.signal});
  } catch (error) {
    if (options.signal?.aborted) throw error;
    throw new Error(transport === 'stateless' ? 'Không kết nối được máy chủ. Kiểm tra kết nối mạng và thử lại; kết quả đang mở vẫn được giữ trên trang.' : 'Không kết nối được ứng dụng. Hãy giữ cửa sổ Chay-Windows.bat mở và thử lại.');
  }
  if (!response.ok) {
    let message = response.status === 413 ? `Dữ liệu quá lớn để xử lý. Hãy chia workbook, giảm dung lượng file hoặc chọn ít phiếu hơn khi xuất.` : `Ứng dụng trả về lỗi ${response.status}.`;
    try {const data = await response.json(); message = data.error || message;} catch (_) {}
    throw new Error(message);
  }
  return response;
}
function isImageSheet(sheet) {return sheet?.source_type === 'image' || sheet?.sourceType === 'image' || !!sheet?.image_info;}
function prepareSheet(source) {
  const sheet = {...source};
  sheet.items = (source.items || []).map(item => ({...item, _invalid:{}}));
  sheet._originalItems = clone(sheet.items);
  sheet._excluded = new Set();
  sheet._selectionTouched = false;
  sheet._metadataTouched = new Set();
  sheet._ocrConfirmed = false;
  if (isImageSheet(sheet)) {sheet.selected = false;sheet.requires_review = true;}
  sheet.errors = source.errors || [];
  sheet.warnings = source.warnings || [];
  sheet.columns = source.columns || [];
  sheet.mapping = source.mapping || {};
  metadata.forEach(field => {sheet[field] = sheet[field] ?? '';});
  return sheet;
}
function current() {return sheets[activeIndex];}
function isHidden(sheet) {return !!sheet.hidden || ['hidden','veryHidden'].includes(sheet.sheet_state || sheet.visibility);}
function hasIssues(sheet) {return !!sheet.errors.length || !!sheet.warnings.length || !sheet.customer.trim() || !sheet.date || (Number.isFinite(sheet.confidence) && sheet.confidence < 75);}
function includedItems(sheet) {return sheet.items.filter(item => !sheet._excluded.has(item.row) && !($('exclude-zero').checked && Number(item.quantity) === 0));}
function sheetTotal(sheet) {return includedItems(sheet).reduce((total, item) => total + (Number.isFinite(Number(item.amount)) ? Number(item.amount) : 0), 0);}
function changedRows(sheet) {
  return sheet.items.filter((item, index) => {
    const original = sheet._originalItems[index];
    return sheet._excluded.has(item.row) || Object.values(item._invalid || {}).some(Boolean) || ['name','unit','quantity','price','amount','note','code'].some(key => item[key] !== original?.[key]);
  }).length;
}
function confidence(sheet) {
  if (!Number.isFinite(Number(sheet.confidence))) return '<span class="confidence">Chưa có điểm nhận diện</span>';
  const value = Math.round(Number(sheet.confidence));
  return `<span class="confidence ${value < 55 ? 'low' : value < 80 ? 'medium' : ''}" title="Điểm dựa trên cấu trúc bảng và dữ liệu; hãy kiểm tra các lưu ý trước khi xuất.">Điểm nhận diện ${value}/100</span>`;
}
function labelForSheet(sheet) {
  if (sheet.summary) return '<span class="mini-tag">Tổng hợp</span>';
  if (isHidden(sheet)) return '<span class="mini-tag">Sheet ẩn</span>';
  if (sheet.errors.length) return '<span class="mini-tag error">Cần kiểm tra</span>';
  if (isImageSheet(sheet)) return sheet._ocrConfirmed ? '<span class="mini-tag teal">Đã kiểm tra</span>' : '<span class="mini-tag warn">Cần xác nhận</span>';
  if (sheet.requires_review || sheet.detection_method === 'inferred') return '<span class="mini-tag warn">Cần xác nhận</span>';
  if (hasIssues(sheet)) return '<span class="mini-tag warn">Có lưu ý</span>';
  return '<span class="mini-tag teal">Chi tiết</span>';
}
function renderNavigation() {
  const query = $('sheet-search').value.trim().toLocaleLowerCase('vi');
  let visible = 0;
  $('sheets').innerHTML = sheets.map((sheet, index) => {
    if (query && !`${sheet.sheet} ${sheet.customer}`.toLocaleLowerCase('vi').includes(query)) return '';
    visible++;
    return `<div class="sheet-card ${index === activeIndex ? 'active' : ''}" id="sheet-${index}"><label class="sheet-select-target"><input class="selected" type="checkbox" data-sheet-select="${index}" ${sheet.selected ? 'checked' : ''} aria-label="${isImageSheet(sheet) ? 'Xác nhận đã kiểm tra dữ liệu OCR và chọn xuất ảnh' : 'Chọn xuất sheet'} ${esc(sheet.sheet)}" ${busy ? 'disabled' : ''}></label><button type="button" data-open-sheet="${index}" ${index === activeIndex ? 'aria-current="true"' : ''} ${busy ? 'disabled' : ''}><strong>${esc(sheet.sheet)}${changedRows(sheet) ? '<span class="edited-indicator" title="Có chỉnh sửa">●</span>' : ''}</strong><span class="customer-name">${esc(sheet.customer || 'Chưa có khách hàng')}</span><span class="sheet-meta">${labelForSheet(sheet)}<span>${sheet.items.length} dòng · ${money(sheetTotal(sheet))} ₫</span></span></button></div>`;
  }).join('') || '<div class="empty-state">Không có sheet phù hợp.</div>';
  $('sidebar-count').textContent = `${visible}/${sheets.length}`;
  renderCustomers();
}
function renderCustomers() {
  const container = $('customer-selection');if (!container) return;
  const customers = sheets.map((sheet,index) => ({sheet,index})).filter(({sheet}) => isImageSheet(sheet));
  container.hidden = !customers.length;
  container.innerHTML = customers.length ? `<div class="customer-heading"><h3>Chọn xuất theo khách hàng</h3><p>Tick từng phiếu sau khi đã đối chiếu dữ liệu OCR với ảnh. Sửa tên khách hàng trong chi tiết phiếu nếu cần.</p></div><div class="customer-list">${customers.map(({sheet,index}) => `<label class="customer-option ${sheet.selected ? 'selected' : ''}"><input type="checkbox" data-customer-select="${index}" ${sheet.selected ? 'checked' : ''} ${busy ? 'disabled' : ''} aria-label="Đã kiểm tra và chọn xuất phiếu khách hàng ${esc(sheet.customer || sheet.sheet)}"><span class="customer-copy"><strong>${esc(sheet.customer || 'Chưa đọc được khách hàng')}</strong><small>${esc(sheet.image_info?.filename || sheet.sheet)} · ${(sheet.image_info?.source_rows || []).filter(row => row.type === 'goods').length || sheet.items.length} dòng phát hiện${sheet.errors.length ? ' · Cần sửa dữ liệu' : ''}</small></span><span class="customer-amount">${sheet.errors.length ? 'Tạm tính ' : ''}${money(sheetTotal(sheet))} ₫</span></label>`).join('')}</div>` : '';
}
function syncTotals() {
  if (!sheets.length) return;
  const selected = sheets.filter(sheet => sheet.selected);
  let count = 0, total = 0;
  selected.forEach(sheet => {count += includedItems(sheet).length; total += sheetTotal(sheet);});
  $('metric-sheets').innerHTML = `${selected.length}<small>phiếu</small>`;
  $('metric-items').innerHTML = `${number(count)}<small>dòng</small>`;
  $('metric-total').innerHTML = `${money(total)}<small>₫</small>`;
  $('metric-review').innerHTML = `${sheets.filter(hasIssues).length}<small>sheet</small>`;
  const imageCount = new Set(sheets.filter(isImageSheet).map(sheet => sheet.image_info?.filename || sheet.image_info?.image_name || sheet.sheet)).size;
  $('counts').textContent = sheets.every(isImageSheet) ? `${imageCount} ảnh · ${sheets.length} phiếu khách hàng` : `${sheets.length} sheet đã quét`;
  $('total').textContent = `${selected.length} phiếu · ${number(count)} dòng hàng · ${money(total)} đồng`;
  $('export').disabled = !!busy || selected.length === 0;
  $('apply-bulk').disabled = !!busy || selected.length === 0;
  if (current()) {
    const totalElement = document.querySelector('.sheet-total');
    if (totalElement) totalElement.textContent = `${money(sheetTotal(current()))} ₫`;
    const countElement = document.querySelector('.table-count');
    if (countElement) countElement.textContent = `${includedItems(current()).length} dòng sẽ xuất${changedRows(current()) ? ` · ${changedRows(current())} dòng đã sửa / bỏ` : ''}`;
    const undo = document.querySelector('[data-action="undo"]');
    if (undo) undo.disabled = !!busy || changedRows(current()) === 0;
  }
}
function renderAlerts(sheet) {
  const imageReview = isImageSheet(sheet) ? `<div class="alert image-confirmation"><strong>Kiểm tra dữ liệu nhận diện từ ảnh</strong><p>Đối chiếu ảnh gốc, tên hàng, số lượng và đơn giá. Chữ mờ hoặc chữ viết tay có thể đọc sai.</p><label class="checkbox-label"><input type="checkbox" data-image-confirm ${sheet.selected ? 'checked' : ''}> Tôi đã kiểm tra dữ liệu OCR và chọn xuất phiếu này</label></div>` : '';
  const errors = sheet.errors.length ? `<details class="alert error" open><summary>${sheet.errors.length} vấn đề từ file nguồn</summary><ul>${sheet.errors.map(text => `<li>${esc(text)}</li>`).join('')}</ul><p>Bạn có thể sửa dòng đã đọc ngay bên dưới. Dòng chưa đọc được cần chọn lại cột hoặc sửa file nguồn. Hệ thống kiểm tra lại khi xuất.</p></details>` : '';
  const warnings = sheet.warnings.length ? `<details class="alert"><summary>${sheet.warnings.length} lưu ý cần kiểm tra</summary><ul>${sheet.warnings.map(text => `<li>${esc(text)}</li>`).join('')}</ul></details>` : '';
  return imageReview + errors + warnings;
}
function renderDetail() {
  const sheet = current();
  if (!sheet) {$('sheet-detail').innerHTML = '<div class="empty-state">Chọn một sheet để xem phiếu.</div>'; return;}
  $('sheet-detail').innerHTML = `<div class="detail-header"><div class="detail-title-row"><div><h3>${esc(sheet.sheet)}</h3><p>${sheet.selected ? 'Phiếu này đang được chọn để xuất' : 'Phiếu này chưa được chọn để xuất'}${sheet.header ? ` · Tiêu đề nguồn: dòng ${sheet.header}${sheet.header_span > 1 ? '–' + (sheet.header + sheet.header_span - 1) : ''}` : ' · Cần xác định cột dữ liệu'}</p></div>${confidence(sheet)}</div></div><div class="detail-fields"><div class="fields">${[['customer','Khách hàng *','text'],['recipient','Người nhận','text'],['date','Ngày xuất *','date'],['address','Địa chỉ khách hàng','text'],['voucher','Số phiếu xuất','text']].map(([field,label,type]) => `<label class="${field === 'address' ? 'wide' : ''}">${label}<input data-field="${field}" type="${type}" value="${esc(sheet[field])}" placeholder="${field === 'customer' ? 'Nhập tên khách hàng' : field === 'voucher' ? 'Để trống nếu chưa có' : ''}" ${field === 'customer' || field === 'date' ? 'required' : ''} maxlength="${field === 'address' ? 500 : 250}"></label>`).join('')}</div></div><div class="alerts">${renderAlerts(sheet)}</div><nav class="detail-tabs" aria-label="Thông tin sheet"><button type="button" data-tab="items" class="${activeTab === 'items' ? 'active' : ''}" ${activeTab === 'items' ? 'aria-current="page"' : ''}>Mặt hàng <span>${sheet.items.length}</span></button><button type="button" data-tab="source" class="${activeTab === 'source' ? 'active' : ''}" ${activeTab === 'source' ? 'aria-current="page"' : ''}>${isImageSheet(sheet) ? 'Ảnh gốc & văn bản' : 'Xem nguồn Excel'}</button><button type="button" data-tab="mapping" class="${activeTab === 'mapping' ? 'active' : ''}" ${activeTab === 'mapping' ? 'aria-current="page"' : ''}>Chỉnh nhận diện</button></nav><div id="tab-content" class="tab-pane"></div>`;
  renderTab();
}
function filteredItems() {
  const query = itemQuery.trim().toLocaleLowerCase('vi');
  return current().items.map((item, index) => ({item,index})).filter(({item}) => !query || `${item.name} ${item.unit} ${item.code || ''} ${item.note || ''} ${item.row}`.toLocaleLowerCase('vi').includes(query));
}
function itemsMarkup() {
  const sheet = current(), filtered = filteredItems();
  const pages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  page = Math.min(Math.max(1,page),pages);
  const shown = filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);
  return `<div class="items-toolbar"><label class="search-box"><span aria-hidden="true">⌕</span><input id="item-search" type="search" value="${esc(itemQuery)}" placeholder="Tìm hàng, mã, ghi chú…" aria-label="Tìm mặt hàng trong sheet"></label><div class="item-actions"><button class="button-ghost" type="button" data-action="recalculate" ${sheet.items.length ? '' : 'disabled'} title="Tính lại thành tiền từ số lượng × đơn giá cho các dòng còn chọn">Tính lại thành tiền</button><button class="button-ghost" type="button" data-action="undo" ${changedRows(sheet) ? '' : 'disabled'}>↶ Hoàn tác sheet</button></div></div><div class="table-wrap"><table class="item-table"><colgroup><col class="row-column"><col class="name-column"><col class="unit-column"><col class="quantity-column"><col class="price-column"><col class="amount-column"><col class="note-column"><col class="action-column"></colgroup><thead><tr><th title="Dòng trong Excel nguồn">Dòng</th><th>Tên hàng</th><th>ĐVT</th><th class="num">Số lượng</th><th class="num">Đơn giá</th><th class="num">Thành tiền</th><th>Ghi chú</th><th><span class="visually-hidden">Bỏ / khôi phục dòng</span></th></tr></thead><tbody>${shown.map(({item,index}) => {
    const removed = sheet._excluded.has(item.row);
    const rowIssues = (sheet.issues || []).filter(issue => issue.row === item.row && !(issue.recoverable && issue.field === 'amount' && Math.abs(Math.round(item.quantity * item.price) - item.amount) <= 1));
    const rowLevel = rowIssues.some(issue => issue.level === 'error') ? 'error' : rowIssues.length ? 'warning' : '';
    const issueText = rowIssues.map(issue => issue.message).join(' · ');
    const edited = ['name','unit','quantity','price','amount','note'].some(key => item[key] !== sheet._originalItems[index]?.[key]);
    const input = (field, numeric = false) => `<input class="item-input ${numeric ? 'numeric' : ''} ${item._invalid[field] ? 'invalid' : ''}" data-item="${index}" data-item-field="${field}" type="text" ${numeric ? 'inputmode="decimal"' : ''} value="${esc(numeric ? (item._draft?.[field] ?? number(item[field])) : item[field])}" aria-label="${fields[field]} dòng ${item.row}" ${removed ? 'disabled' : ''} ${numeric ? '' : 'maxlength="500"'}>`;
    return `<tr class="${removed ? 'excluded' : ''} ${edited ? 'edited' : ''}" data-item-row="${index}"><td><span class="source-row ${rowLevel ? 'source-' + rowLevel : ''}" title="${esc(issueText || 'Dòng trong file nguồn')}">${item.row}${rowLevel ? ' !' : ''}</span></td><td>${input('name')}</td><td>${input('unit')}</td><td>${input('quantity',true)}</td><td>${input('price',true)}</td><td class="num"><span class="item-amount">${money(item.amount)}</span></td><td>${input('note')}</td><td><button class="row-toggle" type="button" data-row-toggle="${index}" title="${removed ? 'Khôi phục dòng' : 'Bỏ dòng khỏi phiếu'}" aria-label="${removed ? 'Khôi phục' : 'Bỏ'} dòng ${item.row}">${removed ? '↶' : '×'}</button></td></tr>`;
  }).join('') || `<tr><td colspan="8" class="empty-state">${sheet.items.length ? 'Không có mặt hàng khớp với từ khóa.' : 'Chưa đọc được mặt hàng. Mở “Chỉnh nhận diện” để chọn tiêu đề và cột dữ liệu.'}</td></tr>`}</tbody></table></div><div class="table-summary"><span class="table-count">${includedItems(sheet).length} dòng sẽ xuất${changedRows(sheet) ? ` · ${changedRows(sheet)} dòng đã sửa / bỏ` : ''}</span><strong class="sheet-total">${money(sheetTotal(sheet))} ₫</strong></div><div class="table-footer"><span>${filtered.length ? `${(page - 1) * PAGE_SIZE + 1}–${Math.min(page * PAGE_SIZE,filtered.length)}` : '0'} / ${filtered.length} dòng${itemQuery ? ' khớp từ khóa' : ''}</span><div class="pagination"><button type="button" data-page="-1" ${page === 1 ? 'disabled' : ''} aria-label="Trang trước">‹</button><span>Trang ${page} / ${pages}</span><button type="button" data-page="1" ${page >= pages ? 'disabled' : ''} aria-label="Trang sau">›</button></div></div><p class="table-hint">Sửa trực tiếp tên hàng, ĐVT, số lượng, đơn giá hoặc ghi chú. Sửa số lượng / đơn giá sẽ tính lại thành tiền. Nhập số dạng Việt Nam: 1.000 hoặc 1,5. Dấu × chỉ bỏ dòng khỏi phiếu xuất. Trên điện thoại, kéo bảng sang ngang để xem số lượng, đơn giá và thành tiền.</p>`;
}
function mappingMarkup() {
  const sheet = current();
  return `<p class="mapping-note">Chọn lại dòng tiêu đề và cột nếu bảng nhận diện chưa đúng. Bạn có thể đối chiếu số dòng và tên cột ở tab “Xem nguồn Excel”.</p><div class="mapping-header"><label>Dòng bắt đầu tiêu đề<input id="mapping-header" class="header-row" type="number" min="1" max="10000" value="${sheet.header || 1}"></label><label>Số dòng tiêu đề<input id="mapping-span" type="number" min="1" max="3" value="${sheet.header_span || 1}"></label></div><div class="mapping">${Object.entries(fields).map(([key,label]) => `<label>${label}${['name','quantity'].includes(key) ? ' *' : ''}<select data-map="${key}"><option value="">Không có</option>${sheet.columns.map(column => `<option value="${esc(column)}" ${sheet.mapping[key] === column ? 'selected' : ''}>Cột ${esc(column)}</option>`).join('')}</select></label>`).join('')}</div><p class="detection-notes">${(sheet.detection_notes || []).map(note => `<span>${esc(note)}</span>`).join('<br>')}</p><div class="mapping-actions"><p>Đọc lại sẽ thay các chỉnh sửa mặt hàng của sheet này bằng dữ liệu từ file nguồn.</p><button class="button-secondary remap" type="button" data-action="remap">Đọc lại theo cột đã chọn</button></div>`;
}
function sourceMarkup() {
  const sheet = current(), rows = sheet.preview_rows || [];
  if (!rows.length) return '<div class="empty-state">Bản xem trước chưa khả dụng. Đối chiếu file Excel nguồn để chọn lại cột.</div>';
  const columns = [...new Set(rows.flatMap(row => Object.keys(row.cells || {})))];
  return `<div class="preview-intro"><p>Trích các dòng đầu từ Excel để kiểm tra tiêu đề và cột. Dữ liệu nguồn được giữ nguyên.</p><span class="mini-tag">${rows.length} dòng xem trước</span></div><div class="table-wrap preview-wrap"><table class="preview-table"><thead><tr><th>Dòng</th>${columns.map(column => `<th>Cột ${esc(column)}</th>`).join('')}</tr></thead><tbody>${rows.map(row => `<tr class="${row.row >= sheet.header && row.row < sheet.header + (sheet.header_span || 1) ? 'header-source' : ''}"><td>${row.row}</td>${columns.map(column => `<td>${esc(row.cells?.[column])}</td>`).join('')}</tr>`).join('')}</tbody></table></div><p class="table-hint">Dòng tiêu đề được tô xanh. Nếu tên cột nằm trên nhiều dòng, chỉnh “Số dòng tiêu đề” trong tab nhận diện.</p>`;
}
function renderTab() {
  const container = $('tab-content');
  const sheet = current();
  container.innerHTML = activeTab === 'items' ? (isImageSheet(sheet) && window.ImageImport ? window.ImageImport.itemsLayout(itemsMarkup(),sheet) : itemsMarkup()) : activeTab === 'mapping' ? mappingMarkup() : (isImageSheet(sheet) && window.ImageImport ? window.ImageImport.sourceMarkup(sheet) : sourceMarkup());
  if (busy) setBusy(busy);
}
function renderAll() {
  clearGeneratedDownload();
  $('workspace').hidden = false;
  $('welcome-guide').hidden = true;
  $('source-filename').textContent = filename;
  $('step-upload').className = 'workflow-step done';
  $('step-review').className = 'workflow-step active';
  $('step-export').className = 'workflow-step';
  renderNavigation(); renderDetail(); syncTotals();
}
function numericValue(raw) {
  let value = String(raw).trim();
  if (/\s/.test(value)) {
    if (!/^[+-]?\d{1,3}(?:[ \u00a0\u202f]\d{3})+(?:[.,]\d+)?$/.test(value)) return null;
    value = value.replace(/\s/g,'');
  }
  if (!value || !/^[+-]?[\d.,]+$/.test(value)) return null;
  if (value.includes(',') && value.includes('.')) {
    const decimal = value.lastIndexOf(',') > value.lastIndexOf('.') ? ',' : '.';
    const grouping = decimal === ',' ? '.' : ',';
    const position = value.lastIndexOf(decimal), integer = value.slice(0,position), fraction = value.slice(position+1);
    const groupedInteger = new RegExp(`^[+-]?\\d{1,3}(?:\\${grouping}\\d{3})+$`);
    if (!groupedInteger.test(integer) || !/^\d+$/.test(fraction)) return null;
    value = integer.split(grouping).join('') + '.' + fraction;
  } else if ((value.match(/,/g) || []).length > 1) {
    if (!/^[+-]?\d{1,3}(,\d{3})+$/.test(value)) return null;
    value = value.replace(/,/g,'');
  } else if (value.includes(',')) value = value.replace(',','.');
  else if (/^[+-]?[1-9]\d{0,2}(\.\d{3})+$/.test(value)) value = value.replace(/\./g,'');
  const result = Number(value);
  return Number.isFinite(result) && result >= 0 ? result : null;
}
function exportSettings() {
  return sheets.filter(sheet => sheet.selected).map(sheet => ({sheet:sheet.sheet, selected:true, ...Object.fromEntries(metadata.map(field => [field,sheet[field]])), items:sheet.items.map(item => ({row:item.row, name:item.name, unit:item.unit || '', code:item.code || '', note:item.note || '', quantity:item.quantity, price:item.price, amount:item.amount})), excluded_rows:[...sheet._excluded], ...(isImageSheet(sheet) ? {ocr_confirmed:sheet._ocrConfirmed} : {})}));
}
function fileInformation(input, kind) {
  const file = input.files[0];
  const container = $(`${kind}-drop`);
  container.classList.toggle('file-ready',!!file);
  if (kind === 'data') {
    $('data-label').textContent = file ? file.name : 'Bấm vào đây để chọn file Excel';
    $('data-info').textContent = file ? `${(file.size / 1024 / 1024).toLocaleString('vi-VN',{maximumFractionDigits:2})} MB · Sẵn sàng đọc file` : 'Có thể kéo thả file vào ô này';
    $('data-note').textContent = file ? 'Bấm để đổi file. Kết quả đang mở được giữ đến khi nhận diện file mới thành công.' : 'File Excel chứa tên hàng, số lượng, đơn giá hoặc thành tiền.';
  } else {
    $('template-label').textContent = file ? file.name : 'Template chuẩn đang sẵn sàng';
    $('template-info').textContent = file ? `${(file.size / 1024).toLocaleString('vi-VN',{maximumFractionDigits:1})} KB · Template tùy chọn` : 'Dùng mẫu Huyền An 68 đã cung cấp';
  }
}
['data','template'].forEach(kind => {
  const input = $(`${kind}-file`), drop = $(`${kind}-drop`);
  input.addEventListener('change',() => fileInformation(input,kind));
  ['dragenter','dragover'].forEach(type => drop.addEventListener(type,event => {event.preventDefault(); if (!busy) drop.classList.add('dragging');}));
  ['dragleave','drop'].forEach(type => drop.addEventListener(type,event => {event.preventDefault();drop.classList.remove('dragging');}));
  drop.addEventListener('drop',event => {
    if (busy) return;
    const files = event.dataTransfer?.files;
    if (!files?.length) return;
    if (files.length > 1) {notify('Mỗi ô chỉ nhận một file Excel. Hãy chọn một workbook có nhiều sheet.',true);return;}
    if (!/\.xlsx$/i.test(files[0].name)) {notify('Hãy chọn file .xlsx. File .xls cần lưu lại thành .xlsx trong Excel.',true);return;}
    try {input.files = files;fileInformation(input,kind);} catch (_) {notify('Trình duyệt này chưa hỗ trợ thả file. Bấm “Chọn file” để tải Excel.',true);}
  });
});
$('upload').addEventListener('submit',async event => {
  event.preventDefault();
  if (busy) return;
  if (window.ImageImport?.mode() === 'images') {await window.ImageImport.analyze();return;}
  const body = new FormData(event.target);
  const candidateFiles = {data:$('data-file').files[0],template:$('template-file').files[0] || null};
  const files = Object.values(candidateFiles).filter(Boolean);
  if (files.some(file => !/\.xlsx$/i.test(file.name))) {notify('Hãy chọn file .xlsx. File .xls hoặc .xlsm cần lưu lại thành .xlsx trong Excel.',true);return;}
  if (token && transport !== 'stateless') body.append('replace_token',token);
  setBusy('analyze');notify('Đang quét các sheet, tìm bảng hàng và đối chiếu dữ liệu…',false,true);
  try {
    await runtimeReady;
    if (files.reduce((total,file) => total + file.size,0) > maxUploadBytes) throw new Error(`Tổng dung lượng dữ liệu và template vượt ${number(maxUploadBytes / 1024 / 1024)} MB. Hãy chia workbook hoặc giảm dung lượng.`);
    const data = await (await api('/api/analyze',body)).json();
    if (((data.transport || transport) !== 'stateless' && !data.token) || !Array.isArray(data.sheets) || !data.sheets.length) throw new Error('Workbook không trả về sheet nào để kiểm tra.');
    const nextSheets = data.sheets.map(prepareSheet);
    configureRuntime(data);
    token = data.token || '';sheets = nextSheets;filename = data.filename || candidateFiles.data?.name || 'File Excel';
    analyzedFiles = candidateFiles;overrides = {};
    window.ImageImport?.clearResults();
    activeIndex = Math.max(0,sheets.findIndex(sheet => sheet.selected));activeTab = 'items';itemQuery = '';page = 1;$('sheet-search').value = '';
    renderAll();
    const skipped = sheets.filter(sheet => sheet.summary || isHidden(sheet)).length;
    notify(`Đã quét ${sheets.length} sheet${skipped ? `; ${skipped} sheet tổng hợp / ẩn được bỏ chọn mặc định` : ''}. Kiểm tra tên khách hàng, ngày xuất và các lưu ý trước khi tải file.`);
  } catch (error) {notify(error.message,true);} finally {setBusy();}
});
$('sheets').addEventListener('change',event => {
  const index = event.target.dataset.sheetSelect;
  if (index === undefined || busy) return;
  sheets[index].selected = event.target.checked;sheets[index]._selectionTouched = true;
  if (isImageSheet(sheets[index])) sheets[index]._ocrConfirmed = event.target.checked;
  renderNavigation();syncTotals();
  if (Number(index) === activeIndex) renderDetail();
});
$('customer-selection')?.addEventListener('change',event => {
  const index = event.target.dataset.customerSelect;if (index === undefined || busy || !sheets[index]) return;
  const sheet = sheets[index];sheet.selected = event.target.checked;sheet._ocrConfirmed = event.target.checked;sheet._selectionTouched = true;
  activeIndex = Number(index);activeTab = 'items';page = 1;itemQuery = '';renderNavigation();renderDetail();syncTotals();
});
$('sheets').addEventListener('click',event => {
  const button = event.target.closest('[data-open-sheet]');
  if (!button || busy) return;
  activeIndex = Number(button.dataset.openSheet);page = 1;itemQuery = '';activeTab = 'items';renderNavigation();renderDetail();syncTotals();
});
$('sheet-search').addEventListener('input',renderNavigation);
$('select-details').addEventListener('click',() => {
  sheets.forEach(sheet => {if (isImageSheet(sheet)) return;sheet.selected = !sheet.summary && !isHidden(sheet) && !!sheet.items.length && !sheet.errors.length && !sheet.requires_review && sheet.detection_method !== 'inferred';sheet._selectionTouched = true;});
  renderNavigation();renderDetail();syncTotals();notify('Đã chọn các sheet chi tiết có dữ liệu hợp lệ. Các sheet có lỗi, cần xác nhận, tổng hợp và ẩn vẫn được bỏ chọn.');
});
$('deselect-all').addEventListener('click',() => {sheets.forEach(sheet => {sheet.selected = false;sheet._selectionTouched = true;});renderNavigation();renderDetail();syncTotals();});
$('exclude-zero').addEventListener('change',() => {renderNavigation();syncTotals();});
$('apply-bulk').addEventListener('click',() => {
  const values = Object.fromEntries(['customer','recipient','date','address'].map(field => [field,$(`bulk-${field}`).value.trim()]).filter(([,value]) => value));
  if (!Object.keys(values).length) {notify('Nhập ít nhất một thông tin muốn áp dụng chung.',true);return;}
  let count = 0;sheets.filter(sheet => sheet.selected).forEach(sheet => {Object.assign(sheet,values);Object.keys(values).forEach(field => sheet._metadataTouched.add(field));count++;});
  renderNavigation();renderDetail();syncTotals();notify(`Đã cập nhật thông tin chung cho ${count} phiếu đã chọn.`);
});
$('sheet-detail').addEventListener('input',event => {
  if (busy) return;
  const element = event.target, sheet = current();
  if (element.hasAttribute('data-image-confirm')) {sheet.selected = element.checked;sheet._ocrConfirmed = element.checked;sheet._selectionTouched = true;renderNavigation();renderDetail();syncTotals();return;}
  if (element.dataset.field) {
    sheet[element.dataset.field] = element.value;sheet._metadataTouched.add(element.dataset.field);element.classList.remove('invalid');
    syncTotals();renderNavigation();return;
  }
  if (element.id === 'item-search') {
    const selectionStart = element.selectionStart;
    itemQuery = element.value;page = 1;renderTab();const input = $('item-search');input.focus();input.setSelectionRange(selectionStart,selectionStart);return;
  }
  if (element.dataset.itemField) {
    const field = element.dataset.itemField, item = sheet.items[Number(element.dataset.item)];
    if (field === 'quantity' || field === 'price') {
      const value = numericValue(element.value);
      item._invalid[field] = value === null;
      item._draft = item._draft || {};
      if (value === null) item._draft[field] = element.value; else delete item._draft[field];
      element.classList.toggle('invalid',value === null);
      if (value !== null) {
        item[field] = value;
        if (!item._invalid.quantity && !item._invalid.price) item.amount = Math.round(item.quantity * item.price);
      }
    } else {item[field] = element.value;item._invalid[field] = field === 'name' && !element.value.trim();element.classList.toggle('invalid',!!item._invalid[field]);}
    const row = element.closest('tr');row.classList.add('edited');row.querySelector('.item-amount').textContent = money(item.amount);
    syncTotals();renderNavigation();
  }
});
$('sheet-detail').addEventListener('focusout',event => {
  const element = event.target;
  if (['quantity','price'].includes(element.dataset.itemField)) {
    const item = current().items[Number(element.dataset.item)];
    if (!item._invalid[element.dataset.itemField]) element.value = number(item[element.dataset.itemField]);
  }
});
$('sheet-detail').addEventListener('click',async event => {
  if (busy) return;
  const button = event.target.closest('button');if (!button) return;
  if (button.dataset.tab) {activeTab = button.dataset.tab;renderDetail();return;}
  if (button.dataset.page) {page += Number(button.dataset.page);renderTab();return;}
  if (button.dataset.rowToggle !== undefined) {
    const sheet = current(), item = sheet.items[Number(button.dataset.rowToggle)];
    if (sheet._excluded.has(item.row)) sheet._excluded.delete(item.row); else sheet._excluded.add(item.row);
    renderTab();renderNavigation();syncTotals();return;
  }
  if (button.dataset.action === 'recalculate') {
    const sheet = current();let count = 0;
    includedItems(sheet).forEach(item => {
      if (!item._invalid.quantity && !item._invalid.price && Number.isFinite(item.quantity * item.price)) {
        const calculated = Math.round(item.quantity * item.price);
        if (item.amount !== calculated) {item.amount = calculated;count++;}
      }
    });
    renderTab();renderNavigation();syncTotals();notify(`Đã tính lại thành tiền từ số lượng × đơn giá${count ? `; cập nhật ${count} dòng` : '; không có dòng nào cần thay đổi'}. Kiểm tra số lượng và đơn giá trước khi xuất.`);return;
  }
  if (button.dataset.action === 'undo') {
    const sheet = current();sheet.items = clone(sheet._originalItems);sheet._excluded.clear();renderTab();renderNavigation();syncTotals();notify(`Đã khôi phục mặt hàng của sheet “${sheet.sheet}” theo lần nhận diện gần nhất.`);return;
  }
  if (button.dataset.action === 'remap') {
    const sheet = current();
    const mapping = Object.fromEntries([...document.querySelectorAll('[data-map]')].map(select => [select.dataset.map,select.value]));
    const header = Number($('mapping-header').value), span = Number($('mapping-span').value);
    if (!Number.isInteger(header) || header < 1 || !Number.isInteger(span) || span < 1 || span > 3) {notify('Dòng tiêu đề phải là số nguyên dương; số dòng tiêu đề từ 1 đến 3.',true);return;}
    if (!mapping.name || !mapping.quantity || (!mapping.price && !mapping.amount)) {notify('Chọn ít nhất cột tên hàng, số lượng và đơn giá hoặc thành tiền.',true);return;}
    const selectedColumns = Object.values(mapping).filter(Boolean);
    if (new Set(selectedColumns).size !== selectedColumns.length) {notify('Mỗi loại dữ liệu cần một cột riêng. Kiểm tra các cột đang chọn trùng nhau.',true);return;}
    const previous = sheets;const target = sheet.sheet;
    const candidateOverrides = {...overrides,[target]:{header,header_span:span,mapping}};
    setBusy('remap');notify(`Đang đọc lại sheet “${target}” theo cột đã chọn…`,false,true);
    try {
      const data = await (await api('/api/remap',{token,overrides:transport === 'stateless' ? candidateOverrides : {[target]:{header,header_span:span,mapping}}})).json();
      if (!Array.isArray(data.sheets) || !data.sheets.length) throw new Error('Máy chủ chưa trả về kết quả nhận diện. Hãy thử lại.');
      sheets = data.sheets.map(source => {
        const fresh = prepareSheet(source), old = previous.find(entry => entry.sheet === source.sheet);
        if (!old) return fresh;
        metadata.forEach(field => {if (source.sheet !== target || old._metadataTouched.has(field)) fresh[field] = old[field];});
        fresh._selectionTouched = old._selectionTouched;
        fresh._metadataTouched = old._metadataTouched;
        if (old._ocrDraft !== undefined) fresh._ocrDraft = old._ocrDraft;
        if (old._selectionTouched || old.items.length) fresh.selected = old.selected;
        if (isImageSheet(old)) {fresh.image_info = old.image_info;fresh.source_type = 'image';fresh.requires_review = true;fresh._ocrConfirmed = source.sheet === target ? false : old._ocrConfirmed;if (source.sheet === target) fresh.selected = false;}
        if (source.sheet !== target) {fresh.items = old.items;fresh._originalItems = old._originalItems;fresh._excluded = old._excluded;}
        return fresh;
      });
      overrides = candidateOverrides;
      activeIndex = Math.max(0,sheets.findIndex(entry => entry.sheet === target));page = 1;itemQuery = '';activeTab = 'items';renderAll();
      notify(`Đã đọc lại sheet “${target}”. Kiểm tra mặt hàng và các lưu ý mới.`);
    } catch (error) {notify(error.message,true);} finally {setBusy();}
  }
});
function exportProgress(message = '') {
  const element = $('export-progress');
  if (element) {element.hidden = !message;element.textContent = message;}
}
function clearGeneratedDownload() {
  if (generatedDownloadUrl) URL.revokeObjectURL(generatedDownloadUrl);
  generatedDownloadUrl = '';
  const link = $('export-download');
  if (link) {link.hidden = true;link.removeAttribute('href');link.removeAttribute('download');}
}
$('export-cancel')?.addEventListener('click',() => {
  if (!exportTask) return;
  exportTask.reason = 'cancelled';exportTask.controller.abort();
});
window.addEventListener('pagehide',clearGeneratedDownload);
$('export').addEventListener('click',async () => {
  if (busy) return;
  const selected = sheets.filter(sheet => sheet.selected);
  if (!selected.length) {notify('Chọn ít nhất một sheet để xuất.',true);return;}
  for (const sheet of selected) {
    let error = '';
    if (isImageSheet(sheet) && !sheet._ocrConfirmed) error = 'Xác nhận đã kiểm tra dữ liệu OCR';
    else if (!sheet.customer.trim()) error = 'Nhập tên khách hàng';
    else if (!sheet.date || !/^\d{4}-\d{2}-\d{2}$/.test(sheet.date)) error = 'Chọn ngày xuất hợp lệ';
    else if (!includedItems(sheet).length) error = 'Phiếu không còn dòng hàng nào để xuất';
    else {
      const invalid = includedItems(sheet).find(item => !String(item.name || '').trim() || Object.values(item._invalid).some(Boolean) || ![item.quantity,item.price,item.amount].every(value => Number.isFinite(Number(value)) && Number(value) >= 0));
      if (invalid) error = `Kiểm tra tên hàng và các số tại dòng ${invalid.row}`;
    }
    if (error) {
      activeIndex = sheets.indexOf(sheet);activeTab = 'items';page = 1;itemQuery = '';renderNavigation();renderDetail();syncTotals();
      if (!sheet.customer.trim()) document.querySelector('[data-field="customer"]')?.classList.add('invalid');
      if (!sheet.date) document.querySelector('[data-field="date"]')?.classList.add('invalid');
      notify(`${error} trong sheet “${sheet.sheet}”.`,true);$('sheet-detail').scrollIntoView({block:'start',behavior:'smooth'});return;
    }
  }
  const format = $('export-format').value;
  const body = {token,sheets:exportSettings(),exclude_zero:$('exclude-zero').checked,format};
  const task = {controller:new AbortController(),reason:''};exportTask = task;
  let abortExport;
  const cancellation = new Promise((resolve,reject) => {abortExport = () => reject(new Error('Đã dừng tạo file.'));});
  task.controller.signal.addEventListener('abort',abortExport,{once:true});
  const waitForExport = promise => Promise.race([promise,cancellation]);
  const started = Date.now();
  const timeout = setTimeout(() => {task.reason = 'timeout';task.controller.abort();},EXPORT_TIMEOUT_MS);
  const elapsed = setInterval(() => {
    const seconds = Math.floor((Date.now()-started)/1000);
    exportProgress(`Đang tạo ${selected.length} phiếu · ${seconds} giây. Bạn có thể bấm “Hủy tạo file” nếu cần.`);
  },5000);
  setBusy('export');notify('Đang kiểm tra dữ liệu và tạo phiếu theo template…',false,true);
  if ($('export-cancel')) {$('export-cancel').hidden = false;$('export-cancel').disabled = false;}
  exportProgress(`Đang tạo ${selected.length} phiếu. Vui lòng chờ; dữ liệu đã sửa được giữ trên trang.`);
  try {
    const response = await waitForExport(api('/api/export',body,{signal:task.controller.signal}));
    const disposition = response.headers.get('Content-Disposition') || '';
    const filenameMatch = disposition.match(/filename\*=UTF-8''([^;]+)/i) || disposition.match(/filename="?([^";]+)"?/i);
    let downloadName = format === 'zip' ? 'Phieu-xuat-kho.zip' : 'Phieu-xuat-kho.xlsx';
    if (filenameMatch) {try {downloadName = decodeURIComponent(filenameMatch[1]);} catch (_) {}}
    const blob = await waitForExport(response.blob());
    const signature = new Uint8Array(await waitForExport(blob.slice(0,4).arrayBuffer()));
    if (signature.length !== 4 || signature[0] !== 80 || signature[1] !== 75 || signature[2] !== 3 || signature[3] !== 4) throw new Error('File trả về chưa phải Excel hoặc ZIP hợp lệ. Bấm “Tạo và tải file” để thử lại.');
    clearGeneratedDownload();
    const url = URL.createObjectURL(blob);generatedDownloadUrl = url;
    const link = $('export-download');
    if (link) {link.href = url;link.download = downloadName;link.textContent = 'Tải lại file đã tạo';link.title = downloadName;link.hidden = false;}
    const anchor = document.createElement('a');anchor.href = url;anchor.download = downloadName;document.body.appendChild(anchor);
    try {anchor.click();} catch (_) { /* The visible link also starts a user-initiated download. */ } finally {anchor.remove();}
    $('step-review').className = 'workflow-step done';$('step-export').className = 'workflow-step active';
    const message = `Đã tạo ${selected.length} phiếu${format === 'zip' ? ' trong file ZIP' : ' trong file Excel'}. Nếu chưa thấy file tải về, bấm “Tải lại file đã tạo”.`;
    notify(message);exportProgress(message);
  } catch (error) {
    const message = task.reason === 'timeout'
      ? `Chưa tạo xong file sau ${EXPORT_TIMEOUT_MS/1000} giây. Dữ liệu đã sửa vẫn được giữ. Bấm “Tạo và tải file” để thử lại, hoặc chọn ít phiếu hơn.`
      : task.reason === 'cancelled'
        ? 'Đã hủy tạo file. Dữ liệu đã sửa vẫn được giữ. Bấm “Tạo và tải file” khi bạn muốn thử lại.'
        : error.message;
    notify(message,task.reason !== 'cancelled');exportProgress(message);
  } finally {
    clearTimeout(timeout);clearInterval(elapsed);task.controller.signal.removeEventListener('abort',abortExport);
    if (exportTask === task) exportTask = null;
    if ($('export-cancel')) $('export-cancel').hidden = true;
    setBusy();
  }
});
const runtimeReady = initializeRuntime();
