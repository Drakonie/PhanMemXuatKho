from __future__ import annotations
import ast
import calendar
import copy
import io
import re
import unicodedata
from datetime import datetime, date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.workbook.properties import CalcProperties


def normalize(value):
    text = str(value or '').lower().replace('đ', 'd')
    return re.sub(r'\s+', ' ', ''.join(c for c in unicodedata.normalize('NFD', text) if not unicodedata.combining(c))).strip()

ALIASES = {
    'name': ['ten hang', 'ten hang hoa', 'ten san pham', 'hang hoa', 'mat hang', 'ten vat tu', 'ten thuc pham', 'san pham'],
    'quantity': ['so luong', 'sl', 'thuc xuat', 'sl xuat', 'so luong xuat', 'khoi luong'],
    'price': ['don gia', 'gia ban', 'gia', 'dg'],
    'amount': ['thanh tien', 'thanh tien cong no', 'tong tien', 'tien hang'],
    'unit': ['dvt', 'don vi tinh', 'don vi'],
    'code': ['ma hang', 'ma hang hoa', 'ma san pham', 'ma vat tu', 'sku'],
    'note': ['ghi chu', 'dien giai'],
}

def field_of(value):
    text = normalize(value).replace('.', '')
    for key, names in ALIASES.items():
        for name in names:
            if text == name or (len(name) > 3 and re.match(r'^' + re.escape(name) + r'(?:\s|\(|/)', text)):
                return key
    return None


def number(value):
    if isinstance(value, bool) or value is None or str(value).strip() == '':
        return None
    if isinstance(value, (int, float, Decimal)):
        result = Decimal(str(value))
    else:
        text = re.sub(r'(?i)\s|vnd|vnđ|đồng|₫|đ', '', str(value))
        if ',' in text and '.' in text:
            # The last separator determines the decimal separator.
            if text.rfind(',') > text.rfind('.'):
                text = text.replace('.', '').replace(',', '.')
            else:
                text = text.replace(',', '')
        elif ',' in text or '.' in text:
            sep = ',' if ',' in text else '.'
            parts = text.split(sep)
            if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3 and parts[0] != '0'):
                text = ''.join(parts)
            elif sep == ',':
                text = text.replace(',', '.')
        try:
            result = Decimal(text)
        except InvalidOperation:
            return None
    return result if result.is_finite() else None


class Values:
    """Prefer Excel's saved results; safely evaluate basic uncached formulas."""
    def __init__(self, raw, cached):
        self.raw, self.cached, self.memo, self.active = raw, cached, {}, set()

    def get(self, sheet, coord):
        key = (sheet, coord)
        if key in self.memo:
            return self.memo[key]
        value = self.raw[sheet][coord].value
        if not isinstance(value, str) or not value.startswith('='):
            return value
        saved = self.cached[sheet][coord].value
        if saved is not None:
            return saved
        if key in self.active or len(self.active) > 80:
            raise ValueError('Công thức vòng hoặc quá phức tạp')
        self.active.add(key)
        try:
            expression = value[1:].replace('$', '')
            def reference(match):
                target = match.group(1) or match.group(2) or sheet
                target = target.replace("''", "'")
                resolved = self.get(target, match.group(3))
                if resolved is None:
                    resolved = 0
                parsed = number(resolved)
                if parsed is None:
                    raise ValueError('Công thức tham chiếu ô không phải số')
                return str(parsed)
            def sum_range(match):
                from openpyxl.utils.cell import range_boundaries
                a,b,c,d = range_boundaries(match.group(1))
                if (c-a+1)*(d-b+1) > 100000:
                    raise ValueError('Vùng SUM quá lớn')
                total = Decimal(0)
                for row in range(b,d+1):
                    for col in range(a,c+1):
                        total += number(self.get(sheet, f'{get_column_letter(col)}{row}')) or Decimal(0)
                return str(total)
            expression = re.sub(r'(?i)SUM\(([A-Z]+\d+:[A-Z]+\d+)\)', sum_range, expression)
            expression = re.sub(r"(?:'((?:[^']|'')+)'!|([A-Za-z_][\w .-]*)!)?([A-Z]{1,3}\d+)", reference, expression)
            def evaluate(node):
                if isinstance(node, ast.Expression): return evaluate(node.body)
                if isinstance(node, ast.Constant) and type(node.value) in (int,float): return Decimal(str(node.value))
                if isinstance(node, ast.UnaryOp) and isinstance(node.op,(ast.UAdd,ast.USub)):
                    return evaluate(node.operand) * (-1 if isinstance(node.op,ast.USub) else 1)
                if isinstance(node, ast.BinOp) and isinstance(node.op,(ast.Add,ast.Sub,ast.Mult,ast.Div)):
                    left,right = evaluate(node.left),evaluate(node.right)
                    if isinstance(node.op,ast.Add): return left+right
                    if isinstance(node.op,ast.Sub): return left-right
                    if isinstance(node.op,ast.Mult): return left*right
                    return left/right
                raise ValueError('Công thức chưa hỗ trợ; hãy tính lại và lưu file trong Excel')
            result = evaluate(ast.parse(expression,mode='eval'))
            self.memo[key] = result
            return result
        finally:
            self.active.remove(key)


def customer_and_date(sheet, header):
    customer = ''
    found_date = ''
    candidates = []
    for row in sheet.iter_rows(min_row=1,max_row=max(1,header-1),max_col=min(sheet.max_column,60)):
        for cell in row:
            value = cell.value
            if isinstance(value,(date,datetime)):
                found_date = value.strftime('%Y-%m-%d')
            if not isinstance(value,str) or value.startswith('='):
                continue
            norm = normalize(value)
            match = re.search(r'ngay\s+(\d{1,2})\s+thang\s+(\d{1,2})\s+nam\s+(\d{4})',norm)
            if not match:
                match = re.search(r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b',norm)
            if match:
                try: found_date = date(int(match[3]),int(match[2]),int(match[1])).isoformat()
                except ValueError: pass
            label = re.match(r'^(?:ten khach hang|khach hang|ten don vi|don vi mua|nguoi mua hang)\s*[:：]?\s*(.*)$',norm)
            if label:
                after = re.split(r'[:：]',value,maxsplit=1)
                if len(after)==2 and after[1].strip(): customer=after[1].strip()
                elif label[1]: customer = re.sub(r'^(?:tên khách hàng|khách hàng|tên đơn vị|đơn vị mua|người mua hàng)\s*', '',value,flags=re.I).strip()
                else:
                    for col in range(cell.column+1,min(sheet.max_column,cell.column+8)+1):
                        adjacent = sheet.cell(cell.row,col).value
                        if adjacent and isinstance(adjacent,str): customer=adjacent.strip(); break
            if any(token in norm for token in ['truong ', 'cong ty ', 'nha hang ', 'mam non ', 'khach san ']):
                candidates.append(value.strip())
    return customer or (candidates[0] if candidates else ''), found_date


def analyze(source, overrides=None):
    raw = load_workbook(io.BytesIO(source),data_only=False)
    cached = load_workbook(io.BytesIO(source),data_only=True)
    values = Values(raw,cached)
    results = []
    for sheet in raw:
        warnings, errors = [], []
        override = (overrides or {}).get(sheet.title)
        mapping, header, score = {},0,0
        if override:
            header=int(override['header'])
            if not 1 <= header <= sheet.max_row: raise ValueError('Dòng tiêu đề ngoài phạm vi sheet')
            mapping={k:column_index_from_string(v) for k,v in override['mapping'].items() if v}
        else:
            for row in sheet.iter_rows(max_row=min(sheet.max_row,100),max_col=min(sheet.max_column,100)):
                detected={}
                for cell in row:
                    field = field_of(cell.value)
                    if field and field not in detected: detected[field]=cell.column
                current = len(detected)+ (5 if 'name' in detected else 0)
                if 'name' in detected and 'quantity' in detected and ('price' in detected or 'amount' in detected) and current>score:
                    mapping,header,score=detected,row[0].row,current
        # Some source workbooks keep notes beside the table without a header.
        if header and not override and 'note' not in mapping:
            for col in range(max(mapping.values(),default=0)+1,min(sheet.max_column,100)+1):
                if sheet.cell(header,col).value is not None: continue
                found=False
                for r in range(header+1,min(sheet.max_row,header+50)+1):
                    item_name=sheet.cell(r,mapping.get('name',1)).value
                    if re.match(r'^(tong cong|tong tien|cong tien|tong hop|subtotal|total)(\b|:)',normalize(item_name)): break
                    v=sheet.cell(r,col).value
                    if item_name and isinstance(v,str) and not v.startswith('=') and number(v) is None:
                        found=True; break
                if found:
                    mapping['note']=col
                    warnings.append(f'Nhận diện ghi chú ở cột {get_column_letter(col)} dù cột chưa có tiêu đề; hãy kiểm tra.')
                    break
        summary = any(term in normalize(sheet.title) for term in ['tong hop','tong cong','summary','total'])
        customer,day = customer_and_date(sheet,header or 1)
        result={'sheet':sheet.title,'selected': bool(header) and not summary and sheet.sheet_state=='visible',
                'summary':summary,'header':header,'mapping':{k:get_column_letter(v) for k,v in mapping.items()},
                'columns':[get_column_letter(i) for i in range(1,min(sheet.max_column,100)+1)],
                'customer':customer,'recipient':sheet.title,'date':day,'address':'','voucher':'',
                'items':[],'warnings':warnings,'errors':errors,'total':0}
        results.append(result)
        if not header or not all(k in mapping for k in ['name','quantity']) or not ('price' in mapping or 'amount' in mapping):
            result['selected']=False
            errors.append('Chưa nhận diện đủ tên hàng, số lượng và đơn giá hoặc thành tiền. Hãy chọn dòng tiêu đề và cột thủ công.')
            continue
        if not customer: warnings.append('Chưa tìm được tên khách hàng. Cần nhập trước khi xuất.')
        if not day: warnings.append('Chưa tìm được ngày xuất. Cần nhập trước khi xuất.')
        if summary: warnings.append('Sheet tổng hợp được bỏ chọn mặc định để tránh xuất trùng với các sheet chi tiết.')
        if sheet.sheet_state!='visible': warnings.append('Sheet ẩn được bỏ chọn mặc định.')
        for row_index in range(header+1,sheet.max_row+1):
            def read(field):
                if field not in mapping: return None
                return values.get(sheet.title,f'{get_column_letter(mapping[field])}{row_index}')
            try:
                name=read('name')
                if name is None or str(name).strip()=='': continue
                text=normalize(name)
                if re.match(r'^(tong cong|tong tien|cong tien|tong hop|subtotal|total)(\b|:)',text) or text in ['cong', 'tong']: break
                if field_of(name)=='name': continue
                qty,price,amount = number(read('quantity')),number(read('price')),number(read('amount'))
                if qty is None: raise ValueError('Thiếu hoặc không đọc được số lượng')
                if price is None and amount is not None and qty!=0:
                    price=amount/qty
                    warnings.append(f'Dòng {row_index}: suy ra đơn giá từ thành tiền / số lượng.')
                if price is None: raise ValueError('Thiếu hoặc không đọc được đơn giá')
                if qty<0 or price<0 or (amount is not None and amount<0): raise ValueError('Có số âm; cần kiểm tra trước khi xuất kho')
                calculated=(qty*price).quantize(Decimal('1'),rounding=ROUND_HALF_UP)
                if amount is None:
                    amount=calculated
                    warnings.append(f'Dòng {row_index}: tính thành tiền từ số lượng × đơn giá.')
                else:
                    amount=amount.quantize(Decimal('1'),rounding=ROUND_HALF_UP)
                    if abs(amount-calculated)>1:
                        errors.append(f'Dòng {row_index}: thành tiền {amount} khác số lượng × đơn giá {calculated}. Cần sửa file nguồn.')
                result['items'].append({'row':row_index,'name':str(name).strip(),'quantity':float(qty),'price':float(price),
                                       'amount':int(amount),'unit':str(read('unit') or ''),'code':str(read('code') or ''),'note':str(read('note') or '')})
            except Exception as exc:
                errors.append(f'Dòng {row_index}: {exc}')
        result['total']=sum(item['amount'] for item in result['items'])
        if not result['items']: errors.append('Không có dòng hàng hợp lệ.')
    return results

DIGITS=['không','một','hai','ba','bốn','năm','sáu','bảy','tám','chín']
def money_words(value):
    value=int(value)
    if value<0 or value>=10**18: raise ValueError('Tổng tiền ngoài phạm vi hỗ trợ')
    if value==0: return 'Không đồng.'
    def triple(n,full=False):
        h,t,u=n//100,n//10%10,n%10
        out=[]
        if h or full: out.extend([DIGITS[h],'trăm'])
        if t>1: out.extend([DIGITS[t],'mươi'])
        elif t==1: out.append('mười')
        elif u and (h or full): out.append('lẻ')
        if u: out.append('mốt' if u==1 and t>1 else 'lăm' if u==5 and t>=1 else DIGITS[u])
        return ' '.join(out)
    groups=[]
    while value: groups.append(value%1000); value//=1000
    units=['','nghìn','triệu','tỷ','nghìn tỷ','triệu tỷ']
    parts=[]
    for index in range(len(groups)-1,-1,-1):
        if groups[index]: parts.append((triple(groups[index],bool(parts) and groups[index]<100)+' '+units[index]).strip())
    text=' '.join(parts)+' đồng.'
    return text[0].upper()+text[1:]


def validate_template(data):
    wb=load_workbook(io.BytesIO(data))
    sheet=wb.active
    if normalize(sheet['A5'].value)!='phieu xuat kho' or field_of(sheet['B11'].value)!='name' or 'tong cong' not in normalize(sheet['B27'].value):
        raise ValueError('Template phải theo mẫu đã cung cấp: tiêu đề A5, bảng hàng dòng 11 và tổng cộng dòng 27.')
    return wb


def put(sheet,coord,value):
    sheet[coord]=value
    if isinstance(value,str): sheet[coord].data_type='s'


def export(template, sheets, exclude_zero=False):
    wb=validate_template(template)
    base=wb.active
    selected=[s for s in sheets if s.get('selected')]
    if not selected: raise ValueError('Chọn ít nhất một sheet để xuất.')
    used=set()
    totals={}
    for entry in selected:
        if entry.get('errors'): raise ValueError(f"Sheet {entry['sheet']} có lỗi dữ liệu. Hãy sửa trước khi xuất.")
        customer=str(entry.get('customer','')).strip()
        if not customer: raise ValueError(f"Chưa nhập khách hàng cho sheet {entry['sheet']}.")
        try: day=datetime.strptime(entry.get('date',''),'%Y-%m-%d')
        except ValueError: raise ValueError(f"Ngày xuất không hợp lệ: {entry['sheet']}")
        items=[i for i in entry['items'] if not exclude_zero or i['quantity']!=0]
        if not items: raise ValueError(f"Sheet {entry['sheet']} không còn dòng hàng để xuất.")
        sheet=wb.copy_worksheet(base)
        title=re.sub(r'[\\/*?:\[\]]','-',entry['sheet'])[:31] or 'Phiếu xuất'
        candidate=title; index=2
        while candidate.lower() in used or candidate==base.title:
            candidate=title[:26]+f' ({index})'; index+=1
        used.add(candidate.lower()); sheet.title=candidate
        extra=max(0,len(items)-15)
        if extra:
            shifted=[]
            for merged in list(sheet.merged_cells.ranges):
                if merged.min_row>=27:
                    shifted.append((merged.min_col,merged.min_row+extra,merged.max_col,merged.max_row+extra))
                    sheet.unmerge_cells(str(merged))
            dims={r:copy.copy(d) for r,d in sheet.row_dimensions.items() if r>=27}
            sheet.insert_rows(27,extra)
            for old,dim in sorted(dims.items(),reverse=True):
                if old in sheet.row_dimensions: del sheet.row_dimensions[old]
                dim.index=old+extra; sheet.row_dimensions[old+extra]=dim
            for bounds in shifted: sheet.merge_cells(start_column=bounds[0],start_row=bounds[1],end_column=bounds[2],end_row=bounds[3])
        total_row=27+extra
        for r in range(12,total_row):
            for c in range(1,10):
                cell=sheet.cell(r,c)
                if r>=27: cell._style=copy.copy(sheet.cell(12,c)._style)
                cell.value=None
            if r>=27: sheet.row_dimensions[r].height=sheet.row_dimensions[12].height
        put(sheet,'B6',entry.get('voucher',''))
        sheet['E6']=day
        put(sheet,'B7',customer)
        put(sheet,'B8',entry.get('recipient',''))
        put(sheet,'B9',entry.get('address',''))
        if entry.get('content'): put(sheet,'B10',entry['content'])
        for index,item in enumerate(items,1):
            r=index+11
            vals=[index,item['name'],item.get('code',''),item.get('unit',''),item['quantity'],item['quantity'],item['price'],item['amount'],item.get('note','')]
            for c,value in enumerate(vals,1): put(sheet,f'{get_column_letter(c)}{r}',value)
        total=sum(item['amount'] for item in items)
        # Numeric amounts preserve source totals and display in readers without formula recalculation.
        sheet.cell(total_row,8,total)
        put(sheet,f'B{29+extra}',money_words(total))
        sheet.print_area=f'A1:I{32+extra}'
        sheet.print_options.horizontalCentered=True
        sheet.sheet_properties.pageSetUpPr.fitToPage=True
        sheet.page_setup.fitToWidth=1; sheet.page_setup.fitToHeight=0
        sheet.print_title_rows='1:11'
        totals[sheet.title]=total
    wb.remove(base)
    # The supplied template contains one worksheet; remove any extra template-only sheets.
    for sheet in list(wb):
        if sheet.title not in totals: wb.remove(sheet)
    wb.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True)
    result=io.BytesIO(); wb.save(result)
    return result.getvalue()
