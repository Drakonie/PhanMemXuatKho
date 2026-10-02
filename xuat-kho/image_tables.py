"""Expand shared goods tables into reviewable customer-specific OCR documents."""
from __future__ import annotations
import csv
import io
from datetime import date
import re
from engine import normalize


def validate_table(value):
    if value is None:
        return None
    if not isinstance(value, dict):
        raise ValueError('Bảng OCR theo khách hàng không hợp lệ.')
    groups, rows = value.get('groups'), value.get('rows')
    if not isinstance(groups, list) or not 2 <= len(groups) <= 30 or not isinstance(rows, list) or not 1 <= len(rows) <= 200:
        raise ValueError('Bảng ảnh hỗ trợ 2–30 khách hàng, tối đa 200 dòng mỗi vùng đọc.')
    result = {'groups': [], 'rows': [], 'method': 'grid-cell-ocr', 'price_inferred': value.get('price_inferred') is True,
              'header_inferred': value.get('header_inferred') is True}
    def text(value):
        if not isinstance(value, str) or len(value) > 1000:
            raise ValueError('Ô dữ liệu OCR không hợp lệ hoặc quá dài.')
        return value.strip()
    def score(value):
        if type(value) not in (int,float) or not 0 <= value <= 100:
            raise ValueError('Điểm OCR không hợp lệ.')
        return value
    for index, group in enumerate(groups, 1):
        if not isinstance(group, dict): raise ValueError('Khách hàng OCR không hợp lệ.')
        result['groups'].append({'name':text(group.get('name','')), 'index':index, 'confidence':score(group.get('confidence',0))})
    for row in rows:
        if not isinstance(row, dict) or not isinstance(row.get('values'), list) or len(row['values']) != len(groups):
            raise ValueError('Số ô khách hàng không khớp số nhóm.')
        entry = {field:text(row.get(field,'')) for field in ('name','unit','price')}
        entry['confidence'] = score(row.get('confidence',0))
        entry['values'] = []
        for cells in row['values']:
            if not isinstance(cells,dict): raise ValueError('Ô số liệu khách hàng không hợp lệ.')
            entry['values'].append({'quantity':text(cells.get('quantity','')), 'amount':text(cells.get('amount','')), 'confidence':score(cells.get('confidence',0))})
        result['rows'].append(entry)
    return result


def _date(text):
    plain=normalize(text)
    match=re.search(r'(\d{1,2})\s+thang\s+(\d{1,2})\s+nam\s+(\d{4})',plain) or re.search(r'\b(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})\b',plain)
    if match:
        try:return date(int(match[3]),int(match[2]),int(match[1])).isoformat()
        except ValueError:pass
    return ''


def expand_customers(documents):
    expanded=[];info=[]
    for image_index, document in enumerate(documents):
        table=validate_table(document.get('table_data'))
        if not table:
            expanded.append(document);info.append(None);continue
        edits=document.get('customer_edits',{})
        unknown=set(edits)-{str(i) for i in range(1,len(table['groups'])+1)}
        if unknown:raise ValueError('Bản sửa tham chiếu nhóm khách hàng không tồn tại.')
        day=_date(document.get('text',''))
        for index, group in enumerate(table['groups']):
            name=group['name']
            unresolved=not name or '[OCR:' in name
            title=(document['filename'].rsplit('.',1)[0]+' · '+(name if not unresolved else f'Khách hàng {index+1}'))
            lines=[]
            if not unresolved:lines.append('Khách hàng: '+name)
            if day:lines.append('Ngày xuất: '+day)
            lines.append('Tên hàng | ĐVT | Số lượng | Đơn giá | Thành tiền')
            sources=[]
            for source_index, row in enumerate(table['rows'],1):
                values=row['values'][index]
                if not values['quantity'] and not values['amount']:continue
                quantity=values['quantity']
                if not row['name'] or '[OCR:' in row['name']:
                    quantity='[OCR: thiếu tên hàng] '+quantity
                cells=[row['name'] or '[OCR: thiếu tên hàng]',row['unit'],quantity,row['price'],values['amount']]
                # csv quoting retains names containing | instead of shifting money columns.
                buffer=io.StringIO();csv.writer(buffer,delimiter='|',lineterminator='').writerow(cells)
                lines.append(buffer.getvalue());sources.append({'source_row':source_index,'confidence':min(row['confidence'],values['confidence'])})
            corrected='\n'.join(lines)
            edited=str(index+1) in edits
            if edited:corrected=edits[str(index+1)]
            expanded.append({'filename':title+'.png','text':corrected,'edited_text':True,'rotation':document.get('rotation',0),'language':document.get('language','vie+eng')})
            info.append({'filename':document['filename'],'image_name':document['filename'],'index':image_index,
                         'customer_group':{'name':name,'index':index+1,'count':len(table['groups'])},
                         'corrected_text':corrected,'detection_method':'ocr-customer-groups',
                         'ocr_text':document.get('text',''),'ocr_confidence':group['confidence'],
                         'edited_text':edited,'sources':sources,'price_inferred':table.get('price_inferred',False),
                         'header_inferred':table.get('header_inferred',False)})
    return expanded,info
