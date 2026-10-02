"""Shared-price customer tables and per-customer OCR edits must stay separate."""
import io,json,sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from openpyxl import load_workbook
from app import analyze_images,read_ocr_documents
from engine import analyze,export
from image_import import build_image_workbook
from image_tables import validate_table
ROOT=Path(__file__).resolve().parents[1]

def document():
 return {'filename':'bang-chung.jpg','text':'Thứ sáu, ngày 02 tháng 10 năm 2026','table_data':{
  'groups':[{'name':'Đào Mỹ','confidence':90},{'name':'Phân hiệu 1 (Cây Đa Hương)','confidence':85}],
  'rows':[{'name':'Gạo tẻ','unit':'kg','price':'10,000','confidence':90,'values':[{'quantity':'2.0','amount':'20,000','confidence':80},{'quantity':'3','amount':'30,000','confidence':90}]},
          {'name':'Rau cải','unit':'kg','price':'5,000','confidence':90,'values':[{'quantity':'1','amount':'5,000','confidence':90},{'quantity':'2','amount':'10,000','confidence':90}]}]}}

class MultiCustomerTests(unittest.TestCase):
 def workbook(self,doc):
  source,report=build_image_workbook([doc]);return analyze(source),report
 def test_separate_customers_shared_prices_date_and_totals(self):
  sheets,report=self.workbook(document());self.assertEqual(len(sheets),2)
  self.assertEqual([s['customer'] for s in sheets],['Đào Mỹ','Phân hiệu 1 (Cây Đa Hương)'])
  self.assertEqual([s['date'] for s in sheets],['2026-10-02']*2)
  self.assertEqual([s['total'] for s in sheets],[25000,40000])
  self.assertTrue(all(not s['errors'] for s in sheets))
  self.assertEqual([r['customer_group']['index'] for r in report['documents']],[1,2])
  self.assertTrue(all(r['filename']=='bang-chung.jpg' for r in report['documents']))
 def test_only_selected_customer_exports(self):
  sheets,_=self.workbook(document());sheets[0]['selected']=False;sheets[1]['selected']=True
  wb=load_workbook(io.BytesIO(export((ROOT/'templates/Template.xlsx').read_bytes(),sheets)))
  self.assertEqual(len(wb.worksheets),1);self.assertEqual(wb.active['H27'].value,40000)
 def test_inferred_price_header_is_disclosed_for_review(self):
  doc=document();doc['table_data']['price_inferred']=True
  _,report=self.workbook(doc)
  self.assertTrue(all(any('đơn giá chung' in warning for warning in record['warnings']) for record in report['documents']))
 def test_inferred_customer_header_survives_api_and_warns(self):
  doc=document();doc['table_data']['header_inferred']=True
  parsed=read_ocr_documents(json.dumps([doc]).encode())
  _,report=self.workbook(parsed[0])
  self.assertTrue(all(record['header_inferred'] and any('tiêu đề Tổng' in warning for warning in record['warnings']) for record in report['documents']))
 def test_reparse_one_group_keeps_sibling_data_and_stable_titles(self):
  doc=document();before,_=self.workbook(doc)
  doc['customer_edits']={'1':'Khách hàng: Tên đã sửa\nNgày xuất: 02/10/2026\nTên hàng | ĐVT | Số lượng | Đơn giá | Thành tiền\nGạo tẻ | kg | 4 | 10000 | 40000'}
  after,report=self.workbook(doc);self.assertEqual(after[0]['customer'],'Tên đã sửa');self.assertEqual(after[0]['total'],40000)
  self.assertEqual([s['sheet'] for s in before],[s['sheet'] for s in after]);self.assertEqual(after[1]['items'],before[1]['items'])
  self.assertTrue(report['documents'][0]['edited_text']);self.assertFalse(report['documents'][1]['edited_text'])
 def test_empty_group_cells_are_not_other_customers_orders(self):
  doc=document();doc['table_data']['rows'][0]['values'][0].update(quantity='',amount='')
  sheets,_=self.workbook(doc);self.assertEqual(len(sheets[0]['items']),1);self.assertEqual(len(sheets[1]['items']),2)
 def test_missing_quantity_stays_blocking(self):
  doc=document();doc['table_data']['rows'][0]['values'][0]['quantity']='[OCR: chưa đọc được]'
  sheets,report=self.workbook(doc);self.assertTrue(sheets[0]['errors']);self.assertTrue(report['documents'][0]['issues'])
 def test_missing_name_cannot_be_exported_as_placeholder(self):
  doc=document();doc['table_data']['rows'][0]['name']='[OCR: chưa đọc được]'
  sheets,_=self.workbook(doc);self.assertTrue(sheets[0]['errors']);self.assertTrue(sheets[1]['errors'])
 def test_customer_name_unresolved_requires_manual_input(self):
  doc=document();doc['table_data']['groups'][0]['name']=''
  sheets,_=self.workbook(doc);self.assertEqual(sheets[0]['customer'],'')
 def test_pipe_in_item_name_does_not_shift_numeric_columns(self):
  doc=document();doc['table_data']['rows'][0]['name']='Gạo | loại 1'
  sheets,_=self.workbook(doc);self.assertEqual(sheets[0]['items'][0]['name'],'Gạo | loại 1');self.assertEqual(sheets[0]['total'],25000)
 def test_unknown_edit_group_rejected(self):
  doc=document();doc['customer_edits']={'3':'sai'}
  with self.assertRaisesRegex(ValueError,'không tồn tại'):self.workbook(doc)
 def test_api_passes_valid_customer_edits_and_table(self):
  doc=document();doc['customer_edits']={'1':'Khách hàng: Test'}
  parsed=read_ocr_documents(json.dumps([doc]).encode());self.assertEqual(parsed[0]['customer_edits'],doc['customer_edits']);self.assertEqual(len(parsed[0]['table_data']['groups']),2)
 def test_invalid_table_shape_and_cell_types_rejected(self):
  for transform in [lambda d:d['rows'][0].update(values=[]),lambda d:d['groups'][0].update(name=123),lambda d:d['rows'][0].update(confidence=float('nan'))]:
   table=document()['table_data'];transform(table)
   with self.assertRaises(ValueError):validate_table(table)
 def test_invalid_edit_keys_or_duplicate_groups_rejected(self):
  for edits in [{'0':'a'},{'01':'a','1':'b'},{'x':'a'},{'31':'a'},{'١':'a'},{'1':[]}]:
   doc=document();doc['customer_edits']=edits
   with self.assertRaises(ValueError):read_ocr_documents(json.dumps([doc]).encode())

if __name__=='__main__':unittest.main()
