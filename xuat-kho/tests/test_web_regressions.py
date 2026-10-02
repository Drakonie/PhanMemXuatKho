"""Real row-loss, formula and numeric regressions found during the web review."""
import io
import sys
import unittest
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from app import apply_settings,zip_name
from engine import analyze,number


def data(rows):
    workbook=Workbook()
    for row in rows:workbook.active.append(row)
    output=io.BytesIO();workbook.save(output)
    return output.getvalue()


PREFIX=[['Khách hàng: Trường ABC'],['Ngày xuất: 01/10/2026']]
HEADER=['Tên hàng','Số lượng','Đơn giá','Thành tiền']


class WebRegressionTests(unittest.TestCase):
    def test_shorthand_header_keeps_both_tables(self):
        sheets=analyze(data(PREFIX+[['Hàng','SL','Đơn giá','Thành tiền'],
            ['Thịt heo',2,100000,200000],HEADER,['Rau cải',1,30000,30000]]))
        self.assertEqual(sheets[0]['total'],230000)
        self.assertEqual(len(sheets[0]['items']),2)
        self.assertFalse(sheets[0]['errors'])

    def test_unrecognized_earlier_table_blocks_partial_export_and_can_be_remapped(self):
        source=data(PREFIX+[['Cột riêng','SL','Đơn giá','Thành tiền'],
            ['Thịt heo',2,100000,200000],HEADER,['Rau cải',1,30000,30000]])
        sheets=analyze(source)
        self.assertFalse(sheets[0]['selected'])
        self.assertTrue(sheets[0]['requires_review'])
        self.assertTrue(any('trước tiêu đề' in e for e in sheets[0]['errors']))
        sheets[0]['selected']=True
        with self.assertRaises(ValueError):apply_settings(sheets,sheets)
        title=sheets[0]['sheet']
        corrected=analyze(source,{title:{'header':3,'mapping':{'name':'A','quantity':'B','price':'C','amount':'D'}}})[0]
        self.assertEqual(corrected['total'],230000)
        self.assertEqual(len(corrected['items']),2)
        self.assertFalse(corrected['errors'])

    def test_total_item_code_does_not_drop_valid_goods(self):
        sheet=analyze(data(PREFIX+[['Mã hàng',*HEADER],
            ['SP2','Cá',1,30000,30000],['TOTAL','Rau xanh',2,10000,20000]]))[0]
        self.assertEqual(sheet['total'],50000)
        self.assertEqual([item['code'] for item in sheet['items']],['SP2','TOTAL'])
        self.assertFalse(sheet['errors'])

    def test_missing_name_and_invalid_numbers_remain_blocking(self):
        sheets=analyze(data(PREFIX+[HEADER,['Gạo',2,10000,20000],[None,'O','O','O']]))
        self.assertEqual(len(sheets[0]['items']),1)
        self.assertTrue(any('thiếu tên' in e for e in sheets[0]['errors']))
        sheets[0]['selected']=True
        with self.assertRaises(ValueError):apply_settings(sheets,sheets)

    def test_double_negative_accounting_number_is_not_positive(self):
        for raw in ['(-1)','(−1)','(-2.500)','(+1)','(VND -1)','(₫ -1)','(đồng −1)']:
            with self.subTest(raw=raw):self.assertIsNone(number(raw))
        self.assertEqual(number('(1)'),Decimal('-1'))
        self.assertEqual(number('(VND 1)'),Decimal('-1'))
        sheet=analyze(data(PREFIX+[HEADER,['Gạo','(-1)',10000,10000]]))[0]
        self.assertFalse(sheet['items'])
        self.assertTrue(sheet['errors'])

    def test_aggregate_boolean_and_reference_semantics(self):
        examples={'=SUM(TRUE,TRUE)':2,'=SUM(H1,I1,J1,TRUE,"2")':6,'=AVERAGE(K1,2)':2,'=MIN(TRUE,H1)':1}
        for formula,expected in examples.items():
            with self.subTest(formula=formula):
                workbook=Workbook();sheet=workbook.active
                for row in PREFIX+[HEADER,['Gạo',formula,10000,'=B4*C4']]:sheet.append(row)
                sheet['H1']=True;sheet['I1']='5';sheet['J1']=3
                output=io.BytesIO();workbook.save(output)
                result=analyze(output.getvalue())[0]
                self.assertFalse(result['errors'],result['errors'])
                self.assertEqual(result['items'][0]['quantity'],expected)

    def test_invalid_direct_aggregate_text_is_not_silently_zero(self):
        sheet=analyze(data(PREFIX+[HEADER,['Gạo','=SUM("abc")',10000,'=B4*C4']]))[0]
        self.assertTrue(sheet['errors'])
        self.assertFalse(sheet['items'])

    def test_ambiguous_uncached_power_formulas_require_excel_calculation(self):
        for formula in ['=2^3^2','=-2^2','=(2^3)^2']:
            with self.subTest(formula=formula):
                sheet=analyze(data(PREFIX+[HEADER,['Gạo',formula,10000,'=B4*C4']]))[0]
                self.assertTrue(sheet['errors'])
                self.assertFalse(sheet['items'])
        sheet=analyze(data(PREFIX+[HEADER,['Gạo','=2^3',10000,'=B4*C4']]))[0]
        self.assertEqual(sheet['total'],80000)
        self.assertFalse(sheet['errors'])

    def test_different_vouchers_or_delivery_info_do_not_merge(self):
        for field,first,second in [('Số phiếu','PX001','PX002'),('Người nhận hàng','Cơ sở 1','Cơ sở 2'),('Địa chỉ','Địa chỉ 1','Địa chỉ 2')]:
            with self.subTest(field=field):
                rows=PREFIX+[[f'{field}: {first}'],HEADER,['Gạo',2,10000,20000],['Tổng cộng'],[],
                    [f'{field}: {second}'],HEADER,['Rau',3,20000,60000]]
                sheet=analyze(data(rows))[0]
                self.assertTrue(any('khác bảng đầu' in e for e in sheet['errors']),sheet['errors'])

    def test_matching_vouchers_and_recipients_can_share_one_receipt(self):
        rows=PREFIX+[['Số phiếu: PX001'],['Người nhận hàng: Cơ sở 1'],HEADER,['Gạo',2,10000,20000],['Tổng cộng'],[],
            ['Số phiếu: PX001'],['Người nhận hàng: Cơ sở 1'],HEADER,['Rau',3,20000,60000]]
        sheet=analyze(data(rows))[0]
        self.assertEqual(sheet['total'],80000)
        self.assertFalse(sheet['errors'],sheet['errors'])

    def test_windows_reserved_device_name_with_extension_is_safe(self):
        for name in ['CON.foo','NUL.foo','COM1.foo','LPT9.foo','AUX','normal.foo']:
            result=zip_name(name,set())
            if name!='normal.foo':self.assertTrue(result.startswith('Phieu-'),result)
            else:self.assertEqual(result,'normal.foo.xlsx')


if __name__=='__main__':unittest.main()
