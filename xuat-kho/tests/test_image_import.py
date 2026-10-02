"""OCR row preservation and numeric validation independent of HTTP sessions."""
import io
import sys
import unittest
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from engine import analyze
from image_import import build_image_workbook


HEADER = [(20, 'Tên hàng'), (400, 'Số lượng'), (560, 'Đơn giá'), (760, 'Thành tiền')]
HEADER_UNIT = [(20, 'Tên hàng'), (280, 'ĐVT'), (400, 'Số lượng'), (560, 'Đơn giá'), (760, 'Thành tiền')]
META = [[(20, 'Khách hàng: Trường ABC')], [(20, 'Ngày xuất: 01/10/2026')]]


def tsv_document(rows, name='photo.png'):
    lines = ['level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext']
    for line, cells in enumerate(rows, 1):
        for word, (left, text) in enumerate(cells, 1):
            lines.append(f'5\t1\t1\t1\t{line}\t{word}\t{left}\t{line*45}\t{max(16, len(text)*8)}\t22\t96\t{text}')
    return {'filename': name, 'text': '\n'.join(' '.join(t for _, t in row) for row in rows), 'tsv': '\n'.join(lines)}


def text_document(text):
    return {'filename': 'photo.png', 'text': text, 'tsv': '', 'edited_text': True}


def import_document(document):
    source, report = build_image_workbook([document])
    return source, analyze(source)[0], report['documents'][0]


class ImageImportTests(unittest.TestCase):
    def test_later_richer_header_does_not_drop_first_table(self):
        document = tsv_document(META + [HEADER,
            [(20, 'Thịt heo'), (400, '2'), (560, '100.000'), (760, '200.000')],
            HEADER_UNIT, [(20, 'Rau cải'), (280, 'kg'), (400, '1'), (560, '30.000'), (760, '30.000')]])
        _, sheet, _ = import_document(document)
        self.assertEqual([i['name'] for i in sheet['items']], ['Thịt heo', 'Rau cải'])
        self.assertEqual(sheet['total'], 230000)
        self.assertFalse(sheet['errors'], sheet['errors'])

    def test_plain_repeated_headers_refresh_shuffled_columns(self):
        document = text_document('''Khách hàng: Trường ABC
Ngày xuất: 01/10/2026
Tên hàng | ĐVT | Số lượng | Đơn giá | Thành tiền
Thịt heo | kg | 2 | 100.000 | 200.000
Đơn giá | Thành tiền | Tên hàng | Số lượng | ĐVT
30.000 | 30.000 | Rau cải | 1 | kg
''')
        _, sheet, _ = import_document(document)
        self.assertEqual(sheet['total'], 230000)
        self.assertEqual([i['name'] for i in sheet['items']], ['Thịt heo', 'Rau cải'])
        self.assertFalse(sheet['errors'], sheet['errors'])

    def test_tsv_split_repeated_header_does_not_consume_adjacent_goods(self):
        upper = [(20,'Tên'), (400,'Số'), (560,'Đơn'), (760,'Thành')]
        lower = [(20,'hàng'), (400,'lượng'), (560,'giá'), (760,'tiền')]
        _, sheet, _ = import_document(tsv_document(META + [upper, lower,
            [(20,'Thịt heo'), (400,'2'), (560,'100000'), (760,'200000')], upper, lower,
            [(20,'Rau cải'), (400,'1'), (560,'30000'), (760,'30000')]]))
        self.assertEqual(sheet['total'],230000)
        self.assertEqual(len(sheet['items']),2)
        self.assertFalse(sheet['errors'],sheet['errors'])

    def test_tsv_goods_before_first_clear_header_are_preserved(self):
        _, sheet, report = import_document(tsv_document(META + [
            [(20,'Thịt heo'), (400,'2'), (560,'100000'), (760,'200000')], HEADER,
            [(20,'Rau cải'), (400,'1'), (560,'30000'), (760,'30000')]]))
        self.assertEqual(sheet['total'],230000)
        self.assertEqual(len(sheet['items']),2)
        self.assertTrue(report['warnings'])

    def test_quoted_delimiter_in_product_name_is_preserved(self):
        _, sheet, _ = import_document(text_document('Tên hàng | Số lượng | Đơn giá | Thành tiền\n"Gạo A | loại 1" | 2 | 10000 | 20000'))
        self.assertEqual(sheet['items'][0]['name'],'Gạo A | loại 1')
        self.assertEqual(sheet['total'],20000)

    def test_structured_goods_before_plain_header_are_preserved(self):
        document = text_document('''Khách hàng: Trường ABC
Ngày xuất: 01/10/2026
Thịt heo | kg | 2 | 100.000 | 200.000
Tên hàng | ĐVT | Số lượng | Đơn giá | Thành tiền
Rau cải | kg | 1 | 30.000 | 30.000
''')
        _, sheet, _ = import_document(document)
        self.assertEqual(sheet['total'], 230000)
        self.assertEqual(len(sheet['items']), 2)

    def test_product_names_are_not_customer_or_date_metadata(self):
        document = text_document('''Khách hàng: Trường ABC
Ngày xuất: 01/10/2026
Tên hàng | ĐVT | Số lượng | Đơn giá | Thành tiền
Bánh Ngày mới | kg | 2 | 100.000 | 200.000
Khách hàng đặc biệt | kg | 1 | 30.000 | 30.000
Date syrup | kg | 1 | 20.000 | 20.000
''')
        _, sheet, report = import_document(document)
        self.assertEqual(sheet['total'], 250000)
        self.assertEqual(len(sheet['items']), 3)
        self.assertEqual(sheet['customer'], 'Trường ABC')
        self.assertEqual(sheet['date'], '2026-10-01')
        self.assertEqual(report['metadata']['customer'], 'Trường ABC')

    def test_tsv_product_name_containing_date_is_retained(self):
        document = tsv_document(META + [HEADER_UNIT,
            [(20, 'Bánh Ngày mới'), (280, 'kg'), (400, '2'), (560, '100.000'), (760, '200.000')]])
        _, sheet, _ = import_document(document)
        self.assertEqual(sheet['items'][0]['name'], 'Bánh Ngày mới')
        self.assertEqual(sheet['total'], 200000)
        self.assertEqual(sheet['date'], '2026-10-01')

    def test_missing_name_blocks_even_when_numbers_are_valid(self):
        for numbers in ['2 | 100.000 | 200.000', 'O | O | O']:
            with self.subTest(numbers=numbers):
                source, sheet, report = import_document(text_document(f'''Khách hàng: Trường ABC
Ngày xuất: 01/10/2026
Tên hàng | ĐVT | Số lượng | Đơn giá | Thành tiền
Thịt heo | kg | 2 | 100.000 | 200.000
 | kg | {numbers}
'''))
                self.assertEqual(len(sheet['items']), 1)
                self.assertTrue(sheet['errors'])
                self.assertTrue(any(i['level'] == 'error' and not i['recoverable'] for i in sheet['issues']))
                self.assertEqual(report['source_rows'][-1]['cells']['name'], '')
                self.assertTrue(analyze(source)[0]['errors'], 'Source-only reanalysis must still block')

    def test_formula_like_ocr_number_stays_invalid(self):
        _, sheet, report = import_document(text_document('''Khách hàng: Trường ABC
Ngày xuất: 01/10/2026
Tên hàng | ĐVT | Số lượng | Đơn giá | Thành tiền
Thịt heo | kg | =2 | 100.000 | 200.000
'''))
        self.assertFalse(sheet['items'])
        self.assertTrue(sheet['errors'])
        self.assertEqual(report['source_rows'][0]['cells']['quantity'], '=2')

    def test_unknown_quantity_preserves_raw_row_and_blocks(self):
        source, sheet, report = import_document(text_document('''Khách hàng: Trường ABC
Ngày xuất: 01/10/2026
Tên hàng | ĐVT | Số lượng | Đơn giá | Thành tiền
Thịt heo | kg | ? | 100.000 | 200.000
'''))
        self.assertTrue(sheet['errors'])
        self.assertEqual(report['source_rows'][0]['cells']['quantity'], '?')
        workbook = load_workbook(io.BytesIO(source))
        self.assertEqual(workbook.active['D4'].value, '?')

    def test_edited_text_has_no_fake_ocr_confidence(self):
        _, _, report = import_document(text_document('Tên hàng | Số lượng | Đơn giá | Thành tiền\nGạo | 2 | 10000 | 20000'))
        self.assertIsNone(report['ocr_confidence'])
        self.assertEqual(report['confidence_origin'], 'edited-text')

    def test_invalid_tsv_coordinates_reject_document(self):
        document = tsv_document(META + [HEADER, [(20, 'Gạo'), (400, '2'), (560, '10000'), (760, '20000')]])
        document['tsv'] = document['tsv'].replace('\t20\t180\t', '\t-20\t180\t')
        with self.assertRaises(ValueError):
            build_image_workbook([document])

    def test_headerless_tsv_proposal_is_complete_and_unapproved(self):
        document = tsv_document(META + [
            [(20, 'Gạo'), (400, '2'), (560, '10000'), (760, '20000')],
            [(20, 'Rau'), (400, '3'), (560, '20000'), (760, '60000')]])
        _, sheet, report = import_document(document)
        self.assertEqual(sheet['total'], 80000)
        self.assertEqual(report['detection_method'], 'ocr-inferred')
        self.assertTrue(report['needs_review'])

    def test_customer_change_in_one_image_blocks_merge(self):
        _, sheet, report = import_document(text_document('''Khách hàng: Trường ABC
Ngày xuất: 01/10/2026
Tên hàng | Số lượng | Đơn giá | Thành tiền
Gạo | 2 | 10000 | 20000
Khách hàng: Trường DEF
Tên hàng | Số lượng | Đơn giá | Thành tiền
Rau | 3 | 20000 | 60000
'''))
        self.assertTrue(sheet['errors'])
        self.assertTrue(any(row['type'] == 'metadata' for row in report['source_rows']))

    def test_duplicate_image_names_make_distinct_sheets(self):
        document = text_document('Tên hàng | Số lượng | Đơn giá | Thành tiền\nGạo | 2 | 10000 | 20000')
        source, report = build_image_workbook([document, document])
        self.assertEqual(len(analyze(source)), 2)
        self.assertEqual(len(report['image_info']), 2)


if __name__ == '__main__':
    unittest.main()
